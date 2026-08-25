"""
TUM ICI - Special/Supplementary Results Extractor
===================================================
Uses pdfplumber (not pymupdf/fitz) because pdfplumber's extract_text()
groups words into properly-ordered visual lines for this document, which
makes every SUMMARY block trivial to parse:

    SUMMARY                              [optional]
    <label1> [<label2> ...] TOTAL
    <n1> [<n2> ...] <grand_total>

This covers all three sheet types seen in the PDF:
  - pass_fail:      PASS FAIL ABSENT TOTAL
  - classification: CREDIT / PASS / DISTINCTION / FIRST CLASS / SECOND CLASS
                     HONOURS(...) / THIRD CLASS  (one or more, then TOTAL)
  - competency:      MTY PFY CPT NYC ABS TOTAL

Watermark/OCR noise sometimes injects stray single characters into lines
(e.g. "1 0 A 0 1"); we extract numbers with re.findall(r'\\d+', line) so
that noise is ignored rather than breaking the token count.

One genuine cross-page split was found in this PDF (summary numbers
appear at the very top of the NEXT page). The line-stream is flattened
across page boundaries so this is handled the same way as same-page
cases.

SP / SUP attribution (pass_fail sections only)
-----------------------------------------------
Each student in a pass_fail section occupies two lines:
  Line 1: "<serial> <student_no> <name> <SEX> <marks...> <T.UNITS>"
  Line 2+: "#  <resit_marks...> <T.UNITS> <T.MARKS> <MEAN> <REMARKS>"  <- Special
          "## <resit_marks...> <T.UNITS> [<T.MARKS> <MEAN> <REMARKS>]" <- Supplementary

The '#' / '##' prefix on the resit row is the authoritative SP/SUP signal
(per the KEY legend printed on every page). A student with BOTH a '#' row
and a '##' row is counted as 1 in SP AND 1 in SUP (Option A).

Students with no '#'/'##' row (edge case) default to SUP.
Classification and competency pages have no '#'/'##' rows and do not
contribute SP/SUP counts.

Known PDF quirks handled
------------------------
  Q1. Classification-section student rows (with historical series dates like
      "NOV2023 ... APRIL2026 ... CLASSIFIED") span into the preceding pass_fail
      section's line range.  _is_classification_summary_row() filters them out.

  Q2. Some students have MEAN='-' (all units absent/failed, no computable mean).
      The REMARKS text (e.g. "6 ABS, 6 FAIL") is still present.
      _student_outcome() falls back to keyword scan when no float MEAN is found.

  Q3. pdfplumber sometimes splits a single PDF text line across two extracted
      lines when watermark characters fall between words.  The most affected
      pattern is REMARKS text split at the end-of-line boundary (e.g. "11 ABS, 1"
      on one line, "FAIL" on the next).  _collect_student_lines() merges the
      continuation line back before passing to _student_outcome().

  Q4. A student row with no unit marks and no REMARKS at all (completely absent,
      no resit attempted) cannot have its outcome determined.  These default to
      "absent" — the most defensively correct assumption.
"""

import re
import pdfplumber
from pathlib import Path


# ═══════════════════════════════ Patterns ════════════════════════════════════

RE_PROGRAMME = re.compile(
    r"(CERTIFICATE IN[^\n]+|DIPLOMA IN[^\n]+|BACHELOR OF[^\n]+|MASTER OF[^\n]+|"
    r"COMPUTER SCIENCTIST[^\n]+|COMPUTER SCIENTIST[^\n]+|"
    r"INFORMATION COMMUNICATION TECHNOLOGIST[^\n]+)",
    re.IGNORECASE,
)
RE_YOS = re.compile(r"YEAR OF STUDY\s*:?\s*(\d+)", re.IGNORECASE)
RE_AY  = re.compile(r"ACADEMIC YEAR\s*:?\s*(\d{4}\s*/\s*\d{4})", re.IGNORECASE)
RE_SER = re.compile(
    r"SERIES\s*:?\s*((?:APR|JAN|MAY|AUG|SEPT?|OCT|NOV|DEC)\w*\.?\s*\d{4})",
    re.IGNORECASE,
)
RE_INTAKE = re.compile(
    r"\b([A-Z]{2,6}/(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\d{4}/[A-Z0-9-]+)",
    re.IGNORECASE,
)

CLASS_LABEL_ALTS = (
    r"FIRST CLASS HONOURS?",
    r"FIRST CLASS",
    r"SECOND CLASS HONOURS\s*\([^)]*\)",
    r"SECOND CLASS HONOURS?",
    r"THIRD CLASS",
    r"DISTINCTION",
    r"CREDIT",
    r"PASS",
    r"FAIL",
)
RE_CLASS_LABEL = re.compile("(" + "|".join(CLASS_LABEL_ALTS) + ")", re.IGNORECASE)

PFAST_LABELS = ["PASS", "FAIL", "ABSENT"]
COMP_LABELS  = ["MTY", "PFY", "CPT", "NYC", "ABS"]

# Student-row identification: "<serial>  <DEPT/NNN[X]/YYYY>  <name> ..."
RE_STUDENT_HDR = re.compile(
    r"^\s*\d+\s+[A-Z]{2,8}/\d{3,5}[A-Z]*/\d{4}\s+",
    re.IGNORECASE,
)

# ── Quirk Q1: classification-section rows that bleed into pass_fail sections ─
# These historical classification summary rows contain series-date tokens
# (e.g. "NOV2023 ... APRIL2026") and/or the word CLASSIFIED/UNCLASSIFIED.
# No genuine pass_fail student row ever contains two or more month+year tokens.
RE_MONTH_YEAR = re.compile(
    r'\b(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\w*\s*\d{4}\b',
    re.IGNORECASE,
)
RE_CLASSIFIED_MARKER = re.compile(r'\b(?:UN)?CLASSIFIED\b', re.IGNORECASE)

# ── Quirk Q2: outcome keyword scan (no float MEAN required) ──────────────────
# Only match as standalone outcome keywords, not inside unit names or noise.
RE_OUTCOME_KEYWORD = re.compile(
    r'\b(PASS|FAIL|ABS(?:ENT)?)\b', re.IGNORECASE
)

# ═══════════════════════════════ Helpers ═════════════════════════════════════

def _clean_prog(raw: str) -> str:
    raw = raw.strip()
    raw = re.split(r"\s+-\s+[A-Z0-9/]+", raw)[0].strip()          # strip " - INTAKE/CODE"
    raw = re.split(r"\s+YEAR OF STUDY\s*:", raw, flags=re.IGNORECASE)[0].strip()
    raw = re.sub(r"\s{2,}", " ", raw)
    return raw


def _update_state(state: dict, text: str) -> None:
    m = RE_PROGRAMME.search(text)
    if m:
        state["programme_name"] = _clean_prog(m.group(1))
    m = RE_YOS.search(text)
    if m:
        state["year_of_study"] = int(m.group(1))
    m = RE_AY.search(text)
    if m:
        state["academic_year"] = re.sub(r"\s+", "", m.group(1))
    m = RE_SER.search(text)
    if m:
        state["series"] = re.sub(r"\s+", " ", m.group(1)).strip()
    m = RE_INTAKE.search(text)
    if m:
        state["intake_code"] = m.group(1).upper()


def _try_label_line(line: str):
    """
    If `line` is a "<label(s)> TOTAL" summary header, return (kind, labels).
    kind in {"pass_fail", "competency", "classification"}.  Otherwise None.
    """
    up = line.upper().strip()
    if not up.endswith("TOTAL"):
        return None
    body = up[: -len("TOTAL")].strip()
    if not body:
        return None

    tokens = body.split()

    if tokens == PFAST_LABELS:
        return "pass_fail", PFAST_LABELS
    if tokens == COMP_LABELS:
        return "competency", COMP_LABELS

    # Classification: body must be fully consumed by known class labels
    labels = []
    pos = 0
    while pos < len(body):
        m = RE_CLASS_LABEL.match(body[pos:])
        if not m:
            if body[pos] == " ":
                pos += 1
                continue
            return None
        labels.append(m.group(1).strip().upper())
        pos += m.end()
        while pos < len(body) and body[pos] == " ":
            pos += 1
    if labels:
        return "classification", labels
    return None


def _find_numbers(lines: list[str], start: int, need: int, max_lookahead: int = 6):
    """
    Search forward from `lines[start]` for a line containing exactly `need`
    digit-runs (ignoring watermark noise glued to the digits). Skips blank
    lines and lines with zero digit-runs. Returns (numbers, line_index) or
    (None, None).
    """
    for j in range(start, min(start + max_lookahead, len(lines))):
        digits = re.findall(r"\d+", lines[j])
        if len(digits) == need:
            return [int(x) for x in digits], j
    return None, None


def _resit_type(line: str) -> str | None:
    """
    Return the type of a resit row from its leading '#' marker.
    Returns 'sp' (Special), 'sup' (Supplementary), 'retake', 'correction',
    or None if not a resit row.

    Must check '####' before '###' before '##' before '#' to avoid
    prefix-collision.
    """
    s = line.lstrip()
    if s.startswith("####"):
        return "correction"
    if s.startswith("###"):
        return "retake"
    if s.startswith("##"):
        return "sup"
    if s.startswith("#"):
        return "sp"
    return None


def _is_classification_summary_row(line: str) -> bool:
    """
    Return True if a student-header-matching line is actually a classification
    summary row from the consolidated history sheet (Quirk Q1).

    Detection (either condition is sufficient):
      1. Contains CLASSIFIED or UNCLASSIFIED.
      2. Contains >=2 month+year tokens (prior-series labels like "NOV2023 … APRIL2026").
         Genuine pass_fail student rows never contain month+year tokens in data fields.
    """
    if RE_CLASSIFIED_MARKER.search(line):
        return True
    if len(RE_MONTH_YEAR.findall(line)) >= 2:
        return True
    return False


def _student_outcome(rows: list[str]) -> str:
    """
    Determine a student's overall outcome from one or more text lines.

    Strategy 1 (preferred): float MEAN followed by REMARKS text.
      Searches for pattern "<float> <REMARKS>" at end of line.
      REMARKS "PASS" -> pass; contains "FAIL" -> fail; contains "ABS" -> absent.

    Strategy 2 (fallback for Quirk Q2): keyword scan without MEAN.
      Scans each non-header line for standalone PASS/FAIL/ABS keywords.
      Used when MEAN is '-' (e.g. all units absent/failed, no mean computed).

    Returns 'pass', 'fail', 'absent', or 'unknown'.
    """
    # Strategy 1: float MEAN + REMARKS
    for row in rows:
        m = re.search(r"\d+\.\d+\s+(.+?)\s*$", row)
        if m:
            rem = m.group(1).strip().upper()
            if rem == "PASS":       return "pass"
            if "FAIL" in rem:      return "fail"
            if "ABS" in rem:       return "absent"

    # Strategy 2: bare REMARKS keyword scan (Quirk Q2).
    # The REMARKS field can contain unit-level counts like "6 ABS, 6 FAIL",
    # meaning 6 units absent and 6 units failed.  "ABS" here refers to absent
    # units, not the overall student outcome.  We therefore use priority ordering:
    #   FAIL > ABS > PASS
    # (if a student has any failing units they are overall FAIL, even if some
    # units were absent; only if the line has ABS but no FAIL is the outcome absent.)
    for row in rows:
        # Skip header-like lines to avoid false matches
        if re.search(r"STUDENT\s+NAME|MEAN\s+REMARKS|UNIT\s+(?:CODE|NAME)", row, re.IGNORECASE):
            continue
        found = {m.group(1).upper() for m in RE_OUTCOME_KEYWORD.finditer(row)}
        if not found:
            continue
        if "FAIL" in found:                     return "fail"
        if "ABS" in found or "ABSENT" in found: return "absent"
        if "PASS" in found:                     return "pass"

    return "unknown"


def _parse_student_blocks(section_lines: list[str]) -> dict:
    """
    Scan lines from a student section (lines between the previous SUMMARY
    record's numbers line and the current SUMMARY label line) and count
    SP vs SUP students.

    Rules:
    - Student with '#'  resit row  -> Special (SP).
    - Student with '##' resit row  -> Supplementary (SUP).
    - Student with BOTH            -> counts as 1 in SP AND 1 in SUP (Option A).
    - Student with NO '#'/'##' row -> defaults to SUP (edge case).
    - '###' (Retake) and '####' (Correction) rows are ignored.

    Outcome (pass/fail/absent) is read from REMARKS (Strategy 1: with MEAN float;
    Strategy 2: bare keyword scan).  Students whose outcome still cannot be
    determined after all strategies default to 'absent' (Quirk Q4 — conservative).

    Quirk Q1: classification summary rows that match RE_STUDENT_HDR are filtered
    out before building the student list.

    Quirk Q3: REMARKS text occasionally split across two consecutive lines by
    pdfplumber (watermark character between words).  The line immediately following
    a student-header or resit row that contains only a bare outcome keyword is
    merged into the preceding line before outcome detection.

    Returns dict with keys:
      sp_registered, sp_absent, sp_sat_exam, sp_pass, sp_fail,
      sup_registered, sup_absent, sup_sat_exam, sup_pass, sup_fail
    """
    # ── Pass 1: build merged line list (handle Quirk Q3 split-line REMARKS) ───
    merged: list[str] = []
    for raw in section_lines:
        line = raw  # already stripped by caller
        # Check if this line is a bare continuation fragment of the previous line.
        # A continuation fragment: short, no student-number prefix, no '#' prefix,
        # and consists only of an outcome keyword (possibly with noise chars).
        stripped = line.strip()
        if (merged
                and not RE_STUDENT_HDR.match(stripped)
                and _resit_type(stripped) is None
                and re.fullmatch(r'[^a-z]*(?:PASS|FAIL|ABS(?:ENT)?)[^a-z]*', stripped, re.IGNORECASE)
                and len(stripped) <= 30):
            # Merge into previous line
            merged[-1] = merged[-1] + " " + stripped
        else:
            merged.append(line)

    # ── Pass 2: collect (first_line, [resit_rows]) per genuine student ────────
    students: list[tuple[str, list[str]]] = []
    cur_resits: list[str] | None = None

    for line in merged:
        if RE_STUDENT_HDR.match(line):
            # Quirk Q1: skip classification summary rows
            if _is_classification_summary_row(line):
                cur_resits = None
                continue
            cur_resits = []
            students.append((line, cur_resits))
        elif cur_resits is not None:
            rt = _resit_type(line)
            if rt in ("sp", "sup"):
                cur_resits.append(line)

    # ── Pass 3: count outcomes ────────────────────────────────────────────────
    sp_reg = sp_pass = sp_fail = sp_abs = 0
    sup_reg = sup_pass = sup_fail = sup_abs = 0

    for first_line, resit_rows in students:
        has_sp  = any(_resit_type(r) == "sp"  for r in resit_rows)
        has_sup = any(_resit_type(r) == "sup" for r in resit_rows)

        if not has_sp and not has_sup:
            # No '#'/'##' marker found — default to SUP (conservative).
            has_sup = True

        # Try resit rows first; fall back to first-attempt row.
        outcome = _student_outcome(resit_rows)
        if outcome == "unknown":
            outcome = _student_outcome([first_line])
        if outcome == "unknown":
            # Quirk Q4: completely absent student with no REMARKS at all.
            # Default to 'absent' — most defensively correct.
            outcome = "absent"
            print(f"    [INFO] no REMARKS found; defaulting to absent: {first_line[:80]}")

        if has_sp:
            sp_reg += 1
            if outcome == "pass":     sp_pass += 1
            elif outcome == "fail":   sp_fail += 1
            elif outcome == "absent": sp_abs  += 1

        if has_sup:
            sup_reg += 1
            if outcome == "pass":     sup_pass += 1
            elif outcome == "fail":   sup_fail += 1
            elif outcome == "absent": sup_abs  += 1

    return {
        "sp_registered":  sp_reg,
        "sp_absent":      sp_abs,
        "sp_sat_exam":    max(0, sp_reg  - sp_abs),
        "sp_pass":        sp_pass,
        "sp_fail":        sp_fail,
        "sup_registered": sup_reg,
        "sup_absent":     sup_abs,
        "sup_sat_exam":   max(0, sup_reg - sup_abs),
        "sup_pass":       sup_pass,
        "sup_fail":       sup_fail,
    }


# ══════════════════════════════ Main extractor ════════════════════════════════

def extract_pdf_data(pdf_path: Path) -> dict:
    print(f"\nOpening: {pdf_path}")
    pdf = pdfplumber.open(str(pdf_path))
    total_pages = len(pdf.pages)

    # Flatten to one long line stream, remembering which page each line came from
    all_lines: list[str] = []
    line_page: list[int] = []
    page_texts: list[str] = []
    for pg_num, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        page_texts.append(text)
        for l in text.split("\n"):
            l = l.strip()
            if l:
                all_lines.append(l)
                line_page.append(pg_num)
    pdf.close()

    RE_BANNER_AY = re.compile(
        r"(\d{4}\s*/\s*\d{4})\s+ACADEMIC YEAR SPECIAL/SUPPLEMENTARY", re.IGNORECASE
    )
    global_series = "Unknown"
    banner_ay_counts: dict[str, int] = {}
    for t in page_texts:
        m = RE_SER.search(t)
        if m and global_series == "Unknown":
            global_series = re.sub(r"\s+", " ", m.group(1)).strip()
        for bm in RE_BANNER_AY.finditer(t):
            y = re.sub(r"\s+", "", bm.group(1))
            banner_ay_counts[y] = banner_ay_counts.get(y, 0) + 1
    global_ay = (max(banner_ay_counts, key=banner_ay_counts.get)
                 if banner_ay_counts else "Unknown")
    print(f"Banner academic-year counts across pages: {banner_ay_counts} -> using {global_ay}")

    state = {
        "programme_name": "Unknown", "year_of_study": 0,
        "academic_year": global_ay, "series": global_series,
        "intake_code": "Unknown",
    }

    records: list[dict] = []
    cur_page_for_state = -1

    # section_start tracks the index of the first line AFTER the previous
    # record's numbers line.  Lines in [section_start, i) are the student
    # rows belonging to the current (about-to-be-parsed) SUMMARY block.
    section_start = 0

    i = 0
    n = len(all_lines)
    while i < n:
        pg = line_page[i]
        if pg != cur_page_for_state:
            _update_state(state, page_texts[pg])
            cur_page_for_state = pg

        line = all_lines[i]
        parsed = _try_label_line(line)

        if parsed is not None:
            kind, labels = parsed
            need = len(labels) + 1
            nums, found_idx = _find_numbers(all_lines, i + 1, need)

            if nums is not None:
                *counts, grand_total = nums
                rec_state = {**state}
                if line_page[found_idx] != pg:
                    tmp_state = {**state}
                    _update_state(tmp_state, page_texts[line_page[found_idx]])
                    if RE_PROGRAMME.search(page_texts[line_page[found_idx]]):
                        rec_state = tmp_state

                # ── SP / SUP attribution ──────────────────────────────────
                if kind == "pass_fail":
                    section_lines = all_lines[section_start:i]
                    sp_sup = _parse_student_blocks(section_lines)
                else:
                    sp_sup = {}
                # ─────────────────────────────────────────────────────────

                if kind == "pass_fail":
                    p, f, a = counts
                    if grand_total == p + f + a:
                        records.append({
                            **rec_state, "summary_type": "pass_fail",
                            "pass": p, "fail": f, "absent": a,
                            "total": grand_total, "sat_exam": grand_total - a,
                            "source_page": pg + 1,
                            **sp_sup,
                        })
                        sp_str = (
                            f"sp={sp_sup.get('sp_registered',0)}/"
                            f"{sp_sup.get('sp_pass',0)}P/{sp_sup.get('sp_fail',0)}F  "
                            f"sup={sp_sup.get('sup_registered',0)}/"
                            f"{sp_sup.get('sup_pass',0)}P/{sp_sup.get('sup_fail',0)}F"
                        )
                        print(f"  [p{pg+1}] pass_fail  {rec_state['programme_name'][:33]:33s} "
                              f"yr{rec_state['year_of_study']} TOT={grand_total} | {sp_str}")
                elif kind == "competency":
                    mty, pfy, cpt, nyc, ab = counts
                    if grand_total == sum(counts):
                        records.append({
                            **rec_state, "summary_type": "competency",
                            "mty": mty, "pfy": pfy, "cpt": cpt, "nyc": nyc, "abs": ab,
                            "total": grand_total,
                            "pass": mty + pfy + cpt, "fail": nyc, "absent": ab,
                            "sat_exam": grand_total - ab,
                            "source_page": pg + 1,
                        })
                        print(f"  [p{pg+1}] competency {rec_state['programme_name'][:33]:33s} "
                              f"yr{rec_state['year_of_study']} MTY={mty} PFY={pfy} CPT={cpt} "
                              f"NYC={nyc} ABS={ab} TOT={grand_total}")
                else:  # classification
                    cls = {}
                    ok = False
                    if len(counts) == len(labels):
                        cand = {}
                        for lbl, cnt in zip(labels, counts):
                            cand[lbl] = cand.get(lbl, 0) + cnt
                        if sum(cand.values()) == grand_total:
                            cls, ok = cand, True
                        elif len(labels) >= 2 and counts == sorted(counts) and counts[-1] == grand_total:
                            cand = {}
                            prev = 0
                            for lbl, cum in zip(labels, counts):
                                cand[lbl] = cand.get(lbl, 0) + (cum - prev)
                                prev = cum
                            if sum(cand.values()) == grand_total and all(v >= 0 for v in cand.values()):
                                cls, ok = cand, True
                    if ok:
                        records.append({
                            **rec_state, "summary_type": "classification",
                            "classification": cls, "total": grand_total,
                            "source_page": pg + 1,
                        })
                        print(f"  [p{pg+1}] classif.   {rec_state['programme_name'][:33]:33s} "
                              f"yr{rec_state['year_of_study']} {cls} TOT={grand_total}")

                section_start = found_idx + 1
                i = found_idx + 1
                continue

        i += 1

    print(f"\nTotal records extracted: {len(records)}")
    return {
        "total_pages": total_pages,
        "series": global_series,
        "academic_year": global_ay,
        "records": records,
    }


if __name__ == "__main__":
    import json, sys
    pdf_arg = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("input.pdf")
    data = extract_pdf_data(pdf_arg)
    out = Path("output_extracted.json")
    out.write_text(json.dumps(data, indent=2))
    print(f"\nSaved {len(data['records'])} records -> {out}")