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
# Matches lines like: "1 DCS/005J/2025 NGUMBI MUSAU ..."
#                     "12 DICT/751J/2023 LENNOX KUJAH ..."
#                     "2  DCS/0106/2023 MWAKISHA MIGUEL ..."
RE_STUDENT_HDR = re.compile(
    r"^\s*\d+\s+[A-Z]{2,8}/\d{3,5}[A-Z]*/\d{4}\s+",
    re.IGNORECASE,
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
            # allow stray whitespace
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


def _student_outcome(rows: list[str]) -> str:
    """
    Determine a student's overall resit outcome from their resit row(s).

    Searches for the row containing a MEAN value (a float like "58.08")
    followed by REMARKS text.  REMARKS determines the outcome:
      - "PASS"        -> 'pass'
      - contains FAIL -> 'fail'
      - contains ABS  -> 'absent'
      - otherwise     -> 'unknown'

    For students with both '#' and '##' rows, the '#' row normally carries
    the full summary (T.MARKS, MEAN, REMARKS); the '##' row may carry only
    unit marks and T.UNITS.  Scanning all rows and returning the first match
    handles both single-type and dual-type students correctly.
    """
    for row in rows:
        m = re.search(r"\d+\.\d+\s+(.+?)\s*$", row)
        if m:
            rem = m.group(1).strip().upper()
            if rem == "PASS":
                return "pass"
            if "FAIL" in rem:
                return "fail"
            if "ABS" in rem:
                return "absent"
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

    Outcome (pass/fail/absent) is read from the REMARKS field on the resit
    row that carries T.MARKS and MEAN.  Students whose outcome cannot be
    determined are counted in registered only (conservative).

    Returns dict with keys:
      sp_registered, sp_absent, sp_sat_exam, sp_pass, sp_fail,
      sup_registered, sup_absent, sup_sat_exam, sup_pass, sup_fail
    """
    # Collect (first_line, [resit_rows]) per student
    students: list[tuple[str, list[str]]] = []
    cur_resits: list[str] | None = None

    for line in section_lines:
        if RE_STUDENT_HDR.match(line):
            cur_resits = []
            students.append((line, cur_resits))
        elif cur_resits is not None:
            rt = _resit_type(line)
            if rt in ("sp", "sup"):
                cur_resits.append(line)
            # retake / correction / header noise lines are skipped silently

    sp_reg = sp_pass = sp_fail = sp_abs = 0
    sup_reg = sup_pass = sup_fail = sup_abs = 0

    for first_line, resit_rows in students:
        has_sp  = any(_resit_type(r) == "sp"  for r in resit_rows)
        has_sup = any(_resit_type(r) == "sup" for r in resit_rows)

        if not has_sp and not has_sup:
            # No '#'/'##' marker found — default to SUP (conservative).
            has_sup = True

        # Prefer resit-row REMARKS; fall back to first-attempt row for the
        # rare case where the resit mark is baked into the first row.
        outcome = _student_outcome(resit_rows)
        if outcome == "unknown":
            outcome = _student_outcome([first_line])
        if outcome == "unknown":
            print(f"    [WARN] could not determine outcome: {first_line[:80]}")

        if has_sp:
            sp_reg += 1
            if outcome == "pass":    sp_pass += 1
            elif outcome == "fail":  sp_fail += 1
            elif outcome == "absent": sp_abs  += 1

        if has_sup:
            sup_reg += 1
            if outcome == "pass":    sup_pass += 1
            elif outcome == "fail":  sup_fail += 1
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
    # The document banner's academic year varies by cohort/intake section;
    # use whichever value appears on the most pages as the document-level one.
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
            # (Re)compute state fresh from this page's own header text,
            # since each page restates its own programme/year/series.
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
                # If the numbers line was on a LATER page than the label line,
                # re-derive state from that later page's own header (rare;
                # keeps programme attribution correct across the one genuine
                # cross-page split found in this document).
                if line_page[found_idx] != pg:
                    tmp_state = {**state}
                    _update_state(tmp_state, page_texts[line_page[found_idx]])
                    # Only trust the later page's header if it actually has one;
                    # otherwise keep the original page's state (continuation page).
                    if RE_PROGRAMME.search(page_texts[line_page[found_idx]]):
                        rec_state = tmp_state

                # ── SP / SUP attribution ──────────────────────────────────
                # Only pass_fail sections have '#'/'##' student resit rows.
                # Classification and competency sections are left without
                # sp_* fields; the generator will show '-' for those rows.
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
                            # Cumulative variant seen on p104: "CREDIT PASS TOTAL" -> "10 12 12"
                            # i.e. counts are a running sum ending at grand_total; back out
                            # per-label counts by differencing consecutive cumulative values.
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

                # Advance section boundary past the numbers line.
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
