"""
Robust buffer-based extractor for TUM mark-sheet PDFs.

Three SUMMARY scenarios observed in the document:
  A) Self-contained   — SUMMARY and its numbers are on the same page.
  B) Cross-page split — SUMMARY (+header row) ends page N; the four numeric
                        values appear at the top of page N+1 (which also has
                        a new programme header and sometimes an intake-code
                        artefact between the header row and the numbers).
  C) Interior page    — A "Page X of Y" continuation page has SUMMARY+nums
                        mid-page with no programme header of its own.

  D) Honours/Classification sheets — use a completely different summary
     layout (classification labels + totals, no PASS/FAIL/ABSENT).
     These are extracted separately with their own pattern.

Strategy
--------
* `state` accumulates heading fields and carries forward unchanged to
  continuation / split pages.
* `summary_buffer` accumulates raw text once "SUMMARY" is spotted and grows
  across page boundaries until resolved, then resets.
* Disambiguation: if "CLASSIFICATION" appears in the summary block we treat
  it as a degree classification sheet, not a pass/fail sheet.
"""

import fitz
import re
import json
from pathlib import Path


# ─────────────────────────── Patterns ──────────────────────────────────── #

RE_PROGRAMME = re.compile(
    r"(CERTIFICATE IN[^\n]+|DIPLOMA IN[^\n]+|BACHELOR OF[^\n]+|MASTER OF[^\n]+)",
    re.IGNORECASE,
)
RE_YEAR_OF_STUDY = re.compile(r"YEAR OF STUDY\s*:?\s*(\d+)", re.IGNORECASE)
RE_ACADEMIC_YEAR = re.compile(r"ACADEMIC YEAR\s*:?\s*([\d/]+)", re.IGNORECASE)
RE_SERIES        = re.compile(r"SERIES\s*:\s*([A-Z]{3}\s+\d{4})", re.IGNORECASE)
RE_SHEET_TYPE    = re.compile(
    r"(SEMESTER ORDINARY MARK SHEET|CONSOLIDATED MARKSHEET|SUPPLEMENTARY|CLASSIFICATION)",
    re.IGNORECASE,
)
RE_SEMESTER = re.compile(
    r"SEM\s*([12])\s+UNITS\s+AND\s+MARKS",
    re.IGNORECASE,
)

# Primary: full header row + values on the same block.
# The \s* between header labels handles the case where labels appear
# on separate lines.  The [\s\S]*? bridge handles the intake-code artefact
# that sometimes appears between the header row and the numbers.
RE_SUMMARY_FULL = re.compile(
    r"PASS\s+FAIL\s+ABSENT\s+TOTAL"   # column headers
    r"[\s\S]*?"                        # bridge (may cross page artefacts)
    r"(\d+)\s+(\d+)\s+(\d+)\s+(\d+)", # four values
    re.IGNORECASE,
)

# Secondary (bare values): when the header row was on the previous page
# the buffer starts with just the four numbers before "KEY:".
# We apply a cross-check pass+fail+absent == total.
RE_BARE_NUMS = re.compile(r"\b(\d+)\b")

# Honours / classification summary:
#   "SUMMARY\n\n<CLASS LABEL>\nTOTAL\nN1\nN2\nN_total"
RE_HONOURS_SUMMARY = re.compile(
    r"SUMMARY[\s\S]*?TOTAL\s+([\d\s]+)",
    re.IGNORECASE,
)


# ─────────────────────────── Helpers ───────────────────────────────────── #

def _update_state(state: dict, text: str) -> None:
    m = RE_PROGRAMME.search(text)
    if m:
        raw = m.group(1).strip()
        prog = re.split(r"\s+-\s+[A-Z]+/", raw)[0].strip()
        state["programme_name"] = prog

    m = RE_YEAR_OF_STUDY.search(text)
    if m:
        state["year_of_study"] = int(m.group(1))

    m = RE_ACADEMIC_YEAR.search(text)
    if m:
        state["academic_year"] = m.group(1).strip()

    m = RE_SERIES.search(text)
    if m:
        state["series"] = m.group(1).strip()

    m = RE_SHEET_TYPE.search(text)
    if m:
        state["sheet_type"] = m.group(1).strip().title()

    m = RE_SEMESTER.search(text)
    if m:
        state["semester"] = int(m.group(1))


def _is_honours_summary(buf: str) -> bool:
    """True when the SUMMARY block uses classification labels, not PASS/FAIL."""
    upper = buf.upper()
    has_honours   = "HONOURS" in upper or "CLASSIFICATION" in upper
    has_pass_label = bool(re.search(r"\bPASS\b", upper))
    return has_honours and not has_pass_label


def _resolve_pass_fail(buf: str):
    """
    Try to extract (pass, fail, absent, total) from a PASS/FAIL-style buffer.
    Returns 4-tuple of ints, or None.
    """
    # Primary: full header row present in this buffer.
    m = RE_SUMMARY_FULL.search(buf)
    if m:
        p, f, a, t = (int(m.group(i)) for i in range(1, 5))
        if t == p + f + a:          # sanity
            return p, f, a, t
        # If sanity fails the regex probably grabbed the wrong numbers.
        # Fall through to secondary.

    # Secondary: header row was on a previous page; scan for four nums
    # before "KEY:" with cross-check.
    key_idx = buf.upper().find("KEY:")
    region  = buf[:key_idx] if key_idx != -1 else buf[:400]
    nums    = RE_BARE_NUMS.findall(region)
    for i in range(len(nums) - 3):
        p, f, a, t = (int(nums[i + j]) for j in range(4))
        if t == p + f + a and t > 0:
            return p, f, a, t

    return None


def _resolve_honours(buf: str) -> dict | None:
    """
    Parse a classification-style SUMMARY block.
    Returns a dict with classification breakdown + total, or None.
    """
    m = RE_HONOURS_SUMMARY.search(buf)
    if not m:
        return None

    raw_nums = [int(x) for x in m.group(1).split() if x.isdigit()]
    if len(raw_nums) < 2:
        return None

    total = raw_nums[-1]
    counts = raw_nums[:-1]

    # Extract classification labels
    labels = re.findall(
        r"(FIRST CLASS|SECOND CLASS HONOURS\s*\([^)]+\)|THIRD CLASS|PASS|FAIL|DISTINCTION|CREDIT)",
        buf,
        re.IGNORECASE,
    )
    labels = [l.strip().upper() for l in labels]

    classification_data = {}
    for i, label in enumerate(labels):
        if i < len(counts):
            classification_data[label] = counts[i]

    return {"classification": classification_data, "total": total}


# ─────────────────────── Main extractor ────────────────────────────────── #

def extract_pdf_data(pdf_path: Path) -> list[dict]:
    print(f"Opening: {pdf_path.name}")
    results: list[dict] = []

    state: dict = {
        "programme_name": "Unknown",
        "year_of_study":  0,
        "academic_year":  "Unknown",
        "series":         "Unknown",
        "sheet_type":     "Unknown",
        "semester": None,
    }

    doc = fitz.open(pdf_path)
    total_pages = len(doc)

    summary_buffer:    str | None  = None
    state_at_summary:  dict | None = None

    for page_num in range(len(doc)):
        raw = doc[page_num].get_text()

        # 1. Update state from this page's headers (no-op on continuation pages).
        _update_state(state, raw)

        upper = raw.upper()
        summary_pos = upper.find("SUMMARY")

        # 2. Buffer management.
        if summary_buffer is None:
            if summary_pos != -1:
                summary_buffer    = raw[summary_pos:]
                state_at_summary  = {**state}
                print(f"  [p{page_num + 1}] SUMMARY found — start buffer")
        else:
            # Append this page, but stop before a new programme heading
            # that sits near the very top (avoids merging two records).
            prog_match = RE_PROGRAMME.search(raw)
            boundary   = prog_match.start() if (prog_match and prog_match.start() < 250) else None
            chunk      = raw[:boundary] if boundary is not None else raw
            summary_buffer += "\n" + chunk
            print(f"  [p{page_num + 1}] Appending to buffer (prog_boundary={boundary is not None})")

        # 3. Try to resolve the buffer.
        if summary_buffer is not None:
            resolved = False

            if _is_honours_summary(summary_buffer):
                hon = _resolve_honours(summary_buffer)
                if hon is not None:
                    record = {**state_at_summary, "summary_type": "classification", **hon}
                    results.append(record)
                    print(f"  [p{page_num + 1}] Resolved (honours) → total={hon['total']}")
                    resolved = True
            else:
                result = _resolve_pass_fail(summary_buffer)
                if result is not None:
                    p, f, a, t = result
                    record = {
                        **state_at_summary,
                        "summary_type": "pass_fail",
                        "pass":         p,
                        "fail":         f,
                        "absent":       a,
                        "total":        t,
                        "sat_exam":     t - a,
                    }
                    results.append(record)
                    print(f"  [p{page_num + 1}] Resolved → PASS={p} FAIL={f} ABSENT={a} TOTAL={t}")
                    resolved = True

            if resolved:
                summary_buffer   = None
                state_at_summary = None
            elif len(summary_buffer) > 8000:
                print(f"  [p{page_num + 1}] WARNING: buffer exceeded limit without resolution — discarding")
                summary_buffer   = None
                state_at_summary = None

    doc.close()
    print(f"\nTotal records extracted: {len(results)}")
    return {
        "total_pages": total_pages,
        "records": results
    }


# ──────────────────────────── Entry point ──────────────────────────────── #

if __name__ == "__main__":
    pdf_path = Path(__file__).parent / "uploads" / "input.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    data = extract_pdf_data(pdf_path)

    # Save output in the script's directory
    out_path = Path(__file__).parent / "output_final.json"
    with open(out_path, "w") as f:
        json.dump(data, f, indent=2)
    print(f"Saved {len(data)} records to {out_path}")