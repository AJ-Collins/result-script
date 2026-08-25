# TUM ICI — Special/Supplementary Results Processor

Turns the multi-hundred-page "CONSOLIDATED SUPPLEMENTARY/SPECIAL MARKSHEET"
PDF into a summarized Word document (grouped by programme type and year,
with a totals table and signature lines).

## Files

- `extract.py` — parses the source PDF and emits a JSON of programme/year
  summary records.
- `generate.py` — reads that JSON and builds the output `.docx`.
- `requirements.txt` — Python dependencies.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
# Step 1: extract structured data from the PDF
python extract.py path/to/RESULTS.pdf
# -> writes ./output_extracted.json (prints every record it resolves, for auditing)

# Step 2: generate the Word document from that JSON
python generate.py output_extracted.json path/to/OUTPUT.docx
```

## How extraction works

The source PDF has three sheet types, each ending in a `SUMMARY` block:

```
SUMMARY                     <- optional, not always present
<label1> [<label2> ...] TOTAL
<n1> [<n2> ...] <grand_total>
```

- **pass_fail**: labels are exactly `PASS FAIL ABSENT`
- **competency** (TVET/ICT Level 6 sheets): labels are exactly
  `MTY PFY CPT NYC ABS`
- **classification** (degree/diploma finalist sheets): one or more of
  `CREDIT / PASS / DISTINCTION / FIRST CLASS / SECOND CLASS HONOURS(...) /
  THIRD CLASS / FAIL`

`extract.py` scans the whole document as one flattened, ordered line stream
(so a summary block whose numbers land on the *next* physical page — this
happens at least once in the source PDF — is still resolved correctly), and
only accepts a record once the per-label counts are checked to sum to the
stated grand total. One page in the sample PDF prints a *cumulative* count
instead of per-label counts (e.g. `CREDIT PASS TOTAL` → `10 12 12`, meaning
CREDIT=10 then a running total of 12, not CREDIT=10 PASS=12); this is
detected and unwound by differencing, not silently miscounted.

Programme name, year of study, series and academic year are re-read from
each page's own header, so state never silently carries over stale values
between programmes — this was a real bug in an earlier draft of this
approach (Bachelor of Science and TVET Level-6 competency sheets shared no
common header pattern, so the regex needed extending — see
`RE_PROGRAMME` in `extract.py`).

## KNOWN LIMITATION — SP vs SUP split

The output groups every metric into SP (Special) / SUP (Supplementary)
sub-columns to match the reference document format. **This split is not
computed** — every combined total from the PDF's own SUMMARY block is
placed under SUP, and SP is left as `-`.

Why: the source PDF marks Special (`#`) vs Supplementary (`##`) exam
attempts *per unit inside a student's row*, not per student, and the
marker can appear on a wrapped continuation line disconnected from the
rest of that student's row (see page 1 of the sample PDF: `# 84 48 - ...`
sits on its own line below the student's main row of marks). Rolling
that up into a reliable per-student SP/SUP tag — and therefore a
trustworthy SP/SUP split of Registered/Absent/Sat/Pass/Fail — is a
separate, materially harder parsing problem than reading the SUMMARY
totals, and I did not want to ship an unverified guess as if it were
data.

If you have a rule for this (e.g. a separate source list of which
students sat Special vs Supplementary, or a convention I'm missing),
tell me and I'll wire real per-student attribution into `extract.py`
rather than leaving the column as `-`.

## Extending

- New sheet type: add its fixed label set next to `PFAST_LABELS` /
  `COMP_LABELS` in `extract.py`, or extend `CLASS_LABEL_ALTS` if it's a
  classification-style ("<label> TOTAL") sheet.
- New programme name pattern not being picked up: extend `RE_PROGRAMME`.
  Sanity-check with:
  ```python
  # pages with a header block but no programme match:
  for pg in pdf.pages:
      text = pg.extract_text() or ""
      if "YEAR OF STUDY" in text.upper() and not RE_PROGRAMME.search(text):
          print("MISSED:", text[:200])
  ```
- Output table layout: `generate.py`'s `build_main_table` /
  `build_summary_table` are plain `python-docx` calls (no raw OOXML
  templating), so column labels/order/widths are straightforward to edit
  in `MAIN_METRICS` / `EXTRA_SINGLE` / `set_col_widths`.
