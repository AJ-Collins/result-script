"""
TUM ICI – Special/Supplementary Results Word Doc Generator
===========================================================
Builds the summarised results Word document from the JSON produced by
extract.py, matching the reference layout exactly:

  Table 1 (Main): grouped by programme type → Year of Study (desc)
  Table 2 (Summary): one row per unique programme name + TOTAL row

Column layout for both tables (matching the reference):
  S/No | Programme Name | No. Registered SP/SUP | No. Absent SP/SUP |
  No. Sat Exam SP/SUP | No. Pass SP/SUP | No. Fail SP/SUP |
  Deregistered SP/SUP | Disciplinary SP/SUP | Incomplete SP/SUP |
  Total SP/SUP   [+ No. Classified in summary table only]

ALL 9 metric groups carry SP and SUP sub-columns.
Programmes with no SP/SUP data (competency / classification-only) show
"-" in the SP column and the combined total in the SUP column.
"""

import sys
import json
import os
import re
import shutil
import zipfile
import tempfile
from collections import defaultdict
from datetime import date
from xml.sax.saxutils import escape as xml_escape

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Book Antiqua"


def _esc(text):
    return xml_escape(str(text))


# ─────────────────────────────────────────────────────────────────────────────
# XML helpers
# ─────────────────────────────────────────────────────────────────────────────

def _find_table_end(content, tbl_start):
    depth, pos = 0, tbl_start
    while pos < len(content):
        no = content.find('<w:tbl>',  pos)
        nc = content.find('</w:tbl>', pos)
        if nc == -1:
            raise ValueError("Malformed XML: unmatched <w:tbl>")
        if no != -1 and no < nc:
            depth += 1; pos = no + 7
        else:
            depth -= 1; pos = nc + 8
            if depth == 0:
                return pos
    raise ValueError("Malformed XML: could not find matching </w:tbl>")


def _find_all_top_level_tables(content):
    tables, pos = [], 0
    while True:
        start = content.find('<w:tbl>', pos)
        if start == -1:
            break
        end = _find_table_end(content, start)
        tables.append((start, end))
        pos = end
    return tables


# ─────────────────────────────────────────────────────────────────────────────
# Cell / table helpers
# ─────────────────────────────────────────────────────────────────────────────

def _set_cell(cell, text, bold=False, size=8, align="center"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if align == "center"
                   else WD_ALIGN_PARAGRAPH.LEFT)
    r = p.add_run(str(text))
    r.font.name = FONT
    r.font.size = Pt(size)
    r.bold = bold
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _shade(cell, color="D9D9D9"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def _set_col_widths(table, widths_cm):
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        for gc, w in zip(tblGrid.findall(qn("w:gridCol")), widths_cm):
            gc.set(qn("w:w"), str(int(Cm(w).twips)))
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            if i < len(row.cells):
                row.cells[i].width = Cm(w)

def _tight_margins(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblCellMar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "20")
        el.set(qn("w:type"), "dxa")
        tblCellMar.append(el)
    tblPr.append(tblCellMar)
    
def _set_borders(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        b.append(el)
    tblPr.append(b)


# ─────────────────────────────────────────────────────────────────────────────
# Aggregation
# ─────────────────────────────────────────────────────────────────────────────

def _zero():
    return {
        "registered": 0, "absent": 0, "sat_exam": 0,
        "pass": 0, "fail": 0, "total": 0, "classified": 0,
        "sp_registered": 0, "sp_absent": 0, "sp_sat_exam": 0,
        "sp_pass": 0, "sp_fail": 0,
        "sup_registered": 0, "sup_absent": 0, "sup_sat_exam": 0,
        "sup_pass": 0, "sup_fail": 0,
    }


SP_SUP_KEYS = (
    "sp_registered", "sp_absent", "sp_sat_exam", "sp_pass", "sp_fail",
    "sup_registered", "sup_absent", "sup_sat_exam", "sup_pass", "sup_fail",
)


def aggregate(records):
    """
    One row per (programme_name, year_of_study).
    Multiple intakes sharing the same programme+year are summed together.
    """
    agg = defaultdict(_zero)
    for r in records:
        key = (r["programme_name"], r["year_of_study"])
        a = agg[key]
        if r["summary_type"] in ("pass_fail", "competency"):
            a["registered"] += r["total"]
            a["absent"]     += r["absent"]
            a["sat_exam"]   += r["sat_exam"]
            a["pass"]       += r["pass"]
            a["fail"]       += r["fail"]
            a["total"]      += r["total"]
            for k in SP_SUP_KEYS:
                a[k] += r.get(k, 0)
        elif r["summary_type"] == "classification":
            a["registered"] += r["total"]
            a["sat_exam"]   += r["total"]
            a["pass"]       += r["total"]
            a["total"]      += r["total"]
            a["classified"] += r["total"]
    return agg


# ─────────────────────────────────────────────────────────────────────────────
# Column definitions
# ─────────────────────────────────────────────────────────────────────────────

def _prog_type(name):
    n = name.upper()
    if n.startswith("MASTER"):      return "Masters"
    if n.startswith("BACHELOR"):    return "Degree"
    if n.startswith("DIPLOMA"):     return "Diploma"
    if n.startswith("CERTIFICATE"): return "Certificate"
    return "Other"


TYPE_ORDER = ["Masters", "Degree", "Diploma", "Certificate", "Other"]


# Every metric column has TWO sub-columns: SP and SUP.
# Tuple: (header_label, combined_key, sp_key, sup_key)
#   None keys  -> show "-"
#   "COMPUTED" -> value is calculated from other fields
ALL_METRICS = [
    ("No. Registered", "registered", "sp_registered",  "sup_registered"),
    ("No. Absent",     "absent",     "sp_absent",       "sup_absent"),
    ("No. Sat Exam",   "sat_exam",   "sp_sat_exam",     "sup_sat_exam"),
    ("No. Pass",       "pass",       "sp_pass",         "sup_pass"),
    ("No. Fail",       "fail",       "sp_fail",         "sup_fail"),
    ("Deregistered",   None,         None,              None),
    ("Disciplinary",   None,         None,              None),
    ("Incomplete",     "incomplete", "COMPUTED",        "COMPUTED"),
    ("Total",          "total",      "TOTAL_PROXY",     "TOTAL_PROXY"),
]

# S/No + Name + 9 metrics x 2 = 20 columns
N_MAIN_COLS = 2 + len(ALL_METRICS) * 2
# Summary table: +1 for No. Classified
N_SUM_COLS  = N_MAIN_COLS + 1


def _compute_incomplete(v, has_sp_sup):
    """Return (combined_inc, sp_inc, sup_inc)."""
    comb = max(0, v["registered"] - v["sat_exam"] - v["absent"])
    if has_sp_sup:
        sp_inc  = max(0, v["sp_registered"]  - v["sp_sat_exam"]  - v["sp_absent"])
        sup_inc = max(0, v["sup_registered"] - v["sup_sat_exam"] - v["sup_absent"])
    else:
        sp_inc  = "-"
        sup_inc = comb
    return comb, sp_inc, sup_inc


def _fill_metric_cells(row_cells, start_col, vals, has_sp_sup, bold=False):
    """
    Fill 2 × len(ALL_METRICS) cells starting at start_col.
    Returns the next free column index.
    """
    inc_comb, inc_sp, inc_sup = _compute_incomplete(vals, has_sp_sup)
    col = start_col

    for label, comb_key, sp_key, sup_key in ALL_METRICS:
        if label in ("Deregistered", "Disciplinary"):
            _set_cell(row_cells[col],     "-", bold=bold)
            _set_cell(row_cells[col + 1], "-", bold=bold)

        elif label == "Incomplete":
            sp_val  = inc_sp  if has_sp_sup else "-"
            sup_val = inc_sup if has_sp_sup else inc_comb
            _set_cell(row_cells[col],     sp_val,  bold=bold)
            _set_cell(row_cells[col + 1], sup_val, bold=bold)

        elif label == "Total":
            # SP and SUP totals are proxied by sp_registered / sup_registered
            # (the most meaningful summary count for each category)
            if has_sp_sup:
                _set_cell(row_cells[col],     vals.get("sp_registered",  0), bold=bold)
                _set_cell(row_cells[col + 1], vals.get("sup_registered", 0), bold=bold)
            else:
                _set_cell(row_cells[col],     "-",                  bold=bold)
                _set_cell(row_cells[col + 1], vals.get("total", 0), bold=bold)

        else:
            # Regular split metric
            sp_val  = vals.get(sp_key,   0) if has_sp_sup else "-"
            sup_val = vals.get(sup_key,  0) if has_sp_sup else vals.get(comb_key, 0)
            _set_cell(row_cells[col],     sp_val,  bold=bold)
            _set_cell(row_cells[col + 1], sup_val, bold=bold)

        col += 2

    return col


# ─────────────────────────────────────────────────────────────────────────────
# Header builder (shared by both tables)
# ─────────────────────────────────────────────────────────────────────────────

def _add_header_rows(table, include_classified=False):
    """
    Two-row header:
      Row 1: S/No (vspan) | Programme Name (vspan) | metric_label (hspan over SP+SUP) ...
      Row 2: (blank)      | (blank)                | SP | SUP | SP | SUP ...
    """
    h1 = table.add_row().cells
    h2 = table.add_row().cells

    _set_cell(h1[0], "S/No", bold=True)
    _set_cell(h1[1], "Programme Name", bold=True)
    h1[0].merge(h2[0])
    h1[1].merge(h2[1])

    col = 2
    for label, *_ in ALL_METRICS:
        merged = h1[col].merge(h1[col + 1])
        _set_cell(merged, label, bold=True)
        _set_cell(h2[col],     "SP",  bold=True)
        _set_cell(h2[col + 1], "SUP", bold=True)
        col += 2

    if include_classified:
        merged = h1[col].merge(h2[col])
        _set_cell(merged, "No. Classified", bold=True)

    for c in list(h1) + list(h2):
        _shade(c)


# ─────────────────────────────────────────────────────────────────────────────
# Table 1: Main (grouped by type / year)
# ─────────────────────────────────────────────────────────────────────────────

def build_main_table(doc, agg):
    table = doc.add_table(rows=0, cols=N_MAIN_COLS); _set_borders(table); _tight_margins(table)
    _add_header_rows(table, include_classified=False)

    by_type = defaultdict(lambda: defaultdict(list))
    for (prog, year), vals in agg.items():
        by_type[_prog_type(prog)][year].append((prog, vals))

    serial = 1
    for ptype in TYPE_ORDER:
        if ptype not in by_type:
            continue
        for year in sorted(by_type[ptype].keys(), reverse=True):
            # Group header spanning all columns
            grow = table.add_row().cells
            merged = grow[0]
            for c in grow[1:]:
                merged = merged.merge(c)
            _set_cell(merged, f"{ptype}: Year of Study: {year}", bold=True)
            _shade(merged, "F2F2F2")

            for prog, vals in sorted(by_type[ptype][year], key=lambda x: x[0]):
                row = table.add_row().cells
                _set_cell(row[0], serial)
                _set_cell(row[1], prog.title(), align="left")
                has_sp_sup = (vals["sp_registered"] + vals["sup_registered"]) > 0
                _fill_metric_cells(row, 2, vals, has_sp_sup)
                serial += 1

    return table


# ─────────────────────────────────────────────────────────────────────────────
# Table 2: Summary (one row per programme, collapsed across years)
# ─────────────────────────────────────────────────────────────────────────────

def build_summary_table(doc, agg):
    prog_sum = defaultdict(_zero)
    for (prog, year), vals in agg.items():
        p = prog_sum[prog]
        for k in _zero().keys():
            p[k] += vals[k]

    table = doc.add_table(rows=0, cols=N_SUM_COLS); _set_borders(table); _tight_margins(table)
    _add_header_rows(table, include_classified=True)

    sorted_progs = sorted(
        prog_sum.keys(),
        key=lambda p: (TYPE_ORDER.index(_prog_type(p)), p)
    )

    sno = 1
    # Accumulators for TOTAL row
    grand = _zero()
    grand_inc = grand_inc_sp = grand_inc_sup = 0
    grand_classified = 0
    has_any_sp = False

    for prog in sorted_progs:
        v = prog_sum[prog]
        row = table.add_row().cells
        _set_cell(row[0], sno)
        _set_cell(row[1], prog.title(), align="left")
        has_sp_sup = (v["sp_registered"] + v["sup_registered"]) > 0
        if has_sp_sup:
            has_any_sp = True
        col = _fill_metric_cells(row, 2, v, has_sp_sup)
        classified_val = v["classified"] if v["classified"] else "-"
        _set_cell(row[col], classified_val)

        # Accumulate grand totals
        for k in _zero().keys():
            grand[k] += v[k]
        inc_c, inc_s, inc_su = _compute_incomplete(v, has_sp_sup)
        grand_inc += inc_c
        if has_sp_sup:
            grand_inc_sp  += inc_s
            grand_inc_sup += inc_su
        else:
            grand_inc_sup += inc_c
        grand_classified += v["classified"]
        sno += 1

    # TOTAL row
    row = table.add_row().cells
    merged = row[0].merge(row[1])
    _set_cell(merged, "TOTAL", bold=True)

    col = 2
    for label, comb_key, sp_key, sup_key in ALL_METRICS:
        if label in ("Deregistered", "Disciplinary"):
            _set_cell(row[col],     "-", bold=True)
            _set_cell(row[col + 1], "-", bold=True)
        elif label == "Incomplete":
            _set_cell(row[col],     grand_inc_sp  if has_any_sp else "-",       bold=True)
            _set_cell(row[col + 1], grand_inc_sup if has_any_sp else grand_inc, bold=True)
        elif label == "Total":
            if has_any_sp:
                _set_cell(row[col],     grand["sp_registered"],  bold=True)
                _set_cell(row[col + 1], grand["sup_registered"], bold=True)
            else:
                _set_cell(row[col],     "-",              bold=True)
                _set_cell(row[col + 1], grand["total"],   bold=True)
        else:
            _set_cell(row[col],     grand.get(sp_key,   0) if has_any_sp else "-",                     bold=True)
            _set_cell(row[col + 1], grand.get(sup_key,  0) if has_any_sp else grand.get(comb_key, 0),  bold=True)
        col += 2

    _set_cell(row[col], grand_classified if grand_classified else "-", bold=True)
    return table


# ─────────────────────────────────────────────────────────────────────────────
# Main generator
# ─────────────────────────────────────────────────────────────────────────────

def generate(template_path, json_path, output_path):
    with open(json_path) as f:
        data = json.load(f)
    records     = data["records"]
    series      = data.get("series", "")
    ay          = data.get("academic_year", "")
    total_pages = data.get("total_pages", "?")
    gen_date    = date.today().strftime("%d %B %Y").lstrip("0")

    agg = aggregate(records)

    # ── Unpack template and replace placeholders ──────────────────────────────
    work_dir   = tempfile.mkdtemp()
    unpack_dir = os.path.join(work_dir, "unpacked")
    with zipfile.ZipFile(template_path, "r") as z:
        z.extractall(unpack_dir)

    doc_xml_path = os.path.join(unpack_dir, "word", "document.xml")
    with open(doc_xml_path, encoding="utf-8") as f:
        content = f.read()

    content = content.replace("{{ semister }}",       "SPECIAL/SUPPLEMENTARY")
    content = content.replace("{{ academic_year }}", _esc(ay))
    content = content.replace("{{ series }}",        _esc(series))
    content = content.replace("{{ total_pages }}",   str(total_pages))
    content = content.replace("{{ generated_date }}", _esc(gen_date))

    tables = _find_all_top_level_tables(content)
    assert len(tables) >= 3, f"Need >=3 tables in template, found {len(tables)}"

    t1_end   = tables[0][1]
    last_end = tables[-1][1]

    between      = content[t1_end: tables[1][0]]
    doc_c1_match = re.search(
        r'<w:p[^>]*>(?:(?!</w:p>).)*DOC\.\s*C[0-9][^<]*(?:(?!</w:p>).)*</w:p>\s*',
        between, re.DOTALL,
    )
    cut_at = t1_end + doc_c1_match.start() if doc_c1_match else tables[1][0]

    MARKER     = "____MAGIC_INSERTION____"
    marker_xml = f'<w:p><w:r><w:t>{MARKER}</w:t></w:r></w:p>'
    new_content = content[:cut_at] + marker_xml + content[last_end:]
    with open(doc_xml_path, "w", encoding="utf-8") as f:
        f.write(new_content)

    tmp_docx = os.path.join(work_dir, "tmp.docx")
    with (zipfile.ZipFile(template_path, "r") as zi,
          zipfile.ZipFile(tmp_docx, "w", zipfile.ZIP_DEFLATED) as zo):
        for item in zi.infolist():
            if item.filename == "word/document.xml":
                zo.write(doc_xml_path, "word/document.xml")
            else:
                zo.writestr(item, zi.read(item.filename))

    # ── Build python-docx tables and insert ───────────────────────────────────
    doc = Document(tmp_docx)
    insert_p = next(
        (p._p for p in doc.paragraphs if MARKER in p.text), None
    )

    def ins(el):
        if insert_p is not None:
            insert_p.addprevious(el)
        else:
            doc._body.append(el)

    ins(doc.add_paragraph()._p)

    # Column widths: S/No(1.0) + Name(4.6) + 9 metric groups x 2 x 0.68cm
    metric_w = [1.0] * (len(ALL_METRICS) * 2)
    t1 = build_main_table(doc, agg)
    _set_col_widths(t1, [0.8, 4.0] + metric_w)
    ins(t1._tbl)

    ins(doc.add_paragraph()._p)

    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rn = tp.add_run("Total no of students per School/Institute")
    rn.bold = True; rn.font.name = FONT; rn.font.size = Pt(9)
    ins(tp._p)

    # Summary table: +1 col for No. Classified
    t2 = build_summary_table(doc, agg)
    _set_col_widths(t2, [0.8, 4.0] + [0.9] * (len(ALL_METRICS) * 2) + [1.0])
    ins(t2._tbl)

    ins(doc.add_paragraph()._p)

    if insert_p is not None and insert_p.getparent() is not None:
        insert_p.getparent().remove(insert_p)

    doc.save(output_path)
    shutil.rmtree(work_dir)
    print(f"\nSaved: {output_path}")
    print(f"Records: {len(records)}   Programme-year rows: {len(agg)}")


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python generate.py <template.docx> <data.json> <output.docx>")
        sys.exit(1)
    generate(sys.argv[1], sys.argv[2], sys.argv[3])
