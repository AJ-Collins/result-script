"""
TUM Results Report Generator
Usage: python generate_results.py <template.docx> <data.json> [output.docx]
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

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def title_case(text):
    return text.title()

def programme_type(name):
    n = name.upper()
    if n.startswith("BACHELOR"):    return "Degree"
    if n.startswith("DIPLOMA"):     return "Diploma"
    if n.startswith("CERTIFICATE"): return "Certificate"
    return "Other"

def ordinal(n):
    return {1: "ONE", 2: "TWO", 3: "THREE", 4: "FOUR"}.get(n, str(n))

def e(text):
    return xml_escape(str(text))

# ─────────────────────────────────────────────────────────────────────────────
# NESTED-AWARE TABLE BOUNDARY FINDER
# ─────────────────────────────────────────────────────────────────────────────

def find_table_end(content, tbl_start):
    """
    Return the index just after the closing </w:tbl> that matches the
    <w:tbl> opening at tbl_start, correctly handling nested <w:tbl> tags.
    """
    depth = 0
    pos   = tbl_start
    while pos < len(content):
        no = content.find('<w:tbl>',  pos)
        nc = content.find('</w:tbl>', pos)
        if nc == -1:
            raise ValueError("Malformed XML: unmatched <w:tbl>")
        if no != -1 and no < nc:
            depth += 1
            pos = no + 7
        else:
            depth -= 1
            pos = nc + 8
            if depth == 0:
                return pos
    raise ValueError("Malformed XML: could not find matching </w:tbl>")


def find_all_top_level_tables(content):
    """
    Walk the document and return a list of (start, end) for every TOP-LEVEL
    <w:tbl> block.  Nested tables are skipped — their start is consumed by
    find_table_end and never returned.
    """
    tables = []
    pos    = 0
    while True:
        start = content.find('<w:tbl>', pos)
        if start == -1:
            break
        end = find_table_end(content, start)
        tables.append((start, end))
        pos = end
    return tables

# ─────────────────────────────────────────────────────────────────────────────
# XML BUILDING BLOCKS
# ─────────────────────────────────────────────────────────────────────────────

def make_run(text, bold=True):
    sz    = "18" if bold else "16"
    b_tag = '<w:b w:val="1"/><w:bCs w:val="1"/>' if bold else ""
    return f"""            <w:r>
              <w:rPr>
                <w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/>
                {b_tag}
                <w:sz w:val="{sz}"/>
                <w:szCs w:val="{sz}"/>
                <w:rtl w:val="0"/>
              </w:rPr>
              <w:t xml:space="preserve">{e(text)}</w:t>
            </w:r>"""

# ─────────────────────────────────────────────────────────────────────────────
# TABLE 2
# ─────────────────────────────────────────────────────────────────────────────

T2_TABLE_OPEN = """\
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="Table2"/>
        <w:tblW w:w="14318" w:type="dxa"/>
        <w:jc w:val="left"/>
        <w:tblInd w:w="-572" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:left w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:bottom w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:right w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:insideH w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:insideV w:color="000000" w:space="0" w:sz="4" w:val="single"/>
        </w:tblBorders>
        <w:tblLayout w:type="fixed"/>
        <w:tblLook w:val="0400"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="710"/>
        <w:gridCol w:w="4960"/>
        <w:gridCol w:w="1277"/>
        <w:gridCol w:w="991"/>
        <w:gridCol w:w="994"/>
        <w:gridCol w:w="707"/>
        <w:gridCol w:w="994"/>
        <w:gridCol w:w="1417"/>
        <w:gridCol w:w="1274"/>
        <w:gridCol w:w="994"/>
      </w:tblGrid>"""

T2_HEADER_ROW = """\
      <w:tr>
        <w:trPr><w:cantSplit w:val="0"/><w:tblHeader w:val="1"/></w:trPr>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>S/No</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>Programme Name</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Registered</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Absent</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Sat Exam</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Pass</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Fail</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>Disciplinary</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>Incomplete</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>Total</w:t></w:r>
          </w:p></w:tc>
      </w:tr>"""

def t2_group_header_row(ptype, year):
    label = e(f"{ptype}: Year of Study - {year}")
    return f"""\
      <w:tr>
        <w:trPr><w:cantSplit w:val="0"/><w:tblHeader w:val="0"/></w:trPr>
        <w:tc>
          <w:tcPr><w:gridSpan w:val="10"/><w:shd w:fill="auto" w:val="clear"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000">
            <w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr>
              <w:t xml:space="preserve">{label}</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>"""

def t2_data_row(serial, prog_name, total, absent, sat, passed, failed):
    inc = max(0, total - sat - absent)
    def cell(val):
        return f"""        <w:tc>
          <w:tcPr><w:shd w:fill="auto" w:val="clear"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000">
            <w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr>
              <w:t xml:space="preserve">{e(val)}</w:t>
            </w:r>
          </w:p>
        </w:tc>"""

    name_cell = f"""        <w:tc>
          <w:tcPr><w:shd w:fill="auto" w:val="clear"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000">
            <w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr>
              <w:t xml:space="preserve">{e(title_case(prog_name))}</w:t>
            </w:r>
          </w:p>
        </w:tc>"""

    return f"""\
      <w:tr>
        <w:trPr><w:cantSplit w:val="0"/></w:trPr>
{cell(str(serial))}
{name_cell}
{cell(str(total))}
{cell(str(absent))}
{cell(str(sat))}
{cell(str(passed))}
{cell(str(failed))}
{cell("-")}
{cell(str(inc))}
{cell(str(total))}
      </w:tr>"""

# ─────────────────────────────────────────────────────────────────────────────
# TABLE 3
# ─────────────────────────────────────────────────────────────────────────────

T3_TABLE_OPEN = """\
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="Table3"/>
        <w:tblW w:w="14317" w:type="dxa"/>
        <w:jc w:val="left"/>
        <w:tblInd w:w="-572" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:left w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:bottom w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:right w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:insideH w:color="000000" w:space="0" w:sz="4" w:val="single"/>
          <w:insideV w:color="000000" w:space="0" w:sz="4" w:val="single"/>
        </w:tblBorders>
        <w:tblLayout w:type="fixed"/>
        <w:tblLook w:val="0400"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="709"/>
        <w:gridCol w:w="4820"/>
        <w:gridCol w:w="1254"/>
        <w:gridCol w:w="872"/>
        <w:gridCol w:w="943"/>
        <w:gridCol w:w="616"/>
        <w:gridCol w:w="567"/>
        <w:gridCol w:w="1418"/>
        <w:gridCol w:w="1275"/>
        <w:gridCol w:w="705"/>
        <w:gridCol w:w="1138"/>
      </w:tblGrid>"""

T3_TITLE_ROW = """\
      <w:tr>
        <w:trPr><w:cantSplit w:val="0"/></w:trPr>
        <w:tc>
          <w:tcPr><w:gridSpan w:val="11"/><w:shd w:fill="auto" w:val="clear"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000">
            <w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr>
              <w:t>Total no of students per School/Institute</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>"""

T3_HEADER_ROW = """\
      <w:tr>
        <w:trPr><w:cantSplit w:val="0"/><w:tblHeader w:val="1"/></w:trPr>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>S/No</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>Programme Name</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Registered</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Absent</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Sat Exam</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>N0. Pass</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Fail</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>Disciplinary</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>Incomplete</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>Total</w:t></w:r>
          </w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:fill="auto" w:val="clear"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000"><w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr><w:t>No. Classified</w:t></w:r>
          </w:p></w:tc>
      </w:tr>"""

def t3_data_row(serial, prog_name, registered, absent, sat, passed, failed, incomplete, total):
    def cell(val, align="center"):
        return f"""        <w:tc>
          <w:tcPr><w:shd w:fill="auto" w:val="clear"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000">
            <w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="{align}"/>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr>
              <w:t xml:space="preserve">{e(val)}</w:t>
            </w:r>
          </w:p>
        </w:tc>"""

    name_cell = f"""        <w:tc>
          <w:tcPr><w:shd w:fill="auto" w:val="clear"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000">
            <w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr>
              <w:t xml:space="preserve">{e(title_case(prog_name))}</w:t>
            </w:r>
          </w:p>
        </w:tc>"""

    return f"""\
      <w:tr>
        <w:trPr><w:cantSplit w:val="0"/></w:trPr>
{cell(str(serial))}
{name_cell}
{cell(str(registered))}
{cell(str(absent))}
{cell(str(sat))}
{cell(str(passed))}
{cell(str(failed))}
{cell("-")}
{cell(str(incomplete))}
{cell(str(total))}
{cell("-")}
      </w:tr>"""

def t3_total_row(registered, absent, sat, passed, failed, incomplete, total):
    def cell(val, bold=True):
        b = '<w:b w:val="1"/><w:bCs w:val="1"/>' if bold else ""
        return f"""        <w:tc>
          <w:tcPr><w:shd w:fill="auto" w:val="clear"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000">
            <w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/>{b}<w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr>
              <w:t xml:space="preserve">{e(val)}</w:t>
            </w:r>
          </w:p>
        </w:tc>"""

    label_cell = f"""        <w:tc>
          <w:tcPr><w:shd w:fill="auto" w:val="clear"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p w:rsidR="00000000" w14:paraId="00000000">
            <w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl w:val="0"/></w:rPr>
              <w:t>TOTAL</w:t>
            </w:r>
          </w:p>
        </w:tc>"""

    return f"""\
      <w:tr>
        <w:trPr><w:cantSplit w:val="0"/></w:trPr>
{cell("")}
{label_cell}
{cell(str(registered))}
{cell(str(absent))}
{cell(str(sat))}
{cell(str(passed))}
{cell(str(failed))}
{cell("-")}
{cell(str(incomplete))}
{cell(str(total))}
{cell("-")}
      </w:tr>"""

# ─────────────────────────────────────────────────────────────────────────────
# PAGE BREAK / DOC C1
# ─────────────────────────────────────────────────────────────────────────────

PAGE_BREAK_PARA = """\
    <w:p w:rsidR="00000000" w14:paraId="00000000">
      <w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:rPr><w:rtl w:val="0"/></w:rPr>
        <w:br w:type="page"/>
      </w:r>
    </w:p>"""

def doc_c1_para(total_pages, semester_label):
    text = e(f"DOC. C1: PAGES 1 \u2013 {total_pages} - SEMESTER {semester_label} ")
    return f"""\
    <w:p w:rsidR="00000000" w14:paraId="00000000">
      <w:pPr><w:jc w:val="center"/>
        <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/></w:rPr>
      </w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Book Antiqua" w:cs="Book Antiqua" w:eastAsia="Book Antiqua" w:hAnsi="Book Antiqua"/><w:b w:val="1"/><w:bCs w:val="1"/><w:rtl w:val="0"/></w:rPr>
        <w:t xml:space="preserve">{text}</w:t>
      </w:r>
    </w:p>"""

# ─────────────────────────────────────────────────────────────────────────────
# SEMESTER SECTION BUILDER
# ─────────────────────────────────────────────────────────────────────────────

ORDINARY_TYPE = "semester ordinary mark sheet"

def build_semester_xml(semester, records, total_pages):
    # Only Semester Ordinary Mark Sheets for this semester
    ordinary = [
        r for r in records
        if r["sheet_type"].lower() == ORDINARY_TYPE
        and r["semester"] == semester
    ]
    if not ordinary:
        return ""

    # Aggregate by (programme, year) — multiple rows = separate intakes, sum them
    agg = defaultdict(lambda: {"pass": 0, "fail": 0, "absent": 0, "sat_exam": 0, "total": 0})
    for rec in ordinary:
        key = (rec["programme_name"], rec["year_of_study"])
        agg[key]["pass"]     += rec.get("pass", 0)
        agg[key]["fail"]     += rec.get("fail", 0)
        agg[key]["absent"]   += rec.get("absent", 0)
        agg[key]["sat_exam"] += rec.get("sat_exam", 0)
        agg[key]["total"]    += rec.get("total", 0)

    flat = [{"programme_name": p, "year_of_study": y, **v} for (p, y), v in agg.items()]

    groups = defaultdict(lambda: defaultdict(list))
    for rec in flat:
        groups[programme_type(rec["programme_name"])][rec["year_of_study"]].append(rec)

    # Table 2
    t2_rows = [T2_HEADER_ROW]
    serial = 1
    for ptype in ["Degree", "Diploma", "Certificate", "Other"]:
        if ptype not in groups:
            continue
        for year in sorted(groups[ptype].keys(), reverse=True):
            t2_rows.append(t2_group_header_row(ptype, year))
            for rec in sorted(groups[ptype][year], key=lambda x: x["programme_name"]):
                t2_rows.append(t2_data_row(serial, rec["programme_name"], rec["total"],
                                           rec["absent"], rec["sat_exam"], rec["pass"], rec["fail"]))
                serial += 1
    table2_xml = T2_TABLE_OPEN + "\n" + "\n".join(t2_rows) + "\n    </w:tbl>"

    # Table 3
    prog_sum = defaultdict(lambda: {"reg":0,"abs":0,"sat":0,"pass":0,"fail":0,"inc":0,"tot":0})
    for rec in flat:
        p = rec["programme_name"]
        prog_sum[p]["reg"]  += rec["total"]
        prog_sum[p]["abs"]  += rec["absent"]
        prog_sum[p]["sat"]  += rec["sat_exam"]
        prog_sum[p]["pass"] += rec["pass"]
        prog_sum[p]["fail"] += rec["fail"]
        prog_sum[p]["inc"]  += max(0, rec["total"] - rec["sat_exam"] - rec["absent"])
        prog_sum[p]["tot"]  += rec["total"]

    t3_rows = [T3_TITLE_ROW, T3_HEADER_ROW]
    sno = 1
    sr = sa = ss = sp = sf = si = st = 0
    for prog in sorted(prog_sum.keys()):
        v = prog_sum[prog]
        t3_rows.append(t3_data_row(sno, prog, v["reg"], v["abs"], v["sat"],
                                   v["pass"], v["fail"], v["inc"], v["tot"]))
        sr+=v["reg"]; sa+=v["abs"]; ss+=v["sat"]; sp+=v["pass"]
        sf+=v["fail"]; si+=v["inc"]; st+=v["tot"]
        sno += 1
    t3_rows.append(t3_total_row(sr, sa, ss, sp, sf, si, st))
    table3_xml = T3_TABLE_OPEN + "\n" + "\n".join(t3_rows) + "\n    </w:tbl>"

    sem_label = ordinal(semester)
    return doc_c1_para(total_pages, sem_label) + "\n" + table2_xml + "\n" + table3_xml


# ─────────────────────────────────────────────────────────────────────────────
# MAIN GENERATOR
# ─────────────────────────────────────────────────────────────────────────────

def generate(template_path, json_path, output_path):
    with open(json_path) as f:
        raw = json.load(f)

    records     = raw if isinstance(raw, list) else raw.get("records", [])
    total_pages = raw.get("total_pages", "?") if isinstance(raw, dict) else "?"

    semesters = sorted(set(
        r["semester"] for r in records
        if r["sheet_type"].lower() == ORDINARY_TYPE
    ))

    meta_rec      = next((r for r in records if r.get("academic_year")), {})
    academic_year = meta_rec.get("academic_year", "")
    series        = meta_rec.get("series", "")
    gen_date      = date.today().strftime("%d %B %Y").lstrip("0")

    work_dir   = tempfile.mkdtemp()
    unpack_dir = os.path.join(work_dir, "unpacked")
    with zipfile.ZipFile(template_path, "r") as z:
        z.extractall(unpack_dir)

    doc_xml_path = os.path.join(unpack_dir, "word", "document.xml")
    content = open(doc_xml_path, encoding="utf-8").read()

    # Simple placeholder replacements
    content = content.replace("{{ semister }}", f"SEMESTER {ordinal(semesters[0]) if semesters else 'ONE'}")
    content = content.replace("{{ academic_year }}", e(academic_year))
    content = content.replace("{{ series }}", e(series))
    content = content.replace("{{ total_pages }}", str(total_pages))
    content = content.replace("{{ generated_date }}", e(gen_date))

    # ── Find ALL top-level tables ─────────────────────────────────────────────
    # The template may already contain pre-filled semester sections (multiple
    # Table2/Table3 pairs).  We ONLY keep Table1 (the header info table) and
    # the document footer that follows the very last table.  Everything in
    # between — all semester data — is replaced with freshly generated content.
    tables = find_all_top_level_tables(content)
    assert len(tables) >= 3, f"Expected at least 3 tables in template, found {len(tables)}"

    # Table[0] = Table1 (header).  Its end is our first anchor.
    t1_end    = tables[0][1]

    # Last table's end = where the document footer/signature section begins.
    last_end  = tables[-1][1]

    # The "before" slice: everything up to (but not including) the first DOC.C1
    # paragraph that precedes the first data table, so we don't get a stray
    # DOC.C1 line sitting above our generated content.
    chunk_between = content[t1_end : tables[1][0]]
    doc_c1_match  = re.search(
        r'<w:p[^>]*>(?:(?!</w:p>).)*DOC\.\s*C[0-9][^<]*(?:(?!</w:p>).)*</w:p>\s*',
        chunk_between, re.DOTALL
    )
    if doc_c1_match:
        cut_at = t1_end + doc_c1_match.start()
    else:
        cut_at = tables[1][0]   # fallback: cut right at the first data table

    before_data = content[:cut_at]
    after_data  = content[last_end:]

    # ── Generate one section per semester ─────────────────────────────────────
    sem_sections = [s for s in
                    (build_semester_xml(sem, records, total_pages) for sem in semesters)
                    if s]

    semesters_xml = ("\n" + PAGE_BREAK_PARA + "\n").join(sem_sections)

    new_content = before_data + semesters_xml + after_data

    with open(doc_xml_path, "w", encoding="utf-8") as f:
        f.write(new_content)

    with (zipfile.ZipFile(template_path, "r") as zin,
          zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout):
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.write(doc_xml_path, "word/document.xml")
            else:
                zout.writestr(item, zin.read(item.filename))

    shutil.rmtree(work_dir)
    print(f"✅  Saved: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generate_results.py <template.docx> <data.json> [output.docx]")
        sys.exit(1)
    template = sys.argv[1]
    data     = sys.argv[2]
    output   = sys.argv[3] if len(sys.argv) > 3 else "ICI_RESULTS.docx"
    generate(template, data, output)
