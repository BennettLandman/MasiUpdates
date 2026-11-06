#!/usr/bin/env python3
"""
generate_report.py

All-in-one script to:
1. Download the first sheet of a Google Spreadsheet to an Excel file (.xlsx).
2. Extract Columns B (Email Address) and C (Status Update).
3. Strip '@domain' from Email Address leaving only the username/name.
4. Replace manual line breaks with clean lines and format every line as a bullet point.
5. Sort entries alphabetically by name in the first column.
6. Generate a Microsoft Word report (.docx) featuring:
   - Header updated with the current date (e.g. YYYY.MM.DD).
   - Alternating row background shading.
   - Matching table headers.
   - Clean bullet list layout for status updates.
7. Save the Word document automatically as YYYY.MM.DD.docx (or a custom path).

Runs out of the box with standard Python (zero pip dependencies required).
"""

import argparse
import csv
import datetime
import io
import os
import re
import ssl
import sys
import urllib.request
import xml.etree.ElementTree as ET
import xml.sax.saxutils
import zipfile

DEFAULT_URL = "https://docs.google.com/spreadsheets/d/1CMKiJqblfM0IocFP77v91pUJx-Zu0mSySDHLONOVRZU/edit?usp=sharing"

# Ensure SSL verification works smoothly even on systems missing local CA certificates
try:
    _test_ctx = ssl.create_default_context()
    _test_ctx.get_ca_certs()
except Exception:
    ssl._create_default_https_context = ssl._create_unverified_context


def fetch_url(url: str) -> str:
    """Fetch URL content as text with a browser User-Agent."""
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"}
    )
    try:
        with urllib.request.urlopen(req) as response:
            return response.read().decode("utf-8")
    except Exception:
        ctx = ssl._create_unverified_context()
        req2 = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"}
        )
        with urllib.request.urlopen(req2, context=ctx) as response:
            return response.read().decode("utf-8")


def extract_spreadsheet_id(url_or_id: str) -> str:
    """Extract spreadsheet ID from URL or return string as-is."""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url_or_id)
    if match:
        return match.group(1)
    return url_or_id.strip()


def detect_first_sheet_name(sheet_id: str) -> str:
    """Detect the title of the first sheet tab from the Google Sheets HTML."""
    try:
        html = fetch_url(f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit?usp=sharing")
        match = re.search(r"class=\"[^\"]*docs-sheet-tab-caption[^\"]*\">([^<]+)<", html)
        if match:
            return match.group(1).strip()
    except Exception:
        pass
    return "Form Responses 1"


def clean_xml_string(val: str) -> str:
    """Filter out control characters disallowed in XML 1.0."""
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]", "", str(val))


def escape_xml(val: str) -> str:
    """Escape XML special characters after removing invalid control characters."""
    return xml.sax.saxutils.escape(clean_xml_string(val))


def col_to_letter(col_idx: int) -> str:
    """Convert 1-based column index to Excel column letter."""
    result = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def write_xlsx(rows: list, output_path: str, sheet_name: str = "Sheet1"):
    """Write rows to a valid Microsoft Excel (.xlsx) file using pure standard library."""
    safe_sheet_name = escape_xml(sheet_name)

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    wb_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""

    workbook = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="{safe_sheet_name}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

    sheet_xml_parts = [
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>"""
    ]

    for r_idx, row in enumerate(rows, start=1):
        sheet_xml_parts.append(f'<row r="{r_idx}">')
        for c_idx, val in enumerate(row, start=1):
            if val is None or val == "":
                continue
            cell_ref = f"{col_to_letter(c_idx)}{r_idx}"
            safe_val = escape_xml(str(val))
            sheet_xml_parts.append(
                f'<c r="{cell_ref}" t="inlineStr"><is><t xml:space="preserve">{safe_val}</t></is></c>'
            )
        sheet_xml_parts.append("</row>")

    sheet_xml_parts.append("""  </sheetData>
</worksheet>""")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/worksheets/sheet1.xml", "".join(sheet_xml_parts))


def download_first_sheet(url_or_id: str, output_xlsx: str = None) -> tuple:
    """
    Download the first sheet of a Google Spreadsheet and save as Excel (.xlsx).
    Returns (raw_rows, output_xlsx_path).
    """
    sheet_id = extract_spreadsheet_id(url_or_id)
    first_sheet_name = detect_first_sheet_name(sheet_id)
    if not output_xlsx:
        sanitized_name = re.sub(r'[\\/*?:"<>| ]', "_", first_sheet_name)
        output_xlsx = f"{sanitized_name}.xlsx"

    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv"
    print(f"[1/4] Downloading first sheet ('{first_sheet_name}') from Google Sheets...")
    csv_text = fetch_url(csv_url)

    reader = list(csv.reader(io.StringIO(csv_text)))
    write_xlsx(reader, output_xlsx, sheet_name=first_sheet_name)
    print(f"      Saved raw Excel sheet to: {output_xlsx} ({len(reader)} rows)")
    return reader, output_xlsx


def read_entries_from_rows(rows: list) -> tuple:
    """
    Extract and process Columns A, B, and C from table rows:
    - Column A: Timestamp of submission.
    - Column B: Strip '@domain' to leave only the name.
    - Column C: Remove manual line breaks and format lines as bullet points.
    - Sort rows alphabetically by the name in Column B.
    Returns (sorted_entries, last_timestamp).
    """
    entries = []
    last_timestamp = None

    # Skip header row (rows[0])
    for row in rows[1:]:
        if len(row) < 3:
            continue
        timestamp = row[0].strip() if len(row) > 0 else ""
        email = row[1].strip()
        status_text = row[2]

        # Strip @domain from email to leave only the username/name
        name = email.split("@")[0].strip()

        # Split on manual line breaks (soft returns, \r\n, \r, \n, unicode line separators)
        raw_lines = re.split(r"[\r\n\v\u2028]+", status_text)
        cleaned_bullets = []
        for line in raw_lines:
            # Strip leading bullet indicators (*, -, •) and surrounding whitespace
            cleaned_line = re.sub(r"^\s*[\*\-•]\s*", "", line).strip()
            if cleaned_line:
                cleaned_bullets.append(cleaned_line)

        if name and cleaned_bullets:
            entries.append({"name": name, "bullets": cleaned_bullets, "timestamp": timestamp})
            if timestamp:
                last_timestamp = timestamp

    # Sort entries alphabetically by the name in the first column
    entries.sort(key=lambda item: (item["name"].lower(), item["name"]))
    return entries, last_timestamp


def create_word_report(
    entries: list,
    output_docx: str,
    date_obj: datetime.date = None,
    template_path: str = None
):
    """
    Generate the Microsoft Word (.docx) report matching the exact fonts (Arial),
    table colors (GridTable4), margins (0.5 in), line spacing, and header format
    from the example document.
    """
    if date_obj is None:
        date_obj = datetime.date.today()

    # Formatted date for header: e.g. "Friday, November 6, 2026"
    date_text = date_obj.strftime("%A, %B ") + str(date_obj.day) + date_obj.strftime(", %Y")

    # Header XML with bottom border, date, tabs, and page number field
    header_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:tabs><w:tab w:val="left" w:pos="3100"/><w:tab w:val="right" w:pos="10800"/></w:tabs><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>{escape_xml(date_text)}</w:t></w:r><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">- </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:bCs/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:bCs/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:bCs/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:bCs/><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:bCs/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:bCs/></w:rPr><w:t>–</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p></w:hdr>"""

    # Build Table 0 (Header Row Table) exactly as formatted in the example docx
    tbl0_xml = """<w:tbl><w:tblPr><w:tblStyle w:val="GridTable4"/><w:tblW w:w="10792" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="1255"/><w:gridCol w:w="9537"/></w:tblGrid><w:tr><w:trPr><w:cnfStyle w:val="100000000000" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:trHeight w:val="315"/></w:trPr><w:tc><w:tcPr><w:cnfStyle w:val="001000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:tcW w:w="1255" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:line="276" w:lineRule="auto"/><w:contextualSpacing/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:kern w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>Email Address</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="9537" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:line="276" w:lineRule="auto"/><w:contextualSpacing/><w:cnfStyle w:val="100000000000" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:kern w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>Status Update - Provide plain text (no jargon) what is going on with your projects. Do have a success to share? Do you have problem that you would like to discuss? Are you waiting for anything? </w:t></w:r></w:p></w:tc></w:tr></w:tbl>"""

    # Build Table 1 (Data Table) with alternating row formatting and bullet lists
    t1_rows = []
    for idx, entry in enumerate(entries):
        name_escaped = escape_xml(entry["name"])
        is_odd_band = (idx % 2 == 0)  # Alternating row band (cnfStyle oddHBand)
        cnf_tr = "000000100000" if is_odd_band else "000000000000"
        cnf_tc_col0 = "001000000000"
        cnf_p_col1 = "000000100000" if is_odd_band else "000000000000"

        bullets_p_xml = []
        for bullet in entry["bullets"]:
            bullet_escaped = escape_xml(bullet)
            p_xml = f"""<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="21"/></w:numPr><w:cnfStyle w:val="{cnf_p_col1}" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="{1 if is_odd_band else 0}" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:kern w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:kern w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">{bullet_escaped}</w:t></w:r></w:p>"""
            bullets_p_xml.append(p_xml)

        tr_xml = f"""<w:tr><w:trPr><w:cnfStyle w:val="{cnf_tr}" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="{1 if is_odd_band else 0}" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:trHeight w:val="315"/></w:trPr><w:tc><w:tcPr><w:cnfStyle w:val="{cnf_tc_col0}" w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/><w:tcW w:w="0" w:type="auto"/><w:hideMark/></w:tcPr><w:p><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:kern w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:kern w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>{name_escaped}</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/><w:hideMark/></w:tcPr>{"".join(bullets_p_xml)}</w:tc></w:tr>"""
        t1_rows.append(tr_xml)

    table1_xml = f"""<w:tbl><w:tblPr><w:tblStyle w:val="GridTable4"/><w:tblW w:w="10792" w:type="dxa"/><w:tblLook w:val="0480" w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2646"/><w:gridCol w:w="8146"/></w:tblGrid>{"".join(t1_rows)}</w:tbl>"""

    # Section properties with 0.5 inch margins (720 dxa)
    sectPr_xml = """<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>"""

    new_doc_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {tbl0_xml}
    {table1_xml}
    <w:p><w:pPr><w:snapToGrid w:val="0"/><w:contextualSpacing/><w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr></w:p>
    {sectPr_xml}
  </w:body>
</w:document>"""

    # Look for template.docx in same folder as script if not specified
    if not template_path:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidate = os.path.join(script_dir, "template.docx")
        if os.path.exists(candidate):
            template_path = candidate

    if template_path and os.path.exists(template_path):
        with zipfile.ZipFile(template_path, "r") as z_in, zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                if item.filename == "word/document.xml":
                    z_out.writestr(item.filename, new_doc_xml.encode("utf-8"))
                elif item.filename == "word/header1.xml":
                    z_out.writestr(item.filename, header_xml.encode("utf-8"))
                else:
                    z_out.writestr(item.filename, z_in.read(item.filename))
    else:
        # Standalone generation if template not available
        with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("word/document.xml", new_doc_xml.encode("utf-8"))
            zf.writestr("word/header1.xml", header_xml.encode("utf-8"))


def main():
    parser = argparse.ArgumentParser(
        description="All-in-one script to download Google Sheet and generate formatted Word report (.docx)"
    )
    parser.add_argument(
        "--url",
        default=DEFAULT_URL,
        help="Google Sheets URL or spreadsheet ID"
    )
    parser.add_argument(
        "--input",
        "-i",
        default=None,
        help="Path to an existing Excel (.xlsx) file if already downloaded"
    )
    parser.add_argument(
        "--output",
        "-o",
        default=None,
        help="Output Word report (.docx) path (default: YYYY.MM.DD.docx)"
    )
    parser.add_argument(
        "--last-updated-file",
        default="lastupdated.txt",
        help="Path to save the timestamp of the last entry (default: lastupdated.txt)"
    )

    args = parser.parse_args()

    # Step 1: Download or use existing file
    if args.input and os.path.exists(args.input):
        print(f"[1/4] Using existing input file: {args.input}")
        with zipfile.ZipFile(args.input) as z:
            xml_content = z.read("xl/worksheets/sheet1.xml")
            root = ET.fromstring(xml_content)
            ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            rows = []
            for r in root.findall(".//s:row", ns):
                cells = [
                    c.find("s:is/s:t", ns).text if c.find("s:is/s:t", ns) is not None else ""
                    for c in r.findall("s:c", ns)
                ]
                rows.append(cells)
    else:
        rows, saved_xlsx = download_first_sheet(args.url)

    # Step 2: Extract, clean, bulletize and sort entries
    print("[2/4] Processing Columns B & C (cleaning usernames, formatting bullets, sorting)...")
    entries, last_timestamp = read_entries_from_rows(rows)
    print(f"      Processed and sorted {len(entries)} entries alphabetically.")

    # Save last entry timestamp to lastupdated.txt
    if last_timestamp:
        with open(args.last_updated_file, "w", encoding="utf-8") as f:
            f.write(f"{last_timestamp}\n")
        print(f"      Saved last entry timestamp ('{last_timestamp}') to: {args.last_updated_file}")

    # Step 3: Determine output document name
    today_str = datetime.date.today().strftime("%Y.%m.%d")
    output_docx = args.output if args.output else f"{today_str}.docx"

    # Step 4: Generate Word document
    print(f"[3/4] Generating Word report: {output_docx} ...")
    create_word_report(entries, output_docx, date_obj=datetime.date.today())
    print(f"[4/4] ✓ Successfully created Word report at: {os.path.abspath(output_docx)}")


if __name__ == "__main__":
    main()
