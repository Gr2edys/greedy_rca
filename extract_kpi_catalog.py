#!/usr/bin/env python3
"""
UNC 5GC KPI Catalog Extractor
==============================
Extract KPI counters from Huawei UNC HedEx product documentation and output an Excel catalog.

Input:  Path to the product documentation resources directory (containing navi.xml and KPI/).
Output: Excel (.xlsx) file with columns:
    KPI ID | NE | KPI名称 | 大类 | 子类 | 含义 | 计算公式

Scope: KPIs under the navi.xml topic:
    AMF/SMF/GW-C/NRF/NSSF/SMSF/NCG业务性能指标

Usage:
    python3 extract_kpi_catalog.py /path/to/resources/ [output.xlsx]
"""

import re
import json
import os
import sys
import csv
import zipfile
from collections import defaultdict, Counter


# ─── Configuration ───────────────────────────────────────────────────────────

# The combined NE topic name in navi.xml (UTF-8 decoded)
COMBINED_TOPIC = "AMF/SMF/GW-C/NRF/NSSF/SMSF/NCG业务性能指标"

# Output columns
HEADERS = ["KPI ID", "NE", "KPI名称", "大类", "子类", "含义", "计算公式"]

# Placeholder KPI IDs to skip (navigation nodes, not real counters)
SKIP_IDS = {"00000000", "00000001", "00000002", "00000003", "00000004",
            "00000005", "00000006", "00000007", "00000008", "00000009"}


# ─── Phase 1: Parse navi.xml ─────────────────────────────────────────────────

def parse_navi(resources_dir):
    """
    Parse navi.xml and extract all KPI counters under the combined NE topic.
    
    Returns a list of dicts:
        {kpi_id, kpi_name, ne, cat_path: [topic1, topic2, ...], url}
    """
    navi_path = os.path.join(resources_dir, "navi.xml")
    if not os.path.exists(navi_path):
        raise FileNotFoundError(f"navi.xml not found at {navi_path}")
    
    with open(navi_path, "rb") as f:
        raw = f.read()
    
    # Decode: try UTF-8, fallback to GB2312
    for enc in ["utf-8", "gb2312"]:
        try:
            content = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        content = raw.decode("utf-8", errors="replace")
    
    lines = content.split("\n")
    
    # ── Find the combined NE topic ──
    ne_start = None
    ne_indent = None
    for i, line in enumerate(lines):
        if COMBINED_TOPIC in line and "<topic" in line:
            ne_start = i
            ne_indent = len(line) - len(line.lstrip())
            break
    
    if ne_start is None:
        raise ValueError(f"Could not find topic '{COMBINED_TOPIC}' in navi.xml")
    
    print(f"  Found NE topic at line {ne_start}")
    
    # ── Walk the tree and collect KPI leaves ──
    stack = []          # [(indent, txt, url), ...]
    kpi_entries = []
    
    for i in range(ne_start + 1, len(lines)):
        line = lines[i]
        indent = len(line) - len(line.lstrip())
        
        # Exit when we close the NE topic
        if indent <= ne_indent and "</topic>" in line:
            break
        
        open_m = re.search(r'<topic\s+txt="([^"]*)"\s+url="([^"]*)"', line)
        close_m = "</topic>" in line
        
        if open_m:
            txt = open_m.group(1)
            url = open_m.group(2)
            
            # Maintain stack: pop until we're at the right depth
            while stack and stack[-1][0] >= indent:
                stack.pop()
            
            stack.append((indent, txt, url))
            
            # Check if this is a leaf KPI counter: KPI/<digits>.html
            if re.match(r"^KPI/\d+\.html$", url):
                kpi_id = re.search(r"KPI/(\d+)\.html", url).group(1)
                
                # Skip placeholder IDs
                if kpi_id in SKIP_IDS:
                    continue
                
                # Build category path from stack (exclude self, exclude "KPI" heading)
                cat_path = []
                for s_indent, s_txt, s_url in stack[:-1]:
                    if s_indent > ne_indent and s_txt != "KPI":
                        cat_path.append(s_txt)
                
                # Determine NE
                ne = determine_ne(txt, cat_path)
                
                kpi_entries.append({
                    "kpi_id": kpi_id,
                    "kpi_name": txt,
                    "ne": ne,
                    "cat_path": cat_path,
                    "url": url
                })
        
        elif close_m:
            if stack:
                stack.pop()
    
    return kpi_entries


def determine_ne(kpi_name, cat_path):
    """Determine which NE a KPI belongs to from its name and category context."""
    name = kpi_name
    
    # ── Explicit NE in KPI name ──
    if "SMSF" in name:
        return "SMSF"
    if "NSSF" in name:
        return "NSSF"
    if "NRF" in name:
        return "NRF"
    if "NCG" in name or "CHF" in name:
        return "NCG"
    if "PGW-C" in name or "SGW-C" in name or "S5/S8" in name or "S11" in name:
        return "GW-C"
    if "GGSN" in name:
        return "GW-C"
    
    # ── Category path hints ──
    for cat in cat_path:
        if "AMF" in cat:
            return "AMF"
        if "SMF" in cat:
            return "SMF"
        if "NRF" in cat:
            return "NRF"
        if "NSSF" in cat:
            return "NSSF"
        if "SMSF" in cat:
            return "SMSF"
        if "NCG" in cat or "CHF" in cat:
            return "NCG"
    
    # ── KPI name pattern matching ──
    if "N2模式" in name or "N11" in name or "N8" in name or "N12" in name:
        return "AMF"
    if "SMF" in name:
        return "SMF"
    if "N4" in name or "N7" in name or "N40" in name:
        return "SMF"
    
    # ── Counter ID range hints (approximate) ──
    kpi_id_str = name.split()[0] if name.split() else ""
    if kpi_id_str.isdigit():
        kpi_id = int(kpi_id_str)
        if 1929449000 <= kpi_id <= 1929480999:
            return "AMF"
        if 1929457000 <= kpi_id <= 1929470999:
            return "SMF"
        if 1929471000 <= kpi_id <= 1929474999:
            return "NRF"
        if 1929475000 <= kpi_id <= 1929477999:
            return "NSSF"
        if 1929478000 <= kpi_id <= 1929479999:
            return "SMSF"
        if 1929686000 <= kpi_id <= 1929686999:
            return "NCG"
    
    return "AMF"


# ─── Phase 2: Parse KPI HTML files ───────────────────────────────────────────

def parse_kpi_html(kpi_id, kpi_dir):
    """Extract 含义 (meaning) and 计算公式 (formula) from a KPI counter HTML file."""
    html_path = os.path.join(kpi_dir, f"{kpi_id}.html")
    
    if not os.path.exists(html_path):
        return {"meaning": "无", "formula": "无"}
    
    with open(html_path, "rb") as f:
        raw = f.read()
    
    # Decode: try GB2312 first (most common), then UTF-8
    for enc in ["gb2312", "utf-8", "gbk"]:
        try:
            text = raw.decode(enc, errors="ignore")
            break
        except UnicodeDecodeError:
            continue
    else:
        return {"meaning": "无", "formula": "无"}
    
    result = {"meaning": "无", "formula": "无"}
    
    # ── Extract 含义 (description) ──
    # Pattern: <div class="pmidesc"> ... <div class="pmidescbody"> TEXT </div>
    m = re.search(
        r'<div class="pmidesc">.*?<div class="pmidescbody">\s*(.*?)\s*</div>',
        text, re.DOTALL
    )
    if m:
        meaning = re.sub(r"<[^>]+>", "", m.group(1)).strip()
        if meaning and meaning not in ("无", "无。", ""):
            result["meaning"] = meaning
    
    # ── Extract 计算公式 (formula) ──
    # Pattern: <div class="pmiformula"> ... <div class="pmiformulabody"> TEXT </div>
    m = re.search(
        r'<div class="pmiformula">.*?<div class="pmiformulabody">\s*(.*?)\s*</div>',
        text, re.DOTALL
    )
    if m:
        formula = re.sub(r"<[^>]+>", "", m.group(1)).strip()
        if formula and formula not in ("无", "无。", "-", ""):
            result["formula"] = formula
    
    return result


# ─── Phase 3: Build category hierarchy ───────────────────────────────────────

def build_categories(kpi_entries):
    """Assign 大类 and 子类 from cat_path."""
    for entry in kpi_entries:
        cp = entry["cat_path"]
        if len(cp) >= 2:
            entry["大类"] = cp[0]
            entry["子类"] = " > ".join(cp[1:])
        elif len(cp) == 1:
            entry["大类"] = cp[0]
            entry["子类"] = ""
        else:
            entry["大类"] = "KPI"
            entry["子类"] = ""


# ─── Phase 4: Generate Excel (.xlsx) ─────────────────────────────────────────

def escape_xml(s):
    """Escape special XML characters and strip control characters."""
    if s is None:
        return ""
    s = str(s)
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    s = s.replace('"', "&quot;").replace("'", "&apos;")
    s = "".join(c for c in s if ord(c) >= 32 or c in "\n\r\t")
    return s


def write_xlsx(entries, output_path):
    """Write KPI entries to an .xlsx file using only Python stdlib (zipfile + XML)."""
    
    # Sort by NE then KPI ID
    entries_sorted = sorted(entries, key=lambda e: (e["ne"], int(e["kpi_id"])))
    
    # Build rows
    rows = [HEADERS]
    for e in entries_sorted:
        raw_name = e["kpi_name"]
        parts = raw_name.split(" ", 1)
        clean_name = parts[1] if len(parts) > 1 and parts[0].isdigit() else raw_name
        
        rows.append([
            e["kpi_id"], e["ne"], clean_name,
            e.get("大类", ""), e.get("子类", ""),
            e.get("meaning", "无"), e.get("formula", "无")
        ])
    
    # Shared strings table
    strings, smap = [], {}
    def _ss(s):
        s = str(s) if s else ""
        if s not in smap:
            smap[s] = len(strings)
            strings.append(s)
        return smap[s]
    
    for row in rows:
        for cell in row:
            _ss(cell)
    
    ss_xml_parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(strings)}" uniqueCount="{len(strings)}">'
    ]
    for s in strings:
        ss_xml_parts.append(f'<si><t xml:space="preserve">{escape_xml(s)}</t></si>')
    ss_xml_parts.append('</sst>')
    ss_xml = "\n".join(ss_xml_parts)
    
    # Styles
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="2">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF2F5496"/></patternFill></fill>'
        '</fills>'
        '<borders count="2">'
        '<border><left style="thin"><color rgb="FFD9D9D9"/></left>'
        '<right style="thin"><color rgb="FFD9D9D9"/></right>'
        '<top style="thin"><color rgb="FFD9D9D9"/></top>'
        '<bottom style="thin"><color rgb="FFD9D9D9"/></bottom></border>'
        '<border><left/><right/><top/><bottom/></border>'
        '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '</cellXfs>'
        '</styleSheet>'
    )
    
    # Sheet data
    cols = "ABCDEFG"
    sheet_row_strs = []
    for r, row in enumerate(rows):
        cells = []
        for c, val in enumerate(row):
            si = _ss(val)
            style = 2 if r == 0 else 0
            cells.append(f'<c r="{cols[c]}{r+1}" t="s" s="{style}"><v>{si}</v></c>')
        sheet_row_strs.append(f'<row r="{r+1}">{"".join(cells)}</row>')
    
    last_col = cols[len(HEADERS) - 1]
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews>'
        '<sheetView tabSelected="1" workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView>'
        '</sheetViews>'
        '<cols>'
        '<col min="1" max="1" width="15" customWidth="1"/>'
        '<col min="2" max="2" width="8" customWidth="1"/>'
        '<col min="3" max="3" width="55" customWidth="1"/>'
        '<col min="4" max="4" width="32" customWidth="1"/>'
        '<col min="5" max="5" width="42" customWidth="1"/>'
        '<col min="6" max="6" width="65" customWidth="1"/>'
        '<col min="7" max="7" width="32" customWidth="1"/>'
        '</cols>'
        f'<sheetData>{"".join(sheet_row_strs)}</sheetData>'
        f'<autoFilter ref="A1:{last_col}{len(rows)}"/>'
        '</worksheet>'
    )
    
    zip_entries = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            '</Relationships>'
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="UNC 5GC KPI Catalog" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
            '<Relationship Id="rId3" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
            'Target="sharedStrings.xml"/>'
            '</Relationships>'
        ),
        "xl/worksheets/sheet1.xml": sheet_xml,
        "xl/styles.xml": styles_xml,
        "xl/sharedStrings.xml": ss_xml,
    }
    
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in zip_entries.items():
            zf.writestr(name, content)


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("Usage: python3 extract_kpi_catalog.py <resources_dir> [output.xlsx]")
        sys.exit(1)
    
    resources_dir = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "unc_5gc_kpi_catalog.xlsx"
    kpi_dir = os.path.join(resources_dir, "KPI")
    
    if not os.path.isdir(resources_dir):
        print(f"ERROR: resources directory not found: {resources_dir}")
        sys.exit(1)
    if not os.path.isdir(kpi_dir):
        print(f"ERROR: KPI directory not found: {kpi_dir}")
        sys.exit(1)
    
    # ── Phase 1: Parse navi.xml ──
    print("Step 1/4: Parsing navi.xml KPI hierarchy...")
    entries = parse_navi(resources_dir)
    print(f"  Found {len(entries)} KPI counters")
    
    ne_counts = Counter(e["ne"] for e in entries)
    for ne, count in sorted(ne_counts.items(), key=lambda x: -x[1]):
        print(f"    {ne}: {count}")
    
    # ── Phase 2: Parse KPI HTML files ──
    print("\nStep 2/4: Extracting 含义 and 计算公式 from KPI HTML files...")
    for i, e in enumerate(entries):
        if (i + 1) % 1000 == 0:
            print(f"  {i+1}/{len(entries)}...")
        details = parse_kpi_html(e["kpi_id"], kpi_dir)
        e["meaning"] = details["meaning"]
        e["formula"] = details["formula"]
    print(f"  Done. Processed {len(entries)} files.")
    
    # Coverage stats
    with_meaning = sum(1 for e in entries if e.get("meaning") and e["meaning"] != "无")
    with_formula = sum(1 for e in entries if e.get("formula") and e["formula"] != "无")
    print(f"  含义: {with_meaning}/{len(entries)}  |  公式: {with_formula}/{len(entries)}")
    
    # ── Phase 3: Build categories ──
    print("\nStep 3/4: Building 大类/子类 hierarchy...")
    build_categories(entries)
    
    # Show top categories per NE
    for ne in sorted(ne_counts):
        ne_entries = [e for e in entries if e["ne"] == ne]
        cat_counts = Counter(e.get("大类", "") for e in ne_entries)
        top = cat_counts.most_common(3)
        cats_str = ", ".join(f"{c}({n})" for c, n in top)
        print(f"  {ne}: {cats_str}")
    
    # ── Phase 4: Generate Excel ──
    print(f"\nStep 4/4: Writing Excel to {output_path}...")
    write_xlsx(entries, output_path)
    
    print(f"\n{'='*60}")
    print(f"DONE: {len(entries)} KPIs written to {output_path}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
