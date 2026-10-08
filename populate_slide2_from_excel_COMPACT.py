#!/usr/bin/env python3
r"""
Credit Risk Metrics slide 2 update with enforced compact summary height.

Reads the latest counts from the existing generated Excel workbook, then
updates ONLY the native top-right PowerPoint summary table in the ORIGINAL
Final Deck.pptx.  Previous-quarter cells are blank and the latest-quarter
cells become 3Q26.

Edits the 8-column native summary table only: replaces text and CORRECTS
its 6 row heights to prevent overlap with the detailed table beneath it.
No table reconstruction, no column merging, no font/color/layout changes
elsewhere. The original is never overwritten.

Windows usage (run from the Scripts folder):
    python populate_slide2_from_excel_COMPACT.py

Dependencies: openpyxl, lxml (lxml is supplied with python-pptx).

Files expected alongside this script:
    Final Deck.pptx
    LatestFile_with_levels_and_summary.xlsx

Output:
    Final_Deck_3Q26_Compact.pptx

Use the ORIGINAL Final Deck, never the previously modified/overlapping deck.
"""

from pathlib import Path
from zipfile import ZipFile
import re
import sys

from openpyxl import load_workbook
from lxml import etree

HERE = Path(__file__).resolve().parent
SOURCE_EXCEL = HERE / "LatestFile_with_levels_and_summary.xlsx"
ORIGINAL_PPTX = HERE / "Final Deck.pptx"
OUTPUT_PPTX = HERE / "Final_Deck_3Q26_Compact.pptx"

# Compact dimensions (inches), matching the original visible 6-row summary.
# Header (title + quarter labels), then Level 1/2/3/Total rows.
COMPACT_ROW_INCHES = (0.50, 0.24, 0.225, 0.225, 0.225, 0.225)
EMU_PER_INCH = 914400
TARGET_SLIDE = "ppt/slides/slide2.xml"
NEW_QUARTER = "3Q26"

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
A_TR = "{%s}tr" % NS["a"]
A_TC = "{%s}tc" % NS["a"]
A_T = "{%s}t" % NS["a"]


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def as_count(value, context):
    if isinstance(value, bool) or value is None:
        raise ValueError(f"Missing/invalid count: {context} = {value!r}")
    if isinstance(value, (int, float)) and float(value).is_integer():
        n = int(value)
    elif isinstance(value, str) and re.fullmatch(r"\d+", value.strip()):
        n = int(value.strip())
    else:
        raise ValueError(f"Non-integer count: {context} = {value!r}")
    if n < 0:
        raise ValueError(f"Negative count: {context} = {n}")
    return n


def read_summary(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Color Level Summary" not in wb.sheetnames:
            raise ValueError("Excel missing sheet 'Color Level Summary'")
        ws = wb["Color Level Summary"]
        data = list(ws.iter_rows(min_row=1, max_row=100, max_col=20, values_only=True))
        header_idx = None
        headers = None
        for i, row in enumerate(data):
            m = {clean(x).lower(): j for j, x in enumerate(row) if x is not None}
            if {"level", "green", "yellow", "red", "total"} <= set(m):
                header_idx, headers = i, m
                break
        if headers is None:
            raise ValueError("Could not find Level, Green, Yellow, Red, Total headers")
        counts = {}
        for row in data[header_idx + 1:]:
            level = clean(row[headers["level"]])
            if level not in {"Level 1", "Level 2", "Level 3", "Grand Total"}:
                continue
            if level in counts:
                raise ValueError(f"Duplicate summary level: {level}")
            counts[level] = {c: as_count(row[headers[c.lower()]], f"{level}/{c}")
                             for c in ("Green", "Yellow", "Red", "Total")}
            if sum(counts[level][c] for c in ("Green", "Yellow", "Red")) != counts[level]["Total"]:
                raise ValueError(f"Count total mismatch in Excel: {level}")
        needed = {"Level 1", "Level 2", "Level 3", "Grand Total"}
        if set(counts) != needed:
            raise ValueError(f"Excel summary missing: {sorted(needed - set(counts))}")
        for color in ("Green", "Yellow", "Red", "Total"):
            if sum(counts[f"Level {i}"][color] for i in (1, 2, 3)) != counts["Grand Total"][color]:
                raise ValueError(f"Grand total mismatch in Excel for {color}")
        return counts
    finally:
        wb.close()


def cell_text(cell):
    return "".join(t.text or "" for t in cell.iter(A_T)).strip()


def set_cell_text(cell, text):
    """Modify only existing <a:t> text nodes, retaining every formatting tag."""
    nodes = list(cell.iter(A_T))
    if not nodes:
        raise ValueError(f"Cell has no editable text run; cannot safely insert {text!r}")
    nodes[0].text = str(text)
    for node in nodes[1:]:
        node.text = ""


def label_kind(value):
    t = clean(value).lower()
    m = re.fullmatch(r"level\s*([123])(?:\s*metric)?", t)
    if m:
        return "Level " + m.group(1)
    if t in {"total", "grand total"}:
        return "Grand Total"
    return None


def find_top_summary(root):
    matches = []
    for frame in root.xpath(".//p:graphicFrame", namespaces=NS):
        tables = frame.xpath("./a:graphic/a:graphicData/a:tbl", namespaces=NS)
        if not tables:
            continue
        table = tables[0]
        rows = [list(r.findall(A_TC)) for r in table.findall(A_TR)]
        if len(rows) < 6 or any(len(r) != 8 for r in rows):
            continue
        row_indexes = {}
        for i, cells in enumerate(rows):
            kind = label_kind(cell_text(cells[0]))
            if kind:
                row_indexes[kind] = i
        if {"Level 1", "Level 2", "Level 3", "Grand Total"} <= set(row_indexes):
            matches.append((table, rows, row_indexes))
    if len(matches) != 1:
        raise ValueError(
            f"Expected one editable 8-column summary table on slide 2; found {len(matches)}. "
            "No presentation changes saved."
        )
    return matches[0]


def normalized_structure(xml):
    """Compare structure ignoring only cell text and intended table dimensions."""
    root = etree.fromstring(xml)
    table, _rows, _indexes = find_top_summary(root)
    for node in root.iter(A_T):
        node.text = "CHECK"
    for tr in table.findall(A_TR):
        tr.set("h", "CHECK")
        for tc in tr.findall(A_TC):
            tc_pr = tc.find("{%s}tcPr" % NS["a"])
            if tc_pr is not None:
                tc_pr.attrib.pop("marT", None)
                tc_pr.attrib.pop("marB", None)
    frame = table.getparent().getparent().getparent()
    xfrm = frame.find("{%s}xfrm" % NS["p"])
    if xfrm is not None:
        ext = xfrm.find("{%s}ext" % NS["a"])
        if ext is not None:
            ext.set("cy", "CHECK")
    return etree.tostring(root, method="c14n")


def rewrite_slide_xml(original_xml, counts):
    root = etree.fromstring(original_xml)
    table, rows, indexes = find_top_summary(root)
    before_w = [col.get("w") for col in table.xpath("./a:tblGrid/a:gridCol", namespaces=NS)]

    # Blank previous quarter only, and replace latest quarter labels in place.
    # Keep all 8 columns -- DO NOT merge/remove previous-quarter columns.
    header_changes = {"previous": 0, "latest": 0}
    for cells in rows:
        for cell in cells:
            label = clean(cell_text(cell)).lower()
            if label == "1q26":
                set_cell_text(cell, "")
                header_changes["previous"] += 1
            elif label == "2q26":
                set_cell_text(cell, NEW_QUARTER)
                header_changes["latest"] += 1

    if header_changes != {"previous": 3, "latest": 3}:
        raise ValueError(
            f"Unexpected quarter labels {header_changes}; need 3 previous and 3 latest. "
            "Stopped to avoid changing wrong elements."
        )

    for level in ("Level 1", "Level 2", "Level 3", "Grand Total"):
        cells = rows[indexes[level]]
        values = counts[level]
        for col in (1, 3, 5):
            set_cell_text(cells[col], "")
        for col, key in ((2, "Green"), (4, "Yellow"), (6, "Red"), (7, "Total")):
            set_cell_text(cells[col], values[key])

    # IMPORTANT: the earlier versions preserved source XML row heights,
    # but those heights produced a 6-row summary that extended into the next
    # table. Fix that actual geometry, not only the contents.
    trs = table.findall(A_TR)
    if len(trs) != len(COMPACT_ROW_INCHES):
        raise ValueError(f"Expected six summary rows, found {len(trs)}")
    for tr, height_inches in zip(trs, COMPACT_ROW_INCHES):
        tr.set("h", str(round(height_inches * EMU_PER_INCH)))
        # PowerPoint's default top/bottom cell padding can force a table row
        # to expand again when the deck is opened. Remove vertical padding
        # in the summary ONLY (do not change fonts, fills or borders).
        for tc in tr.findall(A_TC):
            tc_pr = tc.find("{%s}tcPr" % NS["a"])
            if tc_pr is not None:
                tc_pr.set("marT", "0")
                tc_pr.set("marB", "0")
    frame = table.getparent().getparent().getparent()
    xfrm = frame.find("{%s}xfrm" % NS["p"])
    if xfrm is None or xfrm.find("{%s}ext" % NS["a"]) is None:
        raise ValueError("Summary graphicFrame lacks an editable extent")
    xfrm.find("{%s}ext" % NS["a"]).set(
        "cy", str(sum(round(x * EMU_PER_INCH) for x in COMPACT_ROW_INCHES))
    )
    after_w = [col.get("w") for col in table.xpath("./a:tblGrid/a:gridCol", namespaces=NS)]
    if before_w != after_w:
        raise RuntimeError("Safety check failed: table column widths changed")

    result = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if normalized_structure(original_xml) != normalized_structure(result):
        raise RuntimeError("Safety check failed: non-text slide XML changed")
    return result


def modify_deck(pptx_path, output_path, counts):
    if pptx_path.resolve() == output_path.resolve():
        raise ValueError("Do not overwrite original Final Deck")
    with ZipFile(pptx_path, "r") as src:
        if TARGET_SLIDE not in src.namelist():
            raise ValueError("Slide 2 is missing from the PowerPoint file")
        old_xml = src.read(TARGET_SLIDE)
        new_xml = rewrite_slide_xml(old_xml, counts)
        with ZipFile(output_path, "w") as dst:
            for info in src.infolist():
                dst.writestr(info, new_xml if info.filename == TARGET_SLIDE else src.read(info.filename))
    # Post-write: exact table data and geometric validation.
    with ZipFile(output_path, "r") as out:
        root = etree.fromstring(out.read(TARGET_SLIDE))
        _table, rows, indexes = find_top_summary(root)
        for level, idx in indexes.items():
            cells = rows[idx]
            for col in (1, 3, 5):
                if cell_text(cells[col]):
                    raise RuntimeError(f"Previous-quarter cell not blank: {level}/{col}")
            for col, key in ((2, "Green"), (4, "Yellow"), (6, "Red"), (7, "Total")):
                if cell_text(cells[col]) != str(counts[level][key]):
                    raise RuntimeError(f"Verification failed: {level}/{key}")


def main():
    if not SOURCE_EXCEL.is_file():
        raise FileNotFoundError(f"Generated Excel missing: {SOURCE_EXCEL}")
    if not ORIGINAL_PPTX.is_file():
        raise FileNotFoundError(f"Original presentation missing: {ORIGINAL_PPTX}")
    counts = read_summary(SOURCE_EXCEL)
    print("Counts from generated Excel:")
    for label in ("Level 1", "Level 2", "Level 3", "Grand Total"):
        print(f"  {label:12} " + " | ".join(f"{k}: {counts[label][k]}" for k in ("Green", "Yellow", "Red", "Total")))
    modify_deck(ORIGINAL_PPTX, OUTPUT_PPTX, counts)
    print(f"SUCCESS: {OUTPUT_PPTX}")
    print("Column widths/fonts/colors preserved; compact row heights and vertical margins enforced.")
    print("Previous-quarter cells blank; latest = 3Q26.")
    print("Note: slide title/narrative and detailed lower table were not changed.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
