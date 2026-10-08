#!/usr/bin/env python3
r"""
Update slide 2's ORIGINAL 8-column Credit Risk summary table using the
Excel "Color Level Summary" sheet. PowerPoint itself (Windows COM) edits the
cells: no table rebuilding, no column removal/merging, no XML layout rewrites.

Expected, in SAME Scripts folder as this script:
    Final Deck.pptx
    LatestFile_with_levels_and_summary.xlsx

Output (original presentation is never overwritten):
    Final_Deck_3Q26_Native_Updated.pptx

Behavior:
 - Blank previous-quarter headings and values, keep their original columns.
 - Set the latest-quarter headings to 3Q26.
 - Put current Green/Yellow/Red and row totals into their existing cells.
 - Check table geometry before and after; stop if PowerPoint resized it.
 - Make no other slide changes; slide title and detailed table remain unchanged.

One-time libraries (if missing and installation is allowed):
    python -m pip install openpyxl pywin32

Run:
    python update_slide2_NATIVE_POWERPOINT.py
"""

from pathlib import Path
from tempfile import NamedTemporaryFile
import re
import os
import sys

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
EXCEL_PATH = ROOT / "LatestFile_with_levels_and_summary.xlsx"
POWERPOINT_PATH = ROOT / "Final Deck.pptx"
OUTPUT_PATH = ROOT / "Final_Deck_3Q26_Native_Updated.pptx"
NEW_QUARTER = "3Q26"


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def read_counts(path):
    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        if "Color Level Summary" not in workbook.sheetnames:
            raise ValueError("Excel sheet 'Color Level Summary' not found")
        sheet = workbook["Color Level Summary"]
        values = list(sheet.iter_rows(min_row=1, max_col=5, values_only=True))
        header = [clean(v).lower() for v in values[0]]
        expected = ["level", "green", "yellow", "red", "total"]
        if header != expected:
            raise ValueError(f"Unexpected Excel summary columns: {header}")
        counts = {}
        for row in values[1:]:
            label = clean(row[0]);
            if label not in ("Level 1", "Level 2", "Level 3", "Grand Total"):
                continue
            if label in counts:
                raise ValueError(f"Duplicate Excel summary row: {label}")
            row_counts = {}
            for color, v in zip(expected[1:], row[1:]):
                if isinstance(v, bool) or v is None or not isinstance(v, (int, float)) or int(v) != v or v < 0:
                    raise ValueError(f"Invalid count {label} {color}: {v!r}")
                row_counts[color] = int(v)
            if row_counts["green"] + row_counts["yellow"] + row_counts["red"] != row_counts["total"]:
                raise ValueError(f"Invalid {label} total in Excel")
            counts[label] = row_counts
        if set(counts) != {"Level 1", "Level 2", "Level 3", "Grand Total"}:
            raise ValueError("Excel summary must contain all three levels and Grand Total")
        for color in ("green", "yellow", "red", "total"):
            if sum(counts[f"Level {n}"][color] for n in (1, 2, 3)) != counts["Grand Total"][color]:
                raise ValueError(f"Grand Total mismatch in Excel: {color}")
        return counts
    finally:
        workbook.close()


def cell_text(table, row, col):
    return clean(table.Cell(row, col).Shape.TextFrame.TextRange.Text)


def set_cell_text(table, row, col, text):
    # Microsoft PowerPoint edits the existing cell, not its row/column layout.
    shape = table.Cell(row, col).Shape
    text_range = shape.TextFrame.TextRange
    desired = str(text)
    if clean(text_range.Text) != desired:
        text_range.Text = desired


def locate_summary(slide):
    matches = []
    for shape in slide.Shapes:
        try:
            if not shape.HasTable:
                continue
            table = shape.Table
            if table.Rows.Count != 6 or table.Columns.Count != 8:
                continue
            labels = {}
            for row in range(1, table.Rows.Count + 1):
                value = cell_text(table, row, 1).lower()
                if value in ("level 1", "level 2", "level 3"):
                    labels["Level " + value[-1]] = row
                if value in ("total", "grand total"):
                    labels["Grand Total"] = row
            if len(labels) == 4:
                matches.append((shape, table, labels))
        except Exception:
            continue
    if len(matches) != 1:
        raise RuntimeError(f"Found {len(matches)} matching 6-row / 8-column tables on slide 2 (expected exactly 1). Stopped.")
    return matches[0]


def geometry(shape, table):
    return {
        "shape": tuple(float(getattr(shape, k)) for k in ("Left", "Top", "Width", "Height")),
        "rows": tuple(float(table.Rows(i).Height) for i in range(1, table.Rows.Count + 1)),
        "cols": tuple(float(table.Columns(i).Width) for i in range(1, table.Columns.Count + 1)),
    }


def check_geometry(before, after):
    diffs = []
    for category in ("shape", "rows", "cols"):
        for i, (x, y) in enumerate(zip(before[category], after[category]), start=1):
            if abs(x - y) > 9144:  # more than 0.01 inch in EMUs
                diffs.append(f"{category}[{i}]: {x:.0f} -> {y:.0f}")
    if diffs:
        raise RuntimeError("PowerPoint changed table geometry: " + "; ".join(diffs))


def update_table(table, row_labels, counts):
    # Expect exactly 3 old and 3 current quarter labels in the summary table.
    previous = []
    latest = []
    for row in range(1, table.Rows.Count + 1):
        for col in range(1, table.Columns.Count + 1):
            value = cell_text(table, row, col).lower()
            if value == "1q26":
                previous.append((row, col))
            elif value == "2q26":
                latest.append((row, col))
    if len(previous) != 3 or len(latest) != 3:
        raise RuntimeError(f"Unexpected quarter headers: previous={previous}; latest={latest}. Stopped without saving.")

    for row, col in previous:
        set_cell_text(table, row, col, "")
    for row, col in latest:
        set_cell_text(table, row, col, NEW_QUARTER)

    for level, row in row_labels.items():
        values = counts[level]
        for col in (2, 4, 6):
            set_cell_text(table, row, col, "")
        for col, key in ((3, "green"), (5, "yellow"), (7, "red"), (8, "total")):
            set_cell_text(table, row, col, values[key])


def verify_table(table, row_labels, counts):
    quarter_labels = []
    for row in range(1, table.Rows.Count + 1):
        for col in range(1, table.Columns.Count + 1):
            t = cell_text(table, row, col)
            if t == NEW_QUARTER:
                quarter_labels.append((row, col))
            if t in ("1Q26", "2Q26"):
                raise RuntimeError("Old quarter remains in summary table")
    if len(quarter_labels) != 3:
        raise RuntimeError(f"Expected 3 latest-quarter headings; found {len(quarter_labels)}")
    for level, row in row_labels.items():
        for col in (2, 4, 6):
            if cell_text(table, row, col):
                raise RuntimeError(f"Previous-quarter column is not blank: {level}/{col}")
        for col, key in ((3, "green"), (5, "yellow"), (7, "red"), (8, "total")):
            if cell_text(table, row, col) != str(counts[level][key]):
                raise RuntimeError(f"Wrong {level} {key}")


def main():
    if os.name != "nt":
        raise RuntimeError("This script needs Windows with Microsoft PowerPoint installed.")
    if not EXCEL_PATH.is_file() or not POWERPOINT_PATH.is_file():
        raise FileNotFoundError(f"Required files: {EXCEL_PATH} and {POWERPOINT_PATH}")
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("pywin32 is required. If permitted, install with: python -m pip install pywin32") from exc

    counts = read_counts(EXCEL_PATH)
    app = win32.DispatchEx("PowerPoint.Application")
    presentation = None
    validation_presentation = None
    tmp = OUTPUT_PATH.with_name(OUTPUT_PATH.stem + "_TEMP.pptx")
    try:
        app.Visible = True
        presentation = app.Presentations.Open(str(POWERPOINT_PATH.resolve()), True, False, False)
        if presentation.Slides.Count < 2:
            raise RuntimeError("PowerPoint has fewer than two slides")
        shape, table, levels = locate_summary(presentation.Slides(2))
        baseline = geometry(shape, table)
        update_table(table, levels, counts)
        check_geometry(baseline, geometry(shape, table))
        verify_table(table, levels, counts)
        if tmp.exists():
            tmp.unlink()
        # SaveAs: 24 = ppSaveAsOpenXMLPresentation (.pptx)
        presentation.SaveAs(str(tmp.resolve()), 24)
        presentation.Close()
        presentation = None

        # Re-open the actual saved output for post-write verification.
        validation_presentation = app.Presentations.Open(str(tmp.resolve()), True, False, False)
        after_shape, after_table, after_levels = locate_summary(validation_presentation.Slides(2))
        check_geometry(baseline, geometry(after_shape, after_table))
        verify_table(after_table, after_levels, counts)
        validation_presentation.Close()
        validation_presentation = None
        tmp.replace(OUTPUT_PATH)
        print("SUCCESS: PowerPoint table updated without changing its geometry.")
        print(f"Output: {OUTPUT_PATH}")
        for level in ("Level 1", "Level 2", "Level 3", "Grand Total"):
            vals = counts[level]
            print(f"{level:<12} Green={vals['green']:<3} Yellow={vals['yellow']:<3} Red={vals['red']:<3} Total={vals['total']}")
    finally:
        if validation_presentation is not None:
            validation_presentation.Close()
        if presentation is not None:
            presentation.Close()
        app.Quit()
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass


if __name__ == "__main__":
    main()
