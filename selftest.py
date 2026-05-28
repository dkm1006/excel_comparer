"""End-to-end self-test for excel_comparer.

Builds two small workbooks in a temporary directory with a known set of
differences and prints the resulting diff report.  Used for quick manual
verification:

    python selftest.py
"""


import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from excel_comparer import ExcelComparer, format_differences


def _build_golden(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"

    ws["A1"] = "Item"
    ws["B1"] = "Qty"
    ws["A2"] = "Apples"
    ws["B2"] = 100
    ws["A3"] = "Bananas"
    ws["B3"] = 50

    # Bold header
    bold = Font(bold=True)
    ws["A1"].font = bold
    ws["B1"].font = bold

    # Merged title
    ws["A5"] = "Totals"
    ws.merge_cells("A5:B5")

    # Cell with specific fill + alignment + number format
    ws["B2"].fill = PatternFill(patternType="solid", fgColor="FFFF00")
    ws["B2"].alignment = Alignment(horizontal="right")
    ws["B2"].number_format = "#,##0"

    # A second sheet
    notes = wb.create_sheet("Notes")
    notes["A1"] = "Hello"

    wb.save(str(path))


def _build_candidate(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"

    ws["A1"] = "Item"
    ws["B1"] = "Qty"
    ws["A2"] = "Apples"
    ws["B2"] = 120              # VALUE diff vs golden=100
    # ws["A3"] missing entirely  # VALUE diff (golden=Bananas, other=None)
    # ws["B3"] missing entirely  # VALUE diff (golden=50, other=None)
    ws["C2"] = "extra"          # VALUE diff (golden=None, other='extra')

    # Header no longer bold       # FORMAT diff on font.bold for A1 and B1
    # (default font, no bold)

    # No merged range            # MERGED_CELL missing in other: A5:B5
    ws["A5"] = "Totals"

    # Different fill on B2       # FORMAT diff on fill
    ws["B2"].fill = PatternFill(patternType="solid", fgColor="00FF00")
    # Different number_format    # FORMAT diff on number_format
    ws["B2"].number_format = "0.00"

    # "Notes" sheet missing      # SHEET_MISSING
    # Extra sheet
    extra = wb.create_sheet("Extra")
    extra["A1"] = "surprise"

    wb.save(str(path))


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        golden = tmp_path / "golden.xlsx"
        candidate = tmp_path / "candidate.xlsx"

        _build_golden(golden)
        _build_candidate(candidate)

        comparer = ExcelComparer(golden)
        results = comparer.compare(candidate)

        for path, diffs in results.items():
            print(f"=== {path} ===")
            print(format_differences(diffs))
            print(f"\n({len(diffs)} differences total)")


if __name__ == "__main__":
    main()
