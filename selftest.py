"""End-to-end self-test for excel_comparer.

Builds two small workbooks in a temporary directory with a known set of
differences and exercises:

* the check pipeline (baseline comparison, rating band, required cells,
  min text length, irremediable character),
* the text report (PASSED / FAILED, per-check grouping, row-prefixed lines),
* the annotated Excel export (highlighted cells, applied auto-corrections,
  "Review Documentation" column and summary sheet).

Used for quick manual verification::

    python selftest.py
"""

import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from excel_comparer import (
    BaselineComparisonCheck,
    CheckContext,
    DiffCategory,
    ExportConfig,
    IrremediableCharacterCheck,
    MinTextLengthCheck,
    RatingBandCheck,
    RequiredCellsCheck,
    SheetRanges,
    export_annotated_workbook,
    format_differences,
)


def _build_baseline(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"

    ws["A1"] = "IRO-Classification"
    ws["B1"] = "IRO-Scoring"
    ws["C1"] = "IRO Name"
    ws["D1"] = "Documentation"

    ws["A2"] = "Type"
    ws["B2"] = "Score"
    ws["C2"] = "Name"
    ws["D2"] = "Reasoning"

    ws["A3"] = "Negative Impact"
    ws["B3"] = 3
    ws["C3"] = "iro-1"
    ws["D3"] = "x" * 300

    ws["A4"] = "Positive Impact"
    ws["B4"] = "-"
    ws["C4"] = "iro-2"
    ws["D4"] = "y" * 300

    bold = Font(bold=True)
    for col in ("A", "B", "C", "D"):
        ws[f"{col}2"].font = bold

    notes = wb.create_sheet("Notes")
    notes["A1"] = "Hello"

    wb.save(str(path))


def _build_candidate(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"

    ws["A1"] = "IRO-Classification"
    ws["B1"] = "IRO-Scoring"
    ws["D1"] = "Documentation"

    ws["A2"] = "Type"
    ws["B2"] = "Score"
    ws["C2"] = "Aux"
    ws["D2"] = "Reasoning"

    # Negative impact with an out-of-range score (RatingBandCheck) — also
    # picked up by IrremediableCharacterCheck since 9 is not in {1..4}.
    ws["A3"] = "Negative Impact"
    ws["B3"] = 9
    ws["C3"] = "iro-1"
    ws["D3"] = "short"  # MinTextLengthCheck: reasoning too short

    # Positive impact with the wrong score — IrremediableCharacterCheck will
    # auto-correct to "-".
    ws["A4"] = "Positive Impact"
    ws["B4"] = 4
    ws["C4"] = "iro-2"
    ws["D4"] = "z" * 300

    # Row 5 has plenty of "wrong" values that *would* be flagged, but its
    # anchor column (C, the IRO Name) is empty — so with
    # row_anchor_column="C" configured below, the row-bound checks should
    # all skip it. Used to verify the new "skip rows with empty anchor"
    # behaviour.
    ws["A5"] = "Negative Impact"
    ws["B5"] = 99           # would trip rating_band + irremediable_character
    ws["C5"] = None         # empty anchor -> row should be skipped
    ws["D5"] = "too short"  # would trip min_text_length

    extra = wb.create_sheet("Extra")
    extra["A1"] = "surprise"

    wb.save(str(path))


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        baseline = tmp_path / "baseline.xlsx"
        candidate = tmp_path / "candidate.xlsx"

        _build_baseline(baseline)
        _build_candidate(candidate)

        # All row-bound checks use the IRO Name column (C) as their anchor;
        # row 5 has an empty C and should therefore be skipped.
        anchor = "C"

        checks = [
            BaselineComparisonCheck(
                baseline_path=baseline,
                categories=(
                    DiffCategory.VALUE,
                    DiffCategory.SHEET_MISSING,
                    DiffCategory.SHEET_EXTRA,
                ),
                ranges={"Sales": SheetRanges.parse("A3:D10")},
            ),
            RatingBandCheck(
                min_value=1,
                max_value=4,
                integers_only=True,
                data_start_row=3,
                row_anchor_column=anchor,
                ranges={"Sales": SheetRanges.parse("B")},
            ),
            RequiredCellsCheck(
                data_start_row=3,
                row_anchor_column=anchor,
                ranges={"Sales": SheetRanges.parse("A")},
            ),
            MinTextLengthCheck(
                min_chars=100,
                data_start_row=3,
                row_anchor_column=anchor,
                ranges={"Sales": SheetRanges.parse("D")},
            ),
            IrremediableCharacterCheck(
                type_column="A",
                score_column="B",
                data_start_row=3,
                row_anchor_column=anchor,
                ranges={"Sales": SheetRanges.parse("A:D")},
            ),
        ]

        # Run the pipeline.
        wb = load_workbook(filename=str(candidate), data_only=True)
        ctx = CheckContext(workbook_path=candidate, workbook=wb)
        diffs = []
        for check in checks:
            diffs.extend(check.run(ctx))

        print(format_differences(diffs, file_label=str(candidate)))
        print()

        export_cfg = ExportConfig(
            doc_columns={"Sales": 4},   # column D
            header_row=2,
            output_dir=str(tmp_path / "out"),
        )
        out = export_annotated_workbook(candidate, diffs, export_cfg)
        print(f"Annotated copy written to: {out}")
        print(f"({len(diffs)} findings total)")


if __name__ == "__main__":
    main()
