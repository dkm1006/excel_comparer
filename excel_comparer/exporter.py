"""Annotated Excel export.

Given the original candidate workbook and the list of differences found
by every configured check, produce a new ``<name><suffix>.xlsx`` file
where:

* every cell-bound diff is filled with the configured highlight color,
* every diff that carries an auto-correction (``Difference.correction``)
  has the corresponding cell rewritten to the correction value (and is
  filled with ``correction_color`` instead of ``highlight_color``),
* a new "Review Documentation" column is inserted immediately to the
  right of the existing Documentation column on each sheet listed in
  ``export.doc_columns``. Each row that has at least one cell-bound diff
  gets a human-readable summary (one ``[FLAGGED] <message>`` line per
  diff -- or ``[CORRECTED] <message>`` when an auto-correction was
  applied to the cell -- joined by newlines),

* an optional summary sheet is prepended listing PASSED / FAILED status
  per source sheet plus per-category counts.

``MERGED_CELL`` / ``SHEET_*`` diffs are *not* written into the doc column
(they don't refer to a specific row).
"""

from collections import Counter, defaultdict
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_comparer.config import ExportConfig
from excel_comparer.models import DiffCategory, Difference


# Markers prepended to per-diff lines in the doc column.
FLAGGED_MARKER = "[FLAGGED]"
CORRECTED_MARKER = "[CORRECTED]"


# Categories that aren't bound to a row in any source sheet
# (they shouldn't appear in the doc column).
_WORKBOOK_LEVEL_CATEGORIES: frozenset[DiffCategory] = frozenset({
    DiffCategory.SHEET_MISSING,
    DiffCategory.SHEET_EXTRA,
})


def export_annotated_workbook(
    original_path: str | Path,
    diffs: Iterable[Difference],
    config: ExportConfig,
) -> Path:
    """Write an annotated copy of *original_path* alongside it (or to ``output_dir``).

    Returns the path to the written file.
    """
    original_path = Path(original_path)
    diffs = list(diffs)

    wb: Workbook = load_workbook(filename=str(original_path), rich_text=True)
    unprotect_sheets(wb)

    highlight = PatternFill(
        patternType="solid",
        fgColor=config.highlight_color,
        bgColor=config.highlight_color,
    )
    correction_fill = PatternFill(
        patternType="solid",
        fgColor=config.correction_color,
        bgColor=config.correction_color,
    )

    # Group diffs per sheet for both highlighting and the doc column.
    diffs_by_sheet: dict[str, list[Difference]] = defaultdict(list)
    for d in diffs:
        diffs_by_sheet[d.sheet].append(d)

    for sheet_name, sheet_diffs in diffs_by_sheet.items():
        # A SHEET_MISSING diff points at a sheet that doesn't exist in the
        # candidate workbook; nothing to annotate per-cell for those.
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]

        # 1. Apply corrections + highlight individual cells.
        for d in sheet_diffs:
            if not d.is_cell_bound or d.cell is None:
                continue
            cell = ws[d.cell]
            if config.apply_corrections and d.has_correction:
                cell.value = d.correction
                cell.fill = correction_fill
            else:
                cell.fill = highlight

        # 2. Insert the per-sheet "Review Documentation" column.
        doc_col_idx = config.doc_columns.get(sheet_name)
        if doc_col_idx:
            _insert_doc_column(ws, doc_col_idx, sheet_diffs, config)

    # 3. Optional summary sheet at the front.
    if config.add_summary_sheet:
        _prepend_summary_sheet(wb, diffs, original_path, config)

    output_path = _resolve_output_path(original_path, config)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))
    return output_path


def unprotect_sheets(wb: Workbook) -> None:
    """Unprotect every sheet in *wb* so we can write corrections into them."""
    for ws in wb:
        ws.protection.disable()

# --------------------------------------------------------------------------- #
# Internal helpers                                                            #
# --------------------------------------------------------------------------- #


def _resolve_output_path(other_path: Path, config: ExportConfig) -> Path:
    name = f"{other_path.stem}{config.output_suffix}{other_path.suffix}"
    if config.output_dir:
        return Path(config.output_dir) / name
    return other_path.with_name(name)


def _insert_doc_column(
    ws: Worksheet,
    after_col: int,
    sheet_diffs: list[Difference],
    config: ExportConfig,
) -> None:
    """Insert a new column at ``after_col + 1`` and populate it from *sheet_diffs*.

    Only row-bound diffs contribute -- workbook-level diffs (sheet
    missing / extra / merged cell) are skipped because they don't
    reference a row.
    """
    new_col_idx = after_col + 1
    ws.insert_cols(new_col_idx)

    new_col_letter = get_column_letter(new_col_idx)
    ws.column_dimensions[new_col_letter].width = config.doc_column_width

    # Copy the header row(s) styling from the original doc column, and write
    # the new header into the configured header row.
    header_row = config.header_row
    header_cell = ws.cell(row=header_row, column=new_col_idx, value=config.doc_column_header)
    header_cell.font = Font(bold=True)
    header_cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Group row-bound diffs by row and write one entry per row.
    diffs_by_row: dict[int, list[Difference]] = defaultdict(list)
    for d in sheet_diffs:
        if d.category in _WORKBOOK_LEVEL_CATEGORIES:
            continue
        if d.category == DiffCategory.MERGED_CELL:
            continue
        row = d.row
        if row is None or row <= header_row:
            continue
        diffs_by_row[row].append(d)

    flag_fill = PatternFill(
        patternType="solid",
        fgColor=config.flag_color,
        bgColor=config.flag_color,
    )

    for row, row_diffs in sorted(diffs_by_row.items()):
        text = "\n".join(
            d.message for d in sorted(row_diffs, key=lambda x: (x.category))
        )
        cell = ws.cell(row=row, column=new_col_idx, value=text)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.fill = flag_fill


def _format_doc_line(d: Difference, config: ExportConfig) -> str:
    """Render one line of the in-workbook "Review Documentation" column.

    Format: ``[FLAGGED] <message>`` (or ``[CORRECTED] <message>`` when
    an auto-correction was applied to the cell). The check name is not
    included -- :attr:`Difference.message` is already worded for the
    reviewer, and the marker tells them whether the row was just
    flagged or actually rewritten.
    """
    marker = (
        CORRECTED_MARKER
        if config.apply_corrections and d.has_correction
        else FLAGGED_MARKER
    )
    return f"{marker} {d.message}"



def _prepend_summary_sheet(
    wb: Workbook,
    diffs: list[Difference],
    other_path: Path,
    config: ExportConfig,
) -> None:
    """Create a "Comparison Summary" sheet at index 0 listing PASSED/FAILED per sheet."""
    name = config.summary_sheet_name
    # If a previous summary exists, drop it first.
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(title=name, index=0)

    bold = Font(bold=True)

    ws["A1"] = "File"
    ws["B1"] = str(other_path)
    ws["A2"] = "Overall verdict"
    ws["B2"] = "PASSED" if not diffs else "FAILED"
    for c in ("A1", "A2"):
        ws[c].font = bold

    # Per-check verdicts.
    by_check: dict[str, list[Difference]] = defaultdict(list)
    for d in diffs:
        by_check[d.check_name or "(unnamed check)"].append(d)

    row = 4
    ws.cell(row=row, column=1, value="Check").font = bold
    ws.cell(row=row, column=2, value="Verdict").font = bold
    ws.cell(row=row, column=3, value="Findings").font = bold
    row += 1
    for check_name in sorted(by_check):
        check_diffs = by_check[check_name]
        ws.cell(row=row, column=1, value=check_name)
        ws.cell(row=row, column=2, value="FAILED" if check_diffs else "PASSED")
        ws.cell(row=row, column=3, value=len(check_diffs))
        row += 1

    # Per-sheet table (all checks combined).
    row += 1
    table_header_row = row
    headers = ["Sheet", "Verdict", "Total"]
    # Add a column per observed category.
    seen_categories = sorted(
        {d.category for d in diffs if d.category not in _WORKBOOK_LEVEL_CATEGORIES},
        key=lambda c: c.value,
    )
    headers.extend(c.value for c in seen_categories)
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=table_header_row, column=i, value=h)
        c.font = bold
    row += 1

    per_sheet_counts: dict[str, Counter[DiffCategory]] = defaultdict(Counter)
    workbook_level: list[Difference] = []
    for d in diffs:
        if d.category in _WORKBOOK_LEVEL_CATEGORIES:
            workbook_level.append(d)
        else:
            per_sheet_counts[d.sheet][d.category] += 1

    for sheet in sorted(per_sheet_counts):
        counts = per_sheet_counts[sheet]
        total = sum(counts.values())
        ws.cell(row=row, column=1, value=sheet)
        ws.cell(row=row, column=2, value="FAILED" if total else "PASSED")
        ws.cell(row=row, column=3, value=total)
        for i, cat in enumerate(seen_categories, start=4):
            ws.cell(row=row, column=i, value=counts.get(cat, 0))
        row += 1

    if workbook_level:
        row += 1
        title = ws.cell(row=row, column=1, value="Workbook-level differences")
        title.font = bold
        row += 1
        for d in workbook_level:
            ws.cell(row=row, column=1, value=d.message)
            row += 1

    # Sensible column widths.
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    for i in range(4, 4 + len(seen_categories)):
        ws.column_dimensions[get_column_letter(i)].width = 22


__all__ = [
    "export_annotated_workbook",
    "FLAGGED_MARKER",
    "CORRECTED_MARKER",
]
