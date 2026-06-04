"""Minimum-text-length check.

For every non-empty cell inside the configured ``ranges``, asserts that
the cell's text has at least ``min_chars`` non-whitespace characters.
Emits a :data:`DiffCategory.REASONING_TOO_SHORT` finding per too-short
cell. Empty cells are *not* reported here — that's the
:class:`RequiredCellsCheck`'s job.

The check runs on every sheet listed in ``ranges``. If ``ranges`` is
empty or missing the check is a no-op.

Configuration (TOML)
--------------------

.. code-block:: toml

    [[checks]]
    name = "min_text_length"
    min_chars = 250

    [checks.ranges]
    Impacts = "U3:U"
"""

from typing import ClassVar

from openpyxl.utils import get_column_letter

from excel_comparer.checks.base import CheckContext, register_check
from excel_comparer.checks.helpers import find_last_data_row, is_empty, normalize_column
from excel_comparer.config import SheetRanges
from excel_comparer.models import DiffCategory, Difference


@register_check
class MinTextLengthCheck:
    """Assert that every non-empty cell in ``ranges`` holds at least ``min_chars`` characters."""

    name: ClassVar[str] = "min_text_length"
    description: ClassVar[str] = (
        "Asserts that every non-empty cell in the configured ranges contains "
        "at least a configurable number of non-whitespace characters."
    )

    def __init__(
        self,
        *,
        min_chars: int = 250,
        header_row: int = 2,
        row_anchor_column: int | str | None = None,
        ranges: dict[str, SheetRanges] | None = None,
    ) -> None:
        self.min_chars = int(min_chars)
        self.header_row = header_row
        self.row_anchor_column: int | None = (
            normalize_column(row_anchor_column) if row_anchor_column else None
        )
        self.ranges: dict[str, SheetRanges] = ranges or {}

    def run(self, ctx: CheckContext) -> list[Difference]:
        diffs: list[Difference] = []
        for sheet_name, sr in self.ranges.items():
            if sr.skip or sheet_name not in ctx.workbook.sheetnames:
                continue
            ws = ctx.workbook[sheet_name]

            scan_min_col, scan_max_col = sr.clamp_cols(1, ws.max_column)
            if scan_min_col > scan_max_col:
                continue
            last_row = find_last_data_row(
                ws,
                start_row=self.header_row + 1,
                scan_min_col=scan_min_col,
                scan_max_col=scan_max_col,
            )
            if last_row < self.header_row + 1:
                continue

            anchor_empty: dict[int, bool] = {}

            for row, col in sr.iter_cells(
                row_floor=self.header_row + 1,
                row_ceiling=last_row,
                col_ceiling=ws.max_column,
            ):
                if self.row_anchor_column is not None:
                    if row not in anchor_empty:
                        anchor_empty[row] = is_empty(
                            ws.cell(row=row, column=self.row_anchor_column).value
                        )
                    if anchor_empty[row]:
                        continue
                value = ws.cell(row=row, column=col).value
                if is_empty(value):
                    continue
                text_len = len(str(value).strip())
                if text_len < self.min_chars:
                    diffs.append(
                        Difference(
                            category=DiffCategory.REASONING_TOO_SHORT,
                            sheet=sheet_name,
                            cell=f"{get_column_letter(col)}{row}",
                            golden=f"at least {self.min_chars} chars",
                            other=f"{text_len} chars",
                            check_name=self.name,
                        )
                    )
        return diffs
