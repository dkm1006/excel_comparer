"""Shared helpers for cell-level checks.

These helpers operate on raw cell *values* (whatever ``openpyxl`` returns
when ``data_only=True``), not on ``Cell`` objects, so they're easy to
unit-test and reuse across checks.
"""

import math
from typing import Iterable

from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet


# Default number of consecutive empty rows after which ``find_last_data_row``
# stops scanning. Conservative enough for typical IRO sheets but bounded.
DEFAULT_MAX_TRAILING_EMPTY = 30


def is_empty(value: object) -> bool:
    """Treat ``None``, NaN floats, and whitespace-only strings as empty."""
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def is_in_range(
    value: object,
    lo: float,
    hi: float,
    *,
    integers_only: bool = False,
) -> bool:
    """Return ``True`` iff *value* is a number in ``[lo, hi]``.

    Empty values, booleans and non-numeric values all return ``False`` so
    the caller can distinguish "wrong value" from "no value" (use
    :func:`is_empty` first if you care about that distinction).
    """
    if is_empty(value) or isinstance(value, bool):
        return False
    if not isinstance(value, (int, float)):
        return False
    if integers_only and isinstance(value, float) and not value.is_integer():
        return False
    return lo <= value <= hi


def normalize_column(value: int | str) -> int:
    """Accept a 1-based int or an Excel column letter; return a 1-based index."""
    value = column_index_from_string(value.strip().upper()) if isinstance(value, str) and value.strip() else value

    if value < 1:
        raise ValueError(f"Column index must be >= 1, got {value}")
    
    return int(value)



def normalize_columns(values: Iterable[int | str]) -> list[int]:
    """Apply :func:`normalize_column` to every entry."""
    return [normalize_column(v) for v in values]


def find_last_data_row(
    ws: Worksheet,
    *,
    start_row: int,
    scan_min_col: int,
    scan_max_col: int,
    max_trailing_empty: int = DEFAULT_MAX_TRAILING_EMPTY,
) -> int:
    """Return the last 1-based row index that has any value in the scan range.

    Returns ``start_row - 1`` when no data row is found. Stops scanning
    after *max_trailing_empty* consecutive empty rows so unbounded sheets
    don't hang the check.
    """
    last = start_row - 1
    consecutive_empty = 0
    for row_cells in ws.iter_rows(
        min_row=start_row, min_col=scan_min_col, max_col=scan_max_col
    ):
        if any(cell.value is not None for cell in row_cells):
            last = row_cells[0].row
            consecutive_empty = 0
        else:
            consecutive_empty += 1
            if consecutive_empty > max_trailing_empty:
                break
    return last


def all_equal_case_insensitive(values: Iterable[object], target: str) -> bool:
    """Case-insensitive: are all *values* equal to *target* after str()/strip()?

    Empty cells (``None``) count as not equal to *target*, so a partially
    filled row never satisfies this predicate.
    """
    target_norm = target.strip().lower()
    return all(
        v is not None and str(v).strip().lower() == target_norm
        for v in values
    )
