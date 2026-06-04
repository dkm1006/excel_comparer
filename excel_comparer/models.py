"""Data model for differences between Excel workbooks."""

from dataclasses import dataclass
from enum import StrEnum
from typing import Any

from openpyxl.utils.cell import coordinate_from_string


class DiffCategory(StrEnum):
    """High-level category of a single difference between two workbooks."""

    # Baseline-comparison categories
    SHEET_MISSING = "SHEET_MISSING"          # sheet present in baseline, absent in candidate
    SHEET_EXTRA = "SHEET_EXTRA"              # sheet present in candidate, absent in baseline
    VALUE = "VALUE"                          # cell value differs (covers missing / extra / changed)
    FORMAT = "FORMAT"                        # a specific style attribute differs
    MERGED_CELL = "MERGED_CELL"              # merged-range set differs

    # Standalone-check categories
    RATING_OUT_OF_RANGE = "RATING_OUT_OF_RANGE"        # numeric cell outside an allowed [lo, hi] range
    MISSING_REQUIRED = "MISSING_REQUIRED"              # required cell is empty
    REASONING_TOO_SHORT = "REASONING_TOO_SHORT"        # free-text answer below minimum length
    NO_SELECTION = "NO_SELECTION"                      # mutually-exclusive column group with no selection
    INCONSISTENT_TYPE_SCORE = "INCONSISTENT_TYPE_SCORE"  # score doesn't match a categorical type

    @property
    def is_cell_bound(self) -> bool:
        """True iff this category refers to a single cell (and therefore a row).

        Cell-bound categories require a per-cell sweep of the worksheet; the
        others do not. 
        """
        return self in _CELL_BOUND_CATEGORIES


# Categories that refer to a specific cell (and therefore a specific row).
_CELL_BOUND_CATEGORIES = frozenset({
    DiffCategory.VALUE,
    DiffCategory.FORMAT,
    DiffCategory.RATING_OUT_OF_RANGE,
    DiffCategory.MISSING_REQUIRED,
    DiffCategory.REASONING_TOO_SHORT,
    DiffCategory.NO_SELECTION,
    DiffCategory.INCONSISTENT_TYPE_SCORE,
})


@dataclass(frozen=True)
class Difference:
    """A single difference / finding reported by a check.

    The same dataclass is used by every check (not just the baseline
    comparison); ``check_name`` records which check produced this
    finding, and ``correction`` carries an optional value the exporter
    may write back into the candidate workbook.
    """

    category: DiffCategory
    sheet: str
    cell: str | None = None        # e.g. "B7"; None for sheet-level / merged-range / row-spanning diffs
    attribute: str | None = None   # e.g. "font.bold", or a column-group label; None when not applicable
    golden: Any = None             # expected / reference value (or rule description for non-baseline checks)
    other: Any = None              # actual value found in the candidate workbook
    check_name: str = ""           # name of the Check that produced this diff
    message_override: str | None = None  # set by checks that want to provide a custom message
    correction: Any = None         # optional auto-correction value the exporter may write into ``cell``

    @property
    def message(self) -> str:
        """Short, reviewer-friendly message describing the difference.

        Wording is intentionally concise -- the exporter writes this
        verbatim into the in-workbook "Review Documentation" column,
        where the cell coordinate and sheet name are already obvious
        from context. The text report in :mod:`excel_comparer.reporting`
        prefixes its own ``Sheet "X"`` / ``Row N, Cell XY [CATEGORY]``
        headers, so we don't repeat that information here either.

        The baseline-comparison categories speak of the *benchmark*
        rather than the internal "golden / other" terminology so
        non-developer reviewers understand what's being compared.
        """
        if self.message_override is not None:
            return self.message_override
        if self.category == DiffCategory.SHEET_MISSING:
            return f"Sheet '{self.sheet}' is missing (present in the benchmark)."
        if self.category == DiffCategory.SHEET_EXTRA:
            return f"Sheet '{self.sheet}' is extra (not present in the benchmark)."
        if self.category == DiffCategory.VALUE:
            return (
                f"Column [{self.cell}] flagged. Benchmark is [{self.golden}] instead of [{self.other}]. "
                f"Changed to benchmark value / Difference is reasonable."
            )
        if self.category == DiffCategory.FORMAT:
            return (
                f"Format flagged [{self.attribute}]. "
                f"Benchmark is [{self.golden}], but found [{self.other}]."
            )

        if self.category == DiffCategory.MERGED_CELL:
            return f"Merged-cell range flagged: '{self.attribute}'."
        if self.category == DiffCategory.RATING_OUT_OF_RANGE:
            return (
                f"Cell [{self.cell}] rating out of range: [{self.other}] instead of [{self.golden}]."
            )
        if self.category == DiffCategory.MISSING_REQUIRED:
            return f"Required cell [{self.cell}] is empty."
        if self.category == DiffCategory.REASONING_TOO_SHORT:
            return f"Reasoning text too short ({self.other} characters)."
        if self.category == DiffCategory.NO_SELECTION:
            return (
                f"No selection made in group '{self.attribute}' (expected [{self.golden}])."
            )
        if self.category == DiffCategory.INCONSISTENT_TYPE_SCORE:
            message = (
                f"Score [{self.other}] is inconsistent: [{self.attribute}] requires [{self.golden}]."
                if self.correction is None else
                f"Inconsistent score [{self.other}] was corrected: [{self.attribute}] requires [{self.golden}]."
            )
            return message
        raise ValueError(f"Unknown DiffCategory: {self.category}")


    @property
    def row(self) -> int | None:
        """1-based row number.
        
        Returns ``None`` for diffs that aren't bound to any specific row.
        """
        if self.cell is not None:
            _, row = coordinate_from_string(self.cell)
            return row

    @property
    def is_cell_bound(self) -> bool:
        """True iff this diff refers to a single cell (and therefore a row)."""
        return self.category.is_cell_bound

    @property
    def has_correction(self) -> bool:
        """True iff this diff carries an auto-correction the exporter can apply."""
        return self.correction is not None and self.cell is not None

    def __str__(self) -> str:
        return self.message
