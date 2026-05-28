"""Data model for differences between Excel workbooks."""

from dataclasses import dataclass
from enum import StrEnum


class DiffCategory(StrEnum):
    """High-level category of a single difference between two workbooks."""

    SHEET_MISSING = "SHEET_MISSING"  # sheet present in golden, absent in other
    SHEET_EXTRA = "SHEET_EXTRA"      # sheet present in other,  absent in golden
    VALUE = "VALUE"                  # cell value differs (covers missing / extra / changed)
    FORMAT = "FORMAT"                # a specific style attribute differs
    MERGED_CELL = "MERGED_CELL"      # merged-range set differs


@dataclass(frozen=True)
class Difference:
    """A single difference between the golden workbook and another workbook."""

    category: DiffCategory
    sheet: str
    cell: str | None = None        # e.g. "B7"; None for sheet-level or merged-range diffs
    attribute: str | None = None   # e.g. "font.bold"; None when not applicable
    golden: str | int | float | None = None
    other: str | int | float | None = None

    @property
    def message(self) -> str:
        """Human-readable message describing the difference."""
        if self.category == DiffCategory.SHEET_MISSING:
            message = f"{self.category}: Sheet '{self.sheet}' is missing in other workbook."
        elif self.category == DiffCategory.SHEET_EXTRA:
            message = f"{self.category}: Sheet '{self.sheet}' is extra in other workbook."
        elif self.category == DiffCategory.VALUE:
            message = f"{self.category}: Sheet '{self.sheet}', Cell {self.cell}: golden={self.golden} vs other={self.other}"
        elif self.category == DiffCategory.FORMAT:
            message = f"{self.category}: Sheet '{self.sheet}', Cell {self.cell}, Attribute '{self.attribute}': golden={self.golden} vs other={self.other}"
        elif self.category == DiffCategory.MERGED_CELL:
            # For merged cell differences, 'cell' is None and 'attribute' contains the range (e.g. "B2:D4")
            message = f"{self.category}: Sheet '{self.sheet}', Merged Range '{self.attribute}': golden={self.golden} vs other={self.other}"
        else:
            raise ValueError(f"Unknown DiffCategory: {self.category}")
        return message

    def __str__(self) -> str:
        return self.message
