"""Compare Excel workbooks against a golden reference file."""

from excel_comparer.comparer import ExcelComparer
from excel_comparer.models import DiffCategory, Difference
from excel_comparer.reporting import format_differences

__all__ = [
    'ExcelComparer',
    'DiffCategory',
    'Difference',
    'format_differences',
]
