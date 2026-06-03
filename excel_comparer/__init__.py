"""Run a configurable pipeline of checks against Excel workbooks.

The baseline (golden-workbook) comparison is just one of the available
checks; others validate value ranges, completeness, selection
requirements, and consistency between columns. See
:mod:`excel_comparer.checks` for the full catalogue.
"""

from excel_comparer.checks import (
    CHECK_REGISTRY,
    BaselineComparisonCheck,
    Check,
    CheckContext,
    ColumnGroup,
    IrremediableCharacterCheck,
    MinTextLengthCheck,
    RatingBandCheck,
    RequiredCellsCheck,
    SelectionRequiredCheck,
    build_check,
    register_check,
)
from excel_comparer.config import (
    Config,
    Defaults,
    ExportConfig,
    ReportConfig,
    SheetRange,
    SheetRanges,
)
from excel_comparer.exporter import (
    CORRECTED_MARKER,
    FLAGGED_MARKER,
    export_annotated_workbook,
)
from excel_comparer.models import DiffCategory, Difference
from excel_comparer.reporting import FAILED, PASSED, format_differences, summarize, verdict

__all__ = [
    # Checks framework
    "Check",
    "CheckContext",
    "CHECK_REGISTRY",
    "build_check",
    "register_check",
    # Built-in checks
    "BaselineComparisonCheck",
    "IrremediableCharacterCheck",
    "MinTextLengthCheck",
    "RatingBandCheck",
    "RequiredCellsCheck",
    "SelectionRequiredCheck",
    "ColumnGroup",
    # Models
    "DiffCategory",
    "Difference",
    # Config
    "Config",
    "Defaults",
    "ExportConfig",
    "ReportConfig",
    "SheetRange",
    "SheetRanges",
    # Reporting
    "format_differences",
    "summarize",
    "verdict",
    "PASSED",
    "FAILED",
    # Export
    "export_annotated_workbook",
    "FLAGGED_MARKER",
    "CORRECTED_MARKER",
]
