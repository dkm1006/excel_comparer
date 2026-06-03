"""Pluggable checks subpackage.

Importing this package side-effect-registers every built-in check class
in :data:`CHECK_REGISTRY` (via the ``@register_check`` decorator on each
concrete check module).

To add a new built-in check:

1. Add a module under ``excel_comparer/checks/`` that defines a class
   with class-level ``name``/``description`` and a ``run(ctx)`` method,
   decorated with ``@register_check``.
2. Import it from this ``__init__.py`` so the registration runs.
"""

from excel_comparer.checks.base import (
    CHECK_REGISTRY,
    Check,
    CheckContext,
    build_check,
    register_check,
)

# Side-effect imports: each module's @register_check decorator populates
# CHECK_REGISTRY at import time.
from excel_comparer.checks.baseline_comparison import BaselineComparisonCheck
from excel_comparer.checks.irremediable_character import IrremediableCharacterCheck
from excel_comparer.checks.min_text_length import MinTextLengthCheck
from excel_comparer.checks.rating_band import RatingBandCheck
from excel_comparer.checks.required_cells import RequiredCellsCheck
from excel_comparer.checks.selection_required import ColumnGroup, SelectionRequiredCheck

__all__ = [
    # Protocol / infrastructure
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
]
