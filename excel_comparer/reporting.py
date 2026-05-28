"""Human-readable rendering of :class:`Difference` lists."""

from collections import defaultdict
from typing import Iterable

from excel_comparer.models import DiffCategory, Difference


def format_differences(diffs: Iterable[Difference]) -> str:
    """Render a list of :class:`Difference` objects as a grouped text report."""
    diffs = list(diffs)
    if not diffs:
        return "No differences."

    # Split workbook-level (sheet missing/extra) from per-sheet diffs.
    workbook_level: list[Difference] = []
    by_sheet: dict[str, list[Difference]] = defaultdict(list)
    for d in diffs:
        if d.category in (DiffCategory.SHEET_MISSING, DiffCategory.SHEET_EXTRA):
            workbook_level.append(d)
        else:
            by_sheet[d.sheet].append(d)

    lines: list[str] = []
    for d in workbook_level:
        lines.append(d.message)

    for sheet in sorted(by_sheet):
        lines.append(f'Sheet "{sheet}"')
        for d in by_sheet[sheet]:
            lines.append(f"  {d.message}")

    return "\n".join(lines)
