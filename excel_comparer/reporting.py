"""Human-readable rendering of :class:`Difference` lists.

The text report contains, per workbook:

* a top **PASSED / FAILED** verdict line with per-category counts,
* optionally one block per check (``group_by_check=True``) so failures
  attributable to one check don't drown out another's,
* workbook-level diffs (sheet missing / extra),
* per-sheet sections that list every cell-level diff prefixed with its
  row number.
"""

from collections import Counter, defaultdict
from typing import Iterable

from excel_comparer.models import DiffCategory, Difference


PASSED = "PASSED"
FAILED = "FAILED"

# Categories that are workbook-level (don't reference a sheet's data area).
_WORKBOOK_LEVEL_CATEGORIES: frozenset[DiffCategory] = frozenset({
    DiffCategory.SHEET_MISSING,
    DiffCategory.SHEET_EXTRA,
})


def verdict(diffs: Iterable[Difference]) -> str:
    """Return ``"PASSED"`` if *diffs* is empty, else ``"FAILED"``."""
    return PASSED if not list(diffs) else FAILED


def summarize(diffs: Iterable[Difference]) -> dict[DiffCategory, int]:
    """Count differences per :class:`DiffCategory`."""
    counter: Counter[DiffCategory] = Counter(d.category for d in diffs)
    return dict(counter)


def format_differences(
    diffs: Iterable[Difference],
    *,
    file_label: str | None = None,
    show_passed: bool = True,
    group_by_sheet: bool = True,
    group_by_check: bool = True,
) -> str:
    """Render a list of :class:`Difference` objects as a grouped text report.

    Parameters
    ----------
    diffs:
        Iterable of :class:`Difference` objects from a single candidate workbook
        (typically the union of all checks' findings).
    file_label:
        Optional file path / label to include in the verdict line.
    show_passed:
        When ``True`` (default), still produces a ``PASSED`` block for files
        with no diffs. When ``False``, returns an empty string for passing
        files.
    group_by_sheet:
        When ``True`` (default), cell-level diffs are grouped per sheet and
        ordered by (row, column).
    group_by_check:
        When ``True`` (default), the output is partitioned into one section
        per :attr:`Difference.check_name`, each with its own verdict line.
    """
    diffs = list(diffs)
    lines = _format_top_header(diffs, file_label=file_label)

    if not diffs:
        if not show_passed:
            return ""
        return "\n".join(lines)

    if group_by_check:
        by_check: dict[str, list[Difference]] = defaultdict(list)
        for d in diffs:
            by_check[d.check_name or "(unnamed check)"].append(d)
        for check_name in sorted(by_check):
            lines.append("")
            lines.append(f"--- Check: {check_name} ---")
            lines.extend(_format_check_block(by_check[check_name], group_by_sheet))
    else:
        lines.extend(_format_check_block(diffs, group_by_sheet))

    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# Internal helpers                                                            #
# --------------------------------------------------------------------------- #


def _format_top_header(diffs: list[Difference], *, file_label: str = "") -> list[str]:
    """Top-of-report verdict + per-category count line."""
    if not diffs:
        return [f"{PASSED}: {file_label} no differences."]
    counts = summarize(diffs)
    total = sum(counts.values())
    parts = ", ".join(
        f"{cat.value}: {counts[cat]}"
        for cat in DiffCategory
        if cat in counts
    )
    return [f"{FAILED}: {file_label} {total} difference{'s' if total != 1 else ''} ({parts})"]


def _format_check_block(diffs: list[Difference], group_by_sheet: bool) -> list[str]:
    """Per-check (or whole-file) block: workbook-level diffs + per-sheet rows."""
    workbook_level: list[Difference] = []
    by_sheet: dict[str, list[Difference]] = defaultdict(list)
    for d in diffs:
        if d.category in _WORKBOOK_LEVEL_CATEGORIES:
            workbook_level.append(d)
        else:
            by_sheet[d.sheet].append(d)

    lines: list[str] = []
    for d in workbook_level:
        lines.append(d.message)

    sheet_iter = sorted(by_sheet) if group_by_sheet else list(by_sheet)
    for sheet in sheet_iter:
        lines.append(f'Sheet "{sheet}"')
        for d in sorted(by_sheet[sheet], key=_diff_sort_key):
            lines.append(f"  {d.message}")
    return lines


def _diff_sort_key(d: Difference) -> tuple[int, str, str]:
    """Order by row, then by cell, then by attribute, for stable output."""
    row = d.row if d.row is not None else 0
    cell = d.cell or ""
    attribute = d.attribute or ""
    return (row, cell, attribute)
