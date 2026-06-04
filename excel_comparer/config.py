"""TOML-based configuration for excel_comparer.

The config governs:

* which :class:`Check` instances run (the ``[[checks]]`` array of tables),
* the Excel export (output path, highlight color, per-sheet documentation
  column placement, etc.),
* a few cosmetic flags for the text report.

Loading is lenient: any missing section/key falls back to a sensible
default, so an empty TOML file (or no file at all) yields a config with
*no* checks (and the CLI will report nothing). To run the historical
behaviour, include a ``[[checks]]`` entry with ``name = "baseline_comparison"``.
"""

import re
import sys
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Any, Iterator

if sys.version_info >= (3, 11):
    import tomllib  # type: ignore[import-not-found]
else:  # pragma: no cover
    import tomli as tomllib  # type: ignore[import-not-found, no-redef]

from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import range_boundaries


# --------------------------------------------------------------------------- #
# Dataclasses                                                                 #
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class SheetRange:
    """Restricts which cells of a sheet a check looks at.

    ``None`` for any bound means "use the sheet's natural extent".

    ``skip=True`` short-circuits the whole sheet.
    """

    min_row: int | None = None
    max_row: int | None = None
    min_col: int | None = None        # 1-based, inclusive
    max_col: int | None = None        # 1-based, inclusive
    skip: bool = False

    def row_in_range(self, row: int) -> bool:
        """Return ``True`` iff *row* lies within the configured row bounds."""
        if self.skip:
            return False
        if self.min_row is not None and row < self.min_row:
            return False
        if self.max_row is not None and row > self.max_row:
            return False
        return True

    def col_in_range(self, col: int) -> bool:
        """Return ``True`` iff *col* lies within the configured column bounds."""
        if self.skip:
            return False
        if self.min_col is not None and col < self.min_col:
            return False
        if self.max_col is not None and col > self.max_col:
            return False
        return True

    def contains(self, row: int, col: int) -> bool:
        """Return ``True`` iff ``(row, col)`` lies inside the configured range."""
        return self.row_in_range(row) and self.col_in_range(col)

    def clamp_rows(self, start: int, end: int) -> tuple[int, int]:
        """Intersect ``[start, end]`` (inclusive) with the configured row bounds.

        Returns the clamped ``(start, end)`` tuple. If the intersection is
        empty the returned ``start`` will be greater than ``end`` (callers
        typically just iterate ``range(start, end + 1)`` and get an empty
        loop in that case).
        """
        lo = start if self.min_row is None else max(start, self.min_row)
        hi = end if self.max_row is None else min(end, self.max_row)
        return lo, hi

    def clamp_cols(self, start: int, end: int) -> tuple[int, int]:
        """Intersect ``[start, end]`` (inclusive) with the configured column bounds.

        Returns the clamped ``(start, end)`` tuple. If the intersection is
        empty the returned ``start`` will be greater than ``end``.
        """
        lo = start if self.min_col is None else max(start, self.min_col)
        hi = end if self.max_col is None else min(end, self.max_col)
        return lo, hi


@dataclass(frozen=True)
class SheetRanges:
    """A set of zero or more :class:`SheetRange` rectangles on a single sheet.

    A cell is considered "in range" iff it lies inside **any** contained
    sub-range. An empty ``ranges`` tuple means *no restriction* (every
    cell is in range). ``skip=True`` short-circuits the whole sheet
    regardless of the contained sub-ranges.

    Use :meth:`parse` as the single entry point for coercing whatever the
    TOML loader produced (string, list, dict, ``SheetRange``,
    ``SheetRanges`` or ``None``) into a ``SheetRanges`` instance.
    """

    ranges: tuple[SheetRange, ...] = ()
    skip: bool = False

    # ---------------------------- constructors --------------------------- #

    @classmethod
    def parse(cls, spec: Any) -> "SheetRanges":
        """Coerce *spec* into a ``SheetRanges`` instance.

        Accepts:

        * ``None`` — unrestricted (every cell in range).
        * a :class:`SheetRanges` — returned unchanged.
        * a :class:`SheetRange` — wrapped in a single-element tuple.
        * a ``dict`` with any subset of
          ``{min_row, max_row, min_col, max_col, skip}`` — the legacy
          single-range form (used to mark a sheet as skipped via
          ``{ skip = true }``). ``min_col`` / ``max_col`` may be column
          letters or 1-based ints.
        * a ``str`` — one or more comma-separated A1-style specs (see
          :meth:`_parse_a1_text`).
        * a ``list`` — every item is coerced and the results unioned.
        """
        if spec is None:
            return cls()
        if isinstance(spec, SheetRanges):
            return spec
        if isinstance(spec, SheetRange):
            return cls(ranges=(spec,))
        if isinstance(spec, str):
            return cls(ranges=cls._parse_a1_text(spec))
        if isinstance(spec, dict):
            return cls._parse_dict(spec)
        if isinstance(spec, (list, tuple)):
            collected: list[SheetRange] = []
            skip = False
            for item in spec:
                sub = cls.parse(item)
                if sub.skip:
                    skip = True
                collected.extend(sub.ranges)
            return cls(ranges=tuple(collected), skip=skip)
        raise ValueError(
            f"Cannot coerce {type(spec).__name__} into SheetRanges: {spec!r}"
        )

    @classmethod
    def _parse_dict(cls, spec: dict[str, Any]) -> "SheetRanges":
        """Parse the legacy ``{ min_row=..., max_col=..., skip=... }`` form."""
        allowed = {"min_row", "max_row", "min_col", "max_col", "skip"}
        extra = set(spec) - allowed
        if extra:
            raise ValueError(
                f"Unknown keys in range table: {sorted(extra)} "
                f"(allowed: {sorted(allowed)})"
            )
        skip = bool(spec.get("skip", False))
        sub = SheetRange(
            min_row=_parse_row(spec.get("min_row")),
            max_row=_parse_row(spec.get("max_row")),
            min_col=_parse_col(spec.get("min_col")),
            max_col=_parse_col(spec.get("max_col")),
            skip=False,  # skip lives on the outer SheetRanges
        )
        # If the dict only said `skip = true` with no bounds, don't add an
        # empty SheetRange (which would still match everything); just flag skip.
        empty = (
            sub.min_row is None
            and sub.max_row is None
            and sub.min_col is None
            and sub.max_col is None
        )
        return cls(ranges=() if empty else (sub,), skip=skip)

    # A1 syntax: "A1:B2", "A:D", "3:10", single cell "A1", single col "C",
    # single row "5". Comma-separated tokens are unioned.
    _A1_FULL = re.compile(r"^[A-Z]+\d+:[A-Z]+\d+$")
    _A1_SINGLE_CELL = re.compile(r"^([A-Z]+)(\d+)$")
    _A1_COL_RANGE = re.compile(r"^([A-Z]+):([A-Z]+)$")
    _A1_SINGLE_COL = re.compile(r"^([A-Z]+)$")
    _A1_ROW_RANGE = re.compile(r"^(\d+):(\d+)$")
    _A1_SINGLE_ROW = re.compile(r"^(\d+)$")

    @classmethod
    def _parse_a1_text(cls, text: str) -> tuple[SheetRange, ...]:
        """Parse one or more comma-separated A1 specs into ``SheetRange``s.

        Accepted token shapes (case-insensitive, whitespace-tolerant):

        =========  ======================================================
        ``A1:B2``  rectangle (rows + columns bounded)
        ``A1``     single cell
        ``A:D``    whole-column range (rows unbounded)
        ``C``      single column
        ``3:10``   whole-row range (columns unbounded)
        ``5``      single row
        =========  ======================================================
        """
        tokens = [t.strip().upper() for t in text.split(",") if t.strip()]
        if not tokens:
            return ()
        out: list[SheetRange] = []
        for token in tokens:
            out.append(cls._parse_a1_token(token))
        return tuple(out)

    @classmethod
    def _parse_a1_token(cls, token: str) -> SheetRange:
        if cls._A1_FULL.match(token):
            min_col, min_row, max_col, max_row = range_boundaries(token)
            return SheetRange(
                min_row=min_row, max_row=max_row,
                min_col=min_col, max_col=max_col,
            )
        if m := cls._A1_SINGLE_CELL.match(token):
            col = column_index_from_string(m.group(1))
            row = int(m.group(2))
            return SheetRange(
                min_row=row, max_row=row, min_col=col, max_col=col,
            )
        if m := cls._A1_COL_RANGE.match(token):
            lo = column_index_from_string(m.group(1))
            hi = column_index_from_string(m.group(2))
            if hi < lo:
                lo, hi = hi, lo
            return SheetRange(min_col=lo, max_col=hi)
        if m := cls._A1_SINGLE_COL.match(token):
            col = column_index_from_string(m.group(1))
            return SheetRange(min_col=col, max_col=col)
        if m := cls._A1_ROW_RANGE.match(token):
            lo, hi = int(m.group(1)), int(m.group(2))
            if hi < lo:
                lo, hi = hi, lo
            return SheetRange(min_row=lo, max_row=hi)
        if m := cls._A1_SINGLE_ROW.match(token):
            row = int(m.group(1))
            return SheetRange(min_row=row, max_row=row)
        raise ValueError(f"Cannot parse A1 range token: {token!r}")

    # ----------------------------- query API ----------------------------- #

    def row_in_range(self, row: int) -> bool:
        """Return ``True`` iff *row* is inside any contained sub-range."""
        if self.skip:
            return False
        if not self.ranges:
            return True
        return any(r.row_in_range(row) for r in self.ranges)

    def col_in_range(self, col: int) -> bool:
        """Return ``True`` iff *col* is inside any contained sub-range."""
        if self.skip:
            return False
        if not self.ranges:
            return True
        return any(r.col_in_range(col) for r in self.ranges)

    def contains(self, row: int, col: int) -> bool:
        """Return ``True`` iff ``(row, col)`` is inside any contained sub-range."""
        if self.skip:
            return False
        if not self.ranges:
            return True
        return any(r.contains(row, col) for r in self.ranges)

    def clamp_rows(self, start: int, end: int) -> tuple[int, int]:
        """Bounding-box clamp across all sub-ranges (the union's row span).

        Returns ``(lo, hi)`` such that every row of every sub-range that
        also lies in ``[start, end]`` is covered. Per-row containment
        should still be tested with :meth:`row_in_range` when sub-ranges
        are disjoint. If the result is empty, ``lo > hi``.
        """
        if self.skip:
            return (start, start - 1)
        if not self.ranges:
            return (start, end)
        clamped = [r.clamp_rows(start, end) for r in self.ranges]
        return min(lo for lo, _ in clamped), max(hi for _, hi in clamped)

    def clamp_cols(self, start: int, end: int) -> tuple[int, int]:
        """Bounding-box clamp across all sub-ranges (the union's column span)."""
        if self.skip:
            return (start, start - 1)
        if not self.ranges:
            return (start, end)
        clamped = [r.clamp_cols(start, end) for r in self.ranges]
        return min(lo for lo, _ in clamped), max(hi for _, hi in clamped)

    def iter_cells(
        self,
        *,
        row_floor: int,
        row_ceiling: int,
        col_floor: int = 1,
        col_ceiling: int | None = None,
    ) -> Iterator[tuple[int, int]]:
        """Yield every ``(row, col)`` pair inside the union of sub-ranges.

        Each sub-range is intersected with ``[row_floor, row_ceiling]`` and
        ``[col_floor, col_ceiling]`` first; unbounded sides of a sub-range
        inherit the corresponding floor/ceiling. Useful for "scan every
        cell the user asked about" checks.

        When the union contains overlapping sub-ranges, cells in the
        overlap are yielded once (a small ``set`` deduplicates them).
        When ``ranges`` is empty or ``skip=True``, nothing is yielded —
        callers should treat that as a no-op.
        """
        if self.skip or not self.ranges:
            return
        seen: set[tuple[int, int]] = set()
        for r in self.ranges:
            r_lo = row_floor if r.min_row is None else max(row_floor, r.min_row)
            r_hi = row_ceiling if r.max_row is None else min(row_ceiling, r.max_row)
            c_lo = col_floor if r.min_col is None else max(col_floor, r.min_col)
            c_hi = (
                (col_ceiling if r.max_col is None else min(col_ceiling, r.max_col))
                if col_ceiling is not None
                else r.max_col
            )
            if c_hi is None:
                raise ValueError(
                    "iter_cells: sub-range has unbounded max_col and no col_ceiling "
                    "was provided; pass col_ceiling=ws.max_column to enumerate cells."
                )
            for row in range(r_lo, r_hi + 1):
                for col in range(c_lo, c_hi + 1):
                    if (row, col) not in seen:
                        seen.add((row, col))
                        yield row, col


@dataclass(frozen=True)
class ExportConfig:
    """Settings for the annotated Excel export."""

    enabled: bool = True
    output_suffix: str = "_review"
    output_dir: str | None = None
    highlight_color: str = "FFFFFF00"           # ARGB yellow
    correction_color: str = "FFB6FFB6"          # ARGB light green
    flag_color: str = "FFFFFF00"                # ARGB bright yellow
    header_row: int = 2                            # 1-based index of the header row to copy styling from
    doc_columns: dict[str, int] = field(default_factory=dict)   # sheet -> 1-based col index of existing Documentation col
    doc_column_header: str = "Review Documentation"
    doc_column_width: float = 60.0
    summary_sheet_name: str = "Comparison Summary"
    add_summary_sheet: bool = True
    apply_corrections: bool = True              # whether to write Difference.correction back into the cell

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ExportConfig":
        doc_col_raw = data.get("doc_columns", {}) or {}
        if not isinstance(doc_col_raw, dict):
            raise ValueError("export.doc_columns must be a table mapping sheet -> column")
        doc_columns = {
            sheet: _parse_col(col) or 0
            for sheet, col in doc_col_raw.items()
        }

        header_row = _parse_row(data.get("header_row")) or 2

        return cls(
            enabled=bool(data.get("enabled", True)),
            output_suffix=str(data.get("output_suffix", "_review")),
            output_dir=data.get("output_dir") or None,
            highlight_color=_normalize_color(data.get("highlight_color", "FFFFFF00")),
            correction_color=_normalize_color(data.get("correction_color", "FFB6FFB6")),
            flag_color=_normalize_color(data.get("flag_color", "FFFFFF00")),
            header_row=header_row,
            doc_columns=doc_columns,
            doc_column_header=str(data.get("doc_column_header", "Review Documentation")),
            doc_column_width=float(data.get("doc_column_width", 60.0)),
            summary_sheet_name=str(data.get("summary_sheet_name", "Comparison Summary")),
            add_summary_sheet=bool(data.get("add_summary_sheet", True)),
            apply_corrections=bool(data.get("apply_corrections", True)),
        )


@dataclass(frozen=True)
class ReportConfig:
    """Settings for the text/console report."""

    show_passed: bool = True
    group_by_sheet: bool = True
    group_by_check: bool = True

    @classmethod
    def from_dict(cls, data: dict[str, bool]) -> "ReportConfig":
        return cls(**data)


@dataclass(frozen=True)
class Defaults:
    """Top-level defaults applied to every check.

    Any value here is injected into a ``[[checks]]`` entry's kwargs at
    config-parse time unless the entry already specifies it. Checks
    therefore never have to consult :class:`Defaults` at runtime.
    """

    header_row: int = 2
    # Optional "row key" column. Row-bound checks skip any row whose key
    # column is empty (intent: rows with no IRO name in column B are
    # treated as unused/blank and shouldn't generate findings). ``None``
    # disables this behaviour.
    row_anchor_column: int | None = None

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Defaults":
        header_row = _parse_row(data.get("header_row")) or 2
        # Treat falsy as "no anchor".
        raw_anchor = data.get("row_anchor_column")
        row_anchor_column = None if not raw_anchor else _parse_col(raw_anchor)
        return cls(
            header_row=header_row,
            row_anchor_column=row_anchor_column,
        )


@dataclass
class Config:
    """Top-level configuration object.

    ``checks`` is a list of *spec dicts* — the ``name`` and the kwargs each
    check needs. The orchestrator (CLI) materialises them into :class:`Check`
    instances via :func:`excel_comparer.checks.build_check`, so that the
    check registry only needs to be imported once at the top of the CLI.
    """

    checks: list[dict[str, Any]] = field(default_factory=list)
    export: ExportConfig = field(default_factory=ExportConfig)
    report: ReportConfig = field(default_factory=ReportConfig)
    defaults: Defaults = field(default_factory=Defaults)

    # ----------------------------- loading -------------------------------- #

    @classmethod
    def load(cls, path: str | Path) -> "Config":
        """Load a config from a TOML file. ``None`` / missing file => defaults."""
        with Path(path).open("rb") as f:
            data = tomllib.load(f)
        return cls.from_dict(data)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Config":
        defaults = Defaults.from_dict(data.get("defaults", {}) or {})

        checks_raw = data.get("checks", []) or []
        if not isinstance(checks_raw, list):
            raise ValueError("'checks' must be a TOML array of tables ([[checks]])")
        checks: list[dict[str, Any]] = []
        for i, entry in enumerate(checks_raw):
            if not isinstance(entry, dict):
                raise ValueError(f"checks[{i}] must be a table")
            if "name" not in entry:
                raise ValueError(f"checks[{i}] is missing the required 'name' key")
            # Copy so subsequent mutation by the orchestrator doesn't bleed into the loaded config.
            spec = dict(entry)
            if "ranges" in spec:
                spec["ranges"] = _coerce_ranges_map(
                    spec["ranges"], where=f"checks[{i}].ranges"
                )
            # Inject top-level defaults for any key the check spec didn't supply.
            for f in fields(Defaults):
                spec.setdefault(f.name, getattr(defaults, f.name))
            checks.append(spec)

        export_cfg = ExportConfig.from_dict(data.get("export", {}))
        report_cfg = ReportConfig.from_dict(data.get("report", {}))

        return cls(
            checks=checks,
            export=export_cfg,
            report=report_cfg,
            defaults=defaults,
        )


# --------------------------------------------------------------------------- #
# Parsing helpers                                                             #
# --------------------------------------------------------------------------- #


def _parse_col(value: Any) -> int | None:
    """Accept ``"B"``, ``"AA"`` or a 1-based integer; return a 1-based index."""
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError(f"Column value must be a letter or int, got bool: {value!r}")
    if isinstance(value, int):
        if value < 1:
            raise ValueError(f"Column index must be >= 1, got {value}")
        return value
    if isinstance(value, str) and value.strip():
        return column_index_from_string(value.strip().upper())
    raise ValueError(f"Cannot parse column value: {value!r}")


def _parse_row(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError(f"Row value must be an int, got bool: {value!r}")
    if isinstance(value, int):
        if value < 1:
            raise ValueError(f"Row index must be >= 1, got {value}")
        return value
    raise ValueError(f"Cannot parse row value: {value!r}")


def _coerce_ranges_map(
    raw: Any, *, where: str = "ranges"
) -> dict[str, SheetRanges]:
    """Coerce a TOML ``ranges`` table into ``{sheet_name -> SheetRanges}``.

    Every check accepts the same shape: a table mapping sheet name to a
    range specification (string, list, dict, or already a
    :class:`SheetRanges`). Single-sheet checks just have one entry.
    """
    out = {} if raw is None else raw
    if not isinstance(out, dict):
        raise ValueError(
            f"{where} must be a table mapping sheet name to a range spec, "
            f"got {type(raw).__name__}: {raw!r}"
        )
    for sheet, spec in out.items():
        try:
            out[sheet] = SheetRanges.parse(spec)
        except ValueError as exc:
            raise ValueError(f"{where}[{sheet!r}]: {exc}") from exc
    return out


def _normalize_color(value: Any) -> str:
    """Normalise a hex color string to 8-char ARGB upper-case (e.g. ``FFFFFF00``)."""
    if value is None:
        return "FFFFFF00"
    s = str(value).strip().lstrip("#").upper()
    if len(s) == 6:
        s = "FF" + s
    if len(s) != 8 or any(c not in "0123456789ABCDEF" for c in s):
        raise ValueError(f"Invalid color value (need 6 or 8 hex chars): {value!r}")
    return s
