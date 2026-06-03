"""Command-line interface for excel_comparer."""

import argparse
import sys
from pathlib import Path
from typing import Sequence

from openpyxl import load_workbook

from excel_comparer.checks import (
    BaselineComparisonCheck,
    Check,
    CheckContext,
    build_check,
)
from excel_comparer.config import Config
from excel_comparer.exporter import export_annotated_workbook
from excel_comparer.models import Difference
from excel_comparer.reporting import format_differences


DEFAULT_CONFIG_FILENAME = "excel_comparer.toml"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-comparer",
        description=(
            "Run a configured pipeline of checks against one or more Excel "
            "workbooks. The baseline (golden) comparison is just one of the "
            "available checks; others validate value ranges, completeness, "
            "selection requirements, etc. Configure the pipeline via a TOML "
            "file (see excel_comparer.example.toml). Optionally writes "
            "annotated copies of each compared workbook with highlighted "
            "cells, applied auto-corrections, and a 'Review Documentation' "
            "column."
        ),
    )
    parser.add_argument(
        "baseline",
        help=(
            "Path to the baseline (golden) .xlsx file. Used by any "
            "baseline_comparison check that doesn't set its own baseline_path."
        ),
    )
    parser.add_argument(
        "candidates",
        nargs="+",
        help="One or more .xlsx files to run the configured checks against.",
    )
    parser.add_argument(
        "--config",
        default=None,
        help=(
            f"Path to a TOML config file. If omitted, "
            f"./{DEFAULT_CONFIG_FILENAME} is used when present; otherwise "
            "built-in defaults apply."
        ),
    )
    parser.add_argument(
        "--no-export",
        action="store_true",
        help="Disable Excel export even if enabled in the config.",
    )
    parser.add_argument(
        "--export",
        action="store_true",
        help="Force enable Excel export even if disabled in the config.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    config_path = _resolve_config_path(args.config)
    config = Config.load(config_path)

    checks: list[Check] = _build_checks(config, baseline_path=args.baseline)

    export_enabled = config.export.enabled
    if args.no_export:
        export_enabled = False
    if args.export:
        export_enabled = True

    files_with_diffs = 0
    for candidate in args.candidates:
        diffs = _run_checks_on_candidate(checks, candidate)
        report = format_differences(
            diffs,
            file_label=candidate,
            show_passed=config.report.show_passed,
            group_by_sheet=config.report.group_by_sheet,
            group_by_check=config.report.group_by_check,
        )
        if report:
            print(report)
            print()

        if diffs:
            files_with_diffs += 1

        if export_enabled:
            out_path = export_annotated_workbook(candidate, diffs, config.export)
            print(f"  -> annotated copy written to: {out_path}")
            print()

    return files_with_diffs


# --------------------------------------------------------------------------- #
# Internals                                                                   #
# --------------------------------------------------------------------------- #


def _build_checks(config: Config, *, baseline_path: str) -> list[Check]:
    """Instantiate every configured check, injecting the CLI baseline path."""
    checks: list[Check] = []
    for spec in config.checks:
        spec = dict(spec)  # don't mutate the config
        name = spec.pop("name")
        check = build_check(name, **spec)
        if isinstance(check, BaselineComparisonCheck) and check.baseline_path is None:
            check.baseline_path = Path(baseline_path)
        checks.append(check)
    return checks


def _run_checks_on_candidate(
    checks: Sequence[Check], candidate_path: str
) -> list[Difference]:
    """Load the candidate workbook once and run every check against it."""
    path = Path(candidate_path)
    workbook = load_workbook(filename=str(path), data_only=True)
    ctx = CheckContext(workbook_path=path, workbook=workbook)
    diffs: list[Difference] = []
    for check in checks:
        diffs.extend(check.run(ctx))
    return diffs


def _resolve_config_path(explicit: str | None) -> str | None:
    """Return *explicit* if given; else the default file if it exists."""
    path = Path(explicit) if explicit else Path(DEFAULT_CONFIG_FILENAME)
    return str(path) if path.exists() else None


if __name__ == "__main__":
    sys.exit(main())
