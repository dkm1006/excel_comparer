"""Command-line interface for excel_comparer."""

import argparse
import sys
from typing import Sequence

from excel_comparer.comparer import ExcelComparer
from excel_comparer.reporting import format_differences


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-comparer",
        description=(
            "Compare one or more Excel workbooks against a golden reference "
            "workbook and report any differences in values, formatting, and "
            "merged-cell ranges."
        ),
    )
    parser.add_argument("golden", help="Path to the golden .xlsx file")
    parser.add_argument(
        "others",
        nargs="+",
        help="One or more .xlsx files to compare against the golden file",
    )
    parser.add_argument(
        "--float-tol",
        type=float,
        default=0.0,
        help="Absolute tolerance for floating-point cell value comparisons "
             "(default: 0.0, i.e. exact equality).",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    comparer = ExcelComparer(args.golden, float_tolerance=args.float_tol)
    results = comparer.compare(*args.others)

    files_with_diffs = 0
    for path, diffs in results.items():
        print(f"=== {path} ===")
        print(format_differences(diffs))
        print()
        if diffs:
            files_with_diffs += 1

    return files_with_diffs


if __name__ == "__main__":
    sys.exit(main())
