# excel-comparer

Compare one or more Excel (`.xlsx`) workbooks against a *golden* reference
workbook and list every difference in:

- **Cell values** — cached computed values (`data_only=True`); no formula
  inspection. Missing, extra, and changed values are all reported under
  the single `VALUE` category.
- **Cell formatting** — font (name, size, bold, italic, underline, color),
  fill (pattern, fg/bg color), number format, alignment (horizontal,
  vertical, wrap), and borders (style + color per side). One diff per
  differing attribute.
- **Merged cell ranges** — per-sheet set comparison.
- **Sheets** — sheets present on only one side are reported as
  `SHEET_MISSING` / `SHEET_EXTRA`.

Matching is **strict**: sheets by exact name, cells by coordinate
(`A1 ↔ A1`).

## Installation

```bash
pip install -e .
```

This requires Python 3.13+ and installs `openpyxl` as the only runtime
dependency.

## Library usage

```python
from excel_comparer import ExcelComparer, format_differences

comparer = ExcelComparer("golden.xlsx")

# Compare one workbook
results = comparer.compare("candidate.xlsx")

# Compare many workbooks
results = comparer.compare("a.xlsx", "b.xlsx", "c.xlsx")

# Or unpack a list
paths = ["a.xlsx", "b.xlsx"]
results = comparer.compare(*paths)

# `results` is always a dict: {path_str: list[Difference]}
for path, diffs in results.items():
    print(f"=== {path} ===")
    print(format_differences(diffs))
```

Each `Difference` is a small dataclass:

```python
@dataclass(frozen=True)
class Difference:
    category: DiffCategory     # SHEET_MISSING | SHEET_EXTRA | VALUE | FORMAT | MERGED_CELL
    sheet: str
    cell: str | None           # e.g. "B7", or None for sheet/merged-range diffs
    attribute: str | None      # e.g. "font.bold"; None when not applicable
    golden: Any
    other: Any
    message: str               # human-readable one-liner
```

You can also pass a tolerance for floating-point value comparisons:

```python
comparer = ExcelComparer("golden.xlsx", float_tolerance=1e-9)
```

## CLI

```bash
python main.py golden.xlsx candidate1.xlsx candidate2.xlsx
```

or, after `pip install -e .`:

```bash
excel-comparer golden.xlsx candidate1.xlsx candidate2.xlsx --float-tol 1e-9
```

Exit code = number of compared files that had at least one difference.

## Self-test

A small end-to-end self-test creates two sample workbooks in a temp
directory and prints the resulting report:

```bash
python selftest.py
```
