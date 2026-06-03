# excel-comparer

Run a configurable pipeline of **checks** against Excel (`.xlsx`) workbooks.

The historical "compare cell-by-cell against a golden workbook" behaviour
is now just one of several pluggable checks (`baseline_comparison`).
Other built-in checks validate value ranges, completeness, mutually-
exclusive selections, and consistency between columns. New checks are
trivial to add — implement the `Check` protocol and register the class.

## Built-in checks

| `name` (in TOML)         | Class                       | What it does                                                                          |
| ------------------------ | --------------------------- | ------------------------------------------------------------------------------------- |
| `baseline_comparison`    | `BaselineComparisonCheck`   | Cell-by-cell diff vs. a baseline (golden) workbook (VALUE / FORMAT / merged / sheet). |
| `rating_range`           | `RatingRangeCheck`          | Numeric columns must hold a value in `[min_value, max_value]` (optionally integer).   |
| `survey_completion`      | `SurveyCompletionCheck`     | Required cells are non-empty; free-text reasoning meets a minimum length.             |
| `selection_required`     | `SelectionRequiredCheck`    | At least one cell in each configured column group differs from `forbidden_value`.     |
| `irremediable_character` | `IrremediableCharacterCheck`| Score column is consistent with a categorical type column; can auto-correct.          |

## Installation

```bash
pip install -e .
```

Python 3.10+ is required.
`openpyxl` is the only runtime dependency.

## CLI

```bash
python main.py baseline.xlsx candidate1.xlsx candidate2.xlsx
```

or, after `pip install -e .`:

```bash
excel-comparer baseline.xlsx candidate1.xlsx candidate2.xlsx
```

The positional `baseline` argument is used by any `baseline_comparison`
check that doesn't set its own `baseline_path` in the config.

Useful flags:

- `--config path/to/excel_comparer.toml` — explicit config file
  (otherwise `./excel_comparer.toml` is used if it exists).
- `--no-export` / `--export` — force-disable / force-enable the Excel export.

The exit code equals the number of candidate files that have at least
one finding.

### Console output

The text report shows a top verdict, then one block per check:

```
FAILED: candidate.xlsx 10 differences (SHEET_MISSING: 1, SHEET_EXTRA: 1, VALUE: 4, RATING_OUT_OF_RANGE: 1, REASONING_TOO_SHORT: 1, INCONSISTENT_TYPE_SCORE: 2)

--- Check: baseline_comparison ---
SHEET_MISSING: Sheet 'Notes' is missing in other workbook.
SHEET_EXTRA: Sheet 'Extra' is extra in other workbook.
Sheet "Sales"
  Row 3, Cell B3 [VALUE]: golden=3 vs other=9
  Row 4, Cell B4 [VALUE]: golden='-' vs other=4

--- Check: irremediable_character ---
Sheet "Sales"
  Row 3, Cell B3 [INCONSISTENT_TYPE_SCORE]: expected "'Negative Impact' requires integer in [1, 2, 3, 4]", got 9
  Row 4, Cell B4 [INCONSISTENT_TYPE_SCORE]: expected "'Positive Impact' requires '-'", got 4

--- Check: rating_range ---
Sheet "Sales"
  Row 3, Cell B3 [RATING_OUT_OF_RANGE]: expected 'integer in [1, 4]', got 9

--- Check: survey_completion ---
Sheet "Sales"
  Row 3, Cell D3 [REASONING_TOO_SHORT]: expected 'at least 100 chars', got '5 chars'
```

## Configuration (`excel_comparer.toml`)

Every check is a `[[checks]]` entry. The `name` field selects the
implementation; every other key is forwarded as a keyword argument to
the check's constructor. See
[`excel_comparer.example.toml`](excel_comparer.example.toml) for a full
annotated example.

Minimal example (baseline comparison only):

```toml
[[checks]]
name = "baseline_comparison"
float_tolerance = 0.0
categories = ["VALUE", "SHEET_MISSING", "SHEET_EXTRA"]

[checks.ranges."Introduction"]
skip = true

[checks.ranges."Impacts"]
min_row = 3
min_col = "B"
max_col = "U"

[export]
enabled = true
output_suffix = "_review"

[export.doc_columns]
"Impacts" = "U"
```

## Annotated Excel export

When `[export].enabled = true`, each candidate workbook produces a
``<name><output_suffix>.xlsx`` file (next to the input by default, or in
``output_dir``) with:

1. **A "Comparison Summary" sheet** prepended to the workbook:
   - Overall `PASSED` / `FAILED` verdict.
   - Per-check verdict + finding count.
   - Per-sheet table with one column per observed difference category.
   - Workbook-level differences (sheet missing / extra) listed below.
2. **Highlighted cells** — every cell-bound diff is filled with the
   configured ``highlight_color``.
3. **Applied auto-corrections** — diffs that carry a `correction` value
   (currently produced by `irremediable_character`) have their cell
   rewritten to the correction value and are filled with
   ``correction_color`` instead. Set ``apply_corrections = false`` to
   disable.
4. **A new "Review Documentation" column**, inserted immediately to the
   right of the existing Documentation column on each sheet listed in
   ``[export.doc_columns]``. Each row that has at least one row-bound
   diff gets a cell with one ``[FLAGGED <check>] <Difference.message>``
   line per diff, joined by newlines (``[CORRECTED <check>] ...`` for
   diffs that triggered an auto-correction).

`MERGED_CELL` and `SHEET_*` differences are **not** written into the doc
column because they don't reference a specific row.

## Library usage

```python
from openpyxl import load_workbook
from excel_comparer import (
    BaselineComparisonCheck,
    CheckContext,
    RatingRangeCheck,
    export_annotated_workbook,
    format_differences,
)

candidate = "candidate.xlsx"
wb = load_workbook(candidate, data_only=True)
ctx = CheckContext(workbook_path=candidate, workbook=wb)

checks = [
    BaselineComparisonCheck(baseline_path="baseline.xlsx"),
    RatingRangeCheck(sheet="Impacts", columns=["O", "P", "Q"], min_value=1, max_value=4),
]

diffs = [d for c in checks for d in c.run(ctx)]
print(format_differences(diffs, file_label=candidate))
```

### Writing a new check

```python
from typing import ClassVar
from excel_comparer import Check, CheckContext, Difference, DiffCategory, register_check

@register_check
class MyCheck:
    name: ClassVar[str] = "my_check"
    description: ClassVar[str] = "Does something."

    def __init__(self, *, sheet: str, threshold: int = 0) -> None:
        self.sheet = sheet
        self.threshold = threshold

    def run(self, ctx: CheckContext) -> list[Difference]:
        ws = ctx.workbook[self.sheet]
        # ...inspect cells, build Difference objects...
        return []
```

Import the module once (e.g. from your config loader) and the registry
will know about `"my_check"`. Any auto-correction the check wants to
apply should be returned as a `Difference` with the `correction` field
set — the exporter will write the new value back.

## Self-test

A small end-to-end self-test builds two sample workbooks in a temp
directory and exercises every built-in check plus the annotated export:

```bash
python selftest.py
```
