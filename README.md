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

This project uses [uv](https://docs.astral.sh/uv/) for environment and
dependency management. Install uv first (see the upstream docs), then
sync the project:

```bash
uv sync
```

This requires Python 3.11+ and installs `openpyxl` as the only runtime
dependency. To include the web UI extras, run `uv sync --extra web`.

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
uv run python main.py golden.xlsx candidate1.xlsx candidate2.xlsx
```

or, using the installed entry point:

```bash
uv run excel-comparer golden.xlsx candidate1.xlsx candidate2.xlsx --float-tol 1e-9
```

Exit code = number of compared files that had at least one difference.

## Self-test

A small end-to-end self-test creates two sample workbooks in a temp
directory and prints the resulting report:

```bash
uv run python selftest.py
```

## Web UI (Docker)

A small FastAPI-based frontend is available under `webapp/`. It lets you
upload a golden workbook plus one or more candidates from the browser and
shows a grouped, colour-coded diff report with downloadable JSON / text
reports.

Build and run with Docker:

```bash
docker build -t excel-comparer-web .
docker run --rm -p 8000:8000 excel-comparer-web
```

Or with Docker Compose:

```bash
docker compose up --build
```

Then open <http://localhost:8000>.

Other useful endpoints:

- `GET /docs` – auto-generated OpenAPI / Swagger UI
- `POST /api/compare` – multipart form with fields `golden`,
  `others` (repeatable), `float_tol`, `categories` (repeatable)

To run it locally without Docker:

```bash
uv sync --extra web
uv run uvicorn webapp.main:app --reload
```
