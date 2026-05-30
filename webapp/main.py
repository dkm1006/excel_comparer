"""FastAPI web frontend for excel_comparer.

Exposes:
  - ``GET  /``            – HTML upload page
  - ``POST /api/compare`` – multipart upload (golden + 1..N candidates) → JSON diff report
  - ``GET  /healthz``     – simple healthcheck
"""

import logging
import tempfile
from collections import Counter
from dataclasses import asdict
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from excel_comparer import ExcelComparer, format_differences
from excel_comparer.models import DiffCategory, Difference


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

MAX_FILE_BYTES = 25 * 1024 * 1024  # 25 MB per file
ALLOWED_EXTENSIONS = {".xlsx"}

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES = Jinja2Templates(directory=str(BASE_DIR / "templates"))

logger = logging.getLogger("excel_comparer.webapp")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

app = FastAPI(
    title="Excel Comparer",
    description="Compare Excel workbooks against a golden reference file.",
    version="0.1.0",
)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.get("/healthz")
def healthz() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/")
def index(request: Request):
    return TEMPLATES.TemplateResponse(
        request,
        "index.html",
        {
            "categories": [c.value for c in DiffCategory],
            "default_categories": [c.value for c in ExcelComparer.DEFAULT_DIFF_CATEGORIES],
        },
    )


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------

@app.post("/api/compare")
async def api_compare(
    golden: UploadFile = File(..., description="Golden reference .xlsx file"),
    others: list[UploadFile] = File(..., description="One or more candidate .xlsx files"),
    float_tol: float = Form(0.0),
    categories: list[str] | None = Form(None),
) -> JSONResponse:
    # ---- validate inputs ---------------------------------------------------
    _validate_upload(golden, "golden")
    if not others:
        raise HTTPException(status_code=400, detail="At least one candidate file is required.")
    for f in others:
        _validate_upload(f, f.filename or "candidate")

    if float_tol < 0:
        raise HTTPException(status_code=400, detail="float_tol must be >= 0.")

    diff_categories = _parse_categories(categories)

    # ---- run comparison in a temp dir --------------------------------------
    with tempfile.TemporaryDirectory(prefix="excel_comparer_") as tmpdir:
        tmp = Path(tmpdir)
        golden_path = await _save_upload(golden, tmp / f"golden__{_safe_name(golden.filename)}")

        candidate_paths: list[tuple[str, Path]] = []
        for idx, f in enumerate(others):
            safe = _safe_name(f.filename or f"candidate_{idx}.xlsx")
            dest = tmp / f"cand{idx}__{safe}"
            await _save_upload(f, dest)
            candidate_paths.append((f.filename or safe, dest))

        try:
            comparer = ExcelComparer(
                golden_path=golden_path,
                float_tolerance=float_tol,
                diff_categories=diff_categories,
            )
        except Exception as exc:  # noqa: BLE001 - bubble up user-facing error
            logger.exception("Failed to load golden workbook")
            raise HTTPException(status_code=400, detail=f"Could not open golden file: {exc}") from exc

        results: list[dict[str, Any]] = []
        for original_name, path in candidate_paths:
            try:
                diff_map = comparer.compare(path)
                diffs = next(iter(diff_map.values()))
            except Exception as exc:  # noqa: BLE001
                logger.exception("Failed to compare %s", original_name)
                results.append(
                    {
                        "filename": original_name,
                        "error": f"Could not compare file: {exc}",
                        "summary": {},
                        "diffs": [],
                        "text_report": "",
                    }
                )
                continue

            results.append(
                {
                    "filename": original_name,
                    "error": None,
                    "summary": _summarise(diffs),
                    "diffs": [_serialise_diff(d) for d in diffs],
                    "text_report": format_differences(diffs),
                }
            )

    return JSONResponse(
        {
            "golden_filename": golden.filename,
            "float_tol": float_tol,
            "categories": [c.value for c in diff_categories],
            "results": results,
        }
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _validate_upload(f: UploadFile, label: str) -> None:
    if not f or not f.filename:
        raise HTTPException(status_code=400, detail=f"Missing file: {label}.")
    suffix = Path(f.filename).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f"File '{f.filename}' has unsupported extension '{suffix}'. Only .xlsx is allowed.",
        )


def _safe_name(name: str | None) -> str:
    if not name:
        return "file.xlsx"
    # strip any path components, keep just the basename
    return Path(name).name.replace("/", "_").replace("\\", "_")


async def _save_upload(f: UploadFile, dest: Path) -> Path:
    written = 0
    with dest.open("wb") as out:
        while True:
            chunk = await f.read(1024 * 1024)
            if not chunk:
                break
            written += len(chunk)
            if written > MAX_FILE_BYTES:
                raise HTTPException(
                    status_code=413,
                    detail=(
                        f"File '{f.filename}' exceeds the {MAX_FILE_BYTES // (1024 * 1024)} MB limit."
                    ),
                )
            out.write(chunk)
    return dest


def _parse_categories(raw: list[str] | None) -> tuple[DiffCategory, ...]:
    if not raw:
        return ExcelComparer.DEFAULT_DIFF_CATEGORIES
    # FastAPI may pass a single comma-joined string or a list of strings
    flat: list[str] = []
    for item in raw:
        flat.extend(part.strip() for part in item.split(",") if part.strip())
    try:
        result = tuple(DiffCategory(c) for c in flat)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"Unknown diff category: {exc}") from exc
    return result or ExcelComparer.DEFAULT_DIFF_CATEGORIES


def _summarise(diffs: list[Difference]) -> dict[str, int]:
    counts = Counter(d.category.value for d in diffs)
    summary = {c.value: 0 for c in DiffCategory}
    summary.update(counts)
    summary["TOTAL"] = len(diffs)
    return summary


def _serialise_diff(d: Difference) -> dict[str, Any]:
    raw = asdict(d)
    # category is a StrEnum; asdict already gives the string value, but be explicit
    raw["category"] = d.category.value
    raw["message"] = d.message
    # ensure JSON-safe golden/other values
    raw["golden"] = _json_safe(d.golden)
    raw["other"] = _json_safe(d.other)
    return raw


def _json_safe(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)
