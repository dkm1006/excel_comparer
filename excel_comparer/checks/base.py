"""Pluggable checks: protocol, shared context, and registry.

A *check* is anything that, given a candidate workbook, can return a list
of :class:`Difference` findings. The :class:`Check` protocol formalises
this contract:

* a class-level ``name`` (used as the key in the TOML config and the
  ``check_name`` tag on every emitted :class:`Difference`),
* a class-level ``description`` (shown in reports / the future UI),
* a ``run(ctx)`` method returning ``list[Difference]``.

Checks are stateless with respect to the workbook: they receive a
fresh :class:`CheckContext` per candidate file and **must not mutate**
``ctx.workbook``. Auto-corrections are reported as the ``correction``
field on a :class:`Difference` and applied by the exporter.

A check's TOML sub-table is passed directly to its ``__init__`` as
keyword arguments, so any per-check parsing (e.g. column letter →
index) lives in the check itself.
"""

from dataclasses import dataclass
from pathlib import Path
from typing import Any, ClassVar, Protocol, runtime_checkable

from openpyxl.workbook import Workbook

from excel_comparer.models import Difference


@dataclass(frozen=True)
class CheckContext:
    """Inputs handed to every check on every candidate workbook.

    Loaded once by the orchestrator (``cli.py``) and shared across all
    configured checks for that candidate, so individual checks don't have
    to re-open the file.
    """
    workbook_path: Path
    workbook: Workbook


@runtime_checkable
class Check(Protocol):
    """Anything that can produce :class:`Difference` findings.

    Concrete check classes set ``name`` / ``description`` as ``ClassVar``\\ s
    so they're available without instantiation (e.g. for UI listings).
    """

    name: ClassVar[str]
    description: ClassVar[str]

    def run(self, ctx: CheckContext) -> list[Difference]: ...


CHECK_REGISTRY: dict[str, Check] = {}


def register_check(cls: Check) -> Check:
    """Class decorator: register a check class under its ``name``.

    Raises ``ValueError`` on duplicate registration so silent overrides
    aren't possible.
    """
    name = getattr(cls, "name", None)
    if not name:
        raise ValueError(f"Check class {cls.__name__} is missing a non-empty 'name' attribute")
    if name in CHECK_REGISTRY:
        existing = CHECK_REGISTRY[name].__name__
        raise ValueError(f"Check name {name!r} already registered by {existing}")
    CHECK_REGISTRY[name] = cls
    return cls


def build_check(name: str, **kwargs: Any) -> Check:
    """Look up a registered check class and build it from TOML kwargs."""
    if name not in CHECK_REGISTRY:
        valid = ", ".join(sorted(CHECK_REGISTRY)) or "(none registered)"
        raise ValueError(f"Unknown check {name!r}. Registered checks: {valid}")
    return CHECK_REGISTRY[name](**kwargs)
