"""Typed models for parsed formulas and API series points."""

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class YChartsFormula:
    formula_type: str
    ticker: str
    metric: str
    start_date: Optional[str] = None
    end_date: Optional[str] = None
    raw: str = ""


@dataclass(frozen=True)
class SeriesPoint:
    date: str
    value: float | None
