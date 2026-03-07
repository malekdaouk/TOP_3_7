"""Cached data fetchers and formula resolver for YCharts formulas."""

from typing import Any, Optional

import streamlit as st

from ycharts.client import get_ycharts_client
from ycharts.models import YChartsFormula


@st.cache_data(ttl=3600)
def fetch_latest_value(ticker: str, metric: str, formula_type: str) -> float | None:
    """Fetch the most recent value for a ticker+metric pair."""
    client = get_ycharts_client()
    if formula_type == "security":
        return client.get_latest_security_metric(ticker, metric)
    if formula_type == "indicator":
        series = client.get_indicator(ticker)
        return series[-1][1] if series else None
    if formula_type == "fund":
        series = client.get_fund_metric(ticker, metric)
        return series[-1][1] if series else None
    return None


@st.cache_data(ttl=3600)
def fetch_series(
    ticker: str,
    metric: str,
    formula_type: str,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
) -> list[list[Any]]:
    """Fetch full time series data."""
    client = get_ycharts_client()
    if formula_type == "security":
        return client.get_security_metric(ticker, metric, start_date, end_date)
    if formula_type == "indicator":
        return client.get_indicator(ticker, start_date, end_date)
    if formula_type == "fund":
        return client.get_fund_metric(ticker, metric, start_date, end_date)
    return []


def resolve_formula(formula: YChartsFormula) -> float | list[list[Any]] | None:
    """Resolve a parsed formula to either latest value or full series."""
    if formula.start_date or formula.end_date:
        return fetch_series(
            formula.ticker,
            formula.metric,
            formula.formula_type,
            formula.start_date,
            formula.end_date,
        )
    return fetch_latest_value(formula.ticker, formula.metric, formula.formula_type)
