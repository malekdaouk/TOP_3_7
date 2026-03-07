"""Singleton YCharts API client for Streamlit apps."""

import os
from typing import Any, Optional

import requests
import streamlit as st

from ycharts.exceptions import YChartsAPIError, YChartsAuthError


class YChartsClient:
    BASE_URL = "https://ycharts.com/api/v3"

    def __init__(self, api_key: str):
        self.session = requests.Session()
        self.session.headers.update(
            {
                "X-APIKEY": api_key,
                "Content-Type": "application/json",
            }
        )

    def _get(self, endpoint: str, params: Optional[dict[str, Any]] = None) -> dict[str, Any]:
        url = f"{self.BASE_URL}{endpoint}"
        response = self.session.get(url, params=params or {}, timeout=30)
        if response.status_code == 401:
            raise YChartsAuthError("Invalid or expired YCharts API key.")
        if not response.ok:
            raise YChartsAPIError(f"YCharts API error {response.status_code}: {response.text}")
        return response.json()

    def get_security_metric(
        self,
        ticker: str,
        metric: str,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
    ) -> list[list[Any]]:
        params: dict[str, str] = {}
        if start_date:
            params["start_date"] = start_date
        if end_date:
            params["end_date"] = end_date
        data = self._get(f"/securities/{ticker}/metrics/{metric}", params=params)
        return data.get("results", {}).get(ticker, {}).get(metric, {}).get("data", [])

    def get_latest_security_metric(self, ticker: str, metric: str) -> float | None:
        series = self.get_security_metric(ticker, metric)
        if not series:
            return None
        return series[-1][1]

    def get_indicator(
        self,
        indicator: str,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
    ) -> list[list[Any]]:
        params: dict[str, str] = {}
        if start_date:
            params["start_date"] = start_date
        if end_date:
            params["end_date"] = end_date
        data = self._get(f"/indicators/{indicator}/data", params=params)
        return data.get("results", {}).get(indicator, {}).get("data", [])

    def get_fund_metric(
        self,
        ticker: str,
        metric: str,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
    ) -> list[list[Any]]:
        params: dict[str, str] = {}
        if start_date:
            params["start_date"] = start_date
        if end_date:
            params["end_date"] = end_date
        data = self._get(f"/mutual_funds/{ticker}/metrics/{metric}", params=params)
        return data.get("results", {}).get(ticker, {}).get(metric, {}).get("data", [])


@st.cache_resource
def get_ycharts_client() -> YChartsClient:
    """Create and cache a singleton YCharts API client."""
    api_key = os.environ.get("YCHARTS_API_KEY")
    if not api_key:
        raise YChartsAuthError("YCHARTS_API_KEY environment variable not set.")
    return YChartsClient(api_key)
