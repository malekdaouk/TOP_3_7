"""Custom exceptions for YCharts integration."""


class YChartsAuthError(Exception):
    """Raised for 401 responses or missing API key."""


class YChartsAPIError(Exception):
    """Raised for non-auth YCharts API failures."""


class FormulaParseError(Exception):
    """Raised when a YCharts formula string cannot be parsed."""
