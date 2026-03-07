"""Parser for YCharts Excel add-in formulas."""

import re
from datetime import datetime
from typing import Optional

from ycharts.exceptions import FormulaParseError
from ycharts.models import YChartsFormula

# Maps Excel function name to a normalized formula type.
FORMULA_TYPE_MAP = {
    "YCS": "security",
    "YCI": "indicator",
    "YCF": "fund",
    "YCSMF": "fund",
    "YCP": "security",
}

FORMULA_PATTERN = re.compile(
    r"=(?P<func>YC\w+)\("
    r'"(?P<ticker>[^:]+):(?P<metric>[^"]+)"'
    r'(?:,\s*"(?P<start>[^"]*)")?'
    r'(?:,\s*"(?P<end>[^"]*)")?'
    r"\)",
    re.IGNORECASE,
)

XLL_FORMULA_PATTERN = re.compile(
    r"=\s*(?:_xll\.)?(?:@)?(?P<func>YC\w+)\((?P<args>.*)\)\s*$",
    re.IGNORECASE,
)


def _normalize_date(date_str: Optional[str]) -> Optional[str]:
    """Normalize date values to ISO format (YYYY-MM-DD)."""
    if not date_str:
        return None
    value = date_str.strip()
    if not value:
        return None
    try:
        return datetime.fromisoformat(value).date().isoformat()
    except ValueError as exc:
        raise FormulaParseError(f"Invalid date format {date_str!r}; expected YYYY-MM-DD.") from exc


def parse_formula(formula_str: str) -> YChartsFormula:
    """Parse a YCharts formula string into structured arguments."""
    if not isinstance(formula_str, str):
        raise FormulaParseError(f"Formula must be a string, got {type(formula_str).__name__}.")

    formula_text = formula_str.strip()

    standard_match = FORMULA_PATTERN.match(formula_text)
    if standard_match:
        func = standard_match.group("func").upper()
        formula_type = FORMULA_TYPE_MAP.get(func)
        if not formula_type:
            raise FormulaParseError(f"Unknown YCharts function: {func}")

        return YChartsFormula(
            formula_type=formula_type,
            ticker=standard_match.group("ticker").strip().upper(),
            metric=standard_match.group("metric").strip().lower(),
            start_date=_normalize_date(standard_match.group("start")),
            end_date=_normalize_date(standard_match.group("end")),
            raw=formula_str,
        )

    xll_match = XLL_FORMULA_PATTERN.match(formula_text)
    if xll_match:
        func = xll_match.group("func").upper()
        formula_type = FORMULA_TYPE_MAP.get(func)
        if not formula_type:
            raise FormulaParseError(f"Unknown YCharts function: {func}")

        args = _split_formula_args(xll_match.group("args"))
        if len(args) < 2:
            raise FormulaParseError(f"Expected at least 2 arguments in formula: {formula_str!r}")

        # In _xll formulas, args are often cell refs (e.g. $F276, CN$1).
        # Keep them raw so caller can resolve references against worksheet context.
        return YChartsFormula(
            formula_type=formula_type,
            ticker=_strip_quotes(args[0]).strip(),
            metric=_strip_quotes(args[1]).strip(),
            raw=formula_str,
        )

    raise FormulaParseError(f"Cannot parse YCharts formula: {formula_str!r}")


def is_ycharts_formula(cell_value: str) -> bool:
    """Return True if a cell string appears to be a YCharts formula."""
    if not cell_value:
        return False
    text = str(cell_value).strip()
    return bool(FORMULA_PATTERN.match(text) or XLL_FORMULA_PATTERN.match(text))


def _strip_quotes(value: str) -> str:
    text = value.strip()
    if text.startswith('"') and text.endswith('"') and len(text) >= 2:
        return text[1:-1]
    return text


def _split_formula_args(arg_text: str) -> list[str]:
    args: list[str] = []
    current: list[str] = []
    in_quotes = False

    for ch in arg_text:
        if ch == '"':
            in_quotes = not in_quotes
            current.append(ch)
        elif ch == "," and not in_quotes:
            args.append("".join(current).strip())
            current = []
        else:
            current.append(ch)

    if current:
        args.append("".join(current).strip())
    return args
