"""Load workbooks and extract YCharts formulas."""

from pathlib import Path
from typing import BinaryIO

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

from ycharts.exceptions import FormulaParseError
from ycharts.models import YChartsFormula
from ycharts.parser import is_ycharts_formula, parse_formula


def extract_ycharts_formulas(
    filepath: str | Path | BinaryIO,
    sheet_name: str | None = None,
    min_row: int = 1,
) -> dict[str, YChartsFormula]:
    """
    Open an Excel workbook and return all cells containing YCharts formulas.

    Returns a dict mapping "SheetName!A1" to parsed `YChartsFormula`.
    """
    workbook = openpyxl.load_workbook(filepath, data_only=False)
    formulas: dict[str, YChartsFormula] = {}

    sheets = workbook.worksheets
    if sheet_name:
        sheets = [sheet for sheet in workbook.worksheets if sheet.title == sheet_name]

    for sheet in sheets:
        for row in sheet.iter_rows(min_row=min_row):
            for cell in row:
                value = cell.value
                formula_text = None
                if isinstance(value, ArrayFormula):
                    formula_text = value.text
                elif isinstance(value, str):
                    formula_text = value

                if isinstance(formula_text, str) and is_ycharts_formula(formula_text):
                    key = f"{sheet.title}!{cell.coordinate}"
                    try:
                        formulas[key] = parse_formula(formula_text)
                    except FormulaParseError:
                        # Skip malformed formulas and keep extraction resilient.
                        continue

    return formulas
