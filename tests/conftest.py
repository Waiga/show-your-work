"""Workbook fixtures, built in code so every test states its own evidence."""

from __future__ import annotations

import pytest
from openpyxl import Workbook


def _save(wb: Workbook, tmp_path, name: str) -> str:
    path = tmp_path / name
    wb.save(path)
    return str(path)


@pytest.fixture
def clean_model(tmp_path) -> str:
    """A well-built sheet: a header, a filled formula column, and a total under it.

    Nothing here should be reported. This fixture is the false-positive guard for
    every check in the suite.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Model"
    ws["A1"], ws["B1"], ws["C1"] = "Units", "Price", "Revenue"
    for row in range(2, 12):
        ws.cell(row, 1, row * 3)
        ws.cell(row, 2, 25)
        ws.cell(row, 3, f"=A{row}*B{row}")
    ws["C12"] = "=SUM(C2:C11)"
    return _save(wb, tmp_path, "clean.xlsx")


@pytest.fixture
def overwritten(tmp_path) -> str:
    """The same sheet with one formula replaced by a typed number."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Model"
    ws["A1"], ws["B1"], ws["C1"] = "Units", "Price", "Revenue"
    for row in range(2, 12):
        ws.cell(row, 1, row * 3)
        ws.cell(row, 2, 25)
        ws.cell(row, 3, f"=A{row}*B{row}")
    ws["C6"] = 4200  # typed over the formula
    ws["C12"] = "=SUM(C2:C11)"
    return _save(wb, tmp_path, "overwritten.xlsx")


@pytest.fixture
def short_total(tmp_path) -> str:
    """A total whose range stops two rows above it."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Model"
    ws["A1"] = "Amount"
    for row in range(2, 10):
        ws.cell(row, 1, row * 5)
    ws["A10"] = "=SUM(A2:A7)"  # misses rows 8 and 9
    return _save(wb, tmp_path, "short_total.xlsx")


@pytest.fixture
def odd_formula(tmp_path) -> str:
    """One cell in a filled column doing something different."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Model"
    for row in range(2, 12):
        ws.cell(row, 1, row)
        ws.cell(row, 2, f"=A{row}*2")
    ws["B7"] = "=A7*3"  # the odd one
    return _save(wb, tmp_path, "odd.xlsx")


@pytest.fixture
def looping(tmp_path) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Loop"
    ws["A1"] = "=B1+1"
    ws["B1"] = "=C1*2"
    ws["C1"] = "=A1"
    return _save(wb, tmp_path, "loop.xlsx")


@pytest.fixture
def hidden_things(tmp_path) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Visible"
    ws["A1"] = 1
    ws.row_dimensions[3].hidden = True
    ws.row_dimensions[4].hidden = True
    ws.row_dimensions[5].hidden = True
    ws.column_dimensions["D"].hidden = True
    wb.create_sheet("Quiet").sheet_state = "hidden"
    wb.create_sheet("Buried").sheet_state = "veryHidden"
    return _save(wb, tmp_path, "hidden.xlsx")


@pytest.fixture
def text_numbers(tmp_path) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Amount"
    for row in range(2, 9):
        ws.cell(row, 1, row * 10)
    ws["A5"] = "1250"  # digits stored as text
    return _save(wb, tmp_path, "text_numbers.xlsx")


@pytest.fixture
def linked(tmp_path) -> str:
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "=[1]Budget!B4"
    return _save(wb, tmp_path, "linked.xlsx")
