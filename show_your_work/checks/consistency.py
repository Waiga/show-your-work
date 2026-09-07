"""Checks for a formula that breaks the pattern its neighbours follow.

The single odd formula in an otherwise uniform block is the error class with the
strongest research behind it, and it is invisible in Excel: the cell looks like every
other cell in the column.
"""

from __future__ import annotations

import re

from show_your_work.checks.formulas import (
    _bounded,
    _contiguous_blocks,
    _is_number_like,
    _iter_lines,
    _shape,
    is_subtotal_like,
)
from show_your_work.findings import Finding, Level, Report
from show_your_work.workbook import LoadedWorkbook

# A pattern needs this many agreeing formulas before a differing one counts as odd.
MIN_MAJORITY = 4
# Above this share of oddities the block simply has no single pattern.
MAX_ODD_SHARE = 0.25

_NUMERIC_TEXT_RE = re.compile(r"^\s*[-+(]?\s*[\d,]*\.?\d+\s*\)?%?\s*$")


def inconsistent_formulas(book: LoadedWorkbook, report: Report) -> None:
    """One formula differing from the many identical formulas around it."""
    seen: set[tuple[str, str]] = set()

    for sheet in book.formulas.worksheets:
        max_row, max_col, _ = _bounded(sheet)
        if max_row == 0 or max_col == 0:
            continue

        for axis, _label, cells in _iter_lines(sheet, max_row, max_col):
            for block in _contiguous_blocks(cells):
                formula_cells = [c for c in block if c.data_type == "f"]
                if len(formula_cells) < MIN_MAJORITY + 1:
                    continue

                # A total closing a column is meant to differ from the column.
                formula_cells = [c for c in formula_cells if not is_subtotal_like(c, block)]
                if len(formula_cells) < MIN_MAJORITY + 1:
                    continue

                shapes: dict[str, list] = {}
                for cell in formula_cells:
                    shapes.setdefault(_shape(str(cell.value)), []).append(cell)

                if len(shapes) < 2:
                    continue

                majority_shape, majority_cells = max(shapes.items(), key=lambda kv: len(kv[1]))
                if len(majority_cells) < MIN_MAJORITY:
                    continue

                odd = [
                    cell
                    for shape, group in shapes.items()
                    if shape != majority_shape
                    for cell in group
                ]
                if len(odd) / len(formula_cells) > MAX_ODD_SHARE:
                    continue

                for cell in odd:
                    key = (sheet.title, cell.coordinate)
                    if key in seen:
                        continue
                    seen.add(key)
                    report.add(
                        Finding(
                            check="inconsistent_formula",
                            level=Level.HIGH,
                            sheet=sheet.title,
                            location=cell.coordinate,
                            summary=f"Formula differs from the {len(majority_cells)} "
                            f"matching formulas in this {axis}",
                            detail=(
                                "A filled column normally repeats one formula. A cell that "
                                "quietly does something else looks identical on screen, and "
                                "is the pattern most often found behind a wrong total."
                            ),
                            sample=str(cell.value),
                        )
                    )


def numbers_stored_as_text(book: LoadedWorkbook, report: Report) -> None:
    """Digits saved as text, which arithmetic silently skips."""
    for sheet in book.formulas.worksheets:
        max_row, max_col, _ = _bounded(sheet)
        if max_row == 0 or max_col == 0:
            continue

        for _axis, label, cells in _iter_lines(sheet, max_row, max_col):
            if not label.isalpha():
                continue  # columns only; a row of mixed types is normal

            numeric = [c for c in cells if _is_number_like(c)]
            if len(numeric) < 3:
                continue

            text_numbers = [
                c
                for c in cells
                if isinstance(c.value, str) and _NUMERIC_TEXT_RE.match(c.value)
            ]
            if not text_numbers:
                continue

            locations = [c.coordinate for c in text_numbers]
            report.add(
                Finding(
                    check="number_stored_as_text",
                    level=Level.MEDIUM,
                    sheet=sheet.title,
                    location=", ".join(locations[:5])
                    + ("" if len(locations) <= 5 else f" (+{len(locations) - 5} more)"),
                    summary=f"{len(text_numbers)} value{'s' if len(text_numbers) > 1 else ''} "
                    "in a numeric column stored as text",
                    detail=(
                        "SUM and lookups skip text, so these are excluded from totals without "
                        "any error appearing. The column still looks correct on screen."
                    ),
                )
            )
