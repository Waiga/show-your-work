"""Building a reference graph from the formulas, and finding loops in it.

Excel refuses to calculate a circular reference and shows zero instead, which is easy
to miss in a large model. The loop can be found without evaluating anything: read which
cells each formula names, then look for a path that returns to where it started.
"""

from __future__ import annotations

import re

from openpyxl.utils import column_index_from_string, get_column_letter

from show_your_work.findings import Finding, Level, Report
from show_your_work.workbook import LoadedWorkbook

# Ceilings that keep a large model from turning into an unbounded graph walk.
MAX_FORMULA_CELLS = 60_000
MAX_RANGE_CELLS = 2_000

_STRING_LITERAL_RE = re.compile(r'"[^"]*"')
_EXTERNAL_RE = re.compile(r"\[[^\]\[]+\]")

# 'My Sheet'!A1:B2  |  Sheet1!A1  |  A1
_REFERENCE_RE = re.compile(
    r"(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?"
    r"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,3}\$?\d+)"
    r"(?::(\$?[A-Za-z]{1,3}\$?\d+))?"
    r"(?![A-Za-z0-9_(])"
)


def _split_ref(ref: str) -> tuple[int, int]:
    """'$B$7' -> (7, 2)."""
    plain = ref.replace("$", "").upper()
    letters = "".join(ch for ch in plain if ch.isalpha())
    digits = "".join(ch for ch in plain if ch.isdigit())
    return int(digits), column_index_from_string(letters)


def _node(sheet: str, row: int, col: int) -> str:
    return f"{sheet}!{get_column_letter(col)}{row}"


def _dependencies(formula: str, home_sheet: str, known_sheets: set[str]) -> set[str]:
    """Cells a formula names, as graph nodes. External references are ignored."""
    body = _STRING_LITERAL_RE.sub('""', formula)
    if _EXTERNAL_RE.search(body):
        body = _EXTERNAL_RE.sub("[]", body)

    deps: set[str] = set()
    for quoted, bare, first, last in _REFERENCE_RE.findall(body):
        sheet_name = quoted or bare or home_sheet
        if sheet_name not in known_sheets:
            continue

        try:
            row1, col1 = _split_ref(first)
            row2, col2 = _split_ref(last) if last else (row1, col1)
        except (ValueError, KeyError):
            continue

        row_lo, row_hi = sorted((row1, row2))
        col_lo, col_hi = sorted((col1, col2))
        if (row_hi - row_lo + 1) * (col_hi - col_lo + 1) > MAX_RANGE_CELLS:
            continue

        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                deps.add(_node(sheet_name, row, col))
    return deps


def _build_graph(book: LoadedWorkbook) -> tuple[dict[str, set[str]], bool]:
    known = set(book.formulas.sheetnames)
    graph: dict[str, set[str]] = {}
    truncated = False

    for sheet in book.formulas.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type != "f" or not isinstance(cell.value, str):
                    continue
                if len(graph) >= MAX_FORMULA_CELLS:
                    truncated = True
                    return graph, truncated
                graph[_node(sheet.title, cell.row, cell.column)] = _dependencies(
                    cell.value, sheet.title, known
                )

    # A dependency on a cell holding no formula cannot continue a loop.
    graph = {node: {d for d in deps if d in graph} for node, deps in graph.items()}
    return graph, truncated


def _find_cycles(graph: dict[str, set[str]]) -> list[list[str]]:
    """Every distinct loop reachable in the graph, each reported once."""
    WHITE, GREY, BLACK = 0, 1, 2
    colour = dict.fromkeys(graph, WHITE)
    cycles: list[list[str]] = []
    seen_signatures: set[frozenset] = set()

    for start in graph:
        if colour[start] != WHITE:
            continue
        stack: list[tuple[str, iter]] = [(start, iter(sorted(graph[start])))]
        path: list[str] = [start]
        colour[start] = GREY

        while stack:
            node, children = stack[-1]
            advanced = False
            for child in children:
                if colour.get(child, BLACK) == GREY:
                    loop = path[path.index(child) :]
                    signature = frozenset(loop)
                    if signature not in seen_signatures:
                        seen_signatures.add(signature)
                        cycles.append(loop)
                elif colour.get(child, BLACK) == WHITE:
                    colour[child] = GREY
                    path.append(child)
                    stack.append((child, iter(sorted(graph[child]))))
                    advanced = True
                    break
            if not advanced:
                colour[node] = BLACK
                stack.pop()
                path.pop()
    return cycles


def circular_references(book: LoadedWorkbook, report: Report) -> None:
    """Formulas that depend, directly or through other cells, on themselves."""
    graph, truncated = _build_graph(book)
    if truncated:
        report.note_unchecked(
            "Circular references in the largest sheets",
            f"The workbook holds more than {MAX_FORMULA_CELLS:,} formulas, so the "
            "reference graph was cut short and later formulas were not traced.",
        )

    iterative = bool(getattr(getattr(book.formulas, "calculation", None), "iterate", False))
    if iterative:
        report.add(
            Finding(
                check="iterative_calculation",
                level=Level.MEDIUM,
                sheet="(workbook)",
                location="calculation settings",
                summary="Iterative calculation is switched on",
                detail=(
                    "Excel only needs this setting when formulas depend on each other in a "
                    "loop. The results then depend on how many passes Excel was told to run, "
                    "so the same file can produce different numbers on another machine."
                ),
            )
        )

    for loop in _find_cycles(graph):
        shown = " -> ".join(loop[:4]) + (" -> ..." if len(loop) > 4 else f" -> {loop[0]}")
        sheet = loop[0].split("!", 1)[0]
        report.add(
            Finding(
                check="circular_reference",
                level=Level.HIGH,
                sheet=sheet,
                location=loop[0].split("!", 1)[1],
                summary=(
                    "Cell depends on itself"
                    if len(loop) == 1
                    else f"{len(loop)} cells depend on each other in a loop"
                ),
                detail=(
                    f"Chain: {shown}. Excel cannot resolve a loop and normally shows zero, "
                    "so a total built on these cells is silently wrong rather than flagged."
                ),
            )
        )
