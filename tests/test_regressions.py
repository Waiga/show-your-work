"""Cases an adversarial pass found. Each one shipped as a bug once.

Every test here names the wrong behaviour it prevents, so a future change that
reintroduces it fails with an explanation rather than a bare assertion.
"""

from __future__ import annotations

from openpyxl import Workbook

from show_your_work.cli import run


def _book(tmp_path, name: str, build) -> str:
    wb = Workbook()
    build(wb)
    path = tmp_path / name
    wb.save(path)
    return str(path)


def checks(path: str) -> list[str]:
    return [f.check for f in run(path).findings]


def at(path: str, check: str) -> list:
    return [f for f in run(path).findings if f.check == check]


# --- 1. an external reference must not be re-read as a local one ---


def test_external_reference_does_not_invent_a_circular_reference(tmp_path):
    def build(wb):
        first = wb.active
        first.title = "Sheet1"
        second = wb.create_sheet("Sheet2")
        # Stripping only the [Book] part would leave "Sheet2!$B$1", which reads as a
        # local reference and closes a loop that does not exist.
        first["A1"] = "=[OtherBook.xlsx]Sheet2!$B$1"
        second["B1"] = "=Sheet1!A1"

    path = _book(tmp_path, "external.xlsx", build)
    assert "circular_reference" not in checks(path)
    assert "external_link" in checks(path)


def test_a_real_cross_sheet_cycle_is_still_caught(tmp_path):
    def build(wb):
        first = wb.active
        first.title = "A"
        second = wb.create_sheet("B")
        first["A1"] = "=B!B1"
        second["B1"] = "=A!A1"

    assert "circular_reference" in checks(_book(tmp_path, "cycle.xlsx", build))


# --- 2 and 3. a seed value is not an overwritten formula ---


def test_opening_balance_is_not_reported_as_an_overwritten_formula(tmp_path):
    def build(wb):
        ws = wb.active
        ws.title = "Ledger"
        ws["C2"] = 1000  # opening balance, typed on purpose
        for row in range(3, 30):
            ws.cell(row, 2, row * 7)
            ws.cell(row, 3, f"=C{row - 1}+B{row}")

    path = _book(tmp_path, "balance.xlsx", build)
    assert run(path).findings == [], "a running balance is correct spreadsheet practice"


def test_override_in_the_middle_of_a_column_is_still_high(tmp_path):
    def build(wb):
        ws = wb.active
        for row in range(2, 12):
            ws.cell(row, 1, row)
            ws.cell(row, 2, f"=A{row}*2")
        ws["B6"] = 999

    hits = at(_book(tmp_path, "override.xlsx", build), "overwritten_formula")
    assert [f.location for f in hits] == ["B6"]
    assert hits[0].level.value == "high"


def test_override_at_the_edge_is_medium_because_a_typed_edge_is_often_deliberate(tmp_path):
    def build(wb):
        ws = wb.active
        ws["B2"] = 4200  # nothing below refers back to it, so it is not a seed
        for row in range(3, 15):
            ws.cell(row, 1, row)
            ws.cell(row, 2, f"=A{row}*2")

    hits = at(_book(tmp_path, "edge.xlsx", build), "overwritten_formula")
    assert [f.location for f in hits] == ["B2"]
    assert hits[0].level.value == "medium", "README documents the edge case as medium"


# --- 4. stacked section subtotals are correct, not a gap ---


def test_grand_total_over_section_subtotals_is_not_a_gap(tmp_path):
    def build(wb):
        ws = wb.active
        ws.title = "Rollup"
        for row, value in ((1, 10), (2, 20), (3, 30)):
            ws.cell(row, 2, value)
        ws["B4"] = "=SUM(B1:B3)"
        for row, value in ((5, 40), (6, 50), (7, 60)):
            ws.cell(row, 2, value)
        ws["B8"] = "=SUM(B5:B7)"  # correctly excludes the subtotal in B4
        ws["B9"] = "=B4+B8"

    path = _book(tmp_path, "rollup.xlsx", build)
    assert run(path).findings == [], "stacked sections are how a budget is meant to be built"


# --- 5. the horizontal axis is checked too ---


def test_row_sum_that_misses_the_last_column(tmp_path):
    def build(wb):
        ws = wb.active
        ws.title = "Quarters"
        for col, value in ((2, 100), (3, 110), (4, 120), (5, 130)):
            ws.cell(1, col, value)
        ws["F1"] = "=SUM(B1:D1)"  # misses column E

    hits = at(_book(tmp_path, "quarters.xlsx", build), "total_misses_rows")
    assert [f.location for f in hits] == ["F1"]
    assert "column" in hits[0].summary


def test_column_sum_gap_still_found(tmp_path):
    def build(wb):
        ws = wb.active
        for row in range(2, 10):
            ws.cell(row, 1, row * 5)
        ws["A10"] = "=SUM(A2:A7)"

    hits = at(_book(tmp_path, "vertical.xlsx", build), "total_misses_rows")
    assert [f.location for f in hits] == ["A10"]
    assert "2 rows" in hits[0].summary


# --- 6. a total does not have to start with its aggregate ---


def _column_closed_by(total: str):
    def build(wb):
        ws = wb.active
        ws.title = "Model"
        ws["A1"], ws["B1"], ws["C1"] = "Units", "Price", "Revenue"
        for row in range(2, 10):
            ws.cell(row, 1, row * 3)
            ws.cell(row, 2, 25)
            ws.cell(row, 3, f"=A{row}*B{row}")
        ws["C10"] = total

    return build


def test_lotus_plus_prefixed_total_is_not_an_inconsistent_formula(tmp_path):
    # "=+SUM(...)" is the Lotus carry-over idiom and is everywhere in finance
    # workbooks. Anchoring the exemption to the start of the formula reported it,
    # and one public Ontario workbook produced 48 false high findings from this.
    path = _book(tmp_path, "lotus.xlsx", _column_closed_by("=+SUM(C2:C9)"))
    assert run(path).findings == [], "=+SUM is a total, written the way Lotus wrote it"


def test_negated_total_is_not_an_inconsistent_formula(tmp_path):
    path = _book(tmp_path, "negated.xlsx", _column_closed_by("=-SUM(C2:C9)"))
    assert run(path).findings == []


def test_rounded_total_is_not_an_inconsistent_formula(tmp_path):
    path = _book(tmp_path, "rounded.xlsx", _column_closed_by("=ROUND(SUM(C2:C9),0)"))
    assert run(path).findings == [], "a rounded total is still a total"


def test_a_conditional_aggregate_is_not_treated_as_a_plain_total(tmp_path):
    from show_your_work.checks.formulas import _calls_an_aggregate

    # The name must be followed directly by its bracket, or widening the pattern
    # would exempt every filtered aggregate in the workbook as well.
    for conditional in ("=SUMIF(A1:A9,\">5\")", "=SUMIFS(A1:A9,B1:B9,1)",
                        "=COUNTIF(A1:A9,1)", "=COUNTIFS(A1:A9,1)",
                        "=AVERAGEIF(A1:A9,1)"):
        assert not _calls_an_aggregate(conditional), conditional
    for total in ("=SUM(C2:C9)", "=+SUM(C2:C9)", "=-SUM(C2:C9)", "=ROUND(SUM(C2:C9),0)"):
        assert _calls_an_aggregate(total), total


def test_an_aggregate_in_the_middle_of_a_column_is_still_reported(tmp_path):
    # The exemption is widened in what it matches, never in where it applies: a
    # subtotal is still only exempt at either end of the block it closes.
    def build(wb):
        ws = wb.active
        for row in range(2, 12):
            ws.cell(row, 1, row)
            ws.cell(row, 2, f"=A{row}*2")
        ws["B6"] = "=+SUM(A2:A5)"

    hits = at(_book(tmp_path, "mid_aggregate.xlsx", build), "inconsistent_formula")
    assert [f.location for f in hits] == ["B6"], "position still decides, not the text"


# --- 7. a term added outside the range is not a skipped row ---


def test_a_row_added_outside_the_sum_is_not_a_skipped_row(tmp_path):
    # Ireland's Local Authority Budget series writes totals as "=SUM(D93:D103)-D104"
    # where row 104 is a deduction. Reading the formula only as far as its first
    # bracket produced 48 false high findings per edition, across seven editions.
    def build(wb):
        ws = wb.active
        ws.title = "Budget"
        for row in range(2, 5):
            ws.cell(row, 2, row * 10)
        ws["B5"] = 7  # an adjustment line, added on purpose
        ws["B6"] = "=SUM(B2:B4)+B5"

    path = _book(tmp_path, "adjustment.xlsx", build)
    assert run(path).findings == [], "B5 is in the formula, so it is not left out"


def test_a_deduction_named_with_an_absolute_reference_is_not_a_skipped_row(tmp_path):
    def build(wb):
        ws = wb.active
        ws.title = "Budget"
        for row in range(2, 5):
            ws.cell(row, 2, row * 10)
        ws["B5"] = 7
        ws["B6"] = "=SUM(B2:B4)-$B$5"

    assert run(_book(tmp_path, "deduction.xlsx", build)).findings == []


def test_a_gap_row_the_formula_never_names_is_still_reported(tmp_path):
    # The exemption must depend on the formula naming that exact cell, not on the
    # formula merely having a second term.
    def build(wb):
        ws = wb.active
        ws.title = "Budget"
        for row in range(2, 5):
            ws.cell(row, 2, row * 10)
        ws["B5"] = 7
        ws["B6"] = "=SUM(B2:B4)+B99"

    hits = at(_book(tmp_path, "elsewhere.xlsx", build), "total_misses_rows")
    assert [f.location for f in hits] == ["B6"]


# --- 8. a repeated partial aggregate is a design, not a slip ---


def test_a_summary_column_of_partial_sums_is_not_six_broken_totals(tmp_path):
    # A detail table with a summary column that deliberately re-aggregates only part
    # of it. Stopping the scan at the blank separator column does not fix this: the
    # skipped columns sit before the blank, not after it.
    def build(wb):
        ws = wb.active
        ws.title = "Detail"
        for row in range(2, 8):
            for col in range(2, 7):  # B to F
                ws.cell(row, col, row * col)
            ws.cell(row, 8, f"=SUM(B{row}:D{row})")  # H, past a blank column G

    path = _book(tmp_path, "subset.xlsx", build)
    assert run(path).findings == [], "one shape repeated six times is a decision"


def test_a_partial_aggregate_repeated_only_twice_is_still_reported(tmp_path):
    # Repetition is the signal, so it has to have a threshold, and the threshold has
    # to bite. Two of them is not yet a pattern.
    def build(wb):
        ws = wb.active
        ws.title = "Detail"
        for row in range(2, 4):
            for col in range(2, 7):
                ws.cell(row, col, row * col)
            ws.cell(row, 8, f"=SUM(B{row}:D{row})")

    hits = at(_book(tmp_path, "twice.xlsx", build), "total_misses_rows")
    assert [f.location for f in hits] == ["H2", "H3"]


# --- 9. the guards: widening an exemption is how a tool goes quiet ---


def test_a_genuine_mid_column_overwrite_survives_every_widened_exemption(tmp_path):
    def build(wb):
        ws = wb.active
        ws.title = "Model"
        for row in range(2, 12):
            ws.cell(row, 1, row)
            ws.cell(row, 2, f"=A{row}*2")
        ws["B6"] = 999
        ws["B12"] = "=+SUM(B2:B11)"  # the widened case, in the same sheet

    hits = at(_book(tmp_path, "guard_overwrite.xlsx", build), "overwritten_formula")
    assert [f.location for f in hits] == ["B6"]
    assert hits[0].level.value == "high"


def test_a_genuine_short_total_survives_every_widened_exemption(tmp_path):
    def build(wb):
        ws = wb.active
        ws.title = "Model"
        for row in range(2, 10):
            ws.cell(row, 1, row * 5)
        ws["A10"] = "=+SUM(A2:A7)-A99"  # widened prefix, a term naming a far cell

    hits = at(_book(tmp_path, "guard_total.xlsx", build), "total_misses_rows")
    assert [f.location for f in hits] == ["A10"]
    assert hits[0].level.value == "high"
    assert "2 rows" in hits[0].summary


# --- a header row is not a mistake ---


def test_a_header_row_computed_differently_is_not_high(tmp_path):
    """The top row of a percentage block divides by a different denominator.

    Found in a published Ontario census table: the block total's percentage is
    taken against the grand total, while every row beneath it is taken against
    that block total. Both are correct and the difference is the whole design, so
    ten cells like this arrived as HIGH on a file with nothing wrong in them.
    """

    def build(wb):
        ws = wb.active
        ws.title = "Population"
        ws["A4"] = 10_000_000  # the grand total, off to the left
        ws["B4"] = 7_600_000  # this block's total
        ws["C4"] = "=B4/A4"  # the block total as a share of everything
        for row in range(5, 12):
            ws.cell(row, 2, row * 1000)
            ws.cell(row, 3, f"=B{row}/B$4")  # each row as a share of the block

    hits = at(_book(tmp_path, "header_row.xlsx", build), "inconsistent_formula")
    assert [f.location for f in hits] == ["C4"]
    assert hits[0].level.value == "medium", "a cell bounding the block is not a slip"
    assert "end of the block" in hits[0].summary


def test_an_odd_formula_inside_the_run_is_still_high(tmp_path):
    """The other half: the exemption is about position, not about being unusual.

    Note the oddity has to differ in more than which row it anchors to. Shapes are
    compared with row numbers stripped, so ``=B8/B$4`` and ``=B8/B$99`` are the same
    shape and a wrong anchor row alone is invisible here. That limit is stated in
    the README rather than left for someone to discover.
    """

    def build(wb):
        ws = wb.active
        ws.title = "Population"
        ws["B4"] = 7_600_000
        for row in range(5, 12):
            ws.cell(row, 2, row * 1000)
            ws.cell(row, 3, f"=B{row}/B$4")
        ws["C8"] = "=B8/A$4"  # the wrong denominator column, mid-run

    hits = at(_book(tmp_path, "odd_inside.xlsx", build), "inconsistent_formula")
    assert [f.location for f in hits] == ["C8"]
    assert hits[0].level.value == "high"


# --- error reporting ---


def test_zip_without_a_workbook_gets_a_plain_message(tmp_path, capsys):
    import zipfile

    from show_your_work.cli import EXIT_UNREADABLE, main

    path = tmp_path / "empty.xlsx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("readme.txt", "not a workbook")
    assert main([str(path)]) == EXIT_UNREADABLE
    message = capsys.readouterr().err
    assert "does not contain a workbook" in message
    assert "KeyError" not in message
