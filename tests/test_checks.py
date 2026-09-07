"""Each check fires on the thing it looks for, and stays quiet otherwise."""

from __future__ import annotations

from show_your_work.cli import run


def checks_in(path: str) -> set[str]:
    return {f.check for f in run(path).findings}


def find(path: str, check: str) -> list:
    return [f for f in run(path).findings if f.check == check]


# --- false positives are the expensive failure, so they come first ---


def test_clean_model_reports_nothing(clean_model):
    report = run(clean_model)
    assert report.findings == [], (
        "A well-built sheet must produce no findings. Got: "
        + ", ".join(f"{f.check}@{f.location}" for f in report.findings)
    )


def test_total_row_is_not_an_inconsistent_formula(clean_model):
    assert "inconsistent_formula" not in checks_in(clean_model)


def test_total_row_does_not_look_like_an_overwritten_formula(clean_model):
    assert "overwritten_formula" not in checks_in(clean_model)


def test_function_name_is_not_read_as_a_cell_reference(tmp_path):
    from openpyxl import Workbook

    wb = Workbook()
    wb.active["A1"] = "=LOG10(100)"
    path = tmp_path / "log.xlsx"
    wb.save(path)
    assert "circular_reference" not in checks_in(str(path))


# --- positive cases ---


def test_typed_number_inside_a_formula_column(overwritten):
    hits = find(overwritten, "overwritten_formula")
    assert [f.location for f in hits] == ["C6"]
    assert hits[0].level.value == "high"


def test_total_that_stops_short(short_total):
    hits = find(short_total, "total_misses_rows")
    assert len(hits) == 1
    assert hits[0].location == "A10"
    assert "2 rows" in hits[0].summary


def test_odd_formula_in_a_filled_column(odd_formula):
    hits = find(odd_formula, "inconsistent_formula")
    assert [f.location for f in hits] == ["B7"]


def test_circular_chain(looping):
    hits = find(looping, "circular_reference")
    assert len(hits) == 1
    assert "loop" in hits[0].summary


def test_hidden_sheets_rows_and_columns(hidden_things):
    found = checks_in(hidden_things)
    assert {"hidden_sheet", "hidden_rows", "hidden_columns"} <= found
    very = [f for f in find(hidden_things, "hidden_sheet") if f.sheet == "Buried"]
    assert very[0].level.value == "high"


def test_digits_stored_as_text(text_numbers):
    hits = find(text_numbers, "number_stored_as_text")
    assert len(hits) == 1
    assert "A5" in hits[0].location


def test_link_to_another_workbook(linked):
    assert find(linked, "external_link")


def test_missing_cached_values_is_reported_as_unchecked(clean_model):
    topics = " ".join(u.topic for u in run(clean_model).unchecked)
    assert "Formula results" in topics
