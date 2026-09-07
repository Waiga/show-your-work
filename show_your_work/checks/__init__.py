"""The checks, and the order they run in.

Each check is a callable taking (LoadedWorkbook, Report) and appending findings.
A check that cannot run records why, rather than staying silent and reading as a pass.
"""

from __future__ import annotations

from show_your_work.checks import consistency, formulas, references, structure

ALL_CHECKS = (
    formulas.overwritten_formulas,
    consistency.inconsistent_formulas,
    formulas.totals_that_miss_rows,
    references.circular_references,
    formulas.cached_error_values,
    formulas.broken_defined_names,
    consistency.numbers_stored_as_text,
    formulas.external_workbook_links,
    structure.hidden_sheets,
    structure.hidden_rows_and_columns,
    structure.active_filters,
    formulas.volatile_functions,
    structure.merged_cells_in_data,
    structure.macro_enabled_file,
    structure.protected_sheets,
)

__all__ = ["ALL_CHECKS", "consistency", "formulas", "references", "structure"]
