# Show Your Work

Finds the numbers in a spreadsheet that nobody can explain.

Point it at an `.xlsx` file and it reports the things that make a total untrustworthy:
a formula someone typed over, a `SUM` that stops one row short, a cell that quietly
does something different from the column it sits in, a loop, a hidden sheet, a link to
a file that may have moved.

It runs entirely on your machine. No upload, no account, no API key, no network call.

```
$ show-your-work examples/messy-forecast.xlsx

Show Your Work — examples/messy-forecast.xlsx
========================================================================
Read 2 sheet(s). 7 finding(s): 3 high, 2 medium, 2 low.

HIGH
------------------------------------------------------------------------
HIGH    Forecast!D9   [inconsistent_formula]
        Formula differs from the 8 matching formulas in this column

HIGH    Forecast!D5   [overwritten_formula]
        Typed number inside a column of formulas

HIGH    Forecast!D12  [total_misses_rows]
        SUM skips 2 rows that sit between its range and the total
```

That output is real. `examples/messy-forecast.xlsx` is built by
`examples/make_examples.py` with seven problems planted in it, and
`examples/clean-forecast.xlsx` is the same sheet built properly, on which the tool
reports nothing.

```bash
python examples/make_examples.py
show-your-work examples/messy-forecast.xlsx   # 7 findings, exit 1
show-your-work examples/clean-forecast.xlsx   # nothing found, exit 0
```

## Why

A spreadsheet is the only document people quote from without checking how it was built.
The dangerous cell is never the one showing an error. It is the one showing a number
that looks exactly like the eleven around it and was typed in by hand three quarters ago.

This tool does not tell you a number is wrong. It tells you which numbers the sheet
cannot account for, so a human can go and look at those instead of all of them.

## Install

```bash
pip install show-your-work
```

Python 3.9 or newer. The only dependency is `openpyxl`.

From source:

```bash
git clone https://github.com/Waiga/show-your-work
cd show-your-work
pip install -e .
```

## Use

```bash
show-your-work model.xlsx                 # findings, no cell contents
show-your-work model.xlsx --verbose       # add why each finding matters
show-your-work model.xlsx --show-values   # include the actual cell contents
show-your-work model.xlsx --format json   # for scripts
```

Exit codes make it usable in a pipeline:

| Code | Meaning |
|---|---|
| `0` | nothing found at or above the fail level |
| `1` | findings at or above the fail level |
| `2` | the file could not be read |

`--fail-on` sets that level: `high` (default), `medium`, `low`, or `never`.

```yaml
# fail a pull request that introduces a hardcoded number into a model
- run: show-your-work models/pricing.xlsx --fail-on high
```

## What it checks

| Check | Level | What it means |
|---|---|---|
| `overwritten_formula` | high | A typed number sits inside a run of identical formulas. Someone replaced a calculation with a figure, so the sheet no longer explains it and it will not update. |
| `inconsistent_formula` | high | One formula differs from the many matching formulas around it. Invisible on screen, and the pattern most often found behind a wrong total. |
| `total_misses_rows` | high | A `SUM` range stops short of numbers sitting between it and the total. Rows added under a table fall outside a total nobody extended. |
| `circular_reference` | high | Cells depend on themselves, directly or through a chain. Excel shows zero rather than an error. |
| `iterative_calculation` | medium | Excel's iterative calculation setting is on. It is only needed when formulas depend on each other, and it makes results depend on how many passes Excel was told to run. |
| `error_value` | high | A saved result is `#REF!`, `#DIV/0!`, `#VALUE!` or similar. |
| `broken_defined_name` | high | A named range points at deleted cells. |
| `number_stored_as_text` | medium | Digits stored as text in a numeric column. `SUM` and lookups skip them silently. |
| `external_link` | medium | A formula reads from another workbook. The figure is whatever that file last saved. |
| `hidden_sheet` | medium / high | A sheet is hidden. "Very hidden" sheets cannot be revealed from the Excel menu at all. |
| `hidden_rows`, `hidden_columns` | medium / low | Content excluded from view but still inside totals. |
| `active_filter` | medium / low | A filter is applied, so the rows on screen may not be all the rows. |
| `volatile_function` | low | `TODAY`, `NOW`, `RAND`, `OFFSET`, `INDIRECT`. The file can show different numbers tomorrow with no edit in between. |
| `merged_cells` | low | Merged blocks spanning rows, which read as blanks inside ranges. |
| `macro_enabled` | medium | An `.xlsm` file. Code in it may change values on open. |
| `protected_sheet` | low | Recorded for context. It does not block any other check. |

## What it does not check

This list is part of the tool, not a disclaimer. Every run prints it.

- **Whether any number is correct.** A hardcoded figure may well be the right one. The
  tool reports that a number is unexplained, never that it is wrong.
- **Whether the formula logic matches the intent.** A model can be perfectly consistent
  and completely wrong about the business.
- **Anything outside the cells.** Charts, pivot caches, VBA code and embedded objects
  are not inspected.
- **Calculated results, when the file has none.** A workbook generated by a program and
  never opened in Excel stores formulas without results. Value-based checks cannot run,
  and the report says so rather than reporting a pass.
- **`.xls` files.** The old binary format is not supported. Re-save as `.xlsx`.

Absence of a finding is reported as "not checked", never as a confirmed no.

## Privacy

Cell contents are withheld from the report unless you pass `--show-values`. Findings
identify a cell by its address and describe the shape of the problem, so a report can be
pasted into a ticket or a chat without carrying the data with it.

Nothing is written anywhere except standard output. The workbook is opened read-only and
is never modified.

## How this compares

Spreadsheet auditing is an old field and this tool is not the first to look for these
things. Most of what checks a workbook today is either a paid Windows add-in
([PerfectXL](https://www.perfectxl.com/), [Operis Analysis Kit](https://www.operisanalysiskit.com/),
[Spreadsheet Detective](https://spreadsheetdetective.com/main/)), a feature of a
particular Excel licence tier
([Spreadsheet Inquire](https://support.microsoft.com/en-us/office/analyze-a-workbook-with-spreadsheet-inquire-5991e8fa-f1c1-401a-ae3f-469384ae3e3b)),
or a website you upload the file to.

The detection here is not novel. What is different is that it is free, open, runs as a
command, never sends the file anywhere, and returns an exit code you can put in CI.

Two of the checks come from published work rather than invention:
`inconsistent_formula` targets the error class studied by
[ExceLint](https://github.com/ExceLint/ExceLint) (OOPSLA 2018), and the error taxonomy
behind the severity levels follows
[Panko and Halverson](https://arxiv.org/pdf/0809.3613). `active_filter` is the one check
with no published evidence behind it; it is included on judgement, and rated low.

## Contributing

Useful tests, reproducible examples, honest limitations, and false-positive reports are
first-class contributions. A workbook that produces a wrong finding is the single most
valuable thing you can open an issue with, provided it contains no real data.

## Licence

MIT.
