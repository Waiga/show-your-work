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

## A real one

The example above is built to be broken. This one was not.

[Socio-economic statistics for rural and urban Ontario](https://data.ontario.ca/dataset/c30aa695-4735-466a-bc6e-fd31f1290973),
published by the Ontario Ministry of Agriculture, Food and Agribusiness under the Open
Government Licence – Ontario.

On the `Population by age` sheet, the 2016 census block occupies columns R to X, and rows
5 to 22 are eighteen five-year age bands. Two of them are gone. The age labels in R19 and
R20 are empty, the Ontario counts in S19 and S20 are empty, and T19 and T20 — which should
divide the Ontario count by the block total — have a `#REF!` numerator and a saved `#REF!`
result.

Which two bands were lost is legible from the sheet itself. The 2021 block immediately to
the left still labels its rows 19 and 20 "Both sexes: 70-74 years" and "Both sexes: 75-79
years", and those are exactly the bands the 2016 sequence skips: it runs 65-69 and then
jumps to 80-84. The Urban and Rural columns on those same two rows still hold their
numbers, and on all sixteen surviving 2016 rows the Ontario figure equals Urban plus Rural
exactly — so the identity that holds everywhere else in the block would reconstruct both
missing figures from cells that are still there.

```bash
curl -s -o population_counts.xlsx \
  'https://data.ontario.ca/dataset/c30aa695-4735-466a-bc6e-fd31f1290973/resource/e07b6d92-31ef-437d-85f2-da88ef563515/download/population_statistics_-_rural_and_urban_ontario_population__counts.en.xlsx'

show-your-work population_counts.xlsx | grep 'Population by age!T'
```

The counts below describe the file as published on 7 September 2026, whose SHA-256 is
`ae3cc972acb90bfe40aedad5d233475e8db5abb9b690164a9e71d8c4fdd0aa19`. Ontario may correct or
republish it, in which case the numbers here are of a file that no longer exists and the
hash is how you can tell. That would be good news.

```
HIGH    Population by age!T19  [error_value]
HIGH    Population by age!T20  [error_value]
HIGH    Population by age!T19  [inconsistent_formula]
HIGH    Population by age!T20  [inconsistent_formula]
```

**The `grep` is not decoration, and leaving it out would misrepresent the run.** The whole
workbook produces 195 findings. Fifteen are on this sheet; the other 180 are on two sheets
that carry many `#REF!` and `#N/A` cells of their own, and 172 of the 195 are those error
cells. On a file this size the tool hands you a list to search, not an answer. That is the
honest shape of the result.

Of the fifteen on this sheet, the four above are the ones that matter. Ten more are the
block's header row, reported at medium because a cell bounding a block is where a total
belongs; the last is a typed-in cell that the sheet's own layout explains. Those ten were
high until this file was run against the tool, and being wrong about a real published
workbook is what got the level fixed.

Note also which check caught it. `error_value` is the least clever row in the table below,
and this README's own argument is that the dangerous cell is the one *not* showing an
error. Both are true. The error is only the marker; the finding is the two missing rows —
three columns wide, two thirds of the way down a 246-row sheet, in a block that reads as
complete because every band still present is correct.

## Why

A spreadsheet is the only document people quote from without checking how it was built.
The dangerous cell is never the one showing an error. It is the one showing a number
that looks exactly like the eleven around it and was typed in by hand three quarters ago.

This tool does not tell you a number is wrong. It tells you which numbers the sheet
cannot account for, so a human can go and look at those instead of all of them.

## Install

Python 3.9 or newer. The only dependency is `openpyxl`.

```bash
pip install unexplained-cells
```

Or from source:

```bash
git clone https://github.com/Waiga/show-your-work
cd show-your-work
pip install -e .
```

The package is called `unexplained-cells`, which is what it reports. The name
`show-your-work` on PyPI belongs to [showyourwork](https://pypi.org/project/showyourwork/),
an established and unrelated project for reproducible scientific articles, and a
near-miss name beside theirs would help nobody. The command you type is still
`show-your-work`.

## Use

```bash
show-your-work model.xlsx                 # findings, no cell contents
show-your-work model.xlsx --verbose       # add why each finding matters
show-your-work model.xlsx --show-values   # include the actual cell contents
show-your-work model.xlsx --format json   # for scripts
show-your-work a.xlsx b.xlsx c.xlsx       # several files in one run
```

More than one file may be given. Each is opened and reported on its own, one
unreadable file does not stop the rest, and the exit code is the worst any single
file earned. In `--format json`, one file prints a single object and several print an
array, so the whole of standard output stays valid JSON either way.

Exit codes make it usable in a pipeline:

| Code | Meaning |
|---|---|
| `0` | nothing found at or above the fail level |
| `1` | findings at or above the fail level |
| `2` | a file could not be read, or refused before reading, or the command line was wrong |

`--fail-on` sets that level: `high` (default), `medium`, `low`, or `never`.

A file that is not a valid `.xlsx`, or that is shaped like a decompression bomb or an
XML entity-expansion payload, is refused with a plain message and exit `2`. It is never
a silent empty report, because "nothing was examined" must not read like "nothing was
found".

```yaml
# fail a pull request that introduces a hardcoded number into a model
- run: show-your-work models/pricing.xlsx --fail-on high
```

### As a pre-commit hook

This repository ships a [pre-commit](https://pre-commit.com) hook, so every `.xlsx` or
`.xlsm` being committed is audited before it lands. In a consuming repository's
`.pre-commit-config.yaml`:

```yaml
repos:
  - repo: https://github.com/Waiga/show-your-work
    rev: v0.2.0            # pin to a tag or commit
    hooks:
      - id: show-your-work
        args: [--fail-on, high]   # optional; high is the default
```

pre-commit passes every staged spreadsheet to one run of the tool, and a commit fails
when any of them has a finding at or above the fail level.

### In GitHub Actions

The same command runs on the spreadsheets in a repository. This reusable job audits
every `.xlsx` under the checkout and fails the build on any high finding:

```yaml
# .github/workflows/spreadsheets.yml
name: Audit spreadsheets

on:
  pull_request:
  push:
    branches: [main]

permissions:
  contents: read

jobs:
  show-your-work:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.13"
      - run: pip install git+https://github.com/Waiga/show-your-work
      - name: Audit every spreadsheet in the repo
        run: |
          shopt -s globstar nullglob
          files=(**/*.xlsx **/*.xlsm)
          if [ ${#files[@]} -eq 0 ]; then
            echo "No spreadsheets to audit."
            exit 0
          fi
          show-your-work "${files[@]}" --fail-on high
```

The tool never uploads the file or makes a network call, so it needs no secrets and the
`contents: read` permission above is all it uses.

## What it checks

| Check | Level | What it means |
|---|---|---|
| `overwritten_formula` | high / medium | A typed number sits inside a run of identical formulas. Someone replaced a calculation with a figure, so the sheet no longer explains it and it will not update. **High** when formulas sit on both sides of it. **Medium** at the top or bottom of a run, where a typed number is more often deliberate. A starting value that the formulas below it refer back to, such as an opening balance, is not reported at all. |
| `inconsistent_formula` | high / medium | One formula differs from the many matching formulas around it. Invisible on screen, and the pattern most often found behind a wrong total. Medium when the odd cell is the first or last of the block, where a header or a total belongs. A total closing the block is exempt however it is written: `=SUM(...)`, the Lotus-style `=+SUM(...)` and `=ROUND(SUM(...),0)` are all read as totals. |
| `total_misses_rows` | high / medium | A `SUM` range stops short of the numbers next to it, checked both down a column and across a row. Rows added to a table fall outside a total nobody extended. A subtotal in the gap is ignored, because stacked sections are meant to be built that way. So is a row the formula itself names, as in `=SUM(D93:D103)-D104`, where row 104 is a deduction added outside the range on purpose. So is a partial aggregate repeated in three or more cells, which is a deliberate subset rather than a slip. |
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

**Most spreadsheets are not models, and three of these checks need one.** In a sweep of 527
valid public government `.xlsx` files, only 21% contained a formula of any kind. On the
other 79% — flat, value-only exports — `overwritten_formula`, `inconsistent_formula` and
`total_misses_rows` are structurally incapable of firing: there is no formula to overwrite,
no repeated pattern to break, and no range to fall short of. Those files can still produce
`error_value`, `number_stored_as_text`, `hidden_sheet` and the rest, but the three checks
this tool exists for have nothing to read, and a clean report on such a file means only
that the value-based checks found nothing. That sample leaned toward budget and statistics
files, so 21% is what those 527 workbooks showed, not a rate for spreadsheets in general.

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

## The boundary — what the tool structurally cannot see

The section above is what the tool chooses not to judge. This section is different: it is
what the tool *cannot* reach, because the file format hides it from the way the tool
reads a workbook. It reads cells: a formula's text and, when Excel saved them, the last
results. Anything that is not a cell, or that lives only in a part the reader does not
open, is outside its reach.

- **Macros.** The tool detects that an `.xlsm` carries a VBA project and flags it, so you
  know code is present. It does not read, run, or reason about that code. A macro can
  rewrite any value on open, and the tool cannot tell you whether it does.
- **External workbook links.** A formula that reads from another file is flagged, but the
  other file is never opened. The number you see is whatever that file last saved into
  this one; if it moved or changed, the tool cannot resolve the link to find out.
- **Array and dynamic-array formulas.** A legacy array (Ctrl+Shift+Enter) formula reaches
  the reader as an object rather than plain text, so the consistency, total and
  overwrite checks skip it, and the cells it spills into read as blank. A modern
  dynamic-array formula such as `SORT`, `FILTER`, `UNIQUE` or `SEQUENCE` is stored only
  on its anchor cell; Excel produces the spilled cells at open time, so the tool sees
  neither the spilled values nor how far they reach. Neither kind is analysed.
- **Cached results versus live computation.** The tool reads the results Excel last saved,
  never a value it computed itself. A workbook written by a program and never opened in
  Excel has formulas with no saved results; the value-based checks cannot run and the
  report says so. When saved results do exist, they are only as current as the last save,
  and a formula changed since then may show a stale number the tool takes at face value.
- **A one-off partial aggregate.** A total that deliberately sums part of a range —
  `=SUM(B2:D2)` in a table that runs out to F — is indistinguishable from a total that
  stops short by mistake. The file records the range, never the intent. The tool settles
  it by repetition: the same partial-aggregate formula appearing in three or more cells is
  a design decision and is not reported, on the reasoning that a slip happens once. A
  genuinely one-off deliberate partial aggregate is therefore still reported, and nothing
  in the file could tell it apart. The repetition test compares formulas with their row
  numbers stripped, so a partial aggregate filled *down* a column collapses to one shape
  and is recognised; the same one filled *across* a row does not, and is still reported.
- **A formula that differs only in which row it anchors to.** Formulas are compared with
  their row numbers stripped, which is what lets a filled-down column collapse to a single
  shape. The cost is that `=B8/B$4` and `=B8/B$99` are the same shape, so a total pointed
  at the wrong anchor row is not an inconsistency this check can see. It is exactly the
  kind of mistake the tool exists to catch, and it is the one shape it is blind to.
- **Pivot caches.** A pivot table keeps its own snapshot of the source data. The tool does
  not open that cache, so it cannot tell whether a pivot is stale or what it summarises.
- **Charts.** Charts, and the series and cached points inside them, are not read. A chart
  can plot numbers that no longer match the cells it was built from, and the tool will
  not see the difference.

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
