# Corpus manifest

Two measured claims in the README rest on material outside this repository: the
Ontario workbook the "A real one" section is about, and the sweep of 527 public
government spreadsheets behind the 21% figure. This document says exactly what
each of them is, so a stranger can obtain the same files and check.

The README already pins the Ontario file properly — download URL, publication
date and SHA-256. That is the standard the rest of this document is written to.

Where something was not recorded at the time, this document says so rather than
reconstructing it.

## 1. The sweep: 527 public government spreadsheets

**What it is.** 527 valid `.xlsx` files downloaded from public open-data portals
on **7 September 2026**. This is the corpus that found four false-positive
classes in the tool, all fixed in v0.2, and it is the corpus the 21% figure is
computed over.

**The list.** [`docs/corpus/workbooks.tsv`](corpus/workbooks.tsv) — 527 rows
plus a header:

| column | meaning |
|---|---|
| `batch` | `1` or `2`, see the two collection rounds below |
| `local_name` | the filename the run used; the prefix is a collection index, the rest is the published filename |
| `bytes` | size of the downloaded file |
| `sha256` | SHA-256 of the downloaded bytes, unmodified |
| `formula_bearing` | see the caveat in section 2 |
| `tool_run` | whether the tool was actually run over this file |
| `source_url` | the download URL |

**Every one of the 527 rows carries a source URL.** The files were written to
disk exactly as served, so `sha256` is directly comparable against a fresh
download. These portals republish, so a changed hash means the file moved on,
not that the manifest is wrong.

**Where they came from.** Resources were enumerated through the CKAN-style APIs
of eight open-data portals. The download URLs those APIs return often point at
the publishing department's own host rather than at the portal, so the 527 files
arrive from **32 distinct hosts** across Australian, Canadian and Irish
government publishers. The `source_url` column is the authority for every row;
the largest hosts are:

| host | files |
|---|---|
| `data.gov.au` | 335 |
| `opendata.housing.gov.ie` | 51 |
| `opendata.transport.nsw.gov.au` | 21 |
| `data.nsw.gov.au` | 18 |
| `data.ontario.ca` | 13 |
| `www.budget.canada.ca` | 12 |
| 26 further hosts | 77 |

The Irish share is why the README's second false-positive class names Ireland's
local authority budgets: `opendata.housing.gov.ie` publishes seven years of them
and every edition tripped the same defect.

**Selection rule.** Resources advertising an `.xlsx` format were enumerated from
each portal's API, then downloaded. Anything that failed to download, or
downloaded but would not open as a valid `.xlsx`, was dropped. Nothing was
filtered on content, size, subject or whether the tool would find anything.

**The funnel, as recorded at the time:**

| stage | count |
|---|---|
| public `.xlsx` resources enumerated across the eight portals | ~2,430 |
| download attempts | ~1,060 |
| **valid `.xlsx` obtained — the 527** | **527** |
| of which contain any formula at all | 113 (21%) |
| tool run to completion on | 245 |

The two batches: **batch 1** is 159 files collected first, of which the tool was
run on 148; **batch 2** is 368 files collected second, of which the tool was run
on the 97 that carry formulas. 159 + 368 = 527, and 148 + 97 = 245.

**Why only 245 of 527 were run.** The remaining 271 formula-free workbooks in
batch 2 were abandoned on time. They can only produce structural and
value-based findings, which would push the "produced at least one finding" rate
down rather than up, so stopping was a decision that made the reported hit rate
*more* favourable than a complete sweep would have. That is worth knowing before
reading it.

**What the 245 produced**, recorded at the time and not recomputed here: 170 of
245 (69%) produced at least one finding of any level; 59 of 245 (24%) produced at
least one HIGH; and 28 of 245 (11%) produced at least one of the three headline
checks, `overwritten_formula`, `inconsistent_formula` or `total_misses_rows`.
None of those three figures appears in the README.

**Licence and redistribution.** Australian, Canadian and Irish government open
data, variously under Creative Commons Attribution and open government
licences; the
Ontario file in section 3 is under the Open Government Licence – Ontario.
**No workbook is redistributed here.** The manifest carries filenames, sizes,
hashes and URLs only, which sidesteps the licence question entirely and is
enough to obtain the same files.

## 2. ⚠️ The 21% figure, and what is and is not pinned about it

The README says: "In a sweep of 527 valid public government `.xlsx` files, only
21% contained a formula of any kind." That is 113 of 527, recorded at the time.

**What is pinned.** The 97 formula-bearing files of batch 2 were recorded by
name and are marked `formula_bearing = yes` in the manifest. The other 271 files
of batch 2 are marked `no`.

**What is not pinned.** The 16 formula-bearing files of batch 1 — 113 minus 97 —
**were never recorded individually**. Batch 1's rows therefore carry an empty
`formula_bearing` column rather than a guess. The column is blank because the
information does not exist, not because it was not looked for.

**And the method behind 113 is not recorded either.** No surviving script says
how "contains a formula" was decided. A crude re-count over the same 527 files —
scanning each worksheet's XML for an `<f>` element — finds 99, not 113, and its
result is a strict subset of the 97 recorded batch-2 names, so it is a narrower
test than whatever was used, not a contradiction of it. **The published 113 has
not been changed**, because a cruder recount is not grounds to overwrite a
recorded measurement. It is stated here so a reader who repeats the count and
gets a different answer knows why.

The bias the README already discloses stands and is the more important caveat:
the sample leaned toward budget and statistics files, so 21% is what those 527
workbooks showed and not a rate for spreadsheets in general.

## 3. The named file: Ontario rural and urban population counts

The one file the README makes a specific, checkable claim about. It is already
fully pinned there and is repeated here so the manifest is complete.

**It is also row `c042_…` of the 527.** The README's published SHA-256 and the
hash of that row are the same string, which is the one place in this manifest
where a published figure and the corpus check each other.

**Dataset.** Socio-economic statistics for rural and urban Ontario, published by
the Ontario Ministry of Agriculture, Food and Agribusiness.

```
https://data.ontario.ca/dataset/c30aa695-4735-466a-bc6e-fd31f1290973/resource/e07b6d92-31ef-437d-85f2-da88ef563515/download/population_statistics_-_rural_and_urban_ontario_population__counts.en.xlsx
```

**As published on** 7 September 2026.

**SHA-256** `ae3cc972acb90bfe40aedad5d233475e8db5abb9b690164a9e71d8c4fdd0aa19`

**Figures computed over it,** all in the README: the whole workbook produces 195
findings, of which 172 are error cells; 15 are on the `Population by age` sheet;
4 of those 15 are the ones that matter; 10 more are the block's header row,
reported at medium; the sheet is 246 rows.

**Licence.** Open Government Licence – Ontario. Not redistributed here.

**This claim rots by design.** Ontario may correct or republish the resource, in
which case the hash stops matching and the counts describe a file that no longer
exists. The hash is how a reader tells. As the README says, that would be good
news.

## 4. What a third party can and cannot reproduce

| claim | status |
|---|---|
| which 527 workbooks | **reproducible** — every filename, hash and URL ships here |
| the 245 the tool was run on | **reproducible** — flagged per row |
| the Ontario file and its findings | **reproducible** — URL, date and hash pinned |
| the 97 formula-bearing files of batch 2 | **reproducible** — flagged per row |
| the 16 formula-bearing files of batch 1 | **not reproducible** — never recorded individually |
| the exact method behind "113 contain a formula" | **not recorded** — a stricter recount gives 99 |
| the per-file findings from the sweep | working files on one machine, **not published** |
| the sweep harness | **not published** |
