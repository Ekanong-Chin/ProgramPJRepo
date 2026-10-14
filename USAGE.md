# Excel sales transformation

Local, standard-library-only Python 3.11+ program. Reads the source without
modifying it and does not transmit transaction contents externally.

## Run

```powershell
python excel_sales.py --source source --target target
```

An isolated Python 3.13.7 runtime was downloaded into ignored `.venv/` for this
workspace. No third-party packages were installed:

```powershell
.venv/python.exe excel_sales.py --source source --target target
```

Directory mode reads every direct-child `.xlsx` file (including uppercase
extensions), ignoring Excel `~$` lock files and subdirectories. A single XLSX path
is also supported. Source and target directories must differ.

Each source has its own output history: `aaa.xlsx` produces
`aaa_YYYYMMDD_HHMMSS.xlsx`, while `bbb.xlsx` produces
`bbb_YYYYMMDD_HHMMSS.xlsx`. Collisions append `_001`, `_002`, etc.

The batch JSON report returns per-file status, output, counts and issues, plus
success/failure totals. One invalid file does not stop the others. Exit status
is 0 only when all files succeeded, 1 for any failure. Formula caches
are the last saved values; freshness is not verified.

## Approved decisions

Implementation and the following remaining decisions were explicitly approved in
chat after Revision 8. This supersedes historical implementation gates in the draft.

- Scan raw sentinel candidates from row 3 through the last populated/formula row;
  require exactly one A=`Total`, B-D blank row. Footer sales are ignored.
- Monthly A is first-of-month date displayed `YYYY-MM`, B is actual date count,
  C-E blank, F+ channels. Daily layout follows the specification.
- Read numeric XML tokens directly. Round each daily amount to two decimal places
  using `ROUND_HALF_UP`, as subsequently approved by the user. Aggregate the
  rounded daily values with exact integer coefficients/exponents. Source tokens
  with more than 15 significant digits may be accepted when their rounded values
  pass the output policy. Changed daily cells are counted in rounding warnings.
  Strip insignificant trailing zeros for significant-digit counting. Daily and
  monthly values must have <=15 significant digits and nonzero magnitude within
  `2.2250738585072E-308` to `1.79769313486231E+308`, inclusive.
- Write native numeric tokens, then verify strict Decimal equality for every
  serialized value and exact expected cell coordinates. This guarantee concerns
  stored tokens, not later Excel recalculation/resaving.
- Accept ordinary/shared/array sales formulas with numeric caches. Reject missing,
  text, boolean or error caches, unsupported cell metadata, header/hierarchy formulas,
  and date/sales merges. Do not evaluate formulas.
- Limits: input 50 MiB; expanded ZIP 250 MiB; 100,000 source rows; 366 dates;
  1,000 channels; 1,000,000 intersections.
- Merges wholly within header rows 1-2 and columns A-D are ignored. Merges crossing
  into channel data, the sentinel, or date/sales columns fail. Merge conflict is diagnosed at the
  incompatible explicit change; earlier rows are not retroactively invalidated.
- Output is saved and validated as a temporary file, then atomically linked to a
  new filename without overwriting. A filesystem without safe hard-link publication
  fails explicitly rather than using an unsafe fallback.

## Tests and example

```powershell
.venv/python.exe -c "import sys; sys.path.insert(0, '.'); import unittest; suite=unittest.defaultTestLoader.discover('tests'); result=unittest.TextTestRunner(verbosity=2).run(suite); sys.exit(not result.wasSuccessful())"
.venv/python.exe examples/create_sample.py
```

Tests use synthetic fixtures and temporary directories within the repository.
They cover hierarchy/reset/recovery, exclusions, formula-cache package tokens,
precision, duplicates, sparse/month/year dates, failure preservation, styles,
serialization tampering and concurrent output publication.

The example creates source and output under ignored `target/synthetic/` and refuses
to replace the existing synthetic source. Rounded October totals are HATO=`0.30`,
Scheduled=`0.00`, Gofresh=`1.00`; November totals are `0.00`, `2.13`, `-1.00`.

Interactive Excel opening/recalculation has not been tested.
Test results verify the stored workbook rather than future Excel behavior.
The existing README is preserved; Steps 21-22 describe upsert and per-source batches.

Latest executed validation: 32 unittest methods passed, including table-driven
hierarchy and formula subcases; no skipped tests. CLI synthetic transformation
completed with 3 channels, 3 dates, 2 months and zero critical errors. Tests inspect
stored workbook values; interactive Excel and production files were not used.

## Insert or replace by channel and date

For each source, the default run reads only its latest timestamped workbook in
`--target` (exact source stem, timestamp then numeric suffix), and replaces
matching `(channel.casefold(), date)` daily amounts, inserts new keys and preserves
old history. Monthly totals/Days are rebuilt from the complete merged matrix.
Channels are grouped by their resolved A/B/C parents, retaining first-seen sibling
order. New channels join their group's block; missing intersections are zero-filled. The
result is a new version, preserving all prior workbooks. The JSON `upsert` field
reports baseline, updated/inserted key counts and combined dimensions.

Select a particular baseline explicitly:

```powershell
.venv/python.exe excel_sales.py --source source/TransactionDetail.xlsx --target target --master target/MasterDetail_20261014_005446.xlsx
```

This also allows intentional migration of an old `MasterDetail_*.xlsx` history.
Old global MasterDetail files are never automatically assigned to a source.
`--master` is only supported with a single XLSX source, not directory mode.

An explicit missing/invalid master fails safely. A target-directory lock serializes
automatic read/merge/publish operations, with a 30-second timeout. If a process is
terminated and leaves `.sales-upsert.lock`, remove it only after verifying that
no transformation is active. Duplicate channels/dates inside a single source
are still rejected; source-versus-master duplicate keys are replaced, not summed.

## Channel group ordering

Both Monthly Total and Daily Sales put channels of the same hierarchy group
together (A, then B, then C), without alphabetic renaming or numeric aggregation.
Verified parent paths are stored in a worksheet extension for future partial
upserts. Masters created before this feature acquire grouping from incoming
matching channels; unknown historical channels retain stable ungrouped ordering.
If an external editor removes the extension, the same fallback applies.
Group names compare exactly; upsert channel keys still use casefold.
