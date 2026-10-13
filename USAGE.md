# Excel sales transformation

Local, standard-library-only Python 3.11+ program. Reads the source without
modifying it and does not transmit transaction contents externally.

## Run

```powershell
python excel_sales.py --source source/TransactionDetail.xlsx --target target
```

An isolated Python 3.13.7 runtime was downloaded into ignored `.venv/` for this
workspace. No third-party packages were installed:

```powershell
.venv/python.exe excel_sales.py --source source/TransactionDetail.xlsx --target target
```

The JSON report returns status, output path, counts, issues and complete/blocked
validation. Exit status is 0 for published output, 1 for failure. Formula caches
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

No production workbook or interactive Excel opening/recalculation has been tested.
Test results verify the stored workbook rather than future Excel behavior.
The existing README and requirement draft are preserved.

Latest executed validation: 21 unittest methods passed, including table-driven
hierarchy and formula subcases; no skipped tests. CLI synthetic transformation
completed with 3 channels, 3 dates, 2 months and zero critical errors. Tests inspect
stored workbook values; interactive Excel and production files were not used.
