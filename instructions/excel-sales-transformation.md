# Excel Sales Transformation — Program Instructions

**Version:** 1.4 — Part 1 Revision 8 (Hierarchy & Parent Total; Requirements Draft for Codex Review)  
**Status:** Hierarchy Rules Confirmed; Other Implementation Blockers Remain; Implementation Not Approved  
**Project:** ProgramPJRepo  
**Scope:** Requirements, validation rules, and tests only. Do not implement source code until the user explicitly approves.

## Step 1 — Program Overview & Objectives

Develop a local Excel transformation program that reads `TransactionDetail.xlsx`, extracts sales channels from hierarchical columns A–D and daily amounts from dated columns E onward, pivots daily sales into channel columns, and generates a new `MasterDetail_YYYYMMDD_HHMMSS.xlsx` workbook with Monthly Total and Daily Sales sections.

- Dynamically detect channels from source hierarchy without referring to cell colors or formatting.
- Preserve exact original selected channel names and their source order.
- Generate the output workbook without any existing template.
- Never modify the source, overwrite prior output, transmit transaction contents externally, or perform Git write operations without explicit user approval.
- Follow project `AGENTS.md`.

## Step 2 — Input & Output File Specifications

### Directories and input

- Default source: `./source/TransactionDetail.xlsx`; configurable source and target directories.
- Input format: `.xlsx`; read **the first worksheet**, with **date headers in row 1**, **sales subheaders in row 2**, and **channel data beginning in row 3**. See Step 12 for the confirmed end-of-data marker.
- Columns **A–D**: hierarchical channel labels; columns **E onward**: daily values aligned to date headers, except explicitly recognized non-date summary columns (see Step 10).
- Recognize actual Excel date values and explicitly supported text dates `DD/MM/YYYY` and `YYYY-MM-DD`; normalize internally to calendar dates. Interpret slash dates as day-first. Reject unsupported/ambiguous dates; detect duplicate calendar dates.
- Read-only input; do not modify, move, rename, or delete.

### Output

- Default target: `./target/`.
- Each successful run creates **one new** `.xlsx` workbook named `MasterDetail_YYYYMMDD_HHMMSS.xlsx` (append `_001`, `_002`, etc. on name collisions, without overwriting).
- Exactly one worksheet: `MasterDetail`, containing Monthly Total above Daily Sales, separated by a blank row.
- Daily Sales columns: A `Month / Date`; B `Day`; C `Weekday`; D `Tab WK` (blank); E `Tab Promotion` (blank); F onward: identified channel names.
- Monthly Total columns include `Month`, `Days`, and monthly per-channel totals; keep the same channel ordering in both sections. Exact physical positions for monthly channel headers should be settled during final layout review if needed.
- Source date columns only: only explicitly recognized summary headers from Step 10 are excluded from daily dates; unknown non-date headers are critical errors.

## Step 3 — Channel Detection Rules (superseded in part by Step 13, v1.4)

**Read Step 13 before implementing these rules.** Step 13 supplies the authoritative hierarchy carry-forward/reset, parent/child, and multi-candidate classification rules; if Step 3 conflicts with Step 13, Step 13 prevails.

Determine channel candidates from the hierarchy in columns A–D without relying on colors, highlights, borders, font styles, or other formatting.

1. **Rule 1 — Followed by Total:** when a nonempty label in column A, B, or C is immediately followed by `Total` in the next column, use the preceding label as the candidate channel name. Recognize `Total` case-insensitively after trimming surrounding spaces.
2. **Rule 2 — Column D channel:** when D has a nonempty value other than `Total`, use D as candidate channel.
3. Resolve merged/blank hierarchical parent cells where possible, using the surrounding hierarchy; retain source row numbers and hierarchy paths for diagnostics.
4. **Exclude parent Total rows** whenever the parent has valid child channels; exclude grand totals. Keep terminal channel-level Total rows if they have no selected child channels.
5. Do not use `Total` as a channel name. Preserve the *exact* source spelling, capitalization and internal spaces of selected channel labels; no mapping file or renaming. Preserve first-appearance order.
6. Validate duplicate selected channel names after parent-total exclusion using the **case-insensitive `casefold()` comparison** specified in Step 11.2. Duplicate names are **critical errors**, not candidates for aggregation.
7. If parent-child relationships or selections are ambiguous, report row details and fail rather than guessing.

Examples: B=`HATO`, C=`Total` → HATO; C=`OMS Scheduled CC`, D=`Total` → OMS Scheduled CC; D=`OMS OnDemand Gofresh` → OMS OnDemand Gofresh; B=`OMS`, C=`Total` → exclude when OMS has valid child channels.

## Step 4 — Daily Sales Extraction (Final)

- Read dates from first worksheet row 1, columns E onward; exclude only recognized summary headers per Step 10; reject unknown non-date headers, invalid dates, and duplicate dates.
- Extract sales **only** for selected channels (Step 3), never excluded parent subtotals/grand totals.
- Match each amount to its source date, channel, row and column.
- Blank/null/empty sales cells become **numeric 0**.
- Accept **negative**, positive, and zero numeric amounts; preserve underlying precision.
- Reject unexpected nonnumeric text, Excel error cells, and unsupported data types; report cell address without logging confidential sales values.
- Validate all accessible selected channel/date combinations before allowing output generation.

## Step 5 — Data Transformation: Rows to Columns (Final)

- Pivot source rows into a target matrix with one date row and one column per selected channel.
- Daily channel columns start at **F**; labels are exact source names and remain in first-appearance order.
- Sort distinct source dates ascending; write as native Excel dates. Day is the day-of-month integer, Weekday is the English weekday name, and Tab WK / Tab Promotion remain blank.
- Map each source sales amount to its exact date/channel intersection. Convert blanks to 0, retain negatives, never include excluded subtotals or grand totals.
- **Stop on duplicate selected channel:** if two or more selected source rows yield the same channel name under the **case-insensitive Step 11.2 duplicate rule**, collect all affected row numbers as critical validation error(s), do not sum/skip/overwrite, and do not create final MasterDetail.
- Reject duplicate source calendar dates likewise; no hidden aggregation.

## Step 6 — Monthly Total & Daily Sales (Final)

- In the same worksheet, place Monthly Total first and Daily Sales below it.
- **Only actual source dates** are included in both sections; never generate absent calendar dates.
- Daily Sales: one row per distinct source date in chronological order.
- Monthly Total: one row per distinct **year-month** in chronological order. Sum each channel's actual daily records, including negative and zero values.
- `Days` is the number of unique **actual source dates** for the month, not the calendar's total days.
- Ensure per-channel monthly totals exactly reconcile with daily rows, using unrounded values for calculation.

## Step 7 — MasterDetail File Generation (Final)

- New workbook and single `MasterDetail` worksheet for each successful run. No existing template required.
- Use section titles `MONTHLY TOTAL` and `DAILY SALES`, with one blank row between sections.
- Corporate style: dark blue (`#17365D`) header backgrounds; bold white (`#FFFFFF`) header fonts; readable borders, alignments, and column widths.
- Sales cells are **numeric**, formatted with `#,##0.00` (display only: do not round underlying values); see Step 10 for decimal calculation and Excel storage limitations. Dates use `DD/MM/YYYY` Excel display format.
- For Daily Sales: A date; B day; C weekday; D and E blank; F onward exact source channel headers.
- Validate all source records and reconciliations first. A temporary XLSX file and post-save validation are **mandatory**; publish the final output only after the temporary file has passed validation. Never overwrite existing target, including during collisions; remove only current-run temporary files on failure.
- Report final filename and path, selected channel count, daily date count, month count, processing status and warnings.

## Step 8 — Validation & Error Handling (Final)

**Collect all detectable, independent critical errors.** Continue validation where safe; if upstream prerequisites are missing, report that dependent validation was blocked rather than falsely claiming full validation. If any critical error exists, **do not create final MasterDetail**.

Checks include:

- Source presence, `.xlsx` validity, readability, first worksheet availability, required columns and target location.
- Date headers, supported Excel/text dates, summary-column identification, duplicate dates and absence of valid date columns.
- Channel selection per Step 3, parent/child relationships, no selected channels, duplicate selected names and ambiguous hierarchy.
- Sales amounts for selected channels: blank→0, negatives allowed; malformed/non-numeric/error cells are critical.
- Mapping and monthly/daily reconciliation.
- Safe output creation and no overwritten prior file.

For each issue capture: error code, severity, category, worksheet, source row/column (where applicable), concise explanation, and recommended action. Do **not** log confidential transaction amounts or credentials. Provide counts of critical errors/warnings, channels/dates detected where available, and whether validation was complete or blocked. Warnings can cover accepted blank→0 and missing calendar dates; they do not block output. Never modify original source or existing output.

## Step 9 — Unit Tests, End-to-End Tests & Acceptance Criteria (Final)

### 9.1 General test policy

1. Create tests for pure parsing, channel classification, pivoting, date normalization, aggregation, validation, formatting and file generation.
2. Use **synthetic fixture data only**; do not copy production/customer/financial records into fixtures, logs, or test reports.
3. Tests must run locally, without uploading data or invoking external services. Avoid external dependencies unless explicitly approved.
4. Use isolated temporary source and target directories; clean only files the test created. Never touch existing user outputs.
5. Every negative test must verify that the run generates **no final MasterDetail** and does not modify existing output or source.
6. Verify workbook content by reading the generated `.xlsx`, not merely by inspecting console output.
7. Do not claim tests passed unless actually executed; record pass/fail and causes.

### 9.2 Unit test scenarios

| ID | Scenario | Expected result |
|---|---|---|
| UT-01 | B=`HATO`, C=`Total` | Select exact channel `HATO` |
| UT-02 | C=`OMS Scheduled CC`, D=`Total` | Select `OMS Scheduled CC` |
| UT-03 | D=`OMS OnDemand Gofresh` | Select exact D label |
| UT-04 | Parent `OMS` Total with valid child channels | Exclude parent subtotal; include qualifying children |
| UT-05 | Overall grand total | Exclude from selected channels |
| UT-06 | ` total ` in next hierarchy column | Detect Total after trim/case-fold; preserve selected channel label |
| UT-07 | Colors changed, hierarchy unchanged | Identical detected channels |
| UT-08 | Unresolvable hierarchy | Critical error with source location |
| UT-09 | Two selected rows named `Lazada` | Critical duplicate error listing both rows |
| UT-10 | Excel date cell | Normalize to expected calendar date |
| UT-11 | Text `13/10/2026`, `2026-10-13` | Parse correctly as the same date |
| UT-12 | Invalid or repeated date headers | Critical errors, identify headers/cells |
| UT-13 | Each approved summary header (`Total Sales`, `Grand Total`, `Total`, `ยอดขายรวม`) | Exclude from daily dates; test trim and case-insensitive English matching |
| UT-14 | Blank selected sales cell | Use numeric 0 |
| UT-15 | Negative selected sales amount | Preserve negative value |
| UT-16 | Text/error sales cell | Critical error with cell reference |
| UT-17 | Sparse source dates (e.g. Oct 1, 2, 5) | Only three Daily rows; monthly Days=3 |
| UT-18 | Month boundary / same month in different years | Distinct year-month aggregates |
| UT-19 | Pivot mapping, nonalphabetical channel order | Correct values and first-seen channel order |
| UT-20 | Monthly reconciliation incl. negatives/decimals (e.g., `0.1` + `0.2`, and >2 decimals) | Decimal arithmetic, no intermediate rounding; exact internal reconciliation |
| UT-21 | Filename collision | Unique suffix; prior output byte-for-byte unchanged |
| UT-22 | Collect multiple independent errors | Aggregate report; no final workbook |
| UT-23 | Unknown non-date header, e.g. `Sales Amount` | Critical error with header location; do not silently ignore |
| UT-24 | Summary header ` TOTAL SALES ` and unapproved partial match `Total Sales Adjusted` | Skip the former only; flag latter critical |
| UT-25 | Excel number with source representation limitations | Preserve precision available from workbook; report unsupported precision instead of fabricating digits |

### 9.3 End-to-end scenarios

| ID | Scenario | Expected result |
|---|---|---|
| E2E-01 | Valid synthetic source across multiple months, mixed Channel Rule 1/Rule 2 | One new correctly formatted MasterDetail, correct Monthly and Daily sections |
| E2E-02 | Synthetic source with blank, zero and negative sales | 0 for blanks, negatives retained, monthly totals reconcile |
| E2E-03 | Source with excluded parent subtotal and grand total | No double counting; excluded rows not output as channels |
| E2E-04 | Duplicate selected channel + invalid sales + duplicate date | All safely detectable errors reported; no output |
| E2E-05 | Run twice with same input | Two distinct files; first file unchanged; equivalent contents except output filename metadata |
| E2E-06 | File absent, invalid workbook, unwritable output | Clear errors; no partial final file; existing files preserved |
| E2E-07 | Dates only on 1, 2, and 5 of month | Exactly 3 Daily rows; Monthly Days=3; no rows for absent dates |
| E2E-08 | Input formatting/colors changed but values identical | Same extracted channels and numerical results |
| E2E-09 | Inspect output workbook via Excel parser | One `MasterDetail` sheet, native numeric and date cells, dark-blue/white headings, number/date formats, blank D/E fields |
| E2E-10 | Recognized summary headers plus unknown non-date header | Approved summaries skipped; unknown header causes failure with no output |
| E2E-11 | Mixed fractional sales values with >2 decimal places | Unrounded decimal totals reconcile internally; Excel displays 2 places without intentional pre-rounding |

### 9.3a v1.4 Part 1 hierarchy acceptance additions

The complete synthetic hierarchy test matrix **UT-HIER-01 through UT-HIER-14** is defined in **Step 13.7** and is mandatory in addition to the earlier UT/E2E cases. For every critical structural-error fixture, validate the exact error class and locations; for full-run fixtures, verify that no final MasterDetail is published. Step 13 takes precedence wherever a test expectation in Step 9 conflicts with an approved v1.4 hierarchy rule.

### 9.4 Acceptance criteria

Implementation may be accepted only when all criteria below are demonstrated using synthetic Excel files:

- [ ] A valid `TransactionDetail.xlsx` from the first worksheet is parsed with Row 1 date headers and date columns E onward.
- [ ] Both approved channel rules work, and selected parent totals/grand totals are excluded.
- [ ] Hierarchy carry-forward/reset, raw-before-fill sentinel classification, no D carry-forward, before/after parent totals, invalid child without parent fallback, and multi-candidate ambiguity pass UT-HIER-01 through UT-HIER-14.
- [ ] Original selected channel names and first-appearance order are preserved exactly; duplicate selected names fail.
- [ ] Date columns accept Excel dates and day-first/ISO text dates; invalid and duplicate dates fail.
- [ ] Only the four explicitly approved summary headers are skipped; unrecognized non-date headers cause critical errors.
- [ ] Blank amounts become 0; negative values remain negative; bad values fail with useful locations.
- [ ] Daily section has one row per **actual** source date and channel columns starting at F, with D/E blank.
- [ ] Monthly section uses actual date counts and accurately reconciles per-channel sales.
- [ ] All sums use decimal arithmetic with no intermediate rounding; display has two decimal places only. Excel numeric-storage precision limits are explicitly handled.
- [ ] Output contains exactly one `MasterDetail` worksheet; headers are dark blue/white, sales display 2 decimals without rounding calculations.
- [ ] Each successful execution creates a new uniquely named file, never overwriting an existing output.
- [ ] On any critical error, all safely detectable errors are reported; no final workbook is created and existing files are untouched.
- [ ] Source is never modified; no data transmission, unapproved dependency installation, commit or push occurs.
- [ ] Unit tests and end-to-end tests actually execute and their results are reported, including any failures.

### 9.5 Test evidence and implementation gate

After implementation is **separately authorized**, provide: test commands, test results (passed/failed/skipped), sample generated synthetic workbook, reconciliation summary, and any unresolved risks. Until then, this document is **requirements only**; do not generate source code, install dependencies, run Git write operations, or push changes.

---

## Step 10 — Approved Summary Headers & Decimal Precision (Final, additional requirement)

### 10.1 Explicit non-date summary header allowlist

While scanning **Row 1, columns E onward**, classify each header as either a recognized transaction date, an **explicitly approved summary header**, or an invalid/unknown header. The only approved summary headers are:

- `Total Sales`
- `Grand Total`
- `Total`
- `ยอดขายรวม`

Rules:

1. For identification only, trim leading/trailing whitespace; compare English names case-insensitively. Do not alter source values or channel names.
2. Recognized summary columns are **excluded** from daily date extraction, sales calculations, and monthly reconciliation.
3. Use **whole-cell exact matching** after the normalization above; do not skip a header just because it contains `Total` or any approved phrase as a substring.
4. An unrecognized, nonblank, non-date header in the E+ range is a **critical validation error**; include its column and `Row 1` in the validation report. Do not silently skip it. A blank E+ header associated with populated data is also an error; report ambiguous blank headers rather than guessing column meaning.
5. Detect invalid or duplicate calendar dates as already specified in Steps 2, 4 and 8.
6. Treat this allowlist as the approved default. Changes require an explicit requirement/configuration change and corresponding tests; do not infer new summary labels from data.

### 10.2 Numeric precision, aggregation, and display

1. Use **decimal arithmetic** (base-10 fixed/arbitrary precision suited to the chosen language) for all internal sales arithmetic and monthly aggregations. Do **not** accumulate money using native binary floating-point sums.
2. Convert source numeric cells into a decimal representation without adding unwarranted precision. Preserve whatever decimal digits are actually recoverable from the XLSX stored value. Do not claim to reconstruct digits already lost through Excel's numeric storage.
3. Preserve signs, zero, and source decimal precision, including **more than two fractional digits** where available.
4. **Do not round intermediate values** or round before monthly totals. Compute each month's channel total by summing its unrounded daily values.
5. Reconcile monthly totals against unrounded internal daily values using decimal equality; a mismatch is a critical error.
6. Write sales values as **numeric cells**, not text. Apply Excel display number format `#,##0.00` to daily and monthly sales. Two decimal places are **presentation only**, not a request to quantize values to cents.
7. Be explicit about the limitation that standard Excel numeric cells use floating-point storage and have finite significant-digit precision. On writing output, check that conversion to XLSX numeric values does not silently discard precision beyond the supported policy; if safe preservation is not possible, report the limitation/error rather than silently truncate. Internal decimal reconciliation remains authoritative.
8. Do not infer a currency, rounding mode for monetary settlement, or fixed decimal scale unless separately specified.

### 10.3 Required verification

- Test every approved summary-header label (including trimmed English header variants), a partial-match lookalike, and an unknown non-date header.
- Test `0.1 + 0.2`, negative amounts, zeros, blank→0, and values with more than two decimal digits.
- Confirm the internal monthly total equals the unrounded decimal sum of daily values.
- Confirm the workbook shows two decimal places while retaining numeric cell types and the recoverable source precision within documented XLSX limitations.
- If any critical header/precision validation fails, collect other safely detectable errors and do not generate a final MasterDetail file.

**Mandatory rules:** Skip only explicitly approved summary headers. Use decimal arithmetic without intermediate rounding; show two decimals in Excel while protecting the source precision to the extent the file format permits.


## Open points for final review (not permission to guess)

1. **Source source-row boundary:** Confirm whether Row 2 is a secondary header or actual data, and which row begins channel data.
2. **Hierarchy interpretation:** Confirm handling when merged cells and parent labels appear only once; parent totals must be distinguished from terminal channel totals without relying on formatting.
3. **Monthly physical layout:** Confirm exact column positions for Month, Days and channel headers in the Monthly section (the Daily section is explicitly fixed A–E and F+).
4. **Date formats:** `DD/MM/YYYY` is the approved primary text convention; handling of other formats should raise an informative error rather than guess.
5. **Formula source sales:** Resolved in Step 11: use cached numeric results; missing or nonnumeric caches are critical errors.

None of these points authorizes proceeding to source code before explicit approval.


## Step 11 — Implementation Readiness Rules (v1.2)

**Status:** Confirmed policy additions. Implementation remains prohibited pending final review and explicit approval.

### 11.1 Formula cells and cached numeric values

1. In selected channel rows and valid date columns, accept numeric literal cells and formula cells only when a stored, usable **numeric cached result** exists in the XLSX package.
2. Read the cached result, **not** the formula expression. Never evaluate, recalculate, or replace formulas in the source workbook.
3. A formula with a missing cache, empty cache, nonnumeric cache, or Excel error cache is a **critical error**, even though an ordinary blank literal sales cell means zero.
4. Report source worksheet, row, column and error code, but do not include confidential sales amounts or formula contents in logs.
5. Cached results may be stale. Document that the program uses the value last saved in the source workbook and does not verify formula freshness.
6. If the chosen Excel library cannot distinguish formula-without-cache from an ordinary blank, implementation must use a reliable read strategy or stop with a critical validation error; it must not silently map such cells to zero.

### 11.2 Channel names and case-insensitive duplicate detection

1. Keep the original source channel label unchanged as the target column header (same spelling, case and internal spaces).
2. For **duplicate detection only**, compare channel labels using Unicode `casefold()` or a documented equivalent case-insensitive comparison. Do not change output labels as part of duplicate checking.
3. `Lazada`, `lazada`, and `LAZADA` are duplicates and must produce a critical error, including all selected source row numbers.
4. Do not silently aggregate, keep-first, rename or drop duplicate selected channels.
5. Treat an empty or whitespace-only candidate label as invalid. Do not silently invent a channel name.
6. Apply duplicate detection **after** excluding parent subtotal and grand total rows under the approved Step 3 rules.

### 11.3 Datetime handling

1. Accept valid Excel date/datetime cells and approved text date formats in Row 1, E onward.
2. If a date/datetime contains a time component, discard the time **for grouping and output dates**; do not shift time zones or convert the calendar day.
3. Detect duplicate headers **after** normalization to the calendar date. For example, `2026-10-01 09:00` and `2026-10-01 15:00` are duplicate dates and cause a critical error.
4. Respect the source workbook's Excel 1900/1904 date system; invalid/ambiguous numeric date serials must be reported, not guessed.
5. Exclude approved summary headers before validating date columns.

### 11.4 Strict numeric precision policy

1. Parse numeric values from the **stored XLSX numeric representation**, rather than converting through a binary float first whenever possible. Preserve the recoverable numeric value, not insignificant trailing zeros such as the difference between `1.2300` and `1.23`.
2. Use exact decimal addition for aggregation, with sufficient context precision or an exact integer-coefficient/exponent method; detect and reject any internal `Inexact` or `Rounded` signal in computations requiring exactness.
3. Do not round intermediate daily or monthly amounts. Display formatting `#,##0.00` does not change underlying values.
4. Apply the same output validation to **every daily value and every monthly total**. A value containing more than **15 significant decimal digits** is outside the approved conservative Excel-output policy and is a critical error. Leading zeros are not significant; trailing zeros count only if they are necessary to represent the numeric value. This threshold is necessary but **not sufficient** for acceptance.
5. Before publishing, serialize the workbook to a temporary XLSX and inspect its stored numeric tokens. Parse the tokens as Decimal and compare each against its exact internal Decimal value. **Any inequality is a critical precision error**; do not substitute a tolerance, round, or convert numeric cells to text to force acceptance.
6. Reject non-finite values, values outside the supported Excel numeric range, and any conversion that would overflow, underflow to zero, or alter the stored decimal numeric value. Report the affected date/channel or summary cell without logging transaction values.
7. Recognize the distinction between stored XLSX numeric tokens and binary floating-point behavior when Excel opens, recalculates or resaves the workbook. This policy guarantees only the validated stored decimal numeric representation, **not** exact binary representation or the results of future Excel formulas/recalculation.
8. Internal reconciliation uses strict Decimal equality. Reconcile independently: (a) source selected cells to daily matrix by date/channel, (b) daily matrix to monthly totals, and (c) serialized workbook values back to internal values.
9. For example, daily `0.004` and `0.004` may both **display** as `0.00`, while their exact monthly total `0.008` **displays** as `0.01`. This is expected and is not a reconciliation failure.
10. If the chosen library cannot guarantee or verify the above strict policy, stop with a critical error and request a design decision; do not claim strict precision compliance.

### 11.5 Summary header boundaries and normalization

1. Check headers in Row 1 starting at Column E through the last column containing a meaningful value or formula in the header or data region. Columns with formatting only do not extend the range.
2. Approved non-date summary headers remain exactly: `Total Sales`, `Grand Total`, `Total`, `ยอดขายรวม`.
3. For header identification only, trim leading/trailing whitespace including ordinary spaces, tabs, line breaks and nonbreaking spaces. Compare English summary labels case-insensitively; use whole-cell matches and **do not collapse internal whitespace**. For example, `Total  Sales` is not an approved match.
4. A blank header over a column with any populated cell or formula in the data region is a critical error. A wholly empty column may be ignored only if it has no meaningful values or formulas; formatting alone does not count as data.
5. Unknown non-date headers in the inspected range are critical errors; report Excel column coordinates.

### 11.6 Safe output publication and error collection

1. Perform all possible independent validations, collect errors, and do not publish a final workbook if any critical error exists.
2. Generate and validate the output as a temporary XLSX. Inspect numeric tokens and all expected rows/columns before publishing.
3. Publish to a newly chosen filename using a **no-overwrite, race-safe** operation; never replace an existing target even when concurrent executions select the same timestamp. If a collision occurs, choose a new suffix and retry safely.
4. Delete only temporary files created by the current execution on failure; preserve source and all previously generated outputs.
5. Do not install dependencies, create implementation source files, commit or push until the user explicitly approves.

### 11.7 Required additional unit and end-to-end tests

- Numeric formula with valid cached result: use cached value; missing, blank, text or error cache: critical error (not zero).
- Case-insensitive duplicate channel: `Lazada` and `lazada` cause critical error; output retains original casing when unique.
- Datetime headers on the same date at different times: duplicate-date critical error.
- Header ` Total Sales ` is recognized; `Total  Sales` is rejected; whitespace-only header with a formula below is rejected.
- `0.1 + 0.2` yields exact Decimal `0.3` internally; output stored token must round-trip to Decimal `0.3`.
- Two daily values `0.004` produce monthly `0.008` internally, with two-decimal display formatting.
- **Precision threshold example:** `1234567890123456` (16 significant digits) is a critical precision error, even if supplied as a valid source numeric token. `123456789012345` (15 significant digits) may proceed only if stored-token round-trip equality also passes.
- Monthly-total overflow of 15 significant digits must fail even when each daily value individually passes.
- Large-exponent and cancellation cases must be evaluated with exact internal Decimal arithmetic and strict stored-token verification.
- Output collision and concurrent runs must never overwrite any previous file.

### 11.8 Remaining blockers to resolve with real source structure

- CONFIRMED in Step 12: row 1 dates, row 2 Sales subheaders, row 3 first channel row, overall Grand Total marks end of data; footer and notes follow. Exact sentinel validation and tests are specified in Step 12.
- RESOLVED by Step 13 (v1.4 Part 1): hierarchy carry-forward/reset, parent subtotal exclusion, and candidate ambiguity. Remaining Step 13 edge cases require tests against the agreed synthetic fixtures.
- Confirm exact physical layout of Monthly Total (Month, Days and channel column positions) before implementation.
- Confirm Python runtime and approved Excel library; installation requires explicit approval.
- Establish practical workbook limits (file size, date columns, channel count and rows) and error behavior when limits are exceeded.

**Mandatory:** This is an instruction update and review artifact only. Do not implement source code until the user explicitly approves.


## Step 12 — Confirmed Source Row Boundaries (v1.3)

**Approval status:** User confirmed the source row layout is consistent across TransactionDetail workbooks. This section resolves the row-boundary portion of Step 11.8; other implementation blockers remain open.

### 12.1 Row roles

1. **Row 1:** Date headers from column E onward and recognized non-date summary headers, as defined in Steps 2 and 10–11.
2. **Row 2:** Sales subheaders, such as `Sales`, for the corresponding date/summary columns. Row 2 is **not** a transaction/channel record and must not contribute to totals.
3. **Row 3:** First row in the channel data region. Scan subsequent rows for candidate channels and totals until the overall Grand Total row is reached. The Grand Total row number is **dynamic** and must not be hardcoded to a sample row number.
4. **Overall Grand Total row:** The table-ending row has `Total` as the cell value in column A, with no further specific channel hierarchy in B–D. This is a sentinel, never a selectable channel or a sales row. Match `Total` as a whole cell after trimming outer whitespace and comparing case-insensitively.
5. **Rows below the overall Grand Total:** Footer text, notes, applied-filter descriptions and blank rows are outside the data region and must not be parsed as channels or sales. Do not rely on fill color, font or borders to locate these rows.

### 12.2 Boundary validation and fail-safe behavior

1. Identify the overall Grand Total sentinel by content and hierarchy position. Require exactly one unambiguous sentinel row and report a critical structural error if the sentinel is missing or multiple candidate sentinels exist in the inspected worksheet.
2. Never include the Grand Total sales amounts in the Daily matrix or Monthly Total computation.
3. When parsing hierarchy, exclude rows 1 and 2 as well as the Grand Total row and all rows below it.
4. Do not silently interpret an early, ambiguous `Total` in column A as the end of the table if B–D indicate a real hierarchy. Report the ambiguity.
5. Within the data region, fully blank channel cells require hierarchy-aware classification, not automatic creation of a blank channel; do not treat rows with meaningful sales and no resolvable channel as harmless blank rows.
6. Once the Grand Total row is reached, ignore footer/notes for sales calculations; leave the source workbook unchanged.
7. Grand Total is useful as a **diagnostic cross-check**, but it must never replace the approved detail-channel sum or silently repair missing channels. Reconciliation against the original overall total is not automatically an acceptance criterion because source summary categories and selected detail channels may differ; any proposed equality check needs separate business approval.

### 12.3 Added test cases (synthetic data only)

- **UT-ROW-01:** Header row 1 and sales subheader row 2 are excluded; first candidate sales channel is evaluated at row 3.
- **UT-ROW-02:** The Grand Total appears at two different row numbers in two fixtures; both runs detect the sentinel by values rather than absolute row number.
- **UT-ROW-03:** Footer text such as `Applied filters:` below the sentinel is ignored and does not produce a channel/date/sales validation error.
- **UT-ROW-04:** Missing Grand Total marker causes a critical structural error and no final output.
- **UT-ROW-05:** Two candidate Grand Total markers cause a critical structural error and no final output.
- **UT-ROW-06:** A parent subtotal in B–D is not mistaken for the overall Grand Total in column A.
- **E2E-ROW-01:** Generate the complete workbook from a synthetic fixture with row 1 date headers, row 2 Sales labels, row 3 onward channel data, a Grand Total sentinel, then footer notes. Verify no footer/summary values leak into the output.

### 12.4 Remaining decisions before implementation

This confirmation **does not** authorize guessing unresolved rules. Hierarchy carry-forward/reset and subtotal classification are now specified in Step 13 (v1.4 Part 1). Still require final approval for sentinel/footer scan scope, Monthly Total physical layout, strict numeric range/precision edge cases, formula variants and cached-result handling, approved runtime/library, and workbook limits. No source code, dependency installation, Git commit or push without explicit user authorization.


---

## Step 13 — Hierarchy & Parent Total Rules (v1.4 Part 1; Confirmed)

**Status:** Business rules confirmed by the user. Requirements only; **implementation is not authorized**.

### 13.1 Scope and precedence

1. Read the **first worksheet**, using row 1 for dates, row 2 for `Sales` subheaders, and rows **3 through the row immediately before the overall Grand Total sentinel** for sales-channel classification (Step 12).
2. Apply this Step **13** wherever it clarifies or conflicts with Step 3's ambiguous hierarchy interpretation or parent-total wording. Keep the confirmed **Step 11** policy for case-insensitive duplicate detection and decimal/formula handling. Step 12 controls the source row boundaries.
3. Only raw source cell values and merged-cell **structure** can determine the hierarchy. Never use fill color, text color, bolding, borders, or other visual styles to select channels.
4. Perform hierarchy structure resolution and parent/child classification **before sales validation**. Do not let the presence or validity of an amount decide whether a channel is a parent or a child.

### 13.2 Raw values, labels, and `Total` markers

1. Retain both the **raw cell value** and its **resolved hierarchical context** for each source data row; keep the Excel row number and A–D coordinates for diagnostics.
2. For structural recognition only, treat a null, empty string, or whitespace-only label as blank. Recognize `Total` as a full-cell keyword after trimming outer whitespace and comparing case-insensitively. `Total Sales` is **not** a hierarchy `Total` marker.
3. Preserve the original selected channel label exactly for output (spelling, capitalization, and internal whitespace). Do not replace it with a normalized/case-folded label.
4. A raw `Total` is a **control marker**, not an inherited parent or a channel name; do not carry it into following rows.
5. Test for the overall Grand Total sentinel using the **raw, unfilled A–D cells before any carry-forward**. The known sentinel has raw A=`Total` and raw B–D blank (including whitespace-only cells). It is never eligible for Rule 1 or Rule 2. The exact footer/sentinel re-scan scope remains for v1.4 Part 2; do not infer a resolution here.

### 13.3 Hierarchy context and carry-forward (A–C)

1. Treat A, B and C as possible **parent/context levels**. Resolve parent labels from merged cells covering the current row, or from the previous row **within the same still-active group** where the hierarchy clearly continues.
2. For an actual merged range, use its top-left anchor value only for cells **inside that exact merged range**. An anchor outside the inspected region or a conflicting overlap is a critical structural ambiguity.
3. For a blank **unmerged** cell at A, B or C, inherit the last valid active parent at that level **only if** no higher-level group boundary has been crossed and the surrounding structure makes the continuation unambiguous. Do not invent a parent when none exists; report an unresolved hierarchy if one is required to classify the row.
4. Process **explicit** nonblank labels from left to right. When an explicit label at hierarchy level A, B or C establishes a **new parent** different from the currently active label at that level, **reset all deeper active levels** before processing subsequent cells in that row (A resets B/C/D; B resets C/D; C resets D). This prevents inheriting children from a previous group.
5. Repeated explicit labels at the same level may refer to the same active parent if their resolved higher-level ancestry matches. A repeated name under a **different** ancestor is a different hierarchy path (but duplicate selected output names are still checked per Step 11).
6. `Total` markers do not create or update a parent context. A `Total` boundary must not cause an old child to leak into a later group. If an inherited parent could plausibly belong to different groups, mark the row **critical ambiguous** instead of guessing.
7. **Column D must never be carried forward**, even if the previous D value is part of a merged-looking visual area. Select D only when the **current row explicitly has a nonblank, non-`Total` D label**. When a real merge spans D over multiple sales rows, do not silently generate repeated channels: report a critical ambiguity for review.
8. Keep raw and resolved paths separate, e.g. raw `A='', B='', C='', D='OMS OnDemand Hyper'` may resolve through valid active parents to `A=1.GHS / B=OMS / C=OMS OnDemand HI / D=OMS OnDemand Hyper`. Inherited A–C values are context, **not new explicit channel candidates** by themselves.

### 13.4 Channel rules and precedence (exactly one candidate per row)

Apply both confirmed rules to each row in the data region **after** resolving valid parent context:

- **Rule 1 — Channel before `Total`:** a raw `Total` in B, C or D qualifies the immediately preceding A, B or C label as a candidate **only when the preceding level is nonblank and unambiguously resolved**. Do not choose any other ancestor as an additional channel merely because it was carried forward.
- **Rule 2 — Column D:** an **explicit** nonblank D cell that is not `Total` is a candidate, without requiring a marker to the right.

Selection and ambiguity:

1. Gather the candidate(s) identified by Rule 1 and Rule 2 on the current row, noting rule, hierarchy path, and cell reference.
2. **Zero candidates:** do not select a channel. If the row contains meaningful daily amounts but its channel classification is unresolved, report a critical structural error; do not silently drop the row.
3. **Exactly one candidate:** classify it provisionally; still check for parent/subtotal exclusion in §13.5.
4. **More than one candidate on one row:** **critical ambiguity**, even if candidate labels look identical. Never copy the same daily amounts to multiple output channels or silently pick one candidate.
5. An explicit `Total` with no valid immediate predecessor is a critical structural error. Do not bypass it by applying Rule 2 or falling back to a higher ancestor.
6. A selected channel name cannot be `Total`. Grand Total sentinel rows are excluded before candidate selection.

### 13.5 Parent totals versus terminal channels (two-pass structure check)

1. **Pass 1 — structural hierarchy:** resolve hierarchy paths for all rows in the data region and identify parent/descendant relationships **without validating sales amounts first**. Structural descendants may be detected even if their own sales cells are invalid or their candidate classification will later fail.
2. **Pass 2 — channel classification:** for each provisional candidate, check whether its hierarchy path is a prefix/ancestor of one or more more-specific structural child paths in the **same parent group**.
3. If a candidate represents a **parent subtotal** and such children exist, **exclude the subtotal row**, regardless of whether its children appear before or after it or whether child sales validation succeeds.
4. If a candidate is a **terminal channel** (no structural descendants) and satisfies Rule 1 or Rule 2, keep it as a selected channel, subject to duplicate and other validations.
5. If a child path or its sales amount is invalid, report the corresponding **critical error**; **never resurrect, substitute, or merge the parent subtotal** to hide or repair child errors.
6. Use complete resolved ancestry, not matching visible channel names alone, to identify actual parent/child relationships. Do not group same-name labels from unrelated parent paths as one hierarchy; nevertheless, the output-name duplication check still applies.
7. Never compute a grand total by adding both selected child sales and excluded parent subtotal sales.
8. **No guesswork**: where grouping boundaries or parent/child relationships cannot be established reliably, report a critical ambiguity and do not publish MasterDetail.

### 13.6 Case-insensitive duplicate selected channel names

1. **After excluding parent subtotals and the overall Grand Total**, collect selected channel labels.
2. Compare selected names with the previously approved **Unicode `casefold()`** rule. Do not add new Unicode normalization or trim/merge internal whitespace beyond the agreed matching rule.
3. If two different selected source rows produce names equal under `casefold()`, report a **critical duplicate channel error** listing all source rows. Do not sum, skip, or overwrite either record.
4. Keep the original source label for each unique output header; matching normalization is for uniqueness checks only.

### 13.7 Required synthetic fixtures and explicit expected results

Use synthetic examples (no real transaction data). Every fixture must provide raw A–D cells, sales cells where needed, merged ranges where relevant, and exact expected selected/excluded channels or error codes. The following minimum results are mandatory:

| Test ID | Raw structure / situation | Expected behavior |
|---|---|---|
| UT-HIER-01 | Row 3: B=`HATO`, C=`Total`; no structural children | Select `HATO` via Rule 1 |
| UT-HIER-02 | B=`OMS`, C=`OMS OnDemand HI`, D=`OMS OnDemand Gofresh`; next row has blank B/C, D=`OMS OnDemand Hyper` in the same group | Select both D labels, inheriting B/C **only as parent context**; do **not** inherit D |
| UT-HIER-03 | After OMS, an explicit new A group and B=`Lazada`, C=`Total` | Select `Lazada`; prior OMS C/D context is reset |
| UT-HIER-04 | C=`OMS Scheduled CC`, D=`Total`, no structural children below C | Select `OMS Scheduled CC` via Rule 1 |
| UT-HIER-05 | C=`OMS OnDemand HI`, D=`Total`, with D-level `Gofresh` and `Hyper` child paths elsewhere in same group | Exclude `OMS OnDemand HI` parent subtotal, regardless of subtotal appearing before/after children |
| UT-HIER-06 | B=`OMS`, C=`Total`, with more-specific child paths in same group | Exclude OMS subtotal even if it appears **before** those child rows |
| UT-HIER-07 | Parent subtotal is numeric, but one structural child has text `not-a-number` in a dated sales cell | Exclude subtotal; report critical invalid child sales; no final output; no fallback to parent |
| UT-HIER-08 | Row contains C=`Total` and explicit D=`Another Channel`, providing Rule 1 and Rule 2 candidates | Critical multiple-candidates ambiguity; do not choose either |
| UT-HIER-09 | A/B/C merged parent anchor covering two child rows; D explicitly different in each | Correct parent context on both rows; select each distinct D; do not rely on styling |
| UT-HIER-10A | D merged vertically across two sales rows | Critical ambiguous D-level channel; no final output |
| UT-HIER-10B | Row with blank D after a previous explicit D, without a Total marker | Do not carry D or select the previous D channel; if the row has sales amounts and no valid candidate, critical unclassified-sales-row error |
| UT-HIER-14B | Raw A=`Group`, B=`Total`, C=`Channel`, D=`Total` | Critical multiple-candidate ambiguity (A via B=`Total` and C via D=`Total`) |
| UT-HIER-11 | Two distinct selected rows named `Lazada` and `lazada` | Critical duplicate-name error via `casefold()`; no final output |
| UT-HIER-12 | Raw A=` Total ` and raw B–D blank (or whitespace-only) after channel rows | Detect overall Grand Total sentinel **before context filling**; do not select or inspect its sales for detail totals |
| UT-HIER-13 | Blank required parent after a group switch, no supporting merged range or active context | Critical unresolved hierarchy; do not infer old-group parent |
| UT-HIER-14 | Raw A=`Group`, B=`Total`, C=`Total`, D blank | Critical malformed hierarchy: C=`Total` has invalid immediate predecessor B=`Total`; not a valid multiple-candidate case |

**Additional acceptance checks:**

- Changing cell fill/font/borders does not change selected channels.
- A structural child with invalid sales never causes its parent subtotal to become selected.
- Parent totals are excluded whether they appear **before or after** their children.
- Each selected channel is associated with **one and only one** source row and retains exact raw output spelling.
- On any critical ambiguity, preserve the source/existing targets and create **no final MasterDetail**.
- Report errors with worksheet name, source row, A–D coordinates, and a non-sensitive explanation; do not log confidential sales amounts.

### 13.8 Definition of done for v1.4 Part 1

This hierarchy section is review-ready only when the fixtures above have deterministic classifications and are incorporated into the project's Step 9 acceptance-test matrix. Do **not** claim the entire program is implementation-ready: the remaining v1.4 parts must settle sentinel/footer re-scan scope, Monthly Total layout, numeric range/precision boundaries, formula variants, runtime/library approval, and workbook/resource limits.

**Non-implementation gate:** Do not create or modify source code, install dependencies, run transactions, commit, or push until the user explicitly approves implementation after completing all outstanding specifications.


## 14. v1.4 Part 1 — Revision 2: Confirmed Hierarchy Reset and Deterministic Classification

**Status:** Hierarchy reset approved; implementation remains **NOT APPROVED**. This section supersedes conflicting context-reset and test statements in §13. Tests below are specification fixtures, not executed tests.

### 14.1 Total marker reset (approved)

1. Examine raw A–D cells and recognize the overall Grand Total sentinel **before** applying any carry-forward. Sentinel processing is governed by the separately pending Part 2.
2. For ordinary data rows, resolve the current row and enumerate candidates **before** resetting context. Never retain `Total` as a parent value.
3. After processing a row with `Total` in hierarchy column X, clear active context **at X and every deeper level**. Keep only the levels strictly above X.
4. Reset table:

| Total column | Clear after current row | Preserve after current row |
|---|---|---|
| A | A, B, C, D | none |
| B | B, C, D | A |
| C | C, D | A, B |
| D | D | A, B, C |

5. Column D is never carried forward, whether or not a Total marker appears. An explicit changed parent at level X clears all deeper context before the new row is resolved.
6. Multiple `Total` markers are classified using Step 15 error precedence: an invalid marker predecessor causes malformed hierarchy; otherwise multiple valid channel candidates cause multiple-candidates ambiguity. Do not reset context sequentially to justify a candidate.
7. An immediate predecessor that resolves to blank or `Total` is malformed. A raw blank predecessor may inherit an unambiguous active A–C context before candidate detection, but never inherit `Total` itself. See Step 15.
8. Build a candidate's ancestry path only through the candidate's own column; exclude marker cells and stale lower-level context.
9. A parent subtotal is excluded based on **structural descendants** within the same contiguous group instance, regardless of child sales validity. Never replace an invalid child with the subtotal.
10. Compare active hierarchy labels and paths exactly; use Unicode `casefold()` solely for duplicate selected-channel-name detection. Do not silently apply additional Unicode normalization.
11. If a merged parent range conflicts with an explicit higher-level parent change, report a critical structural ambiguity rather than overriding the reset.

### 14.2 Deterministic fixture corrections

| Test | Raw cells / condition | Expected classification |
|---|---|---|
| UT-HIER-14A | A=`Group`, B=`Total`, C=`Total`, D blank | Critical malformed hierarchy; C marker's immediate predecessor is another marker. Not two valid candidates |
| UT-HIER-14B | A=`Group`, B=`Total`, C=`Channel`, D=`Total` | Critical multiple candidates; candidate at A and candidate at C; do not duplicate daily sales |
| UT-HIER-10A | D is vertically merged across distinct sales rows | Critical structural ambiguity; do not select a repeated D channel |
| UT-HIER-10B | D blank after prior row with D=`Gofresh`, no Total marker | D remains blank; do not select Gofresh again. If row has meaningful sales and no candidate, critical unclassified-sales-row error |
| UT-HIER-RESET-C | B=`OMS`, C=`Total`, D blank; next row B blank, C=`OMS Scheduled CC`, D=`Total` | After first row clear C/D, preserve A/B; second row selects OMS Scheduled CC, assuming no other structural ambiguity |
| UT-HIER-RESET-D | B=`OMS`, C=`OMS Scheduled CC`, D=`Total`; next row B/C blank and D blank | Preserve A/B/C after first row, clear D; second row selects no channel and must not inherit D |

### 14.3 Open decisions — do not infer

The user has approved the **Total reset rule** and the **non-contiguous returning parent path = new group instance** rule (Step 15). Before implementation, confirm: which parent levels are required for each row type; horizontal merged-cell policy; and classification of context-only/separator rows with or without numeric sales. Fixture sets must include full raw A–D values, row numbers, and merged ranges before claiming deterministic acceptance coverage.

**Next specification part:** Sentinel & Footer Detection (Part 2). Do not generate program source code, install dependencies, commit, or push.


## Step 15 — v1.4 Part 1 Revision 3: Candidate Precedence and Contiguous Group Identity

**Status:** Confirmed policy for returning parent paths; technical clarifications proposed for Codex review. **Implementation NOT APPROVED.** This section supersedes contradictory wording in Steps 13–14. Fixtures below are specifications, not executed tests.

### 15.1 Approved contiguous group identity

1. A **group instance** is a contiguous run of rows belonging to the same resolved parent hierarchy, with no intervening different parent group or confirmed boundary. Track an instance identifier in source row order, not only a text path.
2. When an explicit parent path changes to a different parent group, end the previous instance and start a new instance. Clear lower-level inherited context as specified in Step 14.
3. If a previously seen parent path returns **after a different group**, start a **new group instance**. Do not automatically combine its descendants with the earlier instance for subtotal classification.
4. Compare hierarchy parent labels and paths exactly (case-sensitive, no implicit Unicode normalization). Continue to use `casefold()` for **duplicate final output channel names**, after subtotal exclusion. A new group instance does **not** exempt a repeated final channel name from duplicate validation.
5. Subtotal/descendant checks operate only inside the candidate's own group instance and use a path truncated at the candidate level; exclude `Total` and any stale deeper context.
6. Parent/child structure is classified before sales validation. A malformed child must not cause fallback to a parent subtotal.
7. A row containing a `Total` marker is not by itself proof that the entire ancestor group has ended: apply the approved level-specific context reset. Whether a subsequent row starts a new group instance depends on its resolved parent path and intervening group boundaries.
8. Blank separator and footer boundaries, and the detailed lifecycle of instances at such boundaries, remain open for Part 2. Do not guess these cases.

### 15.2 Candidate resolution and error precedence

1. Detect the raw grand-total sentinel before any context inheritance (Step 12; footer scope remains Part 2).
2. For ordinary data rows, apply explicit parent changes, then resolve permitted A–C inherited context; never inherit D.
3. For each `Total` in B, C, or D, inspect the **resolved** immediately preceding hierarchy cell. A raw blank predecessor is permitted only when unambiguously resolved from the active parent context within the same group. Neither blank unresolved predecessor nor `Total` is a valid channel name.
4. Evaluate Rule 1 and Rule 2 independently to enumerate **provisional** candidates. If any marker has an invalid immediate predecessor, report **ERR-HIER-MALFORMED** first, even if another candidate could be valid.
5. If all markers are structurally valid but the row yields more than one provisional candidate, report **ERR-HIER-MULTI-CANDIDATE**. Never assign a single sales row to multiple channels.
6. With exactly one valid provisional candidate, apply structural descendant exclusion to decide whether it is a final selected channel or an excluded parent subtotal. Do not label it final until this check completes.
7. Only after current-row classification, apply Step 14's level-specific `Total` context reset. If a row contains multiple markers and is invalid, validation records the error; no output is produced.
8. Preserve original channel labels in output. Validate duplicate final channel names using Unicode `casefold()`; report **ERR-CHANNEL-DUPLICATE** with all selected source row references.

### 15.3 Deterministic example fixtures (no merged ranges unless specified)

All blank cells below mean raw empty cells, not `Total`. These fixtures demonstrate the specific policy; remaining required-parent and separator cases are explicitly **not yet approved**.

| Test | Raw source rows (A, B, C, D) | Expected result |
|---|---|---|
| UT-HIER-14A-R3 | Row 3: (`Group`, `Total`, `Total`, blank) | **ERR-HIER-MALFORMED**: C marker has B=`Total` predecessor; no final channel |
| UT-HIER-14B-R3 | Row 3: (`Group`, `Total`, `Channel`, `Total`) | **ERR-HIER-MULTI-CANDIDATE**: A and C are two valid provisional candidates; no final channel |
| UT-HIER-15-R3 | Row 3: (blank, `HATO`, blank, blank); Row 4: (blank, blank, `Total`, blank), no structural children | Row 4 inherits B=`HATO`, selects `HATO` provisionally and finally; reset C/D after Row 4. Row 3 is a context row only if it has no sales data |
| UT-HIER-16-R3 | Row 3: (blank, `OMS`, `Total`, blank); Row 4: (blank, blank, `OMS Scheduled CC`, `Total`) | After Row 3 reset C/D, preserve B=`OMS`; Row 4 provisional `OMS Scheduled CC`; final selection conditional on no descendants; Row 3 OMS is excluded if it has structural descendants in same instance |
| UT-HIER-17-R3 | Row 3: (blank, `OMS`, `OMS Scheduled CC`, `Total`); Row 4: (blank, blank, blank, blank) | After Row 3 reset D only, retain B/C; Row 4 cannot inherit D and creates no D-channel candidate. If Row 4 contains sales, its classification requires the pending context-only-row policy |
| UT-HIER-18-R3 | Row 3: (blank, `OMS`, `Gofresh`, `Total`); Row 4: (blank, `Lazada`, `Total`, blank); Row 5: (blank, `OMS`, `Super`, `Total`) | Two distinct OMS group instances separated by Lazada. Do not use Row 5 as a descendant for a subtotal in the earlier OMS instance. Selected output names remain subject to global casefold duplicate validation |
| UT-HIER-19-R3 | Row 3: (blank, `OMS`, `Gofresh`, `Total`); Row 4: (blank, `oms`, `Hyper`, `Total`) | Distinct exact parent labels `OMS` and `oms`; selected child names `Gofresh` and `Hyper` are not duplicates |
| UT-HIER-20-R3 | Row 3: (blank, `OMS`, `Total`, blank); Row 4: (blank, blank, `Gofresh`, `Total`) with invalid sales text | Row 3 is a parent subtotal excluded by structural descendant Row 4; report invalid child sales, never fallback to Row 3; no final workbook |

**Fixture qualification:** Some examples intentionally leave A blank, following the established sample structure; whether A is mandatory for other row types remains an open blocker. Tests requiring a fully specified parent level, blank separator, or footer are not declared complete until those rules are approved.

### 15.4 Acceptance checklist amendments

- [ ] Step 9 hierarchy coverage explicitly includes UT-HIER-10A, UT-HIER-10B, UT-HIER-14A-R3, UT-HIER-14B-R3, UT-HIER-15-R3 through UT-HIER-20-R3, and the existing RESET-C/RESET-D fixtures.
- [ ] Tests assert provisional candidate(s), structural exclusions, final selected channel(s), context after reset, and exact error code.
- [ ] Full-run negative tests assert no final MasterDetail and no changes to original source or previously generated files.
- [ ] Group reappearance tests assert new instance identity and no cross-instance subtotal exclusion.
- [ ] Distinguish requirement-fixture review from actually executed Unit and End-to-End tests.

### 15.5 Remaining blockers — explicitly not approved

- Required parent levels for each row type; classification of blank/context-only/separator rows, especially rows with numeric zero or formulas.
- Horizontal merged-cell policy and conflicting merged parent anchors.
- Part 2 sentinel/footer scan scope and exact group boundaries at separators/footers.
- Monthly layout, numeric precision boundaries, formula variants, runtime/library choices, resource limits, and comprehensive test expected values.

**Do not generate source code, modify source Excel files, install dependencies, commit, push, or begin implementation without separate explicit user approval.**


## Step 16 — v1.4 Part 1 Revision 6 (incorporating Revision 4): Approved Multi-level Instances, Recovery, and Row Classification

**Status:** Hierarchy rules in this section are approved for requirements review only. **Implementation Not Approved.** For conflicting hierarchy statements, apply **Step 16 > Step 15 > Step 14 > Step 13 > Step 3**. Step 11 remains authoritative for casefold duplicate detection and precision; Step 12 remains authoritative for row boundaries until Part 2 refines sentinel/footer scanning. Supersede older references to “valid children” with **structural descendants**, independent of sales validity. This section does not resolve the remaining non-hierarchy blockers.

### 16.1 Multi-level group instances

1. Maintain distinct active instance identifiers at levels A, B, and C, when populated. A blank A is allowed when B is explicitly present; do not invent an A value. A C-level value requires a resolved B; a D-level channel requires resolved B and C. The hierarchy may begin at B.
2. An explicit changed A closes its previous A instance and all active B/C descendants. An explicit changed B closes its previous B instance and active C descendants while preserving A. An explicit changed C closes only its previous C instance while preserving A/B. D never establishes inherited context.
3. A new instance receives a new monotonic identifier, even if its label/path occurred earlier. A repeated exact parent label in a *continuous* active group retains its existing instance; a path returning after another group or a confirmed separator gets a new instance.
4. `Total` markers are evaluated on the current row before resetting context. A `Total` in B clears B/C; in C clears C; in D clears D only. Preserve ancestors above the marker. Close only the instances at and below the marker, not their ancestors.
5. Candidate path ends at the candidate's level, excluding `Total` and deeper inherited values. A parent subtotal is excluded only if structural descendants belong to **the same candidate-level instance**. Changing C within B does not close B or sever B's relationship with both C groups.
6. Compare hierarchy paths/labels exactly (case-sensitive, no implicit Unicode normalization). Only final output channel duplicate detection uses `casefold()` globally across all instances.
7. A structurally invalid child must never cause a parent subtotal to be selected as a fallback. Determine structural descendants before validating sales values.

### 16.2 Critical-error recovery for collect-all validation

1. Record the error and invalidate the hierarchy context derived from the malformed row. Never carry its values forward.
2. Clear affected active hierarchy context conservatively; if the affected level cannot be determined, clear all A/B/C context. Mark dependent classifications as blocked until a subsequent row supplies an independently resolvable explicit hierarchy.
3. Continue scanning independently resolvable rows and collect further errors. If a subsequent row requires invalidated inherited context, emit `ERR-HIER-CONTEXT-BLOCKED` instead of guessing a channel.
4. Do not use values from invalid rows to infer subtotal/child relationships. Do not publish any final workbook when a critical error exists.
5. The error recovery policy does not guarantee classification of every subsequent row; report validations blocked by missing trustworthy context separately from successfully checked rows.

### 16.3 Required parent levels and no-candidate rows

1. A may be blank when B is explicitly present. A valid C candidate requires resolved B. A valid D candidate requires resolved B and C. A B candidate needs no A.
2. A context-only row has hierarchy content but no valid channel candidate and **no sales data**, including no formula cells. It may establish or update active parent context without producing a channel.
3. A completely blank A–D row with no sales content is a **blank separator**; close active A/B/C instances and clear context. The sentinel and footer are handled separately under Step 12 and pending Part 2.
4. A row without a valid candidate that contains a sales numeric value (including numeric zero), nonempty text, an Excel error, or **any formula** in the relevant daily-sales region is a critical `ERR-HIER-UNASSIGNED-SALES`. Do not treat a formula with empty/missing cache as blank.
5. Blank D must never be inherited. A row with no candidate and no sales may be context-only or separator as above; it must not generate a duplicate channel.
6. For row classification, inspect sales only in recognized date columns within the established source data region; do not inspect summary columns or footer data as sales.

### 16.4 Merged-cell handling

1. Vertical merged cells in A–C may provide a parent anchor only when the entire range belongs to a compatible active parent group. Conflicting explicit parent changes within the merged range cause `ERR-HIER-MERGE-CONFLICT`.
2. Horizontal merges spanning multiple hierarchy levels A–D are unsupported: `ERR-HIER-HORIZONTAL-MERGE`.
3. Never inherit D, even if a vertical merge covers D. Treat merged D spanning data rows as `ERR-HIER-MERGED-D` when it could affect channel identification.
4. A merge crossing a confirmed separator, sentinel, or incompatible group boundary is a critical structural ambiguity. Do not use formatting to resolve it.

### 16.5 Acceptance fixtures HIER-21 through HIER-30

All cells shown are raw A/B/C/D. Unless stated otherwise, Row 1 contains dates, Row 2 contains Sales headers, there are no merges, and sales values in the recognized date column E are ordinary numeric literals. `—` means a truly blank cell. Instance IDs below are illustrative **relative identities** (same/different), not literal runtime ID strings. All full-run critical cases require no final workbook and unchanged source/previous outputs.

| Test | Raw rows (row: A, B, C, D; E sales) | Expected result |
|---|---|---|
| HIER-21 | 3: —, OMS, OnDemand, Gofresh; 10. 4: —, —, Scheduled, Total; 20 | B instance remains identical; C changes to new instance; provisional channels Gofresh (D) and Scheduled (C); final selection subject to same-instance structural descendant rule |
| HIER-22 | 3: —, OMS, OnDemand, Gofresh; 10. 4: —, Lazada, Total, —; 20 | New B instance at row 4; old C instance closed; row 4 candidate Lazada; no stale C |
| HIER-23 | 3: —, OMS, OnDemand, Gofresh; 10. 4: —, Lazada, Total, —; 20. 5: —, OMS, OnDemand, Super; 30 | OMS row 5 is a new B instance distinct from row 3; no cross-instance descendant exclusion |
| HIER-24 | 3: Group, Total, Total, —; 10 | `ERR-HIER-MALFORMED` (C has Total predecessor); no final channel |
| HIER-25 | 3: Group, Total, Total, —; 10. 4: —, Lazada, Total, —; 20 | Row 3 malformed; row 4 has explicit B and can be independently classified as provisional Lazada; full run fails due to row 3 |
| HIER-26 | 3: Group, Total, Total, —; 10. 4: —, —, Total, —; 20 | Row 3 invalidates context; row 4 requires inherited B; `ERR-HIER-CONTEXT-BLOCKED` (and row 3 malformed); no output |
| HIER-27 | 3: —, OMS, OnDemand, Gofresh; 10. 4: —, —, —, —; blank E. 5: —, —, Scheduled, Total; 20 | Row 4 is separator, clears context; row 5 lacks B and yields missing-parent/context-blocked critical error; no output |
| HIER-28 | 3: —, OMS, —, —; E=0 | No candidate with numeric zero sales: `ERR-HIER-UNASSIGNED-SALES`; no output |
| HIER-29 | 3: —, OMS, OnDemand, Gofresh; 10, with horizontal merge B3:C3 | `ERR-HIER-HORIZONTAL-MERGE`; no output |
| HIER-30 | 3: —, OMS, Gofresh, Total; 10. 4: —, Lazada, Total, —; 20. 5: —, OMS, gofresh, Total; 30 | After structural exclusions, `Gofresh` and `gofresh` collide under casefold: `ERR-CHANNEL-DUPLICATE`; no output |

**Additional assertions:** Test context before and after resets; candidate level and truncated path; parent instance identity; provisional vs final channel; exact error codes; no inherited D. For `HIER-21`, the parent/child exclusion must be verified separately with an explicit parent subtotal fixture. These fixtures define the listed behavior but do not replace full end-to-end fixtures with complete date headers and sentinel.

### 16.6 Remaining open items / implementation gate

- **Part 2:** sentinel/footer scanning scope, repeated sentinel policy, and how boundaries interact with merged ranges.
- Monthly Total exact row/column layout and Month value type.
- Numeric significant-digit/range boundaries and precise serialization acceptance criteria.
- Supported formula variants and cache handling outside selected sales cells.
- Python runtime/library and resource limits.
- Expand Step 9 acceptance checklist to include all Step 14–16 fixtures and complete synthetic end-to-end cases.

**IMPLEMENTATION NOT APPROVED.** This revision is a requirements artifact only. Do not write program source code, install dependencies, modify source data, commit, or push without separate explicit user approval.


---

## Step 17 — v1.4 Part 1 Revision 5: Verified-Context Recovery and Deterministic Hierarchy Tests

**Status: Approved for requirements review only. IMPLEMENTATION NOT APPROVED.** In hierarchy conflicts, **Step 17 > Step 16 > Step 15 > Step 14 > Step 13 > Step 3**. Steps 11 and 12 remain authoritative within their own scopes. This revision supersedes the conditional expected outcomes for HIER-21–30 in Step 16.5. It does not resolve Part 2 or other program-level blockers.

### 17.1 Two-pass classification and structural evidence

1. **Pass 1 (structure):** scan raw A–D and permitted resolved A–C context, assign provisional candidates and group instances, record structural descendant evidence, and collect hierarchy errors. Do not validate sales as a prerequisite to determining whether a structural child exists.
2. **Pass 2 (final selection):** exclude parent subtotal candidates when their same-instance structural descendants are established. If an invalid/ambiguous child might belong to a parent but the relationship cannot be proved, mark the affected parent classification **BLOCKED**, never **SELECTED** by absence of valid children.
3. A structurally established child with invalid sales still excludes its parent subtotal; separately report the child sales error. A structurally invalid child is not discarded as proof of 'no children'. Preserve the evidence and mark the potentially affected parent **BLOCKED** if association cannot be resolved.
4. No blocked, malformed, or invalid row contributes sales to the output. Any critical error prevents final publication, even if other rows were independently classified.
5. Track `provisional`, `excluded_subtotal`, `blocked`, and `final_selected` as distinct classification outcomes. Never use a parent amount as fallback for an invalid child.

### 17.2 Deterministic recovery table

Apply this table after recording the primary error. 'Invalidate' means the context is unavailable for inheritance until independently re-established. A fresh explicit B can establish a new B instance without A; explicit C requires verified B; explicit D requires verified B and C.

| Primary error / condition | Context after row | Next-row behavior |
|---|---|---|
| `ERR-HIER-MALFORMED` (invalid Total predecessor) | Invalidate A/B/C, because marker interpretation is unreliable | Independently explicit B may restart; inherited predecessor produces `ERR-HIER-CONTEXT-BLOCKED` |
| `ERR-HIER-MULTI-CANDIDATE` | Invalidate A/B/C; do not choose one candidate | Restart only with independently resolvable explicit hierarchy |
| `ERR-HIER-MISSING-PARENT` (ordinary missing required ancestor) | Invalidate missing level and lower; preserve independently verified higher ancestors | Missing required parent stays `ERR-HIER-MISSING-PARENT` until explicitly established |
| `ERR-HIER-CONTEXT-BLOCKED` (required inherited value invalidated by earlier error) | Preserve verified ancestors only; blocked levels remain unavailable | Emit context-blocked while the dependency persists; do not invent parents |
| `ERR-HIER-HORIZONTAL-MERGE` / `ERR-HIER-MERGE-CONFLICT` / `ERR-HIER-MERGED-D` | Invalidate A/B/C for the affected merged span; ignore its anchors as context | Rows within affected span remain blocked; after span, independently explicit hierarchy can restart |
| `ERR-HIER-UNASSIGNED-SALES` with otherwise valid hierarchy | Keep independently verified A/B/C context; no channel selected | Continue using verified context; no fallback from invalid sales row |
| Raw blank separator (no A–D values and no sales content) | Close all A/B/C instances and clear context **normally**, not as an error | A subsequent missing ancestor is `ERR-HIER-MISSING-PARENT`, not context-blocked |
| Valid `Total` marker | First classify row, then clear marker level and below | Ancestors above marker remain verified; reintroduced closed level gets a **new** instance |

**Error precedence:** `ERR-HIER-HORIZONTAL-MERGE`/merge conflict for an affected span; then `ERR-HIER-MALFORMED`; then `ERR-HIER-MULTI-CANDIDATE`; then `ERR-HIER-CONTEXT-BLOCKED` for an invalidated inherited dependency; then `ERR-HIER-MISSING-PARENT` for an ordinarily absent ancestor; then `ERR-HIER-UNASSIGNED-SALES` if no candidate and sales content exists. Collect additional independent diagnostics when safe, but use exactly one primary error per row. If a cell or row cannot be classified because prerequisites failed, report the dependent check as **blocked**, not as a fabricated secondary error.

**Row classification precedence:** check sentinel using raw cells before any inheritance (Step 12; exact footer scope pending Part 2); then identify merged-range errors; then classify a raw A–D blank row with no sales as a separator; otherwise resolve parent context and detect candidates. A raw blank A–D row with sales content is not a separator and is `ERR-HIER-UNASSIGNED-SALES` unless a higher-priority structural error applies. Recognized summary columns do not count as daily sales content.

### 17.3 Fixture conventions

All fixtures below are **independent**; reset state before each test. Row 1 has E1 as a valid Excel date, Row 2 has E2=`Sales`, and the indicated data rows begin at Row 3. A dash (`—`) means a physically blank cell. Unless specified, there are no merges, the source has a raw grand-total sentinel after the fixture rows, and no other rows exist. E is the single sales column. Tests with critical errors expect no final output, no source mutation, and no changes to prior outputs. Instance assertions use relative identity, not implementation-specific numeric IDs. A full-run test must include a valid sentinel according to Step 12. **These are requirements fixtures, not executed test results.**

### 17.4 Corrected acceptance fixtures HIER-21–HIER-30

| ID | Raw data rows `(row: A, B, C, D; E)` | Exact expected result |
|---|---|---|
| **HIER-21** | `3: —, OMS, OnDemand, Gofresh; 10`; `4: —, —, Scheduled, Total; 20` | Same B instance in rows 3–4; C instance changes; final selected channels **Gofresh** (row 3) and **Scheduled** (row 4); no errors. No parent subtotal is present. |
| **HIER-22** | `3: —, OMS, OnDemand, Gofresh; 10`; `4: —, Lazada, Total, —; 20` | B instance changes; old C closes; final **Gofresh** (row 3), **Lazada** (row 4); C after row 4 is empty; no errors. |
| **HIER-23** | `3: —, OMS, Total, —; 100`; `4: —, Lazada, Total, —; 20`; `5: —, OMS, OnDemand, Super; 30` | OMS row 3 and OMS row 5 belong to **different B instances**; row 5 does not exclude row 3. Final **OMS** (row 3), **Lazada** (row 4), **Super** (row 5); no duplicate final names. |
| **HIER-24** | `3: Group, Total, Total, —; 10` | Primary `ERR-HIER-MALFORMED` at row 3 (C marker's predecessor is Total); A/B/C context invalidated; no output. |
| **HIER-25** | `3: Group, Total, Total, —; 10`; `4: —, Lazada, Total, —; 20` | Row 3 malformed and clears A/B/C; row 4 establishes explicit B=Lazada and independently yields provisional/final **Lazada**; full run fails due to row 3. |
| **HIER-26** | `3: Group, Total, Total, —; 10`; `4: —, —, Total, —; 20` | Row 3 primary `ERR-HIER-MALFORMED`; row 4 primary **`ERR-HIER-CONTEXT-BLOCKED`** (B inherited dependency invalidated), not malformed/missing-parent; no output. |
| **HIER-27** | `3: —, OMS, OnDemand, Gofresh; 10`; `4: —, —, —, —; blank`; `5: —, —, Scheduled, Total; 20` | Row 4 raw blank separator closes B/C; row 5 primary **`ERR-HIER-MISSING-PARENT`** (B absent after normal reset), not context-blocked; no output. |
| **HIER-28** | `3: —, OMS, —, —; 0`; `4: —, —, OnDemand, Gofresh; 10` | Row 3 primary **`ERR-HIER-UNASSIGNED-SALES`** despite numeric zero; verified B=OMS remains available; row 4 independently resolves B and selects **Gofresh**; full run fails due to row 3. |
| **HIER-29** | `3: —, OMS, —, Gofresh; 10` with actual horizontal merged range **B3:C3** (B3 anchor=`OMS`, C3 non-anchor physically blank) | Primary **`ERR-HIER-HORIZONTAL-MERGE`**; do not use merged B/C as inherited context; no output. This fixture does **not** put a value in C3. |
| **HIER-30** | `3: —, OMS, Gofresh, Total; 10`; `4: —, Lazada, Total, —; 20`; `5: —, OMS, gofresh, Total; 30` | Rows 3 and 5 both final selected in distinct OMS B instances; `casefold()` collision `Gofresh` vs `gofresh` produces **`ERR-CHANNEL-DUPLICATE`** referencing rows 3 and 5; no output. |

### 17.5 Additional deterministic acceptance fixtures

| ID | Raw rows / condition | Expected |
|---|---|---|
| **HIER-31** | `3: —, OMS, Total, —; 100`; `4: —, —, Scheduled, Total; 20` | Row 3 provisional OMS, excluded as parent subtotal by same B-instance descendant row 4; final **Scheduled** only; B stays active after C=`Total`. |
| **HIER-32** | `3: —, OMS, OnDemand, Gofresh; 10`; `4: —, —, —, —; blank`; `5: —, OMS, OnDemand, Hyper; 20` | Separator closes instances; new OMS B and OnDemand C instances at row 5; final **Gofresh**, **Hyper**. |
| **HIER-33** | `3: —, OMS, Total, —; 100`; `4: —, —, Scheduled, Total; text 'INVALID'` | Structural descendant row 4 excludes parent OMS before sales validation; row 4 invalid sales is critical; never fallback to OMS; no output. |
| **HIER-34** | `3: —, OMS, Total, —; 100`; `4: —, —, Total, Total; 20` | Row 4 malformed (D marker predecessor is Total); any parent classification that depends on interpreting row 4's child relationship is **BLOCKED**, not selected by assuming no child; no output. |
| **HIER-35** | `3: —, OMS, OnDemand, Gofresh; 10`; `4: —, —, —, —; blank`; `5: —, —, Scheduled, Total; 20` | Normal separator causes missing-parent at row 5, not context-blocked; same assertion as HIER-27, included for regression. |

### 17.6 Test assertion requirements

For every fixture assert: raw A–D and E types, source row numbers, relevant merged ranges, resolved parent values before/after row, same/different A/B/C instance identities, provisional candidates, structural descendants, excluded/blocked/final selections, primary error codes, and any blocked dependent checks. Do not assert a final selected channel for a run that is invalid as a whole without distinguishing its **locally classifiable** rows from overall run success. Negative full-run fixtures assert no new final workbook and unchanged existing files; pure unit tests assert classifications and error codes only.

### 17.7 Superseded text and outstanding blockers

- Replace older Step 3 wording 'valid child channels' with **structural descendants**. Step 16.2's ambiguous 'clear affected context conservatively' is superseded by the deterministic table in Step 17.2. HIER-21–30 in Step 16.5 are superseded by Step 17.4. Step 9 acceptance checklist must include HIER-21–35 and Step 17 error-recovery assertions.
- **Still open:** Part 2 sentinel/footer scan scope and merge boundaries; exact Monthly layout; decimal significant-digit and numeric-range thresholds; formula variants; runtime/library and resource limits; complete end-to-end fixtures. No implementation is authorized.

**Implementation gate:** Do not generate source code, install dependencies, edit transaction data, commit, or push. Await explicit approval after remaining blockers are resolved.


---

## Step 18 — v1.4 Part 1 Revision 6: Verified / Absent / Untrusted Context and Error Precedence

**Status: APPROVED REQUIREMENTS FOR REVIEW ONLY — IMPLEMENTATION NOT APPROVED.** For hierarchy interpretation, Step **18 > 17 > 16 > 15 > 14 > 13 > 3**. Step 18 supersedes any inconsistent recovery and conditional fixture expectations in earlier steps; source boundaries and financial precision remain subject to their own later requirements. Do not write source code or install dependencies.

### 18.1 Context states and instance closure

Each hierarchy level A, B and C maintains (a) a state, (b) its verified label, if any, and (c) a distinct group-instance identity.

- `VERIFIED`: established by an explicit value or a permitted vertical-merge anchor, or safely inherited from a previously verified level within the same active instance.
- `ABSENT`: no active label because none has been established, or because a valid `Total` reset or raw blank separator normally closed the level.
- `UNTRUSTED`: inheritance was invalidated by a preceding structural error; its former label must not be used to classify later rows.
- A level closed by an error or reset cannot silently revive its old instance. An explicit valid restart creates a **new instance ID**, even if its label is identical.
- On a malformed or multi-candidate row, close all active A/B/C instances and set their inheritance state to `UNTRUSTED`. A subsequent explicit B may independently establish a new B instance without A; explicit C requires verified B; explicit D requires verified B and C. Any still-unestablished required ancestor retains `UNTRUSTED` only when it was invalidated by an earlier error; ordinary absence is `ABSENT`.
- A raw blank A–D row with no sales content is classified as a separator **before inheritance** and sets A/B/C to `ABSENT`, closing all active instances. A valid `Total` clears its marker level and lower levels to `ABSENT` only after the current row's candidate detection.
- A missing-parent error arising from `ABSENT` does not turn the absent level into `UNTRUSTED`; dependent future rows remain missing-parent until an explicit valid parent appears. Other independently verified ancestors may remain active.

### 18.2 Error classification and primary-error precedence

Apply evidence-aware precedence to each row, and report one primary error per row plus independently provable secondary diagnostics. A blocked prerequisite is **not** proof of a malformed predecessor.

1. **Merge-origin structural failure**: `ERR-HIER-HORIZONTAL-MERGE`, `ERR-HIER-MERGE-CONFLICT`, or `ERR-HIER-MERGED-D` at the originating merge/range. Where multiple merge failures overlap, choose the error tied to the earliest affected coordinate in row-major order; for the same coordinate, use the preceding code order. Report other independent merge errors as secondary diagnostics.
2. **Proven malformed predecessor**: `ERR-HIER-MALFORMED` only if the actual or reliably resolved immediate predecessor is demonstrably `Total` or otherwise invalid. A predecessor unavailable because of earlier `UNTRUSTED` context does **not** qualify.
3. **Proven multiple candidates**: `ERR-HIER-MULTI-CANDIDATE` when two or more independently valid candidates occur in the row and no higher-priority error applies.
4. **Blocked prerequisite**: `ERR-HIER-CONTEXT-BLOCKED` when a required inherited predecessor cannot be verified because its context is `UNTRUSTED` from a previous error. This takes precedence over a speculative malformed or missing-parent diagnosis.
5. **Ordinarily absent prerequisite**: `ERR-HIER-MISSING-PARENT` when a required parent is `ABSENT` (never established or normally reset), not `UNTRUSTED`.
6. **Sales without a channel**: `ERR-HIER-UNASSIGNED-SALES` when there is sales content but no candidate and no higher-priority structural error.

For merge recovery, report the originating merge error at its origin; dependent later rows inside its affected span receive primary `ERR-HIER-CONTEXT-BLOCKED` if their classification relies on that damaged merge. Do not repeat the same merge-origin primary error on each dependent row. If a later row independently contains another provable merge defect, report that separate defect at its own origin. After the affected span, an independently explicit valid hierarchy may restart with new instance IDs.

### 18.3 Parent classification with structural evidence

Classify each parent subtotal candidate using evidence confined to its **own verified group instance**. Structural evidence is collected before sales validation.

1. `EXCLUDED_SUBTOTAL`: at least one **proven** same-instance structural descendant exists. Keep this status even if another ambiguous/invalid child appears later. A child's sales error is reported separately; no fallback is allowed.
2. `BLOCKED`: no proven descendant is available **and** an invalid/ambiguous child could affect this parent's classification based on the verified ancestry and active instance **immediately before the invalid row**. Do not infer association from channel-name similarity or from untrusted labels.
3. `FINAL_SELECTED`: a provisional parent candidate with no proven descendant and no relevant blocked structural evidence, after all rows in that instance have been scanned and validated.
4. If an invalid row has **no verified ancestry** connecting it to any parent, do not arbitrarily block unrelated parents. Record the hierarchy error and dependent checks as blocked; the entire run still fails and publishes nothing.
5. A new explicit parent or ancestor change establishes a different instance and cannot retrospectively block a previous instance. A subtotal must never be selected as a replacement for an invalid child.

**HIER-34 deterministic interpretation:** row 3 `B=OMS,C=Total` establishes a provisional OMS parent in a verified B instance. Row 4 `B=blank,C=Total,D=Total` has verified inherited B=OMS but a provably malformed D marker (its predecessor is C=`Total`). The invalid row is associated with the active OMS B instance from its verified pre-row ancestry; no proven descendant is available. Therefore **OMS row 3 = BLOCKED** (not `FINAL_SELECTED`), row 4 primary = `ERR-HIER-MALFORMED`, then A/B/C are closed and marked `UNTRUSTED`. No output is published.

### 18.4 Deterministic fixture extensions

Fixtures are independent; each has E1 = a valid Excel date, E2 = `Sales`, all unspecified A–D cells physically blank, no merges unless stated, and a valid raw grand-total sentinel after the last data row. `E` values are numeric unless explicitly stated. IDs are relative identities, not literal generated identifiers. Critical full-run tests expect no new final workbook and unchanged source and previously generated outputs. Each unit test checks primary error, pre/post context states, provisional and final classification, and instance closure.

| ID | Raw rows (row: A, B, C, D; E) | Expected result |
|---|---|---|
| HIER-26-R6 | `3: Group, Total, Total, —; 10`; `4: —, —, Total, —; 20` | Row 3 primary `ERR-HIER-MALFORMED`; A/B/C `UNTRUSTED` and prior instances closed. Row 4 predecessor B is unavailable due to prior error: primary **`ERR-HIER-CONTEXT-BLOCKED`**, not malformed or missing-parent; no selected row 4. |
| HIER-29-R6 | `3: —, OMS, —, Gofresh; 10` with horizontal merge B3:C3; `4: —, —, OnDemand, Hyper; 20`; `5: —, Lazada, Total, —; 30` | Row 3 primary `ERR-HIER-HORIZONTAL-MERGE`; damaged B/C context is `UNTRUSTED`. Row 4 primary `ERR-HIER-CONTEXT-BLOCKED` due to missing verified B. Row 5 explicit B independently restarts and is locally classified as Lazada; full run fails. |
| HIER-33-R6 | `3: —, OMS, Total, —; 100`; `4: —, —, Scheduled, Total; E4=text INVALID` | Row 4 structurally selected Scheduled and causes **`ERR-SALES-INVALID-NUMERIC`** at E4; row 3 OMS remains `EXCLUDED_SUBTOTAL` due to proven same-instance descendant. No parent fallback; no output. |
| HIER-34-R6 | `3: —, OMS, Total, —; 100`; `4: —, —, Total, Total; 20` | Row 4 primary `ERR-HIER-MALFORMED` at D4. Verified pre-row B=OMS associates uncertain child evidence with row 3's B instance. Row 3 classification **BLOCKED** (not selected); A/B/C `UNTRUSTED` after row 4; no output. |
| HIER-36 | `3: —, OMS, OnDemand, Gofresh; 10`; `4: Group, Total, Channel, Total; 20`; `5: —, Lazada, Total, —; 30` | Row 4 primary `ERR-HIER-MULTI-CANDIDATE` (valid A and C predecessors); all active A/B/C instances closed/`UNTRUSTED`. Row 5 explicit B=Lazada starts a new B instance, locally selected Lazada; full run fails. |
| HIER-37 | `3: —, —, Scheduled, Total; 20`; `4: —, —, OnDemand, Hyper; 10`; `5: —, OMS, OnDemand, Gofresh; 30` | Rows 3 and 4 primary `ERR-HIER-MISSING-PARENT` because B is ordinarily `ABSENT`, not untrusted. Row 5 explicit B=OMS restarts with new verified B/C instances and locally selects Gofresh; no output due to earlier errors. |
| HIER-38 | `3: —, OMS, OnDemand, Gofresh; 10`; `4: Group, Total, Total, —; 20`; `5: —, OMS, OnDemand, Hyper; 30` | Row 4 primary malformed, closes previous OMS B instance and invalidates A/B/C. Row 5 explicit OMS establishes a **different** B instance (even though label repeats); locally selected Hyper; full run fails. |
| HIER-39 | `3: —, OMS, Total, —; 100`; `4: —, —, Scheduled, Total; 20`; `5: —, —, Total, Total; 30` | Row 4 is a proven structural descendant of row 3 in the same B instance; row 3 remains **EXCLUDED_SUBTOTAL** despite row 5 malformed/ambiguous child. Row 5 primary `ERR-HIER-MALFORMED`; no fallback or output. |
| HIER-40 | `3: —, OMS, Total, —; 100`; `4: —, —, Scheduled, Total; 20`; `5: —, Lazada, Total, —; 30`; `6: —, —, Total, Total; 40` | Row 3 OMS stays **EXCLUDED_SUBTOTAL** based on row 4. Row 6 malformed is associated only with the verified Lazada B instance immediately before row 6, not the closed OMS instance. OMS remains excluded; row 6 primary `ERR-HIER-MALFORMED`; no output. |

### 18.5 Acceptance checklist and supersession

- Step 9's acceptance checklist **must include** HIER-21–40 and their revision-specific replacements; HIER-26-R6, 29-R6, 33-R6 and 34-R6 supersede conflicting older fixture expectations.
- Test each primary error and blocked dependent check separately; test `VERIFIED`/`ABSENT`/`UNTRUSTED` states and closure/new instance identity explicitly.
- Check `FINAL_SELECTED`, `EXCLUDED_SUBTOTAL`, `BLOCKED`, and locally classified rows separately from overall run success.
- Negative end-to-end runs must leave source and existing output files unchanged and publish no new final workbook. Unit tests need not create output files.
- All hierarchy test fixtures are **specified**, not executed. Additional tests for sentinel/footer, Monthly layout, numeric range, formula variants, runtime and resource limits remain outstanding.

### 18.6 Implementation gate

**IMPLEMENTATION NOT APPROVED.** This is a requirements-only revision. Do not generate application source code, install dependencies, modify source Excel files, commit, or push. Next phase: Part 2 — Sentinel & Footer Detection, followed by remaining program-level blockers and separate user approval.


---

## Step 19 — v1.4 Part 1 Revision 7: Candidate Validation, Parent Snapshot, and Merge Recovery

**Status: CONFIRMED FOR REQUIREMENTS REVIEW; IMPLEMENTATION NOT APPROVED.**
For hierarchy rules, precedence is **Step 19 > 18 > 17 > 16 > 15 > 14 > 13 > 3**.
Where older fixtures or text conflict with this section, the explicit replacements below prevail. This section does not authorize coding, dependency installation, or Git writes.

### 19.1 Three-stage candidate pipeline

1. **Discovery:** From raw and permitted resolved A–D values, identify *potential* Rule 1 candidates (name immediately before `Total`) and Rule 2 candidates (nonempty D other than `Total`). Discovery does not imply structural validity.
2. **Structural validation:** Check marker predecessor, required ancestors, active context trust, merged-cell restrictions, and uniqueness of structurally valid candidates. A channel in C requires a verified B ancestor; a channel in D requires verified B and C. A `Total` marker is never a verified ancestor. A candidate missing a required ancestor is invalid, even if its immediate predecessor is a nonempty label.
3. **Final selection:** After all group instances and structural descendants have been examined, exclude parent subtotals, block unresolved parents, and check duplicate selected names with `casefold()`. No invalid or merely potential candidate becomes a final channel.

**Multiple-candidate error** requires at least two *structurally valid* candidates in the same row. A potential candidate failing required-parent validation does not count toward that threshold. Proven malformed predecessor errors remain higher priority. Where an invalid candidate lacks a required parent because a control marker occupies that position, classify its missing ancestor as `ERR-HIER-MISSING-PARENT` under this revision; report any independent proven malformed marker as a secondary diagnostic only when applicable.

**Explicit replacement for UT-HIER-14B / UT-HIER-14B-R3 / HIER-36:** the fixture `A=Group, B=Total, C=Channel, D=Total` has potential candidates `Group` (A) and `Channel` (C). `Channel` fails because B is `Total`, not a verified parent. Primary error is **`ERR-HIER-MISSING-PARENT`**, not `ERR-HIER-MULTI-CANDIDATE`. No output is published. The previous expectations are superseded. An actual multiple-valid-candidates fixture must be separately constructed and proven to satisfy all required-parent rules before it can be used as an acceptance test; do not fabricate one.

### 19.2 Parent association snapshot

For each row, retain three separate records:
- **Pre-row snapshot:** verified active A/B/C labels and instance IDs before applying the row.
- **Post-explicit-change snapshot:** apply explicit valid higher-level changes and close the corresponding lower-level instances. Preserve only ancestry independently verified in the current row or still verified under unchanged ancestors.
- **Error snapshot:** the trustworthy ancestry remaining after detecting a structural error; never resurrect an invalidated instance.

Associate a defective child with a parent only if its **post-explicit-change verified ancestry** proves membership in that parent's current group instance. If the row explicitly changes B from `OMS` to `Lazada`, errors on that row cannot block the closed OMS instance. If the new parent is malformed or unverified, do not guess an association from the previous label or from name similarity.

If a parent has at least one proven same-instance structural descendant, it stays `EXCLUDED_SUBTOTAL` even when another child is invalid. If it has no proven descendant but verified evidence of a relevant uncertain child, mark it `BLOCKED`. `BLOCKED` is a classification diagnostic tied to the root error, not an additional independent critical-error count.

### 19.3 Deterministic merge recovery

| Condition | State transition and expected primary error |
|---|---|
| Horizontal merge crossing A–D hierarchy levels | At origin: `ERR-HIER-HORIZONTAL-MERGE`; close A/B/C instances and mark A/B/C `UNTRUSTED` |
| Vertical merge conflicts with explicit parent change at level L | At origin: `ERR-HIER-MERGE-CONFLICT`; close instance at L and all lower levels; mark these levels `UNTRUSTED`; independently verified ancestors above L remain |
| Merged D or unsupported D carry-forward | At origin: `ERR-HIER-MERGED-D`; invalidate D classification; dependent A/B/C context is preserved only when independently verified |
| Subsequent dependent row inside invalid merge span | Primary `ERR-HIER-CONTEXT-BLOCKED` if it needs damaged context; do not repeat the origin error as primary |
| After invalid merge span ends | Damaged levels remain `UNTRUSTED` until valid explicit restart or a true blank separator |
| Valid explicit restart | New group instance IDs for restarted levels; no revival of closed instances |
| True raw blank separator outside valid vertical merges | Close active instances and set A/B/C `ABSENT` |
| Raw blank non-anchor inside valid vertical merge | **Not** a separator; resolve using the permitted anchor before classifying the row |

A vertical merge spanning a confirmed separator, sentinel, or incompatible parent boundary is a merge conflict; do not inherit across it. Independently provable overlapping merge errors are secondary diagnostics, while the origin primary is selected by earliest affected coordinate (row-major), then fixed code order: horizontal merge, merge conflict, merged D.

### 19.4 Updated HIER-40 expected classification

Fixture:
- Row 3: B=`OMS`, C=`Total`, E=100
- Row 4: C=`Scheduled`, D=`Total`, E=20 (inherits verified OMS)
- Row 5: B=`Lazada`, C=`Total`, E=30 (explicit new B instance)
- Row 6: C=`Total`, D=`Total`, E=40 (inherits verified Lazada B)

Expected:
- Row 3 OMS = `EXCLUDED_SUBTOTAL` due to proven descendant row 4.
- Row 4 Scheduled = locally structurally selected.
- Row 5 Lazada = **`BLOCKED`** because row 6 provides relevant uncertain child evidence within its own verified Lazada instance, and no proven descendant exists.
- Row 6 primary = `ERR-HIER-MALFORMED` at D6; affected context becomes `UNTRUSTED`.
- No final workbook. Do not block OMS due to Lazada's error.

### 19.5 Additional acceptance fixtures (requirements, not executed)

Each fixture has E1 as a valid date, E2=`Sales`, a grand-total sentinel after the listed rows, no unspecified merges, and blank unspecified A–D cells. Check raw/resolved hierarchy, instance closure, primary error, affected parent status, and unchanged source/prior outputs.

| Test | Input rows (A,B,C,D; E) | Deterministic expected result |
|---|---|---|
| HIER-36-R7 | `3: Group,Total,Channel,Total;20` | Potential A and C; C lacks verified B; primary `ERR-HIER-MISSING-PARENT`; no multiple-valid-candidates classification; no output |
| HIER-40-R7 | As §19.4 | OMS `EXCLUDED_SUBTOTAL`; Lazada `BLOCKED`; D6 malformed; no output |
| HIER-41 | `3: —,OMS,Total,—;100`; `4: —,Lazada,Total,Total;20` | Row 4 explicitly changes B to Lazada before error; primary malformed at D4; OMS instance closed, **not blocked** by row 4; no output |
| HIER-42 | `3: —,OMS,OnDemand,Gofresh;10`; `4: —,—,OnDemand,Hyper;20`; B3:B4 valid vertical merge | Row 4 is continuation, not separator; same B instance; both Gofresh and Hyper locally selected; no hierarchy error |
| HIER-43 | `3: —,OMS,OnDemand,Gofresh;10` with B3:C3 horizontal merge; `4: —,—,OnDemand,Hyper;20`; `5: —,Lazada,Total,—;30` | Row 3 merge origin error; row 4 context-blocked; row 5 new verified B instance and locally selected Lazada; no output |
| HIER-44 | `3: —,OMS,OnDemand,Gofresh;10`; `4: —,Lazada,OnDemand,Hyper;20` with C3:C4 vertical merge | Row 4 explicit B change conflicts with C merge; primary merge conflict at affected range; B/C dependent context untrusted; no output |
| HIER-45 | `3: —,OMS,Total,—;100`; `4: —,—,Scheduled,Total;20`; `5: —,—,Total,Total;30` | OMS remains `EXCLUDED_SUBTOTAL` because row 4 is a proven descendant; row 5 malformed; no parent fallback; no output |

**Fixture integrity note:** For valid vertical merge B3:B4, only B3 contains the anchor value and B4 is physically blank. For horizontal B3:C3, only B3 is an anchor and C3 is physically blank. Do not place conflicting raw values in non-anchor cells.

### 19.6 Acceptance and documentation maintenance

- Replace the HIER-36, UT-HIER-14B, and HIER-40 expected outcomes in the authoritative acceptance matrix with §19.1/§19.4. Do not count contradictory older expectations as active tests.
- Include HIER-41–45, merge span closure/restart, valid non-anchor continuation, and explicit-parent-change association.
- Verify that each primary error is counted once; `BLOCKED` classifications and dependent blocked checks are tracked separately.
- Maintain source-to-daily, daily-to-monthly, and serialized-workbook validation rules from earlier steps.
- All tests here are specified only; no tests have been executed.

### 19.7 Outstanding blockers and implementation gate

**Still unresolved for the full program:** sentinel/footer scanning and boundary merges (Part 2), exact Monthly layout, numerical range and significant-digit boundary examples, formula variants, runtime/library, resource limits, and end-to-end acceptance fixtures. A fully verified multiple-valid-candidates fixture remains to be specified.

**IMPLEMENTATION NOT APPROVED.** Do not generate source code, install dependencies, modify source transaction data, commit, or push. Request explicit user approval only after the remaining blockers are resolved.


---

## Step 20 — Revision 8: Merge Fixtures, Total Recovery and Defensive Invariant

**Status: CONFIRMED REQUIREMENTS; IMPLEMENTATION NOT APPROVED.**

**Precedence:** Step 20 supersedes conflicting expectations in Steps 19, 18, 17, 16, 15, 14, 13, and 3. Earlier test cases named below are **inactive** where they conflict with these explicit replacements. This section is a specification only; no tests have been executed.

### 20.1 Candidate discovery and error precedence

1. Discover potential candidates from the permitted raw/resolved A–D hierarchy, then validate required parents, then perform final channel selection after descendant classification.
2. Under the currently approved required-parent rules, multiple structurally valid candidates in one row are not known to be reachable. Keep `ERR-HIER-MULTI-CANDIDATE` as a **defensive invariant**, not a mandatory positive fixture. Do not invent an invalid fixture to force this error.
3. If a provably malformed immediate predecessor exists, `ERR-HIER-MALFORMED` is primary. Otherwise a candidate without a verified required parent produces `ERR-HIER-MISSING-PARENT` (or `ERR-HIER-CONTEXT-BLOCKED` if the required parent is UNTRUSTED due to an earlier error).
4. Independently established additional errors may be included as secondary diagnostics, but a missing parent is not by itself proof of a malformed predecessor.
5. No source sales row may be mapped to more than one final channel.

### 20.2 Missing-parent caused by a Total control marker

For raw `A=Group, B=Total, C=Channel, D=Total`, with no preceding context:
- Potential Rule 1 candidates are `Group` at A and `Channel` at C.
- `Channel` is structurally invalid: B is a control marker, not a verified parent. Primary: `ERR-HIER-MISSING-PARENT` for the C candidate. The potential A candidate must not be published from this invalid row.
- Evaluate marker reset **after** current-row classification. The valid B=`Total` control marker closes B/C instances and sets B/C to `ABSENT`; verified A=`Group` remains. D=`Total` does not restore C or create a channel.
- A subsequent row requiring B without an explicit valid B reports `ERR-HIER-MISSING-PARENT`, not `ERR-HIER-CONTEXT-BLOCKED`, unless a separate error has rendered B `UNTRUSTED`.
- A subsequent explicit valid B can restart with a new instance ID. A critical error anywhere in the run prevents final publication.
- This replaces UT-HIER-14B, UT-HIER-14B-R3, HIER-36, and HIER-36-R7 expectations wherever contradictory.

### 20.3 Merge origin, state transitions and restart

**Valid merge handling:** For vertical merges, only the top-left anchor contains a raw value. Non-anchor raw blanks inside an approved vertical merged range are resolved from that anchor and are **not** blank separators merely because raw A–D and E are blank. Never carry-forward D. Horizontal merges across hierarchy levels are critical errors.

**Horizontal merge crossing hierarchy levels:** The earliest affected row is the error origin (`ERR-HIER-HORIZONTAL-MERGE`). Close all active A/B/C instances and mark A/B/C `UNTRUSTED`. Rows depending on these levels report `ERR-HIER-CONTEXT-BLOCKED` until a valid explicit restart or true separator.

**Vertical merge conflict with explicit higher-level parent change:** Report `ERR-HIER-MERGE-CONFLICT` at the row where the incompatible explicit change is encountered, not automatically at the anchor row. If B changes while C is inherited from a vertical merge under old B, close old B/C instances and mark **B/C `UNTRUSTED`** for this row and dependent subsequent rows; independently verified A remains `VERIFIED` (or `ABSENT` if none exists). The row containing the earlier valid anchor is not retroactively marked erroneous; the run still fails. If conflict is at A, mark A/B/C `UNTRUSTED`. If conflict is at C under unchanged verified B, mark C `UNTRUSTED` and retain verified A/B.

**After invalid span:** Damaged levels remain `UNTRUSTED` after the merged span ends until a valid explicit restart at the necessary level(s) or a confirmed blank separator sets them `ABSENT`. New explicit labels create new instance IDs; closed instances never resume. A valid vertical merge must not cross a confirmed separator, sentinel or incompatible parent boundary.

**Error precedence:** Proven structural/merge error > proven malformed immediate predecessor > structurally valid multiple-candidate invariant > context-blocked check > missing parent > unassigned sales. For an already untrusted predecessor, do not misclassify uncertainty as proven malformed. Merge-origin errors are counted once per origin; dependent errors are reported separately. For overlapping independently provable merge origins, choose earliest affected coordinate (row-major), then horizontal-merge, merge-conflict, merged-D code order.

### 20.4 Corrected authoritative fixtures

Unless otherwise stated, fixtures use valid E1 date, E2=`Sales`, no unlisted merges, blank unlisted hierarchy cells, a later valid grand-total sentinel, and synthetic sales values. These are **expected results**, not executed tests.

| Test | Raw rows and merges | Expected result |
|---|---|---|
| **HIER-43-R8** | Row 3: A blank, B=`OMS`, C blank, D=`Gofresh`, E=10; horizontal merge **B3:C3**. Row 4: B blank, C=`OnDemand`, D=`Hyper`, E=20. Row 5: B=`Lazada`, C=`Total`, E=30. | Row 3 primary `ERR-HIER-HORIZONTAL-MERGE`, A/B/C `UNTRUSTED`, old instances closed; row 4 dependent `ERR-HIER-CONTEXT-BLOCKED`; row 5 explicit valid B restarts new B instance and is locally selected `Lazada`; no final workbook. |
| **HIER-44-R8** | Row 3: B=`OMS`, C=`OnDemand`, D=`Gofresh`, E=10; vertical merge **C3:C4**. Row 4: B=`Lazada`, C physically blank (non-anchor), D=`Hyper`, E=20. | Row 3 anchor is provisionally valid and is not retroactively errored. Row 4 explicit B change conflicts with inherited C anchor: primary `ERR-HIER-MERGE-CONFLICT` at row 4; old B/C instances closed, B/C `UNTRUSTED`, A retains independently verified state; no final workbook. |
| **HIER-36-R8** | Row 3: A=`Group`, B=`Total`, C=`Channel`, D=`Total`, E=20. Row 4: B blank, C=`Scheduled`, D=`Total`, E=30. Row 5: B=`OMS`, C=`Total`, E=40. | Row 3 primary `ERR-HIER-MISSING-PARENT`; after valid B marker reset, A verified, B/C `ABSENT`. Row 4 primary `ERR-HIER-MISSING-PARENT` (not blocked). Row 5 explicit B creates a new verified B instance; `OMS` locally selected. No final workbook because prior errors exist. |
| **HIER-46** | Same rows as HIER-36-R8. | Assert context after row 3: A `VERIFIED`, B/C `ABSENT`; after row 4 B remains `ABSENT`; row 5 valid explicit restart with new B instance ID. |
| **HIER-47** | Row 3: B=`OMS`, C=`OnDemand`, D=`Gofresh`, E=10. Row 4: A–D and E raw blank; vertical merge **B3:B4** and **C3:C4**. Row 5: B=`Lazada`, C=`Total`, E=30. | Row 4 is a valid non-anchor continuation, **not** separator; B/C instance IDs unchanged; no candidate and no sales; row 5 explicit B change creates new B instance. No hierarchy error from row 4. |
| **HIER-48** | Same vertical C3:C4 merge and rows as HIER-44-R8, followed by row 5 with B blank, C=`Scheduled`, D=`Total`, E=30. | Row 4 is conflict origin; row 5 primary `ERR-HIER-CONTEXT-BLOCKED` if B remains untrusted; no retroactive error on row 3; no final workbook. |
| **HIER-49** | Same rows 3–4 as HIER-44-R8; row 5: B=`Shopee`, C=`Total`, E=30. | Row 5 explicit valid B restart creates a **new** verified B instance and locally selected Shopee; earlier conflict still blocks final publication. |
| **HIER-50** | Property-level defensive check on a synthetic candidate collection containing two *structurally valid* candidates for one row (not claimed to be reachable from valid A–D input). | Validator reports `ERR-HIER-MULTI-CANDIDATE`. No positive source-workbook fixture required unless a reachable hierarchy is demonstrated. |

**HIER-40-R7 remains authoritative:** OMS is `EXCLUDED_SUBTOTAL`; Lazada is `BLOCKED`; row 6 malformed; no final output.

### 20.5 Acceptance criteria and outstanding blockers

- Assert raw cells, merge anchors, pre/post context states, closed/new instance IDs, provisional and final classifications, primary and secondary errors, and no final output for full-run failures.
- Replace active expectations for HIER-43/44/36 in the Step 9 acceptance matrix with the **R8** cases above; retain superseded fixtures for historical reference only.
- A blank non-anchor in a valid merge is not a separator; a true raw blank separator outside merges closes active instances.
- Root errors and `BLOCKED` parent diagnostics must be counted separately. Previously generated outputs and source files remain unchanged.
- Remaining blockers for the **full program**: sentinel/footer scan scope and boundary merges; exact Monthly layout; numeric precision/range boundaries; formula variants; runtime/library; resource limits; and end-to-end acceptance fixtures.

**IMPLEMENTATION NOT APPROVED.** Do not generate or modify source code, install dependencies, commit, or push without separate explicit approval.
