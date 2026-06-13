# Aggregates Input Roadmap

Branch: `aggregates_input`

Status: implemented in R/VBA source; pending Windows Excel ribbon validation.

## Goal

Implement the `Input_Agg` tab update action for the XL pricing tool.

The action will normalize aggregate exposure input data into a method-level R
output table and overwrite the existing `Input_Agg!AH:AK` output block. Existing
Excel formulas in the `U:Y` summary section remain workbook-owned and should not
be replaced by R.

## Action Contract

```text
Sheet marker:        <<Agg>>
Ribbon action:       Update
Dispatched action:   agg_update
R function:          xl_pricing_tool_agg_update
CSV handoff:         _BERTToolkitTemp/xl_pricing_tool/agg/agg_output.csv
Excel target:        Input_Agg!AH18:AK...
```

The CSV handoff schema is:

```text
Peril,TreatyYear,Method,AggSum
```

Supported perils:

```text
EQ
WF
```

Supported methods:

```text
Nationwide
AIR
RMS
Blended
```

## Inputs

### Aggregate Input Table

Source range starts at `Input_Agg!B12`.

Expected shape:

```text
Peril | Province | <year columns>
```

Rules:

- Read year columns from row `12`; do not hardcode the first or last year.
- Treat the data range as flexible length.
- Only `EQ` and `WF` are valid perils.
- Province names should match the keyzone table by trimmed text.

### Keyzone Table

Source range starts at `Control_Module!K15`.

Expected shape:

```text
Province | Peril | AIR | RMS | Blended
```

Rules:

- `1` means include the province/peril row in the method.
- blank or `0` means exclude it.
- `Blended` uses only the explicit `Blended` flag.
- `Blended` is not derived from AIR/RMS flags.

## Processing Logic

1. Validate required sheets: `Input_Agg` and `Control_Module`.
2. Read the aggregate input table from `Input_Agg!B12`.
3. Convert the wide year columns into long form:

```text
Peril | Province | TreatyYear | Agg
```

4. Filter to supported perils: `EQ`, `WF`.
5. Read keyzone flags from `Control_Module!K15:O...`.
6. Calculate method summaries:
   - `Nationwide`: sum all provinces by peril/year.
   - `AIR`: sum only province/peril rows where AIR flag is `1`.
   - `RMS`: sum only province/peril rows where RMS flag is `1`.
   - `Blended`: sum only province/peril rows where Blended flag is `1`.
7. Write `agg_output.csv`.
8. Return a text result that includes `Output folder: ...`.
9. VBA imports `agg_output.csv`, clears the old `AH:AK` output area, and writes
   the new output beginning at `Input_Agg!AH18`.

## Excel Post-Processing

VBA should own workbook mutation for this action.

VBA function:

```text
BTK_RefreshAggOutput(outputFolder)
```

Responsibilities:

- Locate `_BERTToolkitTemp/xl_pricing_tool/agg/agg_output.csv`.
- Clear the previous `Input_Agg!AH18:AK...` output block.
- Preserve the header and note rows above the output block.
- Import CSV into `Input_Agg!AH18`.
- Save the workbook.
- Leave formulas in `U:Y` unchanged so Excel recalculates summary prior/current/change.

## Validation Plan

R-side validation:

- Reject missing `Input_Agg` or `Control_Module`.
- Reject unsupported perils outside `EQ`/`WF`.
- Reject aggregate rows with blank peril or province.
- Reject nonnumeric year headers.
- Warn or fail on aggregate province/peril combinations missing from keyzone when
  method-specific flags are needed.

Output validation:

- Confirm output columns are exactly:

```text
Peril,TreatyYear,Method,AggSum
```

- Confirm each peril/year has one row for each method.
- Confirm `Nationwide` equals the total of all province rows for that peril/year.
- Confirm `AIR`, `RMS`, and `Blended` are based only on explicit keyzone flags.
- Confirm no charts are created.

Excel validation:

- Run `Input_Agg > GRe Tool Ribbon > Update`.
- Confirm `AH18:AK...` is overwritten.
- Confirm `U:Y` formulas remain present.
- Confirm summary values recalculate from the refreshed `AH:AK` table.

## Implementation Milestones

1. Done: register `agg_gather`, `agg_update`, and `agg_build` in `r/registry.R`.
2. Done: extend VBA action mapping so `<<Agg>>` maps to `xl_pricing_tool` and
   `agg_update`.
3. Done: add `r/tools/xl_pricing_tool/agg.R` with read, validate, transform,
   summarize, and CSV write functions.
4. Done: source `agg.R` from `r/registry.R`.
5. Done: add `BTK_RefreshAggOutput` in `excel/vba/BERTToolkitClient.bas`.
6. Done: run R smoke tests using the current template.
7. Pending: validate in the Windows VM through the ribbon.
8. Pending: commit workbook-side changes only if a workbook binary edit is made
   and validated in Excel.

## Current R Validation Result

The R dispatcher path has been tested with:

```text
BTK.DispatchTool("xl_pricing_tool", "agg_update", <template>, <output_dir>)
```

Observed output:

```text
Input rows: 992
Output rows: 128
```

The output row count matches:

```text
2 perils x 16 treaty years x 4 methods = 128 rows
```

Sample reconciliation for `EQ` / `2025`:

```text
Nationwide = 592,830,404,568
AIR        = 151,122,249,306
RMS        = 194,921,464,664
Blended    = 281,989,877,521
```

These values reconcile to the raw aggregate input and explicit keyzone flags.

## Safety Rules

- Do not edit the `.xlsm` package XML directly.
- Do not commit workbook binary changes unless the workbook was edited and
  validated in Excel.
- Use CSV as the R-to-VBA handoff format for reliability.
- Keep `.rds` optional only for debugging, not as the Excel handoff.
