# RMStoYELT Roadmap

Branch:

```text
RMStoYELT
```

Workbook:

```text
excel/templates/Template_RMStoYELT_Tool.xlsm
```

Worksheet:

```text
RMStoYELT
```

Tool marker:

```text
A1 = <<RMStoYELT>>
C4 = <<RMStoYELT>>
```

## Core Job

`RMStoYELT` converts an RMS-style event loss table into a YELT-style simulated
event-year loss table.

Reference RMS input:

```text
/Users/michaelzhang/Documents/Map/XL Pricing Templates/Input/Net_Excl_IndRes_withFB_WS_V21.csv
```

Reference YELT output:

```text
/Users/michaelzhang/Documents/Map/XL Pricing Templates/Input/N1_EQ (30-05-2025, 13-11).csv
```

The conversion should:

1. Read a raw RMS-style input ELT file.
2. Parse event-level frequency and severity fields.
3. Use Monte Carlo simulation to generate event occurrences by simulation year.
4. Sample event losses for each occurrence.
5. Write a YELT-style output CSV that can be consumed by other pricing actions.

## Input Shape

The reference RMS input contains metadata/comment rows followed by six-column
numeric event rows.

Metadata rows begin with:

```text
//
```

Observed metadata includes:

```text
SERVER
RDM
ANALYSIS
ID
NAME
DATE
CURRENCY
PERSPCODE
# RECORDS
AAL
1 in 100
1 in 200
1 in 250
```

Observed event row schema:

```text
Event ID
Event Rate
Mean Loss
Std Dev Ind
Std Dev Corr
Exposure
```

Reference sample observations:

```text
Event rows: 18,261
Sum Event Rate: 9.11168969707
Sum Event Rate * Mean Loss: 439,688,028.40462816
Metadata AAL: 439,688,027.935597
```

The close AAL reconciliation confirms that the RMS event rate and mean loss
columns define the annual expected loss basis.

## Output Shape

The V1 reference YELT output had this header:

```text
SIMULATIONNUMBER,PHYSICALEVENTID,LOSS
```

Observed output behavior:

```text
Rows: 78,902
Simulation years represented: 9,997
Minimum simulation number: 1
Maximum simulation number: 10,000
Minimum events in represented simulation year: 1
Maximum events in represented simulation year: 22
Average events in represented simulation year: 7.892567770331099
Simulation years with multiple event occurrences: 9,967
```

`SIMULATIONNUMBER` is the simulated year/trial identifier.

`PHYSICALEVENTID` maps back to the RMS source event identifier. Multiple rows
with the same `SIMULATIONNUMBER` represent multiple event occurrences in that
simulation year.

`LOSS` is the simulated event occurrence loss.

The V2 output contract is the CTR/TS-style event file used by XL Simulation so
the generated file can be recycled directly into the pricing workflow.

Reference CTR/TS event file:

```text
/Users/michaelzhang/Documents/Map/XL Pricing Templates/Input/Net_withFB_Excl_IndRes_WS.csv
```

V2 output shape:

```text
// escape-delay,,,,,
Year,Loss,catastrophic,0,Delay,EventID
```

The data rows do not include a header row. `Delay` is a generated uniform random
draw in `[0, 1]`, matching the reference file's fifth column convention.

## Proposed Toolkit Action

This action belongs to the existing `xl_pricing_tool` scheme.

Action names:

```text
rmstoyelt_gather
rmstoyelt_update
```

`gather` should discover or validate source files and refresh the workbook's
loss cause declarations.

`update` should perform the RMS-to-YELT conversion for active rows.

Current implementation status:

- `rmstoyelt_gather`, `rmstoyelt_update`, and reserved `rmstoyelt_build` are
  registered under `xl_pricing_tool`.
- `<<RMStoYELT>>`, `<<RMS to YELT>>`, and `<<RMS_ToYELT>>` route to the
  RMStoYELT action family through the VBA client.
- VBA browse helpers are fixed-button compatible:
  `BTK_BrowseRMStoYELTFolder` writes the selected folder to `C14` or `D14`,
  and `BTK_BrowseRMStoYELTSourceFile` writes selected CSV paths to
  `D27:D114`.
- `gather` creates `rmstoyelt/rmstoyelt_declarations.csv`; the VBA refresh
  overwrites only `B27:G114`.
- `update` writes one CTR/TS-style event CSV per active declaration row and
  protects against overwriting the source file.

## Workbook Inputs

The workbook has the following intended user inputs:

```text
C6      Simulation Years
C14     Bulk input folder
D14     Bulk output folder
B27:G114 Loss cause declarations
```

Current template verification:

```text
A1  = <<RMStoYELT>> via formula =$C$4
C4  = <<RMStoYELT>>
B13 = Bulk import
B14 = Folder Path:
B25 = Loss Cause Declarations
B26:G26 = CauseID | Peril | FilePath | Active | OutDir | OutFileName
E27:E114 = Active default formulas, returning Y when CauseID is populated
```

`C6`, `C14`, and `D14` are available as worksheet cells but should be labeled
and documented in the template before final release.

## User Input Modes

V1 should support two input modes.

### Bulk Input

Bulk input is folder-driven.

User workflow:

1. Select `C14`.
2. Click `Browse Folder`.
3. VBA writes the selected input folder path to `C14`.
4. Select `D14`.
5. Click `Browse Folder`.
6. VBA writes the selected output folder path to `D14`.
7. Click `Gather`.

Gather behavior:

- If `C14` is blank, gather must not update the declaration table.
- If `C14` is populated, gather scans the input folder for source files.
- Gather reads the folder contents and refreshes declarations in `B27:G...`.
- Each discovered file should produce one declaration row.
- `FilePath` should be the absolute source file path.
- `OutDir` should default from `D14` when `D14` is not blank.
- `OutFileName` should default to the source file base name, without directory,
  adjusted to the desired YELT output naming convention.
- `Active` should default to `Y` for populated rows, while remaining editable by
  the user.

### Single File Input

Single file input is row-driven.

User workflow:

1. Select one or more cells in `D27:D114`.
2. Click `Browse Source`.
3. VBA writes selected source file path(s) to the selected `FilePath` cell(s).
4. User reviews or edits `CauseID`, `Peril`, `Active`, `OutDir`, and
   `OutFileName`.

Single-file behavior:

- The browse action should only write to selected cells in `D27:D114`.
- It should not overwrite unrelated blocks outside `B:G`.
- `Active` is controlled by user-editable `Y` / `N` identifiers.
- `OutDir` can be manually set by the user.
- If `D14` is not blank, `OutDir` can be automatically populated from `D14`.
- `OutFileName` should default to the selected source file name without its
  directory, but should remain user-editable.

## Gather Contract

`rmstoyelt_gather` should:

- Read bulk input folder from `C14`.
- Read bulk output folder from `D14`.
- Do nothing to declarations when `C14` is blank.
- When `C14` is populated, scan that folder for candidate RMS ELT files.
- Populate declaration rows starting at `B27`.
- Return a clear summary: folder scanned, files found, rows populated, and
  whether `D14` was used for output folders.

Gather should not run the Monte Carlo simulation and should not create YELT
output files.

## Update Contract

`rmstoyelt_update` should:

- Read `Simulation Years` from `C6`.
- Read active declaration rows from `B27:G114`.
- Treat only rows with `Active = Y` as conversion jobs.
- For each active row, read the input file from `FilePath`.
- Parse the RMS ELT from that input file.
- Run the RMStoYELT Monte Carlo simulation.
- Create one YELT output file for that declaration row.
- Use `OutDir` as the output folder.
- Use `OutFileName` as the declared output file name.
- Save the generated file under:

```text
<OutDir>/<OutFileName>
```

- Write output as a CTR/TS-style event file:

```text
// escape-delay,,,,,
Year,Loss,catastrophic,0,Delay,EventID
```

- Write one output row per simulated event occurrence.
- Represent multiple event occurrences in the same simulated year by repeated
  `Year` values.
- Return a summary of active jobs, input files read, output files written, event
  rows generated, and any skipped or failed rows.

Update should not infer new files from `C14`; it should operate from the
declaration table as the authoritative job list. `C14` and `D14` are gather /
convenience inputs, while `FilePath`, `OutDir`, and `OutFileName` on each row
are the update contract.

## Methodology Baseline

The RMS ELT cannot be converted directly into a YELT as a deterministic file
transform. A YELT must be generated through Monte Carlo simulation using the ELT
event rates and loss uncertainty parameters.

Core process:

1. Start with RMS ELT fields:

```text
EventID
Rate
Mean
Sdi
Sdc
Exposure
```

2. Calculate total annual event frequency:

```text
lambda = sum(Rate)
```

3. For each simulated year, draw the number of event occurrences:

```text
N ~ Poisson(lambda)
```

4. For each occurrence, sample an event from the ELT with probability
   proportional to event rate:

```text
P(event_i) = Rate_i / lambda
```

5. For the selected event, simulate loss severity using mean loss and
   uncertainty terms. A common approximation is a scaled Beta distribution
   parameterized by:

```text
Mean
Sdi
Sdc
Exposure
```

6. Write each simulated occurrence as a YELT row, for example:

```text
SIMULATIONNUMBER, PHYSICALEVENTID, LOSS
```

Each occurrence gets its own row. Event occurrence within a simulation year is
captured by repeated `SIMULATIONNUMBER` values, one row per occurrence.

7. If a strict YLT is needed instead of a YELT, aggregate simulated event losses
   by trial or year.

Important caveat: preserving correlation across coverages, locations, or
sub-portfolios requires careful treatment of `Sdi` and `Sdc`. A simple
independent event simulation may be acceptable for a rough prototype, but it may
not reproduce RMS platform behavior exactly.

## Methodology Discussion

The core Monte Carlo method should be explicit and reproducible.

Recommended first-pass model:

1. Treat each RMS row as an event with annual rate `lambda_i = Event Rate`.
2. Let total annual event frequency be `lambda = sum(lambda_i)`.
3. For each simulation year, draw event count:

```text
N_year ~ Poisson(lambda)
```

4. For each simulated occurrence, sample one source event using event rates as
   weights:

```text
P(event = i) = lambda_i / sum(lambda_i)
```

5. For the selected event, sample loss from a severity distribution calibrated
   to source `Mean Loss` and uncertainty terms.

The implemented RMStoYELT V1 severity method uses a scaled Beta approximation
parameterized by `Mean`, `Sdi + Sdc`, and `Exposure`:

```text
s = Sdi + Sdc
a = (Mean / s)^2 * (1 - Mean / Exposure) - (Mean / Exposure)
b = a * (Exposure / Mean - 1)
Loss = qbeta(U, a, b) * Exposure
```

When uncertainty is zero, invalid, or not compatible with the Beta moment
conditions, loss defaults conservatively to `Mean`, capped at `Exposure`.

6. Write one YELT row per simulated occurrence:

```text
SIMULATIONNUMBER = simulation year / trial id
PHYSICALEVENTID = selected RMS Event ID
LOSS = simulated event occurrence loss
```

The design captures event occurrence within the simulation year because a year
with `N_year > 1` produces `N_year` rows with the same `SIMULATIONNUMBER`.

## Open Questions

1. Should the random seed be user-controlled from the workbook for repeatable
   runs?
2. Should gather eventually validate discovered CSVs, or is fast folder listing
   preferred for V1?
3. Should output file names default to the source file name exactly, or should
   browse/gather add a `_YELT` suffix to avoid accidental source/output name
   collisions when users choose the same folder?
