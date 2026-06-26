# XL Simulation V2 Source Input Roadmap

Branch:

```text
xl_simulation_v2
```

Workbook:

```text
excel/templates/Template_XOL_Pricing_Tool.xlsm
```

Worksheet:

```text
Sim_Variations
```

## Goal

V2 should let users point each active loss cause to an external source file and
have `xlsimulation_update` read that file before falling back to the pasted
template source block.

The workbook should become the control panel for source selection, layer
configuration, and output review. Large model or CDF source data should not have
to live inside the workbook unless the user chooses the legacy paste workflow.

## Desired Outcome

For active loss causes, the update action should resolve source data in this
order:

1. Use the user-selected external source file when configured.
2. If the external file cannot be used and fallback is allowed, read the pasted
   source block from the template.
3. If neither source is usable, stop with a clear cause-specific validation
   error.

The result message should report which source was used for each active cause and
should make fallbacks visible. Silent fallback is not acceptable because it can
hide stale workbook data.

## User Workflow

The loss cause declaration table already has the required control columns:

```text
CauseID | CauseFamily | ModelSource | Peril | InputMode | FilePath | Active | Notes | Location
```

V2 should make `InputMode` and `FilePath` meaningful.

Supported `InputMode` values:

```text
Paste
File
Auto
```

Mode behavior:

```text
Paste  Read only the workbook source block declared in Location.
File   Read only the external file declared in FilePath; do not fallback.
Auto   Try FilePath first, then fallback to Location if the file is missing,
       unreadable, or fails schema validation.
```

Recommended default:

```text
Auto
```

for externally sourced `ModeledCAT` and `CDF` causes.

## Browse Controls

Add workbook file browse support for:

```text
CauseFamily = ModeledCAT
CauseFamily = CDF
```

The browse action should:

- open a file picker;
- write the selected path into the selected cause row's `FilePath` cell;
- leave `InputMode` unchanged because source mode is user-controlled;
- avoid copying file contents into the workbook;
- leave `Location` unchanged as the fallback source block.

Use one fixed worksheet browse button for the currently selected loss cause row.
The button should only run when the active selection is one `FilePath` cell in
`G50:G61`. It should read only the matching declaration row `B:J`, validate the
cause, and write the selected absolute path back to that row's `FilePath` cell.
The button can be a normal worksheet shape assigned to
`BTK_BrowseXLSimulationSourceFile`; it does not need custom ribbon XML.

`FS` causes should remain paste-only in the first V2 pass unless a real external
file workflow is confirmed.

## Source Resolution Architecture

The current V1 implementation exports active workbook source blocks to:

```text
<workbook folder>/_BERTToolkitTemp/xl_pricing_tool/xlsimulation/source_inputs/
```

and writes:

```text
xlsimulation_source_manifest.csv
<CauseID>.csv
```

V2 should extend this manifest rather than introduce a second handoff channel.

Recommended manifest columns:

```text
CauseID
CauseFamily
ModelSource
InputMode
FilePath
Location
CsvPath
Rows
Cols
SourceKind
FallbackAllowed
ExportedBy
Status
Message
```

`SourceKind` values:

```text
file
template
```

VBA should only export large worksheet source blocks when needed. If an active
cause has `InputMode = File`, or `InputMode = Auto` with a non-empty `FilePath`,
the manifest can point to the external file first and should not eagerly export
the workbook block unless fallback is needed.

R should own final source resolution and validation because it already owns the
family/model schema rules.

## External File Formats

V2 should support both clean template-shaped CSV files and selected raw source
formats that users actually receive.

### ModeledCAT RMS

Support clean six-column CSV files with headers:

```text
Event ID | Event Rate | Mean Loss | Std Dev Ind | Std Dev Corr | Exposure
```

Also support raw RMS-style CSV exports where:

- metadata/comment rows appear before the data;
- data rows have six comma-separated fields;
- the source file may not include a header row.

The parser should skip leading metadata/comment/blank rows and infer the RMS
six-column schema when the active cause has:

```text
CauseFamily = ModeledCAT
ModelSource = RMS
```

### ModeledCAT CTR and TS

Support clean three-column CSV files with headers:

```text
EventID | Year | Loss
```

Raw source formats that do not directly contain this schema should not be
guessed in the first pass unless a reliable mapping is approved.

### CDF

Support clean CDF parameter/value CSV files, including files shaped like:

```text
Parameter | Value
Mean Frequency | ...
Frequency Type | ...
Interpolation | ...
Loss Cap | ...
Minimum Loss | ...
Percentile | Loss Severity
...
```

The existing template-style CDF blocks should remain valid fallback sources.

## Fallback Rules

Recommended behavior:

```text
Paste
  FilePath is ignored.
  Location is required.

File
  FilePath is required.
  External file must exist, parse, and validate.
  No fallback to Location.

Auto
  If FilePath is blank, use Location.
  If FilePath is present, try external file first.
  If external read or validation fails, use Location.
  Emit a prominent warning in the update result.
```

Fallback should apply to:

- missing file;
- unreadable file;
- parse failure;
- schema validation failure.

Schema validation fallback is allowed in `Auto`, but the warning should include
the original file error so the user can fix the source.

## Validation

Validation should remain tab/action-specific. `xlsimulation_update` should
validate only active loss causes and should validate the resolved source data for
each active cause.

Suggested checks:

- `InputMode` is one of `Paste`, `File`, or `Auto`;
- `FilePath` exists when `InputMode = File`;
- `Location` exists when fallback is required or `InputMode = Paste`;
- external files parse into the expected family/model schema;
- fallback source blocks parse into the expected family/model schema;
- result summary includes one row per active cause with source status.

## Output and User Feedback

The update action should produce a source resolution summary, either in the
normal action message or as a handoff CSV that VBA can surface.

Suggested fields:

```text
CauseID
InputMode
ResolvedSource
ResolvedPathOrRange
FallbackUsed
Status
Message
```

Example messages:

```text
RMS__EQ: used external file C:\...\Net_Incl_IndRes_withFB_EQ_V21.csv.
CDF_MB: external file failed schema validation; used template range AV134:AW1000.
TS_WF: File mode selected but FilePath is blank.
```

## Implementation Phases

### Phase 1: Contract and Workbook Controls

- Document the V2 source resolution contract.
- Add file browse VBA support for selected loss cause rows.
- Store selected paths in `FilePath`.
- Set or preserve `InputMode` according to the user action.
- Do not copy external file contents into the workbook.

### Phase 2: Manifest Extension

- Extend the source manifest with `InputMode`, `FilePath`, `SourceKind`,
  `FallbackAllowed`, `Status`, and `Message`.
- Avoid exporting workbook source blocks when an external file is configured and
  fallback is not immediately needed.
- Preserve V1 manifest compatibility where possible for smoke tests.

### Phase 3: R Source Resolution

- Add a source resolver that tries external files before template sources.
- Keep existing template-block parsing as the fallback path.
- Add external parsers for RMS raw CSV, clean RMS CSV, clean CTR/TS CSV, and CDF
  parameter/value CSV.
- Return structured source resolution statuses for VBA/user messaging.

### Phase 4: Tests and Examples

- Add tests or smoke scripts using the sample source files under:

```text
/Users/michaelzhang/Documents/Map/XL Pricing Templates/Input
```

- Verify RMS raw files skip metadata rows and infer six data columns.
- Verify CDF files read without requiring manual paste.
- Verify `Auto` fallback reports warnings.
- Verify `File` mode errors instead of falling back.

## Decisions

- Browse should use one generic button for the selected loss cause row.
- `InputMode` is set by the user; browse should not change it automatically.
- Raw CTR/TS files that are not already shaped as `EventID | Year | Loss` are
  out of scope for the first V2 pass.
- `FilePath` should store absolute paths.
- Source resolution warnings should appear only in the action result message.
