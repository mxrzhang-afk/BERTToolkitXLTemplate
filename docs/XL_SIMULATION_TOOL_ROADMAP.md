# XL Simulation Tool Roadmap

Branch:

```text
xl_simulation_tool
```

Workbook:

```text
excel/templates/Template_XOL_Pricing_Tool.xlsm
```

Worksheet:

```text
Sim_Variations
```

Tool marker:

```text
A1 = <<XLSimulation>>
C5 = <<XLSimulation>>
```

This document records the V1 contract for the XL simulation tab before the R
and VBA implementation work starts. Keep it updated when the worksheet layout,
simulation assumptions, or handoff files change.

## V1 Scope

The XL simulation tab will support tab-level actions under the existing
`xl_pricing_tool` registry:

```text
xlsimulation_gather
xlsimulation_update
xlsimulation_build
```

V1 priorities:

- Keep input reads strict and fast.
- Read only declared rectangular ranges.
- Validate active loss causes only.
- Avoid broad worksheet scans and relaxed address guessing.
- Write CSV handoff files for VBA to load into fixed worksheet output ranges.

`build` is reserved for a later pricing handoff/export step unless a specific
deliverable is approved.

## Sheet Address Map

Simulation settings:

```text
B13:D18
B14:C14 Simulation Years
B15:C15 Random Seed
B16:C16 Return Period
B17:C17 Risk Measure
B18:C18 CAT Scaling Factor
```

Layer input:

```text
B26:H41
B26:H26 headers:
LayerID | LayerName | Limit | Deductible | # of Reinstatements | ReinstatementPct | InputROL
```

Layer output:

```text
J26:AB41
```

Layer output headers:

```text
J  LayerID
K  LayerName
L  IncludedCauses
M  Limit
N  Deductible
O  Reinstatements
P  InputROL
Q  InputPremium
R  ReinstatementFactor
S  ExpectedLoss_AEP
T  StdDev_AEP
U  SelectedVaR_OEP
V  SelectedTVaR_OEP
W  SelectedRiskMetric_OEP
X  CapitalRequired
Y  LossCostROL
Z  InputLossRatio
AA InputSDMultiple
AB InputROE
```

Loss cause declarations:

```text
B49:J61
B49:J49 headers:
CauseID | CauseFamily | ModelSource | Peril | InputMode | FilePath | Active | Notes | Location
```

Layer/cause coverage matrix:

```text
M49:V61
```

Cause breakdown output:

```text
AE26:AO41
```

The gather action should refresh `AE26:AO26` so `AE26` is `LayerID` and the
remaining headers are the active loss causes in declaration order.

## Loss Cause Families

V1 supports up to 12 declarable loss causes across three families. `CauseID`
values are user-facing names and may be changed in the workbook. Simulation
behavior is controlled by `CauseFamily` and, for Modeled CAT rows,
`ModelSource`.

Modeled CAT rows:

```text
CauseFamily = ModeledCAT
ModelSource = CTR | TS | RMS
Maximum active rows = 6
```

Frequency-severity parameter rows:

```text
CauseFamily = FS
ModelSource = blank
Maximum active rows = 3
```

Severity CDF rows:

```text
CauseFamily = CDF
ModelSource = blank
Maximum active rows = 3
```

The declaration row order controls the fixed OEP output block assignment.

## Source Block Configuration

All source block addresses come from the `Location` column of the loss cause
declarations table. V1 should parse only explicit rectangular ranges:

```text
<start_col><start_row>:<end_col><end_row>
```

Examples:

```text
B134:D100000
R134:W100000
AF134:AG143
AP134:AQ1000
```

Malformed ranges should raise a clear error when the corresponding cause is
active. Inactive causes should not block gather or update.

Current configured source blocks:

```text
CTR_EQ        B134:D100000    EventID | Year | Loss
CTR_WF        F134:H:100000   inactive; malformed in current template
TS_EQ         J134:L100000    EventID | Year | Loss
TS_WF         N134:P100000    EventID | Year | Loss
RMS__EQ       R134:W100000    Event ID | Event Rate | Mean Loss | Std Dev Ind | Std Dev Corr | Exposure
RMS__WF       Y134:AD100000   Event ID | Event Rate | Mean Loss | Std Dev Ind | Std Dev Corr | Exposure
FS_Property   AF134:AG143     parameter/value table
FS_CEAR       AI134:AJ143     parameter/value table
FS_MB         AL134:AM143     parameter/value table
CDF_Porperty  AP134:AQ1000    percentile/loss severity table
CDF_CEAR      AS134:AT1000    percentile/loss severity table
CDF_MB        AV134:AW1000    percentile/loss severity table
```

## Gather Action

`xlsimulation_gather` should:

- Read layer definitions from `B26:H41`.
- Read active loss cause declarations from `B49:J61`.
- Refresh the layer/cause inclusion matrix at `M49:V61`.
- Preserve existing `Y`/`N` selections where layer/cause intersections still
  exist.
- Add active layers and active causes to the matrix.
- Refresh `AE26:AO26` with `LayerID` plus active loss cause headers.
- Validate active source block schemas for ModeledCAT, FS, and CDF inputs.
- Return a clear validation summary.

Expected summary items:

- active layer count and IDs
- active cause count and IDs
- inactive causes skipped
- source blocks validated
- matrix dimensions refreshed
- any schema or location errors

## Update Action

`xlsimulation_update` should:

- Read simulation settings from `B13:D18`.
- Gather layer information from `B26:H41`.
- Gather coverage selections from `M49:V61`.
- Read active loss cause declarations from `B49:J61`.
- Read each active cause source block from its declared `Location`.
- Simulate event-based losses for each active cause.
- Apply CAT scaling factor to all `ModeledCAT` causes.
- Allocate event losses into selected layers.
- Apply reinstatement assumptions to annual aggregate layer loss.
- Produce layer output metrics in `J26:AB41`.
- Produce expected loss contribution by cause/layer in `AE26:AO41`.
- Produce simulated OEP tables for active causes.
- Leave inactive cause OEP tables as zero.
- Write CSV handoff files for VBA to load into workbook ranges.

## Simulation Settings

`Simulation Years` is read from the worksheet and controls the number of
simulated years/trials. It must not be hardcoded to 10000.

`Random Seed` is optional but should be used for deterministic stochastic runs
when provided.

`Return Period` controls the selected VaR/TVaR metric.

`Risk Measure` supports:

```text
VaR
TVaR
```

`CAT Scaling Factor` applies to all `ModeledCAT` causes.

## Simulation Mechanics

### CTR and TS ELTs

CTR and TS event loss tables are treated as pre-simulated event-year tables:

```text
EventID | Year | Loss
```

V1 behavior:

- Do not simulate additional events from CTR/TS.
- Group events by source `Year`.
- Internally remap selected/reused years to simulation year `1:N`.
- If requested simulation years are fewer than available years, truncate.
- If requested simulation years exceed available years, reuse/recycle years.
- Preserve all events within each selected source year.

### RMS Frequency/Severity CDF Tables

RMS blocks use a frequency plus severity-CDF simulation derived from the event
table:

```text
Event ID | Event Rate | Mean Loss | Std Dev Ind | Std Dev Corr | Exposure
```

V1 behavior:

- Annual frequency follows `Poisson(sum(Event Rate))`.
- Severity is sampled through an inverse-CDF selection over cumulative
  `Event Rate` weights.
- Conditional event severity follows a lognormal distribution matched to the
  selected row's `Mean Loss` and `Std Dev Ind`.
- Ignore `Std Dev Corr`.
- Ignore `Exposure`.
- Simulated RMS rows use generated event identifiers; source `Event ID` is an
  input to the severity mixture, not a pre-simulated event-year identity.

### Frequency-Severity Parameter Tables

FS causes use Poisson frequency and declared parametric severity:

```text
Mean Frequency
Frequency Type
Severity Family
Loss Cap
Minimum Loss
Parameter
Value
```

V1 frequency:

```text
Poisson(Mean Frequency)
```

Supported severity families:

```text
pareto
loggamma
weibull
lognormal
gamma
loglogistic
```

`Minimum Loss` and `Loss Cap` are applied by truncating/capping sampled severity.

### Severity CDF Tables

CDF causes use Poisson frequency and inverse-CDF severity sampling:

```text
Mean Frequency
Frequency Type
Interpolation
Loss Cap
Minimum Loss
Percentile | Loss Severity
```

V1 behavior:

- Frequency follows `Poisson(Mean Frequency)`.
- Severity is sampled by drawing `U ~ Uniform(0, 1)`.
- Severity is the linearly interpolated loss at percentile `U`.
- `Minimum Loss` and `Loss Cap` are applied by truncating/capping sampled
  severity when provided.

## Layer Allocation

For each event loss and selected layer:

```text
event_layer_loss = min(max(event_loss - deductible, 0), limit)
```

Annual aggregate capacity:

```text
annual_layer_capacity = limit * (1 + reinstatement_count * reinstatement_pct)
```

AEP annual loss:

```text
annual_aep_loss = min(sum(event_layer_loss), annual_layer_capacity)
```

OEP annual loss:

```text
annual_oep_loss = max(event_layer_loss)
```

Reinstatement capacity does not separately cap OEP beyond the per-event layer
limit.

## Layer Metrics

Metric formulas:

```text
ReinstatementFactor = 1 + ReinstatementCount * ReinstatementPct
InputPremium = Limit * InputROL * ReinstatementFactor
ExpectedLoss_AEP = mean(annual_aep_loss)
StdDev_AEP = sd(annual_aep_loss)
SelectedVaR_OEP = OEP quantile at Return Period
SelectedTVaR_OEP = mean(OEP losses >= SelectedVaR_OEP)
SelectedRiskMetric_OEP = SelectedVaR_OEP if Risk Measure is VaR
SelectedRiskMetric_OEP = SelectedTVaR_OEP if Risk Measure is TVaR
CapitalRequired = SelectedRiskMetric_OEP - ExpectedLoss_AEP
LossCostROL = ExpectedLoss_AEP / Limit
InputLossRatio = ExpectedLoss_AEP / InputPremium
InputSDMultiple = (InputPremium - ExpectedLoss_AEP) / StdDev_AEP
InputROE = (InputPremium - ExpectedLoss_AEP) / CapitalRequired
```

Cause breakdown values are expected loss contributions only:

```text
mean annual AEP layer loss attributable to that cause
```

## OEP Output Blocks

Each cause has a fixed OEP block:

```text
CTR_EQ        B107:C121
CTR_WF        F107:G121
TS_EQ         J107:K121
TS_WF         N107:O121
RMS__EQ       R107:S121
RMS__WF       Y107:Z121
FS_Property   AF107:AG121
FS_CEAR       AI107:AJ121
FS_MB         AL107:AM121
CDF_Porperty  AP107:AQ121
CDF_CEAR      AS107:AT121
CDF_MB        AV107:AW121
```

Rows:

```text
107 title
108 Return Period | OEP
109:121 return period values
```

Update should refresh active cause OEP values and leave inactive cause values
as zero.

## Proposed Handoff Files

R should write files under:

```text
<workbook folder>/_BERTToolkitTemp/xl_pricing_tool/xlsimulation/
```

Proposed files:

```text
xlsimulation_gather_summary.csv
xlsimulation_layer_output.csv
xlsimulation_layer_breakdown.csv
xlsimulation_oep_output.csv
```

For `xlsimulation_update`, VBA should also export active source blocks before
calling BERT:

```text
<workbook folder>/_BERTToolkitTemp/xl_pricing_tool/xlsimulation/source_inputs/
```

Input handoff files:

```text
xlsimulation_source_manifest.csv
<CauseID>.csv
```

This keeps memory-heavy worksheet reads inside Excel/VBA and lets BERT R operate
on compact CSV inputs. R can still fall back to workbook XML source reads when
the manifest is missing, so local smoke tests can run without Excel.

VBA should load these into:

```text
J26:AB41
AE26:AO41
B107:AW121
```

The exact handoff shape may be adjusted during implementation, but the workbook
output ranges above are the V1 contract.

## Implementation Notes

- Add `xlsimulation.R` under `r/tools/xl_pricing_tool/`.
- Source it from `r/registry.R`.
- Register `xlsimulation_gather`, `xlsimulation_update`, and
  `xlsimulation_build` under `xl_pricing_tool`.
- Add `<<XLSimulation>>` routing to `excel/vba/BERTToolkitClient.bas`.
- Add VBA refresh handling for the xlsimulation CSV handoffs.
- Keep the strict location parser local to this action or shared only if another
  action needs the same contract.
