# XL Pricing Tool

This module is the BERT-ready home for the XL pricing workbook workflow.

Template:

```text
excel/templates/Template_XOL_Pricing_Tool.xlsm
```

Source workbook:

```text
XL Pricing Templates/XL Pricing Template.xlsm
```

Tool ID:

```text
xl_pricing_tool
```

Objective markers accepted by the shared VBA dispatcher:

```text
<xl pricing>
<xl pricing tool>
<xol pricing>
```

Generic ribbon actions:

```text
gather
update
build
```

Tab-specific actions are registered separately under the same tool ID. For
example, `Input_GNPI` maps the ribbon Update command to `gnpi_update`.
The actions are intentionally separate from `china_exposure_map`.

By default, outputs are written to a stable workbook-local folder:

```text
<workbook folder>\_BERTToolkitTemp\xl_pricing_tool\
```

## Tab: Input_GNPI

Marker:

```text
<<GNPI>>
```

Default behavior:

- `gather`: not defined when `Gather From` is blank.
- `update`: summarize GNPI data for the two right-side chart outputs.
- `build`: no standalone output currently defined.

`update` reads the flexible table starting at `B12:I` and stops after the
meaningful input rows end. The right-side user inputs control chart
presentation. `Start_Year` is read from `L18`, or from a label/value setting
named `Start Year` / `Start_Year`. The LOB stack sequence starts at `L39`.

The chart aesthetics panel starts at `AA10`, with labels in column `AA` and
editable values in column `AB`. The R reader also supports the same labels
elsewhere on the sheet as a fallback. Supported labels include:

```text
Level Chart Title
Level Chart Subtitle
Composition Chart Title
Composition Chart Subtitle
Actual Color
Revised Color
Estimate Color
Label Color
Grid Color
PNG Width
PNG Height
Font Scale
Label Font Size
Axis Font Size
Title Font Size
LOB Color 1
LOB Color 2
```

Generated files:

```text
_BERTToolkitTemp/xl_pricing_tool/gnpi/gnpi_level_data.csv
_BERTToolkitTemp/xl_pricing_tool/gnpi/gnpi_lob_comparison_data.csv
_BERTToolkitTemp/xl_pricing_tool/gnpi/gnpi_chart_config.csv
_BERTToolkitTemp/xl_pricing_tool/gnpi/gnpi_level_chart.png
_BERTToolkitTemp/xl_pricing_tool/gnpi/gnpi_lob_comparison_chart.png
```

The first chart shows separate columns for each year/type section, such as
`2025 Estimate`, `2025 Revised`, and `2026 Estimate`. The renewal-year section
keeps the required order, without extra highlight outlines. The second chart is
a stacked percentage LOB composition chart across those same year/type sections
from the configured start year, with stack order controlled by the `New
Sequence` list. Charts are rendered by R with `ggplot2` as PNG images and then
inserted into the workbook by VBA.

## Tab: Input_Agg

Marker:

```text
<<Agg>>
```

Implemented action:

```text
agg_update
```

Default behavior:

- `gather`: not defined when `Gather From` is blank.
- `update`: normalize aggregate inputs into the R Output block.
- `build`: no standalone output currently defined.

`update` will read the flexible aggregate input table beginning at
`Input_Agg!B12`. The expected columns are:

```text
Peril | Province | <year columns>
```

Only `EQ` and `WF` are expected perils for this action. Year columns should be
read from row `12` so the action does not depend on a fixed start or end year.

The keyzone lookup table is read from `Control_Module!K15:O...`:

```text
Province | Peril | AIR | RMS | Blended
```

Flag handling:

- `1` means include the province/peril row in that method.
- blank or `0` means exclude it.
- `Blended` is explicit only; it is not derived from AIR/RMS flags.

The R action will produce a CSV handoff file:

```text
_BERTToolkitTemp/xl_pricing_tool/agg/agg_output.csv
```

The CSV schema is:

```text
Peril | TreatyYear | Method | AggSum
```

Methods:

```text
Nationwide
AIR
RMS
Blended
```

VBA will use the CSV handoff to overwrite `Input_Agg!AH18:AK...`, preserving
the existing Excel formulas in the `U:Y` summary area. The summary formulas are
expected to use:

```text
Prior   = RNL_Year - 2
Current = RNL_Year - 1
```

No charts are planned for this tab.

## Tab: Input_Profile

Marker:

```text
<<Profile>>
```

Implemented first-stage action:

```text
profile_update
```

Default behavior:

- `gather`: not defined when `Gather From` is blank.
- `update`: calculate exposure-in-layer and refresh the R Output block.
- `build`: planned for the simulated loss table stage.

`profile_update` reads risk profile input from `Input_Profile!B12:I...`:

```text
AsAt | UY | LOB | Nrisk | SI | Premium | Incurred | Remarks
```

Layer definitions are read from `Input_Profile!X13:AC...`:

```text
LayerID | Limit | Deductible | LayerName | Prior_LOBs | Current_LOBs
```

Coverage LOB rules:

- `Prior` rows use `Prior_LOBs`.
- `Current` rows use `Current_LOBs`.
- blank coverage LOB cells include all LOBs.
- nonblank coverage LOB cells can use comma, slash, or semicolon delimiters.

Exposure-in-layer is calculated per risk band as:

```text
Avg_SI = SI / Nrisk
expo_in_layer = Nrisk * min(max(Avg_SI - Deductible, 0), Limit)
```

Adjusted exposure is calculated from the matching prior/current LOB exposure
rating parameters:

```text
adj_expo_in_layer = expo_in_layer * SubjectPrem / ActualPrem
```

The R action writes:

```text
_BERTToolkitTemp/xl_pricing_tool/profile/profile_output.csv
_BERTToolkitTemp/xl_pricing_tool/profile/profile_si_composition_data.csv
_BERTToolkitTemp/xl_pricing_tool/profile/profile_si_composition_chart.png
```

VBA overwrites only:

```text
Input_Profile!BH12:BS...
```

The left and middle workbook formulas remain untouched. VBA also inserts the SI
composition chart at the anchor address in `L56`, or `K57` if `L56` is blank.

The SI composition chart is a faceted 100% stacked bar chart. It uses average
SI risk bands from `K37:K44`; a `Band Cut` label and trailing blank slots are
allowed, but nonblank cut values must be numeric and nonnegative and at least
two unique cut points must be present. Facets are read from `K47:K55`; a header
value of `LOB` is ignored, blank facet inputs mean all LOBs, and
comma/slash/semicolon delimiters are supported.

## Tab: Risk Loss Fitting

Marker:

```text
<<RiskFit>>
```

Implemented chart-stage action:

```text
riskfit_update
```

Default behavior:

- `gather`: not defined when `Gather From` is blank.
- `update`: produces the first diagnostic loss comparison chart.
- `build`: no build action is defined for this tab.

The tab currently combines loss input, Excel-side on-level adjustments, frequency
and severity summaries, and a compact R input/output block.

The primary loss table starts at `B12:P`:

```text
AsAt | ClaimID | LOB | Insured | AY | TreatyYear | Actual_Incurred
OnLevel_Incurred | ReturnPeriod | OriginalRate | GNPI_Scale | RP_Scale
FinalRate | Sev_Trend | Final_Incurred
```

The current workbook formulas already compute the adjusted loss amount in
`Final_Incurred`, using:

- inflation severity trend toggle from `W5`,
- GNPI occurrence adjustment toggle from `W6`,
- customized return-period adjustment toggle from `W7`,
- `Min`, `Max`, and `StartYear` from `S5:S7`.

The middle summary area calculates:

```text
R14:T32   Exposure by year
V14:W33   Actual claim counts
Y14:Z35   Adjusted frequency
AB17:AE21 Severity threshold comparison
```

The existing R handoff layout is:

```text
BC12:BG18  R Dataspec
BI12:BL14  R Input:  AsAt | LossCause | lambda | alpha
BN12:BO14  R Output: AsAt | Alpha
```

Current `riskfit_update` behavior:

- read the loss table and controls from the current workbook;
- support only `Prior` and `Current`;
- keep frequency Excel-owned, including lambda and adjusted frequency formulas;
- produce the first diagnostic chart as an R-generated PNG;
- insert that PNG into `Risk Loss Fitting!R54:AK98`.

Loss comparison chart controls:

```text
T56 = minimum loss included
T57 = chart start year
T58 = showActual
Z56 = main title, shared by both panels
Z57 = adjusted loss subtitle
Z58 = actual loss subtitle
S61:T... = layer structure, Limit | Ded
```

Chart behavior:

- individual losses, not annual aggregate losses;
- two separate PNG charts are produced;
- the final incurred chart uses `Final_Incurred`;
- the actual incurred chart uses `Actual_Incurred` only when `showActual` is true;
- `TreatyYear` is the facet;
- individual losses are columns inside each treaty-year facet;
- `Prior` and `Current` are side-by-side columns;
- layers define the y-axis maximum and background banding only;
- no layer legend or layer labels are produced.

Generated files:

```text
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_loss_comparison_data.csv
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_final_incurred_loss_chart.png
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_actual_incurred_loss_chart.png
```

Later `riskfit_update` stages will add MLE severity fitting across selected
distributions.

Planned severity fit family:

```text
Pareto
Lognormal
Gamma
Loggamma
Weibull
Loglogistic
```

Open design questions:

- Which distributions should be included in the first implementation versus
  staged after Pareto?
- What output schema should replace or extend `BN12:BO14` once multiple
  distributions are fitted?
