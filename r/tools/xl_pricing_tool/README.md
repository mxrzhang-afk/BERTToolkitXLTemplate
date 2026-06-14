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

The shared VBA dispatcher passes the active worksheet name to R. Implemented
tab actions read from that active worksheet and refresh outputs back to the same
worksheet, so copied tabs can be renamed and run independently as long as the
`A1` marker and required layout are preserved. The original template tab names
remain fallback defaults for direct R calls that do not provide an active sheet.

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

Implemented action:

```text
riskfit_update
```

Default behavior:

- `gather`: not defined when `Gather From` is blank.
- `update`: produces loss comparison charts, weighted MLE severity fits, and
  distribution CDF diagnostic charts.
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
`Final_Incurred` and the fitting weight in `FinalRate`, using:

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
- fit severity distributions in R using `Final_Incurred` as the loss and
  `FinalRate` as the observation weight;
- produce the loss comparison charts as R-generated PNGs;
- produce one weighted CDF diagnostic chart per distribution family;
- insert the PNGs into the Risk Loss Fitting tab.

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
- no layer legend or layer labels are produced;
- loss comparison chart image width expands with the number of unique losses and
  treaty-year facets;
- Excel inserts the final chart at `X61` and the actual chart at `X94`, preserving
  the generated image aspect ratio.

Generated files:

```text
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_loss_comparison_data.csv
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_fit_summary.csv
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_fit_parameters.csv
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_final_incurred_loss_chart.png
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_actual_incurred_loss_chart.png
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_pareto_cdf_chart.png
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_loggamma_cdf_chart.png
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_weibull_cdf_chart.png
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_lognormal_cdf_chart.png
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_gamma_cdf_chart.png
_BERTToolkitTemp/xl_pricing_tool/riskfit/riskfit_loglogistic_cdf_chart.png
```

Severity fitting:

```text
loss_i   = Final_Incurred
weight_i = FinalRate
```

The weighted log-likelihood is:

```text
sum(weight_i * log density(loss_i | theta))
```

The empirical CDF diagnostics use the same weights:

```text
cumulative FinalRate for losses <= x / total FinalRate
```

Supported severity fit families:

```text
Pareto
Lognormal
Gamma
Loggamma
Weibull
Loglogistic
```

Pareto uses the user's minimum loss threshold as a fixed threshold and fits only
`alpha`. The other families fit the parameter sets shown in the workbook blocks.
The full fit summary is loaded to `CG7`, and the parameter table is loaded to
`CV7`. The distribution blocks in `AQ:CC` are Excel-owned; native formulas should
pull from the R output area rather than VBA overwriting those cells.

```text
CG7:CS...  Fit summary
CV7:DA...  Fit parameters
AQ:CC      Distribution blocks controlled by native Excel formulas
```

## Tab: CatLossOnLevel

Marker:

```text
<<CatOnLevel>>
```

Implemented first-stage action:

```text
cat_onlevel_gather
```

The gather action validates the external source tab names configured on the
active `<<CatOnLevel>>` sheet:

```text
C6 = GNPI source tab
C7 = CPI source tab
C8 = Aggregate source tab
```

It checks:

- `C6:C8` are filled;
- the named tabs exist in the active workbook;
- the GNPI source has expected input headers and non-empty data in `B12:I...`;
- the CPI source has expected inflation headers and non-empty data in `E15:H...`;
- the aggregate source has expected input headers and non-empty data in `B12:S...`.

No workbook ranges are overwritten by gather.

```text
cat_onlevel_update
```

The update action calculates the on-level handoff columns for the active
`<<CatOnLevel>>` sheet and writes a CSV consumed by the VBA client:

```text
cat_onlevel/cat_onlevel_output.csv -> active sheet T13:U...
```

Only `T:U` is refreshed. The source tabs remain driven by `C6:C8`.

Supported methods:

- `Province`: sums aggregate exposure for the listed provinces, peril, and
  from/to years.
- `Nationwide`: sums all aggregate provinces for the row peril.
- `Keyzone`: uses blended keyzone province flags from the configured
  control/CPI source tab.
- `User-defined`: first-pass placeholder, treated as `No On-level` and
  flagged with a warning.
- `No On-level`: returns `1` and `1`.
- `GNPI`: looks up GNPI from the configured GNPI tab and chooses one type per
  year using `Actual > Revised > Estimate`.
- `CPI`: returns `from_Expo = 1` and `to_Expo` from the row `AsAt` CPI factor
  column for `from_year`.

Rows where `In Prior/Current` is blank or zero return blank output. Missing
source years use the nearest available fallback and are counted in the action
message.
