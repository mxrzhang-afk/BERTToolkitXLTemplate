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

Planned action:

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
