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
meaningful input rows end. The right-side settings block controls chart
presentation:

```text
L19  Start Year
L20  Level Chart Title
L21  Composition Chart Title
L22  Actual Color
L23  Revised Color
L24  Estimate Color
L25  Label Color
L26  Grid Color
L27  PNG Width
L28  PNG Height
L29  Font Scale
L30  Excel Image Width
L31  Level Image Height
L32  Composition Image Height
L34  Level Chart Anchor
L35  Composition Chart Anchor
L39  LOB Sequence Start
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
`2025 Revised`, `2025 Estimate`, and `2026 Estimate`. The second chart is a
stacked percentage LOB composition chart across those same year/type sections
from the configured start year. Charts are rendered by R as PNG images and then
inserted into the workbook by VBA at the configured anchor cells.
