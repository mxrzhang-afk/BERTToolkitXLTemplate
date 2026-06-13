# China Exposure Map Tool

This module is the BERT-ready home for the current China exposure map workflow.

Current source script to migrate:

```text
Map_user.Rmd
```

Expected workbook sheets:

- `Control`
- `Parameters`
- `Input_Data`
- `Output`
- `Log`

Template:

```text
excel/templates/Template_China_Exposure_Mapping.xlsm
```

Expected input columns:

```text
Province | Lat | Lon | EQ_New | EQ_Old | WS_New | WS_Old
```

Generated outputs:

```text
plot_EQ.png
plot_WS.png
plot_EQ_yoy.png
plot_WS_yoy.png
```

By default these outputs are overwritten in a stable workbook-local folder:

```text
<workbook folder>\_BERTToolkitTemp\china_exposure_map\
```

The `assets` folder contains the required China province and boundary shapefile components.
