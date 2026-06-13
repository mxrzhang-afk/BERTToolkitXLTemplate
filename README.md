# BERTToolkit

For a new Windows deployment, start with:

```text
WINDOWS_DEPLOYMENT_GUIDE.md
```

Windows target path:

```text
C:\CompanyTools\BERTToolkit
```

Copy the `CompanyTools` folder to `C:\` so this repository becomes:

```text
C:\CompanyTools\BERTToolkit
```

## BERT Startup

In BERT's startup `functions.R`, source the toolkit bootstrap:

```r
source("C:/CompanyTools/BERTToolkit/r/bootstrap.R")
```

Excel/VBA should call stable public functions such as:

```vb
Application.Run "BERT.Call", "BTK.RunTool", "china_exposure_map", ThisWorkbook.FullName, ""
Application.Run "BERT.Call", "BTK.DispatchTool", "china_exposure_map", "update", ThisWorkbook.FullName, ""
Application.Run "BERT.Call", "BTK.ValidateTool", "china_exposure_map", ThisWorkbook.FullName
Application.Run "BERT.Call", "BTK.Version"
```

If Excel shows runtime error 1004 for `R.BTK.Version`, it usually means Excel cannot see
the BERT-exported function name yet. Import `excel\vba\BERTToolkitClient.bas` into the
workbook and run `BTK_CheckInstall`; the VBA wrapper calls BERT through `BERT.Call` and
shows a setup checklist when BERT or the toolkit bootstrap is not loaded.

If `BTK_RunChinaExposureMap` reports missing R packages, run:

```text
Alt+F8 > BTK_InstallPackages > Run
```

This installs the required packages into the R library used by BERT:

```r
c("sf", "ggplot2", "dplyr", "magick", "readxl", "openxlsx")
```

Because BERT 2.4.4 uses R 3.5.0, the installer uses a dated Posit Package
Manager snapshot with Windows binaries:

```r
install.packages(
  c("sf", "ggplot2", "dplyr", "magick", "readxl", "openxlsx"),
  repos = "https://packagemanager.posit.co/cran/2019-06-01",
  type = "win.binary",
  dependencies = c("Depends", "Imports", "LinkingTo")
)
```

After installation, restart Excel or run `BTK_CheckInstall` again before running the map.

Import the `.bas` file from the Visual Basic Editor, not from Excel's XML import
dialog:

```text
Excel > Alt+F11 > File > Import File... > BERTToolkitClient.bas
```

The Developer tab also has XML import controls; those will fail with an "XML parse"
error because `.bas` is a VBA source file, not an XML data file.

## GRe Workbook Actions

The macro-enabled template contains a `GRe Tool Ribbon` custom UI with three
buttons:

```text
Gather
Update
Build
```

Each button dispatches from the active sheet objective marker in `A1`, for example:

```text
<exposure map>
<curve fit risk>
```

Current objective mappings:

```text
<exposure map>     -> china_exposure_map
<curve fit risk>  -> curve_fit_risk
```

Tools do not need to support every action. Unsupported actions return a clean
message listing the actions currently available for that tool. The exposure map
tool currently supports `update`; `gather` and `build` are unsupported for this
template.

The workbook has the ribbon XML embedded and the source XML is stored at:

```text
excel/customUI/customUI.xml
```

The VBA source is stored at:

```text
excel/vba/BERTToolkitClient.bas
```

To embed the VBA module into `Template_China_Exposure_Mapping.xlsm` without manually using the Visual Basic
Editor, run this script inside the Windows VM:

```text
C:\CompanyTools\BERTToolkit\excel\scripts\embed_vba.vbs
```

Excel must allow programmatic VBA project access:

```text
File > Options > Trust Center > Trust Center Settings > Macro Settings >
Trust access to the VBA project object model
```

## Workbook-Local Runs

The toolkit code stays installed under:

```text
C:\CompanyTools\BERTToolkit
```

Excel files can be saved anywhere. When VBA calls a tool with a blank output folder, the toolkit creates a temporary run folder beside the calling workbook:

```text
D:\Any\Client\Folder\Template_China_Exposure_Mapping.xlsm
D:\Any\Client\Folder\_BERTToolkitTemp\china_exposure_map\
```

This means templates can move freely and do not need to be stored under `C:\CompanyTools\BERTToolkit`.
The same output folder is reused on every run. `plot_EQ.png`, `plot_WS.png`,
`plot_EQ_yoy.png`, and `plot_WS_yoy.png` are overwritten. `BTK_RunChinaExposureMap`
then replaces the embedded YOY images on the `Parameters` sheet starting at `N17`.

Recommended VBA call pattern:

```vb
Sub RunChinaExposureMap()
    Dim result As Variant

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Please save the workbook before running the tool.", vbExclamation
        Exit Sub
    End If

    ThisWorkbook.Save

    result = Application.Run( _
        "BERT.Call", _
        "BTK.RunTool", _
        "china_exposure_map", _
        ThisWorkbook.FullName, _
        "" _
    )

    MsgBox CStr(result), vbInformation
End Sub
```

Reusable VBA helper:

```text
excel\vba\BERTToolkitClient.bas
```

If a specific output folder is needed, pass it as the fourth argument. Otherwise leave it blank and the workbook-local temporary folder will be used.

## Structure

```text
config/                         Shared toolkit configuration
r/bootstrap.R                   Single file sourced by BERT
r/registry.R                    Tool registry and dispatch
r/utils/                        Shared helpers
r/tools/china_exposure_map/     Current map tool module
excel/templates/                Workbook templates and references
excel/customUI/                 Ribbon XML source
excel/scripts/                  Workbook packaging helpers
excel/vba/                      Reusable VBA caller module
outputs/                        Reserved; workbook outputs normally do not go here
logs/                           Toolkit-level logs
```

Each tool should expose:

```r
<tool>_run(workbook_path, output_dir = NULL, context = list())
<tool>_validate(workbook_path, context = list())
```

The current scaffold includes the China map shapefile assets and the macro-enabled mapping input template at `excel/templates/Template_China_Exposure_Mapping.xlsm`.

Template-level user documentation is maintained in:

```text
excel/TEMPLATE_MANUAL.md
```
