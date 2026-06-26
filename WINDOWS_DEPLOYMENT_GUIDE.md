# Windows Deployment Guide

This guide explains how to deploy the full BERT/GRe toolkit on a Windows machine
when R is already installed but no required R packages are available yet.

## 1. Confirm Prerequisites

Install or confirm the following software:

```text
Microsoft Excel for Windows
R for Windows
BERT 2.4.4 for Excel
```

The toolkit has been developed against BERT 2.4.4, which uses an older R runtime.
The package install steps below use a dated package repository so Windows binary
packages are available for that R version.

## 2. Copy the Toolkit Folder

Copy the entire `CompanyTools` folder to the root of the Windows `C:` drive.

The final path must be:

```text
C:\CompanyTools\BERTToolkit
```

Confirm these files exist:

```text
C:\CompanyTools\BERTToolkit\r\bootstrap.R
C:\CompanyTools\BERTToolkit\r\registry.R
C:\CompanyTools\BERTToolkit\excel\vba\BERTToolkitClient.bas
C:\CompanyTools\BERTToolkit\excel\templates\Template_China_Exposure_Mapping.xlsm
```

## 3. Configure BERT Startup

Find BERT's startup `functions.R` file.

Add this line:

```r
source("C:/CompanyTools/BERTToolkit/r/bootstrap.R")
```

Restart Excel after saving `functions.R`.

## 4. Install Required R Packages

Open Excel, then open the BERT console.

The current toolkit package set is:

```text
sf
ggplot2
dplyr
magick
readxl
openxlsx
xml2
```

Run this command as one complete command:

```r
install.packages(c("sf","ggplot2","dplyr","magick","readxl","openxlsx","xml2"), repos="https://packagemanager.posit.co/cran/2019-06-01", type="win.binary", dependencies=c("Depends","Imports","LinkingTo"))
```

Do not paste only the repository URL by itself. It must remain inside quotes as
the `repos = "..."` argument.

If R asks whether to create a personal library, answer:

```text
yes
```

If R asks whether to use a personal library, answer:

```text
yes
```

Restart Excel after package installation completes.

## 5. Enable VBA Project Access

Excel must allow the deployment script to embed the VBA module into the workbook.

In Excel:

```text
File > Options > Trust Center > Trust Center Settings > Macro Settings
```

Enable:

```text
Trust access to the VBA project object model
```

Also enable macros for trusted toolkit workbooks according to your company policy.

## 6. Embed the VBA Module Into the Template

Run this script in Windows:

```text
C:\CompanyTools\BERTToolkit\excel\scripts\embed_vba.vbs
```

Expected success message:

```text
Embedded VBA module into: C:\CompanyTools\BERTToolkit\excel\templates\Template_China_Exposure_Mapping.xlsm
```

If the script fails with a VBA project access message, return to Step 5 and
confirm the Trust Center setting is enabled.

## 7. Open the Template Workbook

Open:

```text
C:\CompanyTools\BERTToolkit\excel\templates\Template_China_Exposure_Mapping.xlsm
```

Save a working copy somewhere outside the toolkit install folder, for example:

```text
D:\ClientWork\Template_China_Exposure_Mapping.xlsm
```

The toolkit code can stay installed under `C:\CompanyTools\BERTToolkit`; working
copies of templates can be saved anywhere.

## 8. Verify the GRe Ribbon

Open the workbook and confirm the ribbon includes:

```text
GRe Tool Ribbon
```

The ribbon should show three buttons:

```text
Gather
Update
Build
```

If the ribbon appears but buttons do nothing, the VBA module is probably not
embedded. Rerun Step 6.

## 9. Verify the Toolkit Connection

In Excel, run:

```text
Alt+F8 > BTK_CheckInstall > Run
```

Expected result:

```text
BERT Toolkit is available.
Version: 0.1.0
Missing packages: No missing packages.
```

If packages are missing, return to Step 4.

## 10. Run the China Exposure Mapping Template

Go to the `Parameters` sheet.

Confirm cell `A1` contains:

```text
<exposure map>
```

Click:

```text
GRe Tool Ribbon > Update
```

The tool will run the China exposure map workflow through BERT/R.

Generated files are overwritten in:

```text
<workbook folder>\_BERTToolkitTemp\china_exposure_map\
```

The workbook should refresh embedded images:

```text
plot_EQ_yoy.png -> Parameters!N17
plot_WS_yoy.png -> Parameters!N50
```

## 11. Understand Supported Actions

The GRe ribbon uses the active sheet's `A1` objective marker to dispatch actions.

Current mappings:

```text
<exposure map>     -> china_exposure_map
<curve fit risk>   -> curve_fit_risk
```

The China exposure mapping template currently supports only:

```text
Update
```

`Gather` and `Build` are intentionally unsupported for this template. They should
show a clean unsupported-action message.

## 12. Troubleshooting

### BERT Toolkit Version Works, But Tool Fails With Missing Packages

Install packages using the exact Step 4 command. BERT 2.4.4 uses an old R
runtime, so modern CRAN may not provide compatible binaries.

Current package checks are driven by `BTK.RequiredPackages()` in:

```text
C:\CompanyTools\BERTToolkit\r\bootstrap.R
```

### Error: Unexpected `/` in `https://...`

The repository URL was pasted directly into the R console. Paste the full
`install.packages(...)` command from Step 4.

### Excel Shows XML Import Error

Do not import `.bas` files using Excel's XML import controls. Use the deployment
script in Step 6, or manually import VBA through:

```text
Alt+F11 > File > Import File...
```

### Ribbon Appears, But Buttons Cannot Run Macros

Make sure the VBA module was embedded by Step 6 and macros are enabled for the
workbook.

### Unsupported Action Message

This is expected when a template does not support the clicked action. For China
exposure mapping, only `Update` is currently supported.

## 13. Files Maintained by the Toolkit

Main toolkit folder:

```text
C:\CompanyTools\BERTToolkit
```

Template manual:

```text
C:\CompanyTools\BERTToolkit\excel\TEMPLATE_MANUAL.md
```

Current production template:

```text
C:\CompanyTools\BERTToolkit\excel\templates\Template_China_Exposure_Mapping.xlsm
```

VBA source:

```text
C:\CompanyTools\BERTToolkit\excel\vba\BERTToolkitClient.bas
```

VBA embedding script:

```text
C:\CompanyTools\BERTToolkit\excel\scripts\embed_vba.vbs
```

R bootstrap:

```text
C:\CompanyTools\BERTToolkit\r\bootstrap.R
```

R tool registry:

```text
C:\CompanyTools\BERTToolkit\r\registry.R
```

## 14. Deployment Checklist

```text
[ ] CompanyTools copied to C:\
[ ] C:\CompanyTools\BERTToolkit exists
[ ] BERT functions.R sources C:/CompanyTools/BERTToolkit/r/bootstrap.R
[ ] Excel restarted after BERT startup edit
[ ] R packages installed from the 2019-06-01 snapshot: sf, ggplot2, dplyr, magick, readxl, openxlsx, xml2
[ ] VBA project access enabled
[ ] embed_vba.vbs completed successfully
[ ] Template workbook opens with GRe Tool Ribbon
[ ] BTK_CheckInstall reports no missing packages
[ ] Parameters!A1 contains <exposure map>
[ ] GRe Tool Ribbon > Update runs successfully
[ ] Output folder and embedded images refresh
```
