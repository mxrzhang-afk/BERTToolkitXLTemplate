Option Explicit

Dim fso
Dim scriptDir
Dim excelDir
Dim workbookPath
Dim workbookName
Dim modulePath
Dim xl
Dim wb
Dim components
Dim component

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
excelDir = fso.GetParentFolderName(scriptDir)
If WScript.Arguments.Count > 0 Then
    workbookName = WScript.Arguments(0)
Else
    workbookName = "Template_China_Exposure_Mapping.xlsm"
End If
workbookPath = fso.BuildPath(fso.BuildPath(excelDir, "templates"), workbookName)
modulePath = fso.BuildPath(fso.BuildPath(excelDir, "vba"), "BERTToolkitClient.bas")

If Not fso.FileExists(workbookPath) Then
    WScript.Echo "Workbook not found: " & workbookPath
    WScript.Quit 1
End If

If Not fso.FileExists(modulePath) Then
    WScript.Echo "VBA module not found: " & modulePath
    WScript.Quit 1
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

On Error Resume Next
Set wb = xl.Workbooks.Open(workbookPath)
If Err.Number <> 0 Then
    WScript.Echo "Failed to open workbook: " & Err.Description
    xl.Quit
    WScript.Quit 1
End If
On Error GoTo 0

Set components = wb.VBProject.VBComponents
For Each component In components
    If component.Name = "BERTToolkitClient" Then
        components.Remove component
        Exit For
    End If
Next

On Error Resume Next
components.Import modulePath
If Err.Number <> 0 Then
    WScript.Echo "Failed to import VBA module: " & Err.Description
    WScript.Echo "In Excel, enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model."
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If
On Error GoTo 0

wb.Save
wb.Close False
xl.Quit

WScript.Echo "Embedded VBA module into: " & workbookPath
