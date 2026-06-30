Option Explicit

Dim fso
Dim scriptDir
Dim excelDir
Dim workbookPath
Dim xl
Dim wb
Dim components

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
excelDir = fso.GetParentFolderName(scriptDir)
workbookPath = fso.BuildPath(fso.BuildPath(excelDir, "templates"), "Template_Geocode_Tool.xlsm")

If Not fso.FileExists(workbookPath) Then
    WScript.Echo "Workbook not found: " & workbookPath
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
ImportModule components, fso.BuildPath(fso.BuildPath(excelDir, "vba"), "GeoCodingToolkitClient.bas"), "GeoCodingToolkitClient"

wb.Save
wb.Close False
xl.Quit

WScript.Echo "Embedded geocode VBA modules into: " & workbookPath

Sub ImportModule(ByVal components, ByVal modulePath, ByVal moduleName)
    Dim component

    If Not fso.FileExists(modulePath) Then
        WScript.Echo "VBA module not found: " & modulePath
        wb.Close False
        xl.Quit
        WScript.Quit 1
    End If

    For Each component In components
        If component.Name = moduleName Then
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
End Sub
