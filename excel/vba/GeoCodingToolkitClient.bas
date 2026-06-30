Attribute VB_Name = "GeoCodingToolkitClient"
Option Explicit

Private Const GCT_SOURCE As String = "GeoCodingToolkitClient"
Private Const GCT_BOOTSTRAP As String = "C:/CompanyTools/BERTToolkit/r/bootstrap.R"

Private Function GCT_InstallHelp() As String
    GCT_InstallHelp = _
        "BERT is not available to Excel yet." & vbCrLf & vbCrLf & _
        "In the Parallels Windows VM, check:" & vbCrLf & _
        "1. BERT is installed and enabled in Excel." & vbCrLf & _
        "2. C:\CompanyTools\BERTToolkit exists, or bootstrap.R was sourced from the real toolkit folder." & vbCrLf & _
        "3. BERT's startup functions.R contains:" & vbCrLf & _
        "   source(""C:/CompanyTools/BERTToolkit/r/bootstrap.R"")" & vbCrLf & _
        "4. Excel was restarted after changing functions.R."
End Function

Private Function GCT_CallR(ByVal functionName As String, ParamArray args() As Variant) As Variant
    Dim argCount As Long

    On Error Resume Next
    argCount = UBound(args) - LBound(args) + 1
    If Err.Number <> 0 Then
        argCount = 0
        Err.Clear
    End If
    On Error GoTo BertCallFailed

    Select Case argCount
        Case 0
            GCT_CallR = Application.Run("BERT.Call", functionName)
        Case 1
            GCT_CallR = Application.Run("BERT.Call", functionName, args(0))
        Case 2
            GCT_CallR = Application.Run("BERT.Call", functionName, args(0), args(1))
        Case 3
            GCT_CallR = Application.Run("BERT.Call", functionName, args(0), args(1), args(2))
        Case 4
            GCT_CallR = Application.Run("BERT.Call", functionName, args(0), args(1), args(2), args(3))
        Case 5
            GCT_CallR = Application.Run("BERT.Call", functionName, args(0), args(1), args(2), args(3), args(4))
        Case Else
            Err.Raise vbObjectError + 5302, GCT_SOURCE, "GCT_CallR supports up to 5 arguments."
    End Select
    Exit Function

BertCallFailed:
    Err.Raise vbObjectError + 5303, GCT_SOURCE, _
        GCT_InstallHelp() & vbCrLf & vbCrLf & _
        "Failed R call: " & functionName & vbCrLf & _
        "Excel error " & Err.Number & ": " & Err.Description
End Function

Private Sub GCT_SourceBootstrap()
    On Error GoTo BootstrapFailed
    Application.Run "BERT.Exec", "source(""" & GCT_BOOTSTRAP & """)"
    Exit Sub

BootstrapFailed:
    Err.Raise vbObjectError + 5304, GCT_SOURCE, _
        GCT_InstallHelp() & vbCrLf & vbCrLf & _
        "Failed to source: " & GCT_BOOTSTRAP & vbCrLf & _
        "Excel error " & Err.Number & ": " & Err.Description
End Sub

Private Function GCT_OutputFolderFromResult(ByVal resultText As String) As String
    Dim normalized As String
    Dim lines() As String
    Dim i As Long
    Dim prefix As String

    prefix = "Output folder:"
    normalized = Replace(resultText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    lines = Split(normalized, vbLf)

    For i = LBound(lines) To UBound(lines)
        If Left$(Trim$(lines(i)), Len(prefix)) = prefix Then
            GCT_OutputFolderFromResult = Trim$(Mid$(Trim$(lines(i)), Len(prefix) + 1))
            Exit Function
        End If
    Next i
End Function

Private Function GCT_ObjectiveKey(ByVal objectiveText As String) As String
    Dim key As String

    key = LCase$(Trim$(objectiveText))
    Do While Left$(key, 1) = "<" And Len(key) > 0
        key = Mid$(key, 2)
    Loop
    Do While Right$(key, 1) = ">" And Len(key) > 0
        key = Left$(key, Len(key) - 1)
    Loop

    key = Replace(key, "-", " ")
    key = Replace(key, "_", " ")
    key = Application.WorksheetFunction.Trim(key)
    GCT_ObjectiveKey = key
End Function

Private Function GCT_IsInputNormSheet(ByVal ws As Worksheet) As Boolean
    GCT_IsInputNormSheet = (GCT_ObjectiveKey(CStr(ws.Range("A1").Value)) = "inputnorm")
End Function

Private Function GCT_OutputPath(ByVal outputFolder As String, ByVal fileName As String) As String
    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If
    GCT_OutputPath = outputFolder & "inputnorm" & Application.PathSeparator & fileName
End Function

Private Sub GCT_ClearRange(ByVal ws As Worksheet, ByVal firstCell As String, ByVal maxColumns As Long)
    Dim first As Range
    Dim lastRow As Long

    Set first = ws.Range(firstCell)
    lastRow = ws.Rows.Count
    ws.Range(first, first.Offset(lastRow - first.Row, maxColumns - 1)).ClearContents
End Sub

Private Sub GCT_LoadCsvToRange(ByVal csvPath As String, ByVal destination As Range)
    If Len(Dir(csvPath)) = 0 Then
        Err.Raise vbObjectError + 5301, GCT_SOURCE, "Expected output file was not created: " & csvPath
    End If

    With destination.Worksheet.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=destination)
        .TextFileParseType = xlDelimited
        .TextFilePlatform = 65001
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .AdjustColumnWidth = False
        .Refresh BackgroundQuery:=False
        .Delete
    End With
End Sub

Public Sub GCT_BrowseInputNormSourceFile()
    Dim selectedPath As Variant

    On Error GoTo BrowseFailed

    If Not GCT_IsInputNormSheet(ActiveSheet) Then
        MsgBox "Open an <<inputnorm>> sheet before browsing for a source file.", vbExclamation, "Geocode Tool"
        Exit Sub
    End If

    If ActiveCell.CountLarge <> 1 Or Intersect(ActiveCell, ActiveSheet.Range("C14")) Is Nothing Then
        MsgBox "Select cell C14 on the active <<inputnorm>> sheet first.", vbExclamation, "Geocode Tool"
        Exit Sub
    End If

    selectedPath = Application.GetOpenFilename( _
        FileFilter:="Source files (*.csv;*.txt;*.xlsx;*.xlsm),*.csv;*.txt;*.xlsx;*.xlsm,All files (*.*),*.*", _
        Title:="Select source data file", _
        MultiSelect:=False)

    If VarType(selectedPath) = vbBoolean And selectedPath = False Then
        Exit Sub
    End If

    ActiveSheet.Range("C14").Value = CStr(selectedPath)
    Exit Sub

BrowseFailed:
    MsgBox Err.Description, vbCritical, "Geocode Tool"
End Sub

Public Function GCT_RefreshInputNormGather(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim mappingPath As String

    Set ws = ThisWorkbook.Worksheets(sheetName)
    mappingPath = GCT_OutputPath(outputFolder, "inputnorm_header_mapping.csv")

    GCT_ClearRange ws, "B26", 3
    GCT_LoadCsvToRange mappingPath, ws.Range("B26")

    GCT_RefreshInputNormGather = "InputNorm header mapping refreshed on " & sheetName & "."
End Function

Public Function GCT_RefreshInputNormUpdate(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim previewPath As String

    Set ws = ThisWorkbook.Worksheets(sheetName)
    previewPath = GCT_OutputPath(outputFolder, "inputnorm_data_preview.csv")

    GCT_ClearRange ws, "N14", 27
    GCT_LoadCsvToRange previewPath, ws.Range("N14")

    GCT_RefreshInputNormUpdate = "InputNorm normalized data preview refreshed on " & sheetName & "."
End Function

Private Function GCT_DispatchTool(ByVal action As String, Optional ByVal outputDir As String = "") As Variant
    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 5305, GCT_SOURCE, "Please save the workbook before running a Geocode tool action."
    End If

    If Not GCT_IsInputNormSheet(ActiveSheet) Then
        Err.Raise vbObjectError + 5306, GCT_SOURCE, "Open an <<inputnorm>> sheet before running this action."
    End If

    ThisWorkbook.Save
    GCT_SourceBootstrap

    GCT_DispatchTool = GCT_CallR( _
        "BTK.DispatchTool", _
        "goecode_tool", _
        "inputnorm_" & LCase$(action), _
        ThisWorkbook.FullName, _
        outputDir, _
        ActiveSheet.Name _
    )
End Function

Private Sub GCT_HandleResult(ByVal action As String, ByVal result As Variant)
    Dim resultText As String
    Dim outputFolder As String

    If IsError(result) Then
        MsgBox "BERT returned an Excel error before returning a text result." & vbCrLf & _
               "Check that BERT is loaded and C:\CompanyTools\BERTToolkit\r\bootstrap.R can be sourced.", _
               vbCritical, "Geocode Tool"
        Exit Sub
    End If

    resultText = CStr(result)
    If Left$(resultText, 6) = "ERROR:" Then
        MsgBox resultText, vbCritical, "Geocode Tool"
        Exit Sub
    End If

    outputFolder = GCT_OutputFolderFromResult(resultText)
    If action = "gather" And InStr(1, resultText, "InputNorm gather completed.", vbTextCompare) > 0 Then
        resultText = resultText & vbCrLf & vbCrLf & GCT_RefreshInputNormGather(outputFolder, ActiveSheet.Name)
        ThisWorkbook.Save
    End If

    If action = "update" And InStr(1, resultText, "InputNorm update completed.", vbTextCompare) > 0 Then
        resultText = resultText & vbCrLf & vbCrLf & GCT_RefreshInputNormUpdate(outputFolder, ActiveSheet.Name)
        ThisWorkbook.Save
    End If

    MsgBox resultText, vbInformation, "Geocode Tool"
End Sub

Public Sub GCT_Gather()
    GCT_HandleResult "gather", GCT_DispatchTool("gather")
End Sub

Public Sub GCT_Update()
    GCT_HandleResult "update", GCT_DispatchTool("update")
End Sub

Public Sub GCT_Build()
    GCT_HandleResult "build", GCT_DispatchTool("build")
End Sub

Public Sub GCT_RibbonGather(ByVal control As Object)
    On Error GoTo DispatchFailed
    GCT_Gather
    Exit Sub

DispatchFailed:
    MsgBox Err.Description, vbCritical, "Geocode Tool"
End Sub

Public Sub GCT_RibbonUpdate(ByVal control As Object)
    On Error GoTo DispatchFailed
    GCT_Update
    Exit Sub

DispatchFailed:
    MsgBox Err.Description, vbCritical, "Geocode Tool"
End Sub

Public Sub GCT_RibbonBuild(ByVal control As Object)
    On Error GoTo DispatchFailed
    GCT_Build
    Exit Sub

DispatchFailed:
    MsgBox Err.Description, vbCritical, "Geocode Tool"
End Sub

Public Sub GRe_RibbonGather(ByVal control As Object)
    GCT_RibbonGather control
End Sub

Public Sub GRe_RibbonUpdate(ByVal control As Object)
    GCT_RibbonUpdate control
End Sub

Public Sub GRe_RibbonBuild(ByVal control As Object)
    GCT_RibbonBuild control
End Sub
