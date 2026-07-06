Attribute VB_Name = "GeoCodingToolkitClient"
Option Explicit

Private Const GCT_SOURCE As String = "GeoCodingToolkitClient"
Private Const GCT_BOOTSTRAP As String = "C:/CompanyTools/BERTToolkit/r/bootstrap.R"
Private Const GCT_PREVIEW_MAX_ROWS As Long = 10000

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

Private Function GCT_IsHistCrossSheet(ByVal ws As Worksheet) As Boolean
    GCT_IsHistCrossSheet = (GCT_ObjectiveKey(CStr(ws.Range("A1").Value)) = "histcross")
End Function

Private Function GCT_IsGeneralCrossWalkSheet(ByVal ws As Worksheet) As Boolean
    GCT_IsGeneralCrossWalkSheet = (GCT_ObjectiveKey(CStr(ws.Range("A1").Value)) = "generalcrosswalk")
End Function

Private Function GCT_IsSupportedSheet(ByVal ws As Worksheet) As Boolean
    GCT_IsSupportedSheet = GCT_IsInputNormSheet(ws) Or GCT_IsHistCrossSheet(ws) Or GCT_IsGeneralCrossWalkSheet(ws)
End Function

Private Function GCT_ActionPrefix(ByVal ws As Worksheet) As String
    If GCT_IsInputNormSheet(ws) Then
        GCT_ActionPrefix = "inputnorm"
    ElseIf GCT_IsHistCrossSheet(ws) Then
        GCT_ActionPrefix = "histcross"
    ElseIf GCT_IsGeneralCrossWalkSheet(ws) Then
        GCT_ActionPrefix = "generalcrosswalk"
    Else
        Err.Raise vbObjectError + 5307, GCT_SOURCE, "Open an <<inputnorm>>, <<HistCross>>, or <<GeneralCrossWalk>> sheet before running this action."
    End If
End Function

Private Function GCT_OutputPath(ByVal outputFolder As String, ByVal folderName As String, ByVal fileName As String) As String
    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If
    GCT_OutputPath = outputFolder & folderName & Application.PathSeparator & fileName
End Function

Private Function GCT_BlockHasData(ByVal ws As Worksheet, ByVal firstCell As String, ByVal maxColumns As Long) As Boolean
    Dim first As Range
    Dim searchRange As Range
    Dim found As Range

    Set first = ws.Range(firstCell)
    Set searchRange = ws.Range(first, ws.Cells(ws.Rows.Count, first.Column + maxColumns - 1))
    Set found = searchRange.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    GCT_BlockHasData = Not found Is Nothing
End Function

Private Function GCT_LastDataRowInBlock(ByVal ws As Worksheet, ByVal firstCell As String, ByVal maxColumns As Long) As Long
    Dim first As Range
    Dim searchRange As Range
    Dim found As Range

    Set first = ws.Range(firstCell)
    Set searchRange = ws.Range(first, ws.Cells(ws.Rows.Count, first.Column + maxColumns - 1))
    Set found = searchRange.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If found Is Nothing Then
        GCT_LastDataRowInBlock = first.Row - 1
    Else
        GCT_LastDataRowInBlock = found.Row
    End If
End Function

Private Function GCT_AreaLabel(ByVal firstCell As String, ByVal maxColumns As Long) As String
    Dim first As Range
    Set first = ActiveSheet.Range(firstCell)
    GCT_AreaLabel = firstCell & ":" & ActiveSheet.Cells(first.Row + GCT_PREVIEW_MAX_ROWS, first.Column + maxColumns - 1).Address(False, False)
End Function

Private Function GCT_ColumnAreaLabel(ByVal firstCell As String, ByVal maxColumns As Long) As String
    Dim first As Range
    Set first = ActiveSheet.Range(firstCell)
    GCT_ColumnAreaLabel = firstCell & ":" & ActiveSheet.Cells(ActiveSheet.Rows.Count, first.Column + maxColumns - 1).Address(False, False)
End Function

Private Sub GCT_ClearRange(ByVal ws As Worksheet, ByVal firstCell As String, ByVal maxColumns As Long)
    Dim first As Range
    Dim lastRow As Long

    Set first = ws.Range(firstCell)
    lastRow = GCT_LastDataRowInBlock(ws, firstCell, maxColumns)
    If lastRow >= first.Row Then
        ws.Range(first, ws.Cells(lastRow, first.Column + maxColumns - 1)).ClearContents
    End If
End Sub

Private Function GCT_ConfirmClearBeforeRefresh(ByVal action As String) As Boolean
    Dim prefix As String
    Dim warning As String

    GCT_ConfirmClearBeforeRefresh = True
    prefix = GCT_ActionPrefix(ActiveSheet)

    If action = "gather" Then
        If prefix = "generalcrosswalk" Then
            If GCT_BlockHasData(ActiveSheet, "B26", 4) Then
                warning = "B26:E"
            End If
            If Len(Trim$(CStr(ActiveSheet.Range("C14").Value))) > 0 And GCT_BlockHasData(ActiveSheet, "N14", 29) Then
                If Len(warning) > 0 Then warning = warning & " and "
                warning = warning & GCT_ColumnAreaLabel("N14", 29)
            End If
        Else
            If GCT_BlockHasData(ActiveSheet, "B26", 3) Then
                warning = "B26:D"
            End If
        End If
    ElseIf action = "update" Then
        If prefix = "inputnorm" Then
            If GCT_BlockHasData(ActiveSheet, "N14", 27) Then
                warning = GCT_AreaLabel("N14", 27)
            End If
        ElseIf prefix = "histcross" Then
            If GCT_BlockHasData(ActiveSheet, "N14", 27) Then
                warning = GCT_AreaLabel("N14", 27)
            End If
            If GCT_BlockHasData(ActiveSheet, "AP14", 27) Then
                If Len(warning) > 0 Then warning = warning & " and "
                warning = warning & GCT_ColumnAreaLabel("AP14", 27)
            End If
        ElseIf prefix = "generalcrosswalk" Then
            If GCT_BlockHasData(ActiveSheet, "AQ14", 54) Then
                warning = GCT_ColumnAreaLabel("AQ14", 54)
            End If
            If GCT_BlockHasData(ActiveSheet, "CS14", 40) Then
                If Len(warning) > 0 Then warning = warning & " and "
                warning = warning & GCT_ColumnAreaLabel("CS14", 40)
            End If
        End If
    End If

    If Len(warning) = 0 Then Exit Function

    If MsgBox( _
        "The refresh area " & warning & " already contains data." & vbCrLf & vbCrLf & _
        "Clear this area before running " & UCase$(action) & "?", _
        vbQuestion + vbYesNo, _
        "Geocode Tool" _
    ) <> vbYes Then
        GCT_ConfirmClearBeforeRefresh = False
    End If
End Function

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

    If Not GCT_IsSupportedSheet(ActiveSheet) Then
        MsgBox "Open an <<inputnorm>>, <<HistCross>>, or <<GeneralCrossWalk>> sheet before browsing for a source file.", vbExclamation, "Geocode Tool"
        Exit Sub
    End If

    If ActiveCell.CountLarge <> 1 Or Intersect(ActiveCell, ActiveSheet.Range("C14")) Is Nothing Then
        MsgBox "Select cell C14 on the active Geocode action sheet first.", vbExclamation, "Geocode Tool"
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
    mappingPath = GCT_OutputPath(outputFolder, "inputnorm", "inputnorm_header_mapping.csv")

    GCT_ClearRange ws, "B26", 3
    GCT_LoadCsvToRange mappingPath, ws.Range("B26")

    GCT_RefreshInputNormGather = "InputNorm header mapping refreshed on " & sheetName & "."
End Function

Public Function GCT_RefreshHistCrossGather(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim mappingPath As String

    Set ws = ThisWorkbook.Worksheets(sheetName)
    mappingPath = GCT_OutputPath(outputFolder, "histcross", "histcross_header_mapping.csv")

    GCT_ClearRange ws, "B26", 3
    GCT_LoadCsvToRange mappingPath, ws.Range("B26")

    GCT_RefreshHistCrossGather = "HistCross header mapping refreshed on " & sheetName & "."
End Function

Public Function GCT_RefreshGeneralCrossWalkGather(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim inputPath As String
    Dim configPath As String

    Set ws = ThisWorkbook.Worksheets(sheetName)
    inputPath = GCT_OutputPath(outputFolder, "generalcrosswalk", "generalcrosswalk_input_preview.csv")
    configPath = GCT_OutputPath(outputFolder, "generalcrosswalk", "generalcrosswalk_config.csv")

    GCT_ClearRange ws, "N14", 29
    GCT_ClearRange ws, "B26", 4
    GCT_LoadCsvToRange inputPath, ws.Range("N14")
    GCT_LoadCsvToRange configPath, ws.Range("B26")

    GCT_RefreshGeneralCrossWalkGather = "GeneralCrossWalk input and configuration refreshed on " & sheetName & "."
End Function

Public Function GCT_RefreshInputNormUpdate(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim previewPath As String

    Set ws = ThisWorkbook.Worksheets(sheetName)
    previewPath = GCT_OutputPath(outputFolder, "inputnorm", "inputnorm_data_preview.csv")

    GCT_ClearRange ws, "N14", 27
    GCT_LoadCsvToRange previewPath, ws.Range("N14")

    GCT_RefreshInputNormUpdate = "InputNorm normalized data preview refreshed on " & sheetName & "."
End Function

Public Function GCT_RefreshHistCrossUpdate(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim matchedPath As String
    Dim unmatchedPath As String

    Set ws = ThisWorkbook.Worksheets(sheetName)
    matchedPath = GCT_OutputPath(outputFolder, "histcross", "histcross_matched_preview.csv")
    unmatchedPath = GCT_OutputPath(outputFolder, "histcross", "histcross_unmatched_preview.csv")

    GCT_ClearRange ws, "N14", 27
    GCT_ClearRange ws, "AP14", 27
    GCT_LoadCsvToRange matchedPath, ws.Range("N14")
    GCT_LoadCsvToRange unmatchedPath, ws.Range("AP14")

    GCT_RefreshHistCrossUpdate = "HistCross matched and unmatched previews refreshed on " & sheetName & "."
End Function

Public Function GCT_RefreshGeneralCrossWalkUpdate(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim mappedPath As String
    Dim reviewPath As String

    Set ws = ThisWorkbook.Worksheets(sheetName)
    mappedPath = GCT_OutputPath(outputFolder, "generalcrosswalk", "generalcrosswalk_final_mapped.csv")
    reviewPath = GCT_OutputPath(outputFolder, "generalcrosswalk", "generalcrosswalk_manual_review.csv")

    GCT_ClearRange ws, "AQ14", 54
    GCT_ClearRange ws, "CS14", 40
    GCT_LoadCsvToRange mappedPath, ws.Range("AQ14")
    GCT_LoadCsvToRange reviewPath, ws.Range("CS14")

    GCT_RefreshGeneralCrossWalkUpdate = "GeneralCrossWalk mapped output and review queue refreshed on " & sheetName & "."
End Function

Private Function GCT_DispatchTool(ByVal action As String, Optional ByVal outputDir As String = "") As Variant
    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 5305, GCT_SOURCE, "Please save the workbook before running a Geocode tool action."
    End If

    If Not GCT_IsSupportedSheet(ActiveSheet) Then
        Err.Raise vbObjectError + 5306, GCT_SOURCE, "Open an <<inputnorm>>, <<HistCross>>, or <<GeneralCrossWalk>> sheet before running this action."
    End If

    ThisWorkbook.Save
    GCT_SourceBootstrap

    GCT_DispatchTool = GCT_CallR( _
        "BTK.DispatchTool", _
        "goecode_tool", _
        GCT_ActionPrefix(ActiveSheet) & "_" & LCase$(action), _
        ThisWorkbook.FullName, _
        outputDir, _
        ActiveSheet.Name _
    )
End Function

Private Sub GCT_HandleResult(ByVal action As String, ByVal result As Variant)
    Dim resultText As String
    Dim outputFolder As String
    Dim actionPrefix As String

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
    actionPrefix = GCT_ActionPrefix(ActiveSheet)
    If action = "gather" Then
        If actionPrefix = "inputnorm" And InStr(1, resultText, "InputNorm gather completed.", vbTextCompare) > 0 Then
            resultText = resultText & vbCrLf & vbCrLf & GCT_RefreshInputNormGather(outputFolder, ActiveSheet.Name)
            ThisWorkbook.Save
        ElseIf actionPrefix = "histcross" And InStr(1, resultText, "HistCross gather completed.", vbTextCompare) > 0 Then
            resultText = resultText & vbCrLf & vbCrLf & GCT_RefreshHistCrossGather(outputFolder, ActiveSheet.Name)
            ThisWorkbook.Save
        ElseIf actionPrefix = "generalcrosswalk" And InStr(1, resultText, "GeneralCrossWalk gather completed.", vbTextCompare) > 0 Then
            resultText = resultText & vbCrLf & vbCrLf & GCT_RefreshGeneralCrossWalkGather(outputFolder, ActiveSheet.Name)
            ThisWorkbook.Save
        End If
    End If

    If action = "update" Then
        If actionPrefix = "inputnorm" And InStr(1, resultText, "InputNorm update completed.", vbTextCompare) > 0 Then
            resultText = resultText & vbCrLf & vbCrLf & GCT_RefreshInputNormUpdate(outputFolder, ActiveSheet.Name)
            ThisWorkbook.Save
        ElseIf actionPrefix = "histcross" And InStr(1, resultText, "HistCross update completed.", vbTextCompare) > 0 Then
            resultText = resultText & vbCrLf & vbCrLf & GCT_RefreshHistCrossUpdate(outputFolder, ActiveSheet.Name)
            ThisWorkbook.Save
        ElseIf actionPrefix = "generalcrosswalk" And InStr(1, resultText, "GeneralCrossWalk update completed.", vbTextCompare) > 0 Then
            resultText = resultText & vbCrLf & vbCrLf & GCT_RefreshGeneralCrossWalkUpdate(outputFolder, ActiveSheet.Name)
            ThisWorkbook.Save
        End If
    End If

    MsgBox resultText, vbInformation, "Geocode Tool"
End Sub

Public Sub GCT_Gather()
    If Not GCT_ConfirmClearBeforeRefresh("gather") Then Exit Sub
    GCT_HandleResult "gather", GCT_DispatchTool("gather")
End Sub

Public Sub GCT_Update()
    If Not GCT_ConfirmClearBeforeRefresh("update") Then Exit Sub
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
