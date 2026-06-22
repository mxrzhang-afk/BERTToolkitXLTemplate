Attribute VB_Name = "BERTToolkitClient"
Option Explicit

Private Const BTK_SOURCE As String = "BERTToolkitClient"
Private Const BTK_BOOTSTRAP As String = "C:/CompanyTools/BERTToolkit/r/bootstrap.R"
Private Function BTK_InstallHelp() As String
    BTK_InstallHelp = _
        "BERT is not available to Excel yet." & vbCrLf & vbCrLf & _
        "In the Parallels Windows VM, check:" & vbCrLf & _
        "1. BERT is installed and enabled in Excel." & vbCrLf & _
        "2. C:\CompanyTools\BERTToolkit exists, or bootstrap.R was sourced from the real toolkit folder." & vbCrLf & _
        "3. BERT's startup functions.R contains:" & vbCrLf & _
        "   source(""C:/CompanyTools/BERTToolkit/r/bootstrap.R"")" & vbCrLf & _
        "4. Excel was restarted after changing functions.R."
End Function

Private Function BTK_CallR(ByVal functionName As String, ParamArray args() As Variant) As Variant
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
            BTK_CallR = Application.Run("BERT.Call", functionName)
        Case 1
            BTK_CallR = Application.Run("BERT.Call", functionName, args(0))
        Case 2
            BTK_CallR = Application.Run("BERT.Call", functionName, args(0), args(1))
        Case 3
            BTK_CallR = Application.Run("BERT.Call", functionName, args(0), args(1), args(2))
        Case 4
            BTK_CallR = Application.Run("BERT.Call", functionName, args(0), args(1), args(2), args(3))
        Case 5
            BTK_CallR = Application.Run("BERT.Call", functionName, args(0), args(1), args(2), args(3), args(4))
        Case Else
            Err.Raise vbObjectError + 5102, BTK_SOURCE, "BTK_CallR supports up to 5 arguments."
    End Select
    Exit Function

BertCallFailed:
    Err.Raise vbObjectError + 5101, BTK_SOURCE, _
        BTK_InstallHelp() & vbCrLf & vbCrLf & _
        "Failed R call: " & functionName & vbCrLf & _
        "Excel error " & Err.Number & ": " & Err.Description
End Function

Private Sub BTK_SourceBootstrap()
    On Error GoTo BootstrapFailed
    Application.Run "BERT.Exec", "source(""" & BTK_BOOTSTRAP & """)"
    Exit Sub

BootstrapFailed:
    Err.Raise vbObjectError + 5103, BTK_SOURCE, _
        BTK_InstallHelp() & vbCrLf & vbCrLf & _
        "Failed to source: " & BTK_BOOTSTRAP & vbCrLf & _
        "Excel error " & Err.Number & ": " & Err.Description
End Sub

Private Function BTK_OutputFolderFromResult(ByVal resultText As String) As String
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
            BTK_OutputFolderFromResult = Trim$(Mid$(Trim$(lines(i)), Len(prefix) + 1))
            Exit Function
        End If
    Next i
End Function

Private Sub BTK_DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub BTK_DeleteChartObjects(ByVal ws As Worksheet)
    Dim i As Long

    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
End Sub

Private Sub BTK_InsertMapImage(ByVal ws As Worksheet, ByVal imagePath As String, ByVal shapeName As String, ByVal targetCell As Range)
    Dim picture As Object

    If Len(Dir(imagePath)) = 0 Then
        Err.Raise vbObjectError + 5110, BTK_SOURCE, "Map image was not created: " & imagePath
    End If

    BTK_DeleteShapeIfExists ws, shapeName

    Set picture = ws.Shapes.AddPicture( _
        Filename:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=targetCell.Left, _
        Top:=targetCell.Top, _
        Width:=420, _
        Height:=315 _
    )
    picture.Name = shapeName
    picture.LockAspectRatio = True
    picture.Placement = xlMoveAndSize
End Sub

Private Sub BTK_RefreshMapImages(ByVal outputFolder As String)
    Dim ws As Worksheet
    Dim eqPath As String
    Dim wsPath As String
    Dim eqAnchor As Range
    Dim wsAnchor As Range

    If Len(outputFolder) = 0 Then
        Err.Raise vbObjectError + 5111, BTK_SOURCE, "Could not find output folder in R result."
    End If

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    eqPath = outputFolder & "plot_EQ_yoy.png"
    wsPath = outputFolder & "plot_WS_yoy.png"

    Set ws = ThisWorkbook.Worksheets("Parameters")
    Set eqAnchor = ws.Range("N17")
    Set wsAnchor = ws.Range("N50")

    BTK_InsertMapImage ws, eqPath, "BTK_plot_EQ", eqAnchor
    BTK_InsertMapImage ws, wsPath, "BTK_plot_WS", wsAnchor
End Sub

Private Sub BTK_InsertImage(ByVal ws As Worksheet, ByVal imagePath As String, ByVal shapeName As String, ByVal targetCell As Range, ByVal width As Double, ByVal height As Double)
    Dim picture As Object

    If Len(Dir(imagePath)) = 0 Then
        Err.Raise vbObjectError + 5112, BTK_SOURCE, "Chart image was not created: " & imagePath
    End If

    BTK_DeleteShapeIfExists ws, shapeName

    Set picture = ws.Shapes.AddPicture( _
        Filename:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=targetCell.Left, _
        Top:=targetCell.Top, _
        Width:=width, _
        Height:=height _
    )
    picture.Name = shapeName
    picture.Placement = xlMoveAndSize
End Sub

Private Sub BTK_InsertImageFitWidth(ByVal ws As Worksheet, ByVal imagePath As String, ByVal shapeName As String, ByVal targetCell As Range, ByVal width As Double)
    Dim picture As Object

    If Len(Dir(imagePath)) = 0 Then
        Err.Raise vbObjectError + 5114, BTK_SOURCE, "Chart image was not created: " & imagePath
    End If

    BTK_DeleteShapeIfExists ws, shapeName

    Set picture = ws.Shapes.AddPicture( _
        Filename:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=targetCell.Left, _
        Top:=targetCell.Top, _
        Width:=-1, _
        Height:=-1 _
    )
    picture.Name = shapeName
    picture.LockAspectRatio = True
    picture.Width = width
    picture.Placement = xlMoveAndSize
End Sub

Private Function BTK_RangeFromSetting(ByVal ws As Worksheet, ByVal settingCell As String, ByVal defaultAddress As String) As Range
    Dim addressText As String

    addressText = Trim$(CStr(ws.Range(settingCell).Value))
    If Len(addressText) = 0 Then
        addressText = defaultAddress
    End If

    On Error Resume Next
    Set BTK_RangeFromSetting = ws.Range(addressText)
    On Error GoTo 0

    If BTK_RangeFromSetting Is Nothing Then
        Set BTK_RangeFromSetting = ws.Range(defaultAddress)
    End If
End Function

Private Function BTK_NumberFromSetting(ByVal ws As Worksheet, ByVal settingCell As String, ByVal defaultValue As Double) As Double
    Dim rawValue As Variant

    rawValue = ws.Range(settingCell).Value
    If IsNumeric(rawValue) And CDbl(rawValue) > 0 Then
        BTK_NumberFromSetting = CDbl(rawValue)
    Else
        BTK_NumberFromSetting = defaultValue
    End If
End Function

Private Sub BTK_LoadCsvToRange(ByVal csvPath As String, ByVal destination As Range)
    If Len(Dir(csvPath)) = 0 Then
        Err.Raise vbObjectError + 5113, BTK_SOURCE, "CSV output was not created: " & csvPath
    End If

    With destination.Worksheet.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=destination)
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .AdjustColumnWidth = False
        .Refresh BackgroundQuery:=False
        .Delete
    End With
End Sub

Private Function BTK_DefaultRunDir(ByVal toolId As String) As String
    BTK_DefaultRunDir = ThisWorkbook.Path & Application.PathSeparator & "_BERTToolkitTemp" & Application.PathSeparator & toolId
End Function

Private Sub BTK_EnsureFolder(ByVal folderPath As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long

    folderPath = Replace(folderPath, "/", Application.PathSeparator)
    If Right$(folderPath, 1) = Application.PathSeparator Then
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If
    If Len(folderPath) = 0 Or Len(Dir(folderPath, vbDirectory)) > 0 Then
        Exit Sub
    End If

    parts = Split(folderPath, Application.PathSeparator)
    currentPath = parts(LBound(parts))
    For i = LBound(parts) + 1 To UBound(parts)
        currentPath = currentPath & Application.PathSeparator & parts(i)
        If Len(Dir(currentPath, vbDirectory)) = 0 Then
            MkDir currentPath
        End If
    Next i
End Sub

Private Function BTK_CsvEscape(ByVal value As Variant) As String
    Dim text As String

    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then
        text = ""
    Else
        text = CStr(value)
    End If

    text = Replace(text, """", """""")
    BTK_CsvEscape = """" & text & """"
End Function

Private Function BTK_SafeFileName(ByVal value As String) As String
    Dim badChars As Variant
    Dim i As Long

    value = Trim$(value)
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
    For i = LBound(badChars) To UBound(badChars)
        value = Replace(value, CStr(badChars(i)), "_")
    Next i
    If Len(value) = 0 Then
        value = "source"
    End If
    BTK_SafeFileName = value
End Function

Private Sub BTK_WriteCsvRow(ByVal fileNo As Integer, ByRef values() As String)
    Dim i As Long
    Dim line As String

    line = values(LBound(values))
    For i = LBound(values) + 1 To UBound(values)
        line = line & "," & values(i)
    Next i
    Print #fileNo, line
End Sub

Private Function BTK_ExportRangeToCsv(ByVal sourceRange As Range, ByVal csvPath As String) As Long
    Const CHUNK_ROWS As Long = 5000
    Dim found As Range
    Dim lastRowOffset As Long
    Dim startOffset As Long
    Dim rowCount As Long
    Dim values As Variant
    Dim fileNo As Integer
    Dim r As Long
    Dim c As Long
    Dim lineValues() As String

    Set found = sourceRange.Find(What:="*", After:=sourceRange.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If found Is Nothing Then
        lastRowOffset = 1
    Else
        lastRowOffset = found.Row - sourceRange.Row + 1
        If lastRowOffset < 1 Then
            lastRowOffset = 1
        End If
    End If

    fileNo = FreeFile
    Open csvPath For Output As #fileNo
    ReDim lineValues(1 To sourceRange.Columns.Count)

    For startOffset = 1 To lastRowOffset Step CHUNK_ROWS
        rowCount = Application.WorksheetFunction.Min(CHUNK_ROWS, lastRowOffset - startOffset + 1)
        values = sourceRange.Cells(startOffset, 1).Resize(rowCount, sourceRange.Columns.Count).Value2
        If rowCount = 1 And sourceRange.Columns.Count = 1 Then
            lineValues(1) = BTK_CsvEscape(values)
            BTK_WriteCsvRow fileNo, lineValues
        Else
            For r = 1 To rowCount
                For c = 1 To sourceRange.Columns.Count
                    lineValues(c) = BTK_CsvEscape(values(r, c))
                Next c
                BTK_WriteCsvRow fileNo, lineValues
            Next r
        End If
    Next startOffset

    Close #fileNo
    BTK_ExportRangeToCsv = lastRowOffset
End Function

Private Function BTK_PrepareXLSimulationUpdateInputs(ByVal sheetName As String, ByVal outputFolder As String) As String
    Dim ws As Worksheet
    Dim sourceFolder As String
    Dim manifestPath As String
    Dim manifestNo As Integer
    Dim row As Long
    Dim causeId As String
    Dim causeFamily As String
    Dim modelSource As String
    Dim activeFlag As String
    Dim locationText As String
    Dim sourceRange As Range
    Dim csvPath As String
    Dim exportedRows As Long
    Dim manifestRow() As String

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    sourceFolder = outputFolder & "xlsimulation" & Application.PathSeparator & "source_inputs"
    BTK_EnsureFolder sourceFolder
    manifestPath = sourceFolder & Application.PathSeparator & "xlsimulation_source_manifest.csv"

    Set ws = ThisWorkbook.Worksheets(sheetName)
    manifestNo = FreeFile
    Open manifestPath For Output As #manifestNo
    ReDim manifestRow(1 To 8)

    manifestRow(1) = BTK_CsvEscape("CauseID")
    manifestRow(2) = BTK_CsvEscape("CauseFamily")
    manifestRow(3) = BTK_CsvEscape("ModelSource")
    manifestRow(4) = BTK_CsvEscape("Location")
    manifestRow(5) = BTK_CsvEscape("CsvPath")
    manifestRow(6) = BTK_CsvEscape("Rows")
    manifestRow(7) = BTK_CsvEscape("Cols")
    manifestRow(8) = BTK_CsvEscape("ExportedBy")
    BTK_WriteCsvRow manifestNo, manifestRow

    For row = 50 To 61
        causeId = Trim$(CStr(ws.Cells(row, "B").Value))
        causeFamily = Trim$(CStr(ws.Cells(row, "C").Value))
        modelSource = Trim$(CStr(ws.Cells(row, "D").Value))
        activeFlag = UCase$(Trim$(CStr(ws.Cells(row, "H").Value)))
        locationText = Trim$(CStr(ws.Cells(row, "J").Value))

        If Len(causeId) > 0 And activeFlag = "Y" Then
            If Len(locationText) = 0 Then
                Err.Raise vbObjectError + 5120, BTK_SOURCE, "XLSimulation active cause has no Location: " & causeId
            End If

            Set sourceRange = Nothing
            On Error Resume Next
            Set sourceRange = ws.Range(locationText)
            On Error GoTo 0
            If sourceRange Is Nothing Then
                Err.Raise vbObjectError + 5121, BTK_SOURCE, "XLSimulation active cause has invalid Location: " & causeId & " = " & locationText
            End If

            csvPath = sourceFolder & Application.PathSeparator & BTK_SafeFileName(causeId) & ".csv"
            exportedRows = BTK_ExportRangeToCsv(sourceRange, csvPath)

            manifestRow(1) = BTK_CsvEscape(causeId)
            manifestRow(2) = BTK_CsvEscape(causeFamily)
            manifestRow(3) = BTK_CsvEscape(modelSource)
            manifestRow(4) = BTK_CsvEscape(locationText)
            manifestRow(5) = BTK_CsvEscape(csvPath)
            manifestRow(6) = BTK_CsvEscape(CStr(exportedRows))
            manifestRow(7) = BTK_CsvEscape(CStr(sourceRange.Columns.Count))
            manifestRow(8) = BTK_CsvEscape("VBA")
            BTK_WriteCsvRow manifestNo, manifestRow

            Set sourceRange = Nothing
        End If
    Next row

    Close #manifestNo
    BTK_PrepareXLSimulationUpdateInputs = outputFolder
End Function

Private Function BTK_RefreshAggOutput(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim aggFolder As String
    Dim outputPath As String
    Dim lastRow As Long

    If Len(outputFolder) = 0 Then
        BTK_RefreshAggOutput = "Aggregate output refresh skipped: could not find output folder in R result."
        Exit Function
    End If

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    aggFolder = outputFolder & "agg" & Application.PathSeparator
    outputPath = aggFolder & "agg_output.csv"

    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, "AH").End(xlUp).Row
    If lastRow < 18 Then
        lastRow = 18
    End If

    ws.Range("AH18:AK" & lastRow).ClearContents
    BTK_LoadCsvToRange outputPath, ws.Range("AH18")

    BTK_RefreshAggOutput = "Aggregate R output refreshed on " & sheetName & "."
End Function

Private Function BTK_RefreshProfileOutput(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim profileFolder As String
    Dim outputPath As String
    Dim chartPath As String
    Dim lastRow As Long

    If Len(outputFolder) = 0 Then
        BTK_RefreshProfileOutput = "Profile output refresh skipped: could not find output folder in R result."
        Exit Function
    End If

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    profileFolder = outputFolder & "profile" & Application.PathSeparator
    outputPath = profileFolder & "profile_output.csv"
    chartPath = profileFolder & "profile_si_composition_chart.png"

    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, "BH").End(xlUp).Row
    If lastRow < 12 Then
        lastRow = 12
    End If

    ws.Range("BH12:BS" & lastRow).ClearContents
    BTK_LoadCsvToRange outputPath, ws.Range("BH12")
    BTK_InsertImage ws, chartPath, "BTK_Profile_SI_Composition_Chart", BTK_RangeFromSetting(ws, "L56", "K57"), 660, 370

    BTK_RefreshProfileOutput = "Profile R output and SI composition chart refreshed on " & sheetName & "."
End Function

Private Function BTK_RefreshCatOnLevelOutput(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim catFolder As String
    Dim outputPath As String
    Dim lastInputRow As Long
    Dim lastOutputRow As Long
    Dim lastRow As Long

    If Len(outputFolder) = 0 Then
        BTK_RefreshCatOnLevelOutput = "CatOnLevel output refresh skipped: could not find output folder in R result."
        Exit Function
    End If

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    catFolder = outputFolder & "cat_onlevel" & Application.PathSeparator
    outputPath = catFolder & "cat_onlevel_output.csv"

    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastInputRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastOutputRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastInputRow, lastOutputRow)
    If lastRow < 13 Then
        lastRow = 13
    End If

    ws.Range("T13:U" & lastRow).ClearContents
    BTK_LoadCsvToRange outputPath, ws.Range("T13")

    BTK_RefreshCatOnLevelOutput = "CatOnLevel output refreshed on " & sheetName & "."
End Function

Private Function BTK_RefreshXLSimulationGather(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim simFolder As String
    Dim coveragePath As String
    Dim breakdownPath As String

    If Len(outputFolder) = 0 Then
        BTK_RefreshXLSimulationGather = "XLSimulation gather refresh skipped: could not find output folder in R result."
        Exit Function
    End If

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    simFolder = outputFolder & "xlsimulation" & Application.PathSeparator
    coveragePath = simFolder & "xlsimulation_coverage_matrix.csv"
    breakdownPath = simFolder & "xlsimulation_layer_breakdown.csv"

    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Range("M49:V61").ClearContents
    ws.Range("AE26:AO41").ClearContents
    BTK_LoadCsvToRange coveragePath, ws.Range("M49")
    BTK_LoadCsvToRange breakdownPath, ws.Range("AE26")

    BTK_RefreshXLSimulationGather = "XLSimulation coverage matrix and breakdown headers refreshed on " & sheetName & "."
End Function

Private Function BTK_RefreshXLSimulationUpdate(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim simFolder As String
    Dim layerOutputPath As String
    Dim breakdownPath As String
    Dim oepPath As String

    If Len(outputFolder) = 0 Then
        BTK_RefreshXLSimulationUpdate = "XLSimulation update refresh skipped: could not find output folder in R result."
        Exit Function
    End If

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    simFolder = outputFolder & "xlsimulation" & Application.PathSeparator
    layerOutputPath = simFolder & "xlsimulation_layer_output.csv"
    breakdownPath = simFolder & "xlsimulation_layer_breakdown.csv"
    oepPath = simFolder & "xlsimulation_oep_output.csv"

    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Range("J26:AB41").ClearContents
    ws.Range("AE26:AO41").ClearContents
    ws.Range("B107:AW121").ClearContents
    BTK_LoadCsvToRange layerOutputPath, ws.Range("J26")
    BTK_LoadCsvToRange breakdownPath, ws.Range("AE26")
    BTK_LoadCsvToRange oepPath, ws.Range("B107")

    BTK_RefreshXLSimulationUpdate = "XLSimulation layer output, breakdown, and OEP tables refreshed on " & sheetName & "."
End Function

Private Function BTK_RefreshRiskFitCharts(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim riskfitFolder As String
    Dim finalChartPath As String
    Dim actualChartPath As String
    Dim summaryPath As String
    Dim parametersPath As String
    Dim finalTarget As Range
    Dim actualTarget As Range
    Dim refreshed As String

    If Len(outputFolder) = 0 Then
        BTK_RefreshRiskFitCharts = "RiskFit chart refresh skipped: could not find output folder in R result."
        Exit Function
    End If

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    riskfitFolder = outputFolder & "riskfit" & Application.PathSeparator
    finalChartPath = riskfitFolder & "riskfit_final_incurred_loss_chart.png"
    actualChartPath = riskfitFolder & "riskfit_actual_incurred_loss_chart.png"
    summaryPath = riskfitFolder & "riskfit_fit_summary.csv"
    parametersPath = riskfitFolder & "riskfit_fit_parameters.csv"

    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Range("CG7:DA200").ClearContents
    BTK_LoadCsvToRange summaryPath, ws.Range("CG7")
    BTK_LoadCsvToRange parametersPath, ws.Range("CV7")

    Set finalTarget = ws.Range("X61:AK91")
    Set actualTarget = ws.Range("X94:AK124")
    BTK_InsertImageFitWidth ws, finalChartPath, "BTK_RiskFit_Final_Incurred_Loss_Chart", ws.Range("X61"), finalTarget.Width
    refreshed = "RiskFit final incurred chart refreshed"

    If Len(Dir(actualChartPath)) > 0 Then
        BTK_InsertImageFitWidth ws, actualChartPath, "BTK_RiskFit_Actual_Incurred_Loss_Chart", ws.Range("X94"), actualTarget.Width
        refreshed = refreshed & "; actual incurred chart refreshed"
    Else
        BTK_DeleteShapeIfExists ws, "BTK_RiskFit_Actual_Incurred_Loss_Chart"
        refreshed = refreshed & "; actual incurred chart skipped"
    End If

    BTK_InsertRiskFitCdfChart ws, riskfitFolder, "pareto", "BTK_RiskFit_Pareto_CDF_Chart", "AQ23:AW43"
    BTK_InsertRiskFitCdfChart ws, riskfitFolder, "loggamma", "BTK_RiskFit_Loggamma_CDF_Chart", "AY23:BE43"
    BTK_InsertRiskFitCdfChart ws, riskfitFolder, "weibull", "BTK_RiskFit_Weibull_CDF_Chart", "BG23:BM43"
    BTK_InsertRiskFitCdfChart ws, riskfitFolder, "lognormal", "BTK_RiskFit_Lognormal_CDF_Chart", "BO23:BU43"
    BTK_InsertRiskFitCdfChart ws, riskfitFolder, "gamma", "BTK_RiskFit_Gamma_CDF_Chart", "BW23:CC43"

    BTK_RefreshRiskFitCharts = refreshed & "; severity fit output and CDF charts refreshed on " & sheetName & "."
End Function

Private Sub BTK_InsertRiskFitCdfChart(ByVal ws As Worksheet, ByVal riskfitFolder As String, ByVal familyKey As String, ByVal shapeName As String, ByVal targetAddress As String)
    Dim chartPath As String
    Dim target As Range

    chartPath = riskfitFolder & "riskfit_" & familyKey & "_cdf_chart.png"
    Set target = ws.Range(targetAddress)

    If Len(Dir(chartPath)) > 0 Then
        BTK_InsertImage ws, chartPath, shapeName, target.Cells(1, 1), target.Width, target.Height
    Else
        BTK_DeleteShapeIfExists ws, shapeName
    End If
End Sub

Private Function BTK_RefreshGNPICharts(ByVal outputFolder As String, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim gnpiFolder As String
    Dim levelImagePath As String
    Dim compositionImagePath As String
    Dim chartWidth As Double
    Dim levelHeight As Double
    Dim compositionHeight As Double
    Dim missing As String

    If Len(outputFolder) = 0 Then
        BTK_RefreshGNPICharts = "GNPI chart refresh skipped: could not find output folder in R result."
        Exit Function
    End If

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    gnpiFolder = outputFolder & "gnpi" & Application.PathSeparator
    levelImagePath = gnpiFolder & "gnpi_level_chart.png"
    compositionImagePath = gnpiFolder & "gnpi_lob_comparison_chart.png"

    If Len(Dir(levelImagePath)) = 0 Then
        missing = missing & vbCrLf & levelImagePath
    End If
    If Len(Dir(compositionImagePath)) = 0 Then
        missing = missing & vbCrLf & compositionImagePath
    End If

    If Len(missing) > 0 Then
        BTK_RefreshGNPICharts = "GNPI chart image files were not visible to Excel:" & missing
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(sheetName)
    chartWidth = BTK_NumberFromSetting(ws, "L30", 560)
    levelHeight = BTK_NumberFromSetting(ws, "L31", 285)
    compositionHeight = BTK_NumberFromSetting(ws, "L32", 285)

    BTK_DeleteChartObjects ws
    BTK_DeleteShapeIfExists ws, "Pic_GNPI_Level"
    BTK_DeleteShapeIfExists ws, "Pic_GNPI_Breakdown"
    BTK_InsertImage ws, levelImagePath, "BTK_GNPI_Level_Chart", BTK_RangeFromSetting(ws, "L34", "O18"), chartWidth, levelHeight
    BTK_InsertImage ws, compositionImagePath, "BTK_GNPI_LOB_Comparison_Chart", BTK_RangeFromSetting(ws, "L35", "O43"), chartWidth, compositionHeight

    BTK_RefreshGNPICharts = "GNPI R chart images refreshed on " & sheetName & "."
End Function

Private Function GRe_ObjectiveText() As String
    GRe_ObjectiveText = Trim$(CStr(ActiveSheet.Range("A1").Value))
End Function

Private Function GRe_ObjectiveKey(ByVal objectiveText As String) As String
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
    GRe_ObjectiveKey = key
End Function

Private Function GRe_ToolIdForObjective(ByVal objectiveText As String) As String
    Select Case GRe_ObjectiveKey(objectiveText)
        Case "exposure map", "china exposure map"
            GRe_ToolIdForObjective = "china_exposure_map"
        Case "xl pricing", "xl pricing tool", "xol pricing"
            GRe_ToolIdForObjective = "xl_pricing_tool"
        Case "gnpi"
            GRe_ToolIdForObjective = "xl_pricing_tool"
        Case "agg"
            GRe_ToolIdForObjective = "xl_pricing_tool"
        Case "profile"
            GRe_ToolIdForObjective = "xl_pricing_tool"
        Case "riskfit"
            GRe_ToolIdForObjective = "xl_pricing_tool"
        Case "catonlevel", "cat onlevel", "cat on level"
            GRe_ToolIdForObjective = "xl_pricing_tool"
        Case "xlsimulation", "xl simulation", "xl simulations", "sim variations"
            GRe_ToolIdForObjective = "xl_pricing_tool"
        Case "curve fit risk"
            GRe_ToolIdForObjective = "curve_fit_risk"
        Case Else
            GRe_ToolIdForObjective = ""
    End Select
End Function

Private Function GRe_ActionForObjective(ByVal objectiveText As String, ByVal action As String) As String
    Select Case GRe_ObjectiveKey(objectiveText)
        Case "gnpi"
            GRe_ActionForObjective = "gnpi_" & LCase$(action)
        Case "agg"
            GRe_ActionForObjective = "agg_" & LCase$(action)
        Case "profile"
            GRe_ActionForObjective = "profile_" & LCase$(action)
        Case "riskfit"
            GRe_ActionForObjective = "riskfit_" & LCase$(action)
        Case "catonlevel", "cat onlevel", "cat on level"
            GRe_ActionForObjective = "cat_onlevel_" & LCase$(action)
        Case "xlsimulation", "xl simulation", "xl simulations", "sim variations"
            GRe_ActionForObjective = "xlsimulation_" & LCase$(action)
        Case Else
            GRe_ActionForObjective = LCase$(action)
    End Select
End Function

Private Function BTK_DispatchTool(ByVal toolId As String, ByVal action As String, Optional ByVal outputDir As String = "", Optional ByVal activeSheetName As String = "") As Variant
    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 5100, BTK_SOURCE, "Please save the workbook before running a GRe tool action."
    End If

    ThisWorkbook.Save
    BTK_SourceBootstrap

    BTK_DispatchTool = BTK_CallR( _
        "BTK.DispatchTool", _
            toolId, _
            action, _
            ThisWorkbook.FullName, _
            outputDir, _
            activeSheetName _
    )
End Function

Private Sub GRe_HandleResult(ByVal toolId As String, ByVal action As String, ByVal result As Variant, ByVal refreshWorkbook As Boolean, ByVal activeSheetName As String)
    Dim resultText As String
    Dim outputFolder As String

    If IsError(result) Then
        MsgBox "BERT returned an Excel error before returning a text result." & vbCrLf & _
               "Run BTK_CheckInstall, then try again. If this persists, re-source bootstrap.R and re-import BERTToolkitClient.bas.", _
               vbCritical, "GRe Tools"
        Exit Sub
    End If

    resultText = CStr(result)
    If Left$(resultText, 6) = "ERROR:" Then
        MsgBox resultText, vbCritical, "GRe Tools"
        Exit Sub
    End If

    If refreshWorkbook And toolId = "china_exposure_map" And action = "update" Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        BTK_RefreshMapImages outputFolder
        ThisWorkbook.Save
        resultText = resultText & vbCrLf & vbCrLf & "Workbook images refreshed on Parameters sheet."
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "gnpi_update" And InStr(1, resultText, "GNPI update completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshGNPICharts(outputFolder, activeSheetName)
        ThisWorkbook.Save
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "agg_update" And InStr(1, resultText, "Aggregate update completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshAggOutput(outputFolder, activeSheetName)
        ThisWorkbook.Save
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "profile_update" And InStr(1, resultText, "Profile update completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshProfileOutput(outputFolder, activeSheetName)
        ThisWorkbook.Save
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "riskfit_update" And InStr(1, resultText, "RiskFit update completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshRiskFitCharts(outputFolder, activeSheetName)
        ThisWorkbook.Save
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "cat_onlevel_update" And InStr(1, resultText, "CatOnLevel update completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshCatOnLevelOutput(outputFolder, activeSheetName)
        ThisWorkbook.Save
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "xlsimulation_gather" And InStr(1, resultText, "XLSimulation gather completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshXLSimulationGather(outputFolder, activeSheetName)
        ThisWorkbook.Save
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "xlsimulation_update" And InStr(1, resultText, "XLSimulation update completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshXLSimulationUpdate(outputFolder, activeSheetName)
        ThisWorkbook.Save
    End If

    MsgBox resultText, vbInformation, "GRe Tools"
End Sub

Public Sub GRe_Dispatch(ByVal action As String)
    Dim objectiveText As String
    Dim toolId As String
    Dim actionName As String
    Dim result As Variant
    Dim activeSheetName As String
    Dim outputDir As String

    On Error GoTo DispatchFailed

    activeSheetName = ActiveSheet.Name
    objectiveText = GRe_ObjectiveText()
    If Len(objectiveText) = 0 Then
        MsgBox "no definable action for this tab", vbInformation, "GRe Tools"
        Exit Sub
    End If

    toolId = GRe_ToolIdForObjective(objectiveText)
    If Len(toolId) = 0 Then
        MsgBox "No GRe tool is registered for active sheet objective: " & objectiveText, vbInformation, "GRe Tools"
        Exit Sub
    End If

    actionName = GRe_ActionForObjective(objectiveText, action)
    outputDir = ""
    If toolId = "xl_pricing_tool" And actionName = "xlsimulation_update" Then
        outputDir = BTK_PrepareXLSimulationUpdateInputs(activeSheetName, BTK_DefaultRunDir(toolId))
    End If

    result = BTK_DispatchTool(toolId, actionName, outputDir, activeSheetName)
    GRe_HandleResult toolId, actionName, result, True, activeSheetName
    Exit Sub

DispatchFailed:
    MsgBox Err.Description, vbCritical, "GRe Tools"
End Sub

Public Sub GRe_Gather()
    GRe_Dispatch "gather"
End Sub

Public Sub GRe_Update()
    GRe_Dispatch "update"
End Sub

Public Sub GRe_Build()
    GRe_Dispatch "build"
End Sub

Public Sub GRe_RibbonGather(ByVal control As Object)
    GRe_Gather
End Sub

Public Sub GRe_RibbonUpdate(ByVal control As Object)
    GRe_Update
End Sub

Public Sub GRe_RibbonBuild(ByVal control As Object)
    GRe_Build
End Sub

Public Sub GRe_RibbonGatherFallback()
    GRe_Gather
End Sub

Public Sub GRe_RibbonUpdateFallback()
    GRe_Update
End Sub

Public Sub GRe_RibbonBuildFallback()
    GRe_Build
End Sub

Public Function BTK_Version() As Variant
    Dim version As Variant

    version = BTK_CallR("BTK.Version")
    If IsError(version) Then
        BTK_SourceBootstrap
        version = BTK_CallR("BTK.Version")
    End If

    If IsError(version) Then
        Err.Raise vbObjectError + 5104, BTK_SOURCE, _
            BTK_InstallHelp() & vbCrLf & vbCrLf & _
            "BERT returned an Excel error for BTK.Version."
    End If

    BTK_Version = version
End Function

Public Sub BTK_CheckInstall()
    Dim version As Variant
    Dim missingPackages As Variant

    On Error GoTo CheckFailed
    BTK_SourceBootstrap
    version = BTK_Version()
    missingPackages = BTK_CallR("BTK.MissingPackages")
    MsgBox "BERT Toolkit is available." & vbCrLf & _
           "Version: " & CStr(version) & vbCrLf & _
           "Missing packages: " & CStr(missingPackages), vbInformation, "BERT Toolkit"
    Exit Sub

CheckFailed:
    MsgBox Err.Description, vbExclamation, "BERT Toolkit"
End Sub

Public Sub BTK_InstallPackages()
    Dim result As Variant

    On Error GoTo InstallFailed
    BTK_SourceBootstrap
    result = BTK_CallR("BTK.InstallPackages")

    If IsError(result) Then
        MsgBox "BERT returned an Excel error while installing packages.", vbCritical, "BERT Toolkit"
    Else
        MsgBox CStr(result), vbInformation, "BERT Toolkit"
    End If
    Exit Sub

InstallFailed:
    MsgBox Err.Description, vbCritical, "BERT Toolkit"
End Sub

Public Function BTK_RunTool(ByVal toolId As String, Optional ByVal outputDir As String = "") As Variant
    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 5100, BTK_SOURCE, "Please save the workbook before running a BERT toolkit function."
    End If

    ThisWorkbook.Save

    BTK_SourceBootstrap

    If IsError(BTK_Version()) Then
        Err.Raise vbObjectError + 5105, BTK_SOURCE, "BERT Toolkit is not loaded."
    End If

    BTK_RunTool = BTK_CallR( _
        "BTK.RunTool", _
        toolId, _
        ThisWorkbook.FullName, _
        outputDir _
    )
End Function

Public Sub BTK_RunChinaExposureMap()
    Dim result As Variant
    Dim activeSheetName As String

    activeSheetName = ActiveSheet.Name
    result = BTK_DispatchTool("china_exposure_map", "update", "", activeSheetName)
    GRe_HandleResult "china_exposure_map", "update", result, True, activeSheetName
End Sub
