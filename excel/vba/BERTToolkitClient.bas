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
    Dim picture As Shape

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
    Dim picture As Shape

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

Private Function BTK_GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set BTK_GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If BTK_GetOrCreateWorksheet Is Nothing Then
        Set BTK_GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        BTK_GetOrCreateWorksheet.Name = sheetName
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

Private Function BTK_ConfigValue(ByVal helperWs As Worksheet, ByVal key As String, ByVal defaultValue As String) As String
    Dim found As Range

    Set found = helperWs.Range("P:P").Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If found Is Nothing Then
        BTK_ConfigValue = defaultValue
    Else
        BTK_ConfigValue = CStr(found.Offset(0, 1).Value)
        If Len(Trim$(BTK_ConfigValue)) = 0 Then
            BTK_ConfigValue = defaultValue
        End If
    End If
End Function

Private Function BTK_ConfigNumber(ByVal helperWs As Worksheet, ByVal key As String, ByVal defaultValue As Double) As Double
    Dim rawValue As String

    rawValue = BTK_ConfigValue(helperWs, key, CStr(defaultValue))
    If IsNumeric(rawValue) And CDbl(rawValue) > 0 Then
        BTK_ConfigNumber = CDbl(rawValue)
    Else
        BTK_ConfigNumber = defaultValue
    End If
End Function

Private Function BTK_ColorFromHex(ByVal hexColor As String, ByVal fallbackColor As Long) As Long
    Dim cleanColor As String

    cleanColor = Trim$(hexColor)
    cleanColor = Replace(cleanColor, "#", "")
    If Len(cleanColor) <> 6 Then
        BTK_ColorFromHex = fallbackColor
        Exit Function
    End If

    On Error GoTo BadColor
    BTK_ColorFromHex = RGB( _
        CLng("&H" & Mid$(cleanColor, 1, 2)), _
        CLng("&H" & Mid$(cleanColor, 3, 2)), _
        CLng("&H" & Mid$(cleanColor, 5, 2)) _
    )
    Exit Function

BadColor:
    BTK_ColorFromHex = fallbackColor
End Function

Private Function BTK_ListItem(ByVal listText As String, ByVal itemIndex As Long, ByVal defaultValue As String) As String
    Dim parts() As String

    If Len(Trim$(listText)) = 0 Then
        BTK_ListItem = defaultValue
        Exit Function
    End If

    parts = Split(listText, ",")
    If itemIndex - 1 <= UBound(parts) Then
        BTK_ListItem = Trim$(parts(itemIndex - 1))
    Else
        BTK_ListItem = defaultValue
    End If

    If Len(BTK_ListItem) = 0 Then
        BTK_ListItem = defaultValue
    End If
End Function

Private Function BTK_DefaultLOBColor(ByVal itemIndex As Long) As String
    Dim defaults As Variant

    defaults = Array("#A7C7DC", "#3F7FBF", "#7EAA92", "#D9A441", "#C96C5A", "#8E7CC3", "#6FA8DC", "#93C47D")
    BTK_DefaultLOBColor = defaults((itemIndex - 1) Mod (UBound(defaults) + 1))
End Function

Private Sub BTK_FormatNativeChart(ByVal chartObj As ChartObject, ByVal titleText As String, ByVal subtitleText As String, ByVal titleFontSize As Double, ByVal axisFontSize As Double)
    With chartObj.Chart
        .HasTitle = True
        .ChartTitle.Text = titleText & IIf(Len(subtitleText) > 0, vbLf & subtitleText, "")
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = titleFontSize
        .Legend.Position = xlLegendPositionTop
        .Axes(xlCategory).TickLabels.Font.Size = axisFontSize
        .Axes(xlValue).TickLabels.Font.Size = axisFontSize
    End With
End Sub

Private Sub BTK_BuildGNPINativeLevelTable(ByVal helperWs As Worksheet, ByVal lastRawRow As Long)
    Dim rowIndex As Long
    Dim targetRow As Long
    Dim sectionType As String
    Dim amount As Double

    helperWs.Range("S:V").Clear
    helperWs.Range("S1:V1").Value = Array("Label", "Actual", "Revised", "Estimate")

    targetRow = 2
    For rowIndex = 2 To lastRawRow
        sectionType = CStr(helperWs.Cells(rowIndex, "B").Value)
        amount = CDbl(helperWs.Cells(rowIndex, "D").Value) / 1000000#

        helperWs.Cells(targetRow, "S").Value = helperWs.Cells(rowIndex, "C").Value
        Select Case LCase$(sectionType)
            Case "actual"
                helperWs.Cells(targetRow, "T").Value = amount
                helperWs.Cells(targetRow, "U").Formula = "=NA()"
                helperWs.Cells(targetRow, "V").Formula = "=NA()"
            Case "revised"
                helperWs.Cells(targetRow, "T").Formula = "=NA()"
                helperWs.Cells(targetRow, "U").Value = amount
                helperWs.Cells(targetRow, "V").Formula = "=NA()"
            Case "estimate"
                helperWs.Cells(targetRow, "T").Formula = "=NA()"
                helperWs.Cells(targetRow, "U").Formula = "=NA()"
                helperWs.Cells(targetRow, "V").Value = amount
        End Select
        targetRow = targetRow + 1
    Next rowIndex
End Sub

Private Sub BTK_BuildGNPINativeCompositionTable(ByVal helperWs As Worksheet, ByVal firstRawRow As Long, ByVal lastRawRow As Long, ByRef lastHelperRow As Long, ByRef lastHelperCol As Long)
    Dim labels As Object
    Dim lobs As Object
    Dim rowIndex As Long
    Dim labelText As String
    Dim lobText As String
    Dim labelIndex As Long
    Dim lobIndex As Long
    Dim key As Variant

    Set labels = CreateObject("Scripting.Dictionary")
    Set lobs = CreateObject("Scripting.Dictionary")

    helperWs.Range("X:AJ").Clear

    For rowIndex = firstRawRow To lastRawRow
        labelText = CStr(helperWs.Cells(rowIndex, "J").Value)
        lobText = CStr(helperWs.Cells(rowIndex, "K").Value)
        If Len(labelText) > 0 And Not labels.Exists(labelText) Then labels.Add labelText, labels.Count + 1
        If Len(lobText) > 0 And Not lobs.Exists(lobText) Then lobs.Add lobText, lobs.Count + 1
    Next rowIndex

    helperWs.Cells(1, "X").Value = "Label"
    For Each key In labels.Keys
        helperWs.Cells(labels(key) + 1, "X").Value = key
    Next key

    For Each key In lobs.Keys
        helperWs.Cells(1, 24 + lobs(key)).Value = key
    Next key

    For rowIndex = firstRawRow To lastRawRow
        labelText = CStr(helperWs.Cells(rowIndex, "J").Value)
        lobText = CStr(helperWs.Cells(rowIndex, "K").Value)
        If labels.Exists(labelText) And lobs.Exists(lobText) Then
            labelIndex = labels(labelText) + 1
            lobIndex = 24 + lobs(lobText)
            helperWs.Cells(labelIndex, lobIndex).Value = CDbl(helperWs.Cells(rowIndex, "M").Value)
        End If
    Next rowIndex

    lastHelperRow = labels.Count + 1
    lastHelperCol = 24 + lobs.Count
End Sub

Private Function BTK_RefreshGNPINativeCharts(ByVal outputFolder As String) As String
    Dim ws As Worksheet
    Dim helperWs As Worksheet
    Dim gnpiFolder As String
    Dim levelDataPath As String
    Dim compositionDataPath As String
    Dim configPath As String
    Dim chartWidth As Double
    Dim levelHeight As Double
    Dim compositionHeight As Double
    Dim levelLastRow As Long
    Dim compositionLastRow As Long
    Dim compositionHelperLastRow As Long
    Dim compositionHelperLastCol As Long
    Dim levelChart As ChartObject
    Dim compositionChart As ChartObject
    Dim seriesIndex As Long
    Dim lobPalette As String

    If Len(outputFolder) = 0 Then
        BTK_RefreshGNPINativeCharts = "GNPI chart refresh skipped: could not find output folder in R result."
        Exit Function
    End If

    outputFolder = Replace(outputFolder, "/", Application.PathSeparator)
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    gnpiFolder = outputFolder & "gnpi" & Application.PathSeparator
    levelDataPath = gnpiFolder & "gnpi_level_data.csv"
    compositionDataPath = gnpiFolder & "gnpi_lob_comparison_data.csv"
    configPath = gnpiFolder & "gnpi_chart_config.csv"

    Set ws = ActiveSheet
    Set helperWs = BTK_GetOrCreateWorksheet("_BTK_GNPI_Data")
    helperWs.Visible = xlSheetVisible
    helperWs.Cells.Clear

    BTK_LoadCsvToRange levelDataPath, helperWs.Range("A1")
    BTK_LoadCsvToRange compositionDataPath, helperWs.Range("H1")
    BTK_LoadCsvToRange configPath, helperWs.Range("P1")

    levelLastRow = helperWs.Cells(helperWs.Rows.Count, "A").End(xlUp).Row
    compositionLastRow = helperWs.Cells(helperWs.Rows.Count, "H").End(xlUp).Row

    BTK_BuildGNPINativeLevelTable helperWs, levelLastRow
    BTK_BuildGNPINativeCompositionTable helperWs, 2, compositionLastRow, compositionHelperLastRow, compositionHelperLastCol

    chartWidth = BTK_NumberFromSetting(ws, "L30", 560)
    levelHeight = BTK_NumberFromSetting(ws, "L31", 285)
    compositionHeight = BTK_NumberFromSetting(ws, "L32", 285)

    BTK_DeleteChartObjects ws
    BTK_DeleteShapeIfExists ws, "BTK_GNPI_Level_Chart"
    BTK_DeleteShapeIfExists ws, "BTK_GNPI_LOB_Comparison_Chart"

    Set levelChart = ws.ChartObjects.Add( _
        Left:=BTK_RangeFromSetting(ws, "L34", "O18").Left, _
        Top:=BTK_RangeFromSetting(ws, "L34", "O18").Top, _
        Width:=chartWidth, _
        Height:=levelHeight _
    )
    levelChart.Name = "BTK_GNPI_Level_Chart"
    levelChart.Chart.ChartType = xlColumnClustered
    levelChart.Chart.SetSourceData helperWs.Range("S1:V" & levelLastRow)
    levelChart.Chart.Axes(xlValue).TickLabels.NumberFormat = "#,##0M"
    levelChart.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = BTK_ColorFromHex(BTK_ConfigValue(helperWs, "Actual Color", "#4F9FBC"), RGB(79, 159, 188))
    levelChart.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = BTK_ColorFromHex(BTK_ConfigValue(helperWs, "Revised Color", "#3A7F72"), RGB(58, 127, 114))
    levelChart.Chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = BTK_ColorFromHex(BTK_ConfigValue(helperWs, "Estimate Color", "#8EC5D6"), RGB(142, 197, 214))
    BTK_FormatNativeChart levelChart, _
        BTK_ConfigValue(helperWs, "Chart 1 Title", "GNPI by Treaty Year and Type"), _
        BTK_ConfigValue(helperWs, "Chart 1 Subtitle", "Total GNPI by year and type"), _
        BTK_ConfigNumber(helperWs, "Title Font Size", 16), _
        BTK_ConfigNumber(helperWs, "Axis Font Size", 9)

    Set compositionChart = ws.ChartObjects.Add( _
        Left:=BTK_RangeFromSetting(ws, "L35", "O43").Left, _
        Top:=BTK_RangeFromSetting(ws, "L35", "O43").Top, _
        Width:=chartWidth, _
        Height:=compositionHeight _
    )
    compositionChart.Name = "BTK_GNPI_LOB_Comparison_Chart"
    compositionChart.Chart.ChartType = xlColumnStacked100
    compositionChart.Chart.SetSourceData helperWs.Range(helperWs.Cells(1, "X"), helperWs.Cells(compositionHelperLastRow, compositionHelperLastCol))
    compositionChart.Chart.Axes(xlValue).TickLabels.NumberFormat = "0%"
    lobPalette = BTK_ConfigValue(helperWs, "LOB Palette", "")
    For seriesIndex = 1 To compositionChart.Chart.SeriesCollection.Count
        compositionChart.Chart.SeriesCollection(seriesIndex).Format.Fill.ForeColor.RGB = BTK_ColorFromHex( _
            BTK_ListItem(lobPalette, seriesIndex, BTK_DefaultLOBColor(seriesIndex)), _
            RGB(120, 160, 190) _
        )
    Next seriesIndex
    BTK_FormatNativeChart compositionChart, _
        BTK_ConfigValue(helperWs, "Chart 2 Title", "LOB Composition by Treaty Year"), _
        BTK_ConfigValue(helperWs, "Chart 2 Subtitle", "Stacked percentage composition by year/type section from configured start year"), _
        BTK_ConfigNumber(helperWs, "Title Font Size", 16), _
        BTK_ConfigNumber(helperWs, "Axis Font Size", 9)

    helperWs.Visible = xlSheetVeryHidden
    BTK_RefreshGNPINativeCharts = "GNPI native Excel charts refreshed."
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
        Case Else
            GRe_ActionForObjective = LCase$(action)
    End Select
End Function

Private Function BTK_DispatchTool(ByVal toolId As String, ByVal action As String, Optional ByVal outputDir As String = "") As Variant
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
            outputDir _
    )
End Function

Private Sub GRe_HandleResult(ByVal toolId As String, ByVal action As String, ByVal result As Variant, ByVal refreshWorkbook As Boolean)
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
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshGNPINativeCharts(outputFolder)
        ThisWorkbook.Save
    End If

    MsgBox resultText, vbInformation, "GRe Tools"
End Sub

Public Sub GRe_Dispatch(ByVal action As String)
    Dim objectiveText As String
    Dim toolId As String
    Dim actionName As String
    Dim result As Variant

    On Error GoTo DispatchFailed

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
    result = BTK_DispatchTool(toolId, actionName, "")
    GRe_HandleResult toolId, actionName, result, True
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

Public Sub GRe_RibbonGather(ByVal control As IRibbonControl)
    GRe_Gather
End Sub

Public Sub GRe_RibbonUpdate(ByVal control As IRibbonControl)
    GRe_Update
End Sub

Public Sub GRe_RibbonBuild(ByVal control As IRibbonControl)
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

    result = BTK_DispatchTool("china_exposure_map", "update", "")
    GRe_HandleResult "china_exposure_map", "update", result, True
End Sub
