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

Private Function BTK_RefreshAggOutput(ByVal outputFolder As String) As String
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

    Set ws = ThisWorkbook.Worksheets("Input_Agg")
    lastRow = ws.Cells(ws.Rows.Count, "AH").End(xlUp).Row
    If lastRow < 18 Then
        lastRow = 18
    End If

    ws.Range("AH18:AK" & lastRow).ClearContents
    BTK_LoadCsvToRange outputPath, ws.Range("AH18")

    BTK_RefreshAggOutput = "Aggregate R output refreshed on Input_Agg."
End Function

Private Function BTK_RefreshProfileOutput(ByVal outputFolder As String) As String
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

    Set ws = ThisWorkbook.Worksheets("Input_Profile")
    lastRow = ws.Cells(ws.Rows.Count, "BH").End(xlUp).Row
    If lastRow < 12 Then
        lastRow = 12
    End If

    ws.Range("BH12:BS" & lastRow).ClearContents
    BTK_LoadCsvToRange outputPath, ws.Range("BH12")
    BTK_InsertImage ws, chartPath, "BTK_Profile_SI_Composition_Chart", BTK_RangeFromSetting(ws, "L56", "K57"), 660, 370

    BTK_RefreshProfileOutput = "Profile R output and SI composition chart refreshed on Input_Profile."
End Function

Private Function BTK_RefreshRiskFitCharts(ByVal outputFolder As String) As String
    Dim ws As Worksheet
    Dim riskfitFolder As String
    Dim finalChartPath As String
    Dim actualChartPath As String
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

    Set ws = ThisWorkbook.Worksheets("Risk Loss Fitting")
    Set finalTarget = ws.Range("R54:AK75")
    Set actualTarget = ws.Range("R77:AK98")
    BTK_InsertImage ws, finalChartPath, "BTK_RiskFit_Final_Incurred_Loss_Chart", finalTarget.Cells(1, 1), finalTarget.Width, finalTarget.Height
    refreshed = "RiskFit final incurred chart refreshed"

    If Len(Dir(actualChartPath)) > 0 Then
        BTK_InsertImage ws, actualChartPath, "BTK_RiskFit_Actual_Incurred_Loss_Chart", actualTarget.Cells(1, 1), actualTarget.Width, actualTarget.Height
        refreshed = refreshed & "; actual incurred chart refreshed"
    Else
        BTK_DeleteShapeIfExists ws, "BTK_RiskFit_Actual_Incurred_Loss_Chart"
        refreshed = refreshed & "; actual incurred chart skipped"
    End If

    BTK_RefreshRiskFitCharts = refreshed & " on Risk Loss Fitting."
End Function

Private Function BTK_RefreshGNPICharts(ByVal outputFolder As String) As String
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

    Set ws = ActiveSheet
    chartWidth = BTK_NumberFromSetting(ws, "L30", 560)
    levelHeight = BTK_NumberFromSetting(ws, "L31", 285)
    compositionHeight = BTK_NumberFromSetting(ws, "L32", 285)

    BTK_DeleteChartObjects ws
    BTK_DeleteShapeIfExists ws, "Pic_GNPI_Level"
    BTK_DeleteShapeIfExists ws, "Pic_GNPI_Breakdown"
    BTK_InsertImage ws, levelImagePath, "BTK_GNPI_Level_Chart", BTK_RangeFromSetting(ws, "L34", "O18"), chartWidth, levelHeight
    BTK_InsertImage ws, compositionImagePath, "BTK_GNPI_LOB_Comparison_Chart", BTK_RangeFromSetting(ws, "L35", "O43"), chartWidth, compositionHeight

    BTK_RefreshGNPICharts = "GNPI R chart images refreshed."
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
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshGNPICharts(outputFolder)
        ThisWorkbook.Save
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "agg_update" And InStr(1, resultText, "Aggregate update completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshAggOutput(outputFolder)
        ThisWorkbook.Save
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "profile_update" And InStr(1, resultText, "Profile update completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshProfileOutput(outputFolder)
        ThisWorkbook.Save
    End If

    If refreshWorkbook And toolId = "xl_pricing_tool" And action = "riskfit_update" And InStr(1, resultText, "RiskFit update completed.", vbTextCompare) > 0 Then
        outputFolder = BTK_OutputFolderFromResult(resultText)
        resultText = resultText & vbCrLf & vbCrLf & BTK_RefreshRiskFitCharts(outputFolder)
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
