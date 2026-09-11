Attribute VB_Name = "Output"
Option Explicit

'=====================================================================================
' Output module, v1.2.0  (replaces the v1.1.2 Output module)
'
' SetupSummaryWs is still the entry point called from frmAlternativeSetup.cmdClose_Click.
' It now writes five linked blocks to the Summary worksheet and builds four charts:
'   1. Results table: initial cost, PW by category, NPW, delta to lowest, closure days
'   2. Present worth by category (stacked column, one per alternative)
'   3. Expenditure stream by calendar year (clustered column, undiscounted)
'   4. Cumulative discounted cost by calendar year (line)
'   5. Discount-rate sensitivity 2% to 8% (line) with the analysis rate marked
' Everything on the sheet is a live formula into the alternative worksheets, so the
' Summary updates when quantities change. Charts are rebuilt each time the form closes.
'
' To install: in the VBA editor remove the existing "Output" module, then
' File > Import File... > Output.bas. No other code changes are needed.
'=====================================================================================

Private Const ROW_HDR As Long = 3          'header row of the results table
Private Const ROW_FIRST As Long = 4        'first alternative row
Private Const RATE_MIN As Double = 2
Private Const RATE_MAX As Double = 8
Private Const RATE_STEP As Double = 0.5

'Series colors: one hue per pavement type, tinted for a second alternative of the same type
Private Function AltColor(altType As Integer, nth As Long) As Long
    Select Case altType
        Case EnumAltType.NewHMA
            If nth = 0 Then AltColor = RGB(42, 120, 214) Else AltColor = RGB(109, 167, 236)
        Case EnumAltType.NewPCC
            If nth = 0 Then AltColor = RGB(235, 104, 52) Else AltColor = RGB(243, 156, 122)
        Case Else
            AltColor = RGB(137, 135, 129)
    End Select
End Function

Public Sub SetupSummaryWs()
'This sub writes data to the summary worksheet and rebuilds its charts
    Dim ws As Worksheet
    Dim curAltObj As clsAlternative
    Dim numAlts As Long, i As Long, r As Long, k As Long
    Dim wsAlt As Worksheet, altName As String
    Dim rowInit As Long, rowNPW As Long, rowTotal As Long, rowFirstAct As Long, rowLastAct As Long
    Dim q As String
    Dim rngB As String, rngC As String, rngD As String, rngE As String
    Dim rowStream As Long, rowCum As Long, rowSens As Long, rowCat As Long
    Dim typeCount(0 To 3) As Long
    Dim colorOf() As Long
    Dim lastSummaryRow As Long

    Set ws = GlobalVars.wsSummary
    numAlts = MainDoc.AlternativeWsCollection.Count
    If numAlts = 0 Then
        ws.Range("A4:Z400").Clear
        Exit Sub
    End If
    ReDim colorOf(1 To numAlts)

    Application.ScreenUpdating = False
    On Error GoTo CleanExit

    'Clear previous output (rows 4 and below, plus all charts)
    ws.Range("A4:Z400").Clear
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    '---------------------------------------------------------------- 1. Results table
    ws.Range("A3:L3").Value = Array("Alternative", "Name", "Type", "Initial Construction", _
        "Maintenance PW", "Rehabilitation PW", "Lost Revenue PW", "Salvage PW", _
        "Net Present Worth", "vs. Lowest", "Closure Days (30 yr)", "Alternative Description")
    ws.Range("A3:L3").Font.Bold = True
    ws.Range("A3:L3").WrapText = True

    r = ROW_FIRST
    i = 0
    For Each curAltObj In MainDoc.AlternativeWsCollection
        i = i + 1
        Set wsAlt = Worksheets(curAltObj.NameWsAlt)
        altName = "'" & curAltObj.NameWsAlt & "'!"

        'Locate the NPW table on the alternative worksheet (layout differs by type)
        rowInit = wsAlt.Columns("B").Find(What:="Initial Construction", LookIn:=xlValues, LookAt:=xlWhole).Row
        rowNPW = wsAlt.Columns("B").Find(What:="Net Present Worth", LookIn:=xlValues, LookAt:=xlWhole).Row
        rowTotal = wsAlt.Columns("A").Find(What:="Total", LookIn:=xlValues, LookAt:=xlWhole).Row
        rowFirstAct = rowInit + 1
        rowLastAct = rowNPW - 1                      'salvage row
        rngB = altName & "$B$" & rowFirstAct & ":$B$" & rowLastAct
        rngC = altName & "$C$" & rowFirstAct & ":$C$" & rowLastAct
        rngD = altName & "$D$" & rowFirstAct & ":$D$" & rowLastAct
        rngE = altName & "$E$" & rowFirstAct & ":$E$" & rowLastAct

        ws.Cells(r, 1).Value = "Alt " & i
        ws.Cells(r, 2).Value = curAltObj.Name
        ws.Cells(r, 3).Value = GetEnumStringFromInt_AltTypeDispString(curAltObj.AltType)
        ws.Cells(r, 4).Formula = "=" & altName & "$G$" & rowTotal
        ws.Cells(r, 5).Formula = "=SUMIFS(" & rngE & "," & rngB & ",""Maintenance*""," & rngB & ",""<>*Indirect*"")"
        ws.Cells(r, 6).Formula = "=SUMIFS(" & rngE & "," & rngB & ",""Rehabilitation*""," & rngB & ",""<>*Indirect*"")"
        ws.Cells(r, 7).Formula = "=SUMIFS(" & rngE & "," & rngB & ",""*Indirect*"")"
        ws.Cells(r, 8).Formula = "=SUMIFS(" & rngE & "," & rngB & ",""Salvage*"")"
        ws.Cells(r, 9).Formula = "=" & altName & "$E$" & rowNPW
        ws.Cells(r, 11).Formula = "=SUM(" & altName & "$F$4:$F$10)"
        ws.Cells(r, 12).Value = curAltObj.Description

        colorOf(i) = AltColor(curAltObj.AltType, typeCount(curAltObj.AltType))
        typeCount(curAltObj.AltType) = typeCount(curAltObj.AltType) + 1
        r = r + 1
    Next curAltObj
    lastSummaryRow = r - 1

    'Delta to the lowest NPW, written after the loop so the MIN range is known
    For r = ROW_FIRST To lastSummaryRow
        ws.Cells(r, 10).Formula = "=I" & r & "-MIN($I$" & ROW_FIRST & ":$I$" & lastSummaryRow & ")"
    Next r
    ws.Range(ws.Cells(ROW_FIRST, 4), ws.Cells(lastSummaryRow, 10)).NumberFormat = "$#,##0;($#,##0);-"
    ws.Range(ws.Cells(ROW_FIRST, 11), ws.Cells(lastSummaryRow, 11)).NumberFormat = "0"

    'Verdict line
    r = lastSummaryRow + 1
    ws.Cells(r, 2).Formula = "=""Lowest NPW: ""&INDEX($B$" & ROW_FIRST & ":$B$" & lastSummaryRow & ",MATCH(MIN($I$" & ROW_FIRST & ":$I$" & lastSummaryRow & "),$I$" & ROW_FIRST & ":$I$" & lastSummaryRow & ",0))&""  |  Analysis: ""&'General Information'!$D$33&"" yr at ""&'General Information'!$D$34&""%  |  Lost revenue: ""&'General Information'!$D$38"
    ws.Cells(r, 2).Font.Italic = True

    '---------------------------------------------------------------- 2. PW by category (helper block)
    rowCat = lastSummaryRow + 3
    ws.Cells(rowCat, 1).Value = "Present worth by category"
    ws.Cells(rowCat, 1).Font.Bold = True
    ws.Cells(rowCat + 1, 1).Value = "Category"
    For i = 1 To numAlts
        ws.Cells(rowCat + 1, 1 + i).Formula = "=$B$" & (ROW_FIRST + i - 1)
    Next i
    Dim catNames As Variant, catCols As Variant
    catNames = Array("Initial construction", "Maintenance", "Rehabilitation", "Lost revenue", "Salvage")
    catCols = Array("D", "E", "F", "G", "H")
    For k = 0 To 4
        ws.Cells(rowCat + 2 + k, 1).Value = catNames(k)
        For i = 1 To numAlts
            ws.Cells(rowCat + 2 + k, 1 + i).Formula = "=$" & catCols(k) & "$" & (ROW_FIRST + i - 1)
        Next i
    Next k
    ws.Range(ws.Cells(rowCat + 2, 2), ws.Cells(rowCat + 6, 1 + numAlts)).NumberFormat = "$#,##0;($#,##0);-"

    '---------------------------------------------------------------- 3. Expenditure stream + 4. cumulative (helper blocks)
    rowStream = rowCat + 9
    ws.Cells(rowStream, 1).Value = "Expenditure stream by year (undiscounted)"
    ws.Cells(rowStream, 1).Font.Bold = True
    ws.Cells(rowStream + 1, 1).Value = "Year"
    rowCum = rowStream + 35
    ws.Cells(rowCum, 1).Value = "Cumulative discounted cost by year"
    ws.Cells(rowCum, 1).Font.Bold = True
    ws.Cells(rowCum + 1, 1).Value = "Year"
    For k = 0 To 30
        ws.Cells(rowStream + 2 + k, 1).Formula = "='General Information'!$D$25+" & k
        ws.Cells(rowCum + 2 + k, 1).Formula = "='General Information'!$D$25+" & k
    Next k
    i = 0
    For Each curAltObj In MainDoc.AlternativeWsCollection
        i = i + 1
        Set wsAlt = Worksheets(curAltObj.NameWsAlt)
        altName = "'" & curAltObj.NameWsAlt & "'!"
        rowInit = wsAlt.Columns("B").Find(What:="Initial Construction", LookIn:=xlValues, LookAt:=xlWhole).Row
        rowNPW = wsAlt.Columns("B").Find(What:="Net Present Worth", LookIn:=xlValues, LookAt:=xlWhole).Row
        rowFirstAct = rowInit + 1: rowLastAct = rowNPW - 1
        rngC = altName & "$C$" & rowFirstAct & ":$C$" & rowLastAct
        rngD = altName & "$D$" & rowFirstAct & ":$D$" & rowLastAct
        rngE = altName & "$E$" & rowFirstAct & ":$E$" & rowLastAct
        ws.Cells(rowStream + 1, 1 + i).Formula = "=$B$" & (ROW_FIRST + i - 1)
        ws.Cells(rowCum + 1, 1 + i).Formula = "=$B$" & (ROW_FIRST + i - 1)
        For k = 0 To 30
            q = "IF(" & k & ">'General Information'!$D$33,0,"
            ws.Cells(rowStream + 2 + k, 1 + i).Formula = "=" & q & IIf(k = 0, altName & "$D$" & rowInit & "+", "") & _
                "SUMIF(" & rngC & "," & k & "," & rngD & "))"
            ws.Cells(rowCum + 2 + k, 1 + i).Formula = "=" & IIf(k = 0, "", ws.Cells(rowCum + 1 + k, 1 + i).Address(False, False) & "+") & _
                q & IIf(k = 0, altName & "$E$" & rowInit & "+", "") & "SUMIF(" & rngC & "," & k & "," & rngE & "))"
        Next k
    Next curAltObj
    ws.Range(ws.Cells(rowStream + 2, 2), ws.Cells(rowStream + 32, 1 + numAlts)).NumberFormat = "$#,##0;($#,##0);-"
    ws.Range(ws.Cells(rowCum + 2, 2), ws.Cells(rowCum + 32, 1 + numAlts)).NumberFormat = "$#,##0;($#,##0);-"

    '---------------------------------------------------------------- 5. Discount-rate sensitivity (helper block)
    rowSens = rowCum + 35
    ws.Cells(rowSens, 1).Value = "Net present worth vs. discount rate"
    ws.Cells(rowSens, 1).Font.Bold = True
    ws.Cells(rowSens + 1, 1).Value = "Rate (%)"
    Dim nRates As Long: nRates = CLng((RATE_MAX - RATE_MIN) / RATE_STEP)
    For k = 0 To nRates
        ws.Cells(rowSens + 2 + k, 1).Value = RATE_MIN + k * RATE_STEP
    Next k
    i = 0
    For Each curAltObj In MainDoc.AlternativeWsCollection
        i = i + 1
        Set wsAlt = Worksheets(curAltObj.NameWsAlt)
        altName = "'" & curAltObj.NameWsAlt & "'!"
        rowInit = wsAlt.Columns("B").Find(What:="Initial Construction", LookIn:=xlValues, LookAt:=xlWhole).Row
        rowNPW = wsAlt.Columns("B").Find(What:="Net Present Worth", LookIn:=xlValues, LookAt:=xlWhole).Row
        rowFirstAct = rowInit + 1: rowLastAct = rowNPW - 1
        rngC = altName & "$C$" & rowFirstAct & ":$C$" & rowLastAct
        rngD = altName & "$D$" & rowFirstAct & ":$D$" & rowLastAct
        ws.Cells(rowSens + 1, 1 + i).Formula = "=$B$" & (ROW_FIRST + i - 1)
        For k = 0 To nRates
            ws.Cells(rowSens + 2 + k, 1 + i).Formula = "=" & altName & "$D$" & rowInit & "+SUMPRODUCT((" & rngD & ")*(" & rngC & _
                "<='General Information'!$D$33)/(1+$A" & (rowSens + 2 + k) & "/100)^(" & rngC & "))"
        Next k
    Next curAltObj
    ws.Range(ws.Cells(rowSens + 2, 2), ws.Cells(rowSens + 2 + nRates, 1 + numAlts)).NumberFormat = "$#,##0;($#,##0);-"
    'Row that flags a change of winner between the analysis rate and 7% (FAA AIP handbook rate)
    ws.Cells(rowSens + 3 + nRates, 1).Value = "Winner at analysis rate"
    ws.Cells(rowSens + 3 + nRates, 2).Formula = "=INDEX($B$" & ROW_FIRST & ":$B$" & lastSummaryRow & ",MATCH(MIN($I$" & ROW_FIRST & ":$I$" & lastSummaryRow & "),$I$" & ROW_FIRST & ":$I$" & lastSummaryRow & ",0))"
    ws.Cells(rowSens + 4 + nRates, 1).Value = "Winner at 7% (FAA AIP)"
    ws.Cells(rowSens + 4 + nRates, 2).Formula = "=INDEX($B$" & (rowSens + 1) & ":" & ws.Cells(rowSens + 1, 1 + numAlts).Address(False, False) & _
        ",MATCH(MIN($B$" & (rowSens + 12) & ":" & ws.Cells(rowSens + 12, 1 + numAlts).Address(False, False) & "),$B$" & (rowSens + 12) & ":" & ws.Cells(rowSens + 12, 1 + numAlts).Address(False, False) & ",0))"
    ws.Cells(rowSens + 5 + nRates, 1).Value = "Same winner?"
    ws.Cells(rowSens + 5 + nRates, 2).Formula = "=IF(B" & (rowSens + 3 + nRates) & "=B" & (rowSens + 4 + nRates) & ",""Yes"",""No - result is sensitive to the discount rate"")"

    ws.Columns("A:L").AutoFit
    ws.Columns("L").ColumnWidth = 40

    '---------------------------------------------------------------- Charts
    Dim ch As Chart, s As Series
    Dim chartLeft As Double: chartLeft = ws.Columns("N").Left
    Dim chartW As Double: chartW = 480
    Dim chartH As Double: chartH = 260

    'Chart A: present worth by category (stacked column, one column per alternative)
    Set ch = ws.Shapes.AddChart2(297, xlColumnStacked, chartLeft, ws.Rows(ROW_HDR).Top, chartW, chartH).Chart
    ch.SetSourceData Source:=ws.Range(ws.Cells(rowCat + 1, 1), ws.Cells(rowCat + 6, 1 + numAlts)), PlotBy:=xlRows
    ch.HasTitle = True: ch.ChartTitle.Text = "Present worth by category"
    ch.Axes(xlValue).TickLabels.NumberFormat = "$#,##0,,""M"""
    ch.ChartGroups(1).GapWidth = 60
    ch.HasLegend = True: ch.Legend.Position = xlLegendPositionBottom
    Dim catShade As Variant: catShade = Array(RGB(66, 66, 66), RGB(140, 140, 140), RGB(100, 100, 100), RGB(184, 182, 174), RGB(220, 220, 220))
    For k = 1 To ch.SeriesCollection.Count
        ch.SeriesCollection(k).Format.Fill.ForeColor.RGB = catShade(k - 1)
    Next k

    'Chart B: expenditure stream by calendar year (undiscounted)
    Set ch = ws.Shapes.AddChart2(201, xlColumnClustered, chartLeft, ws.Rows(ROW_HDR).Top + chartH + 12, chartW, chartH).Chart
    ch.SetSourceData Source:=ws.Range(ws.Cells(rowStream + 1, 1), ws.Cells(rowStream + 32, 1 + numAlts)), PlotBy:=xlColumns
    ch.HasTitle = True: ch.ChartTitle.Text = "Expenditure stream by year (undiscounted)"
    ch.Axes(xlValue).TickLabels.NumberFormat = "$#,##0,,""M"""
    ch.Axes(xlCategory).TickLabelSpacing = 5
    ch.ChartGroups(1).GapWidth = 40: ch.ChartGroups(1).Overlap = 0
    ch.HasLegend = True: ch.Legend.Position = xlLegendPositionBottom
    For k = 1 To ch.SeriesCollection.Count
        ch.SeriesCollection(k).Format.Fill.ForeColor.RGB = colorOf(k)
    Next k

    'Chart C: cumulative discounted cost
    Set ch = ws.Shapes.AddChart2(227, xlLine, chartLeft, ws.Rows(ROW_HDR).Top + 2 * (chartH + 12), chartW, chartH).Chart
    ch.SetSourceData Source:=ws.Range(ws.Cells(rowCum + 1, 1), ws.Cells(rowCum + 32, 1 + numAlts)), PlotBy:=xlColumns
    ch.HasTitle = True: ch.ChartTitle.Text = "Cumulative discounted cost"
    ch.Axes(xlValue).TickLabels.NumberFormat = "$#,##0,,""M"""
    ch.Axes(xlCategory).TickLabelSpacing = 5
    ch.HasLegend = True: ch.Legend.Position = xlLegendPositionBottom
    For k = 1 To ch.SeriesCollection.Count
        With ch.SeriesCollection(k)
            .Format.Line.ForeColor.RGB = colorOf(k)
            .Format.Line.Weight = 2.25
            .MarkerStyle = xlMarkerStyleNone
        End With
    Next k

    'Chart D: discount-rate sensitivity
    Set ch = ws.Shapes.AddChart2(227, xlLine, chartLeft, ws.Rows(ROW_HDR).Top + 3 * (chartH + 12), chartW, chartH).Chart
    ch.SetSourceData Source:=ws.Range(ws.Cells(rowSens + 1, 1), ws.Cells(rowSens + 2 + nRates, 1 + numAlts)), PlotBy:=xlColumns
    ch.HasTitle = True: ch.ChartTitle.Text = "Net present worth vs. discount rate (3% TDOT, 7% FAA AIP)"
    ch.Axes(xlValue).TickLabels.NumberFormat = "$#,##0,,""M"""
    ch.Axes(xlCategory).TickLabels.NumberFormat = "0.0""%"""
    ch.HasLegend = True: ch.Legend.Position = xlLegendPositionBottom
    For k = 1 To ch.SeriesCollection.Count
        With ch.SeriesCollection(k)
            .Format.Line.ForeColor.RGB = colorOf(k)
            .Format.Line.Weight = 2.25
            .MarkerStyle = xlMarkerStyleNone
        End With
    Next k

CleanExit:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Summary could not be fully built: " & Err.Description, vbExclamation, "LCCA Summary"
    End If
    Set curAltObj = Nothing
    Set wsAlt = Nothing
End Sub
