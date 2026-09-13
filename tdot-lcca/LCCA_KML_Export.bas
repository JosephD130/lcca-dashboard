Attribute VB_Name = "LCCA_KML_Export"
'==============================================================================================
' TDOT Aeronautics LCCA Framework - Google Earth export
'
' Writes a .kml file beside the workbook holding, for the project on General Information:
'   * the airport, with the whole LCCA result in its description bubble
'   * a schematic runway footprint, oriented from the runway number and sized from the area
'     entered, one per alternative, colored green for the lowest net present worth
'   * one extruded bar per alternative whose height is its net present worth
'   * every maintenance and rehabilitation event as a placemark stamped with the calendar year
'     it happens, so the Google Earth time slider walks the analysis period
'
' Run it with Alt+F8, ExportLCCAKML, or assign it to a button. Nothing here changes a cell or
' touches the internet; it only reads the workbook and writes a text file.
'
' ARA for TDOT Aeronautics, v1.2.0
'==============================================================================================
Option Explicit

Private Const DEFAULT_WIDTH_FT As Double = 100#   ' used if Summary T27 is empty
' Where the Summary keeps things. The dashboard strip sits above the results table, so these moved in
' v1.2.0; they are the same numbers build_summary.py lays the sheet out with.
Private Const ROW0 As Long = 12                   ' first alternative row of the results table
Private Const ROW1 As Long = 15                   ' last
Private Const SEC_OFF As Long = 81                ' section description row = SEC_OFF + results row
Private Const TALLEST_BAR_M As Double = 700#      ' the largest present worth stands this high
Private Const FT_PER_DEG_LAT As Double = 364000#

Public Sub ExportLCCAKML()
    Dim gi As Worksheet, sm As Worksheet
    Dim path As String, f As Integer
    Dim lat As Double, lon As Double, hasGeo As Boolean

    On Error GoTo Fail
    Set gi = ThisWorkbook.Worksheets("General Information")
    Set sm = ThisWorkbook.Worksheets("Summary")

    If Len(Trim$(CStr(gi.Range("D9").Value))) = 0 Then
        MsgBox "Pick an airport on General Information first.", vbExclamation, "LCCA export"
        Exit Sub
    End If
    If ThisWorkbook.path = "" Then
        MsgBox "Save the workbook once before exporting, so the file has somewhere to go.", vbExclamation, "LCCA export"
        Exit Sub
    End If

    ' the coordinates sit in the map-data block on the Summary (embedded reference values)
    hasGeo = IsNumeric(sm.Range("AD132").Value) And IsNumeric(sm.Range("AE132").Value)
    If hasGeo Then
        lon = CDbl(sm.Range("AD132").Value)
        lat = CDbl(sm.Range("AE132").Value)
    Else
        MsgBox "This airport has no published coordinates in the reference list, so the export has no" & vbCrLf & _
               "place to put. Pick another airport or add its coordinates to the map-data block.", vbExclamation, "LCCA export"
        Exit Sub
    End If

    path = ThisWorkbook.path & Application.PathSeparator & KmlFileName(gi)
    f = FreeFile
    Open path For Output As #f

    Print #f, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #f, "<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"">"
    Print #f, "<Document>"
    Print #f, "  <name>" & X(DocName(gi)) & "</name>"
    WriteLegend f, gi
    Print #f, "  <LookAt><longitude>" & Num(lon, 6) & "</longitude><latitude>" & Num(lat, 6) & "</latitude>" & _
              "<altitude>0</altitude><heading>" & Num(RunwayBearing(CStr(gi.Range("D22").Value)), 1) & "</heading>" & _
              "<tilt>55</tilt><range>4200</range><altitudeMode>relativeToGround</altitudeMode></LookAt>"
    WriteStyles f
    WriteAirport f, gi, sm, lat, lon
    WriteAlternatives f, gi, sm, lat, lon
    Print #f, "</Document>"
    Print #f, "</kml>"
    Close #f

    MsgBox "Written:" & vbCrLf & path & vbCrLf & vbCrLf & _
           "Open it in Google Earth. Use the time slider to step through the analysis period.", vbInformation, "LCCA export"
    Exit Sub
Fail:
    On Error Resume Next
    Close #f
    MsgBox "Export failed: " & Err.description, vbCritical, "LCCA export"
End Sub

'---------------------------------------------------------------- document pieces

' What the colors mean. Google Earth shows this when the reader clicks the document in Places.
Private Sub WriteLegend(ByVal f As Integer, gi As Worksheet)
    Dim d As String
    d = "<p>Life-cycle cost analysis, " & CStr(gi.Range("D33").Value) & " years at " & CStr(gi.Range("D34").Value) & _
        "%. Generated from the TDOT Aeronautics LCCA Framework v1.2.0.</p>"
    d = d & "<p><b>Legend</b></p><table cellpadding=" & Q & "3" & Q & " cellspacing=" & Q & "0" & Q & ">"
    d = d & LegendRow("#1f8b2e", "Lowest net present worth")
    d = d & LegendRow("#1f78d6", "Other alternatives")
    d = d & LegendRow("#c1440e", "Maintenance or rehabilitation event, shown in the year it happens")
    d = d & "</table>"
    d = d & "<p>Each alternative has a runway footprint (schematic: sized from the area entered and the width on the " & _
        "Summary, turned to the runway number) and a bar whose height is its net present worth, the tallest drawn " & _
        "700 m high. The events carry a date, so the time slider at the top of Google Earth walks the analysis period.</p>"
    Print #f, "  <description><![CDATA[" & d & "]]></description>"
End Sub

Private Function LegendRow(ByVal rgb As String, ByVal label As String) As String
    LegendRow = "<tr><td bgcolor=" & Q & rgb & Q & " width=" & Q & "18" & Q & ">&nbsp;</td><td>" & label & "</td></tr>"
End Function

Private Function Q() As String
    Q = Chr$(34)
End Function

Private Sub WriteStyles(ByVal f As Integer)
    WriteStyle f, "altLow", "ff2e8b1f", "662e8b1f"      ' green: lowest present worth
    WriteStyle f, "alt", "ffd6781f", "66d6781f"         ' blue-orange: the others
    WriteStyle f, "event", "ff0e44c1", "660e44c1"       ' orange: a maintenance or rehabilitation event
    Print #f, "  <Style id=""airport""><IconStyle><scale>1.2</scale>" & _
              "</IconStyle>" & _
              "<LabelStyle><scale>1.0</scale></LabelStyle></Style>"
End Sub

Private Sub WriteStyle(ByVal f As Integer, ByVal id As String, ByVal lineAbgr As String, ByVal polyAbgr As String)
    Print #f, "  <Style id=""" & id & """>"
    Print #f, "    <LineStyle><color>" & lineAbgr & "</color><width>2</width></LineStyle>"
    Print #f, "    <PolyStyle><color>" & polyAbgr & "</color></PolyStyle>"
    Print #f, "    <IconStyle><color>" & lineAbgr & "</color><scale>0.9</scale>" & _
              "</IconStyle>"
    Print #f, "  </Style>"
End Sub

Private Sub WriteAirport(ByVal f As Integer, gi As Worksheet, sm As Worksheet, ByVal lat As Double, ByVal lon As Double)
    Dim d As String, r As Long
    d = "<h3>" & CStr(gi.Range("D9").Value) & " (" & CStr(gi.Range("D10").Value) & ")</h3>"
    d = d & "<p>" & CStr(gi.Range("D11").Value) & ", " & CStr(gi.Range("D13").Value) & " region"
    If Len(CStr(gi.Range("D22").Value)) > 0 Then d = d & "<br/>" & CStr(gi.Range("D22").Value) & ", " & CStr(gi.Range("D23").Value)
    d = d & "<br/>Construction " & CStr(gi.Range("D25").Value) & ", " & CStr(gi.Range("D33").Value) & " years at " & _
        CStr(gi.Range("D34").Value) & "%</p>"

    d = d & "<table border=""1"" cellpadding=""4"" cellspacing=""0"">"
    d = d & "<tr><th>Alternative</th><th>Type</th><th>Initial</th><th>Net present worth</th><th>Closure days</th><th>Section</th></tr>"
    For r = ROW0 To ROW1
        If Len(CStr(sm.Cells(r, 7).Value)) > 0 Then
            d = d & "<tr><td>" & CStr(sm.Cells(r, 8).Value) & "</td><td>" & CStr(sm.Cells(r, 9).Value) & "</td>" & _
                "<td align=""right"">" & Money(sm.Cells(r, 10).Value) & "</td>" & _
                "<td align=""right"">" & Money(sm.Cells(r, 15).Value) & "</td>" & _
                "<td align=""right"">" & CStr(sm.Cells(r, 17).Value) & "</td>" & _
                "<td>" & CStr(sm.Cells(SEC_OFF + r, 16).Value) & "</td></tr>"
        End If
    Next r
    d = d & "</table>"
    d = d & "<p><b>" & CStr(sm.Range("G17").Value) & "</b><br/>" & CStr(sm.Range("G18").Value) & "</p>"

    Print #f, "  <Placemark>"
    Print #f, "    <name>" & X(CStr(gi.Range("D9").Value) & " (" & CStr(gi.Range("D10").Value) & ")") & "</name>"
    Print #f, "    <description><![CDATA[" & d & "]]></description>"
    Print #f, "    <styleUrl>#airport</styleUrl>"
    Print #f, "    <Point><coordinates>" & Num(lon, 6) & "," & Num(lat, 6) & ",0</coordinates></Point>"
    Print #f, "  </Placemark>"
End Sub

Private Sub WriteAlternatives(ByVal f As Integer, gi As Worksheet, sm As Worksheet, _
                              ByVal lat As Double, ByVal lon As Double)
    Dim r As Long, k As Long
    Dim wsName As String, altName As String, ws As Worksheet
    Dim bearing As Double, lengthFt As Double, npw As Double, best As Double
    Dim widthFt As Double, worst As Double, barScale As Double
    Dim yr0 As Long, period As Long, style As String

    bearing = RunwayBearing(CStr(gi.Range("D22").Value))
    widthFt = DEFAULT_WIDTH_FT
    If IsNumeric(sm.Range("T27").Value) Then
        If CDbl(sm.Range("T27").Value) > 0 Then widthFt = CDbl(sm.Range("T27").Value)
    End If
    lengthFt = 0
    If IsNumeric(gi.Range("D26").Value) Then
        If CDbl(gi.Range("D26").Value) > 0 Then lengthFt = CDbl(gi.Range("D26").Value) * 9# / widthFt
    End If
    yr0 = CLng(gi.Range("D25").Value)
    period = CLng(gi.Range("D33").Value)

    ' seeded on the first alternative found, not on zero: a net present worth of exactly 0 is a value,
    ' not an empty slot, and seeding on zero would hand the "lowest" style to the wrong alternative
    Dim seeded As Boolean
    seeded = False: best = 0: worst = 0
    For r = ROW0 To ROW1
        If Len(CStr(sm.Cells(r, 7).Value)) > 0 Then
            npw = CDbl(sm.Cells(r, 15).Value)
            If Not seeded Then
                best = npw: worst = npw: seeded = True
            Else
                If npw < best Then best = npw
                If npw > worst Then worst = npw
            End If
        End If
    Next r
    barScale = 0
    If worst > 0 Then barScale = TALLEST_BAR_M / (worst / 1000000#)   ' the tallest bar is always the same height

    k = 0
    For r = ROW0 To ROW1
        wsName = CStr(sm.Cells(r, 7).Value)
        If Len(wsName) > 0 Then
            altName = CStr(sm.Cells(r, 8).Value)
            npw = CDbl(sm.Cells(r, 15).Value)
            style = IIf(npw = best, "altLow", "alt")
            Print #f, "  <Folder>"
            Print #f, "    <name>" & X(altName & " - " & Money(npw) & IIf(npw = best, " (lowest)", "")) & "</name>"

            If lengthFt > 0 Then WriteRunway f, lat, lon, bearing, lengthFt, widthFt, k, style, _
                                              altName & " footprint", CStr(sm.Cells(SEC_OFF + r, 16).Value)
            WriteBar f, lat, lon, k, style, altName, npw, barScale
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(wsName)
            On Error GoTo 0
            If Not ws Is Nothing Then WriteEvents f, ws, lat, lon, bearing, lengthFt, k, yr0, period, altName
            Set ws = Nothing
            Print #f, "  </Folder>"
            k = k + 1
        End If
    Next r
End Sub

' Schematic runway footprint: a rectangle centered on the airport, turned to the runway heading.
Private Sub WriteRunway(ByVal f As Integer, ByVal lat As Double, ByVal lon As Double, ByVal bearing As Double, _
                        ByVal lengthFt As Double, ByVal widthFt As Double, ByVal k As Long, ByVal style As String, _
                        ByVal nm As String, ByVal section As String)
    Dim hx As Double, hy As Double, off As Double
    Dim c(0 To 4) As String, i As Long
    off = (k - 1.5) * widthFt * 2.2                      ' lay the alternatives side by side so all are visible
    hx = lengthFt / 2#: hy = widthFt / 2#
    c(0) = Corner(lat, lon, bearing, -hx, -hy + off)
    c(1) = Corner(lat, lon, bearing, hx, -hy + off)
    c(2) = Corner(lat, lon, bearing, hx, hy + off)
    c(3) = Corner(lat, lon, bearing, -hx, hy + off)
    c(4) = c(0)
    Print #f, "    <Placemark>"
    Print #f, "      <name>" & X(nm) & "</name>"
    Print #f, "      <description><![CDATA[" & X2(section) & "<br/>Schematic footprint: " & Format$(lengthFt, "#,##0") & _
              " x " & Format$(widthFt, "#,##0") & " ft from the area entered, turned to the runway number. Not survey geometry.]]></description>"
    Print #f, "      <styleUrl>#" & style & "</styleUrl>"
    Print #f, "      <Polygon><tessellate>1</tessellate><outerBoundaryIs><LinearRing><coordinates>"
    For i = 0 To 4
        Print #f, "        " & c(i)
    Next i
    Print #f, "      </coordinates></LinearRing></outerBoundaryIs></Polygon>"
    Print #f, "    </Placemark>"
End Sub

' One extruded bar per alternative: height proportional to net present worth.
Private Sub WriteBar(ByVal f As Integer, ByVal lat As Double, ByVal lon As Double, ByVal k As Long, _
                     ByVal style As String, ByVal altName As String, ByVal npw As Double, ByVal barScale As Double)
    Dim h As Double, s As Double, i As Long, e As Double
    Dim c(0 To 4) As String
    h = (npw / 1000000#) * barScale
    s = 260#                                             ' bar footprint, feet
    e = 1200# + k * 700#                                 ' set the bars off to one side, in a row
    c(0) = Corner3(lat, lon, 90#, e, 0#, h)
    c(1) = Corner3(lat, lon, 90#, e + s, 0#, h)
    c(2) = Corner3(lat, lon, 90#, e + s, s, h)
    c(3) = Corner3(lat, lon, 90#, e, s, h)
    c(4) = c(0)
    Print #f, "    <Placemark>"
    Print #f, "      <name>" & X(altName & ": " & Money(npw)) & "</name>"
    Print #f, "      <styleUrl>#" & style & "</styleUrl>"
    Print #f, "      <Polygon><extrude>1</extrude><altitudeMode>relativeToGround</altitudeMode>"
    Print #f, "        <outerBoundaryIs><LinearRing><coordinates>"
    For i = 0 To 4
        Print #f, "          " & c(i)
    Next i
    Print #f, "      </coordinates></LinearRing></outerBoundaryIs></Polygon>"
    Print #f, "    </Placemark>"
End Sub

' Every activity on the alternative worksheet, stamped with the calendar year it happens.
Private Sub WriteEvents(ByVal f As Integer, ws As Worksheet, ByVal lat As Double, ByVal lon As Double, _
                        ByVal bearing As Double, ByVal lengthFt As Double, ByVal k As Long, _
                        ByVal yr0 As Long, ByVal period As Long, ByVal altName As String)
    Dim r As Long, hdrRow As Long, npwRow As Long, n As Long, pos As Double
    Dim item As String, yrOff As Double, cost As Double, days As Double, daily As Double
    Dim lab As String

    hdrRow = 0: npwRow = 0
    For r = 30 To 60
        If Trim$(CStr(ws.Cells(r, 2).Value)) = "Item" Then hdrRow = r
        If Trim$(CStr(ws.Cells(r, 2).Value)) = "Net Present Worth" Then npwRow = r
    Next r
    If hdrRow = 0 Or npwRow = 0 Then Exit Sub
    daily = 0
    If IsNumeric(ws.Range("F2").Value) Then daily = CDbl(ws.Range("F2").Value)

    n = 0
    For r = hdrRow + 1 To npwRow - 1
        item = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(item) > 0 And InStr(1, item, "Indirect", vbTextCompare) = 0 And item <> "Item" Then
            If IsNumeric(ws.Cells(r, 3).Value) And IsNumeric(ws.Cells(r, 4).Value) Then
                yrOff = CDbl(ws.Cells(r, 3).Value)
                cost = CDbl(ws.Cells(r, 4).Value)
                If yrOff >= 1000 Then yrOff = 0                      ' the initial row carries the calendar year
                If cost <> 0 And yrOff <= period Then
                    days = 0
                    If daily > 0 And IsNumeric(ws.Cells(r + 1, 4).Value) Then
                        If InStr(1, CStr(ws.Cells(r + 1, 2).Value), "Indirect", vbTextCompare) > 0 Then
                            days = CDbl(ws.Cells(r + 1, 4).Value) / daily
                        End If
                    End If
                    pos = (n / 8# - 0.45) * lengthFt                  ' walk the events along the runway
                    lab = ShortAlt(altName) & " " & item & " " & CStr(yr0 + CLng(yrOff))
                    Print #f, "    <Placemark>"
                    Print #f, "      <name>" & X(lab) & "</name>"
                    Print #f, "      <description><![CDATA[<b>" & X2(altName) & "</b><br/>" & X2(item) & _
                              "<br/>Year " & CStr(yr0 + CLng(yrOff)) & " (" & CStr(CLng(yrOff)) & " years after construction)" & _
                              "<br/>Cost in the year: " & Money(cost) & _
                              IIf(days > 0, "<br/>Runway closed about " & Format$(days, "0") & " days", "") & "]]></description>"
                    Print #f, "      <TimeSpan><begin>" & CStr(yr0 + CLng(yrOff)) & "-01-01</begin><end>" & _
                              CStr(yr0 + CLng(yrOff)) & "-12-31</end></TimeSpan>"
                    Print #f, "      <styleUrl>#event</styleUrl>"
                    Print #f, "      <Point><coordinates>" & Corner3(lat, lon, bearing, pos, (k - 1.5) * 220#, 0#) & "</coordinates></Point>"
                    Print #f, "    </Placemark>"
                    n = n + 1
                End If
            End If
        End If
    Next r
End Sub

'---------------------------------------------------------------- helpers

' Offset in feet along (a) the runway direction and (b) across it, returned as "lon,lat,0".
Private Function Corner(ByVal lat As Double, ByVal lon As Double, ByVal bearing As Double, _
                        ByVal alongFt As Double, ByVal acrossFt As Double) As String
    Corner = Corner3(lat, lon, bearing, alongFt, acrossFt, 0#)
End Function

Private Function Corner3(ByVal lat As Double, ByVal lon As Double, ByVal bearing As Double, _
                         ByVal alongFt As Double, ByVal acrossFt As Double, ByVal altM As Double) As String
    Dim rad As Double, dN As Double, dE As Double, la As Double, lo As Double
    rad = bearing * 3.14159265358979 / 180#
    dN = alongFt * Cos(rad) - acrossFt * Sin(rad)
    dE = alongFt * Sin(rad) + acrossFt * Cos(rad)
    la = lat + dN / FT_PER_DEG_LAT
    lo = lon + dE / (FT_PER_DEG_LAT * Cos(lat * 3.14159265358979 / 180#))
    Corner3 = Num(lo, 6) & "," & Num(la, 6) & "," & Num(altM, 1)
End Function

' "Runway 2-20" -> 20 degrees. Anything unparseable points the footprint north-south.
Private Function RunwayBearing(ByVal branch As String) As Double
    Dim i As Long, ch As String, digits As String
    For i = 1 To Len(branch)
        ch = Mid$(branch, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i
    If Len(digits) = 0 Then
        RunwayBearing = 0
    Else
        RunwayBearing = (CDbl(digits) Mod 36) * 10#
    End If
End Function

' "Alternative 2" -> "Alt 2", so the labels on the globe stay short.
Private Function ShortAlt(ByVal s As String) As String
    If InStr(1, s, "Alternative ", vbTextCompare) = 1 Then
        ShortAlt = "Alt " & Mid$(s, 13)
    Else
        ShortAlt = s
    End If
End Function

Private Function KmlFileName(gi As Worksheet) As String
    Dim s As String
    s = CStr(gi.Range("D10").Value) & "_" & CStr(gi.Range("D22").Value) & "_LCCA"
    KmlFileName = Clean(s) & ".kml"
End Function

Private Function DocName(gi As Worksheet) As String
    DocName = CStr(gi.Range("D9").Value) & " (" & CStr(gi.Range("D10").Value) & ") - " & _
              CStr(gi.Range("D16").Value) & IIf(Len(CStr(gi.Range("D16").Value)) = 0, CStr(gi.Range("D22").Value), "")
End Function

Private Function Clean(ByVal s As String) As String
    Dim i As Long, ch As String, o As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (UCase$(ch) >= "A" And UCase$(ch) <= "Z") Then
            o = o & ch
        ElseIf ch = " " Or ch = "-" Or ch = "_" Then
            o = o & "_"
        End If
    Next i
    Do While InStr(o, "__") > 0
        o = Replace(o, "__", "_")
    Loop
    Clean = o
End Function

' Decimal point, whatever the regional settings say.
Private Function Num(ByVal v As Double, ByVal places As Integer) As String
    Num = Replace(Format$(v, "0." & String$(places, "0")), ",", ".")
End Function

Private Function Money(ByVal v As Variant) As String
    If Not IsNumeric(v) Then
        Money = ""
    Else
        Money = "$" & Format$(CDbl(v), "#,##0")
    End If
End Function

' XML escape for element text.
Private Function X(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    X = s
End Function

' Escape for text that sits inside a CDATA description (only the CDATA terminator matters).
Private Function X2(ByVal s As String) As String
    X2 = Replace(s, "]]>", "]] >")
End Function
