Attribute VB_Name = "LCCA_NewStudy"
' -----------------------------------------------------------------------------
'  TDOT Aeronautics LCCA Framework v1.2.0
'  NewStudy: save a clean copy of this workbook so the next project starts from
'  a blank form, without touching the study that is open.
'
'  Run it from the "+ New Study" button on General Information, or Alt+F8 >
'  NewStudy. This module ships inside the workbook; no import step is needed.
' -----------------------------------------------------------------------------
Option Explicit

Public Sub NewStudy()
    Dim src As Workbook, copyWb As Workbook
    Dim target As String, baseName As String, folder As String, ext As String
    Dim n As Long, i As Long, answer As VbMsgBoxResult
    Dim ws As Worksheet, keep As Object

    Set src = ThisWorkbook

    If src.Path = "" Then
        MsgBox "Save this workbook somewhere first, then click New Study again." & vbCrLf & vbCrLf & _
               "New Study writes the fresh copy next to the file it starts from, so the file needs a home.", _
               vbExclamation, "New Study"
        Exit Sub
    End If

    ' --- next free name beside this one: <name>_1.xlsm, _2, _3 ...
    folder = src.Path
    i = InStrRev(src.Name, ".")
    baseName = Left$(src.Name, i - 1)
    ext = Mid$(src.Name, i)
    n = 1
    Do While Len(Dir$(folder & Application.PathSeparator & baseName & "_" & n & ext)) > 0
        n = n + 1
    Loop
    target = folder & Application.PathSeparator & baseName & "_" & n & ext

    answer = MsgBox("Start a new study?" & vbCrLf & vbCrLf & _
                    "This saves a clean copy with every input cleared and every alternative removed:" & vbCrLf & vbCrLf & _
                    baseName & "_" & n & ext & vbCrLf & vbCrLf & _
                    "The study you have open is not changed.", _
                    vbQuestion + vbOKCancel + vbDefaultButton1, "New Study")
    If answer <> vbOK Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Failed

    src.SaveCopyAs target
    Set copyWb = Workbooks.Open(target)

    ' --- the sheets that ship with the framework; anything else is an alternative worksheet
    Set keep = CreateObject("Scripting.Dictionary")
    keep.CompareMode = 1
    For Each ws In src.Worksheets
        If Not IsAlternativeSheet(src, ws.Name) Then keep(ws.Name) = True
    Next ws

    For i = copyWb.Worksheets.Count To 1 Step -1
        Set ws = copyWb.Worksheets(i)
        If Not keep.Exists(ws.Name) Then
            ws.Visible = xlSheetVisible
            ws.Delete
        End If
    Next i

    ClearStudy copyWb

    copyWb.Worksheets("General Information").Activate
    copyWb.Worksheets("General Information").Range("D9").Select
    copyWb.Save
    copyWb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "New study saved as:" & vbCrLf & vbCrLf & target & vbCrLf & vbCrLf & _
           "Open it and start at General Information. This file is unchanged.", _
           vbInformation, "New Study"
    Exit Sub

Failed:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not copyWb Is Nothing Then
        On Error Resume Next
        copyWb.Close SaveChanges:=False
    End If
    MsgBox "New Study could not finish:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "New Study"
End Sub

Private Function IsAlternativeSheet(wb As Workbook, sheetName As String) As Boolean
' An alternative worksheet is one the Alternative Setup form registered on Database column D.
    Dim db As Worksheet, r As Long, last As Long
    On Error Resume Next
    Set db = wb.Worksheets("Database")
    On Error GoTo 0
    If db Is Nothing Then Exit Function
    last = db.Cells(db.Rows.Count, 4).End(xlUp).Row
    For r = 4 To last
        If StrComp(Trim$(CStr(db.Cells(r, 4).Value)), sheetName, vbTextCompare) = 0 Then
            IsAlternativeSheet = True
            Exit Function
        End If
    Next r
End Function

Private Sub ClearStudy(wb As Workbook)
' Empty every cell a user fills in, and put the workbook defaults back.
    Dim gi As Worksheet, db As Worksheet, sm As Worksheet

    Set gi = wb.Worksheets("General Information")
    gi.Range("D9").ClearContents                     ' airport
    gi.Range("D14:D17").ClearContents                ' consultant, project number, name, date
    gi.Range("D21:D24").ClearContents                ' branch and project description
    gi.Range("D25:D29").ClearContents                ' year, areas, markings
    gi.Range("D33").Value = 30                       ' analysis period, TDOT policy
    gi.Range("D34").Value = 3                        ' discount rate, TDOT policy
    gi.Range("D36").Value = 10                       ' mobilization percent
    gi.Range("D37").Value = 5                        ' engineering percent
    gi.Range("D38").Value = "No"                     ' account for indirect cost

    On Error Resume Next
    Set db = wb.Worksheets("Database")
    If Not db Is Nothing Then db.Range("A4:D1000").ClearContents
    Set sm = wb.Worksheets("Summary")
    If Not sm Is Nothing Then
        sm.Range("A4:E1000").ClearContents           ' the small table the setup form maintains
        sm.Range("T27").ClearContents                ' runway width, for the Google Earth footprint
    End If
    On Error GoTo 0
End Sub
