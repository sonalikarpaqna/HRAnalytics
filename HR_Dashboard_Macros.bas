Attribute VB_Name = "HR_Dashboard_Macros"
'==========================================================================
' HR Analytics Dashboard - VBA macros
' Author : Sonali K
' Import : Alt+F11 > File > Import File > choose this .bas file,
'          then save the workbook as "Excel Macro-Enabled Workbook (.xlsm)".
' First run: BuildDashboardButtons (Alt+F8) to place buttons on the Dashboard.
'==========================================================================
Option Explicit

Private Const DASH As String = "Dashboard"
Private Const DEPT_CELL As String = "D4"
Private Const EMP_CELL As String = "J4"
Private Const ALL_DEPTS As String = "All Departments"
Private Const ENGINE As String = "Calc_Engine"

'---------- helpers ----------
Private Function SelectedDept() As String
    SelectedDept = CStr(Worksheets(DASH).Range(DEPT_CELL).Value)
End Function

Private Function DeptMatches(ByVal deptName As String) As Boolean
    DeptMatches = (SelectedDept() = ALL_DEPTS) Or (deptName = SelectedDept())
End Function

Private Function LastRow(ByVal ws As Worksheet) As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

Private Function FreshSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Set FreshSheet = ws
End Function

Private Sub StyleHeader(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(27, 42, 65)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    rng.EntireRow.RowHeight = 30
End Sub

Private Sub StartFast()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub EndFast()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'---------- 1. refresh / reset ----------
Public Sub RefreshDashboard()
    Application.CalculateFull
    Application.StatusBar = "HR dashboard refreshed at " & Format(Now, "hh:mm:ss")
    Application.OnTime Now + TimeValue("00:00:04"), "ClearStatus"
End Sub

Public Sub ClearStatus()
    Application.StatusBar = False
End Sub

Public Sub ResetDashboard()
    With Worksheets(DASH)
        .Range(DEPT_CELL).Value = ALL_DEPTS
        Application.Calculate
        .Range(EMP_CELL).Value = Worksheets(ENGINE).Range("J2").Value
        .Activate
        .Range("A1").Select
    End With
    ActiveWindow.ScrollRow = 1
End Sub

'---------- 2. browse employees within the selected department ----------
Public Sub NextEmployee()
    MoveEmployee 1
End Sub

Public Sub PreviousEmployee()
    MoveEmployee -1
End Sub

Private Sub MoveEmployee(ByVal stepBy As Long)
    Dim eng As Worksheet, n As Long, cur As String, pos As Variant, i As Long
    Set eng = Worksheets(ENGINE)
    n = Application.WorksheetFunction.Count(eng.Range("H2:H501"))
    If n = 0 Then Exit Sub
    cur = CStr(Worksheets(DASH).Range(EMP_CELL).Value)
    pos = Application.Match(cur, eng.Range("J2:J" & (n + 1)), 0)
    If IsError(pos) Then
        i = 1
    Else
        i = CLng(pos) + stepBy
    End If
    If i > n Then i = 1
    If i < 1 Then i = n
    Worksheets(DASH).Range(EMP_CELL).Value = eng.Range("J" & (i + 1)).Value
End Sub

'---------- 3. search an employee by name or ID ----------
Public Sub FindEmployee()
    Dim q As String, ws As Worksheet, r As Long, lr As Long
    Dim hits As Collection, msg As String, pick As String, idx As Long
    q = Trim(InputBox("Type part of a name or an Emp ID (for example: Priya or EMP1045)", "Find employee"))
    If q = "" Then Exit Sub
    Set ws = Worksheets("Employee_Master")
    lr = LastRow(ws)
    Set hits = New Collection
    For r = 2 To lr
        If InStr(1, ws.Cells(r, 2).Value, q, vbTextCompare) > 0 Or StrComp(ws.Cells(r, 1).Value, q, vbTextCompare) = 0 Then
            hits.Add r
            If hits.Count >= 15 Then Exit For
        End If
    Next r
    If hits.Count = 0 Then
        MsgBox "No employee found for """ & q & """.", vbInformation, "Find employee"
        Exit Sub
    End If
    idx = 1
    If hits.Count > 1 Then
        For r = 1 To hits.Count
            msg = msg & r & ". " & ws.Cells(hits(r), 1).Value & "  " & ws.Cells(hits(r), 2).Value & "  (" & ws.Cells(hits(r), 9).Value & ")" & vbCrLf
        Next r
        pick = InputBox(msg & vbCrLf & "Enter the number of the employee to show:", "Several matches", "1")
        If pick = "" Or Not IsNumeric(pick) Then Exit Sub
        idx = CLng(pick)
        If idx < 1 Or idx > hits.Count Then Exit Sub
    End If
    r = hits(idx)
    With Worksheets(DASH)
        .Range(DEPT_CELL).Value = ws.Cells(r, 9).Value
        Application.Calculate
        .Range(EMP_CELL).Value = ws.Cells(r, 1).Value & " | " & ws.Cells(r, 2).Value
        .Activate
        .Range("B19").Select
    End With
End Sub

'---------- 4. promotion list for the selected department ----------
Public Sub GeneratePromotionList()
    Dim src As Worksheet, comp As Worksheet, out As Worksheet, r As Long, o As Long, lr As Long
    StartFast
    Set src = Worksheets("Performance")
    Set comp = Worksheets("Compensation")
    Set out = FreshSheet("Promotion_List")
    out.Range("A1").Value = "Promotion-eligible employees - " & SelectedDept() & " - generated " & Format(Now, "dd-mmm-yyyy hh:mm")
    out.Range("A1").Font.Bold = True
    out.Range("A1").Font.Size = 13
    out.Range("A3:L3").Value = Array("Emp ID", "Employee Name", "Manager", "Department", "Level", "Rating", "Rating Label", _
                                     "Yrs in Level", "Potential", "9-Box", "Annual CTC (Rs)", "Compa-Ratio")
    StyleHeader out.Range("A3:L3")
    o = 4
    lr = LastRow(src)
    For r = 2 To lr
        If src.Cells(r, 18).Value = "Eligible" And DeptMatches(CStr(src.Cells(r, 4).Value)) Then
            out.Cells(o, 1).Value = src.Cells(r, 1).Value
            out.Cells(o, 2).Value = src.Cells(r, 2).Value
            out.Cells(o, 3).Value = src.Cells(r, 3).Value
            out.Cells(o, 4).Value = src.Cells(r, 4).Value
            out.Cells(o, 5).Value = src.Cells(r, 5).Value
            out.Cells(o, 6).Value = src.Cells(r, 7).Value
            out.Cells(o, 7).Value = src.Cells(r, 8).Value
            out.Cells(o, 8).Value = Round(src.Cells(r, 17).Value, 1)
            out.Cells(o, 9).Value = src.Cells(r, 13).Value
            out.Cells(o, 10).Value = src.Cells(r, 15).Value
            out.Cells(o, 11).Value = comp.Cells(r, 6).Value
            out.Cells(o, 12).Value = comp.Cells(r, 10).Value
            o = o + 1
        End If
    Next r
    out.Range("K4:K" & o).NumberFormat = "#,##0"
    out.Range("L4:L" & o).NumberFormat = "0.00"
    If o > 4 Then
        out.Range("A3:L" & (o - 1)).Sort Key1:=out.Range("F4"), Order1:=xlDescending, Key2:=out.Range("H4"), Order2:=xlDescending, Header:=xlYes
        out.Range("A3:L" & (o - 1)).AutoFilter
    End If
    out.Cells(o + 1, 1).Value = "Total eligible: " & (o - 4)
    out.Cells(o + 1, 1).Font.Bold = True
    out.Columns("A:L").AutoFit
    out.Activate
    EndFast
    MsgBox (o - 4) & " promotion-eligible employees listed for " & SelectedDept() & ".", vbInformation, "Promotion list"
End Sub

'---------- 5. flight-risk watch list ----------
Public Sub GenerateFlightRiskList()
    Dim src As Worksheet, out As Worksheet, r As Long, o As Long, lr As Long
    StartFast
    Set src = Worksheets("Engagement")
    Set out = FreshSheet("Flight_Risk_List")
    out.Range("A1").Value = "High flight-risk employees - " & SelectedDept() & " - generated " & Format(Now, "dd-mmm-yyyy hh:mm")
    out.Range("A1").Font.Bold = True
    out.Range("A1").Font.Size = 13
    out.Range("A3:I3").Value = Array("Emp ID", "Employee Name", "Department", "Manager", "Engagement (1-5)", "eNPS", "Rating", "Compa-Ratio", "Risk Points")
    StyleHeader out.Range("A3:I3")
    o = 4
    lr = LastRow(src)
    For r = 2 To lr
        If src.Cells(r, 19).Value = "High" And DeptMatches(CStr(src.Cells(r, 3).Value)) Then
            out.Cells(o, 1).Value = src.Cells(r, 1).Value
            out.Cells(o, 2).Value = src.Cells(r, 2).Value
            out.Cells(o, 3).Value = src.Cells(r, 3).Value
            out.Cells(o, 4).Value = src.Cells(r, 4).Value
            out.Cells(o, 5).Value = src.Cells(r, 13).Value
            out.Cells(o, 6).Value = src.Cells(r, 14).Value
            out.Cells(o, 7).Value = src.Cells(r, 16).Value
            out.Cells(o, 8).Value = src.Cells(r, 17).Value
            out.Cells(o, 9).Value = src.Cells(r, 18).Value
            If src.Cells(r, 16).Value >= 4 Then out.Range(out.Cells(o, 1), out.Cells(o, 9)).Interior.Color = RGB(248, 201, 207)
            o = o + 1
        End If
    Next r
    If o > 4 Then
        out.Range("A3:I" & (o - 1)).Sort Key1:=out.Range("G4"), Order1:=xlDescending, Key2:=out.Range("E4"), Order2:=xlAscending, Header:=xlYes
        out.Range("A3:I" & (o - 1)).AutoFilter
    End If
    out.Cells(o + 1, 1).Value = "Rows shaded pink are high performers (rating 4-5) at risk - priority for stay interviews."
    out.Cells(o + 1, 1).Font.Italic = True
    out.Columns("A:I").AutoFit
    out.Activate
    EndFast
End Sub

'---------- 6. export a department report to a new workbook ----------
Public Sub ExportDepartmentReport()
    Dim nb As Workbook, sSum As Worksheet, sEmp As Worksheet, dash As Worksheet, wsExp As Worksheet
    Dim r As Long, c As Long, o As Long, lastExp As Long, fPath As String, fName As String
    Set dash = Worksheets(DASH)
    Set wsExp = Worksheets("Employee_Explorer")
    Application.Calculate
    StartFast
    ThisWorkbook.Activate
    Set nb = Workbooks.Add(xlWBATWorksheet)
    Set sSum = nb.Worksheets(1)
    sSum.Name = "Summary"
    sSum.Range("A1").Value = "HR Report - " & SelectedDept() & " - " & Format(Date, "dd-mmm-yyyy")
    sSum.Range("A1").Font.Bold = True
    sSum.Range("A1").Font.Size = 14
    sSum.Range("A3:C3").Value = Array("Metric", "Value", "Detail")
    StyleHeader sSum.Range("A3:C3")
    o = 4
    For r = 6 To 14 Step 4            ' three rows of KPI tiles
        For c = 2 To 20 Step 3        ' seven tiles per row
            sSum.Cells(o, 1).Value = dash.Cells(r, c).Value
            sSum.Cells(o, 2).Value = dash.Cells(r + 1, c).Text
            sSum.Cells(o, 3).Value = dash.Cells(r + 2, c).Value
            o = o + 1
        Next c
    Next r
    sSum.Columns("A:C").AutoFit
    Set sEmp = nb.Worksheets.Add(After:=sSum)
    sEmp.Name = "Employees"
    lastExp = wsExp.Cells(wsExp.Rows.Count, 2).End(xlUp).Row
    Do While lastExp > 4 And wsExp.Cells(lastExp, 2).Value = ""
        lastExp = lastExp - 1
    Loop
    wsExp.Range("A4:Q" & lastExp).Copy
    sEmp.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    sEmp.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    sEmp.Columns("A:Q").AutoFit
    fPath = ThisWorkbook.Path
    If fPath = "" Then fPath = Application.DefaultFilePath
    fName = fPath & Application.PathSeparator & "HR_Report_" & Replace(Replace(SelectedDept(), " ", "_"), "&", "and") & "_" & Format(Date, "yyyymmdd") & ".xlsx"
    Application.DisplayAlerts = False
    nb.SaveAs Filename:=fName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    EndFast
    MsgBox "Report saved:" & vbCrLf & fName, vbInformation, "Export complete"
End Sub

'---------- 7. dashboard to PDF ----------
Public Sub ExportDashboardPDF()
    Dim fPath As String, fName As String
    fPath = ThisWorkbook.Path
    If fPath = "" Then fPath = Application.DefaultFilePath
    fName = fPath & Application.PathSeparator & "HR_Dashboard_" & Replace(Replace(SelectedDept(), " ", "_"), "&", "and") & "_" & Format(Date, "yyyymmdd") & ".pdf"
    Worksheets(DASH).ExportAsFixedFormat Type:=xlTypePDF, Filename:=fName, Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub

'---------- 8. one-time setup: buttons on the Dashboard ----------
Public Sub BuildDashboardButtons()
    Dim ws As Worksheet, b As Object, i As Long, x As Double, w As Double, topY As Double, h As Double
    Dim caps As Variant, macs As Variant
    Set ws = Worksheets(DASH)
    For i = ws.Buttons.Count To 1 Step -1
        If Left(ws.Buttons(i).Name, 4) = "btn_" Then ws.Buttons(i).Delete
    Next i
    caps = Array("Reset", "< Previous", "Next >", "Find Employee", "Promotion List", "Flight-Risk List", "Export Dept Report", "Save as PDF", "Refresh")
    macs = Array("ResetDashboard", "PreviousEmployee", "NextEmployee", "FindEmployee", "GeneratePromotionList", "GenerateFlightRiskList", "ExportDepartmentReport", "ExportDashboardPDF", "RefreshDashboard")
    x = ws.Range("B5").Left
    w = (ws.Range("W5").Left - x) / (UBound(caps) + 1)
    topY = ws.Range("B5").Top + 3
    h = ws.Range("B5").Height - 6
    ws.Range("B5").MergeArea.ClearContents
    For i = 0 To UBound(caps)
        Set b = ws.Buttons.Add(x + i * w + 2, topY, w - 4, h)
        b.Name = "btn_" & macs(i)
        b.Caption = caps(i)
        b.OnAction = macs(i)
        b.Font.Bold = True
        b.Font.Size = 9
    Next i
    MsgBox "Buttons added. Save the workbook as .xlsm to keep them.", vbInformation, "Dashboard buttons"
End Sub
