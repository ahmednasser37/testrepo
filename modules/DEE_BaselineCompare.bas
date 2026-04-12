Attribute VB_Name = "DEE_BaselineCompare"
Option Explicit

' =============================================================================
' DEE_BaselineCompare.bas -- Baseline Comparison (Section 15.14)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Loads a second XER file as a baseline and compares it against the current
' Schedule sheet. Outputs a side-by-side diff sheet with slip indicators.
'
' STORAGE MODEL: Baseline stored as hidden sheet "_DEE_Baseline" with same
' A-F layout as Schedule. Survives workbook save/reopen without re-import.
'
' COLOR CODING:
'   Green  = on track (0 day slip)
'   Yellow = minor slip (1-5 days)
'   Red    = major slip (> 5 days)
'   Blue   = new activity (not in baseline)
'   Gray   = removed activity (in baseline, not in current)
' =============================================================================

Private Const BASELINE_SHEET As String = "_DEE_Baseline"

Private Const COLOR_ON_TRACK  As Long = RGB(0, 176, 80)
Private Const COLOR_MINOR     As Long = RGB(255, 192, 0)
Private Const COLOR_MAJOR     As Long = RGB(255, 0, 0)
Private Const COLOR_NEW       As Long = RGB(0, 112, 192)
Private Const COLOR_REMOVED   As Long = RGB(191, 191, 191)

' ---------------------------------------------------------------------------
' LoadBaseline -- Import a XER file and store as hidden baseline sheet
' ---------------------------------------------------------------------------
Public Sub LoadBaseline()
    ' File picker
    Dim filePath As String
    filePath = DEE_XERParser.GetXERFilePath()
    If filePath = "" Then Exit Sub

    DEE_Utils.StartProgress "Loading Baseline XER"

    ' Parse XER
    Dim success As Boolean
    success = DEE_XERParser.ParseXERFile(filePath)
    If Not success Then DEE_Utils.EndProgress: Exit Sub

    ' Build baseline schedule in memory using ScheduleBuilder logic
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Remove existing baseline sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(BASELINE_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create hidden baseline sheet
    Dim wsBL As Worksheet
    Set wsBL = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsBL.Name = BASELINE_SHEET
    wsBL.Visible = xlSheetVeryHidden  ' Hidden from user -- not even in sheet tab list

    ' Write header
    wsBL.Cells(1, 1).Value = "Activity ID"
    wsBL.Cells(1, 2).Value = "Activity Name"
    wsBL.Cells(1, 3).Value = "BL Start"
    wsBL.Cells(1, 4).Value = "BL Finish"
    wsBL.Cells(1, 5).Value = "BL Budget"
    wsBL.Cells(1, 6).Value = "BL Budget"

    ' Use DEE_ScheduleBuilder to populate baseline sheet
    ' We temporarily rename it to "Schedule" to reuse the builder, then rename back
    ' Safer: directly call the build logic into wsBL
    BuildBaselineSheet wsBL

    DEE_Utils.EndProgress
    MsgBox "Baseline loaded: " & wb.Worksheets(BASELINE_SHEET).UsedRange.Rows.Count - 1 & _
           " rows stored." & vbCrLf & "Run 'Compare Baseline' to see the diff.", _
           vbInformation, "Protocol DEE -- Baseline Loaded"
End Sub

' ---------------------------------------------------------------------------
' BuildBaselineSheet -- Populate baseline sheet from parsed XER (reuses parser globals)
' ---------------------------------------------------------------------------
Private Sub BuildBaselineSheet(ws As Worksheet)
    If Not DEE_XERParser.TableExists("TASK") Then Exit Sub

    Dim taskFields As Variant
    taskFields = DEE_XERParser.GetFieldNames("TASK")
    Dim taskRows As Collection
    Set taskRows = DEE_XERParser.GetTable("TASK")

    Dim taskCodeIdx As Integer
    Dim taskNameIdx As Integer
    Dim taskStartIdx As Integer
    Dim taskEndIdx As Integer
    Dim taskBudgetIdx As Integer
    Dim taskTypeIdx As Integer

    taskCodeIdx  = DEE_XERParser.FindFieldIndex(taskFields, "task_code")
    taskNameIdx  = DEE_XERParser.FindFieldIndex(taskFields, "task_name")
    taskStartIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_start_date")
    taskEndIdx   = DEE_XERParser.FindFieldIndex(taskFields, "target_end_date")
    taskBudgetIdx = DEE_XERParser.FindFieldIndex(taskFields, "budget_qty")
    taskTypeIdx  = DEE_XERParser.FindFieldIndex(taskFields, "task_type")

    Dim dataArr() As Variant
    ReDim dataArr(0 To taskRows.Count, 0 To 5)
    Dim rowIdx As Long
    rowIdx = 0

    Dim taskRow As Variant
    For Each taskRow In taskRows
        Dim tr As Variant
        tr = taskRow

        Dim tType As String
        If taskTypeIdx >= 0 And taskTypeIdx <= UBound(tr) Then tType = Trim(tr(taskTypeIdx))
        If tType = "TT_WBS" Then GoTo NextBLTask

        Dim tCode As String
        Dim tName As String
        If taskCodeIdx >= 0 And taskCodeIdx <= UBound(tr) Then tCode = Trim(tr(taskCodeIdx))
        If taskNameIdx >= 0 And taskNameIdx <= UBound(tr) Then tName = Trim(tr(taskNameIdx))

        Dim tStart As Variant
        Dim tEnd As Variant
        If taskStartIdx >= 0 And taskStartIdx <= UBound(tr) Then
            tStart = DEE_Utils.SafeParseDate(CStr(tr(taskStartIdx)))
        End If
        If taskEndIdx >= 0 And taskEndIdx <= UBound(tr) Then
            tEnd = DEE_Utils.SafeParseDate(CStr(tr(taskEndIdx)))
        End If

        Dim tBudget As Double
        If taskBudgetIdx >= 0 And taskBudgetIdx <= UBound(tr) Then
            On Error Resume Next
            tBudget = CDbl(tr(taskBudgetIdx))
            On Error GoTo 0
        End If

        dataArr(rowIdx, 0) = tCode
        dataArr(rowIdx, 1) = tName
        dataArr(rowIdx, 2) = tStart
        dataArr(rowIdx, 3) = tEnd
        dataArr(rowIdx, 4) = tBudget
        dataArr(rowIdx, 5) = tBudget
        rowIdx = rowIdx + 1
NextBLTask:
    Next taskRow

    If rowIdx > 0 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(rowIdx + 1, 6)).Value = dataArr
        ws.Columns(3).NumberFormat = "yyyy-mm-dd"
        ws.Columns(4).NumberFormat = "yyyy-mm-dd"
    End If
End Sub

' ---------------------------------------------------------------------------
' CompareBaseline -- Diff current Schedule vs stored baseline
' ---------------------------------------------------------------------------
Public Sub CompareBaseline()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Check baseline exists
    Dim wsBL As Worksheet
    On Error Resume Next
    Set wsBL = wb.Worksheets(BASELINE_SHEET)
    On Error GoTo 0
    If wsBL Is Nothing Then
        MsgBox "No baseline loaded. Click 'Load Baseline' first.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    ' Check Schedule exists
    Dim wsSch As Worksheet
    On Error Resume Next
    Set wsSch = wb.Worksheets("Schedule")
    On Error GoTo 0
    If wsSch Is Nothing Then
        MsgBox "Schedule sheet not found.", vbCritical, "Protocol DEE"
        Exit Sub
    End If

    DEE_Utils.StartProgress "Comparing Baseline"

    ' Build baseline lookup: task_code -> [BL Start, BL Finish, BL Budget]
    Dim blMap As Object
    Set blMap = CreateObject("Scripting.Dictionary")

    Dim blLastRow As Long
    blLastRow = DEE_Utils.LastRow(wsBL, 1)
    Dim b As Long
    For b = 2 To blLastRow
        Dim blCode As String
        blCode = Trim(wsBL.Cells(b, 1).Value)
        If blCode <> "" Then
            blMap(blCode) = Array(wsBL.Cells(b, 3).Value, _
                                  wsBL.Cells(b, 4).Value, _
                                  wsBL.Cells(b, 5).Value)
        End If
    Next b

    ' Create comparison sheet
    Dim wsDiff As Worksheet
    Set wsDiff = DEE_Utils.GetOrCreateSheet(wb, "Baseline Compare")
    wsDiff.Cells.Clear

    ' Headers
    Dim headers As Variant
    headers = Array("Activity ID", "Activity Name", _
                    "BL Start", "Curr Start", "Start Slip", _
                    "BL Finish", "Curr Finish", "Finish Slip", _
                    "BL Budget", "Curr Budget", "Budget Var", _
                    "Status")
    Dim h As Integer
    For h = 0 To UBound(headers)
        wsDiff.Cells(1, h + 1).Value = headers(h)
    Next h
    DEE_Utils.ApplyTableHeader wsDiff, 1, 1, UBound(headers) + 1

    ' Title block
    wsDiff.Rows(1).Insert
    With wsDiff.Range("A1:L1")
        .Merge
        .Value = "Baseline Comparison Report -- Generated: " & Format(Now(), "yyyy-mm-dd HH:MM")
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
    End With

    Dim schLastRow As Long
    schLastRow = DEE_Utils.LastRow(wsSch, 1)
    Dim outRow As Long
    outRow = 3

    Dim currentCodes As Object
    Set currentCodes = CreateObject("Scripting.Dictionary")

    DEE_Utils.UpdateProgress 30, "Diffing activities"

    Dim i As Long
    For i = 2 To schLastRow
        If DEE_Utils.IsWBSRow(wsSch, i) Then GoTo NextDiff

        Dim currCode As String
        currCode = Trim(wsSch.Cells(i, 1).Value)
        If currCode = "" Then GoTo NextDiff
        currentCodes(currCode) = True

        Dim currName As String
        currName = Trim(wsSch.Cells(i, 2).Value)
        Dim currStart As Variant
        Dim currEnd As Variant
        Dim currBudget As Double
        currStart  = wsSch.Cells(i, 3).Value
        currEnd    = wsSch.Cells(i, 4).Value
        On Error Resume Next
        currBudget = CDbl(wsSch.Cells(i, 6).Value)
        On Error GoTo 0

        ' Write row
        wsDiff.Cells(outRow, 1).Value = currCode
        wsDiff.Cells(outRow, 2).Value = currName

        Dim rowColor As Long
        Dim statusText As String

        If blMap.Exists(currCode) Then
            Dim blData As Variant
            blData = blMap(currCode)
            Dim blStart As Variant
            Dim blEnd As Variant
            Dim blBudget As Double
            blStart  = blData(0)
            blEnd    = blData(1)
            blBudget = CDbl(blData(2))

            ' Start slip
            Dim startSlip As Long
            startSlip = 0
            If IsDate(currStart) And IsDate(blStart) Then
                startSlip = CDate(currStart) - CDate(blStart)
            End If

            ' Finish slip
            Dim finishSlip As Long
            finishSlip = 0
            If IsDate(currEnd) And IsDate(blEnd) Then
                finishSlip = CDate(currEnd) - CDate(blEnd)
            End If

            ' Budget variance
            Dim budgetVar As Double
            budgetVar = currBudget - blBudget

            wsDiff.Cells(outRow, 3).Value = blStart
            wsDiff.Cells(outRow, 4).Value = currStart
            wsDiff.Cells(outRow, 5).Value = startSlip
            wsDiff.Cells(outRow, 6).Value = blEnd
            wsDiff.Cells(outRow, 7).Value = currEnd
            wsDiff.Cells(outRow, 8).Value = finishSlip
            wsDiff.Cells(outRow, 9).Value = blBudget
            wsDiff.Cells(outRow, 10).Value = currBudget
            wsDiff.Cells(outRow, 11).Value = budgetVar

            ' Determine status + color
            Dim maxSlip As Long
            maxSlip = IIf(Abs(startSlip) > Abs(finishSlip), Abs(startSlip), Abs(finishSlip))

            If maxSlip = 0 And budgetVar = 0 Then
                statusText = "On Track"
                rowColor = COLOR_ON_TRACK
            ElseIf maxSlip <= 5 And Abs(budgetVar) / IIf(blBudget > 0, blBudget, 1) < 0.1 Then
                statusText = "Minor Slip"
                rowColor = COLOR_MINOR
            Else
                statusText = "Major Slip"
                rowColor = COLOR_MAJOR
            End If

            ' Color slip cells individually
            If startSlip > 5 Or finishSlip > 5 Then
                wsDiff.Cells(outRow, 5).Font.Color = RGB(255, 0, 0)
                wsDiff.Cells(outRow, 8).Font.Color = RGB(255, 0, 0)
            End If
        Else
            ' New activity (not in baseline)
            wsDiff.Cells(outRow, 3).Value = "N/A"
            wsDiff.Cells(outRow, 4).Value = currStart
            wsDiff.Cells(outRow, 6).Value = "N/A"
            wsDiff.Cells(outRow, 7).Value = currEnd
            wsDiff.Cells(outRow, 9).Value = "N/A"
            wsDiff.Cells(outRow, 10).Value = currBudget
            statusText = "New"
            rowColor = COLOR_NEW
        End If

        wsDiff.Cells(outRow, 12).Value = statusText
        wsDiff.Cells(outRow, 12).Interior.Color = rowColor
        wsDiff.Cells(outRow, 12).Font.Color = RGB(255, 255, 255)
        wsDiff.Cells(outRow, 12).Font.Bold = True

        outRow = outRow + 1
NextDiff:
    Next i

    ' Write removed activities (in baseline, not in current)
    Dim blKey As Variant
    For Each blKey In blMap.Keys
        Dim bk As String
        bk = CStr(blKey)
        If Not currentCodes.Exists(bk) Then
            Dim removedData As Variant
            removedData = blMap(bk)
            wsDiff.Cells(outRow, 1).Value = bk
            wsDiff.Cells(outRow, 2).Value = "(Removed)"
            wsDiff.Cells(outRow, 3).Value = removedData(0)
            wsDiff.Cells(outRow, 6).Value = removedData(1)
            wsDiff.Cells(outRow, 9).Value = removedData(2)
            wsDiff.Cells(outRow, 12).Value = "Removed"
            wsDiff.Range(wsDiff.Cells(outRow, 1), wsDiff.Cells(outRow, 12)).Interior.Color = COLOR_REMOVED
            outRow = outRow + 1
        End If
    Next blKey

    ' Format date columns
    wsDiff.Columns("C:D").NumberFormat = "yyyy-mm-dd"
    wsDiff.Columns("F:G").NumberFormat = "yyyy-mm-dd"
    wsDiff.Columns("I:K").NumberFormat = "$#,##0.00"
    wsDiff.Columns("E:E").NumberFormat = "+0;-0;0"
    wsDiff.Columns("H:H").NumberFormat = "+0;-0;0"
    wsDiff.Columns("A:L").AutoFit

    ' Summary legend
    Dim legRow As Long
    legRow = outRow + 2
    wsDiff.Cells(legRow, 1).Value = "LEGEND:"
    wsDiff.Cells(legRow, 1).Font.Bold = True
    Dim legends As Variant
    legends = Array("On Track (0 days)", "Minor Slip (1-5 days)", "Major Slip (>5 days)", "New Activity", "Removed")
    Dim legColors As Variant
    legColors = Array(COLOR_ON_TRACK, COLOR_MINOR, COLOR_MAJOR, COLOR_NEW, COLOR_REMOVED)
    Dim l As Integer
    For l = 0 To 4
        wsDiff.Cells(legRow, l + 2).Value = legends(l)
        wsDiff.Cells(legRow, l + 2).Interior.Color = legColors(l)
        wsDiff.Cells(legRow, l + 2).Font.Color = RGB(255, 255, 255)
        wsDiff.Cells(legRow, l + 2).Font.Bold = True
    Next l

    DEE_Utils.EndProgress
    wsDiff.Activate

    MsgBox "Baseline comparison complete!" & vbCrLf & _
           (outRow - 3) & " activities compared.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' ClearBaseline -- Remove stored baseline
' ---------------------------------------------------------------------------
Public Sub ClearBaseline()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(BASELINE_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    MsgBox "Baseline cleared.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' HasBaseline -- Returns True if a baseline is stored
' ---------------------------------------------------------------------------
Public Function HasBaseline() As Boolean
    On Error Resume Next
    HasBaseline = Not (ActiveWorkbook.Worksheets(BASELINE_SHEET) Is Nothing)
    On Error GoTo 0
End Function
