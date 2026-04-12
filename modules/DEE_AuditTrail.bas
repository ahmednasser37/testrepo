Attribute VB_Name = "DEE_AuditTrail"
Option Explicit

' =============================================================================
' DEE_AuditTrail.bas -- Change Audit Trail (Section 15.13)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Two-mode audit system:
'
' MODE 1 -- EVENT-DRIVEN (manual cell edits to PMS input columns):
'   Worksheet_Change fires on L, N, O, R, T, U edits.
'   Logs: Timestamp | User | Cell | Old Value | New Value | Data Date
'   Storage: hidden sheet "_DEE_AuditLog"
'
' MODE 2 -- DIFF-ON-DEMAND (XER re-import or snapshot compare):
'   CompareSnapshots() diffs two Snap_yyyy-mm-dd sheets and reports
'   added/removed/changed activities as a formal change log.
'   Better for bulk operations where Worksheet_Change is unreliable.
'
' LOG SHEET: "_DEE_AuditLog"
'   A=Timestamp  B=User  C=Sheet  D=Cell  E=OldValue  F=NewValue
'   G=DataDate   H=ChangeType
' =============================================================================

Private Const AUDIT_SHEET As String = "_DEE_AuditLog"

' PMS input columns that trigger event-driven logging (L,N,O,R,T,U = 12,14,15,18,20,21)
Private Const INPUT_COLS As String = "12,14,15,18,20,21"

' ---------------------------------------------------------------------------
' IsInputColumn -- Returns True if column should be audited
' ---------------------------------------------------------------------------
Public Function IsInputColumn(col As Integer) As Boolean
    Dim cols() As String
    cols = Split(INPUT_COLS, ",")
    Dim c As Integer
    For Each c In cols
        If CInt(c) = col Then
            IsInputColumn = True
            Exit Function
        End If
    Next c
    IsInputColumn = False
End Function

' ---------------------------------------------------------------------------
' LogChange -- Record a manual cell change (called from Worksheet_Change)
' ---------------------------------------------------------------------------
Public Sub LogChange(target As Range, oldValue As Variant)
    On Error Resume Next  ' Never crash the user's edit

    Dim wb As Workbook
    Set wb = target.Parent.Parent

    Dim wsLog As Worksheet
    Set wsLog = GetOrCreateAuditSheet(wb)
    If wsLog Is Nothing Then Exit Sub

    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    Dim dataDate As Date
    dataDate = DEE_Utils.GetDataDate()

    wsLog.Cells(nextRow, 1).Value = Now()
    wsLog.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd HH:MM:SS"
    wsLog.Cells(nextRow, 2).Value = Environ("USERNAME")
    wsLog.Cells(nextRow, 3).Value = target.Parent.Name
    wsLog.Cells(nextRow, 4).Value = target.Address(False, False)
    wsLog.Cells(nextRow, 5).Value = oldValue
    wsLog.Cells(nextRow, 6).Value = target.Value
    wsLog.Cells(nextRow, 7).Value = dataDate
    wsLog.Cells(nextRow, 7).NumberFormat = "yyyy-mm-dd"
    wsLog.Cells(nextRow, 8).Value = "EDIT"

    ' Color-code change type
    wsLog.Cells(nextRow, 8).Interior.Color = RGB(255, 192, 0)

    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' CompareSnapshots -- Diff-on-demand between two Snap_ sheets
' ---------------------------------------------------------------------------
Public Sub CompareSnapshots()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Collect all snapshot sheets
    Dim snapNames() As String
    Dim snapCount As Integer
    snapCount = 0
    ReDim snapNames(0 To wb.Worksheets.Count)

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Left(ws.Name, 5) = "Snap_" Then
            snapNames(snapCount) = ws.Name
            snapCount = snapCount + 1
        End If
    Next ws

    If snapCount < 2 Then
        MsgBox "Need at least 2 snapshots to compare. Take more snapshots first.", _
               vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    ' Build selection list
    Dim listStr As String
    listStr = "Available snapshots:" & vbCrLf & vbCrLf
    Dim s As Integer
    For s = 0 To snapCount - 1
        listStr = listStr & (s + 1) & ") " & snapNames(s) & vbCrLf
    Next s

    Dim fromChoice As String
    fromChoice = InputBox(listStr & vbCrLf & "Enter number for FROM (older) snapshot:", _
                          "Compare Snapshots", "1")
    If fromChoice = "" Then Exit Sub

    Dim toChoice As String
    toChoice = InputBox(listStr & vbCrLf & "Enter number for TO (newer) snapshot:", _
                        "Compare Snapshots", CStr(snapCount))
    If toChoice = "" Then Exit Sub

    Dim fromIdx As Integer
    Dim toIdx As Integer
    On Error Resume Next
    fromIdx = CInt(fromChoice) - 1
    toIdx = CInt(toChoice) - 1
    On Error GoTo 0

    If fromIdx < 0 Or fromIdx >= snapCount Or toIdx < 0 Or toIdx >= snapCount Then
        MsgBox "Invalid selection.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    Dim wsFrom As Worksheet
    Dim wsTo As Worksheet
    Set wsFrom = wb.Worksheets(snapNames(fromIdx))
    Set wsTo = wb.Worksheets(snapNames(toIdx))

    DEE_Utils.StartProgress "Comparing snapshots"

    ' Run diff
    Dim wsDiff As Worksheet
    Set wsDiff = RunSnapshotDiff(wb, wsFrom, wsTo)

    DEE_Utils.EndProgress
    wsDiff.Activate
    MsgBox "Snapshot comparison complete!", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' RunSnapshotDiff -- Core diff engine between two snapshot sheets
' ---------------------------------------------------------------------------
Private Function RunSnapshotDiff(wb As Workbook, wsFrom As Worksheet, wsTo As Worksheet) As Worksheet
    ' Build maps: task_code -> [start, finish, budget, AC%]
    Dim fromMap As Object
    Dim toMap As Object
    Set fromMap = BuildSnapshotMap(wsFrom)
    Set toMap = BuildSnapshotMap(wsTo)

    ' Output sheet
    Dim wsDiff As Worksheet
    Dim diffName As String
    diffName = "ChangeLog_" & Format(Now(), "yyyymmdd_HHmm")
    Set wsDiff = DEE_Utils.GetOrCreateSheet(wb, diffName)
    wsDiff.Cells.Clear

    ' Title
    With wsDiff.Range("A1:I1")
        .Merge
        .Value = "Change Log: " & wsFrom.Name & "  →  " & wsTo.Name
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With

    ' Headers
    Dim hdrs As Variant
    hdrs = Array("Activity ID", "Change Type", "Field", _
                 "Old Value", "New Value", "Delta", _
                 "From Snapshot", "To Snapshot", "Severity")
    Dim h As Integer
    For h = 0 To UBound(hdrs)
        wsDiff.Cells(2, h + 1).Value = hdrs(h)
    Next h
    DEE_Utils.ApplyTableHeader wsDiff, 2, 1, UBound(hdrs) + 1

    Dim outRow As Long
    outRow = 3

    ' Check all TO keys (new + changed)
    Dim k As Variant
    For Each k In toMap.Keys
        Dim tCode As String
        tCode = CStr(k)
        Dim toData As Variant
        toData = toMap(tCode)

        If Not fromMap.Exists(tCode) Then
            ' New activity
            outRow = WriteChangeRow(wsDiff, outRow, tCode, "ADDED", "Activity", _
                                    "", tCode, "", wsFrom.Name, wsTo.Name, "INFO")
        Else
            ' Changed -- compare fields
            Dim fromData As Variant
            fromData = fromMap(tCode)

            ' Start date change
            If CStr(fromData(0)) <> CStr(toData(0)) Then
                Dim startDelta As Long
                startDelta = 0
                If IsDate(fromData(0)) And IsDate(toData(0)) Then
                    startDelta = CDate(toData(0)) - CDate(fromData(0))
                End If
                Dim startSev As String
                startSev = IIf(Abs(startDelta) > 5, "HIGH", IIf(Abs(startDelta) > 0, "MEDIUM", "LOW"))
                outRow = WriteChangeRow(wsDiff, outRow, tCode, "CHANGED", "Start Date", _
                                        CStr(fromData(0)), CStr(toData(0)), _
                                        IIf(startDelta >= 0, "+" & startDelta, CStr(startDelta)) & " days", _
                                        wsFrom.Name, wsTo.Name, startSev)
            End If

            ' Finish date change
            If CStr(fromData(1)) <> CStr(toData(1)) Then
                Dim finDelta As Long
                finDelta = 0
                If IsDate(fromData(1)) And IsDate(toData(1)) Then
                    finDelta = CDate(toData(1)) - CDate(fromData(1))
                End If
                Dim finSev As String
                finSev = IIf(Abs(finDelta) > 5, "HIGH", IIf(Abs(finDelta) > 0, "MEDIUM", "LOW"))
                outRow = WriteChangeRow(wsDiff, outRow, tCode, "CHANGED", "Finish Date", _
                                        CStr(fromData(1)), CStr(toData(1)), _
                                        IIf(finDelta >= 0, "+" & finDelta, CStr(finDelta)) & " days", _
                                        wsFrom.Name, wsTo.Name, finSev)
            End If

            ' Budget change
            Dim fromBudget As Double
            Dim toBudget As Double
            On Error Resume Next
            fromBudget = CDbl(fromData(2))
            toBudget = CDbl(toData(2))
            On Error GoTo 0
            If Abs(fromBudget - toBudget) > 0.01 Then
                Dim budgDelta As Double
                budgDelta = toBudget - fromBudget
                outRow = WriteChangeRow(wsDiff, outRow, tCode, "CHANGED", "Budget", _
                                        Format(fromBudget, "$#,##0.00"), _
                                        Format(toBudget, "$#,##0.00"), _
                                        Format(budgDelta, "+$#,##0.00;-$#,##0.00"), _
                                        wsFrom.Name, wsTo.Name, _
                                        IIf(Abs(budgDelta / IIf(fromBudget <> 0, fromBudget, 1)) > 0.1, "HIGH", "MEDIUM"))
            End If

            ' AC% change (column T = index 3 in map)
            Dim fromAC As Double
            Dim toAC As Double
            On Error Resume Next
            fromAC = CDbl(fromData(3))
            toAC = CDbl(toData(3))
            On Error GoTo 0
            If Abs(fromAC - toAC) > 0.001 Then
                outRow = WriteChangeRow(wsDiff, outRow, tCode, "CHANGED", "AC% Cum", _
                                        Format(fromAC, "0.0%"), Format(toAC, "0.0%"), _
                                        Format(toAC - fromAC, "+0.0%;-0.0%"), _
                                        wsFrom.Name, wsTo.Name, "INFO")
            End If
        End If
    Next k

    ' Check for removed activities
    For Each k In fromMap.Keys
        tCode = CStr(k)
        If Not toMap.Exists(tCode) Then
            outRow = WriteChangeRow(wsDiff, outRow, tCode, "REMOVED", "Activity", _
                                    tCode, "", "", wsFrom.Name, wsTo.Name, "HIGH")
        End If
    Next k

    wsDiff.Columns("A:I").AutoFit

    ' Log this comparison to audit log
    LogDiffRun wb, wsFrom.Name, wsTo.Name, outRow - 3

    Set RunSnapshotDiff = wsDiff
End Function

' ---------------------------------------------------------------------------
' BuildSnapshotMap -- Extract activity data from a snapshot sheet
' Returns: task_code -> Array(Start, Finish, Budget, AC%)
' ---------------------------------------------------------------------------
Private Function BuildSnapshotMap(ws As Worksheet) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)

    ' Skip metadata row if present (hidden row 1 in snapshots)
    Dim startRow As Long
    startRow = 2
    If ws.Rows(1).Hidden Then startRow = 3

    Dim i As Long
    For i = startRow To lastRow
        ' Skip WBS rows (col B empty)
        If Len(Trim(ws.Cells(i, 2).Value)) = 0 Then GoTo NextSnapRow

        Dim tCode As String
        tCode = Trim(ws.Cells(i, 1).Value)
        If tCode = "" Then GoTo NextSnapRow

        Dim acPct As Double
        On Error Resume Next
        acPct = CDbl(ws.Cells(i, 20).Value)  ' Column T: AC% Cum
        On Error GoTo 0

        m(tCode) = Array(ws.Cells(i, 3).Value, _  ' Start
                         ws.Cells(i, 4).Value, _  ' Finish
                         ws.Cells(i, 6).Value, _  ' Budget (F)
                         acPct)                    ' AC%
NextSnapRow:
    Next i

    Set BuildSnapshotMap = m
End Function

' ---------------------------------------------------------------------------
' WriteChangeRow -- Write one change record to the diff sheet
' Returns next available row number
' ---------------------------------------------------------------------------
Private Function WriteChangeRow(ws As Worksheet, rowNum As Long, _
    actCode As String, changeType As String, fieldName As String, _
    oldVal As String, newVal As String, delta As String, _
    fromSnap As String, toSnap As String, severity As String) As Long

    ws.Cells(rowNum, 1).Value = actCode
    ws.Cells(rowNum, 2).Value = changeType
    ws.Cells(rowNum, 3).Value = fieldName
    ws.Cells(rowNum, 4).Value = oldVal
    ws.Cells(rowNum, 5).Value = newVal
    ws.Cells(rowNum, 6).Value = delta
    ws.Cells(rowNum, 7).Value = fromSnap
    ws.Cells(rowNum, 8).Value = toSnap
    ws.Cells(rowNum, 9).Value = severity

    ' Color severity column
    Dim sevColor As Long
    Select Case UCase(severity)
        Case "HIGH":   sevColor = RGB(255, 0, 0)
        Case "MEDIUM": sevColor = RGB(255, 192, 0)
        Case "INFO":   sevColor = RGB(0, 112, 192)
        Case Else:     sevColor = RGB(200, 200, 200)
    End Select
    ws.Cells(rowNum, 9).Interior.Color = sevColor
    ws.Cells(rowNum, 9).Font.Color = RGB(255, 255, 255)
    ws.Cells(rowNum, 9).Font.Bold = True

    ' Alternate row shading
    If rowNum Mod 2 = 0 Then
        ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(245, 245, 245)
    End If

    WriteChangeRow = rowNum + 1
End Function

' ---------------------------------------------------------------------------
' LogDiffRun -- Record a diff operation in the audit log
' ---------------------------------------------------------------------------
Private Sub LogDiffRun(wb As Workbook, fromSnap As String, toSnap As String, changeCount As Long)
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = GetOrCreateAuditSheet(wb)
    If wsLog Is Nothing Then Exit Sub

    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    wsLog.Cells(nextRow, 1).Value = Now()
    wsLog.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd HH:MM:SS"
    wsLog.Cells(nextRow, 2).Value = Environ("USERNAME")
    wsLog.Cells(nextRow, 3).Value = "Diff"
    wsLog.Cells(nextRow, 4).Value = fromSnap & " → " & toSnap
    wsLog.Cells(nextRow, 5).Value = ""
    wsLog.Cells(nextRow, 6).Value = changeCount & " changes found"
    wsLog.Cells(nextRow, 7).Value = DEE_Utils.GetDataDate()
    wsLog.Cells(nextRow, 7).NumberFormat = "yyyy-mm-dd"
    wsLog.Cells(nextRow, 8).Value = "DIFF"
    wsLog.Cells(nextRow, 8).Interior.Color = RGB(0, 112, 192)
    wsLog.Cells(nextRow, 8).Font.Color = RGB(255, 255, 255)
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' ShowAuditLog -- Unhide and activate the audit log sheet
' ---------------------------------------------------------------------------
Public Sub ShowAuditLog()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim wsLog As Worksheet
    Set wsLog = GetOrCreateAuditSheet(wb)
    wsLog.Visible = xlSheetVisible
    wsLog.Activate
    wsLog.Columns("A:H").AutoFit
End Sub

' ---------------------------------------------------------------------------
' GetOrCreateAuditSheet -- Return hidden audit log sheet
' ---------------------------------------------------------------------------
Private Function GetOrCreateAuditSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(AUDIT_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = AUDIT_SHEET
        ws.Visible = xlSheetVeryHidden
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "User"
        ws.Cells(1, 3).Value = "Sheet"
        ws.Cells(1, 4).Value = "Cell / Context"
        ws.Cells(1, 5).Value = "Old Value"
        ws.Cells(1, 6).Value = "New Value"
        ws.Cells(1, 7).Value = "Data Date"
        ws.Cells(1, 8).Value = "Change Type"
        DEE_Utils.ApplyTableHeader ws, 1, 1, 8
    End If

    Set GetOrCreateAuditSheet = ws
End Function
