Attribute VB_Name = "DEE_AuditTrail"
Option Explicit

' =============================================================================
' DEE_AuditTrail.bas -- Diff-on-Demand Audit Trail (P1 rewrite)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Compares any two snapshot sheets (Snap_yyyy-mm-dd or _DEE_Version_NNN) and
' writes a structured change log to "_DEE_Audit" hidden sheet.
'
' IMPACT LEVELS:
'   HIGH   -- Finish date change > 5 days, activity added/deleted, logic change
'   MEDIUM -- Finish date change 1-5 days, % complete change > 10%, budget delta > 10%
'   LOW    -- Start date change, duration change, resource change, notes edit
'
' AUDIT SHEET "_DEE_Audit" columns:
'   A  Timestamp  B  User  C  Impact  D  ChangeType  E  TaskCode  F  Field
'   G  FromValue  H  ToValue  I  Delta  J  FromSheet  K  ToSheet
' =============================================================================

Private Const AUDIT_SHEET As String = "_DEE_Audit"

' ---------------------------------------------------------------------------
' CompareSnapshots -- Diff two sheets chosen by user; write to _DEE_Audit
' ---------------------------------------------------------------------------
Public Sub CompareSnapshots()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Collect candidate sheet names
    Dim candidates() As String
    Dim count As Integer
    count = 0

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim n As String
        n = ws.Name
        If Left(n, 5) = "Snap_" Or Left(n, Len("_DEE_Version_")) = "_DEE_Version_" Or _
           n = DEE_Config.SHEET_SCHEDULE Then
            ReDim Preserve candidates(count)
            candidates(count) = n
            count = count + 1
        End If
    Next ws

    If count < 2 Then
        MsgBox "Need at least 2 snapshot sheets to compare." & vbCrLf & _
               "Use 'Snapshot' or 'Save Version' to create snapshots.", _
               vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    ' Build selection list
    Dim listStr As String
    Dim i As Integer
    For i = 0 To count - 1
        listStr = listStr & (i + 1) & ") " & candidates(i) & vbCrLf
    Next i

    Dim fromChoice As String
    fromChoice = InputBox("Select FROM (baseline) sheet number:" & vbCrLf & vbCrLf & listStr, _
                          "Compare Snapshots -- FROM", "1")
    If fromChoice = "" Then Exit Sub

    Dim toChoice As String
    toChoice = InputBox("Select TO (current) sheet number:" & vbCrLf & vbCrLf & listStr, _
                        "Compare Snapshots -- TO", CStr(count))
    If toChoice = "" Then Exit Sub

    Dim fromIdx As Integer
    Dim toIdx As Integer
    On Error Resume Next
    fromIdx = CInt(fromChoice) - 1
    toIdx   = CInt(toChoice)   - 1
    On Error GoTo 0

    If fromIdx < 0 Or fromIdx >= count Or toIdx < 0 Or toIdx >= count Then
        MsgBox "Invalid selection.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    Dim fromSheet As Worksheet
    Dim toSheet As Worksheet
    Set fromSheet = wb.Worksheets(candidates(fromIdx))
    Set toSheet   = wb.Worksheets(candidates(toIdx))

    DEE_Logger.LogInfo "DEE_AuditTrail", "CompareSnapshots: " & fromSheet.Name & " -> " & toSheet.Name
    DEE_Utils.StartProgress "Comparing snapshots"

    Dim auditWs As Worksheet
    Set auditWs = GetAuditSheet(wb)

    RunDiff fromSheet, toSheet, auditWs

    auditWs.Visible = xlSheetVisible
    auditWs.Activate
    DEE_Utils.EndProgress
    MsgBox "Snapshot comparison complete. See '" & AUDIT_SHEET & "' sheet.", _
           vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' RunDiff -- Core field-level diff engine
' ---------------------------------------------------------------------------
Private Sub RunDiff(fromWs As Worksheet, toWs As Worksheet, auditWs As Worksheet)
    ' Build maps from both sheets: task_code -> row data array
    Dim fromMap As Object
    Set fromMap = BuildScheduleMap(fromWs)
    Dim toMap As Object
    Set toMap = BuildScheduleMap(toWs)

    ' Find the audit sheet's next write row
    Dim outRow As Long
    outRow = DEE_Utils.LastRow(auditWs, 1)
    If outRow < 2 Then
        WriteAuditHeader auditWs
        outRow = 2
    Else
        outRow = outRow + 1
    End If

    Dim ts As String
    ts = Format(Now, "yyyy-mm-dd hh:mm:ss")
    Dim userName As String
    On Error Resume Next
    userName = Environ("USERNAME")
    On Error GoTo 0
    If userName = "" Then userName = "Unknown"

    ' Activities in FROM
    Dim fromKey As Variant
    For Each fromKey In fromMap.Keys
        Dim code As String
        code = CStr(fromKey)

        Dim fromData As Variant
        fromData = fromMap(code)

        If toMap.Exists(code) Then
            Dim toData As Variant
            toData = toMap(code)
            outRow = DiffActivity(auditWs, outRow, ts, userName, code, fromData, toData, fromWs.Name, toWs.Name)
        Else
            ' DELETED
            outRow = WriteAuditRow(auditWs, outRow, ts, userName, "HIGH", "DELETED", code, _
                                   "task_code", code, "", 0, fromWs.Name, toWs.Name)
        End If
    Next fromKey

    ' Activities in TO but not FROM = NEW
    Dim toKey As Variant
    For Each toKey In toMap.Keys
        code = CStr(toKey)
        If Not fromMap.Exists(code) Then
            outRow = WriteAuditRow(auditWs, outRow, ts, userName, "HIGH", "ADDED", code, _
                                   "task_code", "", code, 0, fromWs.Name, toWs.Name)
        End If
    Next toKey
End Sub

' ---------------------------------------------------------------------------
' DiffActivity -- Compare field by field; write a row for each changed field
' Returns next available output row
' ---------------------------------------------------------------------------
Private Function DiffActivity(auditWs As Worksheet, startRow As Long, _
                               ts As String, userName As String, code As String, _
                               fromData As Variant, toData As Variant, _
                               fromShName As String, toShName As String) As Long
    ' fromData / toData: Array(name, start, finish, dur, pct, budget, pred)
    Dim r As Long
    r = startRow

    ' Field index constants (matches BuildScheduleMap order)
    Const FLD_NAME   As Integer = 0
    Const FLD_START  As Integer = 1
    Const FLD_FINISH As Integer = 2
    Const FLD_DUR    As Integer = 3
    Const FLD_PCT    As Integer = 4
    Const FLD_BUDGET As Integer = 5
    Const FLD_PRED   As Integer = 6

    ' Finish date
    Dim blFinish As Date
    Dim actFinish As Date
    On Error Resume Next
    If fromData(FLD_FINISH) <> 0 And Not IsEmpty(fromData(FLD_FINISH)) Then blFinish = CDate(fromData(FLD_FINISH))
    If toData(FLD_FINISH) <> 0 And Not IsEmpty(toData(FLD_FINISH)) Then actFinish = CDate(toData(FLD_FINISH))
    On Error GoTo 0

    If blFinish <> actFinish And (blFinish <> 0 Or actFinish <> 0) Then
        Dim finSlip As Long
        If blFinish <> 0 And actFinish <> 0 Then finSlip = CLng(actFinish - blFinish)
        Dim impact As String
        impact = IIf(Abs(finSlip) > 5, "HIGH", "MEDIUM")
        r = WriteAuditRow(auditWs, r, ts, userName, impact, "DATE_CHANGE", code, "target_end_date", _
                          Format(blFinish, "dd-mmm-yy"), Format(actFinish, "dd-mmm-yy"), finSlip, _
                          fromShName, toShName)
    End If

    ' Start date
    Dim blStart As Date
    Dim actStart As Date
    On Error Resume Next
    If fromData(FLD_START) <> 0 And Not IsEmpty(fromData(FLD_START)) Then blStart = CDate(fromData(FLD_START))
    If toData(FLD_START) <> 0 And Not IsEmpty(toData(FLD_START)) Then actStart = CDate(toData(FLD_START))
    On Error GoTo 0

    If blStart <> actStart And (blStart <> 0 Or actStart <> 0) Then
        Dim startSlip As Long
        If blStart <> 0 And actStart <> 0 Then startSlip = CLng(actStart - blStart)
        r = WriteAuditRow(auditWs, r, ts, userName, "LOW", "DATE_CHANGE", code, "target_start_date", _
                          Format(blStart, "dd-mmm-yy"), Format(actStart, "dd-mmm-yy"), startSlip, _
                          fromShName, toShName)
    End If

    ' % complete
    Dim blPct As Double
    Dim actPct As Double
    On Error Resume Next
    blPct  = CDbl(fromData(FLD_PCT))
    actPct = CDbl(toData(FLD_PCT))
    On Error GoTo 0
    If Abs(blPct - actPct) > 0.001 Then
        Dim pctImpact As String
        pctImpact = IIf(Abs(actPct - blPct) > 0.1, "MEDIUM", "LOW")
        r = WriteAuditRow(auditWs, r, ts, userName, pctImpact, "PCT_CHANGE", code, "phys_complete_pct", _
                          Format(blPct, "0.0%"), Format(actPct, "0.0%"), actPct - blPct, _
                          fromShName, toShName)
    End If

    ' Budget
    Dim blBudget As Double
    Dim actBudget As Double
    On Error Resume Next
    blBudget  = CDbl(fromData(FLD_BUDGET))
    actBudget = CDbl(toData(FLD_BUDGET))
    On Error GoTo 0
    If Abs(blBudget - actBudget) > 0.01 Then
        Dim budgetDelta As Double
        budgetDelta = actBudget - blBudget
        Dim budgetPct As Double
        If blBudget <> 0 Then budgetPct = budgetDelta / Abs(blBudget) Else budgetPct = 1
        Dim budgetImpact As String
        budgetImpact = IIf(Abs(budgetPct) > 0.1, "MEDIUM", "LOW")
        r = WriteAuditRow(auditWs, r, ts, userName, budgetImpact, "BUDGET_CHANGE", code, "budget_qty", _
                          Format(blBudget, "#,##0.00"), Format(actBudget, "#,##0.00"), budgetDelta, _
                          fromShName, toShName)
    End If

    ' Predecessor logic
    Dim blPred As String:  blPred  = Trim(CStr(fromData(FLD_PRED)))
    Dim actPred As String: actPred = Trim(CStr(toData(FLD_PRED)))
    If blPred <> actPred And (blPred <> "" Or actPred <> "") Then
        r = WriteAuditRow(auditWs, r, ts, userName, "HIGH", "LOGIC_CHANGE", code, "predecessors", _
                          blPred, actPred, 0, fromShName, toShName)
    End If

    DiffActivity = r
End Function

' ---------------------------------------------------------------------------
' BuildScheduleMap -- task_code -> Array(name, start, finish, dur, pct, budget, pred)
' Skips WBS summary rows (col B empty)
' ---------------------------------------------------------------------------
Private Function BuildScheduleMap(ws As Worksheet) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, DEE_Config.COL_WBS_ID)
    If lastRow < 2 Then
        Set BuildScheduleMap = m
        Exit Function
    End If

    Dim r As Long
    For r = 2 To lastRow
        Dim code As String
        code = Trim(CStr(ws.Cells(r, DEE_Config.COL_WBS_ID).Value))
        If code = "" Then GoTo NextRow

        Dim taskName As String
        taskName = Trim(CStr(ws.Cells(r, DEE_Config.COL_TASK_NAME).Value))
        If taskName = "" Then GoTo NextRow  ' WBS summary row

        m(code) = Array( _
            taskName, _
            ws.Cells(r, DEE_Config.COL_START).Value, _
            ws.Cells(r, DEE_Config.COL_FINISH).Value, _
            ws.Cells(r, DEE_Config.COL_DUR).Value, _
            ws.Cells(r, DEE_Config.COL_PCT).Value, _
            ws.Cells(r, DEE_Config.COL_BUDGET).Value, _
            ws.Cells(r, DEE_Config.COL_PRED).Value)
NextRow:
    Next r

    Set BuildScheduleMap = m
End Function

' ---------------------------------------------------------------------------
' WriteAuditRow -- Append one row to the audit sheet; return next row number
' ---------------------------------------------------------------------------
Private Function WriteAuditRow(ws As Worksheet, r As Long, _
                                ts As String, userName As String, _
                                impact As String, changeType As String, _
                                taskCode As String, fieldName As String, _
                                fromVal As String, toVal As String, delta As Double, _
                                fromShName As String, toShName As String) As Long
    ws.Cells(r, 1).Value  = ts
    ws.Cells(r, 2).Value  = userName
    ws.Cells(r, 3).Value  = impact
    ws.Cells(r, 4).Value  = changeType
    ws.Cells(r, 5).Value  = taskCode
    ws.Cells(r, 6).Value  = fieldName
    ws.Cells(r, 7).Value  = fromVal
    ws.Cells(r, 8).Value  = toVal
    ws.Cells(r, 9).Value  = delta
    ws.Cells(r, 10).Value = fromShName
    ws.Cells(r, 11).Value = toShName

    ' Colour-code impact
    Dim impactColor As Long
    Select Case impact
        Case "HIGH":   impactColor = RGB(255, 180, 180)
        Case "MEDIUM": impactColor = RGB(255, 230, 130)
        Case "LOW":    impactColor = RGB(210, 240, 210)
        Case Else:     impactColor = RGB(240, 240, 240)
    End Select
    ws.Cells(r, 3).Interior.Color = impactColor
    If impact = "HIGH" Then ws.Cells(r, 3).Font.Bold = True

    WriteAuditRow = r + 1
End Function

' ---------------------------------------------------------------------------
' WriteAuditHeader -- Write column headers on first use
' ---------------------------------------------------------------------------
Private Sub WriteAuditHeader(ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:K1").Value = Array("Timestamp", "User", "Impact", "ChangeType", "TaskCode", "Field", _
                                     "FromValue", "ToValue", "Delta", "FromSheet", "ToSheet")
    ws.Range("A1:K1").Font.Bold = True
    ws.Range("A1:K1").Interior.Color = RGB(0, 32, 96)
    ws.Range("A1:K1").Font.Color = RGB(255, 255, 255)
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B:D").ColumnWidth = 14
    ws.Columns("E:F").ColumnWidth = 16
    ws.Columns("G:H").ColumnWidth = 20
    ws.Columns("I").ColumnWidth = 10
    ws.Columns("J:K").ColumnWidth = 20
End Sub

' ---------------------------------------------------------------------------
' GetAuditSheet -- Return (creating if needed) the _DEE_Audit sheet
' ---------------------------------------------------------------------------
Private Function GetAuditSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(AUDIT_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = AUDIT_SHEET
        ws.Visible = xlSheetVeryHidden
        WriteAuditHeader ws
    End If

    Set GetAuditSheet = ws
End Function

' ---------------------------------------------------------------------------
' ShowAuditLog -- Make the audit sheet visible and activate it
' ---------------------------------------------------------------------------
Public Sub ShowAuditLog()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim ws As Worksheet
    Set ws = GetAuditSheet(wb)

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        MsgBox "Audit log is empty. Use 'Diff Snapshots' to generate entries.", _
               vbInformation, "Protocol DEE"
        Exit Sub
    End If

    ws.Visible = xlSheetVisible
    ws.Activate

    MsgBox lastRow - 1 & " audit entries recorded.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' ExportAuditCSV -- Export audit sheet to a CSV file
' ---------------------------------------------------------------------------
Public Sub ExportAuditCSV()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim ws As Worksheet
    Set ws = GetAuditSheet(wb)

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        MsgBox "Audit log is empty.", vbInformation, "Protocol DEE"
        Exit Sub
    End If

    Dim savePath As String
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="DEE_AuditTrail_" & Format(Now, "yyyymmdd") & ".csv", _
        FileFilter:="CSV Files (*.csv),*.csv")
    If savePath = "False" Or savePath = "" Then Exit Sub

    Dim fNum As Integer
    fNum = FreeFile
    Open savePath For Output As #fNum
    Print #fNum, "Timestamp,User,Impact,ChangeType,TaskCode,Field,FromValue,ToValue,Delta,FromSheet,ToSheet"

    Dim r As Long
    For r = 2 To lastRow
        Dim line As String
        Dim c As Integer
        line = ""
        For c = 1 To 11
            Dim cellVal As String
            cellVal = Replace(CStr(ws.Cells(r, c).Value), """", """""")
            If c > 1 Then line = line & ","
            line = line & """" & cellVal & """"
        Next c
        Print #fNum, line
    Next r

    Close #fNum
    MsgBox "Audit log exported to:" & vbCrLf & savePath, vbInformation, "Protocol DEE"
End Sub
