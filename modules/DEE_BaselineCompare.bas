Attribute VB_Name = "DEE_BaselineCompare"
Option Explicit

' =============================================================================
' DEE_BaselineCompare.bas -- Baseline Comparison (P1 rewrite)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Loads a Primavera XER file as a named baseline. Multiple named baselines
' are supported. Each baseline is stored as a hidden sheet prefixed
' "_DEE_Baseline_<name>" and its file path is persisted in CustomDocumentProperties.
'
' COMPARE OUTPUT: "Baseline_Variance" sheet with columns:
'   A  task_code
'   B  task_name
'   C  BL Start        D  BL Finish        E  BL Duration
'   F  ACT Start       G  ACT Finish       H  ACT Duration
'   I  Start Slip (d)  J  Finish Slip (d)  K  Duration Delta
'   L  Status flag     M  % Complete       N  Notes
'
' STATUS FLAGS:
'   NEW          -- activity exists in current but not in baseline
'   DELETED      -- activity exists in baseline but not in current
'   ON_TRACK     -- finish slip <= 0 days
'   SLIPPED      -- finish slip > 0 (positive = later)
'   AHEAD        -- finish slip < 0 (negative = earlier)
'   LOGIC_CHANGED -- predecessor list differs between baseline and current
' =============================================================================

Private Const CDP_PREFIX      As String = "DEE_Baseline_"  ' CustomDocumentProperty prefix
Private Const SH_PREFIX       As String = "_DEE_Baseline_" ' Sheet name prefix

Private Const COL_BL_CODE     As Integer = 1
Private Const COL_BL_NAME     As Integer = 2
Private Const COL_BL_START    As Integer = 3
Private Const COL_BL_FINISH   As Integer = 4
Private Const COL_BL_DUR      As Integer = 5
Private Const COL_BL_PCT      As Integer = 6
Private Const COL_BL_PRED     As Integer = 7

Private Const COLOR_ON_TRACK  As Long = RGB(0,   176,  80)
Private Const COLOR_SLIPPED   As Long = RGB(255,   0,   0)
Private Const COLOR_AHEAD     As Long = RGB(0,   112, 192)
Private Const COLOR_NEW       As Long = RGB(146, 208,  80)
Private Const COLOR_DELETED   As Long = RGB(191, 191, 191)
Private Const COLOR_LOGIC     As Long = RGB(255, 192,   0)

' ---------------------------------------------------------------------------
' LoadBaseline -- Pick a XER, parse it, ask for a name, store hidden
' ---------------------------------------------------------------------------
Public Sub LoadBaseline()
    Dim filePath As String
    filePath = DEE_XERParser.GetXERFilePath()
    If filePath = "" Then Exit Sub

    Dim baselineName As String
    baselineName = InputBox("Enter a name for this baseline:" & vbCrLf & "(e.g. B1, Tender, Contract)", _
                            "Load Baseline", "B1")
    If Trim(baselineName) = "" Then Exit Sub
    baselineName = Left(Trim(baselineName), 30)

    DEE_Logger.LogInfo "DEE_BaselineCompare", "LoadBaseline: " & baselineName & " from " & filePath
    DEE_Utils.StartProgress "Loading Baseline: " & baselineName

    Dim success As Boolean
    success = DEE_XERParser.ParseXERFile(filePath)
    If Not success Then
        DEE_Utils.EndProgress
        MsgBox "Failed to parse XER file.", vbCritical, "Protocol DEE"
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Remove old sheet with same name if it exists
    Dim sheetName As String
    sheetName = SH_PREFIX & baselineName
    DeleteSheetIfExists wb, sheetName

    ' Build the baseline data sheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add
    ws.Name = sheetName
    ws.Visible = xlSheetVeryHidden

    BuildBaselineSheet ws

    ' Persist file path in CustomDocumentProperties for reference
    StoreBaselineCDP wb, baselineName, filePath

    DEE_Utils.EndProgress
    DEE_Logger.LogInfo "DEE_BaselineCompare", "Baseline '" & baselineName & "' stored on sheet " & sheetName

    MsgBox "Baseline '" & baselineName & "' loaded successfully." & vbCrLf & _
           "Activities stored: " & (DEE_Utils.LastRow(ws, 1) - 1), _
           vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' BuildBaselineSheet -- Extract TASK rows from parsed XER into hidden sheet
' ---------------------------------------------------------------------------
Private Sub BuildBaselineSheet(ws As Worksheet)
    ' Write header
    ws.Cells(1, COL_BL_CODE).Value   = "task_code"
    ws.Cells(1, COL_BL_NAME).Value   = "task_name"
    ws.Cells(1, COL_BL_START).Value  = "target_start_date"
    ws.Cells(1, COL_BL_FINISH).Value = "target_end_date"
    ws.Cells(1, COL_BL_DUR).Value    = "target_drtn_hr_cnt"
    ws.Cells(1, COL_BL_PCT).Value    = "phys_complete_pct"
    ws.Cells(1, COL_BL_PRED).Value   = "predecessors"

    If Not DEE_XERParser.TableExists("TASK") Then Exit Sub

    Dim taskFields As Variant
    taskFields = DEE_XERParser.GetFieldNames("TASK")
    Dim taskRows As Collection
    Set taskRows = DEE_XERParser.GetTable("TASK")

    Dim codeIdx As Integer:  codeIdx  = DEE_XERParser.FindFieldIndex(taskFields, "task_code")
    Dim nameIdx As Integer:  nameIdx  = DEE_XERParser.FindFieldIndex(taskFields, "task_name")
    Dim startIdx As Integer: startIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_start_date")
    Dim endIdx As Integer:   endIdx   = DEE_XERParser.FindFieldIndex(taskFields, "target_end_date")
    Dim durIdx As Integer:   durIdx   = DEE_XERParser.FindFieldIndex(taskFields, "target_drtn_hr_cnt")
    Dim pctIdx As Integer:   pctIdx   = DEE_XERParser.FindFieldIndex(taskFields, "phys_complete_pct")

    ' Build predecessor map from TASKPRED
    Dim predMap As Object
    Set predMap = BuildPredMap()

    Dim r As Long
    r = 2
    Dim taskRow As Variant
    For Each taskRow In taskRows
        Dim tr As Variant
        tr = taskRow

        Dim tCode As String
        If codeIdx >= 0 And codeIdx <= UBound(tr) Then tCode = Trim(CStr(tr(codeIdx)))
        If tCode = "" Then GoTo NextRow

        Dim tName As String
        If nameIdx >= 0 And nameIdx <= UBound(tr) Then tName = Trim(CStr(tr(nameIdx)))

        Dim tStart As Date
        Dim tEnd As Date
        Dim tDur As Double
        Dim tPct As Double

        If startIdx >= 0 And startIdx <= UBound(tr) Then
            tStart = DEE_Utils.SafeParseDate(Trim(CStr(tr(startIdx))))
        End If
        If endIdx >= 0 And endIdx <= UBound(tr) Then
            tEnd = DEE_Utils.SafeParseDate(Trim(CStr(tr(endIdx))))
        End If
        If durIdx >= 0 And durIdx <= UBound(tr) Then
            On Error Resume Next
            tDur = CDbl(tr(durIdx)) / 8
            On Error GoTo 0
        End If
        If pctIdx >= 0 And pctIdx <= UBound(tr) Then
            On Error Resume Next
            tPct = CDbl(tr(pctIdx))
            On Error GoTo 0
        End If

        ws.Cells(r, COL_BL_CODE).Value   = tCode
        ws.Cells(r, COL_BL_NAME).Value   = tName
        If tStart <> 0 Then ws.Cells(r, COL_BL_START).Value = tStart
        If tEnd   <> 0 Then ws.Cells(r, COL_BL_FINISH).Value = tEnd
        ws.Cells(r, COL_BL_DUR).Value    = tDur
        ws.Cells(r, COL_BL_PCT).Value    = tPct

        If predMap.Exists(tCode) Then
            ws.Cells(r, COL_BL_PRED).Value = predMap(tCode)
        End If

        r = r + 1
NextRow:
    Next taskRow
End Sub

' ---------------------------------------------------------------------------
' BuildPredMap -- task_code -> comma-separated predecessor list from TASKPRED
' ---------------------------------------------------------------------------
Private Function BuildPredMap() As Object
    Dim pm As Object
    Set pm = CreateObject("Scripting.Dictionary")

    If Not DEE_XERParser.TableExists("TASKPRED") Then
        Set BuildPredMap = pm
        Exit Function
    End If

    Dim predFields As Variant
    predFields = DEE_XERParser.GetFieldNames("TASKPRED")
    Dim predRows As Collection
    Set predRows = DEE_XERParser.GetTable("TASKPRED")

    Dim succIdx As Integer: succIdx = DEE_XERParser.FindFieldIndex(predFields, "task_id")
    Dim predIdx As Integer: predIdx = DEE_XERParser.FindFieldIndex(predFields, "pred_task_id")
    Dim typeIdx As Integer: typeIdx = DEE_XERParser.FindFieldIndex(predFields, "pred_type")

    Dim predRow As Variant
    For Each predRow In predRows
        Dim pr As Variant
        pr = predRow

        Dim succId As String
        Dim predId As String
        If succIdx >= 0 And succIdx <= UBound(pr) Then succId = Trim(CStr(pr(succIdx)))
        If predIdx >= 0 And predIdx <= UBound(pr) Then predId = Trim(CStr(pr(predIdx)))
        If succId = "" Or predId = "" Then GoTo NextPred

        Dim rel As String
        rel = predId
        If typeIdx >= 0 And typeIdx <= UBound(pr) Then
            Dim relType As String
            relType = Trim(CStr(pr(typeIdx)))
            If relType <> "" And relType <> "PR_FS" Then rel = predId & "(" & relType & ")"
        End If

        If pm.Exists(succId) Then
            pm(succId) = pm(succId) & "," & rel
        Else
            pm(succId) = rel
        End If
NextPred:
    Next predRow

    Set BuildPredMap = pm
End Function

' ---------------------------------------------------------------------------
' CompareBaseline -- Diff current Schedule vs a stored baseline
' ---------------------------------------------------------------------------
Public Sub CompareBaseline()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim names As Variant
    names = ListBaselineNames(wb)

    If IsEmpty(names) Then
        MsgBox "No baselines loaded. Use 'Load Baseline' first.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    Dim baselineName As String
    If UBound(names) = 0 Then
        baselineName = CStr(names(0))
    Else
        Dim listStr As String
        Dim i As Integer
        For i = 0 To UBound(names)
            listStr = listStr & (i + 1) & ") " & names(i) & vbCrLf
        Next i
        Dim choice As String
        choice = InputBox("Select baseline to compare:" & vbCrLf & vbCrLf & listStr, _
                          "Compare Baseline", "1")
        If choice = "" Then Exit Sub
        Dim choiceNum As Integer
        On Error Resume Next
        choiceNum = CInt(choice) - 1
        On Error GoTo 0
        If choiceNum < 0 Or choiceNum > UBound(names) Then Exit Sub
        baselineName = CStr(names(choiceNum))
    End If

    ' Load baseline sheet
    Dim blSheet As Worksheet
    On Error Resume Next
    Set blSheet = wb.Worksheets(SH_PREFIX & baselineName)
    On Error GoTo 0

    If blSheet Is Nothing Then
        MsgBox "Baseline sheet not found for '" & baselineName & "'.", vbCritical, "Protocol DEE"
        Exit Sub
    End If

    ' Load current Schedule sheet
    Dim schedSheet As Worksheet
    On Error Resume Next
    Set schedSheet = wb.Worksheets(DEE_Config.SHEET_SCHEDULE)
    On Error GoTo 0

    If schedSheet Is Nothing Then
        MsgBox "Schedule sheet not found.", vbCritical, "Protocol DEE"
        Exit Sub
    End If

    DEE_Logger.LogInfo "DEE_BaselineCompare", "CompareBaseline: " & baselineName
    DEE_Utils.StartProgress "Comparing baseline: " & baselineName

    ' Build maps
    Dim blMap As Object
    Set blMap = BuildBaselineMap(blSheet)

    Dim actMap As Object
    Set actMap = BuildCurrentMap(schedSheet)

    ' Create output sheet
    Dim outSheet As Worksheet
    Set outSheet = DEE_Utils.GetOrCreateSheet(wb, DEE_Config.SHEET_BASELINE)
    outSheet.Cells.Clear

    WriteVarianceHeader outSheet, baselineName
    WriteVarianceRows outSheet, blMap, actMap

    outSheet.Columns("A:N").AutoFit
    outSheet.Columns("B").ColumnWidth = 40
    outSheet.Activate

    DEE_Utils.EndProgress
    DEE_Logger.LogInfo "DEE_BaselineCompare", "CompareBaseline complete"
    MsgBox "Baseline comparison complete. See '" & DEE_Config.SHEET_BASELINE & "' sheet.", _
           vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' BuildBaselineMap -- task_code -> Array(name, start, finish, dur, pred)
' ---------------------------------------------------------------------------
Private Function BuildBaselineMap(ws As Worksheet) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    Dim r As Long
    For r = 2 To lastRow
        Dim code As String
        code = Trim(CStr(ws.Cells(r, COL_BL_CODE).Value))
        If code <> "" Then
            m(code) = Array( _
                ws.Cells(r, COL_BL_NAME).Value, _
                ws.Cells(r, COL_BL_START).Value, _
                ws.Cells(r, COL_BL_FINISH).Value, _
                ws.Cells(r, COL_BL_DUR).Value, _
                ws.Cells(r, COL_BL_PCT).Value, _
                ws.Cells(r, COL_BL_PRED).Value)
        End If
    Next r

    Set BuildBaselineMap = m
End Function

' ---------------------------------------------------------------------------
' BuildCurrentMap -- task_code -> Array(name, start, finish, dur, pct, pred)
' Uses canonical Schedule column constants from DEE_Config
' ---------------------------------------------------------------------------
Private Function BuildCurrentMap(ws As Worksheet) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, DEE_Config.COL_WBS_ID)
    Dim r As Long
    For r = 2 To lastRow
        Dim code As String
        code = Trim(CStr(ws.Cells(r, DEE_Config.COL_WBS_ID).Value))
        If code = "" Then GoTo NextRow

        ' Skip WBS summary rows (col B empty)
        Dim taskName As String
        taskName = Trim(CStr(ws.Cells(r, DEE_Config.COL_TASK_NAME).Value))
        If taskName = "" Then GoTo NextRow

        m(code) = Array( _
            taskName, _
            ws.Cells(r, DEE_Config.COL_START).Value, _
            ws.Cells(r, DEE_Config.COL_FINISH).Value, _
            ws.Cells(r, DEE_Config.COL_DUR).Value, _
            ws.Cells(r, DEE_Config.COL_PCT).Value, _
            ws.Cells(r, DEE_Config.COL_PRED).Value)
NextRow:
    Next r

    Set BuildCurrentMap = m
End Function

' ---------------------------------------------------------------------------
' WriteVarianceHeader -- Write output sheet column headers
' ---------------------------------------------------------------------------
Private Sub WriteVarianceHeader(ws As Worksheet, baselineName As String)
    With ws.Range("A1:N1")
        .Merge
        .Value = "Baseline Variance Report -- Baseline: " & baselineName & "  |  Data Date: " & Format(Date, "dd-MMM-yyyy")
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .RowHeight = 24
    End With

    Dim headers As Variant
    headers = Array("Task Code", "Task Name", _
                    "BL Start", "BL Finish", "BL Dur (d)", _
                    "ACT Start", "ACT Finish", "ACT Dur (d)", _
                    "Start Slip (d)", "Finish Slip (d)", "Dur Delta (d)", _
                    "Status", "% Complete", "Notes")

    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(2, i + 1).Value = headers(i)
    Next i

    DEE_Utils.ApplyTableHeader ws, 2, 1, 14
End Sub

' ---------------------------------------------------------------------------
' WriteVarianceRows -- Compare maps and write diff rows
' ---------------------------------------------------------------------------
Private Sub WriteVarianceRows(ws As Worksheet, blMap As Object, actMap As Object)
    Dim outRow As Long
    outRow = 3

    ' 1. Activities in baseline
    Dim blKey As Variant
    For Each blKey In blMap.Keys
        Dim code As String
        code = CStr(blKey)

        Dim blData As Variant
        blData = blMap(code)

        If actMap.Exists(code) Then
            Dim actData As Variant
            actData = actMap(code)
            outRow = WriteCompareRow(ws, outRow, code, blData, actData)
        Else
            ' DELETED
            outRow = WriteDeletedRow(ws, outRow, code, blData)
        End If
    Next blKey

    ' 2. New activities (in current but not in baseline)
    Dim actKey As Variant
    For Each actKey In actMap.Keys
        code = CStr(actKey)
        If Not blMap.Exists(code) Then
            Dim actD As Variant
            actD = actMap(code)
            outRow = WriteNewRow(ws, outRow, code, actD)
        End If
    Next actKey
End Sub

' ---------------------------------------------------------------------------
' WriteCompareRow -- Write one compared activity row; returns next row number
' ---------------------------------------------------------------------------
Private Function WriteCompareRow(ws As Worksheet, r As Long, _
                                  code As String, blData As Variant, actData As Variant) As Long
    ' blData / actData: Array(name, start, finish, dur, pct, pred)
    Dim blStart  As Date:   blStart  = CDate(IIf(blData(1) = 0 Or IsEmpty(blData(1)), 0, blData(1)))
    Dim blFinish As Date:   blFinish = CDate(IIf(blData(2) = 0 Or IsEmpty(blData(2)), 0, blData(2)))
    Dim blDur    As Double: blDur    = CDbl(IIf(IsNumeric(blData(3)), blData(3), 0))
    Dim blPred   As String: blPred   = CStr(blData(5))

    Dim actStart  As Date:   actStart  = CDate(IIf(actData(1) = 0 Or IsEmpty(actData(1)), 0, actData(1)))
    Dim actFinish As Date:   actFinish = CDate(IIf(actData(2) = 0 Or IsEmpty(actData(2)), 0, actData(2)))
    Dim actDur    As Double: actDur    = CDbl(IIf(IsNumeric(actData(3)), actData(3), 0))
    Dim actPct    As Double: actPct    = CDbl(IIf(IsNumeric(actData(4)), actData(4), 0))
    Dim actPred   As String: actPred   = CStr(actData(5))

    Dim startSlip  As Long
    Dim finishSlip As Long
    Dim durDelta   As Double
    Dim notes      As String

    If blStart  <> 0 And actStart  <> 0 Then startSlip  = CLng(actStart  - blStart)
    If blFinish <> 0 And actFinish <> 0 Then finishSlip = CLng(actFinish - blFinish)
    durDelta = actDur - blDur

    ' Determine status
    Dim statusFlag As String
    If blPred <> actPred And blPred <> "" And actPred <> "" Then
        statusFlag = "LOGIC_CHANGED"
        notes = "Pred: [" & blPred & "] -> [" & actPred & "]"
    ElseIf finishSlip > 0 Then
        statusFlag = "SLIPPED"
    ElseIf finishSlip < 0 Then
        statusFlag = "AHEAD"
    Else
        statusFlag = "ON_TRACK"
    End If

    ' Write cells
    ws.Cells(r, 1).Value  = code
    ws.Cells(r, 2).Value  = CStr(actData(0))
    If blStart  <> 0 Then ws.Cells(r, 3).Value = blStart
    If blFinish <> 0 Then ws.Cells(r, 4).Value = blFinish
    ws.Cells(r, 5).Value  = blDur
    If actStart  <> 0 Then ws.Cells(r, 6).Value = actStart
    If actFinish <> 0 Then ws.Cells(r, 7).Value = actFinish
    ws.Cells(r, 8).Value  = actDur
    ws.Cells(r, 9).Value  = startSlip
    ws.Cells(r, 10).Value = finishSlip
    ws.Cells(r, 11).Value = durDelta
    ws.Cells(r, 12).Value = statusFlag
    ws.Cells(r, 13).Value = actPct / 100
    ws.Cells(r, 14).Value = notes

    ' Format date cells
    ws.Cells(r, 3).NumberFormat = "dd-mmm-yy"
    ws.Cells(r, 4).NumberFormat = "dd-mmm-yy"
    ws.Cells(r, 6).NumberFormat = "dd-mmm-yy"
    ws.Cells(r, 7).NumberFormat = "dd-mmm-yy"
    ws.Cells(r, 13).NumberFormat = "0%"

    ' Colour the status cell
    Dim statusColor As Long
    Select Case statusFlag
        Case "ON_TRACK":      statusColor = COLOR_ON_TRACK
        Case "SLIPPED":       statusColor = COLOR_SLIPPED
        Case "AHEAD":         statusColor = COLOR_AHEAD
        Case "LOGIC_CHANGED": statusColor = COLOR_LOGIC
        Case Else:            statusColor = RGB(240, 240, 240)
    End Select
    ws.Cells(r, 12).Interior.Color = statusColor
    If statusFlag <> "ON_TRACK" Then
        ws.Cells(r, 12).Font.Bold = True
        ws.Cells(r, 12).Font.Color = IIf(statusFlag = "SLIPPED", RGB(255, 255, 255), RGB(0, 0, 0))
    End If

    ' Highlight slip days
    If finishSlip > 5 Then
        ws.Cells(r, 10).Font.Color = RGB(200, 0, 0)
        ws.Cells(r, 10).Font.Bold = True
    ElseIf finishSlip > 0 Then
        ws.Cells(r, 10).Font.Color = RGB(180, 100, 0)
    ElseIf finishSlip < 0 Then
        ws.Cells(r, 10).Font.Color = RGB(0, 130, 60)
    End If

    WriteCompareRow = r + 1
End Function

Private Function WriteDeletedRow(ws As Worksheet, r As Long, code As String, blData As Variant) As Long
    ws.Cells(r, 1).Value = code
    ws.Cells(r, 2).Value = CStr(blData(0))
    If Not IsEmpty(blData(1)) And blData(1) <> 0 Then
        ws.Cells(r, 3).Value = blData(1)
        ws.Cells(r, 3).NumberFormat = "dd-mmm-yy"
    End If
    If Not IsEmpty(blData(2)) And blData(2) <> 0 Then
        ws.Cells(r, 4).Value = blData(2)
        ws.Cells(r, 4).NumberFormat = "dd-mmm-yy"
    End If
    ws.Cells(r, 5).Value  = blData(3)
    ws.Cells(r, 12).Value = "DELETED"
    ws.Cells(r, 12).Interior.Color = COLOR_DELETED
    ws.Cells(r, 12).Font.Color = RGB(80, 80, 80)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Interior.Color = RGB(230, 230, 230)
    WriteDeletedRow = r + 1
End Function

Private Function WriteNewRow(ws As Worksheet, r As Long, code As String, actData As Variant) As Long
    ws.Cells(r, 1).Value = code
    ws.Cells(r, 2).Value = CStr(actData(0))
    If Not IsEmpty(actData(1)) And actData(1) <> 0 Then
        ws.Cells(r, 6).Value = actData(1)
        ws.Cells(r, 6).NumberFormat = "dd-mmm-yy"
    End If
    If Not IsEmpty(actData(2)) And actData(2) <> 0 Then
        ws.Cells(r, 7).Value = actData(2)
        ws.Cells(r, 7).NumberFormat = "dd-mmm-yy"
    End If
    ws.Cells(r, 8).Value  = actData(3)
    ws.Cells(r, 13).Value = CDbl(IIf(IsNumeric(actData(4)), actData(4), 0)) / 100
    ws.Cells(r, 13).NumberFormat = "0%"
    ws.Cells(r, 12).Value = "NEW"
    ws.Cells(r, 12).Interior.Color = COLOR_NEW
    ws.Cells(r, 12).Font.Bold = True
    WriteNewRow = r + 1
End Function

' ---------------------------------------------------------------------------
' ClearBaseline -- Delete a named baseline sheet and its CDP entry
' ---------------------------------------------------------------------------
Public Sub ClearBaseline()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim names As Variant
    names = ListBaselineNames(wb)

    If IsEmpty(names) Then
        MsgBox "No baselines found.", vbInformation, "Protocol DEE"
        Exit Sub
    End If

    Dim listStr As String
    Dim i As Integer
    For i = 0 To UBound(names)
        listStr = listStr & (i + 1) & ") " & names(i) & vbCrLf
    Next i

    Dim choice As String
    choice = InputBox("Select baseline to clear:" & vbCrLf & vbCrLf & listStr, _
                      "Clear Baseline", "1")
    If choice = "" Then Exit Sub

    Dim choiceNum As Integer
    On Error Resume Next
    choiceNum = CInt(choice) - 1
    On Error GoTo 0
    If choiceNum < 0 Or choiceNum > UBound(names) Then Exit Sub

    Dim name As String
    name = CStr(names(choiceNum))

    DeleteSheetIfExists wb, SH_PREFIX & name
    RemoveBaselineCDP wb, name
    DEE_Logger.LogInfo "DEE_BaselineCompare", "Baseline cleared: " & name
    MsgBox "Baseline '" & name & "' cleared.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' Helpers: CDP storage, sheet deletion, name listing
' ---------------------------------------------------------------------------
Private Sub StoreBaselineCDP(wb As Workbook, name As String, filePath As String)
    Dim cdpName As String
    cdpName = CDP_PREFIX & name
    On Error Resume Next
    wb.CustomDocumentProperties(cdpName).Delete
    On Error GoTo 0
    On Error Resume Next
    wb.CustomDocumentProperties.Add cdpName, False, msoPropertyTypeString, filePath
    On Error GoTo 0
End Sub

Private Sub RemoveBaselineCDP(wb As Workbook, name As String)
    On Error Resume Next
    wb.CustomDocumentProperties(CDP_PREFIX & name).Delete
    On Error GoTo 0
End Sub

Private Sub DeleteSheetIfExists(wb As Workbook, sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

Private Function ListBaselineNames(wb As Workbook) As Variant
    Dim names() As String
    Dim count As Integer
    count = 0

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Left(ws.Name, Len(SH_PREFIX)) = SH_PREFIX Then
            ReDim Preserve names(count)
            names(count) = Mid(ws.Name, Len(SH_PREFIX) + 1)
            count = count + 1
        End If
    Next ws

    If count = 0 Then
        ListBaselineNames = Empty
    Else
        ListBaselineNames = names
    End If
End Function
