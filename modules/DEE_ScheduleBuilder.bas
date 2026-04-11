Attribute VB_Name = "DEE_ScheduleBuilder"
Option Explicit

' =============================================================================
' DEE_ScheduleBuilder.bas -- Schedule Compiler
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Compiles raw PROJWBS + TASK XER tables into the hierarchical Schedule sheet.
'
' SCHEDULE SHEET LAYOUT CONTRACT (Columns A-F):
'   A = Activity ID (WBS code or task_code)
'   B = Activity Name (EMPTY for WBS rows -- canonical WBS detection rule)
'   C = Start  (Excel Date)
'   D = Finish (Excel Date)
'   E = Budget (user-editable)
'   F = Budget (read-only locked copy; PMS formulas reference F)
'
' WBS rows: B = empty, C-D = empty, E-F = 0
' Activity rows: B = indented task name, C-D = dates, E-F = budget
' =============================================================================

' ---------------------------------------------------------------------------
' BuildSchedule -- Main entry point
' budgetField: XER field name to use for budget (e.g. "budget_qty")
' filterProjId: if non-empty, only include activities for this project
' ---------------------------------------------------------------------------
Public Sub BuildSchedule(Optional budgetField As String = "budget_qty", _
                          Optional filterProjId As String = "")
    If Not DEE_XERParser.IsXERParsed() Then
        MsgBox "No XER file loaded. Please load a schedule first.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    DEE_Utils.StartProgress "Building Schedule"

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Initialize warnings sheet
    Dim wsWarn As Worksheet
    Set wsWarn = DEE_Utils.GetOrCreateSheet(wb, "Warnings")
    wsWarn.Cells.Clear

    ' Get or create Schedule sheet
    Dim ws As Worksheet
    Set ws = DEE_Utils.GetOrCreateSheet(wb, "Schedule")
    ws.Cells.Clear

    ' -----------------------------------------------------------------------
    ' 1. Read PROJWBS fields
    ' -----------------------------------------------------------------------
    Dim wbsFields As Variant
    Dim wbsRows As Collection
    Dim wbsHasData As Boolean
    wbsHasData = DEE_XERParser.TableExists("PROJWBS")

    Dim wbsIdIdx As Integer, wbsParentIdx As Integer
    Dim wbsShortIdx As Integer, wbsNameIdx As Integer
    Dim wbsProjIdx As Integer

    If wbsHasData Then
        wbsFields = DEE_XERParser.GetFieldNames("PROJWBS")
        Set wbsRows = DEE_XERParser.GetTable("PROJWBS")
        wbsIdIdx = DEE_XERParser.FindFieldIndex(wbsFields, "wbs_id")
        wbsParentIdx = DEE_XERParser.FindFieldIndex(wbsFields, "parent_wbs_id")
        wbsShortIdx = DEE_XERParser.FindFieldIndex(wbsFields, "wbs_short_name")
        wbsNameIdx = DEE_XERParser.FindFieldIndex(wbsFields, "wbs_name")
        wbsProjIdx = DEE_XERParser.FindFieldIndex(wbsFields, "proj_id")
    End If

    ' -----------------------------------------------------------------------
    ' 2. Read TASK fields
    ' -----------------------------------------------------------------------
    If Not DEE_XERParser.TableExists("TASK") Then
        MsgBox "XER has no activities (TASK table missing).", vbCritical, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    Dim taskFields As Variant
    taskFields = DEE_XERParser.GetFieldNames("TASK")
    Dim taskRows As Collection
    Set taskRows = DEE_XERParser.GetTable("TASK")

    Dim taskIdIdx As Integer, taskWbsIdx As Integer, taskCodeIdx As Integer
    Dim taskNameIdx As Integer, taskStartIdx As Integer, taskEndIdx As Integer
    Dim taskBudgetIdx As Integer, taskProjIdx As Integer, taskTypeIdx As Integer
    Dim taskPctIdx As Integer

    taskIdIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_id")
    taskWbsIdx = DEE_XERParser.FindFieldIndex(taskFields, "wbs_id")
    taskCodeIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_code")
    taskNameIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_name")
    taskStartIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_start_date")
    taskEndIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_end_date")
    taskBudgetIdx = DEE_XERParser.FindFieldIndex(taskFields, budgetField)
    taskProjIdx = DEE_XERParser.FindFieldIndex(taskFields, "proj_id")
    taskTypeIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_type")
    taskPctIdx = DEE_XERParser.FindFieldIndex(taskFields, "phys_complete_pct")

    ' -----------------------------------------------------------------------
    ' 3. Build WBS maps (only if PROJWBS exists)
    ' -----------------------------------------------------------------------
    Dim wbsMap As Object       ' wbs_id -> Variant row array
    Dim parentMap As Object    ' parent_wbs_id -> Collection of child wbs_ids
    Set wbsMap = CreateObject("Scripting.Dictionary")
    Set parentMap = CreateObject("Scripting.Dictionary")

    Dim warnCount As Long
    Dim errCount As Long
    warnCount = 0
    errCount = 0

    If wbsHasData Then
        Dim wbsRow As Variant
        For Each wbsRow In wbsRows
            Dim rd As Variant
            rd = wbsRow
            ' Filter by project if requested
            If filterProjId <> "" And wbsProjIdx >= 0 Then
                If wbsProjIdx <= UBound(rd) Then
                    If rd(wbsProjIdx) <> filterProjId Then GoTo NextWBS
                End If
            End If

            Dim thisWbsId As String
            If wbsIdIdx >= 0 And wbsIdIdx <= UBound(rd) Then
                thisWbsId = Trim(rd(wbsIdIdx))
            Else
                GoTo NextWBS
            End If

            If thisWbsId = "" Then GoTo NextWBS

            If Not wbsMap.Exists(thisWbsId) Then
                wbsMap.Add thisWbsId, rd
            End If

            Dim thisParentId As String
            If wbsParentIdx >= 0 And wbsParentIdx <= UBound(rd) Then
                thisParentId = Trim(rd(wbsParentIdx))
            End If

            If thisParentId <> "" Then
                If Not parentMap.Exists(thisParentId) Then
                    parentMap.Add thisParentId, New Collection
                End If
                parentMap(thisParentId).Add thisWbsId
            End If
NextWBS:
        Next wbsRow
    End If

    ' -----------------------------------------------------------------------
    ' 4. Build task map: wbs_id -> Collection of task row arrays
    ' -----------------------------------------------------------------------
    Dim taskMap As Object
    Set taskMap = CreateObject("Scripting.Dictionary")

    Dim taskCodeSet As Object   ' For duplicate detection
    Set taskCodeSet = CreateObject("Scripting.Dictionary")

    Dim taskRow As Variant
    For Each taskRow In taskRows
        Dim tr As Variant
        tr = taskRow

        ' Filter by project if requested
        If filterProjId <> "" And taskProjIdx >= 0 Then
            If taskProjIdx <= UBound(tr) Then
                If tr(taskProjIdx) <> filterProjId Then GoTo NextTask
            End If
        End If

        ' Skip WBS-type tasks (TT_WBS)
        If taskTypeIdx >= 0 And taskTypeIdx <= UBound(tr) Then
            If Trim(tr(taskTypeIdx)) = "TT_WBS" Then GoTo NextTask
        End If

        Dim thisTaskWbs As String
        If taskWbsIdx >= 0 And taskWbsIdx <= UBound(tr) Then
            thisTaskWbs = Trim(tr(taskWbsIdx))
        End If

        ' Duplicate task_code check
        Dim tCode As String
        If taskCodeIdx >= 0 And taskCodeIdx <= UBound(tr) Then
            tCode = Trim(tr(taskCodeIdx))
        End If
        If tCode <> "" Then
            If taskCodeSet.Exists(tCode) Then
                Dim dupCount As Integer
                dupCount = taskCodeSet(tCode) + 1
                taskCodeSet(tCode) = dupCount
                ' Modify task code to be unique
                tr(taskCodeIdx) = tCode & "_DUP" & dupCount
                DEE_Utils.LogWarning wb, "WARNING", _
                    "Duplicate task_code '" & tCode & "' renamed to '" & tr(taskCodeIdx) & "'", _
                    "TASK table"
                warnCount = warnCount + 1
            Else
                taskCodeSet.Add tCode, 1
            End If
        End If

        ' Add to taskMap
        If Not taskMap.Exists(thisTaskWbs) Then
            taskMap.Add thisTaskWbs, New Collection
        End If
        taskMap(thisTaskWbs).Add tr
NextTask:
    Next taskRow

    ' -----------------------------------------------------------------------
    ' 5. Write header row
    ' -----------------------------------------------------------------------
    ws.Cells(1, 1).Value = "Activity ID"
    ws.Cells(1, 2).Value = "Activity Name"
    ws.Cells(1, 3).Value = "Start"
    ws.Cells(1, 4).Value = "Finish"
    ws.Cells(1, 5).Value = "Budget"
    ws.Cells(1, 6).Value = "Budget"
    DEE_Utils.FormatHeaderRow ws, RGB(0, 32, 96), RGB(255, 255, 255), 1

    ' -----------------------------------------------------------------------
    ' 6. Recursive traversal into output array
    ' -----------------------------------------------------------------------
    Dim outputRows() As Variant
    Dim rowCount As Long
    ' Pre-allocate large array (trim later)
    Dim maxExpected As Long
    maxExpected = taskRows.Count * 3 + 100
    ReDim outputRows(0 To maxExpected, 0 To 5)
    rowCount = 0

    If wbsHasData And wbsMap.Count > 0 Then
        ' Find root WBS nodes: those whose parent_wbs_id is NOT in wbsMap
        Dim rootNodes() As String
        Dim rootCount As Integer
        ReDim rootNodes(0 To wbsMap.Count - 1)
        rootCount = 0

        Dim wbsKey As Variant
        For Each wbsKey In wbsMap.Keys
            Dim nodeData As Variant
            nodeData = wbsMap(wbsKey)
            Dim nodeParent As String
            If wbsParentIdx >= 0 And wbsParentIdx <= UBound(nodeData) Then
                nodeParent = Trim(nodeData(wbsParentIdx))
            End If
            ' Root: parent not in wbsMap
            If nodeParent = "" Or Not wbsMap.Exists(nodeParent) Then
                rootNodes(rootCount) = CStr(wbsKey)
                rootCount = rootCount + 1
            End If
        Next wbsKey

        ' Sort root nodes by seq_num if available (use order they appear)
        ' Traverse each root node
        Dim rn As Integer
        For rn = 0 To rootCount - 1
            TraverseWBS rootNodes(rn), 0, wbsMap, parentMap, taskMap, _
                        wbsShortIdx, wbsNameIdx, taskCodeIdx, taskNameIdx, _
                        taskStartIdx, taskEndIdx, taskBudgetIdx, _
                        outputRows, rowCount, warnCount, wb
        Next rn
    Else
        ' Flat mode: no PROJWBS -- write all tasks directly
        For Each taskRow In taskRows
            Dim flatTr As Variant
            flatTr = taskRow
            AppendTaskRow flatTr, 0, taskCodeIdx, taskNameIdx, _
                          taskStartIdx, taskEndIdx, taskBudgetIdx, _
                          outputRows, rowCount, warnCount, wb
        Next taskRow
    End If

    ' -----------------------------------------------------------------------
    ' 7. Bulk write to Schedule sheet
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 80, "Writing Schedule sheet"

    If rowCount > 0 Then
        ' Trim array
        ReDim Preserve outputRows(0 To rowCount - 1, 0 To 5)
        ws.Range(ws.Cells(2, 1), ws.Cells(rowCount + 1, 6)).Value = outputRows

        ' Format date columns
        ws.Columns(3).NumberFormat = "yyyy-mm-dd"
        ws.Columns(4).NumberFormat = "yyyy-mm-dd"

        ' Format budget columns
        ws.Columns(5).NumberFormat = """$""#,##0.00"
        ws.Columns(6).NumberFormat = """$""#,##0.00"
    End If

    ' -----------------------------------------------------------------------
    ' 8. Auto-apply formatting
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 90, "Applying formatting"

    ' Freeze pane
    DEE_Utils.FreezePaneRow1 ws

    ' Auto-fit
    ws.Columns("A:F").AutoFit

    ' Apply WBS colors
    DEE_WBS.ApplyWBSColoring ws

    ' Apply row grouping
    DEE_WBS.ApplyWBSGrouping ws, False

    ' -----------------------------------------------------------------------
    ' 9. Show completion summary
    ' -----------------------------------------------------------------------
    DEE_Utils.EndProgress

    Dim msg As String
    msg = "Schedule built successfully!" & vbCrLf
    msg = msg & rowCount & " rows compiled." & vbCrLf

    If warnCount > 0 Or errCount > 0 Then
        msg = msg & vbCrLf
        If errCount > 0 Then msg = msg & "[!] " & errCount & " errors logged to Warnings sheet." & vbCrLf
        If warnCount > 0 Then msg = msg & "[!] " & warnCount & " warnings logged to Warnings sheet." & vbCrLf
    End If

    msg = msg & vbCrLf & "You can now use: Apply Colors, Draw Gantt, Create PMS"

    MsgBox msg, vbInformation, "Protocol DEE -- Schedule Built"
End Sub

' ---------------------------------------------------------------------------
' TraverseWBS -- Recursive WBS traversal
' ---------------------------------------------------------------------------
Private Sub TraverseWBS(wbsId As String, level As Integer, _
                         wbsMap As Object, parentMap As Object, taskMap As Object, _
                         wbsShortIdx As Integer, wbsNameIdx As Integer, _
                         taskCodeIdx As Integer, taskNameIdx As Integer, _
                         taskStartIdx As Integer, taskEndIdx As Integer, _
                         taskBudgetIdx As Integer, _
                         outputRows() As Variant, ByRef rowCount As Long, _
                         ByRef warnCount As Long, wb As Workbook)

    If Not wbsMap.Exists(wbsId) Then Exit Sub

    Dim nodeData As Variant
    nodeData = wbsMap(wbsId)

    ' Build WBS row
    Dim indent As String
    indent = Space(level * 2)

    Dim wbsCode As String
    Dim wbsName As String
    If wbsShortIdx >= 0 And wbsShortIdx <= UBound(nodeData) Then wbsCode = Trim(nodeData(wbsShortIdx))
    If wbsNameIdx >= 0 And wbsNameIdx <= UBound(nodeData) Then wbsName = Trim(nodeData(wbsNameIdx))

    ' WBS row: B = empty (canonical detection)
    If rowCount > UBound(outputRows, 1) Then
        ReDim Preserve outputRows(0 To rowCount * 2, 0 To 5)
    End If

    outputRows(rowCount, 0) = indent & wbsCode  ' A: WBS code (indented)
    outputRows(rowCount, 1) = ""                 ' B: EMPTY (WBS marker)
    outputRows(rowCount, 2) = ""                 ' C: no date
    outputRows(rowCount, 3) = ""                 ' D: no date
    outputRows(rowCount, 4) = 0                  ' E: budget = 0 (PMS will SUBTOTAL)
    outputRows(rowCount, 5) = 0                  ' F: budget locked copy
    rowCount = rowCount + 1

    ' Add tasks for this WBS node
    If taskMap.Exists(wbsId) Then
        Dim taskColl As Collection
        Set taskColl = taskMap(wbsId)
        Dim tr As Variant
        For Each tr In taskColl
            AppendTaskRow tr, level + 1, taskCodeIdx, taskNameIdx, _
                          taskStartIdx, taskEndIdx, taskBudgetIdx, _
                          outputRows, rowCount, warnCount, wb
        Next tr
    End If

    ' Recurse into child WBS nodes
    If parentMap.Exists(wbsId) Then
        Dim childColl As Collection
        Set childColl = parentMap(wbsId)
        Dim childId As Variant
        For Each childId In childColl
            TraverseWBS CStr(childId), level + 1, wbsMap, parentMap, taskMap, _
                        wbsShortIdx, wbsNameIdx, taskCodeIdx, taskNameIdx, _
                        taskStartIdx, taskEndIdx, taskBudgetIdx, _
                        outputRows, rowCount, warnCount, wb
        Next childId
    End If
End Sub

' ---------------------------------------------------------------------------
' AppendTaskRow -- Add a single activity row to outputRows
' ---------------------------------------------------------------------------
Private Sub AppendTaskRow(tr As Variant, level As Integer, _
                           taskCodeIdx As Integer, taskNameIdx As Integer, _
                           taskStartIdx As Integer, taskEndIdx As Integer, _
                           taskBudgetIdx As Integer, _
                           outputRows() As Variant, ByRef rowCount As Long, _
                           ByRef warnCount As Long, wb As Workbook)

    If rowCount > UBound(outputRows, 1) Then
        ReDim Preserve outputRows(0 To rowCount * 2, 0 To 5)
    End If

    Dim indent As String
    indent = Space(level * 2)

    ' Activity ID (task_code)
    Dim tCode As String
    If taskCodeIdx >= 0 And taskCodeIdx <= UBound(tr) Then tCode = Trim(tr(taskCodeIdx))

    ' Activity Name
    Dim tName As String
    If taskNameIdx >= 0 And taskNameIdx <= UBound(tr) Then tName = Trim(tr(taskNameIdx))

    ' Dates (locale-safe)
    Dim startDate As Variant
    Dim endDate As Variant
    If taskStartIdx >= 0 And taskStartIdx <= UBound(tr) Then
        startDate = DEE_Utils.SafeParseDate(CStr(tr(taskStartIdx)))
        If IsEmpty(startDate) And Len(Trim(tr(taskStartIdx))) > 0 Then
            DEE_Utils.LogWarning wb, "WARNING", _
                "Date parse failure for task '" & tCode & "' start", _
                "TASK row", CStr(tr(taskStartIdx))
            warnCount = warnCount + 1
        End If
    End If
    If taskEndIdx >= 0 And taskEndIdx <= UBound(tr) Then
        endDate = DEE_Utils.SafeParseDate(CStr(tr(taskEndIdx)))
        If IsEmpty(endDate) And Len(Trim(tr(taskEndIdx))) > 0 Then
            DEE_Utils.LogWarning wb, "WARNING", _
                "Date parse failure for task '" & tCode & "' finish", _
                "TASK row", CStr(tr(taskEndIdx))
            warnCount = warnCount + 1
        End If
    End If

    ' Budget
    Dim budget As Double
    budget = 0
    If taskBudgetIdx >= 0 And taskBudgetIdx <= UBound(tr) Then
        On Error Resume Next
        budget = CDbl(tr(taskBudgetIdx))
        If Err.Number <> 0 Then budget = 0
        On Error GoTo 0
    End If

    ' Write row: A=code, B=name (NOT empty -- activity signal), C=start, D=finish, E-F=budget
    outputRows(rowCount, 0) = indent & tCode    ' A: task_code (indented)
    outputRows(rowCount, 1) = indent & tName    ' B: task_name (non-empty = activity)
    outputRows(rowCount, 2) = startDate         ' C: Start (Excel Date)
    outputRows(rowCount, 3) = endDate           ' D: Finish (Excel Date)
    outputRows(rowCount, 4) = budget            ' E: Budget
    outputRows(rowCount, 5) = budget            ' F: Budget (locked copy)
    rowCount = rowCount + 1
End Sub
