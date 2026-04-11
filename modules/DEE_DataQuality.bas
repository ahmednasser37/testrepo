Attribute VB_Name = "DEE_DataQuality"
Option Explicit

' =============================================================================
' DEE_DataQuality.bas -- Data Quality Checks (Pre-Build Validation)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Runs validation checks on parsed XER data before compiling Schedule sheet.
' Output: "Warnings" sheet with severity-coded rows.
'
' CHECKS (Section 15.6):
'   Duplicate task_code         -> Error   (auto-fix: append _DUP1)
'   Finish < Start              -> Error   (no auto-fix)
'   Duration > 365 working days -> Warning (no auto-fix)
'   Activity with no dates      -> Warning (no auto-fix)
'   Budget = 0 or null          -> Info    (default to 0)
'   Missing predecessors (dang) -> Warning (no auto-fix)
'   Circular dependency         -> Error   (no auto-fix)
' =============================================================================

' ---------------------------------------------------------------------------
' RunDataQualityChecks -- Main entry point
' Returns True if no errors found (warnings OK)
' ---------------------------------------------------------------------------
Public Function RunDataQualityChecks() As Boolean
    RunDataQualityChecks = True

    If Not DEE_XERParser.IsXERParsed() Then
        MsgBox "No XER file loaded.", vbExclamation, "Protocol DEE"
        RunDataQualityChecks = False
        Exit Function
    End If

    DEE_Utils.StartProgress "Data Quality Checks"

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Clear warnings sheet
    Dim wsWarn As Worksheet
    Set wsWarn = DEE_Utils.GetOrCreateSheet(wb, "Warnings")
    wsWarn.Cells.Clear

    Dim errorCount As Long
    Dim warnCount As Long
    Dim infoCount As Long
    errorCount = 0
    warnCount = 0
    infoCount = 0

    ' -----------------------------------------------------------------------
    ' 1. TASK table checks
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 10, "Checking TASK table"

    If Not DEE_XERParser.TableExists("TASK") Then
        DEE_Utils.LogWarning wb, "ERROR", "TASK table not found in XER file"
        errorCount = errorCount + 1
        GoTo SummaryOutput
    End If

    Dim taskFields As Variant
    taskFields = DEE_XERParser.GetFieldNames("TASK")
    Dim taskRows As Collection
    Set taskRows = DEE_XERParser.GetTable("TASK")

    Dim taskCodeIdx As Integer
    Dim taskStartIdx As Integer
    Dim taskEndIdx As Integer
    Dim taskBudgetIdx As Integer
    Dim taskDurIdx As Integer
    Dim taskTypeIdx As Integer

    taskCodeIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_code")
    taskStartIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_start_date")
    taskEndIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_end_date")
    taskBudgetIdx = DEE_XERParser.FindFieldIndex(taskFields, "budget_qty")
    taskDurIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_drtn_hr_cnt")
    taskTypeIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_type")

    ' Build task_code uniqueness set
    Dim codeSet As Object
    Set codeSet = CreateObject("Scripting.Dictionary")

    Dim taskIdSet As Object
    Set taskIdSet = CreateObject("Scripting.Dictionary")

    Dim taskIdIdx As Integer
    taskIdIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_id")

    Dim rowNum As Long
    rowNum = 0
    Dim taskRow As Variant
    For Each taskRow In taskRows
        rowNum = rowNum + 1
        Dim tr As Variant
        tr = taskRow

        ' Skip WBS summary tasks
        Dim taskType As String
        If taskTypeIdx >= 0 And taskTypeIdx <= UBound(tr) Then
            taskType = Trim(tr(taskTypeIdx))
        End If
        If taskType = "TT_WBS" Then GoTo NextDQTask

        ' Track task IDs
        Dim tId As String
        If taskIdIdx >= 0 And taskIdIdx <= UBound(tr) Then tId = Trim(tr(taskIdIdx))
        If tId <> "" Then taskIdSet(tId) = True

        ' Check 1: Duplicate task_code
        Dim tCode As String
        If taskCodeIdx >= 0 And taskCodeIdx <= UBound(tr) Then tCode = Trim(tr(taskCodeIdx))
        If tCode <> "" Then
            If codeSet.Exists(tCode) Then
                DEE_Utils.LogWarning wb, "ERROR", _
                    "Duplicate task_code: '" & tCode & "' (will be renamed _DUP" & codeSet(tCode) + 1 & ")", _
                    "TASK row " & rowNum, tCode
                codeSet(tCode) = codeSet(tCode) + 1
                errorCount = errorCount + 1
            Else
                codeSet.Add tCode, 1
            End If
        End If

        ' Check 2: Finish < Start
        Dim tStart As Variant
        Dim tEnd As Variant
        If taskStartIdx >= 0 And taskStartIdx <= UBound(tr) Then
            tStart = DEE_Utils.SafeParseDate(CStr(tr(taskStartIdx)))
        End If
        If taskEndIdx >= 0 And taskEndIdx <= UBound(tr) Then
            tEnd = DEE_Utils.SafeParseDate(CStr(tr(taskEndIdx)))
        End If

        If Not IsEmpty(tStart) And Not IsEmpty(tEnd) Then
            If CDate(tEnd) < CDate(tStart) Then
                DEE_Utils.LogWarning wb, "ERROR", _
                    "Finish before Start for activity '" & tCode & "'", _
                    "TASK row " & rowNum, _
                    CStr(tr(taskStartIdx)) & " -> " & CStr(tr(taskEndIdx))
                errorCount = errorCount + 1
            End If

            ' Check 3: Duration > 365 working days
            Dim durDays As Double
            If taskDurIdx >= 0 And taskDurIdx <= UBound(tr) Then
                On Error Resume Next
                durDays = CDbl(tr(taskDurIdx)) / 8
                On Error GoTo 0
                If durDays > 365 Then
                    DEE_Utils.LogWarning wb, "WARNING", _
                        "Unusually long duration: " & Format(durDays, "0") & " days for '" & tCode & "'", _
                        "TASK row " & rowNum
                    warnCount = warnCount + 1
                End If
            End If
        End If

        ' Check 4: Activity with no dates
        If IsEmpty(tStart) And IsEmpty(tEnd) Then
            Dim rawStart As String
            Dim rawEnd As String
            If taskStartIdx >= 0 And taskStartIdx <= UBound(tr) Then rawStart = Trim(CStr(tr(taskStartIdx)))
            If taskEndIdx >= 0 And taskEndIdx <= UBound(tr) Then rawEnd = Trim(CStr(tr(taskEndIdx)))
            If rawStart = "" And rawEnd = "" Then
                DEE_Utils.LogWarning wb, "WARNING", _
                    "Activity has no dates: '" & tCode & "'", _
                    "TASK row " & rowNum
                warnCount = warnCount + 1
            End If
        End If

        ' Check 5: Budget = 0 or null
        Dim budget As Double
        budget = 0
        If taskBudgetIdx >= 0 And taskBudgetIdx <= UBound(tr) Then
            On Error Resume Next
            budget = CDbl(tr(taskBudgetIdx))
            If Err.Number <> 0 Then budget = 0
            On Error GoTo 0
        End If
        If budget = 0 Then
            DEE_Utils.LogWarning wb, "INFO", _
                "Zero/null budget for activity '" & tCode & "'", _
                "TASK row " & rowNum
            infoCount = infoCount + 1
        End If

NextDQTask:
    Next taskRow

    ' -----------------------------------------------------------------------
    ' 2. TASKPRED checks (predecessor/successor completeness + circular)
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 40, "Checking predecessor relationships"

    If DEE_XERParser.TableExists("TASKPRED") Then
        Dim predFields As Variant
        predFields = DEE_XERParser.GetFieldNames("TASKPRED")
        Dim predRows As Collection
        Set predRows = DEE_XERParser.GetTable("TASKPRED")

        Dim predTaskIdIdx As Integer
        Dim succTaskIdIdx As Integer
        succTaskIdIdx = DEE_XERParser.FindFieldIndex(predFields, "task_id")
        predTaskIdIdx = DEE_XERParser.FindFieldIndex(predFields, "pred_task_id")

        Dim tasksWithPred As Object
        Set tasksWithPred = CreateObject("Scripting.Dictionary")
        Dim tasksWithSucc As Object
        Set tasksWithSucc = CreateObject("Scripting.Dictionary")
        Dim adjList As Object  ' For circular dependency check
        Set adjList = CreateObject("Scripting.Dictionary")

        For Each k In taskIdSet.Keys
            adjList(CStr(k)) = New Collection
        Next k

        Dim predRow As Variant
        For Each predRow In predRows
            Dim pr As Variant
            pr = predRow

            Dim succId As String
            Dim predId As String
            If succTaskIdIdx >= 0 And succTaskIdIdx <= UBound(pr) Then succId = Trim(pr(succTaskIdIdx))
            If predTaskIdIdx >= 0 And predTaskIdIdx <= UBound(pr) Then predId = Trim(pr(predTaskIdIdx))

            If taskIdSet.Exists(succId) And taskIdSet.Exists(predId) Then
                tasksWithPred(succId) = True
                tasksWithSucc(predId) = True
                If adjList.Exists(predId) Then
                    adjList(predId).Add succId
                End If
            End If
        Next predRow

        ' Check 6: Missing predecessors (dangling start)
        For Each k In taskIdSet.Keys
            If Not tasksWithPred.Exists(CStr(k)) Then
                DEE_Utils.LogWarning wb, "WARNING", _
                    "Activity '" & CStr(k) & "' has no predecessors (dangling start)", _
                    "TASKPRED"
                warnCount = warnCount + 1
            End If
        Next k

        ' Missing successors
        For Each k In taskIdSet.Keys
            If Not tasksWithSucc.Exists(CStr(k)) Then
                DEE_Utils.LogWarning wb, "WARNING", _
                    "Activity '" & CStr(k) & "' has no successors (dangling end)", _
                    "TASKPRED"
                warnCount = warnCount + 1
            End If
        Next k

        ' Check 7: Circular dependency (simple cycle detection via DFS)
        DEE_Utils.UpdateProgress 70, "Checking for circular dependencies"
        Dim hasCycle As Boolean
        hasCycle = DetectCycles(adjList)
        If hasCycle Then
            DEE_Utils.LogWarning wb, "ERROR", _
                "Circular dependency detected in predecessor network. CPM analysis may fail.", _
                "TASKPRED"
            errorCount = errorCount + 1
        End If
    End If

    ' -----------------------------------------------------------------------
    ' Summary
    ' -----------------------------------------------------------------------
SummaryOutput:
    DEE_Utils.UpdateProgress 95, "Generating summary"

    ' Write summary at top of Warnings sheet
    wsWarn.Rows("1:1").Insert
    wsWarn.Rows("1:1").Insert
    wsWarn.Rows("1:1").Insert

    wsWarn.Cells(1, 1).Value = "Data Quality Summary"
    wsWarn.Cells(1, 1).Font.Bold = True
    wsWarn.Cells(1, 1).Font.Size = 14

    wsWarn.Cells(2, 1).Value = "Errors: " & errorCount
    wsWarn.Cells(2, 1).Font.Bold = True
    wsWarn.Cells(2, 1).Font.Color = IIf(errorCount > 0, RGB(255, 0, 0), RGB(0, 176, 80))

    wsWarn.Cells(2, 2).Value = "Warnings: " & warnCount
    wsWarn.Cells(2, 2).Font.Bold = True
    wsWarn.Cells(2, 2).Font.Color = IIf(warnCount > 0, RGB(200, 100, 0), RGB(0, 176, 80))

    wsWarn.Cells(2, 3).Value = "Info: " & infoCount
    wsWarn.Cells(2, 3).Font.Color = RGB(0, 112, 192)

    wsWarn.Columns("A:E").AutoFit

    DEE_Utils.EndProgress

    Dim msg As String
    msg = "Data Quality Check Complete:" & vbCrLf & vbCrLf
    If errorCount > 0 Then
        msg = msg & "[X] " & errorCount & " Errors found" & vbCrLf
    End If
    If warnCount > 0 Then
        msg = msg & "[!] " & warnCount & " Warnings found" & vbCrLf
    End If
    msg = msg & "[i] " & infoCount & " Info items" & vbCrLf & vbCrLf
    msg = msg & "See the Warnings sheet for details."

    MsgBox msg, IIf(errorCount > 0, vbExclamation, vbInformation), "Protocol DEE -- Data Quality"

    RunDataQualityChecks = (errorCount = 0)
End Function

' ---------------------------------------------------------------------------
' DetectCycles -- DFS-based cycle detection in predecessor graph
' Returns True if a cycle exists
' ---------------------------------------------------------------------------
Private Function DetectCycles(adjList As Object) As Boolean
    Dim visited As Object
    Dim recursionStack As Object
    Set visited = CreateObject("Scripting.Dictionary")
    Set recursionStack = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In adjList.Keys
        If Not visited.Exists(CStr(k)) Then
            If DFSCycleCheck(adjList, CStr(k), visited, recursionStack) Then
                DetectCycles = True
                Exit Function
            End If
        End If
    Next k

    DetectCycles = False
End Function

Private Function DFSCycleCheck(adjList As Object, nodeId As String, _
                                 visited As Object, recursionStack As Object) As Boolean
    visited(nodeId) = True
    recursionStack(nodeId) = True

    If adjList.Exists(nodeId) Then
        Dim neighbor As Variant
        For Each neighbor In adjList(nodeId)
            Dim nId As String
            nId = CStr(neighbor)
            If Not visited.Exists(nId) Then
                If DFSCycleCheck(adjList, nId, visited, recursionStack) Then
                    DFSCycleCheck = True
                    Exit Function
                End If
            ElseIf recursionStack.Exists(nId) Then
                DFSCycleCheck = True
                Exit Function
            End If
        Next neighbor
    End If

    recursionStack.Remove nodeId
    DFSCycleCheck = False
End Function
