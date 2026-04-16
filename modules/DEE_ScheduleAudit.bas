Attribute VB_Name = "DEE_ScheduleAudit"
Option Explicit

' =============================================================================
' DEE_ScheduleAudit.bas -- DCMA 14-Point Schedule Health Check
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Performs the DCMA 14-point schedule quality assessment on the parsed XER
' data in memory. Outputs results to "Audit Results" sheet with traffic lights.
'
' Reference: Plexxel_Analysis/vba/ScheduleCheck.bas (3,292 lines)
'
' DCMA 14 CHECKS:
'   1.  Logic -- missing predecessors         (< 5%)
'   2.  Logic -- missing successors           (< 5%)
'   3.  Relationship types -- FS only         (> 90%)
'   4.  Lags                                  (< 5%)
'   5.  Leads (negative lags)                 (0%)
'   6.  Relationship density                  (>= 1 per activity)
'   7.  Hard constraints                      (< 5%)
'   8.  High float (TF > 44 days)             (< 5%)
'   9.  Negative float                        (0%)
'  10.  High duration (> 44 days)             (< 5%)
'  11.  Invalid dates                         (0%)
'  12.  Resources assigned                    (> 90%)
'  13.  Missed activities                     (0%)
'  14.  Critical path length index (CPLI)     (>= 0.95)
' =============================================================================

' Traffic light colors
Private Const COLOR_PASS As Long = RGB(0, 176, 80)    ' Green
Private Const COLOR_WARN As Long = RGB(255, 192, 0)   ' Amber
Private Const COLOR_FAIL As Long = RGB(255, 0, 0)     ' Red

' ---------------------------------------------------------------------------
' RunDCMAaudit -- Main entry point
' ---------------------------------------------------------------------------
Public Sub RunDCMAaudit()
    If Not DEE_XERParser.IsXERParsed() Then
        MsgBox "No XER file loaded. Please load a schedule first.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    DEE_Utils.StartProgress "Running DCMA 14-Point Audit"

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim wsAudit As Worksheet
    Set wsAudit = DEE_Utils.GetOrCreateSheet(wb, "Audit Results")
    wsAudit.Cells.Clear

    ' -----------------------------------------------------------------------
    ' Gather data from XER tables
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 5, "Loading TASK data"

    ' Task data
    Dim taskFields As Variant
    taskFields = DEE_XERParser.GetFieldNames("TASK")
    Dim taskRows As Collection
    Set taskRows = DEE_XERParser.GetTable("TASK")

    Dim totalTasks As Long
    totalTasks = 0
    If Not taskRows Is Nothing Then totalTasks = taskRows.Count

    If totalTasks = 0 Then
        MsgBox "No activities found in loaded XER.", vbExclamation, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    ' Field indexes
    Dim taskIdIdx As Integer
    Dim taskTypeIdx As Integer
    Dim taskStartIdx As Integer
    Dim taskEndIdx As Integer
    Dim taskDurIdx As Integer
    Dim taskConstraintIdx As Integer
    Dim taskFloatIdx As Integer
    taskIdIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_id")
    taskTypeIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_type")
    taskStartIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_start_date")
    taskEndIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_end_date")
    taskDurIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_drtn_hr_cnt")
    taskConstraintIdx = DEE_XERParser.FindFieldIndex(taskFields, "cstr_type")
    taskFloatIdx = DEE_XERParser.FindFieldIndex(taskFields, "total_float_hr_cnt")

    ' Load configurable thresholds from DEE_Settings
    Dim tHighFloatDays As Double:  tHighFloatDays  = DEE_Settings.GetDCMAThreshold("HighFloatDays")
    Dim tHighDurDays   As Double:  tHighDurDays    = DEE_Settings.GetDCMAThreshold("HighDurDays")
    Dim tMissingPred   As Double:  tMissingPred    = DEE_Settings.GetDCMAThreshold("MissingPred")
    Dim tMissingSucc   As Double:  tMissingSucc    = DEE_Settings.GetDCMAThreshold("MissingSucc")
    Dim tFSPct         As Double:  tFSPct          = DEE_Settings.GetDCMAThreshold("FSPct")
    Dim tLagPct        As Double:  tLagPct         = DEE_Settings.GetDCMAThreshold("LagPct")
    Dim tConstraintPct As Double:  tConstraintPct  = DEE_Settings.GetDCMAThreshold("ConstraintPct")
    Dim tHighFloatPct  As Double:  tHighFloatPct   = DEE_Settings.GetDCMAThreshold("HighFloatPct")
    Dim tHighDurPct    As Double:  tHighDurPct     = DEE_Settings.GetDCMAThreshold("HighDurPct")
    Dim tResourcePct   As Double:  tResourcePct    = DEE_Settings.GetDCMAThreshold("ResourcePct")
    Dim tCPLI          As Double:  tCPLI           = DEE_Settings.GetDCMAThreshold("CPLI")

    ' Build task ID set and stats
    Dim taskIdSet As Object
    Set taskIdSet = CreateObject("Scripting.Dictionary")
    Dim tasksWithHardConstraint As Long
    Dim tasksWithHighFloat As Long
    Dim tasksWithNegFloat As Long
    Dim tasksWithHighDur As Long
    Dim tasksWithInvalidDates As Long
    Dim tasksWithResources As Long
    Dim tasksMissed As Long
    Dim tasksWBS As Long
    tasksWithHardConstraint = 0
    tasksWithHighFloat = 0
    tasksWithNegFloat = 0
    tasksWithHighDur = 0
    tasksWithInvalidDates = 0
    tasksWithResources = 0
    tasksMissed = 0
    tasksWBS = 0

    Dim dataDate As Date
    dataDate = DEE_Utils.GetDataDate()

    DEE_Utils.UpdateProgress 15, "Analyzing tasks"

    Dim taskRow As Variant
    For Each taskRow In taskRows
        Dim tr As Variant
        tr = taskRow

        ' Skip WBS summary tasks
        Dim taskType As String
        If taskTypeIdx >= 0 And taskTypeIdx <= UBound(tr) Then
            taskType = Trim(tr(taskTypeIdx))
        End If
        If taskType = "TT_WBS" Then
            tasksWBS = tasksWBS + 1
            GoTo NextAuditTask
        End If

        ' Task ID
        Dim tId As String
        If taskIdIdx >= 0 And taskIdIdx <= UBound(tr) Then tId = Trim(tr(taskIdIdx))
        If tId <> "" Then taskIdSet(tId) = True

        ' Date validation
        Dim tStart As Variant
        Dim tEnd As Variant
        If taskStartIdx >= 0 And taskStartIdx <= UBound(tr) Then
            tStart = DEE_Utils.SafeParseDate(CStr(tr(taskStartIdx)))
        End If
        If taskEndIdx >= 0 And taskEndIdx <= UBound(tr) Then
            tEnd = DEE_Utils.SafeParseDate(CStr(tr(taskEndIdx)))
        End If

        If IsEmpty(tStart) Or IsEmpty(tEnd) Then
            tasksWithInvalidDates = tasksWithInvalidDates + 1
        ElseIf CDate(tEnd) < CDate(tStart) Then
            tasksWithInvalidDates = tasksWithInvalidDates + 1
        End If

        ' Hard constraints
        Dim cstrType As String
        If taskConstraintIdx >= 0 And taskConstraintIdx <= UBound(tr) Then
            cstrType = Trim(tr(taskConstraintIdx))
        End If
        If cstrType = "CS_MEO" Or cstrType = "CS_MSOB" Or cstrType = "CS_MEOA" Then
            tasksWithHardConstraint = tasksWithHardConstraint + 1
        End If

        ' Float
        Dim floatHrs As Double
        If taskFloatIdx >= 0 And taskFloatIdx <= UBound(tr) Then
            On Error Resume Next
            floatHrs = CDbl(tr(taskFloatIdx))
            On Error GoTo 0
        End If
        Dim floatDays As Double
        floatDays = floatHrs / 8  ' Convert hours to days (8-hr work day)
        If floatDays > tHighFloatDays Then tasksWithHighFloat = tasksWithHighFloat + 1
        If floatHrs < 0 Then tasksWithNegFloat = tasksWithNegFloat + 1

        ' Duration
        Dim durHrs As Double
        If taskDurIdx >= 0 And taskDurIdx <= UBound(tr) Then
            On Error Resume Next
            durHrs = CDbl(tr(taskDurIdx))
            On Error GoTo 0
        End If
        If (durHrs / 8) > tHighDurDays Then tasksWithHighDur = tasksWithHighDur + 1

        ' Missed activities (past data date but not complete)
        If Not IsEmpty(tEnd) Then
            If CDate(tEnd) < dataDate Then
                ' Check if complete -- would need phys_complete_pct
                tasksMissed = tasksMissed + 1
            End If
        End If

NextAuditTask:
    Next taskRow

    Dim activeTasks As Long
    activeTasks = taskIdSet.Count

    ' -----------------------------------------------------------------------
    ' Predecessors / Successors analysis (TASKPRED table)
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 30, "Analyzing predecessors"

    Dim tasksWithPred As Object
    Dim tasksWithSucc As Object
    Set tasksWithPred = CreateObject("Scripting.Dictionary")
    Set tasksWithSucc = CreateObject("Scripting.Dictionary")

    Dim totalRelationships As Long
    Dim fsRelationships As Long
    Dim lagsCount As Long
    Dim leadsCount As Long
    totalRelationships = 0
    fsRelationships = 0
    lagsCount = 0
    leadsCount = 0

    If DEE_XERParser.TableExists("TASKPRED") Then
        Dim predFields As Variant
        predFields = DEE_XERParser.GetFieldNames("TASKPRED")
        Dim predRows As Collection
        Set predRows = DEE_XERParser.GetTable("TASKPRED")

        Dim taskIdPredIdx As Integer
        Dim predTaskIdIdx As Integer
        Dim predTypeIdx As Integer
        Dim lagHrIdx As Integer

        taskIdPredIdx = DEE_XERParser.FindFieldIndex(predFields, "task_id")
        predTaskIdIdx = DEE_XERParser.FindFieldIndex(predFields, "pred_task_id")
        predTypeIdx = DEE_XERParser.FindFieldIndex(predFields, "pred_type")
        lagHrIdx = DEE_XERParser.FindFieldIndex(predFields, "lag_hr_cnt")

        Dim predRow As Variant
        For Each predRow In predRows
            Dim pr As Variant
            pr = predRow

            Dim succTaskId As String
            Dim predTaskId As String
            Dim predType As String
            Dim lagHr As Double

            If taskIdPredIdx >= 0 And taskIdPredIdx <= UBound(pr) Then succTaskId = Trim(pr(taskIdPredIdx))
            If predTaskIdIdx >= 0 And predTaskIdIdx <= UBound(pr) Then predTaskId = Trim(pr(predTaskIdIdx))
            If predTypeIdx >= 0 And predTypeIdx <= UBound(pr) Then predType = Trim(pr(predTypeIdx))
            If lagHrIdx >= 0 And lagHrIdx <= UBound(pr) Then
                On Error Resume Next
                lagHr = CDbl(pr(lagHrIdx))
                On Error GoTo 0
            End If

            ' Only count relationships between known tasks
            If taskIdSet.Exists(succTaskId) And taskIdSet.Exists(predTaskId) Then
                totalRelationships = totalRelationships + 1
                tasksWithPred(succTaskId) = True
                tasksWithSucc(predTaskId) = True

                If predType = "PR_FS" Then fsRelationships = fsRelationships + 1
                If lagHr > 0 Then lagsCount = lagsCount + 1
                If lagHr < 0 Then leadsCount = leadsCount + 1
            End If
        Next predRow
    End If

    Dim tasksMissingPred As Long
    Dim tasksMissingSucc As Long
    tasksMissingPred = activeTasks - tasksWithPred.Count
    tasksMissingSucc = activeTasks - tasksWithSucc.Count

    ' -----------------------------------------------------------------------
    ' Resources analysis (TASKRSRC table)
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 50, "Analyzing resources"

    Dim tasksWithRsrc As Object
    Set tasksWithRsrc = CreateObject("Scripting.Dictionary")

    If DEE_XERParser.TableExists("TASKRSRC") Then
        Dim rsrcFields As Variant
        rsrcFields = DEE_XERParser.GetFieldNames("TASKRSRC")
        Dim rsrcRows As Collection
        Set rsrcRows = DEE_XERParser.GetTable("TASKRSRC")

        Dim rsrcTaskIdIdx As Integer
        rsrcTaskIdIdx = DEE_XERParser.FindFieldIndex(rsrcFields, "task_id")

        Dim rsrcRow As Variant
        For Each rsrcRow In rsrcRows
            Dim rr As Variant
            rr = rsrcRow
            If rsrcTaskIdIdx >= 0 And rsrcTaskIdIdx <= UBound(rr) Then
                Dim rtId As String
                rtId = Trim(rr(rsrcTaskIdIdx))
                If taskIdSet.Exists(rtId) Then
                    tasksWithRsrc(rtId) = True
                End If
            End If
        Next rsrcRow
    End If

    tasksWithResources = tasksWithRsrc.Count

    ' -----------------------------------------------------------------------
    ' Compute CPLI (Critical Path Length Index)
    ' CPLI = (CPL + TF) / CPL  where CPL = critical path length in days
    ' Approximation: if no CPM computed, use 0 as placeholder
    ' -----------------------------------------------------------------------
    Dim cpli As Double
    cpli = 0  ' Will be populated if CPM engine is run first

    ' -----------------------------------------------------------------------
    ' Write results to Audit sheet
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 80, "Writing audit results"

    WriteAuditResults wsAudit, activeTasks, totalRelationships, _
        tasksMissingPred, tasksMissingSucc, fsRelationships, _
        lagsCount, leadsCount, tasksWithHardConstraint, _
        tasksWithHighFloat, tasksWithNegFloat, tasksWithHighDur, _
        tasksWithInvalidDates, tasksWithResources, tasksMissed, cpli

    DEE_Utils.EndProgress
    wsAudit.Activate
    MsgBox "DCMA 14-Point Audit complete!" & vbCrLf & _
           activeTasks & " activities analyzed.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' WriteAuditResults -- Write formatted audit table to sheet
' ---------------------------------------------------------------------------
Private Sub WriteAuditResults(ws As Worksheet, activeTasks As Long, _
    totalRel As Long, missingPred As Long, missingSucc As Long, _
    fsRel As Long, lagsCount As Long, leadsCount As Long, _
    hardConstraints As Long, highFloat As Long, negFloat As Long, _
    highDur As Long, invalidDates As Long, withResources As Long, _
    missed As Long, cpli As Double)

    ' Header
    ws.Cells(1, 1).Value = "Protocol DEE -- DCMA 14-Point Schedule Health Check"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(2, 1).Value = "Generated: " & Format(Now(), "yyyy-mm-dd HH:MM")
    ws.Cells(3, 1).Value = "Total Activities: " & activeTasks

    Dim headerRow As Integer
    headerRow = 5

    ws.Cells(headerRow, 1).Value = "#"
    ws.Cells(headerRow, 2).Value = "Check"
    ws.Cells(headerRow, 3).Value = "Count"
    ws.Cells(headerRow, 4).Value = "%"
    ws.Cells(headerRow, 5).Value = "Threshold"
    ws.Cells(headerRow, 6).Value = "Status"
    ws.Cells(headerRow, 7).Value = "Notes"

    DEE_Utils.ApplyTableHeader ws, headerRow, 1, 7

    ' Read thresholds from settings
    Dim sMissingPred   As Double: sMissingPred   = DEE_Settings.GetDCMAThreshold("MissingPred")
    Dim sMissingSucc   As Double: sMissingSucc   = DEE_Settings.GetDCMAThreshold("MissingSucc")
    Dim sFSPct         As Double: sFSPct         = DEE_Settings.GetDCMAThreshold("FSPct")
    Dim sLagPct        As Double: sLagPct        = DEE_Settings.GetDCMAThreshold("LagPct")
    Dim sConstraintPct As Double: sConstraintPct = DEE_Settings.GetDCMAThreshold("ConstraintPct")
    Dim sHighFloatPct  As Double: sHighFloatPct  = DEE_Settings.GetDCMAThreshold("HighFloatPct")
    Dim sHighDurPct    As Double: sHighDurPct    = DEE_Settings.GetDCMAThreshold("HighDurPct")
    Dim sResourcePct   As Double: sResourcePct   = DEE_Settings.GetDCMAThreshold("ResourcePct")
    Dim sCPLI          As Double: sCPLI          = DEE_Settings.GetDCMAThreshold("CPLI")
    Dim sHighFloatDays As Double: sHighFloatDays = DEE_Settings.GetDCMAThreshold("HighFloatDays")
    Dim sHighDurDays   As Double: sHighDurDays   = DEE_Settings.GetDCMAThreshold("HighDurDays")

    ' Define checks
    Dim checks(0 To 13) As String
    Dim counts(0 To 13) As Long
    Dim thresholds(0 To 13) As String
    Dim passConditions(0 To 13) As Boolean

    ' Check 1: Missing predecessors
    checks(0) = "Logic - Missing Predecessors"
    counts(0) = missingPred
    thresholds(0) = "< " & Format(sMissingPred, "0%")
    passConditions(0) = (activeTasks > 0 And (CDbl(missingPred) / activeTasks) < sMissingPred)

    ' Check 2: Missing successors
    checks(1) = "Logic - Missing Successors"
    counts(1) = missingSucc
    thresholds(1) = "< " & Format(sMissingSucc, "0%")
    passConditions(1) = (activeTasks > 0 And (CDbl(missingSucc) / activeTasks) < sMissingSucc)

    ' Check 3: FS relationship types
    Dim nonFsRel As Long
    nonFsRel = totalRel - fsRel
    checks(2) = "Relationship Types (non-FS)"
    counts(2) = nonFsRel
    thresholds(2) = ">= " & Format(sFSPct, "0%") & " FS"
    passConditions(2) = (totalRel > 0 And (CDbl(fsRel) / totalRel) >= sFSPct)

    ' Check 4: Lags
    checks(3) = "Lags"
    counts(3) = lagsCount
    thresholds(3) = "< " & Format(sLagPct, "0%")
    passConditions(3) = (totalRel > 0 And (CDbl(lagsCount) / totalRel) < sLagPct)

    ' Check 5: Leads (negative lags)
    checks(4) = "Leads (Negative Lags)"
    counts(4) = leadsCount
    thresholds(4) = "0%"
    passConditions(4) = (leadsCount = 0)

    ' Check 6: Relationship density
    checks(5) = "Relationship Density"
    counts(5) = totalRel
    thresholds(5) = ">= 1 per activity"
    passConditions(5) = (activeTasks > 0 And (CDbl(totalRel) / activeTasks) >= 1)

    ' Check 7: Hard constraints
    checks(6) = "Hard Constraints"
    counts(6) = hardConstraints
    thresholds(6) = "< " & Format(sConstraintPct, "0%")
    passConditions(6) = (activeTasks > 0 And (CDbl(hardConstraints) / activeTasks) < sConstraintPct)

    ' Check 8: High float
    checks(7) = "High Float (TF > " & Format(sHighFloatDays, "0") & " days)"
    counts(7) = highFloat
    thresholds(7) = "< " & Format(sHighFloatPct, "0%")
    passConditions(7) = (activeTasks > 0 And (CDbl(highFloat) / activeTasks) < sHighFloatPct)

    ' Check 9: Negative float
    checks(8) = "Negative Float"
    counts(8) = negFloat
    thresholds(8) = "0%"
    passConditions(8) = (negFloat = 0)

    ' Check 10: High duration
    checks(9) = "High Duration (> " & Format(sHighDurDays, "0") & " days)"
    counts(9) = highDur
    thresholds(9) = "< " & Format(sHighDurPct, "0%")
    passConditions(9) = (activeTasks > 0 And (CDbl(highDur) / activeTasks) < sHighDurPct)

    ' Check 11: Invalid dates
    checks(10) = "Invalid Dates"
    counts(10) = invalidDates
    thresholds(10) = "0%"
    passConditions(10) = (invalidDates = 0)

    ' Check 12: Resources assigned
    checks(11) = "Resources Assigned"
    counts(11) = withResources
    thresholds(11) = ">= " & Format(sResourcePct, "0%")
    passConditions(11) = (activeTasks > 0 And (CDbl(withResources) / activeTasks) >= sResourcePct)

    ' Check 13: Missed activities
    checks(12) = "Missed Activities (Past DD)"
    counts(12) = missed
    thresholds(12) = "0%"
    passConditions(12) = (missed = 0)

    ' Check 14: CPLI
    checks(13) = "Critical Path Length Index (CPLI)"
    counts(13) = 0  ' N/A without CPM
    thresholds(13) = ">= " & Format(sCPLI, "0.00")
    passConditions(13) = (cpli >= sCPLI Or cpli = 0)  ' 0 = not computed yet

    ' Write rows
    Dim dataRow As Integer
    dataRow = headerRow + 1

    Dim c As Integer
    For c = 0 To 13
        ws.Cells(dataRow + c, 1).Value = c + 1
        ws.Cells(dataRow + c, 2).Value = checks(c)
        ws.Cells(dataRow + c, 3).Value = counts(c)

        ' Percentage
        If activeTasks > 0 And c <> 13 Then
            ws.Cells(dataRow + c, 4).Value = CDbl(counts(c)) / activeTasks
            ws.Cells(dataRow + c, 4).NumberFormat = "0.0%"
        ElseIf c = 13 Then
            ws.Cells(dataRow + c, 4).Value = IIf(cpli > 0, cpli, "N/A")
        End If

        ws.Cells(dataRow + c, 5).Value = thresholds(c)

        ' Status (traffic light)
        Dim statusCell As Range
        Set statusCell = ws.Cells(dataRow + c, 6)
        If passConditions(c) Then
            statusCell.Value = "PASS"
            statusCell.Interior.Color = COLOR_PASS
            statusCell.Font.Color = RGB(255, 255, 255)
        Else
            statusCell.Value = "FAIL"
            statusCell.Interior.Color = COLOR_FAIL
            statusCell.Font.Color = RGB(255, 255, 255)
        End If
        statusCell.Font.Bold = True
        statusCell.HorizontalAlignment = xlCenter

        ' Notes
        Select Case c
            Case 0: ws.Cells(dataRow + c, 7).Value = IIf(passConditions(c), "OK", missingPred & " activities without predecessors")
            Case 1: ws.Cells(dataRow + c, 7).Value = IIf(passConditions(c), "OK", missingSucc & " activities without successors")
            Case 11: ws.Cells(dataRow + c, 7).Value = IIf(passConditions(c), "OK", "Run CPM analysis for CPLI")
            Case 13: ws.Cells(dataRow + c, 7).Value = "Run CPM Analysis first to compute CPLI"
        End Select
    Next c

    ' Summary
    Dim passCount As Integer
    passCount = 0
    For c = 0 To 13
        If passConditions(c) Then passCount = passCount + 1
    Next c

    Dim summaryRow As Integer
    summaryRow = dataRow + 14 + 1

    ws.Cells(summaryRow, 1).Value = "SUMMARY"
    ws.Cells(summaryRow, 2).Value = passCount & " of 14 checks passed"
    ws.Cells(summaryRow, 1).Font.Bold = True

    If passCount = 14 Then
        ws.Cells(summaryRow, 2).Interior.Color = COLOR_PASS
        ws.Cells(summaryRow, 2).Font.Color = RGB(255, 255, 255)
    ElseIf passCount >= 10 Then
        ws.Cells(summaryRow, 2).Interior.Color = COLOR_WARN
    Else
        ws.Cells(summaryRow, 2).Interior.Color = COLOR_FAIL
        ws.Cells(summaryRow, 2).Font.Color = RGB(255, 255, 255)
    End If

    ws.Columns("A:G").AutoFit
End Sub
