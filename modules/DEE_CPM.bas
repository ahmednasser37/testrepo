Attribute VB_Name = "DEE_CPM"
Option Explicit

' =============================================================================
' DEE_CPM.bas -- Critical Path Method (CPM) Engine
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Computes full network analysis from TASKPRED table (Section 15.1).
'
' ALGORITHM:
'   Forward Pass: EF = ES + Duration; ES = max(EF of predecessors) + lag
'   Backward Pass: LF = min(LS of successors); LS = LF - Duration
'   Total Float = LS - ES (or LF - EF)
'   Free Float = min(ES of successors) - EF
'   Critical Path = activities where Total Float = 0
'
' OUTPUT COLUMNS (added to Schedule sheet):
'   Y  = Early Start (ES)
'   Z  = Early Finish (EF)
'   AA = Late Start (LS)
'   AB = Late Finish (LF)
'   AC = Total Float (TF) in working days
'   AD = Free Float (FF) in working days
'   AE = Critical? (TRUE/FALSE)
'
' VISUAL: Float heatmap on row left border color
' =============================================================================

' Column constants for CPM output (after A-X = 1-24)
Private Const COL_ES As Integer = 25  ' Y
Private Const COL_EF As Integer = 26  ' Z
Private Const COL_LS As Integer = 27  ' AA
Private Const COL_LF As Integer = 28  ' AB
Private Const COL_TF As Integer = 29  ' AC
Private Const COL_FF As Integer = 30  ' AD
Private Const COL_CP As Integer = 31  ' AE

' Heatmap color thresholds
Private Const TF_CRITICAL As Integer = 0
Private Const TF_NEAR_CRITICAL As Integer = 5

Private Const COLOR_CRITICAL As Long = RGB(255, 0, 0)
Private Const COLOR_NEAR_CRITICAL As Long = RGB(255, 153, 0)
Private Const COLOR_HEALTHY As Long = RGB(0, 176, 80)

' ---------------------------------------------------------------------------
' RunCPMAnalysis -- Main entry point
' ---------------------------------------------------------------------------
Public Sub RunCPMAnalysis()
    If Not DEE_XERParser.IsXERParsed() Then
        MsgBox "No XER file loaded. Please load a schedule first.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    DEE_Utils.StartProgress "Running CPM Analysis"

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("Schedule")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Schedule sheet not found.", vbCritical, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' 1. Read TASK data (durations and IDs)
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 10, "Loading task network"

    Dim taskFields As Variant
    taskFields = DEE_XERParser.GetFieldNames("TASK")
    Dim taskRows As Collection
    Set taskRows = DEE_XERParser.GetTable("TASK")

    Dim taskIdIdx As Integer
    Dim taskCodeIdx As Integer
    Dim taskDurIdx As Integer
    Dim taskTypeIdx As Integer
    Dim taskStartIdx As Integer
    Dim taskEndIdx As Integer

    taskIdIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_id")
    taskCodeIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_code")
    taskDurIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_drtn_hr_cnt")
    taskTypeIdx = DEE_XERParser.FindFieldIndex(taskFields, "task_type")
    taskStartIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_start_date")
    taskEndIdx = DEE_XERParser.FindFieldIndex(taskFields, "target_end_date")

    ' Build task node dictionary: task_id -> node data
    Dim nodes As Object
    Set nodes = CreateObject("Scripting.Dictionary")

    ' Node structure (stored as array):
    ' [0] = task_code, [1] = duration_days, [2] = ES, [3] = EF,
    ' [4] = LS, [5] = LF, [6] = TF, [7] = FF, [8] = is_critical

    Dim taskRow As Variant
    For Each taskRow In taskRows
        Dim tr As Variant
        tr = taskRow

        Dim taskType As String
        If taskTypeIdx >= 0 And taskTypeIdx <= UBound(tr) Then
            taskType = Trim(tr(taskTypeIdx))
        End If
        If taskType = "TT_WBS" Then GoTo NextCPMTask

        Dim tId As String
        If taskIdIdx >= 0 And taskIdIdx <= UBound(tr) Then tId = Trim(tr(taskIdIdx))
        If tId = "" Then GoTo NextCPMTask

        Dim tCode As String
        If taskCodeIdx >= 0 And taskCodeIdx <= UBound(tr) Then tCode = Trim(tr(taskCodeIdx))

        ' Duration in days (convert from hours, assuming 8-hr work day)
        Dim durHrs As Double
        If taskDurIdx >= 0 And taskDurIdx <= UBound(tr) Then
            On Error Resume Next
            durHrs = CDbl(tr(taskDurIdx))
            On Error GoTo 0
        End If
        Dim durDays As Double
        durDays = durHrs / 8
        If durDays < 0 Then durDays = 0

        ' Target dates for initial ES/EF
        Dim tStart As Variant
        Dim tEnd As Variant
        tStart = 0
        tEnd = 0
        If taskStartIdx >= 0 And taskStartIdx <= UBound(tr) Then
            Dim parsedStart As Variant
            parsedStart = DEE_Utils.SafeParseDate(CStr(tr(taskStartIdx)))
            If Not IsEmpty(parsedStart) Then tStart = CDbl(parsedStart)
        End If
        If taskEndIdx >= 0 And taskEndIdx <= UBound(tr) Then
            Dim parsedEnd As Variant
            parsedEnd = DEE_Utils.SafeParseDate(CStr(tr(taskEndIdx)))
            If Not IsEmpty(parsedEnd) Then tEnd = CDbl(parsedEnd)
        End If

        Dim node(0 To 8) As Variant
        node(0) = tCode     ' task_code
        node(1) = durDays   ' duration (days)
        node(2) = tStart    ' ES (initialized to target start)
        node(3) = tEnd      ' EF (initialized to target end)
        node(4) = 0         ' LS
        node(5) = 0         ' LF
        node(6) = 0         ' TF
        node(7) = 0         ' FF
        node(8) = False     ' is_critical

        nodes(tId) = node
NextCPMTask:
    Next taskRow

    ' -----------------------------------------------------------------------
    ' 2. Read TASKPRED relationships
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 20, "Loading predecessor network"

    ' predecessors: task_id -> Collection of [pred_task_id, pred_type, lag_days]
    Dim predecessors As Object
    Set predecessors = CreateObject("Scripting.Dictionary")
    ' successors: pred_task_id -> Collection of successor task_ids
    Dim successors As Object
    Set successors = CreateObject("Scripting.Dictionary")

    ' Initialize for all nodes
    Dim nKey As Variant
    For Each nKey In nodes.Keys
        predecessors(CStr(nKey)) = New Collection
        successors(CStr(nKey)) = New Collection
    Next nKey

    If DEE_XERParser.TableExists("TASKPRED") Then
        Dim predFields As Variant
        predFields = DEE_XERParser.GetFieldNames("TASKPRED")
        Dim predRows As Collection
        Set predRows = DEE_XERParser.GetTable("TASKPRED")

        Dim predTaskIdIdx As Integer
        Dim succTaskIdIdx As Integer
        Dim predTypeIdx As Integer
        Dim lagHrIdx As Integer

        succTaskIdIdx = DEE_XERParser.FindFieldIndex(predFields, "task_id")
        predTaskIdIdx = DEE_XERParser.FindFieldIndex(predFields, "pred_task_id")
        predTypeIdx = DEE_XERParser.FindFieldIndex(predFields, "pred_type")
        lagHrIdx = DEE_XERParser.FindFieldIndex(predFields, "lag_hr_cnt")

        Dim predRow As Variant
        For Each predRow In predRows
            Dim pr As Variant
            pr = predRow

            Dim succId As String
            Dim predId As String
            Dim predType As String
            Dim lagHrs As Double

            If succTaskIdIdx >= 0 And succTaskIdIdx <= UBound(pr) Then succId = Trim(pr(succTaskIdIdx))
            If predTaskIdIdx >= 0 And predTaskIdIdx <= UBound(pr) Then predId = Trim(pr(predTaskIdIdx))
            If predTypeIdx >= 0 And predTypeIdx <= UBound(pr) Then predType = Trim(pr(predTypeIdx))
            If lagHrIdx >= 0 And lagHrIdx <= UBound(pr) Then
                On Error Resume Next
                lagHrs = CDbl(pr(lagHrIdx))
                On Error GoTo 0
            End If

            If nodes.Exists(succId) And nodes.Exists(predId) Then
                Dim lagDays As Double
                lagDays = lagHrs / 8

                ' Add to predecessor list of successor
                predecessors(succId).Add Array(predId, predType, lagDays)
                ' Add to successor list of predecessor
                successors(predId).Add succId
            End If
        Next predRow
    End If

    ' -----------------------------------------------------------------------
    ' 3. Forward Pass (Critical Path Forward)
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 40, "Forward pass"

    ' Topological sort (Kahn's algorithm)
    Dim sortedOrder As Collection
    Set sortedOrder = TopologicalSort(nodes, predecessors, successors)

    ' Forward pass: use target dates as initial ES/EF if no predecessors
    Dim taskKey As Variant
    For Each taskKey In sortedOrder
        Dim taskId As String
        taskId = CStr(taskKey)
        If Not nodes.Exists(taskId) Then GoTo NextForward

        Dim nd As Variant
        nd = nodes(taskId)
        Dim dur As Double
        dur = nd(1)

        ' Compute ES from predecessors
        Dim maxEF As Double
        maxEF = nd(2)  ' Default to target start

        Dim predEntry As Variant
        For Each predEntry In predecessors(taskId)
            Dim pe As Variant
            pe = predEntry
            Dim pId As String
            pId = CStr(pe(0))
            Dim pType As String
            pType = CStr(pe(1))
            Dim pLag As Double
            pLag = CDbl(pe(2))

            If nodes.Exists(pId) Then
                Dim pnd As Variant
                pnd = nodes(pId)
                Dim candidateES As Double

                Select Case pType
                    Case "PR_FS": candidateES = pnd(3) + pLag  ' EF of pred + lag
                    Case "PR_SS": candidateES = pnd(2) + pLag  ' ES of pred + lag
                    Case "PR_FF": candidateES = pnd(3) + pLag - dur  ' EF pred + lag - dur
                    Case "PR_SF": candidateES = pnd(2) + pLag - dur  ' ES pred + lag - dur
                    Case Else: candidateES = pnd(3) + pLag
                End Select

                If candidateES > maxEF Then maxEF = candidateES
            End If
        Next predEntry

        nd(2) = maxEF           ' ES
        nd(3) = maxEF + dur     ' EF
        nodes(taskId) = nd
NextForward:
    Next taskKey

    ' -----------------------------------------------------------------------
    ' 4. Backward Pass
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 60, "Backward pass"

    ' Find max EF (project end)
    Dim projectEnd As Double
    projectEnd = 0
    For Each taskKey In nodes.Keys
        Dim tnd As Variant
        tnd = nodes(CStr(taskKey))
        If tnd(3) > projectEnd Then projectEnd = tnd(3)
    Next taskKey

    ' Backward pass (reverse topological order)
    Dim i As Integer
    For i = sortedOrder.Count To 1 Step -1
        taskId = CStr(sortedOrder(i))
        If Not nodes.Exists(taskId) Then GoTo NextBackward

        nd = nodes(taskId)

        ' LF = min(LS of successors); if no successors, LF = project end
        Dim minLS As Double
        minLS = projectEnd

        Dim succId As Variant
        For Each succId In successors(taskId)
            Dim sId As String
            sId = CStr(succId)
            If nodes.Exists(sId) Then
                Dim snd As Variant
                snd = nodes(sId)
                ' For FS relationships: LS of successor
                Dim candidateLF As Double
                candidateLF = snd(4)  ' LS of successor
                If candidateLF < minLS Then minLS = candidateLF
            End If
        Next succId

        nd(5) = minLS           ' LF
        nd(4) = minLS - nd(1)  ' LS = LF - duration
        nd(6) = nd(4) - nd(2)  ' TF = LS - ES
        nd(8) = (Abs(nd(6)) <= 0.001)  ' Critical if TF ~= 0
        nodes(taskId) = nd
NextBackward:
    Next i

    ' -----------------------------------------------------------------------
    ' 5. Free Float computation
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 75, "Computing free float"

    For Each taskKey In nodes.Keys
        taskId = CStr(taskKey)
        nd = nodes(taskId)

        Dim minSuccES As Double
        minSuccES = projectEnd  ' Default

        For Each succId In successors(taskId)
            If nodes.Exists(CStr(succId)) Then
                Dim ssnd As Variant
                ssnd = nodes(CStr(succId))
                If ssnd(2) < minSuccES Then minSuccES = ssnd(2)
            End If
        Next succId

        nd(7) = minSuccES - nd(3)  ' FF = min(ES successors) - EF
        If nd(7) < 0 Then nd(7) = 0
        nodes(taskId) = nd
    Next taskKey

    ' -----------------------------------------------------------------------
    ' 6. Write CPM results to Schedule sheet
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 85, "Writing CPM results to Schedule sheet"

    WriteHeaders ws
    WriteCPMResults ws, nodes

    ' Apply heatmap
    ApplyFloatHeatmap ws, nodes

    DEE_Utils.EndProgress
    MsgBox "CPM Analysis complete!" & vbCrLf & _
           nodes.Count & " activities processed.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' TopologicalSort -- Kahn's algorithm for dependency ordering
' ---------------------------------------------------------------------------
Private Function TopologicalSort(nodes As Object, predecessors As Object, _
                                  successors As Object) As Collection
    Dim result As Collection
    Set result = New Collection

    ' Count in-degrees
    Dim inDegree As Object
    Set inDegree = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In nodes.Keys
        inDegree(CStr(k)) = predecessors(CStr(k)).Count
    Next k

    ' Queue of nodes with in-degree 0
    Dim queue As Collection
    Set queue = New Collection
    For Each k In nodes.Keys
        If inDegree(CStr(k)) = 0 Then queue.Add CStr(k)
    Next k

    Do While queue.Count > 0
        Dim curr As String
        curr = CStr(queue(1))
        queue.Remove 1
        result.Add curr

        ' Reduce in-degree of successors
        If successors.Exists(curr) Then
            Dim succ As Variant
            For Each succ In successors(curr)
                Dim sKey As String
                sKey = CStr(succ)
                If inDegree.Exists(sKey) Then
                    inDegree(sKey) = inDegree(sKey) - 1
                    If inDegree(sKey) = 0 Then queue.Add sKey
                End If
            Next succ
        End If
    Loop

    Set TopologicalSort = result
End Function

' ---------------------------------------------------------------------------
' WriteHeaders -- Add CPM column headers to Schedule sheet
' ---------------------------------------------------------------------------
Private Sub WriteHeaders(ws As Worksheet)
    ws.Cells(1, COL_ES).Value = "ES"
    ws.Cells(1, COL_EF).Value = "EF"
    ws.Cells(1, COL_LS).Value = "LS"
    ws.Cells(1, COL_LF).Value = "LF"
    ws.Cells(1, COL_TF).Value = "TF (days)"
    ws.Cells(1, COL_FF).Value = "FF (days)"
    ws.Cells(1, COL_CP).Value = "Critical"

    Dim c As Integer
    For c = COL_ES To COL_CP
        With ws.Cells(1, c)
            .Interior.Color = RGB(70, 30, 0)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next c

    ws.Range(ws.Cells(1, COL_ES), ws.Cells(1, COL_LF)).NumberFormat = "yyyy-mm-dd"
End Sub

' ---------------------------------------------------------------------------
' WriteCPMResults -- Write ES/EF/LS/LF/TF/FF/CP to Schedule sheet rows
' ---------------------------------------------------------------------------
Private Sub WriteCPMResults(ws As Worksheet, nodes As Object)
    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)

    ' Build task_code -> node lookup
    Dim codeToNode As Object
    Set codeToNode = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In nodes.Keys
        Dim nd As Variant
        nd = nodes(CStr(k))
        Dim code As String
        code = Trim(CStr(nd(0)))
        If code <> "" Then codeToNode(code) = nd
    Next k

    Dim i As Long
    For i = 2 To lastRow
        If Not DEE_Utils.IsWBSRow(ws, i) Then
            Dim cellCode As String
            cellCode = Trim(ws.Cells(i, 1).Value)

            If codeToNode.Exists(cellCode) Then
                Dim nodeData As Variant
                nodeData = codeToNode(cellCode)

                ' Write dates (convert numeric back to Date)
                If nodeData(2) > 0 Then
                    ws.Cells(i, COL_ES).Value = CDate(nodeData(2))
                    ws.Cells(i, COL_ES).NumberFormat = "yyyy-mm-dd"
                End If
                If nodeData(3) > 0 Then
                    ws.Cells(i, COL_EF).Value = CDate(nodeData(3))
                    ws.Cells(i, COL_EF).NumberFormat = "yyyy-mm-dd"
                End If
                If nodeData(4) > 0 Then
                    ws.Cells(i, COL_LS).Value = CDate(nodeData(4))
                    ws.Cells(i, COL_LS).NumberFormat = "yyyy-mm-dd"
                End If
                If nodeData(5) > 0 Then
                    ws.Cells(i, COL_LF).Value = CDate(nodeData(5))
                    ws.Cells(i, COL_LF).NumberFormat = "yyyy-mm-dd"
                End If

                ws.Cells(i, COL_TF).Value = nodeData(6)
                ws.Cells(i, COL_TF).NumberFormat = "0.0"
                ws.Cells(i, COL_FF).Value = nodeData(7)
                ws.Cells(i, COL_FF).NumberFormat = "0.0"
                ws.Cells(i, COL_CP).Value = nodeData(8)
            End If
        End If
    Next i
End Sub

' ---------------------------------------------------------------------------
' ApplyFloatHeatmap -- Color left border of rows based on Total Float
'   Red border   = Critical (TF = 0)
'   Orange border= Near-Critical (TF <= 5)
'   Green border = Healthy (TF > 5)
' ---------------------------------------------------------------------------
Private Sub ApplyFloatHeatmap(ws As Worksheet, nodes As Object)
    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)

    Dim i As Long
    For i = 2 To lastRow
        If Not DEE_Utils.IsWBSRow(ws, i) Then
            Dim tfVal As Variant
            tfVal = ws.Cells(i, COL_TF).Value

            If IsNumeric(tfVal) Then
                Dim tf As Double
                tf = CDbl(tfVal)

                Dim borderColor As Long
                If tf <= TF_CRITICAL Then
                    borderColor = COLOR_CRITICAL
                ElseIf tf <= TF_NEAR_CRITICAL Then
                    borderColor = COLOR_NEAR_CRITICAL
                Else
                    borderColor = COLOR_HEALTHY
                End If

                Dim lastCol As Integer
                lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

                With ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                    .Color = borderColor
                End With
            End If
        End If
    Next i
End Sub
