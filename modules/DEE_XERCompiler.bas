Attribute VB_Name = "DEE_XERCompiler"
Option Explicit

' =============================================================================
' DEE_XERCompiler.bas -- XER Export (Schedule Sheet -> XER File)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Reads the Schedule sheet (compiled hierarchy) and reconstructs a valid P6
' XER file. Input is ALWAYS the Schedule sheet, NOT raw XER memory.
'
' AUTO-GENERATES all required P6 support tables:
'   CURRTYPE, FINTMPL, OBS, PROJECT, CALENDAR, SCHEDOPTIONS,
'   PROJWBS, TASK, TASKRSRC (resource assignments)
'
' XER FORMAT:
'   ERMHDR<TAB>...
'   %T<TAB>TABLENAME
'   %F<TAB>field1<TAB>field2...
'   %R<TAB>val1<TAB>val2...
'   %E
' =============================================================================

Private Const XER_VERSION As String = "22.12"
Private Const PROJ_ID As String = "1000"
Private Const CLNDR_ID As String = "1"
Private Const OBS_ID As String = "1"
Private Const CURR_ID As String = "1"

' ---------------------------------------------------------------------------
' ExportXER -- Main entry point
' ---------------------------------------------------------------------------
Public Sub ExportXER()
    ' Get Schedule sheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Schedule")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Schedule sheet not found. Cannot export XER.", vbCritical, "Protocol DEE"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        MsgBox "Schedule sheet is empty.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    ' Get save path
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Export XER File"
        .Filters.Clear
        .Filters.Add "XER Files", "*.xer"
        .InitialFileName = ActiveWorkbook.Name & "_export.xer"
        If .Show <> -1 Then Exit Sub
        Dim savePath As String
        savePath = .SelectedItems(1)
    End With

    If Right(LCase(savePath), 4) <> ".xer" Then savePath = savePath & ".xer"

    DEE_Utils.StartProgress "Exporting XER"

    ' Parse schedule
    Dim wbsRows() As XERWBSRow
    Dim taskRows() As XERTaskRow
    Dim wbsCount As Long
    Dim taskCount As Long
    ParseScheduleSheet ws, lastRow, wbsRows, wbsCount, taskRows, taskCount

    ' Write XER
    Dim fileNum As Integer
    fileNum = FreeFile()
    Open savePath For Output As #fileNum

    ' ERMHDR
    Print #fileNum, "ERMHDR" & vbTab & XER_VERSION & vbTab & Format(Date, "yyyy-mm-dd") & _
        vbTab & "Project" & vbTab & "admin" & vbTab & "admin" & vbTab & "Project" & _
        vbTab & "US" & vbTab & "0" & vbTab & "6.2"

    DEE_Utils.UpdateProgress 10, "Writing CURRTYPE"
    WriteCurrType fileNum

    DEE_Utils.UpdateProgress 20, "Writing OBS"
    WriteOBS fileNum

    DEE_Utils.UpdateProgress 30, "Writing PROJECT"
    WriteProject fileNum, ActiveWorkbook.Name

    DEE_Utils.UpdateProgress 40, "Writing CALENDAR"
    WriteCalendar fileNum

    DEE_Utils.UpdateProgress 50, "Writing SCHEDOPTIONS"
    WriteSchedOptions fileNum

    DEE_Utils.UpdateProgress 60, "Writing PROJWBS"
    WritePROJWBS fileNum, wbsRows, wbsCount

    DEE_Utils.UpdateProgress 70, "Writing TASK"
    WriteTask fileNum, taskRows, taskCount

    DEE_Utils.UpdateProgress 85, "Writing TASKRSRC"
    WriteTaskRsrc fileNum, taskRows, taskCount

    ' End of file
    Print #fileNum, "%E"

    Close #fileNum

    DEE_Utils.EndProgress
    MsgBox "XER exported successfully to:" & vbCrLf & savePath, vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' XER row type definitions
' ---------------------------------------------------------------------------
Private Type XERWBSRow
    wbsId As String
    parentId As String
    shortName As String
    fullName As String
    level As Integer
    seqNum As Integer
End Type

Private Type XERTaskRow
    taskId As String
    wbsId As String
    taskCode As String
    taskName As String
    startDate As Variant
    endDate As Variant
    budget As Double
End Type

' ---------------------------------------------------------------------------
' ParseScheduleSheet -- Read Schedule sheet into typed arrays
' ---------------------------------------------------------------------------
Private Sub ParseScheduleSheet(ws As Worksheet, lastRow As Long, _
                                ByRef wbsRows() As XERWBSRow, ByRef wbsCount As Long, _
                                ByRef taskRows() As XERTaskRow, ByRef taskCount As Long)
    ReDim wbsRows(0 To lastRow)
    ReDim taskRows(0 To lastRow)
    wbsCount = 0
    taskCount = 0

    ' Build WBS ID map from indentation level
    Dim wbsIdMap() As String
    ReDim wbsIdMap(0 To 20)  ' Level -> current wbs_id
    Dim wbsSeq As Integer
    wbsSeq = 10000
    Dim taskSeq As Integer
    taskSeq = 20000

    Dim i As Long
    For i = 2 To lastRow
        Dim cellA As String
        cellA = CStr(ws.Cells(i, 1).Value)
        Dim cellB As String
        cellB = Trim(ws.Cells(i, 2).Value)

        If DEE_Utils.IsWBSRow(ws, i) Then
            ' WBS row
            Dim wbsLevel As Integer
            wbsLevel = DEE_Utils.GetWBSLevel(cellA)

            Dim newId As String
            newId = CStr(wbsSeq)
            wbsSeq = wbsSeq + 1

            Dim parentId As String
            parentId = ""
            If wbsLevel > 1 And wbsLevel - 1 <= 20 Then
                parentId = wbsIdMap(wbsLevel - 1)
            End If

            ' Record this WBS at its level
            If wbsLevel <= 20 Then wbsIdMap(wbsLevel) = newId

            wbsRows(wbsCount).wbsId = newId
            wbsRows(wbsCount).parentId = parentId
            wbsRows(wbsCount).shortName = Trim(cellA)
            wbsRows(wbsCount).fullName = Trim(cellA)
            wbsRows(wbsCount).level = wbsLevel
            wbsRows(wbsCount).seqNum = wbsCount * 10
            wbsCount = wbsCount + 1

        Else
            ' Activity row
            Dim taskCode As String
            taskCode = Trim(cellA)

            ' Get parent WBS (look up from indentation)
            Dim taskLevel As Integer
            taskLevel = DEE_Utils.GetWBSLevel(cellA)
            Dim taskWbsId As String
            taskWbsId = ""
            If taskLevel > 0 And taskLevel <= 20 Then
                taskWbsId = wbsIdMap(taskLevel - 1)
            End If
            If taskWbsId = "" And wbsCount > 0 Then
                taskWbsId = wbsRows(wbsCount - 1).wbsId
            End If

            taskRows(taskCount).taskId = CStr(taskSeq)
            taskRows(taskCount).wbsId = taskWbsId
            taskRows(taskCount).taskCode = taskCode
            taskRows(taskCount).taskName = cellB
            taskRows(taskCount).startDate = ws.Cells(i, 3).Value
            taskRows(taskCount).endDate = ws.Cells(i, 4).Value
            taskRows(taskCount).budget = CDblSafe(ws.Cells(i, 6).Value)

            taskSeq = taskSeq + 1
            taskCount = taskCount + 1
        End If
    Next i
End Sub

' ---------------------------------------------------------------------------
' WriteCurrType -- Write CURRTYPE table
' ---------------------------------------------------------------------------
Private Sub WriteCurrType(fileNum As Integer)
    Print #fileNum, "%T" & vbTab & "CURRTYPE"
    Print #fileNum, "%F" & vbTab & "curr_id" & vbTab & "decimal_digit_cnt" & vbTab & _
        "curr_symbol" & vbTab & "curr_short_name" & vbTab & "base_exch_rate"
    Print #fileNum, "%R" & vbTab & CURR_ID & vbTab & "2" & vbTab & "$" & vbTab & "USD" & vbTab & "1"
End Sub

' ---------------------------------------------------------------------------
' WriteOBS -- Write OBS table
' ---------------------------------------------------------------------------
Private Sub WriteOBS(fileNum As Integer)
    Print #fileNum, "%T" & vbTab & "OBS"
    Print #fileNum, "%F" & vbTab & "obs_id" & vbTab & "parent_obs_id" & vbTab & _
        "obs_name" & vbTab & "obs_short_name"
    Print #fileNum, "%R" & vbTab & OBS_ID & vbTab & "" & vbTab & "Project Team" & vbTab & "TEAM"
End Sub

' ---------------------------------------------------------------------------
' WriteProject -- Write PROJECT table
' ---------------------------------------------------------------------------
Private Sub WriteProject(fileNum As Integer, projName As String)
    Print #fileNum, "%T" & vbTab & "PROJECT"
    Print #fileNum, "%F" & vbTab & "proj_id" & vbTab & "fy_start_month_num" & vbTab & _
        "clndr_id" & vbTab & "sum_base_proj_id" & vbTab & "last_recalc_date" & vbTab & _
        "plan_start_date" & vbTab & "scd_end_date" & vbTab & "add_date" & vbTab & _
        "last_tasksum_date" & vbTab & "proj_short_name" & vbTab & "obs_id" & vbTab & _
        "orig_proj_id" & vbTab & "source_proj_id" & vbTab & "base_type_id" & vbTab & "guid"
    Print #fileNum, "%R" & vbTab & PROJ_ID & vbTab & "1" & vbTab & CLNDR_ID & vbTab & "" & _
        vbTab & Format(Date, "yyyy-mm-dd") & " 08:00" & vbTab & _
        Format(Date, "yyyy-mm-dd") & " 08:00" & vbTab & _
        Format(Date + 365, "yyyy-mm-dd") & " 17:00" & vbTab & _
        Format(Date, "yyyy-mm-dd") & " 08:00" & vbTab & _
        Format(Date, "yyyy-mm-dd") & " 08:00" & vbTab & _
        Left(projName, 20) & vbTab & OBS_ID & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
        CreateGUID()
End Sub

' ---------------------------------------------------------------------------
' WriteCalendar -- Write CALENDAR table (standard 5-day work week)
' ---------------------------------------------------------------------------
Private Sub WriteCalendar(fileNum As Integer)
    Print #fileNum, "%T" & vbTab & "CALENDAR"
    Print #fileNum, "%F" & vbTab & "clndr_id" & vbTab & "default_flag" & vbTab & _
        "clndr_name" & vbTab & "proj_id" & vbTab & "clndr_type" & vbTab & _
        "day_hr_cnt" & vbTab & "week_hr_cnt" & vbTab & "month_hr_cnt" & vbTab & "year_hr_cnt"
    Print #fileNum, "%R" & vbTab & CLNDR_ID & vbTab & "Y" & vbTab & "Standard 5-Day" & _
        vbTab & PROJ_ID & vbTab & "CA_BASE" & vbTab & "8" & vbTab & "40" & vbTab & "160" & vbTab & "2080"
End Sub

' ---------------------------------------------------------------------------
' WriteSchedOptions -- Write SCHEDOPTIONS table
' ---------------------------------------------------------------------------
Private Sub WriteSchedOptions(fileNum As Integer)
    Print #fileNum, "%T" & vbTab & "SCHEDOPTIONS"
    Print #fileNum, "%F" & vbTab & "schedoptions_id" & vbTab & "proj_id" & vbTab & _
        "sched_type" & vbTab & "sched_calendar_on_relationship_lag" & vbTab & _
        "sched_use_expect_end_flag" & vbTab & "sched_progress_override"
    Print #fileNum, "%R" & vbTab & "1" & vbTab & PROJ_ID & vbTab & "RA" & vbTab & "N" & _
        vbTab & "Y" & vbTab & "N"
End Sub

' ---------------------------------------------------------------------------
' WritePROJWBS -- Write PROJWBS table
' ---------------------------------------------------------------------------
Private Sub WritePROJWBS(fileNum As Integer, wbsRows() As XERWBSRow, wbsCount As Long)
    Print #fileNum, "%T" & vbTab & "PROJWBS"
    Print #fileNum, "%F" & vbTab & "wbs_id" & vbTab & "proj_id" & vbTab & "obs_id" & vbTab & _
        "seq_num" & vbTab & "est_wt" & vbTab & "proj_node_flag" & vbTab & "sum_data_flag" & _
        vbTab & "status_code" & vbTab & "wbs_short_name" & vbTab & "wbs_name" & vbTab & _
        "phase_id" & vbTab & "parent_wbs_id" & vbTab & "ev_user_pct" & vbTab & "ev_etc_user_value"

    Dim i As Long
    For i = 0 To wbsCount - 1
        Dim wr As XERWBSRow
        wr = wbsRows(i)
        Print #fileNum, "%R" & vbTab & wr.wbsId & vbTab & PROJ_ID & vbTab & OBS_ID & _
            vbTab & wr.seqNum & vbTab & "1" & vbTab & "N" & vbTab & "N" & vbTab & "WS_Open" & _
            vbTab & wr.shortName & vbTab & wr.fullName & vbTab & "" & vbTab & wr.parentId & _
            vbTab & "0" & vbTab & "0"
    Next i
End Sub

' ---------------------------------------------------------------------------
' WriteTask -- Write TASK table
' ---------------------------------------------------------------------------
Private Sub WriteTask(fileNum As Integer, taskRows() As XERTaskRow, taskCount As Long)
    Print #fileNum, "%T" & vbTab & "TASK"
    Print #fileNum, "%F" & vbTab & "task_id" & vbTab & "proj_id" & vbTab & "wbs_id" & vbTab & _
        "clndr_id" & vbTab & "phys_complete_pct" & vbTab & "rev_fdbk_flag" & vbTab & _
        "est_wt" & vbTab & "lock_plan_flag" & vbTab & "auto_compute_act_flag" & vbTab & _
        "complete_pct_type" & vbTab & "task_type" & vbTab & "duration_type" & vbTab & _
        "status_code" & vbTab & "task_code" & vbTab & "task_name" & vbTab & _
        "target_work_qty" & vbTab & "target_equip_qty" & vbTab & "target_start_date" & _
        vbTab & "target_end_date" & vbTab & "target_dur_hr_cnt" & vbTab & "clndr_id" & _
        vbTab & "budget_qty"

    Dim i As Long
    For i = 0 To taskCount - 1
        Dim tr As XERTaskRow
        tr = taskRows(i)

        Dim startStr As String
        Dim endStr As String
        startStr = FormatXERDate(tr.startDate)
        endStr = FormatXERDate(tr.endDate)

        ' Compute duration in hours (5-day, 8-hour work day approximation)
        Dim durDays As Long
        durDays = 0
        If IsDate(tr.startDate) And IsDate(tr.endDate) Then
            durDays = CDate(tr.endDate) - CDate(tr.startDate)
            If durDays < 0 Then durDays = 0
        End If
        Dim durHrs As Long
        durHrs = durDays * 8

        Print #fileNum, "%R" & vbTab & tr.taskId & vbTab & PROJ_ID & vbTab & tr.wbsId & _
            vbTab & CLNDR_ID & vbTab & "0" & vbTab & "N" & vbTab & "1" & vbTab & "N" & _
            vbTab & "Y" & vbTab & "CP_Phys" & vbTab & "TT_Task" & vbTab & "DT_FixedDrtn" & _
            vbTab & "TK_NotStart" & vbTab & tr.taskCode & vbTab & tr.taskName & _
            vbTab & tr.budget & vbTab & "0" & vbTab & startStr & vbTab & endStr & _
            vbTab & durHrs & vbTab & CLNDR_ID & vbTab & tr.budget
    Next i
End Sub

' ---------------------------------------------------------------------------
' WriteTaskRsrc -- Write TASKRSRC table (required for valid P6 re-import)
' Creates one generic resource assignment per task
' ---------------------------------------------------------------------------
Private Sub WriteTaskRsrc(fileNum As Integer, taskRows() As XERTaskRow, taskCount As Long)
    Print #fileNum, "%T" & vbTab & "TASKRSRC"
    Print #fileNum, "%F" & vbTab & "taskrsrc_id" & vbTab & "task_id" & vbTab & "proj_id" & _
        vbTab & "cost_qty_link_flag" & vbTab & "role_id" & vbTab & "acct_id" & vbTab & _
        "rsrc_id" & vbTab & "pobs_id" & vbTab & "remain_qty" & vbTab & "target_qty" & _
        vbTab & "act_reg_qty" & vbTab & "act_ot_qty" & vbTab & "remain_cost" & vbTab & _
        "target_cost" & vbTab & "act_reg_cost" & vbTab & "act_ot_cost"

    Dim rsrcSeq As Long
    rsrcSeq = 30000

    Dim i As Long
    For i = 0 To taskCount - 1
        If taskRows(i).budget > 0 Then
            Print #fileNum, "%R" & vbTab & rsrcSeq & vbTab & taskRows(i).taskId & _
                vbTab & PROJ_ID & vbTab & "N" & vbTab & "" & vbTab & "" & vbTab & "" & _
                vbTab & "" & vbTab & taskRows(i).budget & vbTab & taskRows(i).budget & _
                vbTab & "0" & vbTab & "0" & vbTab & taskRows(i).budget & vbTab & _
                taskRows(i).budget & vbTab & "0" & vbTab & "0"
            rsrcSeq = rsrcSeq + 1
        End If
    Next i
End Sub

' ---------------------------------------------------------------------------
' FormatXERDate -- Convert Excel date to XER date string
' ---------------------------------------------------------------------------
Private Function FormatXERDate(d As Variant) As String
    If IsDate(d) Then
        FormatXERDate = Format(CDate(d), "yyyy-mm-dd") & " 08:00"
    Else
        FormatXERDate = ""
    End If
End Function

' ---------------------------------------------------------------------------
' CDblSafe -- Safe CDbl conversion
' ---------------------------------------------------------------------------
Private Function CDblSafe(v As Variant) As Double
    On Error Resume Next
    CDblSafe = CDbl(v)
    If Err.Number <> 0 Then CDblSafe = 0
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------------
' CreateGUID -- Generate a simple GUID-like string
' ---------------------------------------------------------------------------
Private Function CreateGUID() As String
    Randomize Timer
    Dim g As String
    g = ""
    Dim i As Integer
    Dim hexChars As String
    hexChars = "0123456789ABCDEF"
    Dim parts As Variant
    parts = Array(8, 4, 4, 4, 12)
    Dim p As Variant
    For Each p In parts
        If Len(g) > 0 Then g = g & "-"
        Dim j As Integer
        For j = 1 To p
            g = g & Mid(hexChars, Int(Rnd() * 16) + 1, 1)
        Next j
    Next p
    CreateGUID = "{" & g & "}"
End Function
