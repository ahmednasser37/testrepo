Attribute VB_Name = "DEE_Main"
Option Explicit

' =============================================================================
' DEE_Main.bas -- Ribbon Callback Router
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' All btn_* and chk_* callbacks route Ribbon button clicks to module functions.
'
' CALLBACK SIGNATURE RULES:
'   Button onAction:    Public Sub btn_XXX(control As IRibbonControl)
'   CheckBox onAction:  Public Sub chk_XXX(control As IRibbonControl, pressed As Boolean)
'   CheckBox getPressed: Public Sub chk_XXX_getPressed(control As IRibbonControl, ByRef returnedVal)
'   ComboBox onChange:  Public Sub cmb_XXX(control As IRibbonControl, text As String)
' =============================================================================

' Ribbon object reference (for invalidation)
Private m_Ribbon As Object

' ---------------------------------------------------------------------------
' Ribbon Load callback
' ---------------------------------------------------------------------------
Public Sub Ribbon_Load(ribbon As Object)
    Set m_Ribbon = ribbon
End Sub

Public Sub InvalidateRibbon()
    If Not m_Ribbon Is Nothing Then m_Ribbon.Invalidate
End Sub

' ===========================================================================
' WBS GROUP CALLBACKS
' ===========================================================================

Public Sub btn_WBSColor(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "WBSColor started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_WBS.ApplyWBSColoring ActiveSheet
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "WBSColor: " & Err.Description
    MsgBox "WBS Color error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_WBSGroup(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "WBSGroup started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_WBS.ApplyWBSGrouping ActiveSheet, DEE_Utils.GetSkipFirstRow()
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "WBSGroup: " & Err.Description
    MsgBox "WBS Group error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_WBSAggregator(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "WBSAggregator started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_WBS.ApplySumLevels ActiveSheet
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "WBSAggregator: " & Err.Description
    MsgBox "WBS Aggregator error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_WBSColumns(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "WBSColumns started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_WBS.WBSInColumns ActiveSheet
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "WBSColumns: " & Err.Description
    MsgBox "WBS Columns error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub chk_SkipHeader(control As IRibbonControl, pressed As Boolean)
    DEE_Utils.SetSkipFirstRow pressed
End Sub

Public Sub chk_SkipHeader_getPressed(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = DEE_Utils.GetSkipFirstRow()
End Sub

Public Sub cmb_WBSLevelFilter(control As IRibbonControl, text As String)
    On Error GoTo ErrHandler
    Dim level As Integer
    level = 0
    Select Case LCase(Trim(text))
        Case "all levels": level = 0
        Case "level 1":    level = 1
        Case "level 2":    level = 2
        Case "level 3":    level = 3
        Case "level 4":    level = 4
    End Select
    DEE_WBS.FilterByLevel ActiveSheet, level
    Exit Sub
ErrHandler:
    MsgBox "WBS Level Filter error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' TIMELINE GROUP CALLBACKS
' ===========================================================================

Public Sub btn_DrawGantt(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "DrawGantt started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_Gantt.DrawGanttChart ActiveSheet
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "DrawGantt: " & Err.Description
    MsgBox "Draw Gantt error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_ManageBars(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_Gantt.ManageBars ActiveSheet
    Exit Sub
ErrHandler:
    MsgBox "Manage Bars error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_PrintSetup(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_Reporting.SetupPrint
    Exit Sub
ErrHandler:
    MsgBox "Print Setup error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' PROGRESS GROUP CALLBACKS
' ===========================================================================

Public Sub btn_CreatePMS(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "CreatePMS started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_PMS.CreatePMS
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "CreatePMS: " & Err.Description
    MsgBox "Create PMS error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_CreatePMSEV(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "CreatePMSEV started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_PMS.CreatePMSWithEarnedUnits
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "CreatePMSEV: " & Err.Description
    MsgBox "Create PMS (EV) error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_LockPeriod(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "LockPeriod started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Schedule")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        MsgBox "Schedule sheet not found.", vbExclamation, "Protocol DEE"
        GoTo CleanUp
    End If
    DEE_PMS.LockPeriod ws
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "LockPeriod: " & Err.Description
    MsgBox "Lock Period error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_Snapshot(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "Snapshot started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Schedule")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        MsgBox "Schedule sheet not found.", vbExclamation, "Protocol DEE"
        GoTo CleanUp
    End If
    DEE_PMS.TakeSnapshot ws
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "Snapshot: " & Err.Description
    MsgBox "Snapshot error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_Dashboard(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "Dashboard started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_Dashboard.CreateDashboard
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "Dashboard: " & Err.Description
    MsgBox "Dashboard error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_SCurveMonthly(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "SCurveMonthly started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_SCurve.CreateSCurveMonthly
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "SCurveMonthly: " & Err.Description
    MsgBox "S-Curve error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_SmartSpread(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "SmartSpread started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_Distribute.SmartSpread
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "SmartSpread: " & Err.Description
    MsgBox "Smart Spread error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

' ===========================================================================
' XER GROUP CALLBACKS
' ===========================================================================

Public Sub btn_LoadSchedule(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "LoadSchedule started"
    On Error GoTo ErrHandler
    LoadXERSchedule
    Exit Sub
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "LoadSchedule: " & Err.Description
    MsgBox "Load Schedule error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_ExportXER(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "ExportXER started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_XERCompiler.ExportXER
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "ExportXER: " & Err.Description
    MsgBox "Export XER error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

' ===========================================================================
' AUDIT GROUP CALLBACKS
' ===========================================================================

Public Sub btn_ScheduleAudit(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "ScheduleAudit started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_ScheduleAudit.RunDCMAaudit
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "ScheduleAudit: " & Err.Description
    MsgBox "DCMA Audit error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_CPM(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "CPMAnalysis started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_CPM.RunCPMAnalysis
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "CPMAnalysis: " & Err.Description
    MsgBox "CPM Analysis error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_LookAhead2W(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "LookAhead2W started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_Reporting.CreateLookAhead2W
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "LookAhead2W: " & Err.Description
    MsgBox "Look-Ahead error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_LookAhead4W(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "LookAhead4W started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_Reporting.CreateLookAhead4W
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "LookAhead4W: " & Err.Description
    MsgBox "Look-Ahead error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

' ===========================================================================
' TOOLS GROUP CALLBACKS
' ===========================================================================

Public Sub btn_CopyToLast(control As IRibbonControl)
    On Error GoTo ErrHandler
    CopyToLastRow
    Exit Sub
ErrHandler:
    MsgBox "Copy To Last error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_FillDown(control As IRibbonControl)
    On Error GoTo ErrHandler
    If Not Selection Is Nothing Then Selection.FillDown
    Exit Sub
ErrHandler:
    MsgBox "Fill Down error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_TextCase(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim caseStr As String
    caseStr = InputBox("Enter case type:" & vbCrLf & "0 = UPPER" & vbCrLf & "1 = lower" & vbCrLf & "2 = Proper", _
                       "Text Case", "2")
    If caseStr = "" Then Exit Sub
    Dim caseType As Integer
    caseType = CInt(caseStr)
    If Not Selection Is Nothing Then DEE_Utils.ConvertCase Selection, caseType
    Exit Sub
ErrHandler:
    MsgBox "Text Case error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_DataQuality(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "DataQuality started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_DataQuality.RunDataQualityChecks
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "DataQuality: " & Err.Description
    MsgBox "Data Quality error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

' ===========================================================================
' BASELINE GROUP CALLBACKS
' ===========================================================================

Public Sub btn_LoadBaseline(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "LoadBaseline started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_BaselineCompare.LoadBaseline
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "LoadBaseline: " & Err.Description
    MsgBox "Load Baseline error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_CompareBaseline(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "CompareBaseline started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_BaselineCompare.CompareBaseline
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "CompareBaseline: " & Err.Description
    MsgBox "Compare Baseline error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_ClearBaseline(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "ClearBaseline started"
    On Error GoTo ErrHandler
    DEE_BaselineCompare.ClearBaseline
    Exit Sub
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "ClearBaseline: " & Err.Description
    MsgBox "Clear Baseline error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' VERSION HISTORY CALLBACKS
' ===========================================================================

Public Sub btn_SaveVersion(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "SaveVersion started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_VersionHistory.SaveVersion
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "SaveVersion: " & Err.Description
    MsgBox "Save Version error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_ShowHistory(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_VersionHistory.ShowHistory
    Exit Sub
ErrHandler:
    MsgBox "Show History error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' AUDIT TRAIL CALLBACKS
' ===========================================================================

Public Sub btn_CompareSnapshots(control As IRibbonControl)
    DEE_Logger.LogInfo "DEE_Main", "CompareSnapshots started"
    DEE_Perf.EnterHeavyMode
    On Error GoTo ErrHandler
    DEE_AuditTrail.CompareSnapshots
    GoTo CleanUp
ErrHandler:
    DEE_Logger.LogError "DEE_Main", "CompareSnapshots: " & Err.Description
    MsgBox "Compare Snapshots error: " & Err.Description, vbCritical, "Protocol DEE"
CleanUp:
    DEE_Perf.ExitHeavyMode
End Sub

Public Sub btn_ShowAuditLog(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_AuditTrail.ShowAuditLog
    Exit Sub
ErrHandler:
    MsgBox "Show Audit Log error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' SETTINGS CALLBACK
' ===========================================================================

Public Sub btn_Settings(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_Settings.OpenSettings
    Exit Sub
ErrHandler:
    MsgBox "Settings error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' SMART IMPORT PIPELINE -- LoadXERSchedule
' ===========================================================================
Public Sub LoadXERSchedule()
    ' -----------------------------------------------------------------------
    ' Step 1: File picker -> select .xer
    ' -----------------------------------------------------------------------
    Dim filePath As String
    filePath = DEE_XERParser.GetXERFilePath()
    If filePath = "" Then Exit Sub

    DEE_Perf.EnterHeavyMode
    DEE_Logger.LogInfo "DEE_Main", "LoadXERSchedule: " & filePath

    On Error GoTo ErrHandler

    ' -----------------------------------------------------------------------
    ' Step 2: Parse XER into memory
    ' -----------------------------------------------------------------------
    DEE_Utils.StartProgress "Loading XER file"

    Dim success As Boolean
    success = DEE_XERParser.ParseXERFile(filePath)
    If Not success Then
        DEE_Logger.LogError "DEE_Main", "ParseXERFile failed: " & filePath
        GoTo CleanUp
    End If

    ' -----------------------------------------------------------------------
    ' Step 3: Optional data quality pre-check
    ' -----------------------------------------------------------------------
    Dim runDQ As Integer
    runDQ = MsgBox("Run data quality checks before importing?", _
                   vbYesNo + vbQuestion, "Protocol DEE -- Data Quality")
    If runDQ = vbYes Then
        DEE_DataQuality.RunDataQualityChecks
    End If

    ' -----------------------------------------------------------------------
    ' Step 4: Multi-project detection
    ' -----------------------------------------------------------------------
    Dim selectedProjId As String
    selectedProjId = ""

    Dim projects As Variant
    projects = DEE_XERParser.GetProjectList()

    If Not IsEmpty(projects) And UBound(projects) >= 0 Then
        If UBound(projects) = 0 Then
            selectedProjId = CStr(projects(0)(0))
        Else
            Dim projList As String
            projList = "Multiple projects found. Enter number to select:" & vbCrLf & vbCrLf
            Dim p As Integer
            For p = 0 To UBound(projects)
                projList = projList & (p + 1) & ") " & projects(p)(1) & " (ID: " & projects(p)(0) & ")" & vbCrLf
            Next p

            Dim choice As String
            choice = InputBox(projList, "Select Project", "1")
            If choice = "" Then GoTo CleanUp

            Dim choiceNum As Integer
            On Error Resume Next
            choiceNum = CInt(choice) - 1
            On Error GoTo ErrHandler

            If choiceNum >= 0 And choiceNum <= UBound(projects) Then
                selectedProjId = CStr(projects(choiceNum)(0))
            Else
                MsgBox "Invalid selection.", vbExclamation, "Protocol DEE"
                GoTo CleanUp
            End If
        End If
    End If

    ' -----------------------------------------------------------------------
    ' Step 5: Budget field selection
    ' -----------------------------------------------------------------------
    Dim availableFields As Variant
    availableFields = DEE_XERParser.GetAvailableBudgetFields()

    Dim budgetField As String
    budgetField = "budget_qty"

    If Not IsEmpty(availableFields) And UBound(availableFields) >= 0 Then
        Dim fieldList As String
        fieldList = "Select budget/quantity field:" & vbCrLf & vbCrLf
        fieldList = fieldList & "0) None / Zero" & vbCrLf
        Dim f As Integer
        For f = 0 To UBound(availableFields)
            fieldList = fieldList & (f + 1) & ") " & availableFields(f) & vbCrLf
        Next f

        Dim fieldChoice As String
        fieldChoice = InputBox(fieldList, "Select Budget Field", "1")
        If fieldChoice = "" Then GoTo CleanUp

        Dim fieldChoiceNum As Integer
        On Error Resume Next
        fieldChoiceNum = CInt(fieldChoice)
        On Error GoTo ErrHandler

        If fieldChoiceNum = 0 Then
            budgetField = ""
        ElseIf fieldChoiceNum >= 1 And fieldChoiceNum <= UBound(availableFields) + 1 Then
            budgetField = availableFields(fieldChoiceNum - 1)
        End If
    End If

    ' -----------------------------------------------------------------------
    ' Step 6: Import raw tables to sheets
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 10, "Importing raw tables"

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim tablesToImport() As String
    tablesToImport = Split("TASK,PROJWBS,CALENDAR,TASKPRED,TASKRSRC,PROJECT,ACTVTYPE,ACTVCODE", ",")

    Dim t As Integer
    For t = 0 To UBound(tablesToImport)
        Dim tName As String
        tName = Trim(tablesToImport(t))
        If DEE_XERParser.TableExists(tName) Then
            DEE_Utils.UpdateProgress 10 + (t * 5), "Importing " & tName
            Dim rawSheet As Worksheet
            Set rawSheet = DEE_Utils.GetOrCreateSheet(wb, tName)
            DEE_XERParser.ImportXERToSheet rawSheet, tName
        End If
    Next t

    ' -----------------------------------------------------------------------
    ' Step 7: Auto-compile Schedule sheet
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 60, "Building Schedule hierarchy"
    DEE_ScheduleBuilder.BuildSchedule budgetField, selectedProjId
    DEE_Logger.LogInfo "DEE_Main", "LoadXERSchedule complete"

    GoTo CleanUp

ErrHandler:
    DEE_Logger.LogError "DEE_Main", "LoadXERSchedule: " & Err.Description
    MsgBox "Load Schedule error: " & Err.Description, vbCritical, "Protocol DEE"

CleanUp:
    DEE_Utils.EndProgress
    DEE_Perf.ExitHeavyMode
End Sub

' ===========================================================================
' INLINE UTILITIES
' ===========================================================================

' ---------------------------------------------------------------------------
' CopyToLastRow -- Copy selected cell value to all cells below until empty
' ---------------------------------------------------------------------------
Private Sub CopyToLastRow()
    If Selection Is Nothing Then Exit Sub
    If Selection.Cells.Count = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim startRow As Long
    Dim col As Integer
    startRow = Selection.Row
    col = Selection.Column

    Dim sourceValue As Variant
    sourceValue = ws.Cells(startRow, col).Value

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    If lastRow <= startRow Then
        MsgBox "No rows below to copy to.", vbInformation, "Protocol DEE"
        Exit Sub
    End If

    ws.Range(ws.Cells(startRow, col), ws.Cells(lastRow, col)).Value = sourceValue
End Sub
