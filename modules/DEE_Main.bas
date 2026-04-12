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
    On Error GoTo ErrHandler
    DEE_WBS.ApplyWBSColoring ActiveSheet
    Exit Sub
ErrHandler:
    MsgBox "WBS Color error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_WBSGroup(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_WBS.ApplyWBSGrouping ActiveSheet, DEE_Utils.GetSkipFirstRow()
    Exit Sub
ErrHandler:
    MsgBox "WBS Group error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_WBSAggregator(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_WBS.ApplySumLevels ActiveSheet
    Exit Sub
ErrHandler:
    MsgBox "WBS Aggregator error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_WBSColumns(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_WBS.WBSInColumns ActiveSheet
    Exit Sub
ErrHandler:
    MsgBox "WBS Columns error: " & Err.Description, vbCritical, "Protocol DEE"
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
    level = 0  ' Default: all levels
    Select Case LCase(Trim(text))
        Case "all levels": level = 0
        Case "level 1": level = 1
        Case "level 2": level = 2
        Case "level 3": level = 3
        Case "level 4": level = 4
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
    On Error GoTo ErrHandler
    DEE_Gantt.DrawGanttChart ActiveSheet
    Exit Sub
ErrHandler:
    MsgBox "Draw Gantt error: " & Err.Description, vbCritical, "Protocol DEE"
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
    On Error GoTo ErrHandler
    DEE_PMS.CreatePMS
    Exit Sub
ErrHandler:
    MsgBox "Create PMS error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_CreatePMSEV(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_PMS.CreatePMSWithEarnedUnits
    Exit Sub
ErrHandler:
    MsgBox "Create PMS (EV) error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_LockPeriod(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Schedule")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        MsgBox "Schedule sheet not found.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If
    DEE_PMS.LockPeriod ws
    Exit Sub
ErrHandler:
    MsgBox "Lock Period error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_Snapshot(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Schedule")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        MsgBox "Schedule sheet not found.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If
    DEE_PMS.TakeSnapshot ws
    Exit Sub
ErrHandler:
    MsgBox "Snapshot error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_Dashboard(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_Dashboard.CreateDashboard
    Exit Sub
ErrHandler:
    MsgBox "Dashboard error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_SCurveMonthly(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_SCurve.CreateSCurveMonthly
    Exit Sub
ErrHandler:
    MsgBox "S-Curve error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_SmartSpread(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_Distribute.SmartSpread
    Exit Sub
ErrHandler:
    MsgBox "Smart Spread error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' XER GROUP CALLBACKS
' ===========================================================================

Public Sub btn_LoadSchedule(control As IRibbonControl)
    On Error GoTo ErrHandler
    LoadXERSchedule
    Exit Sub
ErrHandler:
    MsgBox "Load Schedule error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_ExportXER(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_XERCompiler.ExportXER
    Exit Sub
ErrHandler:
    MsgBox "Export XER error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' AUDIT GROUP CALLBACKS
' ===========================================================================

Public Sub btn_ScheduleAudit(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_ScheduleAudit.RunDCMAaudit
    Exit Sub
ErrHandler:
    MsgBox "DCMA Audit error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_CPM(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_CPM.RunCPMAnalysis
    Exit Sub
ErrHandler:
    MsgBox "CPM Analysis error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_LookAhead2W(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_Reporting.CreateLookAhead2W
    Exit Sub
ErrHandler:
    MsgBox "Look-Ahead error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_LookAhead4W(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_Reporting.CreateLookAhead4W
    Exit Sub
ErrHandler:
    MsgBox "Look-Ahead error: " & Err.Description, vbCritical, "Protocol DEE"
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
    On Error GoTo ErrHandler
    DEE_DataQuality.RunDataQualityChecks
    Exit Sub
ErrHandler:
    MsgBox "Data Quality error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' BASELINE GROUP CALLBACKS
' ===========================================================================

Public Sub btn_LoadBaseline(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_BaselineCompare.LoadBaseline
    Exit Sub
ErrHandler:
    MsgBox "Load Baseline error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_CompareBaseline(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_BaselineCompare.CompareBaseline
    Exit Sub
ErrHandler:
    MsgBox "Compare Baseline error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

Public Sub btn_ClearBaseline(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_BaselineCompare.ClearBaseline
    Exit Sub
ErrHandler:
    MsgBox "Clear Baseline error: " & Err.Description, vbCritical, "Protocol DEE"
End Sub

' ===========================================================================
' VERSION HISTORY CALLBACKS
' ===========================================================================

Public Sub btn_SaveVersion(control As IRibbonControl)
    On Error GoTo ErrHandler
    DEE_VersionHistory.SaveVersion
    Exit Sub
ErrHandler:
    MsgBox "Save Version error: " & Err.Description, vbCritical, "Protocol DEE"
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
    On Error GoTo ErrHandler
    DEE_AuditTrail.CompareSnapshots
    Exit Sub
ErrHandler:
    MsgBox "Compare Snapshots error: " & Err.Description, vbCritical, "Protocol DEE"
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
' (Section 9 -- Smart XER Import Pipeline)
' ===========================================================================
Public Sub LoadXERSchedule()
    ' -----------------------------------------------------------------------
    ' Step 1: File picker -> select .xer
    ' -----------------------------------------------------------------------
    Dim filePath As String
    filePath = DEE_XERParser.GetXERFilePath()
    If filePath = "" Then Exit Sub  ' User cancelled

    ' -----------------------------------------------------------------------
    ' Step 2: Parse XER into memory
    ' -----------------------------------------------------------------------
    DEE_Utils.StartProgress "Loading XER file"

    Dim success As Boolean
    success = DEE_XERParser.ParseXERFile(filePath)
    If Not success Then
        DEE_Utils.EndProgress
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' Step 3: Run data quality checks (optional but recommended)
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
            ' Single project -- auto-select
            selectedProjId = CStr(projects(0)(0))
        Else
            ' Multiple projects -- let user pick
            Dim projList As String
            projList = "Multiple projects found. Enter number to select:" & vbCrLf & vbCrLf
            Dim p As Integer
            For p = 0 To UBound(projects)
                projList = projList & (p + 1) & ") " & projects(p)(1) & " (ID: " & projects(p)(0) & ")" & vbCrLf
            Next p

            Dim choice As String
            choice = InputBox(projList, "Select Project", "1")
            If choice = "" Then
                DEE_Utils.EndProgress
                Exit Sub
            End If

            Dim choiceNum As Integer
            On Error Resume Next
            choiceNum = CInt(choice) - 1
            On Error GoTo 0

            If choiceNum >= 0 And choiceNum <= UBound(projects) Then
                selectedProjId = CStr(projects(choiceNum)(0))
            Else
                MsgBox "Invalid selection.", vbExclamation, "Protocol DEE"
                DEE_Utils.EndProgress
                Exit Sub
            End If
        End If
    End If

    ' -----------------------------------------------------------------------
    ' Step 5: Budget field selection
    ' -----------------------------------------------------------------------
    Dim availableFields As Variant
    availableFields = DEE_XERParser.GetAvailableBudgetFields()

    Dim budgetField As String
    budgetField = "budget_qty"  ' Default

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

        If fieldChoice = "" Then
            DEE_Utils.EndProgress
            Exit Sub
        End If

        Dim fieldChoiceNum As Integer
        On Error Resume Next
        fieldChoiceNum = CInt(fieldChoice)
        On Error GoTo 0

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

    ' (BuildSchedule handles steps 8-9: auto-apply colors, freeze, completion msg)
    DEE_Utils.EndProgress
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

    ' Find last non-empty row in this column
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    If lastRow <= startRow Then
        MsgBox "No rows below to copy to.", vbInformation, "Protocol DEE"
        Exit Sub
    End If

    ws.Range(ws.Cells(startRow, col), ws.Cells(lastRow, col)).Value = sourceValue
End Sub
