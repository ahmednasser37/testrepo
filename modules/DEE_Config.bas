Attribute VB_Name = "DEE_Config"
Option Explicit

' =============================================================================
' DEE_Config.bas -- Centralized Sheet and Column Constants (P0.5)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Single source of truth for all sheet names, column indexes, and add-in
' metadata. Reference these constants throughout all modules instead of
' using string/integer literals.
'
' USAGE:
'   Dim ws As Worksheet
'   Set ws = wb.Worksheets(DEE_Config.SHEET_SCHEDULE)
'   ws.Cells(r, DEE_Config.COL_TASK_CODE).Value = "A1000"
' =============================================================================

' ---------------------------------------------------------------------------
' Add-in identity
' ---------------------------------------------------------------------------
Public Const ADDIN_NAME       As String = "Protocol DEE"
Public Const ADDIN_VERSION    As String = "2.0"
Public Const ADDIN_AUTHOR     As String = "NWC Project Controls"

' ---------------------------------------------------------------------------
' Visible sheet names
' ---------------------------------------------------------------------------
Public Const SHEET_SCHEDULE   As String = "Schedule"
Public Const SHEET_SETTINGS   As String = "DEE Settings"
Public Const SHEET_GANTT      As String = "Gantt"
Public Const SHEET_PMS        As String = "PMS"
Public Const SHEET_DASHBOARD  As String = "Dashboard"
Public Const SHEET_SCURVE     As String = "S-Curve"
Public Const SHEET_AUDIT      As String = "DCMA Audit"
Public Const SHEET_CPM        As String = "CPM"
Public Const SHEET_DQ         As String = "Data Quality"
Public Const SHEET_LOOKAHEAD  As String = "Look-Ahead"
Public Const SHEET_BASELINE   As String = "Baseline_Variance"

' ---------------------------------------------------------------------------
' Hidden / system sheet names (prefix _ so they sort together)
' ---------------------------------------------------------------------------
Public Const SHEET_LOG        As String = "_DEE_Log"
Public Const SHEET_VERSIONS   As String = "_DEE_Versions"
Public Const SHEET_AUDIT_LOG  As String = "_DEE_Audit"

' Prefixes used with dynamic names
Public Const PREFIX_BASELINE  As String = "_DEE_Baseline_"   ' + name
Public Const PREFIX_VERSION   As String = "_DEE_Version_"    ' + 001, 002...
Public Const PREFIX_SNAP      As String = "Snap_"            ' + yyyy-mm-dd

' ---------------------------------------------------------------------------
' Schedule sheet column indexes (1-based)
' ---------------------------------------------------------------------------
Public Const COL_WBS_ID       As Integer = 1   ' A  WBS / task_code
Public Const COL_TASK_NAME    As Integer = 2   ' B  task_name (empty = WBS parent)
Public Const COL_START        As Integer = 3   ' C  target_start_date
Public Const COL_FINISH       As Integer = 4   ' D  target_end_date
Public Const COL_DUR          As Integer = 5   ' E  target_drtn_hr_cnt (days)
Public Const COL_PCT          As Integer = 6   ' F  phys_complete_pct
Public Const COL_BUDGET       As Integer = 7   ' G  budget field
Public Const COL_ACTUAL       As Integer = 8   ' H  actual_cost / actual_qty
Public Const COL_REMAIN       As Integer = 9   ' I  remain_drtn_hr_cnt (days)
Public Const COL_CALENDAR     As Integer = 10  ' J  clndr_id
Public Const COL_RESOURCE     As Integer = 11  ' K  rsrc_id (first resource)
Public Const COL_CONSTRAINT   As Integer = 12  ' L  cstr_type
Public Const COL_FLOAT        As Integer = 13  ' M  float_path (from CPM)
Public Const COL_PRED         As Integer = 14  ' N  predecessor list
Public Const COL_NOTES        As Integer = 15  ' O  task_memo

' CPM output columns (written by DEE_CPM)
Public Const COL_CPM_ES       As Integer = 25  ' Y
Public Const COL_CPM_EF       As Integer = 26  ' Z
Public Const COL_CPM_LS       As Integer = 27  ' AA
Public Const COL_CPM_LF       As Integer = 28  ' AB
Public Const COL_CPM_TF       As Integer = 29  ' AC
Public Const COL_CPM_FF       As Integer = 30  ' AD
Public Const COL_CPM_CRIT     As Integer = 31  ' AE

' ---------------------------------------------------------------------------
' PMS sheet -- period columns start at column G (index 7)
' ---------------------------------------------------------------------------
Public Const PMS_FIRST_PERIOD_COL As Integer = 7   ' G

' ---------------------------------------------------------------------------
' NWC calendar defaults
' ---------------------------------------------------------------------------
Public Const NWC_WEEKEND_CODE     As Integer = 17  ' NETWORKDAYS.INTL: Fri only
Public Const NWC_WORK_HOURS_DAY   As Integer = 8
Public Const NWC_CURRENCY         As String  = "SAR"

' ---------------------------------------------------------------------------
' Performance / capacity limits
' ---------------------------------------------------------------------------
Public Const LOG_MAX_ROWS         As Long = 10000
Public Const MAX_VERSIONS         As Integer = 20
Public Const MAX_SCHEDULE_ROWS    As Long = 100000
Public Const PROGRESS_INTERVAL    As Integer = 1000  ' rows between status bar updates

' ---------------------------------------------------------------------------
' Traffic-light colours (shared by Audit, CPM, DataQuality)
' ---------------------------------------------------------------------------
Public Function COLOR_PASS() As Long:  COLOR_PASS  = RGB(0,   176,  80):  End Function
Public Function COLOR_WARN() As Long:  COLOR_WARN  = RGB(255, 192,   0):  End Function
Public Function COLOR_FAIL() As Long:  COLOR_FAIL  = RGB(255,   0,   0):  End Function
Public Function COLOR_INFO() As Long:  COLOR_INFO  = RGB(173, 216, 230):  End Function
Public Function COLOR_NONE() As Long:  COLOR_NONE  = RGB(242, 242, 242):  End Function

' ---------------------------------------------------------------------------
' WBS level colours (canonical palette)
' ---------------------------------------------------------------------------
Public Function WBS_COLOR(level As Integer) As Long
    Select Case level
        Case 1:  WBS_COLOR = RGB(0,   32,  96)   ' Dark Navy
        Case 2:  WBS_COLOR = RGB(0,   70, 127)   ' Dark Blue
        Case 3:  WBS_COLOR = RGB(0,  112, 192)   ' Medium Blue
        Case 4:  WBS_COLOR = RGB(31, 117, 254)   ' Bright Blue
        Case 5:  WBS_COLOR = RGB(68, 114, 196)   ' Steel Blue
        Case 6:  WBS_COLOR = RGB(141, 180, 227)  ' Light Blue
        Case 7:  WBS_COLOR = RGB(189, 215, 238)  ' Very Light Blue
        Case 8:  WBS_COLOR = RGB(191, 191, 191)  ' Gray
        Case Else: WBS_COLOR = RGB(255, 255, 255) ' White (activity row)
    End Select
End Function
