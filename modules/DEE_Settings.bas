Attribute VB_Name = "DEE_Settings"
Option Explicit

' =============================================================================
' DEE_Settings.bas -- Configurable Rule Engine
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Stores all configurable thresholds and settings in a visible "Settings" sheet.
' Named cells make values easy to find and override without touching VBA code.
'
' SETTINGS SHEET: "DEE Settings"
'   Section 1: DCMA Audit Thresholds
'   Section 2: Gantt Colors
'   Section 3: PMS Options
'   Section 4: General
'
' USAGE:
'   Dim threshold As Double
'   threshold = DEE_Settings.GetDCMAThreshold("MissingPred")  ' Returns 0.05
' =============================================================================

Private Const SETTINGS_SHEET As String = "DEE Settings"

' ---------------------------------------------------------------------------
' Default values (used when Settings sheet doesn't exist or cell is blank)
' ---------------------------------------------------------------------------
Private Const DEFAULT_DCMA_MISSING_PRED    As Double = 0.05
Private Const DEFAULT_DCMA_MISSING_SUCC    As Double = 0.05
Private Const DEFAULT_DCMA_FS_PCT          As Double = 0.9
Private Const DEFAULT_DCMA_LAG_PCT         As Double = 0.05
Private Const DEFAULT_DCMA_CONSTRAINT_PCT  As Double = 0.05
Private Const DEFAULT_DCMA_HIGH_FLOAT_DAYS As Integer = 44
Private Const DEFAULT_DCMA_HIGH_FLOAT_PCT  As Double = 0.05
Private Const DEFAULT_DCMA_HIGH_DUR_DAYS   As Integer = 44
Private Const DEFAULT_DCMA_HIGH_DUR_PCT    As Double = 0.05
Private Const DEFAULT_DCMA_RESOURCE_PCT    As Double = 0.9
Private Const DEFAULT_DCMA_CPLI            As Double = 0.95
Private Const DEFAULT_TF_NEAR_CRITICAL     As Integer = 5
Private Const DEFAULT_WEEKEND_CODE         As Integer = 1

' ---------------------------------------------------------------------------
' EnsureSettingsSheet -- Create the Settings sheet if it doesn't exist
' ---------------------------------------------------------------------------
Public Sub EnsureSettingsSheet()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If Not ws Is Nothing Then Exit Sub  ' Already exists

    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = SETTINGS_SHEET
    ws.Tab.Color = RGB(0, 112, 192)

    BuildSettingsLayout ws
    MsgBox "Settings sheet created. Adjust thresholds as needed.", _
           vbInformation, "Protocol DEE -- Settings"
End Sub

' ---------------------------------------------------------------------------
' BuildSettingsLayout -- Write all setting rows with labels + defaults
' ---------------------------------------------------------------------------
Private Sub BuildSettingsLayout(ws As Worksheet)
    ws.Cells.Clear

    ' Title
    With ws.Range("A1:D1")
        .Merge
        .Value = "Protocol DEE -- Configuration Settings"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .RowHeight = 28
    End With

    ws.Cells(2, 1).Value = "Change values in Column C. Column D shows the named range key."
    ws.Cells(2, 1).Font.Italic = True
    ws.Cells(2, 1).Font.Color = RGB(100, 100, 100)

    ' Column headers
    Dim headerRow As Integer
    headerRow = 3
    ws.Cells(headerRow, 1).Value = "Setting"
    ws.Cells(headerRow, 2).Value = "Description"
    ws.Cells(headerRow, 3).Value = "Value"
    ws.Cells(headerRow, 4).Value = "Key"
    DEE_Utils.ApplyTableHeader ws, headerRow, 1, 4

    Dim r As Integer
    r = headerRow + 1

    ' -----------------------------------------------------------------------
    ' Section 1: DCMA Audit Thresholds
    ' -----------------------------------------------------------------------
    r = WriteSectionHeader(ws, r, "DCMA 14-POINT AUDIT THRESHOLDS")
    r = WriteSettingRow(ws, r, "Missing Predecessors %",   "Max % of activities with no predecessor",   DEFAULT_DCMA_MISSING_PRED,    "DCMA_MissingPred",   "0.00%")
    r = WriteSettingRow(ws, r, "Missing Successors %",     "Max % of activities with no successor",     DEFAULT_DCMA_MISSING_SUCC,    "DCMA_MissingSucc",   "0.00%")
    r = WriteSettingRow(ws, r, "FS Relationship % (min)",  "Minimum % of Finish-to-Start relationships", DEFAULT_DCMA_FS_PCT,          "DCMA_FSPct",         "0.00%")
    r = WriteSettingRow(ws, r, "Lags %",                   "Max % of relationships with lag",           DEFAULT_DCMA_LAG_PCT,          "DCMA_LagPct",        "0.00%")
    r = WriteSettingRow(ws, r, "Hard Constraints %",       "Max % of activities with hard constraints", DEFAULT_DCMA_CONSTRAINT_PCT,   "DCMA_ConstraintPct", "0.00%")
    r = WriteSettingRow(ws, r, "High Float Days",          "Float (days) above which = 'High Float'",  DEFAULT_DCMA_HIGH_FLOAT_DAYS,  "DCMA_HighFloatDays", "0")
    r = WriteSettingRow(ws, r, "High Float %",             "Max % of activities with high float",       DEFAULT_DCMA_HIGH_FLOAT_PCT,   "DCMA_HighFloatPct",  "0.00%")
    r = WriteSettingRow(ws, r, "High Duration Days",       "Duration (days) above which = 'High Dur'", DEFAULT_DCMA_HIGH_DUR_DAYS,    "DCMA_HighDurDays",   "0")
    r = WriteSettingRow(ws, r, "High Duration %",          "Max % of activities with high duration",    DEFAULT_DCMA_HIGH_DUR_PCT,     "DCMA_HighDurPct",    "0.00%")
    r = WriteSettingRow(ws, r, "Resources Assigned % (min)", "Min % of activities with resources",      DEFAULT_DCMA_RESOURCE_PCT,     "DCMA_ResourcePct",   "0.00%")
    r = WriteSettingRow(ws, r, "CPLI (min)",               "Minimum Critical Path Length Index",        DEFAULT_DCMA_CPLI,             "DCMA_CPLI",          "0.00")

    r = r + 1  ' spacer

    ' -----------------------------------------------------------------------
    ' Section 2: CPM / Float Settings
    ' -----------------------------------------------------------------------
    r = WriteSectionHeader(ws, r, "CPM / FLOAT SETTINGS")
    r = WriteSettingRow(ws, r, "Near-Critical Float Days", "TF threshold for 'Near-Critical' (orange border)", DEFAULT_TF_NEAR_CRITICAL, "CPM_NearCritDays", "0")

    r = r + 1

    ' -----------------------------------------------------------------------
    ' Section 3: PMS Options
    ' -----------------------------------------------------------------------
    r = WriteSectionHeader(ws, r, "PMS OPTIONS")
    r = WriteSettingRow(ws, r, "Weekend Code",             "NETWORKDAYS.INTL code (1=Sat+Sun, 7=Fri+Sat, 11=Sun only)", DEFAULT_WEEKEND_CODE, "PMS_WeekendCode", "0")

    r = r + 1

    ' -----------------------------------------------------------------------
    ' Section 4: General
    ' -----------------------------------------------------------------------
    r = WriteSectionHeader(ws, r, "GENERAL")
    r = WriteSettingRow(ws, r, "Max Version History",      "Maximum stored version snapshots (FIFO)", 20, "GEN_MaxVersions", "0")
    r = WriteSettingRow(ws, r, "Company Name",             "Shown in report headers",                 "My Company", "GEN_CompanyName", "@")
    r = WriteSettingRow(ws, r, "Report Logo Path",         "Full path to logo image file (optional)", "", "GEN_LogoPath", "@")

    ws.Columns("A:D").AutoFit
    ws.Columns("B").ColumnWidth = 45
    ws.Columns("C").ColumnWidth = 20
End Sub

' ---------------------------------------------------------------------------
' WriteSectionHeader -- Write a bold section separator row; returns next row
' ---------------------------------------------------------------------------
Private Function WriteSectionHeader(ws As Worksheet, rowNum As Integer, title As String) As Integer
    With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 4))
        .Merge
        .Value = title
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
        .RowHeight = 18
    End With
    WriteSectionHeader = rowNum + 1
End Function

' ---------------------------------------------------------------------------
' WriteSettingRow -- Write one setting; returns next row
' ---------------------------------------------------------------------------
Private Function WriteSettingRow(ws As Worksheet, rowNum As Integer, _
                                  label As String, description As String, _
                                  defaultVal As Variant, namedKey As String, _
                                  numFmt As String) As Integer
    ws.Cells(rowNum, 1).Value = label
    ws.Cells(rowNum, 2).Value = description
    ws.Cells(rowNum, 3).Value = defaultVal
    ws.Cells(rowNum, 4).Value = namedKey

    If numFmt <> "@" And numFmt <> "" Then
        ws.Cells(rowNum, 3).NumberFormat = numFmt
    End If

    ' Highlight value cell as editable
    ws.Cells(rowNum, 3).Interior.Color = RGB(255, 255, 204)
    ws.Cells(rowNum, 3).Borders.LineStyle = xlContinuous
    ws.Cells(rowNum, 3).Borders.Color = RGB(180, 180, 180)

    ' Alternate row shading
    If rowNum Mod 2 = 0 Then
        ws.Cells(rowNum, 1).Interior.Color = RGB(242, 242, 242)
        ws.Cells(rowNum, 2).Interior.Color = RGB(242, 242, 242)
        ws.Cells(rowNum, 4).Interior.Color = RGB(242, 242, 242)
    End If

    WriteSettingRow = rowNum + 1
End Function

' ---------------------------------------------------------------------------
' GetValue -- Generic setting reader by named key
' ---------------------------------------------------------------------------
Public Function GetValue(namedKey As String, defaultVal As Variant) As Variant
    On Error GoTo UseDefault
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(SETTINGS_SHEET)

    ' Search column D for key
    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 4)
    Dim i As Long
    For i = 1 To lastRow
        If Trim(ws.Cells(i, 4).Value) = namedKey Then
            Dim v As Variant
            v = ws.Cells(i, 3).Value
            If IsEmpty(v) Or v = "" Then GoTo UseDefault
            GetValue = v
            Exit Function
        End If
    Next i

UseDefault:
    GetValue = defaultVal
End Function

' ---------------------------------------------------------------------------
' Typed getters for common settings (used by other modules)
' ---------------------------------------------------------------------------

Public Function GetDCMAThreshold(key As String) As Double
    Select Case key
        Case "MissingPred":    GetDCMAThreshold = CDbl(GetValue("DCMA_MissingPred",    DEFAULT_DCMA_MISSING_PRED))
        Case "MissingSucc":    GetDCMAThreshold = CDbl(GetValue("DCMA_MissingSucc",    DEFAULT_DCMA_MISSING_SUCC))
        Case "FSPct":          GetDCMAThreshold = CDbl(GetValue("DCMA_FSPct",          DEFAULT_DCMA_FS_PCT))
        Case "LagPct":         GetDCMAThreshold = CDbl(GetValue("DCMA_LagPct",         DEFAULT_DCMA_LAG_PCT))
        Case "ConstraintPct":  GetDCMAThreshold = CDbl(GetValue("DCMA_ConstraintPct",  DEFAULT_DCMA_CONSTRAINT_PCT))
        Case "HighFloatDays":  GetDCMAThreshold = CDbl(GetValue("DCMA_HighFloatDays",  DEFAULT_DCMA_HIGH_FLOAT_DAYS))
        Case "HighFloatPct":   GetDCMAThreshold = CDbl(GetValue("DCMA_HighFloatPct",   DEFAULT_DCMA_HIGH_FLOAT_PCT))
        Case "HighDurDays":    GetDCMAThreshold = CDbl(GetValue("DCMA_HighDurDays",    DEFAULT_DCMA_HIGH_DUR_DAYS))
        Case "HighDurPct":     GetDCMAThreshold = CDbl(GetValue("DCMA_HighDurPct",     DEFAULT_DCMA_HIGH_DUR_PCT))
        Case "ResourcePct":    GetDCMAThreshold = CDbl(GetValue("DCMA_ResourcePct",    DEFAULT_DCMA_RESOURCE_PCT))
        Case "CPLI":           GetDCMAThreshold = CDbl(GetValue("DCMA_CPLI",           DEFAULT_DCMA_CPLI))
        Case Else:             GetDCMAThreshold = 0
    End Select
End Function

Public Function GetWeekendCode() As Integer
    GetWeekendCode = CInt(GetValue("PMS_WeekendCode", DEFAULT_WEEKEND_CODE))
End Function

Public Function GetNearCriticalDays() As Integer
    GetNearCriticalDays = CInt(GetValue("CPM_NearCritDays", DEFAULT_TF_NEAR_CRITICAL))
End Function

Public Function GetCompanyName() As String
    GetCompanyName = CStr(GetValue("GEN_CompanyName", ""))
End Function

Public Function GetMaxVersions() As Integer
    GetMaxVersions = CInt(GetValue("GEN_MaxVersions", 20))
End Function

' ---------------------------------------------------------------------------
' OpenSettings -- Activate the settings sheet (Ribbon button handler)
' ---------------------------------------------------------------------------
Public Sub OpenSettings()
    EnsureSettingsSheet
    ActiveWorkbook.Worksheets(SETTINGS_SHEET).Activate
End Sub
