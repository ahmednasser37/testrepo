Attribute VB_Name = "DEE_Settings"
Option Explicit

' =============================================================================
' DEE_Settings.bas -- Configurable Rule Engine (P1 rewrite)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Stores all configurable thresholds in a hidden "_DEE_Settings" sheet so
' values persist in the workbook without being visible to end-users by default.
' A user-facing "DEE Settings" visible sheet can be opened from the Ribbon for
' editing; it reads/writes to the hidden sheet.
'
' NWC DEFAULTS:
'   - Weekend = Friday only (NETWORKDAYS.INTL code 17, Sat-Thu working week)
'   - Currency = SAR
'   - DCMA thresholds aligned to PMM v3 compliance levels
'
' USAGE:
'   DEE_Settings.GetSetting("DCMA_MissingPred", 0.05)   -> 0.05 (Double)
'   DEE_Settings.GetDCMAThreshold("MissingPred")         -> 0.05 (Double)
'   DEE_Settings.SetSetting "DCMA_MissingPred", 0.03
' =============================================================================

Private Const HIDDEN_SHEET  As String = "_DEE_Settings"
Private Const VISIBLE_SHEET As String = "DEE Settings"

' ---------------------------------------------------------------------------
' Default values -- NWC / PMM v3 aligned
' ---------------------------------------------------------------------------
Private Const DEFAULT_DCMA_MISSING_PRED    As Double  = 0.05
Private Const DEFAULT_DCMA_MISSING_SUCC    As Double  = 0.05
Private Const DEFAULT_DCMA_FS_PCT          As Double  = 0.9
Private Const DEFAULT_DCMA_LAG_PCT         As Double  = 0.05
Private Const DEFAULT_DCMA_CONSTRAINT_PCT  As Double  = 0.05
Private Const DEFAULT_DCMA_HIGH_FLOAT_DAYS As Integer = 44
Private Const DEFAULT_DCMA_HIGH_FLOAT_PCT  As Double  = 0.05
Private Const DEFAULT_DCMA_HIGH_DUR_DAYS   As Integer = 44
Private Const DEFAULT_DCMA_HIGH_DUR_PCT    As Double  = 0.05
Private Const DEFAULT_DCMA_RESOURCE_PCT    As Double  = 0.9
Private Const DEFAULT_DCMA_CPLI            As Double  = 0.95
Private Const DEFAULT_TF_NEAR_CRITICAL     As Integer = 5
Private Const DEFAULT_WEEKEND_CODE         As Integer = 17   ' NWC: Fri only (Sat-Thu work)
Private Const DEFAULT_CURRENCY             As String  = "SAR"
Private Const DEFAULT_COMPANY              As String  = "National Water Company"
Private Const DEFAULT_MAX_VERSIONS         As Integer = 20

' ---------------------------------------------------------------------------
' GetSetting -- Generic reader by key; returns defaultVal if not found
' ---------------------------------------------------------------------------
Public Function GetSetting(key As String, defaultVal As Variant) As Variant
    On Error GoTo UseDefault

    Dim ws As Worksheet
    Set ws = GetHiddenSheet(False)  ' don't create if missing
    If ws Is Nothing Then GoTo UseDefault

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) = key Then
            Dim v As Variant
            v = ws.Cells(i, 2).Value
            If IsEmpty(v) Or v = "" Then GoTo UseDefault
            GetSetting = v
            Exit Function
        End If
    Next i

UseDefault:
    GetSetting = defaultVal
End Function

' ---------------------------------------------------------------------------
' SetSetting -- Write one key/value pair to hidden sheet
' ---------------------------------------------------------------------------
Public Sub SetSetting(key As String, val As Variant)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = GetHiddenSheet(True)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) = key Then
            ws.Cells(i, 2).Value = val
            Exit Sub
        End If
    Next i

    ' Not found -- append
    Dim newRow As Long
    newRow = lastRow + 1
    If newRow < 2 Then newRow = 2
    ws.Cells(newRow, 1).Value = key
    ws.Cells(newRow, 2).Value = val

    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' GetHiddenSheet -- Return _DEE_Settings worksheet (create if createIfMissing)
' ---------------------------------------------------------------------------
Private Function GetHiddenSheet(createIfMissing As Boolean) As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(HIDDEN_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        If Not createIfMissing Then Exit Function
        Set ws = wb.Worksheets.Add
        ws.Name = HIDDEN_SHEET
        ws.Visible = xlSheetVeryHidden
        InitHiddenSheet ws
    End If

    Set GetHiddenSheet = ws
End Function

' ---------------------------------------------------------------------------
' InitHiddenSheet -- Seed with all defaults
' ---------------------------------------------------------------------------
Private Sub InitHiddenSheet(ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("Key", "Value")
    ws.Range("A1:B1").Font.Bold = True

    Dim r As Integer
    r = 2

    Dim defaults As Variant
    defaults = Array( _
        Array("DCMA_MissingPred",    DEFAULT_DCMA_MISSING_PRED), _
        Array("DCMA_MissingSucc",    DEFAULT_DCMA_MISSING_SUCC), _
        Array("DCMA_FSPct",          DEFAULT_DCMA_FS_PCT), _
        Array("DCMA_LagPct",         DEFAULT_DCMA_LAG_PCT), _
        Array("DCMA_ConstraintPct",  DEFAULT_DCMA_CONSTRAINT_PCT), _
        Array("DCMA_HighFloatDays",  DEFAULT_DCMA_HIGH_FLOAT_DAYS), _
        Array("DCMA_HighFloatPct",   DEFAULT_DCMA_HIGH_FLOAT_PCT), _
        Array("DCMA_HighDurDays",    DEFAULT_DCMA_HIGH_DUR_DAYS), _
        Array("DCMA_HighDurPct",     DEFAULT_DCMA_HIGH_DUR_PCT), _
        Array("DCMA_ResourcePct",    DEFAULT_DCMA_RESOURCE_PCT), _
        Array("DCMA_CPLI",           DEFAULT_DCMA_CPLI), _
        Array("CPM_NearCritDays",    DEFAULT_TF_NEAR_CRITICAL), _
        Array("PMS_WeekendCode",     DEFAULT_WEEKEND_CODE), _
        Array("GEN_Currency",        DEFAULT_CURRENCY), _
        Array("GEN_CompanyName",     DEFAULT_COMPANY), _
        Array("GEN_MaxVersions",     DEFAULT_MAX_VERSIONS), _
        Array("GEN_LogoPath",        "") _
    )

    Dim i As Integer
    For i = 0 To UBound(defaults)
        ws.Cells(r, 1).Value = defaults(i)(0)
        ws.Cells(r, 2).Value = defaults(i)(1)
        r = r + 1
    Next i

    ws.Columns("A:B").AutoFit
End Sub

' ---------------------------------------------------------------------------
' EnsureSettingsExist -- Create hidden sheet with defaults if not present
' ---------------------------------------------------------------------------
Public Sub EnsureSettingsExist()
    GetHiddenSheet True  ' triggers InitHiddenSheet if needed
End Sub

' ---------------------------------------------------------------------------
' OpenSettings -- Show the user-facing visible settings editor
' ---------------------------------------------------------------------------
Public Sub OpenSettings()
    EnsureSettingsExist
    BuildVisibleEditor
    ActiveWorkbook.Worksheets(VISIBLE_SHEET).Activate
End Sub

' ---------------------------------------------------------------------------
' BuildVisibleEditor -- Create/refresh the visible "DEE Settings" sheet
' ---------------------------------------------------------------------------
Private Sub BuildVisibleEditor()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(VISIBLE_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = VISIBLE_SHEET
        ws.Tab.Color = RGB(0, 112, 192)
    End If

    ws.Cells.Clear

    ' Title
    With ws.Range("A1:D1")
        .Merge
        .Value = "Protocol DEE -- Configuration Settings  (NWC / PMM v3)"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .RowHeight = 28
    End With

    ws.Cells(2, 1).Value = "Edit values in Column C. Changes are saved to the hidden settings store when you click Save Settings."
    ws.Cells(2, 1).Font.Italic = True
    ws.Cells(2, 1).Font.Color = RGB(100, 100, 100)

    Dim headerRow As Integer
    headerRow = 3
    ws.Cells(headerRow, 1).Value = "Setting"
    ws.Cells(headerRow, 2).Value = "Description"
    ws.Cells(headerRow, 3).Value = "Value"
    ws.Cells(headerRow, 4).Value = "Key"
    DEE_Utils.ApplyTableHeader ws, headerRow, 1, 4

    Dim r As Integer
    r = headerRow + 1

    r = WriteSection(ws, r, "DCMA 14-POINT AUDIT THRESHOLDS")
    r = WriteRow(ws, r, "Missing Predecessors %",     "Max % of activities with no predecessor",         GetSetting("DCMA_MissingPred",   DEFAULT_DCMA_MISSING_PRED),   "DCMA_MissingPred",   "0.00%")
    r = WriteRow(ws, r, "Missing Successors %",       "Max % of activities with no successor",           GetSetting("DCMA_MissingSucc",   DEFAULT_DCMA_MISSING_SUCC),   "DCMA_MissingSucc",   "0.00%")
    r = WriteRow(ws, r, "FS Relationship % (min)",    "Minimum % of Finish-to-Start relationships",      GetSetting("DCMA_FSPct",         DEFAULT_DCMA_FS_PCT),         "DCMA_FSPct",         "0.00%")
    r = WriteRow(ws, r, "Lags %",                     "Max % of relationships with lag",                 GetSetting("DCMA_LagPct",        DEFAULT_DCMA_LAG_PCT),        "DCMA_LagPct",        "0.00%")
    r = WriteRow(ws, r, "Hard Constraints %",         "Max % of activities with hard constraints",       GetSetting("DCMA_ConstraintPct", DEFAULT_DCMA_CONSTRAINT_PCT), "DCMA_ConstraintPct", "0.00%")
    r = WriteRow(ws, r, "High Float Days",             "Float (days) above which = High Float",          GetSetting("DCMA_HighFloatDays", DEFAULT_DCMA_HIGH_FLOAT_DAYS),"DCMA_HighFloatDays", "0")
    r = WriteRow(ws, r, "High Float %",               "Max % of activities with high float",             GetSetting("DCMA_HighFloatPct",  DEFAULT_DCMA_HIGH_FLOAT_PCT), "DCMA_HighFloatPct",  "0.00%")
    r = WriteRow(ws, r, "High Duration Days",          "Duration (days) above which = High Duration",   GetSetting("DCMA_HighDurDays",   DEFAULT_DCMA_HIGH_DUR_DAYS),  "DCMA_HighDurDays",   "0")
    r = WriteRow(ws, r, "High Duration %",            "Max % of activities with high duration",          GetSetting("DCMA_HighDurPct",    DEFAULT_DCMA_HIGH_DUR_PCT),   "DCMA_HighDurPct",    "0.00%")
    r = WriteRow(ws, r, "Resources Assigned % (min)", "Min % of activities with resources",              GetSetting("DCMA_ResourcePct",   DEFAULT_DCMA_RESOURCE_PCT),   "DCMA_ResourcePct",   "0.00%")
    r = WriteRow(ws, r, "CPLI (min)",                 "Minimum Critical Path Length Index",              GetSetting("DCMA_CPLI",          DEFAULT_DCMA_CPLI),           "DCMA_CPLI",          "0.00")
    r = r + 1

    r = WriteSection(ws, r, "CPM / FLOAT SETTINGS")
    r = WriteRow(ws, r, "Near-Critical Float Days", "TF threshold for Near-Critical (orange border)", GetSetting("CPM_NearCritDays", DEFAULT_TF_NEAR_CRITICAL), "CPM_NearCritDays", "0")
    r = r + 1

    r = WriteSection(ws, r, "PMS / CALENDAR OPTIONS  (NWC: Saturday-Thursday, Friday off)")
    r = WriteRow(ws, r, "Weekend Code", "NETWORKDAYS.INTL code (17=Fri only [NWC], 1=Sat+Sun, 7=Fri+Sat)", GetSetting("PMS_WeekendCode", DEFAULT_WEEKEND_CODE), "PMS_WeekendCode", "0")
    r = r + 1

    r = WriteSection(ws, r, "GENERAL")
    r = WriteRow(ws, r, "Currency",           "Currency code shown in cost headers",                   GetSetting("GEN_Currency",    DEFAULT_CURRENCY),     "GEN_Currency",    "@")
    r = WriteRow(ws, r, "Company Name",       "Shown in report headers (NWC default)",                 GetSetting("GEN_CompanyName", DEFAULT_COMPANY),      "GEN_CompanyName", "@")
    r = WriteRow(ws, r, "Max Version History","Maximum stored version snapshots (FIFO)",               GetSetting("GEN_MaxVersions", DEFAULT_MAX_VERSIONS), "GEN_MaxVersions", "0")
    r = WriteRow(ws, r, "Report Logo Path",   "Full path to logo image file (optional)",               GetSetting("GEN_LogoPath",    ""),                   "GEN_LogoPath",    "@")

    ' Save button row
    r = r + 1
    With ws.Cells(r, 3)
        .Value = "[ Save Settings ]"
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ws.Cells(r, 1).Value = "Click the cell in Column C to save all values to the settings store."
    ws.Cells(r, 1).Font.Italic = True

    ws.Columns("A:D").AutoFit
    ws.Columns("B").ColumnWidth = 52
    ws.Columns("C").ColumnWidth = 22
End Sub

Private Function WriteSection(ws As Worksheet, rowNum As Integer, title As String) As Integer
    With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 4))
        .Merge
        .Value = title
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
        .RowHeight = 18
    End With
    WriteSection = rowNum + 1
End Function

Private Function WriteRow(ws As Worksheet, rowNum As Integer, _
                           label As String, desc As String, _
                           currentVal As Variant, key As String, _
                           numFmt As String) As Integer
    ws.Cells(rowNum, 1).Value = label
    ws.Cells(rowNum, 2).Value = desc
    ws.Cells(rowNum, 3).Value = currentVal
    ws.Cells(rowNum, 4).Value = key

    If numFmt <> "@" And numFmt <> "" Then
        ws.Cells(rowNum, 3).NumberFormat = numFmt
    End If

    ws.Cells(rowNum, 3).Interior.Color = RGB(255, 255, 204)
    ws.Cells(rowNum, 3).Borders.LineStyle = xlContinuous
    ws.Cells(rowNum, 3).Borders.Color = RGB(180, 180, 180)

    If rowNum Mod 2 = 0 Then
        ws.Cells(rowNum, 1).Interior.Color = RGB(242, 242, 242)
        ws.Cells(rowNum, 2).Interior.Color = RGB(242, 242, 242)
        ws.Cells(rowNum, 4).Interior.Color = RGB(242, 242, 242)
    End If

    WriteRow = rowNum + 1
End Function

' ---------------------------------------------------------------------------
' SaveFromEditor -- Read visible sheet values back to hidden store
' ---------------------------------------------------------------------------
Public Sub SaveFromEditor()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim wsVis As Worksheet
    On Error Resume Next
    Set wsVis = wb.Worksheets(VISIBLE_SHEET)
    On Error GoTo 0
    If wsVis Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(wsVis, 4)

    Dim i As Long
    For i = 1 To lastRow
        Dim key As String
        key = Trim(wsVis.Cells(i, 4).Value)
        If key <> "" And key <> "Key" Then
            Dim val As Variant
            val = wsVis.Cells(i, 3).Value
            SetSetting key, val
        End If
    Next i

    MsgBox "Settings saved.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' Typed getters used by other modules
' ---------------------------------------------------------------------------
Public Function GetDCMAThreshold(key As String) As Double
    Select Case key
        Case "MissingPred":    GetDCMAThreshold = CDbl(GetSetting("DCMA_MissingPred",    DEFAULT_DCMA_MISSING_PRED))
        Case "MissingSucc":    GetDCMAThreshold = CDbl(GetSetting("DCMA_MissingSucc",    DEFAULT_DCMA_MISSING_SUCC))
        Case "FSPct":          GetDCMAThreshold = CDbl(GetSetting("DCMA_FSPct",          DEFAULT_DCMA_FS_PCT))
        Case "LagPct":         GetDCMAThreshold = CDbl(GetSetting("DCMA_LagPct",         DEFAULT_DCMA_LAG_PCT))
        Case "ConstraintPct":  GetDCMAThreshold = CDbl(GetSetting("DCMA_ConstraintPct",  DEFAULT_DCMA_CONSTRAINT_PCT))
        Case "HighFloatDays":  GetDCMAThreshold = CDbl(GetSetting("DCMA_HighFloatDays",  DEFAULT_DCMA_HIGH_FLOAT_DAYS))
        Case "HighFloatPct":   GetDCMAThreshold = CDbl(GetSetting("DCMA_HighFloatPct",   DEFAULT_DCMA_HIGH_FLOAT_PCT))
        Case "HighDurDays":    GetDCMAThreshold = CDbl(GetSetting("DCMA_HighDurDays",    DEFAULT_DCMA_HIGH_DUR_DAYS))
        Case "HighDurPct":     GetDCMAThreshold = CDbl(GetSetting("DCMA_HighDurPct",     DEFAULT_DCMA_HIGH_DUR_PCT))
        Case "ResourcePct":    GetDCMAThreshold = CDbl(GetSetting("DCMA_ResourcePct",    DEFAULT_DCMA_RESOURCE_PCT))
        Case "CPLI":           GetDCMAThreshold = CDbl(GetSetting("DCMA_CPLI",           DEFAULT_DCMA_CPLI))
        Case Else:             GetDCMAThreshold = 0
    End Select
End Function

Public Function GetWeekendCode() As Integer
    GetWeekendCode = CInt(GetSetting("PMS_WeekendCode", DEFAULT_WEEKEND_CODE))
End Function

Public Function GetNearCriticalDays() As Integer
    GetNearCriticalDays = CInt(GetSetting("CPM_NearCritDays", DEFAULT_TF_NEAR_CRITICAL))
End Function

Public Function GetCompanyName() As String
    GetCompanyName = CStr(GetSetting("GEN_CompanyName", DEFAULT_COMPANY))
End Function

Public Function GetCurrency() As String
    GetCurrency = CStr(GetSetting("GEN_Currency", DEFAULT_CURRENCY))
End Function

Public Function GetMaxVersions() As Integer
    GetMaxVersions = CInt(GetSetting("GEN_MaxVersions", DEFAULT_MAX_VERSIONS))
End Function
