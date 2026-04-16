Attribute VB_Name = "DEE_Logger"
Option Explicit

' =============================================================================
' DEE_Logger.bas -- Ring-Buffer Diagnostic Logger (P0.2)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Writes timestamped log entries to a hidden sheet "_DEE_Log".
' Capped at LOG_MAX_ROWS (FIFO -- oldest rows deleted when over cap).
'
' USAGE:
'   DEE_Logger.LogInfo  "DEE_Main",  "LoadSchedule started"
'   DEE_Logger.LogWarn  "DEE_XER",   "Missing TASKPRED table"
'   DEE_Logger.LogError "DEE_Gantt", "Shape create failed: " & Err.Description
' =============================================================================

Private Const LOG_SHEET    As String = "_DEE_Log"
Private Const LOG_MAX_ROWS As Long   = 10000

Public Enum LogLevel
    LOG_INFO  = 0
    LOG_WARN  = 1
    LOG_ERROR = 2
End Enum

' ---------------------------------------------------------------------------
' LogInfo / LogWarn / LogError -- Convenience wrappers
' ---------------------------------------------------------------------------
Public Sub LogInfo(moduleName As String, msg As String)
    WriteEntry LOG_INFO, moduleName, msg
End Sub

Public Sub LogWarn(moduleName As String, msg As String)
    WriteEntry LOG_WARN, moduleName, msg
End Sub

Public Sub LogError(moduleName As String, msg As String)
    WriteEntry LOG_ERROR, moduleName, msg
End Sub

' ---------------------------------------------------------------------------
' WriteEntry -- Append one row; prune oldest if over cap
' ---------------------------------------------------------------------------
Private Sub WriteEntry(level As LogLevel, moduleName As String, msg As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = GetLogSheet()
    If ws Is Nothing Then Exit Sub

    Dim levelStr As String
    Select Case level
        Case LOG_INFO:  levelStr = "INFO"
        Case LOG_WARN:  levelStr = "WARN"
        Case LOG_ERROR: levelStr = "ERROR"
        Case Else:      levelStr = "INFO"
    End Select

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value = levelStr
    ws.Cells(nextRow, 3).Value = moduleName
    ws.Cells(nextRow, 4).Value = Left(msg, 1000)

    ' Colour-code level column
    Select Case level
        Case LOG_WARN:  ws.Cells(nextRow, 2).Font.Color = RGB(180, 120, 0)
        Case LOG_ERROR: ws.Cells(nextRow, 2).Font.Color = RGB(200, 0, 0)
    End Select

    ' FIFO prune
    Dim dataRows As Long
    dataRows = nextRow - 1  ' Row 1 is header
    If dataRows > LOG_MAX_ROWS Then
        Dim excess As Long
        excess = dataRows - LOG_MAX_ROWS
        ws.Rows("2:" & CStr(1 + excess)).Delete Shift:=xlUp
    End If

    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' GetLogSheet -- Return (creating if needed) the hidden _DEE_Log sheet
' ---------------------------------------------------------------------------
Private Function GetLogSheet() As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = LOG_SHEET
        ws.Visible = xlSheetVeryHidden

        With ws.Range("A1:D1")
            .Value = Array("Timestamp", "Level", "Module", "Message")
            .Font.Bold = True
            .Interior.Color = RGB(0, 32, 96)
            .Font.Color = RGB(255, 255, 255)
        End With
        ws.Columns("A").ColumnWidth = 20
        ws.Columns("B").ColumnWidth = 8
        ws.Columns("C").ColumnWidth = 22
        ws.Columns("D").ColumnWidth = 80
    End If

    Set GetLogSheet = ws
End Function

' ---------------------------------------------------------------------------
' ExportLog -- Write _DEE_Log contents to a CSV file chosen by the user
' ---------------------------------------------------------------------------
Public Sub ExportLog()
    Dim ws As Worksheet
    Set ws = GetLogSheet()
    If ws Is Nothing Then
        MsgBox "No log data found.", vbInformation, "Protocol DEE"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Log is empty.", vbInformation, "Protocol DEE"
        Exit Sub
    End If

    Dim savePath As String
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="DEE_Log_" & Format(Now, "yyyymmdd_HHMMSS") & ".csv", _
        FileFilter:="CSV Files (*.csv),*.csv")
    If savePath = "False" Or savePath = "" Then Exit Sub

    Dim fNum As Integer
    fNum = FreeFile
    Open savePath For Output As #fNum
    Print #fNum, "Timestamp,Level,Module,Message"

    Dim r As Long
    For r = 2 To lastRow
        Dim ts As String
        ts = Format(ws.Cells(r, 1).Value, "yyyy-mm-dd hh:mm:ss")
        Dim lvl As String
        lvl = ws.Cells(r, 2).Value
        Dim mod As String
        mod = ws.Cells(r, 3).Value
        Dim msgVal As String
        msgVal = Replace(CStr(ws.Cells(r, 4).Value), """", """""")
        Print #fNum, ts & "," & lvl & "," & mod & ",""" & msgVal & """"
    Next r

    Close #fNum
    MsgBox "Log exported to:" & vbCrLf & savePath, vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' ClearLog -- Erase all log rows (keep header)
' ---------------------------------------------------------------------------
Public Sub ClearLog()
    Dim ws As Worksheet
    Set ws = GetLogSheet()
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & CStr(lastRow)).Delete Shift:=xlUp
End Sub
