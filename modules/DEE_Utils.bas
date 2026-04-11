Attribute VB_Name = "DEE_Utils"
Option Explicit

' =============================================================================
' DEE_Utils.bas -- Shared Utilities Module
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================

' ---------------------------------------------------------------------------
' Column Letter Conversion
' ---------------------------------------------------------------------------
Public Function ColLetter(col As Integer) As String
    Dim result As String
    Dim c As Integer
    c = col
    result = ""
    Do While c > 0
        Dim remainder As Integer
        remainder = (c - 1) Mod 26
        result = Chr(65 + remainder) & result
        c = (c - 1) \ 26
    Loop
    ColLetter = result
End Function

' ---------------------------------------------------------------------------
' SkipFirstRow Setting (per-workbook via CustomDocumentProperties)
' ---------------------------------------------------------------------------
Public Property Get g_SkipFirstRow() As Boolean
    On Error Resume Next
    g_SkipFirstRow = ActiveWorkbook.CustomDocumentProperties("DEE_SkipFirstRow").Value
    If Err.Number <> 0 Then g_SkipFirstRow = False
    On Error GoTo 0
End Property

Public Property Let g_SkipFirstRow(Value As Boolean)
    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties("DEE_SkipFirstRow").Value = Value
    If Err.Number <> 0 Then
        On Error Resume Next
        ActiveWorkbook.CustomDocumentProperties.Add _
            Name:="DEE_SkipFirstRow", LinkToContent:=False, _
            Type:=msoPropertyTypeBoolean, Value:=Value
    End If
    On Error GoTo 0
End Property

Public Function GetSkipFirstRow() As Boolean
    GetSkipFirstRow = g_SkipFirstRow
End Function

Public Sub SetSkipFirstRow(val As Boolean)
    g_SkipFirstRow = val
End Sub

' ---------------------------------------------------------------------------
' Progress Bar (Status Bar)
' ---------------------------------------------------------------------------
Public Sub StartProgress(title As String)
    Application.StatusBar = title & " - 0%"
    Application.ScreenUpdating = False
End Sub

Public Sub UpdateProgress(pct As Long, Optional text As String = "")
    If Len(text) > 0 Then
        Application.StatusBar = text & " - " & pct & "%"
    Else
        Application.StatusBar = pct & "% complete"
    End If
End Sub

Public Sub EndProgress()
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------------------
' Locale-Safe Date Parser (CRITICAL -- do NOT use CDate)
' ---------------------------------------------------------------------------
' Input: XER date string "yyyy-mm-dd HH:MM" or "yyyy-mm-dd"
' Output: Excel Date value, or Empty on failure
Public Function SafeParseDate(raw As String) As Variant
    On Error GoTo Fail
    If Len(Trim(raw)) >= 10 Then
        Dim yr As Integer, mo As Integer, dy As Integer
        yr = CInt(Left(raw, 4))
        mo = CInt(Mid(raw, 6, 2))
        dy = CInt(Mid(raw, 9, 2))
        If yr > 1900 And yr < 2100 And mo >= 1 And mo <= 12 And dy >= 1 And dy <= 31 Then
            SafeParseDate = DateSerial(yr, mo, dy)
            Exit Function
        End If
    End If
Fail:
    SafeParseDate = Empty
End Function

' ---------------------------------------------------------------------------
' Header Formatting
' ---------------------------------------------------------------------------
Public Sub FormatHeaderRow(ws As Worksheet, bgColor As Long, fontColor As Long, _
                            Optional headerRow As Integer = 1)
    With ws.Rows(headerRow)
        .Interior.Color = bgColor
        .Font.Color = fontColor
        .Font.Bold = True
    End With
    ws.Rows(headerRow).RowHeight = 20
End Sub

' ---------------------------------------------------------------------------
' Get Or Create Worksheet
' ---------------------------------------------------------------------------
Public Function GetOrCreateSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

' ---------------------------------------------------------------------------
' Text Case Conversion
' caseType: 0=UPPER, 1=lower, 2=Proper
' ---------------------------------------------------------------------------
Public Sub ConvertCase(rng As Range, caseType As Integer)
    Dim cell As Range
    For Each cell In rng
        If Not cell.HasFormula Then
            Select Case caseType
                Case 0: cell.Value = UCase(cell.Value)
                Case 1: cell.Value = LCase(cell.Value)
                Case 2: cell.Value = StrConv(cell.Value, vbProperCase)
            End Select
        End If
    Next cell
End Sub

' ---------------------------------------------------------------------------
' Log Warning to "Warnings" Sheet
' ---------------------------------------------------------------------------
Public Sub LogWarning(wb As Workbook, severity As String, message As String, _
                       Optional rowRef As String = "", Optional rawValue As String = "")
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, "Warnings")

    ' Write header if first use
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Severity"
        ws.Cells(1, 3).Value = "Message"
        ws.Cells(1, 4).Value = "Row Ref"
        ws.Cells(1, 5).Value = "Raw Value"
        FormatHeaderRow ws, RGB(0, 32, 96), RGB(255, 255, 255), 1
    End If

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 2).Value = severity
    ws.Cells(nextRow, 3).Value = message
    ws.Cells(nextRow, 4).Value = rowRef
    ws.Cells(nextRow, 5).Value = rawValue

    ' Color-code severity
    Select Case UCase(severity)
        Case "ERROR":   ws.Cells(nextRow, 2).Interior.Color = RGB(255, 0, 0)
        Case "WARNING": ws.Cells(nextRow, 2).Interior.Color = RGB(255, 192, 0)
        Case "INFO":    ws.Cells(nextRow, 2).Interior.Color = RGB(0, 176, 240)
    End Select
End Sub

' ---------------------------------------------------------------------------
' Check if a value is a WBS row (canonical rule: Column B empty)
' ---------------------------------------------------------------------------
Public Function IsWBSRow(ws As Worksheet, rowNum As Long) As Boolean
    IsWBSRow = (Len(Trim(ws.Cells(rowNum, 2).Value)) = 0)
End Function

' ---------------------------------------------------------------------------
' Get WBS Level from cell value (leading spaces or dot notation)
' ---------------------------------------------------------------------------
Public Function GetWBSLevel(cellValue As String) As Integer
    Dim raw As String
    raw = cellValue
    ' Leading spaces method
    Dim spaces As Integer
    spaces = Len(raw) - Len(LTrim(raw))
    If spaces > 0 Then
        GetWBSLevel = (spaces \ 2) + 1
        Exit Function
    End If
    ' Dot notation method
    Dim dotCount As Integer
    dotCount = Len(raw) - Len(Replace(raw, ".", ""))
    GetWBSLevel = dotCount + 1
End Function

' ---------------------------------------------------------------------------
' Find last used row in a worksheet
' ---------------------------------------------------------------------------
Public Function LastRow(ws As Worksheet, Optional col As Integer = 1) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ---------------------------------------------------------------------------
' Apply standard freeze panes (row 1)
' ---------------------------------------------------------------------------
Public Sub FreezePaneRow1(ws As Worksheet)
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitRow = 1
        .SplitColumn = 0
        .FreezePanes = True
    End With
End Sub

' ---------------------------------------------------------------------------
' Auto-fit columns in a worksheet
' ---------------------------------------------------------------------------
Public Sub AutoFitColumns(ws As Worksheet)
    ws.Cells.EntireColumn.AutoFit
End Sub

' ---------------------------------------------------------------------------
' Data Date property (stored in named range or cell X2 of Schedule sheet)
' ---------------------------------------------------------------------------
Public Function GetDataDate() As Date
    On Error GoTo DefaultDate
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Schedule")
    If IsDate(ws.Range("X2").Value) Then
        GetDataDate = CDate(ws.Range("X2").Value)
        Exit Function
    End If
DefaultDate:
    GetDataDate = Date
End Function

' ---------------------------------------------------------------------------
' Format a range as a table-style header
' ---------------------------------------------------------------------------
Public Sub ApplyTableHeader(ws As Worksheet, headerRow As Integer, _
                              firstCol As Integer, lastCol As Integer)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(headerRow, lastCol))
    With rng
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub
