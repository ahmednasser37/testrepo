Attribute VB_Name = "DEE_VersionHistory"
Option Explicit

' =============================================================================
' DEE_VersionHistory.bas -- Schedule Version History (Section 15.12)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Stores up to 20 compressed Schedule snapshots in hidden sheet "_DEE_History".
' Each save captures: timestamp, version number, data date, user, row count.
' Snapshots are Base64-encoded CSV blobs (compact, no external dependency).
'
' STORAGE: Hidden sheet "_DEE_History"
'   Col A = Version No    Col B = Timestamp    Col C = Data Date
'   Col D = User          Col E = Row Count    Col F = Compressed Data (Base64 CSV)
'
' MAX VERSIONS: 20 (FIFO rotation -- oldest dropped when limit reached)
' =============================================================================

Private Const HISTORY_SHEET  As String = "_DEE_History"
Private Const MAX_VERSIONS   As Integer = 20

' ---------------------------------------------------------------------------
' SaveVersion -- Capture current Schedule sheet as a new version
' ---------------------------------------------------------------------------
Public Sub SaveVersion()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim wsSch As Worksheet
    On Error Resume Next
    Set wsSch = wb.Worksheets("Schedule")
    On Error GoTo 0
    If wsSch Is Nothing Then
        MsgBox "Schedule sheet not found.", vbCritical, "Protocol DEE"
        Exit Sub
    End If

    DEE_Utils.StartProgress "Saving Version"

    Dim wsHist As Worksheet
    Set wsHist = GetOrCreateHistorySheet(wb)

    ' Get next version number
    Dim versionNo As Integer
    Dim lastHistRow As Long
    lastHistRow = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row
    If lastHistRow < 2 Then
        versionNo = 1
    Else
        versionNo = CInt(wsHist.Cells(lastHistRow, 1).Value) + 1
    End If

    ' Enforce FIFO rotation (max 20 versions)
    Do While (lastHistRow - 1) >= MAX_VERSIONS And lastHistRow >= 2
        wsHist.Rows(2).Delete
        lastHistRow = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row
    Loop

    ' Serialize Schedule sheet to CSV blob
    DEE_Utils.UpdateProgress 30, "Serializing schedule data"
    Dim csvBlob As String
    csvBlob = SerializeSchedule(wsSch)

    ' Encode to Base64
    DEE_Utils.UpdateProgress 70, "Encoding snapshot"
    Dim encoded As String
    encoded = Base64Encode(csvBlob)

    ' Get data date
    Dim dataDate As Date
    dataDate = DEE_Utils.GetDataDate()

    ' Write version record
    Dim newRow As Long
    newRow = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row + 1

    wsHist.Cells(newRow, 1).Value = versionNo
    wsHist.Cells(newRow, 2).Value = Now()
    wsHist.Cells(newRow, 2).NumberFormat = "yyyy-mm-dd HH:MM"
    wsHist.Cells(newRow, 3).Value = dataDate
    wsHist.Cells(newRow, 3).NumberFormat = "yyyy-mm-dd"
    wsHist.Cells(newRow, 4).Value = Environ("USERNAME")
    wsHist.Cells(newRow, 5).Value = DEE_Utils.LastRow(wsSch, 1) - 1
    wsHist.Cells(newRow, 6).Value = encoded

    DEE_Utils.EndProgress
    MsgBox "Version " & versionNo & " saved." & vbCrLf & _
           Format(Now(), "yyyy-mm-dd HH:MM") & " | " & _
           (DEE_Utils.LastRow(wsSch, 1) - 1) & " rows", _
           vbInformation, "Protocol DEE -- Version Saved"
End Sub

' ---------------------------------------------------------------------------
' ShowHistory -- Display version list and let user restore
' ---------------------------------------------------------------------------
Public Sub ShowHistory()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim wsHist As Worksheet
    On Error Resume Next
    Set wsHist = wb.Worksheets(HISTORY_SHEET)
    On Error GoTo 0
    If wsHist Is Nothing Then
        MsgBox "No version history found. Save a version first.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No versions stored yet.", vbInformation, "Protocol DEE"
        Exit Sub
    End If

    ' Build version list string
    Dim listStr As String
    listStr = "Stored Versions (newest first):" & vbCrLf & vbCrLf
    Dim r As Long
    For r = lastRow To 2 Step -1
        listStr = listStr & _
            "v" & wsHist.Cells(r, 1).Value & _
            "  |  " & Format(wsHist.Cells(r, 2).Value, "yyyy-mm-dd HH:MM") & _
            "  |  DD: " & Format(wsHist.Cells(r, 3).Value, "yyyy-mm-dd") & _
            "  |  " & wsHist.Cells(r, 5).Value & " rows" & _
            "  |  " & wsHist.Cells(r, 4).Value & vbCrLf
    Next r

    Dim choice As String
    choice = InputBox(listStr & vbCrLf & "Enter version number to restore (or Cancel to exit):", _
                      "Version History -- Protocol DEE")
    If choice = "" Then Exit Sub

    Dim targetVersion As Integer
    On Error Resume Next
    targetVersion = CInt(choice)
    On Error GoTo 0

    ' Find the row for this version
    Dim foundRow As Long
    foundRow = 0
    For r = 2 To lastRow
        If CInt(wsHist.Cells(r, 1).Value) = targetVersion Then
            foundRow = r
            Exit For
        End If
    Next r

    If foundRow = 0 Then
        MsgBox "Version " & targetVersion & " not found.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    ' Confirm restore
    Dim resp As Integer
    resp = MsgBox("Restore version " & targetVersion & "?" & vbCrLf & _
                  "Saved: " & Format(wsHist.Cells(foundRow, 2).Value, "yyyy-mm-dd HH:MM") & vbCrLf & _
                  "Data Date: " & Format(wsHist.Cells(foundRow, 3).Value, "yyyy-mm-dd") & vbCrLf & vbCrLf & _
                  "WARNING: Current Schedule sheet will be overwritten!", _
                  vbYesNo + vbExclamation, "Protocol DEE -- Restore Version")
    If resp <> vbYes Then Exit Sub

    DEE_Utils.StartProgress "Restoring version " & targetVersion

    ' Decode and restore
    Dim encoded As String
    encoded = CStr(wsHist.Cells(foundRow, 6).Value)
    Dim csvBlob As String
    csvBlob = Base64Decode(encoded)

    DEE_Utils.UpdateProgress 50, "Writing Schedule sheet"
    RestoreSchedule wb, csvBlob

    DEE_Utils.EndProgress
    MsgBox "Version " & targetVersion & " restored successfully.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' SerializeSchedule -- Convert Schedule sheet to CSV string
' ---------------------------------------------------------------------------
Private Function SerializeSchedule(ws As Worksheet) As String
    Dim lastRow As Long
    Dim lastCol As Integer
    lastRow = DEE_Utils.LastRow(ws, 1)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 6 Then lastCol = 6
    ' Cap at column X (24) to keep blob manageable
    If lastCol > 24 Then lastCol = 24

    Dim sb As String
    Dim i As Long
    Dim c As Integer

    For i = 1 To lastRow
        Dim lineArr() As String
        ReDim lineArr(0 To lastCol - 1)
        For c = 1 To lastCol
            Dim v As String
            v = CStr(ws.Cells(i, c).Value)
            ' Escape pipes (delimiter)
            v = Replace(v, "|", "~PIPE~")
            lineArr(c - 1) = v
        Next c
        sb = sb & Join(lineArr, "|") & vbLf
    Next i

    SerializeSchedule = sb
End Function

' ---------------------------------------------------------------------------
' RestoreSchedule -- Rebuild Schedule sheet from CSV blob
' ---------------------------------------------------------------------------
Private Sub RestoreSchedule(wb As Workbook, csvBlob As String)
    Dim ws As Worksheet
    Set ws = DEE_Utils.GetOrCreateSheet(wb, "Schedule")
    ws.Cells.Clear

    Dim lines() As String
    lines = Split(csvBlob, vbLf)

    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(Trim(lines(i))) = 0 Then GoTo NextRestoreLine
        Dim cols() As String
        cols = Split(lines(i), "|")
        Dim c As Integer
        For c = 0 To UBound(cols)
            Dim v As String
            v = Replace(cols(c), "~PIPE~", "|")
            ws.Cells(i + 1, c + 1).Value = v
        Next c
NextRestoreLine:
    Next i

    ' Re-apply date formats and WBS colors
    ws.Columns(3).NumberFormat = "yyyy-mm-dd"
    ws.Columns(4).NumberFormat = "yyyy-mm-dd"
    DEE_WBS.ApplyWBSColoring ws
    DEE_Utils.FreezePaneRow1 ws
End Sub

' ---------------------------------------------------------------------------
' GetOrCreateHistorySheet -- Return hidden history sheet
' ---------------------------------------------------------------------------
Private Function GetOrCreateHistorySheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(HISTORY_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = HISTORY_SHEET
        ws.Visible = xlSheetVeryHidden
        ' Header row
        ws.Cells(1, 1).Value = "Version"
        ws.Cells(1, 2).Value = "Timestamp"
        ws.Cells(1, 3).Value = "Data Date"
        ws.Cells(1, 4).Value = "User"
        ws.Cells(1, 5).Value = "Rows"
        ws.Cells(1, 6).Value = "Data"
        DEE_Utils.ApplyTableHeader ws, 1, 1, 6
    End If

    Set GetOrCreateHistorySheet = ws
End Function

' ---------------------------------------------------------------------------
' Base64Encode -- Encode string to Base64
' ---------------------------------------------------------------------------
Private Function Base64Encode(s As String) As String
    Dim bytes() As Byte
    bytes = StrConv(s, vbFromUnicode)
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Dim node As Object
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    Base64Encode = Replace(node.text, vbLf, "")
End Function

' ---------------------------------------------------------------------------
' Base64Decode -- Decode Base64 string back to string
' ---------------------------------------------------------------------------
Private Function Base64Decode(s As String) As String
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Dim node As Object
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.text = s
    Dim bytes() As Byte
    bytes = node.nodeTypedValue
    Base64Decode = StrConv(bytes, vbUnicode)
End Function
