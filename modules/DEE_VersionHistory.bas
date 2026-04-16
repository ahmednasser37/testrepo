Attribute VB_Name = "DEE_VersionHistory"
Option Explicit

' =============================================================================
' DEE_VersionHistory.bas -- Version History (P1 rewrite)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Each version snapshot is stored as a hidden sheet "_DEE_Version_NNN" (NNN
' zero-padded to 3 digits). A registry sheet "_DEE_Versions" tracks metadata:
'
'   Col A: Version number (integer)
'   Col B: Label (user-supplied or auto "v001")
'   Col C: Timestamp (date/time)
'   Col D: Sheet name
'   Col E: Row count (activities)
'   Col F: Hash fingerprint (MSXML2 Base64-based, first 32 chars)
'
' FIFO: when version count exceeds GetMaxVersions(), oldest sheet+row deleted.
' =============================================================================

Private Const REG_SHEET  As String = "_DEE_Versions"
Private Const VER_PREFIX As String = "_DEE_Version_"

Private Const REG_COL_NUM   As Integer = 1
Private Const REG_COL_LABEL As Integer = 2
Private Const REG_COL_TIME  As Integer = 3
Private Const REG_COL_SHEET As Integer = 4
Private Const REG_COL_ROWS  As Integer = 5
Private Const REG_COL_HASH  As Integer = 6

' ---------------------------------------------------------------------------
' SaveVersion -- Snapshot current Schedule sheet into a hidden version sheet
' ---------------------------------------------------------------------------
Public Sub SaveVersion()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim schedWs As Worksheet
    On Error Resume Next
    Set schedWs = wb.Worksheets(DEE_Config.SHEET_SCHEDULE)
    On Error GoTo 0

    If schedWs Is Nothing Then
        MsgBox "Schedule sheet not found.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    Dim label As String
    label = InputBox("Enter version label (leave blank for auto):", "Save Version", "")
    If label = Chr(0) Then Exit Sub  ' user pressed Cancel

    DEE_Logger.LogInfo "DEE_VersionHistory", "SaveVersion started"
    DEE_Utils.StartProgress "Saving version snapshot"

    Dim reg As Worksheet
    Set reg = GetRegistrySheet(wb)

    Dim nextNum As Integer
    nextNum = NextVersionNumber(reg)
    Dim numStr As String
    numStr = Format(nextNum, "000")

    If Trim(label) = "" Then label = "v" & numStr

    Dim sheetName As String
    sheetName = VER_PREFIX & numStr

    ' Copy Schedule to new hidden sheet
    schedWs.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Dim newWs As Worksheet
    Set newWs = wb.Worksheets(wb.Worksheets.Count)
    newWs.Name = sheetName
    newWs.Visible = xlSheetVeryHidden

    Dim rowCount As Long
    rowCount = DEE_Utils.LastRow(newWs, DEE_Config.COL_WBS_ID) - 1
    If rowCount < 0 Then rowCount = 0

    Dim hashVal As String
    hashVal = ComputeSheetHash(newWs, rowCount)

    ' Append to registry
    Dim lastRegRow As Long
    lastRegRow = DEE_Utils.LastRow(reg, REG_COL_NUM)
    If lastRegRow < 1 Then lastRegRow = 1
    Dim regRow As Long
    regRow = lastRegRow + 1

    reg.Cells(regRow, REG_COL_NUM).Value   = nextNum
    reg.Cells(regRow, REG_COL_LABEL).Value = label
    reg.Cells(regRow, REG_COL_TIME).Value  = Now
    reg.Cells(regRow, REG_COL_TIME).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    reg.Cells(regRow, REG_COL_SHEET).Value = sheetName
    reg.Cells(regRow, REG_COL_ROWS).Value  = rowCount
    reg.Cells(regRow, REG_COL_HASH).Value  = hashVal

    PruneOldVersions wb, reg

    DEE_Utils.EndProgress
    DEE_Logger.LogInfo "DEE_VersionHistory", "Saved version " & label & " (" & rowCount & " rows, hash=" & hashVal & ")"
    MsgBox "Version '" & label & "' saved." & vbCrLf & "Activities: " & rowCount, _
           vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' ShowHistory -- Present version list and let user restore one
' ---------------------------------------------------------------------------
Public Sub ShowHistory()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim reg As Worksheet
    Set reg = GetRegistrySheet(wb)

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(reg, REG_COL_NUM)
    If lastRow < 2 Then
        MsgBox "No version history found. Use 'Save Version' first.", vbInformation, "Protocol DEE"
        Exit Sub
    End If

    Dim listStr As String
    listStr = "Saved Versions (newest last):" & vbCrLf & vbCrLf

    Dim dispIdx As Integer
    dispIdx = 1
    Dim dispMap() As Long
    ReDim dispMap(1 To lastRow - 1)

    Dim r As Long
    For r = 2 To lastRow
        Dim vNum  As Long:   vNum  = CLng(reg.Cells(r, REG_COL_NUM).Value)
        Dim vLbl  As String: vLbl  = CStr(reg.Cells(r, REG_COL_LABEL).Value)
        Dim vTs   As String: vTs   = Format(reg.Cells(r, REG_COL_TIME).Value, "yyyy-mm-dd hh:mm")
        Dim vRows As Long:   vRows = CLng(reg.Cells(r, REG_COL_ROWS).Value)
        Dim vHash As String: vHash = Left(CStr(reg.Cells(r, REG_COL_HASH).Value), 8)
        listStr = listStr & dispIdx & ")  " & vLbl & "  [" & vTs & "]  " & vRows & " rows  #" & vHash & vbCrLf
        dispMap(dispIdx) = r
        dispIdx = dispIdx + 1
    Next r

    listStr = listStr & vbCrLf & "Enter number to RESTORE that version to Schedule (or Cancel):"
    Dim choice As String
    choice = InputBox(listStr, "Version History")
    If choice = "" Then Exit Sub

    Dim choiceNum As Integer
    On Error Resume Next
    choiceNum = CInt(choice)
    On Error GoTo 0
    If choiceNum < 1 Or choiceNum > dispIdx - 1 Then
        MsgBox "Invalid selection.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    Dim targetRow As Long
    targetRow = dispMap(choiceNum)
    Dim targetSheet As String:  targetSheet = CStr(reg.Cells(targetRow, REG_COL_SHEET).Value)
    Dim targetLabel As String:  targetLabel = CStr(reg.Cells(targetRow, REG_COL_LABEL).Value)

    Dim confirm As Integer
    confirm = MsgBox("Restore version '" & targetLabel & "'?" & vbCrLf & _
                     "The current Schedule sheet will be overwritten.", _
                     vbYesNo + vbQuestion, "Protocol DEE")
    If confirm <> vbYes Then Exit Sub

    RestoreVersion wb, targetSheet, targetLabel
End Sub

' ---------------------------------------------------------------------------
' RestoreVersion -- Overwrite Schedule sheet with content from a version sheet
' ---------------------------------------------------------------------------
Private Sub RestoreVersion(wb As Workbook, sheetName As String, label As String)
    Dim verWs As Worksheet
    On Error Resume Next
    Set verWs = wb.Worksheets(sheetName)
    On Error GoTo 0

    If verWs Is Nothing Then
        MsgBox "Version sheet '" & sheetName & "' not found.", vbCritical, "Protocol DEE"
        Exit Sub
    End If

    Dim schedWs As Worksheet
    On Error Resume Next
    Set schedWs = wb.Worksheets(DEE_Config.SHEET_SCHEDULE)
    On Error GoTo 0

    If schedWs Is Nothing Then
        MsgBox "Schedule sheet not found.", vbCritical, "Protocol DEE"
        Exit Sub
    End If

    DEE_Logger.LogInfo "DEE_VersionHistory", "RestoreVersion: " & label & " from " & sheetName
    DEE_Utils.StartProgress "Restoring version: " & label

    schedWs.Cells.Clear
    verWs.UsedRange.Copy schedWs.Range("A1")
    schedWs.Activate

    DEE_Utils.EndProgress
    MsgBox "Version '" & label & "' restored to Schedule sheet.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' GetRegistrySheet -- Return (creating if needed) the _DEE_Versions sheet
' ---------------------------------------------------------------------------
Private Function GetRegistrySheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(REG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = REG_SHEET
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1:F1").Value = Array("Version#", "Label", "Timestamp", "Sheet", "Rows", "Hash")
        ws.Range("A1:F1").Font.Bold = True
        ws.Range("A1:F1").Interior.Color = RGB(0, 32, 96)
        ws.Range("A1:F1").Font.Color = RGB(255, 255, 255)
        ws.Columns("A:F").AutoFit
        ws.Columns("B").ColumnWidth = 20
        ws.Columns("C").ColumnWidth = 22
    End If

    Set GetRegistrySheet = ws
End Function

' ---------------------------------------------------------------------------
' NextVersionNumber -- Returns max version number in registry + 1
' ---------------------------------------------------------------------------
Private Function NextVersionNumber(reg As Worksheet) As Integer
    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(reg, REG_COL_NUM)
    If lastRow < 2 Then
        NextVersionNumber = 1
        Exit Function
    End If

    Dim maxNum As Integer
    maxNum = 0
    Dim r As Long
    For r = 2 To lastRow
        Dim n As Integer
        On Error Resume Next
        n = CInt(reg.Cells(r, REG_COL_NUM).Value)
        On Error GoTo 0
        If n > maxNum Then maxNum = n
    Next r

    NextVersionNumber = maxNum + 1
End Function

' ---------------------------------------------------------------------------
' PruneOldVersions -- FIFO delete oldest versions when over the cap
' ---------------------------------------------------------------------------
Private Sub PruneOldVersions(wb As Workbook, reg As Worksheet)
    Dim maxV As Integer
    maxV = DEE_Settings.GetMaxVersions()
    If maxV < 1 Then maxV = DEE_Config.MAX_VERSIONS

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(reg, REG_COL_NUM)
    Dim currentCount As Long
    currentCount = lastRow - 1

    Do While currentCount > maxV
        Dim oldSheet As String
        oldSheet = CStr(reg.Cells(2, REG_COL_SHEET).Value)

        Dim wsOld As Worksheet
        On Error Resume Next
        Set wsOld = wb.Worksheets(oldSheet)
        On Error GoTo 0
        If Not wsOld Is Nothing Then
            Application.DisplayAlerts = False
            wsOld.Delete
            Application.DisplayAlerts = True
        End If

        reg.Rows(2).Delete Shift:=xlUp
        DEE_Logger.LogInfo "DEE_VersionHistory", "Pruned: " & oldSheet

        lastRow = DEE_Utils.LastRow(reg, REG_COL_NUM)
        currentCount = lastRow - 1
    Loop
End Sub

' ---------------------------------------------------------------------------
' ComputeSheetHash -- Lightweight fingerprint via MSXML2 Base64
' Falls back to timestamp+rowcount string if MSXML2 unavailable
' ---------------------------------------------------------------------------
Private Function ComputeSheetHash(ws As Worksheet, rowCount As Long) As String
    On Error GoTo HashFail

    Dim sample As String
    Dim r As Long
    Dim maxSample As Long
    maxSample = IIf(rowCount > 200, 200, rowCount)

    For r = 2 To 1 + maxSample
        sample = sample & "|" & CStr(ws.Cells(r, 1).Value) & _
                          CStr(ws.Cells(r, 3).Value) & _
                          CStr(ws.Cells(r, 4).Value)
    Next r
    sample = sample & "|n=" & rowCount & "|ts=" & Format(Now, "yyyymmddhhmmss")

    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Dim xmlNode As Object
    Set xmlNode = xmlDoc.createElement("b")
    xmlNode.DataType = "bin.base64"

    Dim enc As Object
    Set enc = CreateObject("System.Text.ASCIIEncoding")
    xmlNode.nodeTypedValue = enc.GetBytes_4(sample)

    ComputeSheetHash = Left(Replace(xmlNode.text, vbCrLf, ""), 32)
    Exit Function

HashFail:
    ComputeSheetHash = Format(Now, "yyyymmddHHMMSS") & Right("000" & rowCount, 4)
End Function
