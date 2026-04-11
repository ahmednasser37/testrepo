Attribute VB_Name = "DEE_PMS"
Option Explicit

' =============================================================================
' DEE_PMS.bas -- Progress Measurement Sheet (PMS) Builder
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Adds columns G through X to the Schedule sheet with progress measurement
' formulas. Two variants:
'   CreatePMS             -- Standard PMS (time-based planned %)
'   CreatePMSWithEarnedUnits -- Earned Value variant (EV-weighted rollup)
'
' COLUMN LAYOUT (G-X):
'   G  = Budgeted Total Cost (display copy of F)
'   H  = Earned Units (user-fills or EV formula)
'   I  = Original Duration (NETWORKDAYS.INTL)
'   J  = Schedule Duration (elapsed to data date)
'   K  = Planned UNIT (J/I * F)
'   L  = Previous Qty (user-fills; locked after period close)
'   M  = Current Qty  (N - L)
'   N  = Cum. Qty     (user-fills)
'   O  = PL% previous (user-fills)
'   P  = PL% Current  (Q - O)
'   Q  = PL% Cum.     (K / F)
'   R  = AC% previous (user-fills)
'   S  = AC% Current  (T - R)
'   T  = AC% Cum.     (N / F for activities; SUBTOTAL for WBS)
'   U  = VR% previous (user-fills)
'   V  = VR% Current  (W - U)
'   W  = VR% Cum.     (T - Q)
'   X  = Data Date    (X2 only; blank otherwise)
'
' SPECIAL CELL: $X$2 = Data Date (set via Set Data Date button)
' WEEKEND CODE: configurable (default 1 = Sat+Sun off)
' =============================================================================

' PMS column constants
Private Const COL_G As Integer = 7    ' Budgeted Total Cost
Private Const COL_H As Integer = 8    ' Earned Units
Private Const COL_I As Integer = 9    ' Original Duration
Private Const COL_J As Integer = 10   ' Schedule Duration
Private Const COL_K As Integer = 11   ' Planned UNIT
Private Const COL_L As Integer = 12   ' Previous Qty
Private Const COL_M As Integer = 13   ' Current Qty
Private Const COL_N As Integer = 14   ' Cum. Qty
Private Const COL_O As Integer = 15   ' PL% previous
Private Const COL_P As Integer = 16   ' PL% Current
Private Const COL_Q As Integer = 17   ' PL% Cum.
Private Const COL_R As Integer = 18   ' AC% previous
Private Const COL_S As Integer = 19   ' AC% Current
Private Const COL_T As Integer = 20   ' AC% Cum.
Private Const COL_U As Integer = 21   ' VR% previous
Private Const COL_V As Integer = 22   ' VR% Current
Private Const COL_W As Integer = 23   ' VR% Cum.
Private Const COL_X As Integer = 24   ' Data Date (X2 only)

Private Const WEEKEND_CODE As Integer = 1 ' Sat+Sun off (default)

' ---------------------------------------------------------------------------
' CreatePMS -- Standard Progress Measurement Sheet
' ---------------------------------------------------------------------------
Public Sub CreatePMS()
    Dim ws As Worksheet
    Set ws = GetScheduleSheet()
    If ws Is Nothing Then Exit Sub

    BuildPMS ws, False
End Sub

' ---------------------------------------------------------------------------
' CreatePMSWithEarnedUnits -- Earned Value variant
' ---------------------------------------------------------------------------
Public Sub CreatePMSWithEarnedUnits()
    Dim ws As Worksheet
    Set ws = GetScheduleSheet()
    If ws Is Nothing Then Exit Sub

    BuildPMS ws, True
End Sub

' ---------------------------------------------------------------------------
' GetScheduleSheet -- Returns Schedule worksheet or shows error
' ---------------------------------------------------------------------------
Private Function GetScheduleSheet() As Worksheet
    On Error Resume Next
    Set GetScheduleSheet = ActiveWorkbook.Worksheets("Schedule")
    On Error GoTo 0
    If GetScheduleSheet Is Nothing Then
        MsgBox "Schedule sheet not found. Please load a schedule first.", vbCritical, "Protocol DEE"
    End If
End Function

' ---------------------------------------------------------------------------
' BuildPMS -- Core PMS builder
' ---------------------------------------------------------------------------
Private Sub BuildPMS(ws As Worksheet, isEV As Boolean)
    DEE_Utils.StartProgress "Building PMS"

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        MsgBox "Schedule sheet is empty.", vbExclamation, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' -----------------------------------------------------------------------
    ' 1. Write header row (G:X)
    ' -----------------------------------------------------------------------
    WriteHeaders ws, isEV

    ' -----------------------------------------------------------------------
    ' 2. Set Data Date in X2 (today if not already set)
    ' -----------------------------------------------------------------------
    If Not IsDate(ws.Cells(2, COL_X).Value) Then
        ws.Cells(2, COL_X).Value = Date
    End If
    ws.Cells(2, COL_X).NumberFormat = "yyyy-mm-dd"

    ' -----------------------------------------------------------------------
    ' 3. Build WBS child-range map (for SUBTOTAL ranges)
    ' -----------------------------------------------------------------------
    ' For each WBS row, find lastChildRow
    Dim wbsChildEnd() As Long
    ReDim wbsChildEnd(2 To lastRow)

    Dim i As Long
    For i = 2 To lastRow
        If DEE_Utils.IsWBSRow(ws, i) Then
            Dim myLevel As Integer
            myLevel = DEE_Utils.GetWBSLevel(CStr(ws.Cells(i, 1).Value))
            Dim lastChild As Long
            lastChild = i
            Dim k As Long
            For k = i + 1 To lastRow
                If DEE_Utils.IsWBSRow(ws, k) Then
                    If DEE_Utils.GetWBSLevel(CStr(ws.Cells(k, 1).Value)) <= myLevel Then
                        Exit For
                    End If
                End If
                lastChild = k
            Next k
            wbsChildEnd(i) = lastChild
        End If
    Next i

    ' -----------------------------------------------------------------------
    ' 4. Write formulas row by row
    ' -----------------------------------------------------------------------
    For i = 2 To lastRow
        DEE_Utils.UpdateProgress CLng(((i - 2) / (lastRow - 1)) * 90), "PMS row " & i

        If DEE_Utils.IsWBSRow(ws, i) Then
            WriteWBSFormulas ws, i, wbsChildEnd(i), isEV
        Else
            WriteActivityFormulas ws, i, isEV
        End If
    Next i

    ' -----------------------------------------------------------------------
    ' 5. Apply number formats
    ' -----------------------------------------------------------------------
    ApplyNumberFormats ws, lastRow

    ' -----------------------------------------------------------------------
    ' 6. Apply column grouping (collapse previous-period columns)
    ' -----------------------------------------------------------------------
    ApplyColumnGroups ws

    ' -----------------------------------------------------------------------
    ' 7. Extend WBS colors to PMS columns (G:X)
    ' -----------------------------------------------------------------------
    ExtendWBSColors ws, lastRow

    ' -----------------------------------------------------------------------
    ' 8. Auto-fit
    ' -----------------------------------------------------------------------
    ws.Columns("G:X").AutoFit

    Application.Calculation = xlCalculationAutomatic
    DEE_Utils.EndProgress
    Application.ScreenUpdating = True

    MsgBox "PMS created successfully!" & IIf(isEV, " (Earned Value mode)", ""), _
           vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' WriteHeaders -- Write PMS column headers in row 1
' ---------------------------------------------------------------------------
Private Sub WriteHeaders(ws As Worksheet, isEV As Boolean)
    Dim headers(COL_G To COL_X) As String
    headers(COL_G) = IIf(isEV, "Progress %", "Budgeted Total Cost")
    headers(COL_H) = "Earned Units"
    headers(COL_I) = "Original Duration"
    headers(COL_J) = "Schedule Duration"
    headers(COL_K) = "Planned UNIT"
    headers(COL_L) = "Previous Qty"
    headers(COL_M) = "Current Qty"
    headers(COL_N) = "Cum. Qty"
    headers(COL_O) = "PL% previous"
    headers(COL_P) = "PL% Current"
    headers(COL_Q) = "PL% Cum."
    headers(COL_R) = "AC% previous"
    headers(COL_S) = "AC% Current"
    headers(COL_T) = "AC% Cum."
    headers(COL_U) = "VR% previous"
    headers(COL_V) = "VR% Current"
    headers(COL_W) = "VR% Cum."
    headers(COL_X) = "Current Data Date"

    Dim c As Integer
    For c = COL_G To COL_X
        ws.Cells(1, c).Value = headers(c)
    Next c

    ' Format header cells to match existing style
    With ws.Range(ws.Cells(1, COL_G), ws.Cells(1, COL_X))
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ---------------------------------------------------------------------------
' WriteActivityFormulas -- Write formulas for a leaf activity row
' ---------------------------------------------------------------------------
Private Sub WriteActivityFormulas(ws As Worksheet, rowNum As Long, isEV As Boolean)
    Dim r As String
    r = CStr(rowNum)
    Dim wc As String
    wc = CStr(WEEKEND_CODE)

    ' G: Budget display (copy of F)
    If isEV Then
        ws.Cells(rowNum, COL_G).Value = ""  ' EV mode: user enters Progress %
    Else
        ws.Cells(rowNum, COL_G).Formula = "=F" & r
    End If

    ' H: Earned Units
    If isEV Then
        ws.Cells(rowNum, COL_H).Formula = "=G" & r & "*F" & r
    Else
        ws.Cells(rowNum, COL_H).Value = ""  ' User fills
    End If

    ' I: Original Duration (NETWORKDAYS.INTL)
    ws.Cells(rowNum, COL_I).Formula = _
        "=IFERROR(NETWORKDAYS.INTL(C" & r & ",D" & r & "," & wc & "),0)"

    ' J: Schedule Duration (elapsed to data date X2)
    ws.Cells(rowNum, COL_J).Formula = _
        "=IF($X$2<C" & r & ",0," & _
        "IF(AND($X$2>=C" & r & ",$X$2<=D" & r & ")," & _
        "NETWORKDAYS.INTL(C" & r & ",$X$2," & wc & ")," & _
        "I" & r & "))"

    ' K: Planned UNIT = (J/I) * F
    ws.Cells(rowNum, COL_K).Formula = _
        "=IFERROR((J" & r & "/I" & r & ")*F" & r & ",0)"

    ' L: Previous Qty (user fills)
    If ws.Cells(rowNum, COL_L).Value = "" Then ws.Cells(rowNum, COL_L).Value = ""

    ' M: Current Qty = N - L
    ws.Cells(rowNum, COL_M).Formula = "=N" & r & "-L" & r

    ' N: Cum. Qty (user fills)
    If ws.Cells(rowNum, COL_N).Value = "" Then ws.Cells(rowNum, COL_N).Value = ""

    ' O: PL% previous (user fills)
    If ws.Cells(rowNum, COL_O).Value = "" Then ws.Cells(rowNum, COL_O).Value = ""

    ' P: PL% Current = Q - O
    ws.Cells(rowNum, COL_P).Formula = "=Q" & r & "-O" & r

    ' Q: PL% Cum. = K/F (time progress)
    ws.Cells(rowNum, COL_Q).Formula = "=IFERROR(K" & r & "/F" & r & ",0)"

    ' R: AC% previous (user fills)
    If ws.Cells(rowNum, COL_R).Value = "" Then ws.Cells(rowNum, COL_R).Value = ""

    ' S: AC% Current = T - R
    ws.Cells(rowNum, COL_S).Formula = "=T" & r & "-R" & r

    ' T: AC% Cum. (user fills for activities)
    If ws.Cells(rowNum, COL_T).Value = "" Then ws.Cells(rowNum, COL_T).Value = ""

    ' U: VR% previous (user fills)
    If ws.Cells(rowNum, COL_U).Value = "" Then ws.Cells(rowNum, COL_U).Value = ""

    ' V: VR% Current = W - U
    ws.Cells(rowNum, COL_V).Formula = "=W" & r & "-U" & r

    ' W: VR% Cum. = T - Q
    ws.Cells(rowNum, COL_W).Formula = "=T" & r & "-Q" & r

    ' X: blank for activity rows (only X2 has data date)
    ws.Cells(rowNum, COL_X).Value = ""
End Sub

' ---------------------------------------------------------------------------
' WriteWBSFormulas -- Write formulas for a WBS summary row
' CRITICAL: lastChildRow computed per WBS node, NOT sheet last row
' ---------------------------------------------------------------------------
Private Sub WriteWBSFormulas(ws As Worksheet, rowNum As Long, lastChild As Long, isEV As Boolean)
    Dim r As String
    r = CStr(rowNum)

    Dim firstChild As String
    firstChild = CStr(rowNum + 1)
    Dim lastChildStr As String
    lastChildStr = CStr(lastChild)

    ' G: Budget display (copy of F)
    If isEV Then
        ws.Cells(rowNum, COL_G).Value = ""
    Else
        ws.Cells(rowNum, COL_G).Formula = "=F" & r
    End If

    ' H: blank
    ws.Cells(rowNum, COL_H).Value = ""

    ' I, J: zero for WBS
    ws.Cells(rowNum, COL_I).Value = 0
    ws.Cells(rowNum, COL_J).Value = 0

    If lastChild > rowNum Then
        ' K: Planned UNIT rollup (SUBTOTAL of K children)
        ws.Cells(rowNum, COL_K).Formula = _
            "=IFERROR(SUBTOTAL(9,K" & firstChild & ":K" & lastChildStr & "),0)"

        ' L: Previous Qty rollup
        ws.Cells(rowNum, COL_L).Formula = _
            "=IFERROR(SUBTOTAL(9,L" & firstChild & ":L" & lastChildStr & "),0)"

        ' M: Current Qty = N - L
        ws.Cells(rowNum, COL_M).Formula = "=N" & r & "-L" & r

        ' N: Cum. Qty rollup
        ws.Cells(rowNum, COL_N).Formula = _
            "=IFERROR(SUBTOTAL(9,N" & firstChild & ":N" & lastChildStr & "),0)"
    Else
        ws.Cells(rowNum, COL_K).Value = 0
        ws.Cells(rowNum, COL_L).Value = 0
        ws.Cells(rowNum, COL_M).Value = 0
        ws.Cells(rowNum, COL_N).Value = 0
    End If

    ' O: blank
    ws.Cells(rowNum, COL_O).Value = ""

    ' P: PL% Current = Q - O
    ws.Cells(rowNum, COL_P).Formula = "=Q" & r & "-O" & r

    ' Q: PL% Cum. = K/F
    ws.Cells(rowNum, COL_Q).Formula = "=IFERROR(K" & r & "/F" & r & ",0)"

    ' R: blank
    ws.Cells(rowNum, COL_R).Value = ""

    ' S: AC% Current = T - R
    ws.Cells(rowNum, COL_S).Formula = "=T" & r & "-R" & r

    ' T: AC% Cum. = N/F (weighted average for WBS)
    If isEV Then
        ' EV: weighted average = SUM(H children) / SUM(F children)
        If lastChild > rowNum Then
            ws.Cells(rowNum, COL_T).Formula = _
                "=IFERROR(SUM(H" & firstChild & ":H" & lastChildStr & ")" & _
                "/SUM(F" & firstChild & ":F" & lastChildStr & "),0)"
        Else
            ws.Cells(rowNum, COL_T).Value = 0
        End If
    Else
        ' Standard: Cum Qty / Budget
        ws.Cells(rowNum, COL_T).Formula = "=IFERROR(N" & r & "/F" & r & ",0)"
    End If

    ' U: blank
    ws.Cells(rowNum, COL_U).Value = ""

    ' V: VR% Current = W - U
    ws.Cells(rowNum, COL_V).Formula = "=W" & r & "-U" & r

    ' W: VR% Cum. = T - Q
    ws.Cells(rowNum, COL_W).Formula = "=T" & r & "-Q" & r

    ' X: Data date only in row 2
    ws.Cells(rowNum, COL_X).Value = ""
End Sub

' ---------------------------------------------------------------------------
' ApplyNumberFormats -- Set number formats for PMS columns
' ---------------------------------------------------------------------------
Private Sub ApplyNumberFormats(ws As Worksheet, lastRow As Long)
    ' Budget/cost columns (G, H, K)
    ws.Range("G2:G" & lastRow).NumberFormat = """$""#,##0.00"
    ws.Range("H2:H" & lastRow).NumberFormat = """$""#,##0.00"
    ws.Range("K2:K" & lastRow).NumberFormat = """$""#,##0.00"

    ' Duration columns (I, J)
    ws.Range("I2:I" & lastRow).NumberFormat = "#,##0.00"
    ws.Range("J2:J" & lastRow).NumberFormat = "#,##0.00"

    ' Qty columns (L, M, N) -- all must match: M = N - L
    ws.Range("L2:L" & lastRow).NumberFormat = "#,##0.00"
    ws.Range("M2:M" & lastRow).NumberFormat = "#,##0.00"
    ws.Range("N2:N" & lastRow).NumberFormat = "#,##0.00"

    ' Percentage columns (O through W)
    ws.Range("O2:W" & lastRow).NumberFormat = "0.00%"

    ' Data date column (X)
    ws.Range("X2:X" & lastRow).NumberFormat = "yyyy-mm-dd"
End Sub

' ---------------------------------------------------------------------------
' ApplyColumnGroups -- Group columns for hide/show workflow
' ---------------------------------------------------------------------------
Private Sub ApplyColumnGroups(ws As Worksheet)
    ' Group "Previous Period" columns: L, O, R, U
    ws.Columns("L:L").Group
    ws.Columns("O:O").Group
    ws.Columns("R:R").Group
    ws.Columns("U:U").Group

    ' Group Earned Units + Original Duration detail
    ws.Columns("H:I").Group
End Sub

' ---------------------------------------------------------------------------
' ExtendWBSColors -- Extend existing row colors from A-F to G-X
' ---------------------------------------------------------------------------
Private Sub ExtendWBSColors(ws As Worksheet, lastRow As Long)
    Dim i As Long
    For i = 2 To lastRow
        Dim existingColor As Long
        existingColor = ws.Cells(i, 1).Interior.Color
        ws.Range(ws.Cells(i, COL_G), ws.Cells(i, COL_X)).Interior.Color = existingColor
    Next i
End Sub

' ---------------------------------------------------------------------------
' LockPeriod -- Lock previous-period columns as static values
' ---------------------------------------------------------------------------
Public Sub LockPeriod(ws As Worksheet)
    Dim resp As Integer
    resp = MsgBox("Lock current period? This will:" & vbCrLf & _
                  " - Convert Previous Qty/% columns (L,O,R,U) to static values" & vbCrLf & _
                  " - Protect those cells from editing" & vbCrLf & vbCrLf & _
                  "This action cannot be undone without unprotecting the sheet.", _
                  vbYesNo + vbExclamation, "Lock Period -- Protocol DEE")

    If resp <> vbYes Then Exit Sub

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)

    ' Unprotect
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ' Unlock all cells first
    ws.Cells.Locked = False

    ' Copy L, O, R, U -> Paste as Values (freeze formulas)
    Dim lockCols() As Integer
    lockCols = Array(COL_L, COL_O, COL_R, COL_U)

    Dim c As Integer
    For Each c In lockCols
        Dim rng As Range
        Set rng = ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c))
        rng.Copy
        rng.PasteSpecial xlPasteValues
        rng.Locked = True
    Next c

    Application.CutCopyMode = False

    ' Protect sheet (allow selecting unlocked cells only)
    ws.Protect UserInterfaceOnly:=True, AllowSorting:=False, AllowFiltering:=True

    ' Color sheet tab red to indicate locked period
    ws.Tab.Color = RGB(255, 0, 0)

    ' Record lock timestamp
    Dim stampCell As Range
    Set stampCell = ws.Cells(1, COL_X)
    stampCell.Value = "LOCKED: " & Format(Now(), "yyyy-mm-dd HH:MM")
    stampCell.Interior.Color = RGB(255, 0, 0)
    stampCell.Font.Color = RGB(255, 255, 255)

    MsgBox "Period locked successfully on " & Format(Now(), "yyyy-mm-dd"), _
           vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' TakeSnapshot -- Save current PMS state as historical snapshot sheet
' ---------------------------------------------------------------------------
Public Sub TakeSnapshot(ws As Worksheet)
    Dim snapName As String
    snapName = "Snap_" & Format(Date, "yyyy-mm-dd")

    ' Check if snapshot already exists for today
    Dim existingSnap As Worksheet
    On Error Resume Next
    Set existingSnap = ActiveWorkbook.Worksheets(snapName)
    On Error GoTo 0

    If Not existingSnap Is Nothing Then
        Dim resp As Integer
        resp = MsgBox("Snapshot '" & snapName & "' already exists. Overwrite?", _
                      vbYesNo + vbQuestion, "Protocol DEE")
        If resp <> vbYes Then Exit Sub
        Application.DisplayAlerts = False
        existingSnap.Delete
        Application.DisplayAlerts = True
    End If

    ' Copy sheet
    ws.Copy After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    Dim snapWs As Worksheet
    Set snapWs = ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    snapWs.Name = snapName

    ' Convert formulas to values (make it static)
    snapWs.UsedRange.Copy
    snapWs.UsedRange.PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    ' Lock entire snapshot
    snapWs.Cells.Locked = True
    snapWs.Protect UserInterfaceOnly:=True

    ' Tag metadata in hidden row
    snapWs.Rows(1).EntireRow.Insert
    snapWs.Cells(1, 1).Value = "SNAPSHOT"
    snapWs.Cells(1, 2).Value = snapName
    snapWs.Cells(1, 3).Value = "DataDate: " & Format(DEE_Utils.GetDataDate(), "yyyy-mm-dd")
    snapWs.Cells(1, 4).Value = "Created: " & Format(Now(), "yyyy-mm-dd HH:MM")
    snapWs.Rows(1).Hidden = True

    ' Color tab blue (historical)
    snapWs.Tab.Color = RGB(0, 112, 192)

    MsgBox "Snapshot '" & snapName & "' created.", vbInformation, "Protocol DEE"
End Sub
