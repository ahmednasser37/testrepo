Attribute VB_Name = "DEE_WBS"
Option Explicit

' =============================================================================
' DEE_WBS.bas -- WBS Functions
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Provides WBS coloring, grouping, aggregation, and column extraction.
'
' CANONICAL WBS DETECTION RULE (used by ALL tools):
'   If Len(Trim(ws.Cells(row, colB).Value)) = 0 Then -> WBS Parent Row
'   Else -> Activity Row
'
' COLOR PALETTE (Plexxel-extracted):
'   Level 1: Dark Navy RGB(0,32,96)   / White / Bold / 14pt
'   Level 2: Med Blue  RGB(0,112,192) / White / Bold / 12pt
'   Level 3: Teal      RGB(0,176,220) / White / Bold / 11pt
'   Level 4: Green     RGB(146,208,80)/ Black / Bold / 11pt
'   Level 5: Yellow    RGB(255,255,0) / Black / Bold / 10pt
'   Level 6: Orange    RGB(255,192,0) / Black / Bold / 10pt
'   Level 7: Red       RGB(255,0,0)   / White / Bold / 10pt
'   Level 8: Gray      RGB(191,191,191)/Black / Bold / 10pt
'   Activity: White    RGB(255,255,255)/ Black / No Bold / 11pt + thin borders
' =============================================================================

' WBS color constants
Private Const WBS_COL_L1_BG As Long = RGB(0, 32, 96)
Private Const WBS_COL_L2_BG As Long = RGB(0, 112, 192)
Private Const WBS_COL_L3_BG As Long = RGB(0, 176, 220)
Private Const WBS_COL_L4_BG As Long = RGB(146, 208, 80)
Private Const WBS_COL_L5_BG As Long = RGB(255, 255, 0)
Private Const WBS_COL_L6_BG As Long = RGB(255, 192, 0)
Private Const WBS_COL_L7_BG As Long = RGB(255, 0, 0)
Private Const WBS_COL_L8_BG As Long = RGB(191, 191, 191)
Private Const WBS_COL_ACT_BG As Long = RGB(255, 255, 255)

' ---------------------------------------------------------------------------
' ApplyWBSColoring -- Apply hierarchical color palette to Schedule sheet
' ---------------------------------------------------------------------------
Public Sub ApplyWBSColoring(ws As Worksheet)
    DEE_Utils.StartProgress "Applying WBS Colors"

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        DEE_Utils.EndProgress
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim i As Long
    For i = 2 To lastRow
        Dim cellValA As String
        cellValA = CStr(ws.Cells(i, 1).Value)

        If DEE_Utils.IsWBSRow(ws, i) Then
            ' WBS row: apply level-based color
            Dim level As Integer
            level = DEE_Utils.GetWBSLevel(cellValA)
            ApplyWBSRowColor ws, i, level
        Else
            ' Activity row: white background, thin borders
            ApplyActivityRowColor ws, i
        End If

        If i Mod 100 = 0 Then
            DEE_Utils.UpdateProgress CLng((i / lastRow) * 100), "Coloring row " & i
        End If
    Next i

    DEE_Utils.EndProgress
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------------------
' ApplyWBSRowColor -- Apply color for a WBS level to all used columns
' ---------------------------------------------------------------------------
Private Sub ApplyWBSRowColor(ws As Worksheet, rowNum As Long, level As Integer)
    Dim bgColor As Long
    Dim fontColor As Long
    Dim fontSize As Integer
    Dim isBold As Boolean

    isBold = True
    Select Case level
        Case 1
            bgColor = WBS_COL_L1_BG: fontColor = RGB(255, 255, 255): fontSize = 14
        Case 2
            bgColor = WBS_COL_L2_BG: fontColor = RGB(255, 255, 255): fontSize = 12
        Case 3
            bgColor = WBS_COL_L3_BG: fontColor = RGB(255, 255, 255): fontSize = 11
        Case 4
            bgColor = WBS_COL_L4_BG: fontColor = RGB(0, 0, 0): fontSize = 11
        Case 5
            bgColor = WBS_COL_L5_BG: fontColor = RGB(0, 0, 0): fontSize = 10
        Case 6
            bgColor = WBS_COL_L6_BG: fontColor = RGB(0, 0, 0): fontSize = 10
        Case 7
            bgColor = WBS_COL_L7_BG: fontColor = RGB(255, 255, 255): fontSize = 10
        Case Else
            bgColor = WBS_COL_L8_BG: fontColor = RGB(0, 0, 0): fontSize = 10
    End Select

    ' Find last used column
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 6 Then lastCol = 6

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, lastCol))

    With rng
        .Interior.Color = bgColor
        .Font.Color = fontColor
        .Font.Bold = isBold
        .Font.Size = fontSize
        .Borders.LineStyle = xlNone
    End With
End Sub

' ---------------------------------------------------------------------------
' ApplyActivityRowColor -- White background, thin borders
' ---------------------------------------------------------------------------
Private Sub ApplyActivityRowColor(ws As Worksheet, rowNum As Long)
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 6 Then lastCol = 6

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, lastCol))

    With rng
        .Interior.Color = WBS_COL_ACT_BG
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Font.Size = 11
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(200, 200, 200)
        End With
    End With
End Sub

' ---------------------------------------------------------------------------
' ApplyWBSGrouping -- Excel outline grouping for collapsible rows
' ---------------------------------------------------------------------------
Public Sub ApplyWBSGrouping(ws As Worksheet, Optional skipFirstRow As Boolean = False)
    DEE_Utils.StartProgress "Applying WBS Grouping"

    ' Clear existing groups
    ws.Rows.ClearOutline

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 3 Then
        DEE_Utils.EndProgress
        Exit Sub
    End If

    Dim startDataRow As Long
    startDataRow = IIf(skipFirstRow, 3, 2)

    ' Build level array for grouping
    Dim i As Long
    Dim levels() As Integer
    ReDim levels(startDataRow To lastRow)

    For i = startDataRow To lastRow
        Dim cellVal As String
        cellVal = CStr(ws.Cells(i, 1).Value)
        If DEE_Utils.IsWBSRow(ws, i) Then
            levels(i) = DEE_Utils.GetWBSLevel(cellVal)
        Else
            levels(i) = 99 ' Activities get deepest level
        End If
    Next i

    ' Group rows below each WBS node
    Application.ScreenUpdating = False

    For i = startDataRow To lastRow
        If DEE_Utils.IsWBSRow(ws, i) Then
            Dim thisLevel As Integer
            thisLevel = levels(i)
            ' Find the span of rows belonging to this WBS node
            Dim j As Long
            j = i + 1
            Do While j <= lastRow
                If DEE_Utils.IsWBSRow(ws, j) Then
                    If levels(j) <= thisLevel Then Exit Do
                End If
                j = j + 1
            Loop
            ' Group rows i+1 to j-1
            If j - 1 >= i + 1 Then
                ws.Rows(i + 1 & ":" & (j - 1)).Group
            End If
        End If
    Next i

    ' Set outline to show level 2 by default
    On Error Resume Next
    ws.Outline.ShowLevels RowLevels:=2
    On Error GoTo 0

    Application.ScreenUpdating = True
    DEE_Utils.EndProgress
End Sub

' ---------------------------------------------------------------------------
' ApplySumLevels -- Insert SUBTOTAL formulas on WBS rows
' ---------------------------------------------------------------------------
Public Sub ApplySumLevels(ws As Worksheet)
    DEE_Utils.StartProgress "Applying SUBTOTAL Aggregation"

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < 5 Then
        DEE_Utils.EndProgress
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Identify numeric columns (skip A-D: text/date; skip B)
    Dim numericCols() As Boolean
    ReDim numericCols(1 To lastCol)
    Dim c As Integer
    For c = 5 To lastCol  ' Start from E (budget and beyond)
        numericCols(c) = True
    Next c

    ' Process WBS rows from bottom up (so inner SUBTOTALs nest correctly)
    Dim i As Long
    For i = lastRow To 2 Step -1
        If DEE_Utils.IsWBSRow(ws, i) Then
            ' Find last child row (first same-or-higher-level WBS below)
            Dim wbsLevel As Integer
            wbsLevel = DEE_Utils.GetWBSLevel(CStr(ws.Cells(i, 1).Value))

            Dim lastChild As Long
            lastChild = i
            Dim k As Long
            For k = i + 1 To lastRow
                If DEE_Utils.IsWBSRow(ws, k) Then
                    If DEE_Utils.GetWBSLevel(CStr(ws.Cells(k, 1).Value)) <= wbsLevel Then
                        Exit For
                    End If
                End If
                lastChild = k
            Next k

            ' Insert SUBTOTAL formula for each numeric column
            If lastChild > i Then
                For c = 5 To lastCol
                    If numericCols(c) Then
                        Dim colLtr As String
                        colLtr = DEE_Utils.ColLetter(c)
                        ws.Cells(i, c).Formula = _
                            "=IFERROR(SUBTOTAL(9," & colLtr & (i + 1) & ":" & colLtr & lastChild & "),0)"
                    End If
                Next c
            End If
        End If

        If i Mod 50 = 0 Then
            DEE_Utils.UpdateProgress CLng(((lastRow - i) / lastRow) * 100), "Aggregating row " & i
        End If
    Next i

    Application.ScreenUpdating = True
    DEE_Utils.EndProgress
End Sub

' ---------------------------------------------------------------------------
' WBSInColumns -- Extract WBS levels into separate columns
' ---------------------------------------------------------------------------
Public Sub WBSInColumns(ws As Worksheet)
    DEE_Utils.StartProgress "WBS to Columns"

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        DEE_Utils.EndProgress
        Exit Sub
    End If

    ' Find max WBS level
    Dim maxLevel As Integer
    maxLevel = 0
    Dim i As Long
    For i = 2 To lastRow
        If DEE_Utils.IsWBSRow(ws, i) Then
            Dim lvl As Integer
            lvl = DEE_Utils.GetWBSLevel(CStr(ws.Cells(i, 1).Value))
            If lvl > maxLevel Then maxLevel = lvl
        End If
    Next i

    If maxLevel = 0 Then
        MsgBox "No WBS levels detected.", vbExclamation, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    ' Find insertion point (after existing data)
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Add headers
    Dim lvlStartCol As Integer
    lvlStartCol = lastCol + 1
    Dim lv As Integer
    For lv = 1 To maxLevel
        ws.Cells(1, lvlStartCol + lv - 1).Value = "WBS Level " & lv
        ws.Cells(1, lvlStartCol + lv - 1).Interior.Color = RGB(0, 112, 192)
        ws.Cells(1, lvlStartCol + lv - 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(1, lvlStartCol + lv - 1).Font.Bold = True
    Next lv

    ' Track current WBS at each level
    Dim currentWBS() As String
    ReDim currentWBS(1 To maxLevel)

    ' Fill level columns
    For i = 2 To lastRow
        If DEE_Utils.IsWBSRow(ws, i) Then
            lvl = DEE_Utils.GetWBSLevel(CStr(ws.Cells(i, 1).Value))
            Dim wbsCode As String
            wbsCode = Trim(ws.Cells(i, 1).Value)
            currentWBS(lvl) = wbsCode
            ' Clear deeper levels
            Dim dl As Integer
            For dl = lvl + 1 To maxLevel
                currentWBS(dl) = ""
            Next dl
        End If

        ' Write level values for this row
        For lv = 1 To maxLevel
            ws.Cells(i, lvlStartCol + lv - 1).Value = currentWBS(lv)
        Next lv

        If i Mod 100 = 0 Then
            DEE_Utils.UpdateProgress CLng((i / lastRow) * 100)
        End If
    Next i

    ws.Columns(lvlStartCol & ":" & (lvlStartCol + maxLevel - 1)).AutoFit
    DEE_Utils.EndProgress
End Sub

' ---------------------------------------------------------------------------
' FilterByLevel -- Show/hide rows based on WBS level
' ---------------------------------------------------------------------------
Public Sub FilterByLevel(ws As Worksheet, showLevel As Integer)
    DEE_Utils.StartProgress "Filtering WBS Level"
    Application.ScreenUpdating = False

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)

    Dim i As Long
    For i = 2 To lastRow
        If DEE_Utils.IsWBSRow(ws, i) Then
            Dim lvl As Integer
            lvl = DEE_Utils.GetWBSLevel(CStr(ws.Cells(i, 1).Value))
            If showLevel = 0 Then
                ws.Rows(i).Hidden = False  ' Show all
            ElseIf lvl > showLevel Then
                ws.Rows(i).Hidden = True
            Else
                ws.Rows(i).Hidden = False
            End If
        Else
            ' Activities: show only if showLevel = 0 (all) or high level requested
            If showLevel = 0 Then
                ws.Rows(i).Hidden = False
            ElseIf showLevel >= 4 Then
                ws.Rows(i).Hidden = False
            Else
                ws.Rows(i).Hidden = True
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    DEE_Utils.EndProgress
End Sub

' ---------------------------------------------------------------------------
' GetWBSColorForLevel -- Returns background color for a given WBS level
' ---------------------------------------------------------------------------
Public Function GetWBSColorForLevel(level As Integer) As Long
    Select Case level
        Case 1: GetWBSColorForLevel = WBS_COL_L1_BG
        Case 2: GetWBSColorForLevel = WBS_COL_L2_BG
        Case 3: GetWBSColorForLevel = WBS_COL_L3_BG
        Case 4: GetWBSColorForLevel = WBS_COL_L4_BG
        Case 5: GetWBSColorForLevel = WBS_COL_L5_BG
        Case 6: GetWBSColorForLevel = WBS_COL_L6_BG
        Case 7: GetWBSColorForLevel = WBS_COL_L7_BG
        Case Else: GetWBSColorForLevel = WBS_COL_L8_BG
    End Select
End Function
