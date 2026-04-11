Attribute VB_Name = "DEE_Gantt"
Option Explicit

' =============================================================================
' DEE_Gantt.bas -- Gantt Chart Drawing
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Draws Excel Shapes (rectangles, diamonds) as Gantt bars over a timeline
' header. Timeline starts after data columns (G+).
'
' NAMING CONVENTION:
'   BL-act-{row}    Primary activity bar
'   BL-mile-{row}   Milestone diamond
'   SEC-act-{row}   Secondary activity bar
'   SEC-mile-{row}  Secondary milestone
'   BASE-act-{row}  Baseline bar (when baseline loaded)
'   DataDateLine    Vertical red data date line
'
' DETECTION: Column B empty = WBS row; Column B non-empty = Activity row
' =============================================================================

' Bar geometry constants
Private Const BAR_HEIGHT As Single = 7      ' Points
Private Const BAR_V_OFFSET As Single = 2    ' Vertical offset from row top
Private Const MILE_SIZE As Single = 6       ' Milestone diamond size
Private Const WBS_BAR_HEIGHT As Single = 5  ' WBS summary bar height

' Color constants
Private Const COLOR_WBS_BAR As Long = RGB(95, 129, 189)
Private Const COLOR_ACT_BAR As Long = RGB(0, 153, 0)
Private Const COLOR_SEC_WBS As Long = RGB(0, 128, 128)
Private Const COLOR_SEC_ACT As Long = RGB(255, 153, 0)
Private Const COLOR_BASELINE As Long = RGB(150, 150, 150)
Private Const COLOR_DATA_DATE As Long = RGB(255, 0, 0)
Private Const COLOR_MILESTONE As Long = RGB(255, 0, 0)

' Timeline mode
Private Const MODE_WEEKLY As Integer = 1
Private Const MODE_MONTHLY As Integer = 2

' ---------------------------------------------------------------------------
' DrawGanttChart -- Main entry point
' ---------------------------------------------------------------------------
Public Sub DrawGanttChart(ws As Worksheet)
    DEE_Utils.StartProgress "Drawing Gantt Chart"

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        MsgBox "No data found in sheet.", vbExclamation, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    ' Ask user for timeline mode
    Dim modeStr As String
    modeStr = InputBox("Enter timeline mode:" & vbCrLf & "1 = Weekly" & vbCrLf & "2 = Monthly", _
                       "Gantt Timeline Mode", "2")
    If modeStr = "" Then DEE_Utils.EndProgress: Exit Sub

    Dim timelineMode As Integer
    timelineMode = CInt(modeStr)
    If timelineMode <> MODE_WEEKLY And timelineMode <> MODE_MONTHLY Then
        timelineMode = MODE_MONTHLY
    End If

    ' Find project date range
    Dim projStart As Date
    Dim projEnd As Date
    FindProjectDateRange ws, lastRow, projStart, projEnd

    If projStart = #1/1/1900# Or projEnd = #1/1/1900# Then
        MsgBox "No valid dates found in schedule.", vbExclamation, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Clear existing Gantt shapes and timeline columns
    ClearGanttShapes ws
    ClearTimelineColumns ws

    ' Build timeline
    Dim timelineStartCol As Integer
    Dim periods() As Date        ' Period start dates
    Dim periodCount As Integer
    Dim periodColWidth As Single

    timelineStartCol = 7  ' Column G (after A-F data columns)

    BuildTimeline ws, projStart, projEnd, timelineMode, timelineStartCol, _
                  periods, periodCount, periodColWidth

    ' Draw bars for each row
    Dim i As Long
    For i = 2 To lastRow
        DEE_Utils.UpdateProgress CLng(((i - 2) / (lastRow - 1)) * 100), "Drawing row " & i

        Dim startDate As Variant
        Dim endDate As Variant
        startDate = ws.Cells(i, 3).Value
        endDate = ws.Cells(i, 4).Value

        If Not IsDate(startDate) Then GoTo NextGanttRow

        Dim isWBS As Boolean
        isWBS = DEE_Utils.IsWBSRow(ws, i)

        ' Map dates to period indexes
        Dim startPeriod As Integer
        Dim endPeriod As Integer
        startPeriod = DateToPeriodIndex(CDate(startDate), periods, periodCount)
        endPeriod = DateToPeriodIndex(CDate(IIf(IsDate(endDate), endDate, startDate)), _
                                      periods, periodCount)

        If startPeriod < 0 Or startPeriod >= periodCount Then GoTo NextGanttRow
        If endPeriod < 0 Then endPeriod = startPeriod
        If endPeriod >= periodCount Then endPeriod = periodCount - 1

        ' Check for milestone (start = end date)
        Dim isMilestone As Boolean
        isMilestone = Not IsDate(endDate) Or (IsDate(endDate) And _
                      CDate(startDate) = CDate(endDate))

        If isMilestone Then
            DrawMilestone ws, i, timelineStartCol, startPeriod, periodColWidth, isWBS
        Else
            DrawBar ws, i, timelineStartCol, startPeriod, endPeriod, periodColWidth, isWBS, _
                    False, "", ""
        End If

NextGanttRow:
    Next i

    ' Draw data date line
    Dim dataDate As Date
    dataDate = DEE_Utils.GetDataDate()
    Dim ddPeriod As Integer
    ddPeriod = DateToPeriodIndex(dataDate, periods, periodCount)
    If ddPeriod >= 0 And ddPeriod < periodCount Then
        DrawDataDateLine ws, timelineStartCol, ddPeriod, periodColWidth, lastRow, dataDate
    End If

    Application.ScreenUpdating = True
    DEE_Utils.EndProgress
    MsgBox "Gantt chart drawn successfully!", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' FindProjectDateRange -- Scan schedule for min/max dates
' ---------------------------------------------------------------------------
Private Sub FindProjectDateRange(ws As Worksheet, lastRow As Long, _
                                  ByRef projStart As Date, ByRef projEnd As Date)
    projStart = #1/1/9999#
    projEnd = #1/1/1900#

    Dim i As Long
    For i = 2 To lastRow
        Dim s As Variant
        Dim e As Variant
        s = ws.Cells(i, 3).Value
        e = ws.Cells(i, 4).Value

        If IsDate(s) Then
            If CDate(s) < projStart Then projStart = CDate(s)
        End If
        If IsDate(e) Then
            If CDate(e) > projEnd Then projEnd = CDate(e)
        End If
    Next i

    If projStart = #1/1/9999# Then projStart = #1/1/1900#
End Sub

' ---------------------------------------------------------------------------
' BuildTimeline -- Create timeline header columns
' ---------------------------------------------------------------------------
Private Sub BuildTimeline(ws As Worksheet, projStart As Date, projEnd As Date, _
                           timelineMode As Integer, timelineStartCol As Integer, _
                           ByRef periods() As Date, ByRef periodCount As Integer, _
                           ByRef periodColWidth As Single)

    Dim periodStart As Date
    Dim periodEnd As Date

    If timelineMode = MODE_WEEKLY Then
        ' Round start to nearest Friday (or Monday -- use Monday for simplicity)
        periodStart = projStart - Weekday(projStart, vbMonday) + 1
        periodColWidth = 15  ' Points per week column
    Else
        ' Monthly: round to 1st of month
        periodStart = DateSerial(Year(projStart), Month(projStart), 1)
        periodColWidth = 40  ' Points per month column
    End If

    ' Count periods
    Dim d As Date
    d = periodStart
    periodCount = 0
    Do While d <= projEnd + 60  ' Extra buffer
        periodCount = periodCount + 1
        If timelineMode = MODE_WEEKLY Then
            d = d + 7
        Else
            d = DateSerial(Year(d), Month(d) + 1, 1)
        End If
    Loop

    ReDim periods(0 To periodCount - 1)

    ' Write timeline headers and set column widths
    d = periodStart
    Dim p As Integer
    For p = 0 To periodCount - 1
        periods(p) = d

        Dim col As Integer
        col = timelineStartCol + p

        ' Header text
        Dim headerText As String
        If timelineMode = MODE_WEEKLY Then
            headerText = Format(d, "dd-mmm")
        Else
            headerText = Format(d, "mmm-yy")
        End If

        ws.Cells(1, col).Value = headerText
        ws.Cells(1, col).Interior.Color = RGB(0, 32, 96)
        ws.Cells(1, col).Font.Color = RGB(255, 255, 255)
        ws.Cells(1, col).Font.Bold = True
        ws.Cells(1, col).Font.Size = 8
        ws.Cells(1, col).HorizontalAlignment = xlCenter

        ' Set column width (points -> characters: approximate)
        ws.Columns(col).ColumnWidth = IIf(timelineMode = MODE_WEEKLY, 5, 8)

        If timelineMode = MODE_WEEKLY Then
            d = d + 7
        Else
            d = DateSerial(Year(d), Month(d) + 1, 1)
        End If
    Next p
End Sub

' ---------------------------------------------------------------------------
' DateToPeriodIndex -- Map a date to a 0-based period index
' Returns -1 if outside timeline
' ---------------------------------------------------------------------------
Private Function DateToPeriodIndex(d As Date, periods() As Date, periodCount As Integer) As Integer
    DateToPeriodIndex = -1
    If periodCount = 0 Then Exit Function

    Dim p As Integer
    For p = 0 To periodCount - 2
        If d >= periods(p) And d < periods(p + 1) Then
            DateToPeriodIndex = p
            Exit Function
        End If
    Next p

    ' Last period
    If d >= periods(periodCount - 1) Then
        DateToPeriodIndex = periodCount - 1
    End If
End Function

' ---------------------------------------------------------------------------
' DrawBar -- Draw a Gantt bar shape
' ---------------------------------------------------------------------------
Private Sub DrawBar(ws As Worksheet, rowNum As Long, timelineStartCol As Integer, _
                     startPeriod As Integer, endPeriod As Integer, periodColWidth As Single, _
                     isWBS As Boolean, isSecondary As Boolean, _
                     Optional barPrefix As String = "BL", _
                     Optional colorOverride As String = "")

    Dim barLeft As Single
    Dim barTop As Single
    Dim barWidth As Single
    Dim barHeight As Single

    barLeft = ws.Cells(rowNum, timelineStartCol + startPeriod).Left + 1
    barTop = ws.Cells(rowNum, 1).Top + BAR_V_OFFSET
    barWidth = ws.Cells(rowNum, timelineStartCol + endPeriod).Left + _
               ws.Cells(rowNum, timelineStartCol + endPeriod).Width - barLeft - 1
    barHeight = IIf(isWBS, WBS_BAR_HEIGHT, BAR_HEIGHT)

    If barWidth <= 0 Then barWidth = 5

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barWidth, barHeight)

    ' Bar color
    Dim barColor As Long
    If colorOverride <> "" Then
        barColor = CLng(colorOverride)
    ElseIf isSecondary Then
        barColor = IIf(isWBS, COLOR_SEC_WBS, COLOR_SEC_ACT)
    Else
        barColor = IIf(isWBS, COLOR_WBS_BAR, COLOR_ACT_BAR)
    End If

    With shp
        .Name = barPrefix & IIf(isWBS, "-wbs-", "-act-") & rowNum
        .Fill.ForeColor.RGB = barColor
        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
        .TextFrame.Characters.Text = ""
    End With
End Sub

' ---------------------------------------------------------------------------
' DrawMilestone -- Draw a diamond milestone shape
' ---------------------------------------------------------------------------
Private Sub DrawMilestone(ws As Worksheet, rowNum As Long, timelineStartCol As Integer, _
                            period As Integer, periodColWidth As Single, isWBS As Boolean)

    Dim shpLeft As Single
    Dim shpTop As Single

    shpLeft = ws.Cells(rowNum, timelineStartCol + period).Left + (periodColWidth / 2) - (MILE_SIZE / 2)
    shpTop = ws.Cells(rowNum, 1).Top + BAR_V_OFFSET

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeDiamond, shpLeft, shpTop, MILE_SIZE, MILE_SIZE)

    With shp
        .Name = "BL-mile-" & rowNum
        .Fill.ForeColor.RGB = COLOR_MILESTONE
        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
    End With
End Sub

' ---------------------------------------------------------------------------
' DrawDataDateLine -- Draw vertical red line at current data date
' ---------------------------------------------------------------------------
Private Sub DrawDataDateLine(ws As Worksheet, timelineStartCol As Integer, _
                              ddPeriod As Integer, periodColWidth As Single, _
                              lastRow As Long, dataDate As Date)

    Dim lineLeft As Single
    Dim lineTop As Single
    Dim lineBottom As Single

    lineLeft = ws.Cells(2, timelineStartCol + ddPeriod).Left + 2
    lineTop = ws.Cells(2, 1).Top
    lineBottom = ws.Cells(lastRow, 1).Top + ws.Cells(lastRow, 1).Height
    Dim lineHeight As Single
    lineHeight = lineBottom - lineTop
    If lineHeight <= 0 Then lineHeight = 100

    ' Remove existing data date line
    On Error Resume Next
    ws.Shapes("DataDateLine").Delete
    ws.Shapes("DataDateLineLabel").Delete
    On Error GoTo 0

    ' Draw line (thin rectangle)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, lineLeft, lineTop, 1.5, lineHeight)
    With shp
        .Name = "DataDateLine"
        .Fill.ForeColor.RGB = COLOR_DATA_DATE
        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
    End With

    ' Label at top
    Dim label As Shape
    Set label = ws.Shapes.AddShape(msoShapeRectangle, lineLeft - 30, lineTop, 65, 14)
    With label
        .Name = "DataDateLineLabel"
        .Fill.ForeColor.RGB = COLOR_DATA_DATE
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "Data Date: " & Format(dataDate, "dd-mmm-yy")
        With .TextFrame.Characters.Font
            .Color = RGB(255, 255, 255)
            .Size = 7
            .Bold = True
        End With
        .Placement = xlFreeFloating
    End With
End Sub

' ---------------------------------------------------------------------------
' ClearGanttShapes -- Remove all Gantt shapes from the sheet
' ---------------------------------------------------------------------------
Private Sub ClearGanttShapes(ws As Worksheet)
    Dim shp As Shape
    Dim toDelete() As String
    Dim count As Integer
    count = 0
    ReDim toDelete(0 To ws.Shapes.Count)

    For Each shp In ws.Shapes
        Dim sName As String
        sName = shp.Name
        If Left(sName, 3) = "BL-" Or Left(sName, 4) = "SEC-" Or _
           Left(sName, 5) = "BASE-" Or sName = "DataDateLine" Or _
           sName = "DataDateLineLabel" Then
            toDelete(count) = sName
            count = count + 1
        End If
    Next shp

    Dim i As Integer
    For i = 0 To count - 1
        On Error Resume Next
        ws.Shapes(toDelete(i)).Delete
        On Error GoTo 0
    Next i
End Sub

' ---------------------------------------------------------------------------
' ClearTimelineColumns -- Remove timeline header columns (G onward)
' ---------------------------------------------------------------------------
Private Sub ClearTimelineColumns(ws As Worksheet)
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol <= 6 Then Exit Sub

    ' Clear timeline header row and column widths from G+
    Dim c As Integer
    For c = 7 To lastCol
        ws.Cells(1, c).ClearContents
        ws.Cells(1, c).Interior.ColorIndex = xlNone
        ws.Cells(1, c).Font.ColorIndex = xlAutomatic
        ws.Columns(c).ColumnWidth = 8.43  ' Default width
    Next c
End Sub

' ---------------------------------------------------------------------------
' ManageBars -- Show/hide bar categories
' ---------------------------------------------------------------------------
Public Sub ManageBars(ws As Worksheet)
    Dim resp As String
    resp = InputBox("Gantt Bar Manager" & vbCrLf & vbCrLf & _
                    "Enter action:" & vbCrLf & _
                    "HIDE WBS  - Hide WBS bars" & vbCrLf & _
                    "SHOW WBS  - Show WBS bars" & vbCrLf & _
                    "HIDE ACT  - Hide activity bars" & vbCrLf & _
                    "SHOW ACT  - Show activity bars" & vbCrLf & _
                    "CLEAR ALL - Remove all bars" & vbCrLf & _
                    "REDRAW    - Redraw all bars", _
                    "Manage Gantt Bars", "REDRAW")

    If resp = "" Then Exit Sub

    Select Case UCase(Trim(resp))
        Case "HIDE WBS":  SetShapeVisibility ws, "BL-wbs-", False
        Case "SHOW WBS":  SetShapeVisibility ws, "BL-wbs-", True
        Case "HIDE ACT":  SetShapeVisibility ws, "BL-act-", False
        Case "SHOW ACT":  SetShapeVisibility ws, "BL-act-", True
        Case "CLEAR ALL": ClearGanttShapes ws
        Case "REDRAW":    DrawGanttChart ws
        Case Else:
            MsgBox "Unrecognized command: " & resp, vbExclamation, "Protocol DEE"
    End Select
End Sub

' ---------------------------------------------------------------------------
' SetShapeVisibility -- Show or hide shapes matching prefix
' ---------------------------------------------------------------------------
Private Sub SetShapeVisibility(ws As Worksheet, prefix As String, visible As Boolean)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, Len(prefix)) = prefix Then
            shp.Visible = IIf(visible, msoTrue, msoFalse)
        End If
    Next shp
End Sub
