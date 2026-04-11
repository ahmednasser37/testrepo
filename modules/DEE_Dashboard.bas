Attribute VB_Name = "DEE_Dashboard"
Option Explicit

' =============================================================================
' DEE_Dashboard.bas -- KPI Dashboard Creation
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Creates a KPI dashboard sheet with key project metrics, progress charts,
' and traffic light indicators derived from the Schedule/PMS sheet.
'
' OUTPUT SHEET: "Dashboard"
' KPI CARDS: SPI, CPI, Total Budget, Planned%, Actual%, Variance%
' CHARTS: Progress pie chart, Budget utilization bar chart
' =============================================================================

' KPI card colors
Private Const KPI_GREEN As Long = RGB(0, 176, 80)
Private Const KPI_AMBER As Long = RGB(255, 192, 0)
Private Const KPI_RED As Long = RGB(255, 0, 0)
Private Const KPI_BLUE As Long = RGB(0, 112, 192)
Private Const KPI_DARK As Long = RGB(0, 32, 96)

' ---------------------------------------------------------------------------
' CreateDashboard -- Main entry point
' ---------------------------------------------------------------------------
Public Sub CreateDashboard()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Get Schedule sheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("Schedule")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Schedule sheet not found.", vbCritical, "Protocol DEE"
        Exit Sub
    End If

    DEE_Utils.StartProgress "Building Dashboard"

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        MsgBox "Schedule sheet is empty.", vbExclamation, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' 1. Compute KPIs from Schedule/PMS data
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 10, "Computing KPIs"

    Dim totalBudget As Double
    Dim totalPlannedK As Double   ' Planned UNIT (col K)
    Dim totalCumQty As Double     ' Cum Qty (col N)
    Dim totalActPct As Double     ' Weighted AC%
    Dim totalPlannedPct As Double ' Weighted PL%
    Dim activityCount As Long
    Dim completedCount As Long

    totalBudget = 0
    totalPlannedK = 0
    totalCumQty = 0
    activityCount = 0
    completedCount = 0

    ' Check if PMS columns exist
    Dim hasPMS As Boolean
    hasPMS = (ws.Cells(1, 11).Value = "Planned UNIT")  ' Column K

    Dim i As Long
    For i = 2 To lastRow
        If Not DEE_Utils.IsWBSRow(ws, i) Then
            activityCount = activityCount + 1

            Dim budget As Double
            On Error Resume Next
            budget = CDbl(ws.Cells(i, 6).Value)  ' Column F
            On Error GoTo 0
            totalBudget = totalBudget + budget

            If hasPMS Then
                Dim planK As Double
                Dim cumN As Double
                Dim acPct As Double
                planK = 0
                cumN = 0
                acPct = 0
                On Error Resume Next
                planK = CDbl(ws.Cells(i, 11).Value)  ' K: Planned UNIT
                cumN = CDbl(ws.Cells(i, 14).Value)   ' N: Cum Qty
                acPct = CDbl(ws.Cells(i, 20).Value)  ' T: AC% Cum
                On Error GoTo 0
                totalPlannedK = totalPlannedK + planK
                totalCumQty = totalCumQty + cumN
                If acPct >= 1 Then completedCount = completedCount + 1
            End If
        End If
    Next i

    ' Compute overall percentages
    Dim overallPlanned As Double
    Dim overallActual As Double
    Dim overallVariance As Double
    Dim spi As Double

    If totalBudget > 0 Then
        overallPlanned = totalPlannedK / totalBudget
        overallActual = totalCumQty / totalBudget
    End If
    overallVariance = overallActual - overallPlanned
    spi = IIf(overallPlanned > 0, overallActual / overallPlanned, 0)

    Dim dataDate As Date
    dataDate = DEE_Utils.GetDataDate()

    ' -----------------------------------------------------------------------
    ' 2. Create or clear Dashboard sheet
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 30, "Building dashboard layout"

    Dim wsDash As Worksheet
    Set wsDash = DEE_Utils.GetOrCreateSheet(wb, "Dashboard")
    wsDash.Cells.Clear
    wsDash.Tab.Color = KPI_DARK

    ' Remove gridlines
    Application.ActiveWindow.DisplayGridlines = False

    ' -----------------------------------------------------------------------
    ' 3. Write Dashboard title
    ' -----------------------------------------------------------------------
    With wsDash.Range("A1:N1")
        .Merge
        .Value = "Project Controls Dashboard -- Protocol DEE"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = KPI_DARK
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 36
    End With

    With wsDash.Range("A2:N2")
        .Merge
        .Value = "Data Date: " & Format(dataDate, "dd-MMM-yyyy") & _
                 "    |    Total Activities: " & activityCount & _
                 "    |    Generated: " & Format(Now(), "dd-MMM-yyyy HH:MM")
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 64, 128)
        .HorizontalAlignment = xlCenter
    End With

    ' -----------------------------------------------------------------------
    ' 4. KPI Cards (row 4-8)
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 50, "Writing KPI cards"

    WriteKPICard wsDash, 4, 1, "Total Budget", Format(totalBudget, "$#,##0"), KPI_BLUE
    WriteKPICard wsDash, 4, 3, "Planned %", Format(overallPlanned, "0.0%"), KPI_BLUE
    WriteKPICard wsDash, 4, 5, "Actual %", Format(overallActual, "0.0%"), _
                 GetSPIColor(overallActual - overallPlanned)
    WriteKPICard wsDash, 4, 7, "Variance %", Format(overallVariance, "+0.0%;-0.0%"), _
                 GetVarianceColor(overallVariance)
    WriteKPICard wsDash, 4, 9, "SPI", Format(spi, "0.00"), GetSPIColor(spi - 1)
    WriteKPICard wsDash, 4, 11, "Completed", completedCount & " / " & activityCount, KPI_BLUE

    ' -----------------------------------------------------------------------
    ' 5. Activity summary table
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 70, "Writing activity summary"

    Dim tableStartRow As Integer
    tableStartRow = 11

    wsDash.Cells(tableStartRow, 1).Value = "ACTIVITY SUMMARY"
    wsDash.Cells(tableStartRow, 1).Font.Bold = True
    wsDash.Cells(tableStartRow, 1).Font.Size = 12
    wsDash.Cells(tableStartRow, 1).Interior.Color = KPI_DARK
    wsDash.Cells(tableStartRow, 1).Font.Color = RGB(255, 255, 255)
    wsDash.Range(wsDash.Cells(tableStartRow, 1), wsDash.Cells(tableStartRow, 6)).Merge

    DEE_Utils.ApplyTableHeader wsDash, tableStartRow + 1, 1, 6
    wsDash.Cells(tableStartRow + 1, 1).Value = "Activity ID"
    wsDash.Cells(tableStartRow + 1, 2).Value = "Name"
    wsDash.Cells(tableStartRow + 1, 3).Value = "Start"
    wsDash.Cells(tableStartRow + 1, 4).Value = "Finish"
    wsDash.Cells(tableStartRow + 1, 5).Value = "Budget"
    wsDash.Cells(tableStartRow + 1, 6).Value = "Status"

    ' Write top-level WBS rows to dashboard
    Dim dashRow As Integer
    dashRow = tableStartRow + 2
    Dim maxDashRows As Integer
    maxDashRows = 20

    For i = 2 To lastRow
        If dashRow > tableStartRow + 1 + maxDashRows Then Exit For
        If DEE_Utils.IsWBSRow(ws, i) Then
            Dim wbsLevel As Integer
            wbsLevel = DEE_Utils.GetWBSLevel(CStr(ws.Cells(i, 1).Value))
            If wbsLevel <= 2 Then
                wsDash.Cells(dashRow, 1).Value = Trim(ws.Cells(i, 1).Value)
                wsDash.Cells(dashRow, 2).Value = Trim(ws.Cells(i, 2).Value)
                wsDash.Cells(dashRow, 3).Value = ""
                wsDash.Cells(dashRow, 4).Value = ""
                wsDash.Cells(dashRow, 5).Value = ws.Cells(i, 6).Value
                wsDash.Cells(dashRow, 5).NumberFormat = "$#,##0"
                ' Color by level
                wsDash.Range(wsDash.Cells(dashRow, 1), wsDash.Cells(dashRow, 6)).Interior.Color = _
                    DEE_WBS.GetWBSColorForLevel(wbsLevel)
                wsDash.Range(wsDash.Cells(dashRow, 1), wsDash.Cells(dashRow, 6)).Font.Color = _
                    IIf(wbsLevel <= 3, RGB(255, 255, 255), RGB(0, 0, 0))
                dashRow = dashRow + 1
            End If
        End If
    Next i

    ' -----------------------------------------------------------------------
    ' 6. Create progress chart
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 85, "Creating charts"
    CreateProgressChart wsDash, overallPlanned, overallActual

    ' -----------------------------------------------------------------------
    ' 7. Formatting
    ' -----------------------------------------------------------------------
    wsDash.Columns("A:N").AutoFit
    wsDash.Rows("1:2").RowHeight = 30

    DEE_Utils.EndProgress
    wsDash.Activate
    MsgBox "Dashboard created successfully!", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' WriteKPICard -- Write a styled KPI card in a 2-column merged area
' ---------------------------------------------------------------------------
Private Sub WriteKPICard(ws As Worksheet, startRow As Integer, startCol As Integer, _
                           title As String, value As String, bgColor As Long)
    ' Title cell
    Dim titleRange As Range
    Set titleRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 1))
    titleRange.Merge
    With titleRange
        .Value = title
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = bgColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With

    ' Value cell
    Dim valueRange As Range
    Set valueRange = ws.Range(ws.Cells(startRow + 1, startCol), ws.Cells(startRow + 2, startCol + 1))
    valueRange.Merge
    With valueRange
        .Value = value
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = bgColor
        .Interior.Color = RGB(245, 245, 245)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = bgColor
        End With
    End With
End Sub

' ---------------------------------------------------------------------------
' GetSPIColor -- Return color based on SPI deviation
' ---------------------------------------------------------------------------
Private Function GetSPIColor(deviation As Double) As Long
    If deviation >= 0 Then
        GetSPIColor = KPI_GREEN
    ElseIf deviation >= -0.05 Then
        GetSPIColor = KPI_AMBER
    Else
        GetSPIColor = KPI_RED
    End If
End Function

' ---------------------------------------------------------------------------
' GetVarianceColor -- Return color for variance
' ---------------------------------------------------------------------------
Private Function GetVarianceColor(variance As Double) As Long
    If variance >= 0 Then
        GetVarianceColor = KPI_GREEN
    ElseIf variance >= -0.05 Then
        GetVarianceColor = KPI_AMBER
    Else
        GetVarianceColor = KPI_RED
    End If
End Function

' ---------------------------------------------------------------------------
' CreateProgressChart -- Doughnut chart showing planned vs actual
' ---------------------------------------------------------------------------
Private Sub CreateProgressChart(ws As Worksheet, plannedPct As Double, actualPct As Double)
    ' Write chart data in a temp area
    Dim chartDataCol As Integer
    chartDataCol = 16  ' Column P

    ws.Cells(4, chartDataCol).Value = "Category"
    ws.Cells(4, chartDataCol + 1).Value = "Value"
    ws.Cells(5, chartDataCol).Value = "Planned"
    ws.Cells(5, chartDataCol + 1).Value = plannedPct
    ws.Cells(6, chartDataCol).Value = "Actual"
    ws.Cells(6, chartDataCol + 1).Value = actualPct
    ws.Cells(7, chartDataCol).Value = "Remaining"
    ws.Cells(7, chartDataCol + 1).Value = IIf(1 - actualPct > 0, 1 - actualPct, 0)

    ws.Range(ws.Cells(4, chartDataCol), ws.Cells(7, chartDataCol + 1)).NumberFormat = "0.0%"

    ' Create chart
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=450, Top:=80, Width:=300, Height:=200)

    With cht.Chart
        .ChartType = xlDoughnut

        Dim dataRange As Range
        Set dataRange = ws.Range(ws.Cells(4, chartDataCol), ws.Cells(7, chartDataCol + 1))
        .SetSourceData Source:=dataRange

        .HasTitle = True
        .ChartTitle.Text = "Progress: Planned vs Actual"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Bold = True

        ' Color series
        If .SeriesCollection.Count > 0 Then
            With .SeriesCollection(1)
                If .Points.Count >= 3 Then
                    .Points(1).Format.Fill.ForeColor.RGB = RGB(0, 112, 192)   ' Planned - blue
                    .Points(2).Format.Fill.ForeColor.RGB = RGB(255, 153, 0)   ' Actual - orange
                    .Points(3).Format.Fill.ForeColor.RGB = RGB(220, 220, 220) ' Remaining - gray
                End If
            End With
        End If

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
End Sub
