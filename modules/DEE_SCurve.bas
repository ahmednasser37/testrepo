Attribute VB_Name = "DEE_SCurve"
Option Explicit

' =============================================================================
' DEE_SCurve.bas -- S-Curve Chart Generation
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Creates monthly S-Curve (Planned vs Actual) from the Schedule sheet PMS data.
' Reads snapshot sheets (Snap_yyyy-mm-dd) for historical actual data.
'
' OUTPUT SHEET: "S-Curve Monthly"
' CHART TYPE: Line chart with two series: Planned (blue) + Actual (orange)
' =============================================================================

' ---------------------------------------------------------------------------
' CreateSCurveMonthly -- Main entry point
' ---------------------------------------------------------------------------
Public Sub CreateSCurveMonthly()
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

    DEE_Utils.StartProgress "Building S-Curve"

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        MsgBox "Schedule sheet is empty.", vbExclamation, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' 1. Find project date range
    ' -----------------------------------------------------------------------
    Dim projStart As Date
    Dim projEnd As Date
    projStart = #1/1/9999#
    projEnd = #1/1/1900#

    Dim i As Long
    For i = 2 To lastRow
        If Not DEE_Utils.IsWBSRow(ws, i) Then
            Dim s As Variant
            Dim e As Variant
            s = ws.Cells(i, 3).Value
            e = ws.Cells(i, 4).Value
            If IsDate(s) And CDate(s) < projStart Then projStart = CDate(s)
            If IsDate(e) And CDate(e) > projEnd Then projEnd = CDate(e)
        End If
    Next i

    If projStart = #1/1/9999# Then
        MsgBox "No valid activity dates found.", vbExclamation, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' 2. Build monthly period array
    ' -----------------------------------------------------------------------
    Dim periodStart As Date
    periodStart = DateSerial(Year(projStart), Month(projStart), 1)
    Dim periodEnd As Date
    periodEnd = DateSerial(Year(projEnd), Month(projEnd) + 1, 1)

    Dim periodCount As Integer
    Dim d As Date
    d = periodStart
    periodCount = 0
    Do While d <= periodEnd
        periodCount = periodCount + 1
        d = DateSerial(Year(d), Month(d) + 1, 1)
    Loop

    Dim periods() As Date
    ReDim periods(0 To periodCount - 1)
    d = periodStart
    Dim p As Integer
    For p = 0 To periodCount - 1
        periods(p) = d
        d = DateSerial(Year(d), Month(d) + 1, 1)
    Next p

    ' -----------------------------------------------------------------------
    ' 3. Compute planned S-curve (cumulative from NETWORKDAYS formula)
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 20, "Computing planned S-curve"

    Dim plannedMonthly() As Double
    ReDim plannedMonthly(0 To periodCount - 1)

    Dim totalBudget As Double
    totalBudget = 0

    For i = 2 To lastRow
        If Not DEE_Utils.IsWBSRow(ws, i) Then
            Dim taskBudget As Double
            Dim taskStart As Variant
            Dim taskEnd As Variant
            taskBudget = 0
            On Error Resume Next
            taskBudget = CDbl(ws.Cells(i, 6).Value)  ' Column F (locked budget)
            On Error GoTo 0
            taskStart = ws.Cells(i, 3).Value
            taskEnd = ws.Cells(i, 4).Value

            If IsDate(taskStart) And IsDate(taskEnd) And taskBudget > 0 Then
                ' Distribute budget linearly across months
                Dim ts As Date
                Dim te As Date
                ts = CDate(taskStart)
                te = CDate(taskEnd)
                Dim durDays As Long
                durDays = te - ts + 1
                If durDays < 1 Then durDays = 1

                ' Find overlap with each period
                For p = 0 To periodCount - 1
                    Dim pStart As Date
                    Dim pEnd As Date
                    pStart = periods(p)
                    If p < periodCount - 1 Then
                        pEnd = periods(p + 1) - 1
                    Else
                        pEnd = DateSerial(Year(pStart), Month(pStart) + 1, 0)
                    End If

                    ' Compute overlap
                    Dim overlapStart As Date
                    Dim overlapEnd As Date
                    overlapStart = IIf(ts > pStart, ts, pStart)
                    overlapEnd = IIf(te < pEnd, te, pEnd)

                    If overlapEnd >= overlapStart Then
                        Dim overlapDays As Long
                        overlapDays = overlapEnd - overlapStart + 1
                        plannedMonthly(p) = plannedMonthly(p) + _
                            (taskBudget * CDbl(overlapDays) / CDbl(durDays))
                    End If
                Next p

                totalBudget = totalBudget + taskBudget
            End If
        End If
    Next i

    ' -----------------------------------------------------------------------
    ' 4. Convert to cumulative planned
    ' -----------------------------------------------------------------------
    Dim plannedCum() As Double
    ReDim plannedCum(0 To periodCount - 1)
    Dim cumPlan As Double
    cumPlan = 0
    For p = 0 To periodCount - 1
        cumPlan = cumPlan + plannedMonthly(p)
        plannedCum(p) = IIf(totalBudget > 0, cumPlan / totalBudget * 100, 0)
    Next p

    ' -----------------------------------------------------------------------
    ' 5. Read actual data from snapshots
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 50, "Reading snapshot history"

    Dim actualCum() As Double
    ReDim actualCum(0 To periodCount - 1)
    ' Initialize to -1 (no data)
    For p = 0 To periodCount - 1
        actualCum(p) = -1
    Next p

    ' Read each Snap_yyyy-mm-dd sheet
    Dim snapWs As Worksheet
    For Each snapWs In wb.Worksheets
        If Left(snapWs.Name, 5) = "Snap_" Then
            ' Extract date from sheet name
            Dim snapDateStr As String
            snapDateStr = Mid(snapWs.Name, 6)
            Dim snapDate As Variant
            snapDate = DEE_Utils.SafeParseDate(snapDateStr)
            If Not IsEmpty(snapDate) Then
                ' Find period index for this snapshot
                Dim snapPeriod As Integer
                snapPeriod = -1
                For p = 0 To periodCount - 1
                    Dim nextPeriodDate As Date
                    If p < periodCount - 1 Then
                        nextPeriodDate = periods(p + 1)
                    Else
                        nextPeriodDate = DateSerial(Year(periods(p)), Month(periods(p)) + 1, 1)
                    End If
                    If CDate(snapDate) >= periods(p) And CDate(snapDate) < nextPeriodDate Then
                        snapPeriod = p
                        Exit For
                    End If
                Next p

                If snapPeriod >= 0 Then
                    ' Read overall AC% from the snapshot (column T aggregated, row 2)
                    ' Column T = 20 (AC% Cum.)
                    Dim snapAC As Double
                    On Error Resume Next
                    ' Skip the metadata row inserted in snapshot
                    snapAC = CDbl(snapWs.Cells(3, 20).Value) * 100  ' Row 3 because row 1 is metadata
                    On Error GoTo 0
                    If snapAC > 0 Then actualCum(snapPeriod) = snapAC
                End If
            End If
        End If
    Next snapWs

    ' -----------------------------------------------------------------------
    ' 6. Write S-Curve data sheet
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 70, "Writing S-Curve data"

    Dim wsSC As Worksheet
    Set wsSC = DEE_Utils.GetOrCreateSheet(wb, "S-Curve Monthly")
    wsSC.Cells.Clear

    ' Headers
    wsSC.Cells(1, 1).Value = "Period"
    wsSC.Cells(1, 2).Value = "Planned Cum %"
    wsSC.Cells(1, 3).Value = "Actual Cum %"
    DEE_Utils.ApplyTableHeader wsSC, 1, 1, 3

    For p = 0 To periodCount - 1
        wsSC.Cells(p + 2, 1).Value = Format(periods(p), "mmm-yy")
        wsSC.Cells(p + 2, 2).Value = plannedCum(p) / 100  ' As decimal
        wsSC.Cells(p + 2, 2).NumberFormat = "0.0%"
        If actualCum(p) >= 0 Then
            wsSC.Cells(p + 2, 3).Value = actualCum(p) / 100
            wsSC.Cells(p + 2, 3).NumberFormat = "0.0%"
        End If
    Next p

    wsSC.Columns("A:C").AutoFit

    ' -----------------------------------------------------------------------
    ' 7. Create chart
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 90, "Creating chart"
    CreateSCurveChart wsSC, periodCount

    DEE_Utils.EndProgress
    wsSC.Activate
    MsgBox "S-Curve created successfully!", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' CreateSCurveChart -- Build the line chart
' ---------------------------------------------------------------------------
Private Sub CreateSCurveChart(ws As Worksheet, periodCount As Integer)
    ' Remove existing chart if any
    Dim cht As ChartObject
    For Each cht In ws.ChartObjects
        cht.Delete
    Next cht

    ' Add chart
    Dim newChart As ChartObject
    Set newChart = ws.ChartObjects.Add(Left:=200, Top:=20, Width:=500, Height:=280)

    With newChart.Chart
        .ChartType = xlLine

        ' Data range: columns A-C
        Dim dataRange As Range
        Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(periodCount + 1, 3))
        .SetSourceData Source:=dataRange

        ' Series formatting
        With .SeriesCollection(1)
            .Name = "Planned"
            .Format.Line.ForeColor.RGB = RGB(0, 112, 192)
            .Format.Line.Weight = 2
        End With

        If .SeriesCollection.Count > 1 Then
            With .SeriesCollection(2)
                .Name = "Actual"
                .Format.Line.ForeColor.RGB = RGB(255, 153, 0)
                .Format.Line.Weight = 2
            End With
        End If

        ' Title
        .HasTitle = True
        .ChartTitle.Text = "S-Curve: Planned vs Actual Progress"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True

        ' Axes
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 1

        ' Legend
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        ' Plot area
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
End Sub
