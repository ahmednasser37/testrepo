Attribute VB_Name = "DEE_Distribute"
Option Explicit

' =============================================================================
' DEE_Distribute.bas -- Smart Spread (Quantity Distribution)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Distributes a total quantity across a date range using various profiles:
'   - Linear (uniform)
'   - Front-loaded (early concentration)
'   - Back-loaded (late concentration)
'   - Bell curve (normal distribution)
'   - Custom (user-defined percentage table)
'
' Reference: Plexxel_Analysis/vba/DistributeFunctions.bas (732 lines)
' =============================================================================

' Distribution profile constants
Private Const DIST_LINEAR As Integer = 1
Private Const DIST_FRONT As Integer = 2
Private Const DIST_BACK As Integer = 3
Private Const DIST_BELL As Integer = 4
Private Const DIST_CUSTOM As Integer = 5

' ---------------------------------------------------------------------------
' SmartSpread -- Main entry point (called from Ribbon)
' ---------------------------------------------------------------------------
Public Sub SmartSpread()
    ' Get active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Prompt for parameters
    Dim totalQtyStr As String
    totalQtyStr = InputBox("Enter total quantity to distribute:", "Smart Spread", "1000")
    If totalQtyStr = "" Then Exit Sub

    Dim totalQty As Double
    On Error Resume Next
    totalQty = CDbl(totalQtyStr)
    On Error GoTo 0
    If totalQty = 0 Then MsgBox "Invalid quantity.", vbExclamation: Exit Sub

    Dim startStr As String
    startStr = InputBox("Enter start date (yyyy-mm-dd):", "Smart Spread", _
                        Format(Date, "yyyy-mm-dd"))
    If startStr = "" Then Exit Sub
    Dim startDate As Variant
    startDate = DEE_Utils.SafeParseDate(startStr)
    If IsEmpty(startDate) Then MsgBox "Invalid start date.", vbExclamation: Exit Sub

    Dim endStr As String
    endStr = InputBox("Enter end date (yyyy-mm-dd):", "Smart Spread", _
                      Format(Date + 180, "yyyy-mm-dd"))
    If endStr = "" Then Exit Sub
    Dim endDate As Variant
    endDate = DEE_Utils.SafeParseDate(endStr)
    If IsEmpty(endDate) Then MsgBox "Invalid end date.", vbExclamation: Exit Sub

    If CDate(endDate) <= CDate(startDate) Then
        MsgBox "End date must be after start date.", vbExclamation: Exit Sub
    End If

    Dim profileStr As String
    profileStr = InputBox("Select distribution profile:" & vbCrLf & _
                          "1 = Linear (uniform)" & vbCrLf & _
                          "2 = Front-loaded" & vbCrLf & _
                          "3 = Back-loaded" & vbCrLf & _
                          "4 = Bell curve" & vbCrLf & _
                          "5 = Custom", "Smart Spread", "1")
    If profileStr = "" Then Exit Sub

    Dim profile As Integer
    On Error Resume Next
    profile = CInt(profileStr)
    On Error GoTo 0
    If profile < 1 Or profile > 5 Then profile = DIST_LINEAR

    Dim granularityStr As String
    granularityStr = InputBox("Granularity:" & vbCrLf & _
                              "1 = Monthly" & vbCrLf & "2 = Weekly", "Smart Spread", "1")
    Dim granularity As Integer
    On Error Resume Next
    granularity = CInt(granularityStr)
    On Error GoTo 0
    If granularity <> 1 And granularity <> 2 Then granularity = 1

    DEE_Utils.StartProgress "Computing Smart Spread"

    ' Build period array
    Dim periods() As Date
    Dim periodCount As Integer
    BuildPeriods CDate(startDate), CDate(endDate), granularity, periods, periodCount

    ' Compute distribution weights
    Dim weights() As Double
    ReDim weights(0 To periodCount - 1)
    ComputeWeights profile, periodCount, weights

    ' Compute quantities
    Dim quantities() As Double
    ReDim quantities(0 To periodCount - 1)
    Dim p As Integer
    For p = 0 To periodCount - 1
        quantities(p) = totalQty * weights(p)
    Next p

    ' Write to active sheet (or new sheet)
    Dim outputChoice As String
    outputChoice = InputBox("Output to:" & vbCrLf & _
                            "1 = Current sheet (at cursor)" & vbCrLf & _
                            "2 = New sheet 'Spread'", "Smart Spread", "2")
    Dim outputWs As Worksheet
    If outputChoice = "1" Then
        Set outputWs = ws
        Dim outputRow As Long
        outputRow = IIf(Selection Is Nothing, 1, Selection.Row)
        Dim outputCol As Integer
        outputCol = IIf(Selection Is Nothing, 1, Selection.Column)
    Else
        Set outputWs = DEE_Utils.GetOrCreateSheet(ActiveWorkbook, "Spread")
        outputWs.Cells.Clear
        outputRow = 1
        outputCol = 1
    End If

    WriteSpreadOutput outputWs, outputRow, outputCol, periods, quantities, periodCount, _
                      totalQty, granularity, profile

    DEE_Utils.EndProgress
    outputWs.Activate
    MsgBox "Smart Spread complete!" & vbCrLf & _
           periodCount & " periods, Total: " & Format(totalQty, "#,##0.00"), _
           vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' BuildPeriods -- Create array of period start dates
' ---------------------------------------------------------------------------
Private Sub BuildPeriods(startDate As Date, endDate As Date, granularity As Integer, _
                          ByRef periods() As Date, ByRef periodCount As Integer)
    Dim d As Date
    Dim maxPeriods As Integer
    maxPeriods = 1000

    ReDim periods(0 To maxPeriods)
    periodCount = 0

    d = IIf(granularity = 1, _
            DateSerial(Year(startDate), Month(startDate), 1), _
            startDate - Weekday(startDate, vbMonday) + 1)

    Do While d <= endDate And periodCount <= maxPeriods
        periods(periodCount) = d
        periodCount = periodCount + 1
        If granularity = 1 Then
            d = DateSerial(Year(d), Month(d) + 1, 1)
        Else
            d = d + 7
        End If
    Loop
End Sub

' ---------------------------------------------------------------------------
' ComputeWeights -- Compute normalized distribution weights
' ---------------------------------------------------------------------------
Private Sub ComputeWeights(profile As Integer, periodCount As Integer, weights() As Double)
    Dim p As Integer
    Dim totalWeight As Double
    totalWeight = 0

    Select Case profile
        Case DIST_LINEAR
            For p = 0 To periodCount - 1
                weights(p) = 1
            Next p

        Case DIST_FRONT
            ' Front-loaded: decreasing weights
            For p = 0 To periodCount - 1
                weights(p) = (periodCount - p)
            Next p

        Case DIST_BACK
            ' Back-loaded: increasing weights
            For p = 0 To periodCount - 1
                weights(p) = p + 1
            Next p

        Case DIST_BELL
            ' Bell curve (normal distribution approximation)
            Dim mid As Double
            mid = (periodCount - 1) / 2
            Dim sigma As Double
            sigma = periodCount / 4
            For p = 0 To periodCount - 1
                Dim x As Double
                x = p - mid
                weights(p) = Exp(-(x * x) / (2 * sigma * sigma))
            Next p

        Case DIST_CUSTOM
            ' Equal for now -- user would fill in via UI
            For p = 0 To periodCount - 1
                weights(p) = 1
            Next p

    End Select

    ' Normalize weights to sum to 1
    For p = 0 To periodCount - 1
        totalWeight = totalWeight + weights(p)
    Next p

    If totalWeight > 0 Then
        For p = 0 To periodCount - 1
            weights(p) = weights(p) / totalWeight
        Next p
    End If
End Sub

' ---------------------------------------------------------------------------
' WriteSpreadOutput -- Write distribution table to worksheet
' ---------------------------------------------------------------------------
Private Sub WriteSpreadOutput(ws As Worksheet, startRow As Long, startCol As Integer, _
                               periods() As Date, quantities() As Double, periodCount As Integer, _
                               totalQty As Double, granularity As Integer, profile As Integer)

    Dim profileName As String
    Select Case profile
        Case 1: profileName = "Linear"
        Case 2: profileName = "Front-Loaded"
        Case 3: profileName = "Back-Loaded"
        Case 4: profileName = "Bell Curve"
        Case 5: profileName = "Custom"
        Case Else: profileName = "Unknown"
    End Select

    Dim granName As String
    granName = IIf(granularity = 1, "Monthly", "Weekly")

    ' Title row
    ws.Cells(startRow, startCol).Value = "Smart Spread -- " & profileName & " (" & granName & ")"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Font.Size = 12

    ' Header
    ws.Cells(startRow + 1, startCol).Value = "Period"
    ws.Cells(startRow + 1, startCol + 1).Value = "Quantity"
    ws.Cells(startRow + 1, startCol + 2).Value = "Percentage"
    ws.Cells(startRow + 1, startCol + 3).Value = "Cumulative Qty"
    ws.Cells(startRow + 1, startCol + 4).Value = "Cumulative %"

    DEE_Utils.ApplyTableHeader ws, startRow + 1, startCol, startCol + 4

    Dim cumQty As Double
    cumQty = 0
    Dim p As Integer
    For p = 0 To periodCount - 1
        Dim dataRow As Long
        dataRow = startRow + 2 + p

        Dim periodLabel As String
        If granularity = 1 Then
            periodLabel = Format(periods(p), "mmm-yy")
        Else
            periodLabel = Format(periods(p), "dd-mmm-yy")
        End If

        cumQty = cumQty + quantities(p)

        ws.Cells(dataRow, startCol).Value = periodLabel
        ws.Cells(dataRow, startCol + 1).Value = quantities(p)
        ws.Cells(dataRow, startCol + 1).NumberFormat = "#,##0.00"
        ws.Cells(dataRow, startCol + 2).Value = IIf(totalQty > 0, quantities(p) / totalQty, 0)
        ws.Cells(dataRow, startCol + 2).NumberFormat = "0.00%"
        ws.Cells(dataRow, startCol + 3).Value = cumQty
        ws.Cells(dataRow, startCol + 3).NumberFormat = "#,##0.00"
        ws.Cells(dataRow, startCol + 4).Value = IIf(totalQty > 0, cumQty / totalQty, 0)
        ws.Cells(dataRow, startCol + 4).NumberFormat = "0.00%"
    Next p

    ' Total row
    Dim totalRow As Long
    totalRow = startRow + 2 + periodCount
    ws.Cells(totalRow, startCol).Value = "TOTAL"
    ws.Cells(totalRow, startCol).Font.Bold = True
    ws.Cells(totalRow, startCol + 1).Formula = "=SUM(" & DEE_Utils.ColLetter(startCol + 1) & _
        (startRow + 2) & ":" & DEE_Utils.ColLetter(startCol + 1) & (totalRow - 1) & ")"
    ws.Cells(totalRow, startCol + 1).Font.Bold = True
    ws.Cells(totalRow, startCol + 1).NumberFormat = "#,##0.00"

    ws.Columns(startCol & ":" & (startCol + 4)).AutoFit

    ' Create a mini chart
    CreateSpreadChart ws, startRow + 1, startCol, periodCount
End Sub

' ---------------------------------------------------------------------------
' CreateSpreadChart -- Bar chart of the spread distribution
' ---------------------------------------------------------------------------
Private Sub CreateSpreadChart(ws As Worksheet, headerRow As Long, startCol As Integer, _
                               periodCount As Integer)
    ' Data range: period + quantity columns
    Dim chartStartRow As Long
    chartStartRow = headerRow
    Dim chartEndRow As Long
    chartEndRow = headerRow + periodCount

    Dim dataRange As Range
    Set dataRange = ws.Range( _
        ws.Cells(chartStartRow, startCol), _
        ws.Cells(chartEndRow, startCol + 1))

    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add( _
        Left:=ws.Cells(headerRow, startCol + 6).Left, _
        Top:=ws.Cells(headerRow, startCol + 6).Top, _
        Width:=380, Height:=200)

    With cht.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=dataRange
        .HasTitle = True
        .ChartTitle.Text = "Distribution Profile"
        .ChartTitle.Font.Size = 10
        .HasLegend = False

        If .SeriesCollection.Count > 0 Then
            With .SeriesCollection(1)
                .Format.Fill.ForeColor.RGB = RGB(0, 112, 192)
            End With
        End If
    End With
End Sub
