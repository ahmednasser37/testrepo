Attribute VB_Name = "DEE_Reporting"
Option Explicit

' =============================================================================
' DEE_Reporting.bas -- Look-Ahead Reports & Reporting Utilities
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Generates printable look-ahead reports from the Schedule sheet.
'
' LOOK-AHEAD REPORT:
'   Filters activities where Start or Finish falls within N days of Data Date.
'   Categories: Starting, Finishing, In-Progress.
'   Output: landscape, dark header, auto-fit columns.
'
' LOOK-AHEAD PERIODS: 2 weeks (14 days), 4 weeks (28 days), or custom.
' =============================================================================

' ---------------------------------------------------------------------------
' CreateLookAhead -- Generic look-ahead report builder
' ---------------------------------------------------------------------------
Public Sub CreateLookAhead(Optional windowDays As Integer = 14)
    ' Allow override via InputBox if called directly
    If windowDays = 0 Then
        Dim daysStr As String
        daysStr = InputBox("Enter look-ahead window (days):", "Look-Ahead Report", "14")
        If daysStr = "" Then Exit Sub
        On Error Resume Next
        windowDays = CInt(daysStr)
        On Error GoTo 0
        If windowDays <= 0 Then windowDays = 14
    End If

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

    DEE_Utils.StartProgress "Building Look-Ahead Report"

    Dim lastRow As Long
    lastRow = DEE_Utils.LastRow(ws, 1)
    If lastRow < 2 Then
        MsgBox "Schedule sheet is empty.", vbExclamation, "Protocol DEE"
        DEE_Utils.EndProgress
        Exit Sub
    End If

    Dim dataDate As Date
    dataDate = DEE_Utils.GetDataDate()
    Dim windowEnd As Date
    windowEnd = dataDate + windowDays

    ' -----------------------------------------------------------------------
    ' Filter activities
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 20, "Filtering activities"

    ' Three categories: Starting, Finishing, In-Progress
    Dim startingRows() As Long
    Dim finishingRows() As Long
    Dim inProgressRows() As Long
    Dim startingCount As Integer
    Dim finishingCount As Integer
    Dim inProgressCount As Integer

    ReDim startingRows(0 To lastRow)
    ReDim finishingRows(0 To lastRow)
    ReDim inProgressRows(0 To lastRow)
    startingCount = 0
    finishingCount = 0
    inProgressCount = 0

    Dim i As Long
    For i = 2 To lastRow
        If Not DEE_Utils.IsWBSRow(ws, i) Then
            Dim taskStart As Variant
            Dim taskEnd As Variant
            taskStart = ws.Cells(i, 3).Value
            taskEnd = ws.Cells(i, 4).Value

            If IsDate(taskStart) And IsDate(taskEnd) Then
                Dim ts As Date
                Dim te As Date
                ts = CDate(taskStart)
                te = CDate(taskEnd)

                ' Starting: starts within window
                If ts >= dataDate And ts <= windowEnd Then
                    startingRows(startingCount) = i
                    startingCount = startingCount + 1
                ' Finishing: finishes within window (but started before)
                ElseIf te >= dataDate And te <= windowEnd And ts < dataDate Then
                    finishingRows(finishingCount) = i
                    finishingCount = finishingCount + 1
                ' In-Progress: started before window, finishes after
                ElseIf ts < dataDate And te > windowEnd Then
                    inProgressRows(inProgressCount) = i
                    inProgressCount = inProgressCount + 1
                End If
            End If
        End If
    Next i

    ' -----------------------------------------------------------------------
    ' Create output sheet
    ' -----------------------------------------------------------------------
    Dim weekNum As Integer
    weekNum = windowDays \ 7
    Dim sheetName As String
    sheetName = "LookAhead_" & weekNum & "Wk"

    Dim wsLA As Worksheet
    Set wsLA = DEE_Utils.GetOrCreateSheet(wb, sheetName)
    wsLA.Cells.Clear

    ' -----------------------------------------------------------------------
    ' Write report
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 60, "Writing report"

    ' Title
    Dim titleRow As Integer
    titleRow = 1
    wsLA.Cells(titleRow, 1).Value = weekNum & "-Week Look-Ahead Report"
    wsLA.Cells(titleRow, 1).Font.Bold = True
    wsLA.Cells(titleRow, 1).Font.Size = 16
    wsLA.Range(wsLA.Cells(titleRow, 1), wsLA.Cells(titleRow, 7)).Merge

    wsLA.Cells(titleRow + 1, 1).Value = _
        "Data Date: " & Format(dataDate, "dd-MMM-yyyy") & _
        "    Look-Ahead End: " & Format(windowEnd, "dd-MMM-yyyy") & _
        "    Activities: " & (startingCount + finishingCount + inProgressCount)
    wsLA.Range(wsLA.Cells(titleRow + 1, 1), wsLA.Cells(titleRow + 1, 7)).Merge
    wsLA.Rows(titleRow & ":" & (titleRow + 1)).Interior.Color = RGB(0, 32, 96)
    wsLA.Rows(titleRow & ":" & (titleRow + 1)).Font.Color = RGB(255, 255, 255)

    Dim currentRow As Integer
    currentRow = titleRow + 3

    ' Write each section
    currentRow = WriteLookAheadSection(ws, wsLA, currentRow, _
                                       "ACTIVITIES STARTING (" & startingCount & ")", _
                                       startingRows, startingCount, RGB(0, 176, 80))

    currentRow = currentRow + 1
    currentRow = WriteLookAheadSection(ws, wsLA, currentRow, _
                                       "ACTIVITIES FINISHING (" & finishingCount & ")", _
                                       finishingRows, finishingCount, RGB(255, 192, 0))

    currentRow = currentRow + 1
    currentRow = WriteLookAheadSection(ws, wsLA, currentRow, _
                                       "ACTIVITIES IN PROGRESS (" & inProgressCount & ")", _
                                       inProgressRows, inProgressCount, RGB(0, 112, 192))

    ' -----------------------------------------------------------------------
    ' Print setup
    ' -----------------------------------------------------------------------
    DEE_Utils.UpdateProgress 90, "Configuring print layout"
    SetupPrintLayout wsLA, dataDate

    wsLA.Columns("A:G").AutoFit

    DEE_Utils.EndProgress
    wsLA.Activate

    Dim totalFound As Integer
    totalFound = startingCount + finishingCount + inProgressCount
    MsgBox weekNum & "-Week Look-Ahead Report created!" & vbCrLf & _
           totalFound & " activities found.", vbInformation, "Protocol DEE"
End Sub

' ---------------------------------------------------------------------------
' WriteLookAheadSection -- Write one section (Starting/Finishing/InProgress)
' Returns the next available row
' ---------------------------------------------------------------------------
Private Function WriteLookAheadSection(wsSource As Worksheet, wsDest As Worksheet, _
                                        startRow As Integer, sectionTitle As String, _
                                        rows() As Long, rowCount As Integer, _
                                        headerColor As Long) As Integer

    ' Section header
    wsDest.Cells(startRow, 1).Value = sectionTitle
    Dim headerRange As Range
    Set headerRange = wsDest.Range(wsDest.Cells(startRow, 1), wsDest.Cells(startRow, 7))
    headerRange.Merge
    With headerRange
        .Interior.Color = headerColor
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .RowHeight = 20
    End With

    ' Column headers
    Dim colHeaderRow As Integer
    colHeaderRow = startRow + 1
    wsDest.Cells(colHeaderRow, 1).Value = "Activity ID"
    wsDest.Cells(colHeaderRow, 2).Value = "Activity Name"
    wsDest.Cells(colHeaderRow, 3).Value = "Start"
    wsDest.Cells(colHeaderRow, 4).Value = "Finish"
    wsDest.Cells(colHeaderRow, 5).Value = "Budget"
    wsDest.Cells(colHeaderRow, 6).Value = "Planned %"
    wsDest.Cells(colHeaderRow, 7).Value = "Actual %"

    With wsDest.Range(wsDest.Cells(colHeaderRow, 1), wsDest.Cells(colHeaderRow, 7))
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 18
    End With

    Dim dataRow As Integer
    dataRow = colHeaderRow + 1

    ' Check if PMS columns exist
    Dim hasPMS As Boolean
    hasPMS = (wsSource.Cells(1, 17).Value = "PL% Cum.")  ' Column Q

    Dim r As Integer
    For r = 0 To rowCount - 1
        Dim srcRow As Long
        srcRow = rows(r)
        If srcRow = 0 Then GoTo NextReportRow

        wsDest.Cells(dataRow, 1).Value = Trim(wsSource.Cells(srcRow, 1).Value)
        wsDest.Cells(dataRow, 2).Value = Trim(wsSource.Cells(srcRow, 2).Value)
        wsDest.Cells(dataRow, 3).Value = wsSource.Cells(srcRow, 3).Value
        wsDest.Cells(dataRow, 3).NumberFormat = "dd-mmm-yyyy"
        wsDest.Cells(dataRow, 4).Value = wsSource.Cells(srcRow, 4).Value
        wsDest.Cells(dataRow, 4).NumberFormat = "dd-mmm-yyyy"
        wsDest.Cells(dataRow, 5).Value = wsSource.Cells(srcRow, 6).Value
        wsDest.Cells(dataRow, 5).NumberFormat = "$#,##0"

        If hasPMS Then
            wsDest.Cells(dataRow, 6).Value = wsSource.Cells(srcRow, 17).Value  ' Q: PL% Cum
            wsDest.Cells(dataRow, 6).NumberFormat = "0.0%"
            wsDest.Cells(dataRow, 7).Value = wsSource.Cells(srcRow, 20).Value  ' T: AC% Cum
            wsDest.Cells(dataRow, 7).NumberFormat = "0.0%"
        End If

        ' Alternate row shading
        If (dataRow Mod 2) = 0 Then
            wsDest.Range(wsDest.Cells(dataRow, 1), wsDest.Cells(dataRow, 7)).Interior.Color = _
                RGB(240, 240, 240)
        End If

        dataRow = dataRow + 1
NextReportRow:
    Next r

    If rowCount = 0 Then
        wsDest.Cells(dataRow, 1).Value = "(No activities in this category)"
        wsDest.Cells(dataRow, 1).Font.Italic = True
        wsDest.Cells(dataRow, 1).Font.Color = RGB(128, 128, 128)
        dataRow = dataRow + 1
    End If

    WriteLookAheadSection = dataRow
End Function

' ---------------------------------------------------------------------------
' SetupPrintLayout -- Configure sheet for printing
' ---------------------------------------------------------------------------
Public Sub SetupPrintLayout(ws As Worksheet, dataDate As Date)
    With ws.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$2"
        .LeftFooter = ws.Parent.Name
        .CenterFooter = "Data Date: " & Format(dataDate, "dd-mmm-yyyy")
        .RightFooter = "Page &P of &N"
        .PrintArea = ws.UsedRange.Address
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
    End With
End Sub

' ---------------------------------------------------------------------------
' CreateLookAhead2W -- 2-week look-ahead (Ribbon button shortcut)
' ---------------------------------------------------------------------------
Public Sub CreateLookAhead2W()
    CreateLookAhead 14
End Sub

' ---------------------------------------------------------------------------
' CreateLookAhead4W -- 4-week look-ahead (Ribbon button shortcut)
' ---------------------------------------------------------------------------
Public Sub CreateLookAhead4W()
    CreateLookAhead 28
End Sub

' ---------------------------------------------------------------------------
' SetupPrint -- Print setup for active sheet (Ribbon button)
' ---------------------------------------------------------------------------
Public Sub SetupPrint()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim dataDate As Date
    On Error Resume Next
    dataDate = CDate(ws.Range("X2").Value)
    If Err.Number <> 0 Then dataDate = Date
    On Error GoTo 0

    SetupPrintLayout ws, dataDate

    MsgBox "Print setup configured:" & vbCrLf & _
           " - Landscape orientation" & vbCrLf & _
           " - Fit to 1 page wide" & vbCrLf & _
           " - Row 1 repeated as title" & vbCrLf & _
           " - Data Date in footer", vbInformation, "Protocol DEE"
End Sub
