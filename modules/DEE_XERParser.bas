Attribute VB_Name = "DEE_XERParser"
Option Explicit

' =============================================================================
' DEE_XERParser.bas -- XER File Parser
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Parses Primavera P6 XER files into in-memory dictionaries.
' XER format: tab-delimited text with %T (table), %F (fields), %R (rows), %E (end).
'
' PERFORMANCE: Uses bulk 2D array writes (100x faster than cell-by-cell).
' LOCALE SAFETY: No CDate() usage. All dates via SafeParseDate().
' =============================================================================

' ---------------------------------------------------------------------------
' Module-level globals
' ---------------------------------------------------------------------------
Public g_XERTables As Object       ' Scripting.Dictionary: tableName -> Collection of row arrays (Variant())
Public g_XERFields As Object       ' Scripting.Dictionary: tableName -> Variant (String array of field names)
Public g_XERTableOrder As Collection ' Ordered list of table names encountered
Public g_XERFilePath As String

' ---------------------------------------------------------------------------
' ParseXERFile -- Main entry point
' Returns True on success, False on failure
' ---------------------------------------------------------------------------
Public Function ParseXERFile(filePath As String) As Boolean
    ParseXERFile = False
    On Error GoTo ParseFail

    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbCritical, "Protocol DEE"
        Exit Function
    End If

    ' Initialize globals
    Set g_XERTables = CreateObject("Scripting.Dictionary")
    Set g_XERFields = CreateObject("Scripting.Dictionary")
    Set g_XERTableOrder = New Collection
    g_XERFilePath = filePath

    DEE_Utils.StartProgress "Parsing XER file"

    Dim fileNum As Integer
    fileNum = FreeFile()
    Open filePath For Input As #fileNum

    Dim currentTable As String
    Dim currentFields As Variant
    Dim lineText As String
    Dim parts() As String
    Dim lineCount As Long
    lineCount = 0

    currentTable = ""

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineCount = lineCount + 1

        If lineCount Mod 5000 = 0 Then
            DEE_Utils.UpdateProgress 50, "Parsing XER: " & lineCount & " lines"
        End If

        ' Skip empty lines
        If Len(Trim(lineText)) = 0 Then GoTo NextLine

        ' Split by tab
        parts = Split(lineText, vbTab)

        If UBound(parts) < 0 Then GoTo NextLine

        Dim marker As String
        marker = Trim(parts(0))

        Select Case marker
            Case "%T"
                ' Table declaration
                If UBound(parts) >= 1 Then
                    currentTable = Trim(parts(1))
                    If Not g_XERTables.Exists(currentTable) Then
                        g_XERTables.Add currentTable, New Collection
                        g_XERTableOrder.Add currentTable
                    End If
                End If

            Case "%F"
                ' Field names row
                ' parts(0) = "%F", parts(1..n) = field names
                ' Note: skip parts(0) which is the marker
                If currentTable <> "" And UBound(parts) >= 1 Then
                    Dim fieldCount As Integer
                    fieldCount = UBound(parts) ' number of fields = UBound (parts(0) is marker)
                    ReDim currentFields(0 To fieldCount - 1)
                    Dim fi As Integer
                    For fi = 1 To UBound(parts)
                        currentFields(fi - 1) = Trim(parts(fi))
                    Next fi
                    If g_XERFields.Exists(currentTable) Then
                        g_XERFields(currentTable) = currentFields
                    Else
                        g_XERFields.Add currentTable, currentFields
                    End If
                End If

            Case "%R"
                ' Data row
                ' parts(0) = "%R", parts(1..n) = values
                If currentTable <> "" Then
                    Dim rowCount As Integer
                    rowCount = UBound(parts)
                    If rowCount >= 1 Then
                        ReDim rowArr(0 To rowCount - 1) As Variant
                        Dim ri As Integer
                        For ri = 1 To UBound(parts)
                            rowArr(ri - 1) = parts(ri)
                        Next ri
                        g_XERTables(currentTable).Add rowArr
                    End If
                End If

            Case "%E"
                ' End of file
                Exit Do

            Case "ERMHDR"
                ' Header line -- skip (file metadata)

        End Select

NextLine:
    Loop

    Close #fileNum

    DEE_Utils.EndProgress
    ParseXERFile = True
    Exit Function

ParseFail:
    On Error Resume Next
    Close #fileNum
    DEE_Utils.EndProgress
    MsgBox "Error parsing XER file: " & Err.Description, vbCritical, "Protocol DEE"
    ParseXERFile = False
End Function

' ---------------------------------------------------------------------------
' GetTable -- Returns collection of row arrays for a table
' ---------------------------------------------------------------------------
Public Function GetTable(tableName As String) As Collection
    If g_XERTables Is Nothing Then Set GetTable = Nothing: Exit Function
    If g_XERTables.Exists(tableName) Then
        Set GetTable = g_XERTables(tableName)
    Else
        Set GetTable = Nothing
    End If
End Function

' ---------------------------------------------------------------------------
' GetFieldNames -- Returns Variant array of field name strings
' ---------------------------------------------------------------------------
Public Function GetFieldNames(tableName As String) As Variant
    If g_XERFields Is Nothing Then GetFieldNames = Array(): Exit Function
    If g_XERFields.Exists(tableName) Then
        GetFieldNames = g_XERFields(tableName)
    Else
        GetFieldNames = Array()
    End If
End Function

' ---------------------------------------------------------------------------
' FindFieldIndex -- Returns 0-based index of fieldName in fields array
' Returns -1 if not found
' ---------------------------------------------------------------------------
Public Function FindFieldIndex(fields As Variant, fieldName As String) As Integer
    FindFieldIndex = -1
    If IsEmpty(fields) Then Exit Function
    Dim i As Integer
    For i = 0 To UBound(fields)
        If LCase(Trim(fields(i))) = LCase(Trim(fieldName)) Then
            FindFieldIndex = i
            Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------------------
' TableExists -- True if table was found in XER
' ---------------------------------------------------------------------------
Public Function TableExists(tableName As String) As Boolean
    If g_XERTables Is Nothing Then TableExists = False: Exit Function
    TableExists = g_XERTables.Exists(tableName)
End Function

' ---------------------------------------------------------------------------
' IsXERParsed -- True if a file has been loaded
' ---------------------------------------------------------------------------
Public Function IsXERParsed() As Boolean
    IsXERParsed = (Not g_XERTables Is Nothing) And (g_XERTables.Count > 0)
End Function

' ---------------------------------------------------------------------------
' GetXERFilePath -- Opens file dialog; returns empty string if cancelled
' ---------------------------------------------------------------------------
Public Function GetXERFilePath() As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Primavera P6 XER File"
        .Filters.Clear
        .Filters.Add "XER Files", "*.xer"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
        .InitialFileName = ""
        If .Show = -1 Then
            GetXERFilePath = .SelectedItems(1)
        Else
            GetXERFilePath = ""
        End If
    End With
End Function

' ---------------------------------------------------------------------------
' ImportXERToSheet -- Writes a table to a worksheet using bulk array write
' ---------------------------------------------------------------------------
Public Sub ImportXERToSheet(ws As Worksheet, tableName As String)
    If Not TableExists(tableName) Then
        MsgBox "Table '" & tableName & "' not found in XER.", vbExclamation, "Protocol DEE"
        Exit Sub
    End If

    Dim fields As Variant
    fields = GetFieldNames(tableName)
    Dim rows As Collection
    Set rows = GetTable(tableName)

    ws.Cells.Clear

    Dim numFields As Integer
    numFields = UBound(fields) + 1
    Dim numRows As Long
    numRows = rows.Count

    ' Write header row
    Dim headerArr() As Variant
    ReDim headerArr(0, 0 To numFields - 1)
    Dim f As Integer
    For f = 0 To numFields - 1
        headerArr(0, f) = fields(f)
    Next f
    ws.Range(ws.Cells(1, 1), ws.Cells(1, numFields)).Value = headerArr
    DEE_Utils.FormatHeaderRow ws, RGB(0, 32, 96), RGB(255, 255, 255), 1

    If numRows = 0 Then Exit Sub

    ' Build bulk data array
    Dim dataArr() As Variant
    ReDim dataArr(0 To numRows - 1, 0 To numFields - 1)

    Dim rowObj As Variant
    Dim i As Long
    i = 0
    For Each rowObj In rows
        Dim rowData As Variant
        rowData = rowObj
        Dim j As Integer
        For j = 0 To numFields - 1
            If j <= UBound(rowData) Then
                dataArr(i, j) = rowData(j)
            Else
                dataArr(i, j) = ""
            End If
        Next j
        i = i + 1
    Next rowObj

    ' Bulk write
    ws.Range(ws.Cells(2, 1), ws.Cells(numRows + 1, numFields)).Value = dataArr

    DEE_Utils.AutoFitColumns ws
End Sub

' ---------------------------------------------------------------------------
' GetProjectList -- Returns array of project info from PROJECT table
' Returns array of arrays: each sub-array = [proj_id, proj_short_name]
' ---------------------------------------------------------------------------
Public Function GetProjectList() As Variant
    If Not TableExists("PROJECT") Then
        GetProjectList = Array()
        Exit Function
    End If

    Dim fields As Variant
    fields = GetFieldNames("PROJECT")
    Dim idIdx As Integer
    Dim nameIdx As Integer
    idIdx = FindFieldIndex(fields, "proj_id")
    nameIdx = FindFieldIndex(fields, "proj_short_name")

    Dim rows As Collection
    Set rows = GetTable("PROJECT")

    If rows.Count = 0 Then
        GetProjectList = Array()
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(0 To rows.Count - 1)

    Dim i As Long
    Dim rowObj As Variant
    i = 0
    For Each rowObj In rows
        Dim rd As Variant
        rd = rowObj
        Dim projId As String
        Dim projName As String
        If idIdx >= 0 And idIdx <= UBound(rd) Then projId = rd(idIdx) Else projId = ""
        If nameIdx >= 0 And nameIdx <= UBound(rd) Then projName = rd(nameIdx) Else projName = ""
        result(i) = Array(projId, projName)
        i = i + 1
    Next rowObj

    GetProjectList = result
End Function

' ---------------------------------------------------------------------------
' GetAvailableBudgetFields -- Scans TASK fields for quantity/budget columns
' ---------------------------------------------------------------------------
Public Function GetAvailableBudgetFields() As Variant
    Dim knownBudgetFields() As String
    knownBudgetFields = Split("target_work_qty,target_equip_qty,budget_qty,target_cost,act_work_qty,remain_work_qty", ",")

    Dim fields As Variant
    fields = GetFieldNames("TASK")

    Dim found() As String
    Dim foundCount As Integer
    ReDim found(0 To UBound(knownBudgetFields))
    foundCount = 0

    Dim kf As Integer
    For kf = 0 To UBound(knownBudgetFields)
        If FindFieldIndex(fields, knownBudgetFields(kf)) >= 0 Then
            found(foundCount) = knownBudgetFields(kf)
            foundCount = foundCount + 1
        End If
    Next kf

    If foundCount = 0 Then
        GetAvailableBudgetFields = Array()
    Else
        ReDim result(0 To foundCount - 1) As String
        Dim i As Integer
        For i = 0 To foundCount - 1
            result(i) = found(i)
        Next i
        GetAvailableBudgetFields = result
    End If
End Function

' ---------------------------------------------------------------------------
' Private: FileExists helper
' ---------------------------------------------------------------------------
Private Function FileExists(path As String) As Boolean
    FileExists = (Dir(path) <> "")
End Function
