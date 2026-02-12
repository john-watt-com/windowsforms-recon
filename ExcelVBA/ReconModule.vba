Option Explicit

' Main reconciliation module
' Ported from C# WindowsForms ExcelReconApp

Public Sub RunReconciliation()
    ' Variable declarations
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim wsResults As Worksheet
    Dim wsConfig As Worksheet
    Dim idCol1 As String
    Dim idCol2 As String
    Dim valueColumns1 As String
    Dim valueColumns2 As String
    Dim additionalColumns1 As String
    Dim additionalColumns2 As String
    Dim tolerance As Double
    Dim detailSheet As String
    Dim dict1 As Object
    Dim dict2 As Object
    Dim dictAdditional1 As Object
    Dim dictAdditional2 As Object
    Dim dictDetailRows As Object
    Dim allIDs As Collection
    
    ' Validate setup
    If Not ValidateSetup() Then Exit Sub
    
    ' Clear previous results
    ClearResults
    
    ' Get configuration
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Sheet2")
    Set wsResults = GetOrCreateWorksheet("All Results")
    Set wsConfig = ThisWorkbook.Worksheets("Recon Config")
    
    ' Get configuration settings from table
    idCol1 = GetConfigValue(wsConfig, "Sheet1 ID Column")
    idCol2 = GetConfigValue(wsConfig, "Sheet2 ID Column")
    valueColumns1 = GetConfigValue(wsConfig, "Sheet1 Value Columns")
    valueColumns2 = GetConfigValue(wsConfig, "Sheet2 Value Columns")
    additionalColumns1 = GetConfigValue(wsConfig, "Sheet1 Additional Columns")
    
    ' Get tolerance (default to 0.01 if not specified or invalid)
    Dim toleranceStr As String
    toleranceStr = GetConfigValue(wsConfig, "Tolerance")
    If IsNumeric(toleranceStr) And CDbl(toleranceStr) >= 0 Then
        tolerance = CDbl(toleranceStr)
    Else
        tolerance = 0.01
    End If
    
    ' Get detail sheet setting and additional columns from second sheet
    detailSheet = Trim(UCase(GetConfigValue(wsConfig, "Detail Sheet")))
    additionalColumns2 = GetConfigValue(wsConfig, "Sheet2 Additional Columns")
    
    ' Get sort order setting
    Dim sortOrder As String
    sortOrder = GetConfigValue(wsConfig, "Sort Order")
    
    ' Build dictionaries of ID -> Total
    Set dict1 = BuildTotalsDictionary(ws1, idCol1, valueColumns1)
    Set dict2 = BuildTotalsDictionary(ws2, idCol2, valueColumns2)
    
    ' Build dictionaries of ID -> additional column values (from both sheets)
    Set dictAdditional1 = BuildAdditionalColumnsDictionary(ws1, idCol1, additionalColumns1)
    Set dictAdditional2 = BuildAdditionalColumnsDictionary(ws2, idCol2, additionalColumns2)
    
    ' Build detail rows dictionary if detail mode is enabled
    Set dictDetailRows = CreateObject("Scripting.Dictionary")
    If detailSheet = "SHEET1" Then
        Set dictDetailRows = BuildAllDetailRowsDictionary(ws1, idCol1, additionalColumns1)
    ElseIf detailSheet = "SHEET2" Then
        Set dictDetailRows = BuildAllDetailRowsDictionary(ws2, idCol2, additionalColumns2)
    End If
    
    ' Get all unique IDs from both sheets
    Set allIDs = GetUniqueIDs(dict1, dict2)
    
    ' Write results
    WriteResults wsResults, allIDs, dict1, dict2, dictAdditional1, additionalColumns1, tolerance, sortOrder
    
    ' Split results into Match Results and Error Results sheets
    SplitResults wsResults, ws1, ws2, idCol1, idCol2, detailSheet, dictDetailRows, dict1, dict2, dictAdditional1, dictAdditional2, additionalColumns1, additionalColumns2
    
    MsgBox "Reconciliation completed! " & allIDs.Count & " records processed.", vbInformation, "Success"
End Sub

Private Function ValidateSetup() As Boolean
    ValidateSetup = False
    
    ' Check worksheets exist
    On Error Resume Next
    Dim ws1 As Worksheet, ws2 As Worksheet, wsConfig As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Sheet2")
    Set wsConfig = ThisWorkbook.Worksheets("Recon Config")
    On Error GoTo 0
    
    If ws1 Is Nothing Or ws2 Is Nothing Or wsConfig Is Nothing Then
        MsgBox "Required worksheets not found. Ensure 'Sheet1', 'Sheet2', and 'Recon Config' exist.", vbCritical, "Error"
        Exit Function
    End If
    
    ' Check that ReconConfigTable exists
    Dim tblConfig As ListObject
    On Error Resume Next
    Set tblConfig = wsConfig.ListObjects("ReconConfigTable")
    On Error GoTo 0
    
    If tblConfig Is Nothing Then
        MsgBox "ReconConfigTable not found in Recon Config sheet. Please convert your configuration to a table named 'ReconConfigTable'.", vbCritical, "Error"
        Exit Function
    End If
    
    ' Check required configuration settings
    Dim idCol1 As String, idCol2 As String
    Dim valueColumns1 As String, valueColumns2 As String
    
    idCol1 = GetConfigValue(wsConfig, "Sheet1 ID Column")
    idCol2 = GetConfigValue(wsConfig, "Sheet2 ID Column")
    valueColumns1 = GetConfigValue(wsConfig, "Sheet1 Value Columns")
    valueColumns2 = GetConfigValue(wsConfig, "Sheet2 Value Columns")
    
    If Len(idCol1) = 0 Or Len(idCol2) = 0 Then
        MsgBox "Please specify both ID columns in Recon Config table.", vbCritical, "Error"
        Exit Function
    End If
    
    If Len(valueColumns1) = 0 Or Len(valueColumns2) = 0 Then
        MsgBox "Please specify value columns for both sheets in Recon Config table.", vbCritical, "Error"
        Exit Function
    End If
    
    ' Check data exists
    If ws1.Cells(2, 1).value = "" Or ws2.Cells(2, 1).value = "" Then
        MsgBox "Please load data into Sheet1 and Sheet2.", vbCritical, "Error"
        Exit Function
    End If
    
    ValidateSetup = True
End Function

Private Function BuildTotalsDictionary(ws As Worksheet, idColumnName As String, valueColumnsStr As String) As Object
    ' Returns Dictionary with ID as key and sum of value columns as value
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Find column indexes
    Dim idColIndex As Long
    idColIndex = FindColumnIndex(ws, idColumnName)
    If idColIndex = 0 Then
        MsgBox "Column '" & idColumnName & "' not found in " & ws.Name, vbCritical, "Error"
        Set BuildTotalsDictionary = dict
        Exit Function
    End If
    
    ' Parse value column names
    Dim valueColNames() As String
    valueColNames = Split(valueColumnsStr, ",")
    
    ' Find value column indexes and operations (+ or -)
    Dim valueColIndexes() As Long
    Dim valueColOperations() As String
    ReDim valueColIndexes(0 To UBound(valueColNames))
    ReDim valueColOperations(0 To UBound(valueColNames))
    Dim i As Long
    Dim colName As String
    Dim operation As String
    
    For i = 0 To UBound(valueColNames)
        colName = Trim(valueColNames(i))
        
        ' Check for leading + or - sign
        If Left(colName, 1) = "+" Or Left(colName, 1) = "-" Then
            operation = Left(colName, 1)
            colName = Trim(Mid(colName, 2))  ' Remove the sign
        Else
            operation = "+"  ' Default to addition for backward compatibility
        End If
        
        valueColNames(i) = colName
        valueColOperations(i) = operation
        valueColIndexes(i) = FindColumnIndex(ws, colName)
        
        If valueColIndexes(i) = 0 Then
            MsgBox "Column '" & colName & "' not found in " & ws.Name, vbCritical, "Error"
            Set BuildTotalsDictionary = dict
            Exit Function
        End If
    Next i
    
    ' Loop through data rows (starting from row 2)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, idColIndex).End(xlUp).row
    
    Dim row As Long, id As String, total As Double, value As Double
    For row = 2 To lastRow
        id = Trim(CStr(ws.Cells(row, idColIndex).value))
        
        ' Skip blank IDs
        If Len(id) = 0 Then GoTo NextRow
        
        ' Sum value columns for this row (applying + or - operations)
        total = 0
        For i = 0 To UBound(valueColIndexes)
            If IsNumeric(ws.Cells(row, valueColIndexes(i)).value) Then
                value = CDbl(ws.Cells(row, valueColIndexes(i)).value)
                If valueColOperations(i) = "-" Then
                    total = total - value
                Else
                    total = total + value
                End If
            End If
        Next i
        
        ' Add or update dictionary
        If dict.Exists(id) Then
            dict(id) = dict(id) + total
        Else
            dict.Add id, total
        End If
        
NextRow:
    Next row
    
    Set BuildTotalsDictionary = dict
End Function

Private Function BuildAdditionalColumnsDictionary(ws As Worksheet, idColumnName As String, additionalColumnsStr As String) As Object
    ' Returns Dictionary with ID as key and Dictionary of additional column values as value
    ' Only captures first occurrence of each ID
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' If no additional columns specified, return empty dictionary
    If Len(Trim(additionalColumnsStr)) = 0 Then
        Set BuildAdditionalColumnsDictionary = dict
        Exit Function
    End If
    
    ' Find ID column index
    Dim idColIndex As Long
    idColIndex = FindColumnIndex(ws, idColumnName)
    If idColIndex = 0 Then
        Set BuildAdditionalColumnsDictionary = dict
        Exit Function
    End If
    
    ' Parse additional column names
    Dim additionalColNames() As String
    additionalColNames = Split(additionalColumnsStr, ",")
    
    ' Find additional column indexes
    Dim additionalColIndexes() As Long
    ReDim additionalColIndexes(0 To UBound(additionalColNames))
    Dim i As Long
    For i = 0 To UBound(additionalColNames)
        additionalColNames(i) = Trim(additionalColNames(i))
        additionalColIndexes(i) = FindColumnIndex(ws, additionalColNames(i))
        If additionalColIndexes(i) = 0 Then
            MsgBox "Additional column '" & additionalColNames(i) & "' not found in " & ws.Name, vbCritical, "Error"
            Set BuildAdditionalColumnsDictionary = dict
            Exit Function
        End If
    Next i
    
    ' Loop through data rows (starting from row 2)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, idColIndex).End(xlUp).row
    
    Dim row As Long, id As String
    Dim additionalDict As Object
    
    For row = 2 To lastRow
        id = Trim(CStr(ws.Cells(row, idColIndex).value))
        
        ' Skip blank IDs
        If Len(id) = 0 Then GoTo NextRowAdditional
        
        ' Only store first occurrence of each ID
        If Not dict.Exists(id) Then
            Set additionalDict = CreateObject("Scripting.Dictionary")
            
            ' Store values from additional columns
            For i = 0 To UBound(additionalColIndexes)
                additionalDict.Add additionalColNames(i), ws.Cells(row, additionalColIndexes(i)).value
            Next i
            
            dict.Add id, additionalDict
        End If
        
NextRowAdditional:
    Next row
    
    Set BuildAdditionalColumnsDictionary = dict
End Function

Private Function BuildAllDetailRowsDictionary(ws As Worksheet, idColumnName As String, additionalColumnsStr As String) As Object
    ' Returns Dictionary with ID as key and Collection of all row dictionaries as value
    ' Each row dictionary contains column name -> value pairs
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' If no additional columns specified, return empty dictionary
    If Len(Trim(additionalColumnsStr)) = 0 Then
        Set BuildAllDetailRowsDictionary = dict
        Exit Function
    End If
    
    ' Find ID column index
    Dim idColIndex As Long
    idColIndex = FindColumnIndex(ws, idColumnName)
    If idColIndex = 0 Then
        Set BuildAllDetailRowsDictionary = dict
        Exit Function
    End If
    
    ' Parse additional column names
    Dim additionalColNames() As String
    additionalColNames = Split(additionalColumnsStr, ",")
    
    ' Find additional column indexes
    Dim additionalColIndexes() As Long
    ReDim additionalColIndexes(0 To UBound(additionalColNames))
    Dim i As Long
    For i = 0 To UBound(additionalColNames)
        additionalColNames(i) = Trim(additionalColNames(i))
        additionalColIndexes(i) = FindColumnIndex(ws, additionalColNames(i))
        If additionalColIndexes(i) = 0 Then
            MsgBox "Additional column '" & additionalColNames(i) & "' not found in " & ws.Name, vbCritical, "Error"
            Set BuildAllDetailRowsDictionary = dict
            Exit Function
        End If
    Next i
    
    ' Loop through data rows (starting from row 2)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, idColIndex).End(xlUp).row
    
    Dim row As Long, id As String
    Dim rowDict As Object
    Dim rowCollection As Collection
    
    For row = 2 To lastRow
        id = Trim(CStr(ws.Cells(row, idColIndex).value))
        
        ' Skip blank IDs
        If Len(id) = 0 Then GoTo NextRowAll
        
        ' Create dictionary for this row
        Set rowDict = CreateObject("Scripting.Dictionary")
        For i = 0 To UBound(additionalColIndexes)
            rowDict.Add additionalColNames(i), ws.Cells(row, additionalColIndexes(i)).value
        Next i
        
        ' Add to collection for this ID
        If Not dict.Exists(id) Then
            Set rowCollection = New Collection
            dict.Add id, rowCollection
        Else
            Set rowCollection = dict(id)
        End If
        rowCollection.Add rowDict
        
NextRowAll:
    Next row
    
    Set BuildAllDetailRowsDictionary = dict
End Function

Private Function FindColumnIndex(ws As Worksheet, columnName As String) As Long
    ' Find column index by header name in row 1
    Dim col As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If Trim(CStr(ws.Cells(1, col).value)) = Trim(columnName) Then
            FindColumnIndex = col
            Exit Function
        End If
    Next col
    
    FindColumnIndex = 0
End Function

Private Function GetConfigValue(wsConfig As Worksheet, settingName As String) As String
    ' Get configuration value from ReconConfigTable by setting name (case-insensitive)
    GetConfigValue = ""
    
    ' Get the config table
    Dim tblConfig As ListObject
    On Error Resume Next
    Set tblConfig = wsConfig.ListObjects("ReconConfigTable")
    On Error GoTo 0
    
    If tblConfig Is Nothing Then Exit Function
    
    ' Find Setting and Value column indexes
    Dim settingColIndex As Long, valueColIndex As Long
    Dim col As Long
    
    For col = 1 To tblConfig.ListColumns.Count
        If UCase(Trim(tblConfig.HeaderRowRange.Cells(1, col).value)) = "SETTING" Then
            settingColIndex = col
        ElseIf UCase(Trim(tblConfig.HeaderRowRange.Cells(1, col).value)) = "VALUE" Then
            valueColIndex = col
        End If
    Next col
    
    If settingColIndex = 0 Or valueColIndex = 0 Then Exit Function
    
    ' Search for the setting (case-insensitive)
    Dim row As Long
    Dim settingNameUpper As String
    settingNameUpper = UCase(Trim(settingName))
    
    For row = 1 To tblConfig.ListRows.Count
        If UCase(Trim(CStr(tblConfig.DataBodyRange.Cells(row, settingColIndex).value))) = settingNameUpper Then
            GetConfigValue = Trim(CStr(tblConfig.DataBodyRange.Cells(row, valueColIndex).value))
            Exit Function
        End If
    Next row
End Function

Private Function GetUniqueIDs(dict1 As Object, dict2 As Object) As Collection
    ' Combine all unique IDs from both dictionaries
    Dim allIDs As New Collection
    Dim key As Variant
    
    On Error Resume Next ' Ignore duplicates
    For Each key In dict1.Keys
        allIDs.Add key, CStr(key)
    Next key
    
    For Each key In dict2.Keys
        allIDs.Add key, CStr(key)
    Next key
    On Error GoTo 0
    
    Set GetUniqueIDs = allIDs
End Function

Private Sub WriteResults(wsResults As Worksheet, allIDs As Collection, dict1 As Object, dict2 As Object, dictAdditional As Object, additionalColumnsStr As String, tolerance As Double, sortOrder As String)
    ' Parse additional column names
    Dim additionalColNames() As String
    Dim hasAdditionalCols As Boolean
    hasAdditionalCols = (Len(Trim(additionalColumnsStr)) > 0)
    
    If hasAdditionalCols Then
        additionalColNames = Split(additionalColumnsStr, ",")
        Dim i As Long
        For i = 0 To UBound(additionalColNames)
            additionalColNames(i) = Trim(additionalColNames(i))
        Next i
    End If
    
    ' Write header - new order: IsMatch, Difference, ID, Sheet1 Total, Sheet2 Total, then additional columns
    Dim col As Long
    col = 1
    wsResults.Cells(1, col).value = "IsMatch": col = col + 1
    wsResults.Cells(1, col).value = "Difference": col = col + 1
    wsResults.Cells(1, col).value = "ID": col = col + 1
    wsResults.Cells(1, col).value = "Sheet1 Total": col = col + 1
    wsResults.Cells(1, col).value = "Sheet2 Total": col = col + 1
    
    ' Add additional column headers
    If hasAdditionalCols Then
        For i = 0 To UBound(additionalColNames)
            wsResults.Cells(1, col).value = additionalColNames(i)
            col = col + 1
        Next i
    End If
    
    Dim lastHeaderCol As Long
    lastHeaderCol = col - 1
    
    ' Format header
    With wsResults.Range(wsResults.Cells(1, 1), wsResults.Cells(1, lastHeaderCol))
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Build and sort row data
    Dim rowData As Collection
    Set rowData = New Collection
    
    Dim id As String
    Dim total1 As Double, total2 As Double, diff As Double
    Dim isMatch As Boolean
    Dim additionalDict As Object
    Dim rowDict As Object
    Dim j As Long
    
    ' Build all row data
    For i = 1 To allIDs.Count
        id = allIDs(i)
        
        ' Get totals (default to 0 if not exists)
        If dict1.Exists(id) Then
            total1 = dict1(id)
        Else
            total1 = 0
        End If
        
        If dict2.Exists(id) Then
            total2 = dict2(id)
        Else
            total2 = 0
        End If
        
        diff = total1 - total2
        isMatch = (Abs(diff) < tolerance)
        
        ' Create row dictionary
        Set rowDict = CreateObject("Scripting.Dictionary")
        rowDict("IsMatch") = isMatch
        rowDict("Difference") = diff
        rowDict("ID") = id
        rowDict("Sheet1 Total") = total1
        rowDict("Sheet2 Total") = total2
        
        ' Add additional columns
        If hasAdditionalCols And dictAdditional.Exists(id) Then
            Set additionalDict = dictAdditional(id)
            For j = 0 To UBound(additionalColNames)
                If additionalDict.Exists(additionalColNames(j)) Then
                    rowDict(additionalColNames(j)) = additionalDict(additionalColNames(j))
                Else
                    rowDict(additionalColNames(j)) = ""
                End If
            Next j
        ElseIf hasAdditionalCols Then
            For j = 0 To UBound(additionalColNames)
                rowDict(additionalColNames(j)) = ""
            Next j
        End If
        
        rowData.Add rowDict
    Next i
    
    ' Sort row data by specified columns
    SortRowData rowData, sortOrder, wsResults
    
    ' Write sorted data rows
    Dim row As Long
    row = 2
    For i = 1 To rowData.Count
        Set rowDict = rowData(i)
        
        ' Write to sheet from row dictionary
        col = 1
        wsResults.Cells(row, col).value = rowDict("IsMatch"): col = col + 1
        wsResults.Cells(row, col).value = rowDict("Difference"): col = col + 1
        wsResults.Cells(row, col).value = rowDict("ID"): col = col + 1
        wsResults.Cells(row, col).value = rowDict("Sheet1 Total"): col = col + 1
        wsResults.Cells(row, col).value = rowDict("Sheet2 Total"): col = col + 1
        
        ' Write additional columns
        If hasAdditionalCols Then
            For j = 0 To UBound(additionalColNames)
                wsResults.Cells(row, col).value = rowDict(additionalColNames(j))
                col = col + 1
            Next j
        End If
        
        ' Highlight non-matches
        If Not rowDict("IsMatch") Then
            wsResults.Rows(row).Interior.Color = RGB(255, 200, 200)
        End If
        
        row = row + 1
    Next i
    
    ' Format number columns (Difference, Sheet1 Total, Sheet2 Total)
    wsResults.Range(wsResults.Cells(2, 2), wsResults.Cells(row - 1, 5)).NumberFormat = "#,##0.00"
    
    ' Convert to Table
    Dim lastRow As Long
    lastRow = row - 1
    If lastRow >= 2 Then
        Dim tableRange As Range
        Set tableRange = wsResults.Range(wsResults.Cells(1, 1), wsResults.Cells(lastRow, lastHeaderCol))
        
        ' Delete existing table if present
        On Error Resume Next
        wsResults.ListObjects("ReconResultsTable").Delete
        On Error GoTo 0
        
        ' Create new table
        wsResults.ListObjects.Add(xlSrcRange, tableRange, , xlYes).Name = "ReconResultsTable"
        
        ' Apply table style
        wsResults.ListObjects("ReconResultsTable").TableStyle = "TableStyleMedium2"
    End If
    
    ' Auto-fit columns
    wsResults.Columns("A:" & ColLetter(lastHeaderCol)).AutoFit
End Sub

Private Function ColLetter(colNum As Long) As String
    ' Convert column number to letter
    Dim letter As String
    Dim num As Long
    num = colNum
    Do While num > 0
        letter = Chr((num - 1) Mod 26 + 65) & letter
        num = (num - 1) \ 26
    Loop
    ColLetter = letter
End Function

Private Sub ClearResults()
    ' Clear All Results sheet
    Dim wsResults As Worksheet
    Set wsResults = GetOrCreateWorksheet("All Results")
    
    ' Delete existing table if present
    On Error Resume Next
    wsResults.ListObjects("ReconResultsTable").Delete
    On Error GoTo 0
    
    ' Clear all cells
    wsResults.Cells.Clear
    
    ' Clear Match Results sheet
    Dim wsMatch As Worksheet
    Set wsMatch = GetOrCreateWorksheet("Match Results")
    On Error Resume Next
    wsMatch.ListObjects("MatchResultsTable").Delete
    On Error GoTo 0
    wsMatch.Cells.Clear
    
    ' Clear Error Results sheet
    Dim wsError As Worksheet
    Set wsError = GetOrCreateWorksheet("Error Results")
    On Error Resume Next
    wsError.ListObjects("ErrorResultsTable").Delete
    On Error GoTo 0
    wsError.Cells.Clear
End Sub

Private Sub SortRowData(rowData As Collection, sortOrder As String, wsResults As Worksheet)
    ' Sort row data by specified columns
    ' sortOrder format: "Column1,Column2 DESC,Column3" (can include ASC/DESC, default is ASC)
    
    ' If no sort order specified, sort by ID ascending (default)
    If Len(Trim(sortOrder)) = 0 Then
        sortOrder = "ID"
    End If
    
    ' Parse sort columns and directions
    Dim sortSpecs() As String
    sortSpecs = Split(sortOrder, ",")
    
    Dim sortColumns() As String
    Dim sortDirections() As String
    ReDim sortColumns(0 To UBound(sortSpecs))
    ReDim sortDirections(0 To UBound(sortSpecs))
    
    Dim i As Long, spec As String, parts() As String
    For i = 0 To UBound(sortSpecs)
        spec = Trim(sortSpecs(i))
        
        ' Check for DESC keyword
        If UCase(Right(spec, 5)) = " DESC" Then
            sortColumns(i) = Trim(Left(spec, Len(spec) - 5))
            sortDirections(i) = "DESC"
        ElseIf UCase(Right(spec, 4)) = " ASC" Then
            sortColumns(i) = Trim(Left(spec, Len(spec) - 4))
            sortDirections(i) = "ASC"
        Else
            sortColumns(i) = spec
            sortDirections(i) = "ASC" ' Default
        End If
    Next i
    
    ' Bubble sort (simple, works for typical result set sizes)
    Dim swapped As Boolean, j As Long
    Dim row1 As Object, row2 As Object
    Dim compareResult As Integer
    
    Do
        swapped = False
        For i = 1 To rowData.Count - 1
            Set row1 = rowData(i)
            Set row2 = rowData(i + 1)
            
            ' Compare by sort columns in order
            compareResult = CompareRows(row1, row2, sortColumns, sortDirections)
            
            If compareResult > 0 Then
                ' Swap
                rowData.Remove i + 1
                rowData.Add row2, , i
                swapped = True
            End If
        Next i
    Loop While swapped
End Sub

Private Function CompareRows(row1 As Object, row2 As Object, sortColumns() As String, sortDirections() As String) As Integer
    ' Compare two row dictionaries by multiple columns
    ' Returns: -1 if row1 < row2, 0 if equal, 1 if row1 > row2
    
    Dim i As Long, colName As String
    Dim val1 As Variant, val2 As Variant
    Dim compareResult As Integer
    
    For i = 0 To UBound(sortColumns)
        colName = sortColumns(i)
        
        ' Get values (handle missing keys)
        If row1.Exists(colName) Then
            val1 = row1(colName)
        Else
            val1 = ""
        End If
        
        If row2.Exists(colName) Then
            val2 = row2(colName)
        Else
            val2 = ""
        End If
        
        ' Compare values
        If IsNumeric(val1) And IsNumeric(val2) Then
            ' Numeric comparison
            If CDbl(val1) < CDbl(val2) Then
                compareResult = -1
            ElseIf CDbl(val1) > CDbl(val2) Then
                compareResult = 1
            Else
                compareResult = 0
            End If
        Else
            ' String comparison
            If CStr(val1) < CStr(val2) Then
                compareResult = -1
            ElseIf CStr(val1) > CStr(val2) Then
                compareResult = 1
            Else
                compareResult = 0
            End If
        End If
        
        ' Apply sort direction
        If sortDirections(i) = "DESC" Then
            compareResult = -compareResult
        End If
        
        ' If not equal, return result; otherwise continue to next sort column
        If compareResult <> 0 Then
            CompareRows = compareResult
            Exit Function
        End If
    Next i
    
    ' All columns equal
    CompareRows = 0
End Function

' Helper: Import external Excel file into a worksheet
Public Sub ImportExcelFile(targetSheet As Worksheet)
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.Title = "Select Excel File to Import"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
    
    If fd.Show = -1 Then
        Dim filePath As String
        filePath = fd.SelectedItems(1)
        
        ' Clear target sheet
        targetSheet.Cells.Clear
        
        ' Open file and copy data
        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath, ReadOnly:=True)
        
        ' Copy first worksheet
        wb.Worksheets(1).UsedRange.Copy targetSheet.Range("A1")
        
        wb.Close False
        
        MsgBox "Successfully loaded data into " & targetSheet.Name, vbInformation, "Success"
    End If
End Sub

' Helper: Show the Recon Form
' You can run this macro directly from Excel (Developer -> Macros -> ShowReconForm)
' Or assign it to a button on your worksheet
Public Sub ShowReconForm()
    ReconForm.Show
End Sub

' ========================================
' TRANSFORMATION MODULE
' ========================================

' Main transformation routine - reads Transform Config sheet and creates transformed sheets
Public Sub RunTransform()
    On Error GoTo ErrorHandler
    
    ' Check if Transform Config sheet exists
    Dim wsTransformConfig As Worksheet
    On Error Resume Next
    Set wsTransformConfig = ThisWorkbook.Worksheets("Transform Config")
    On Error GoTo ErrorHandler
    
    If wsTransformConfig Is Nothing Then
        MsgBox "Transform Config sheet not found. Please create it first.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Read transform configuration (starts at row 2)
    Dim transformName As String, sourceSheetName As String, targetSheetName As String
    Dim configRow As Long
    configRow = 2 ' Row with transform name, source, and target
    
    transformName = Trim(CStr(wsTransformConfig.Cells(configRow, 1).value))
    sourceSheetName = Trim(CStr(wsTransformConfig.Cells(configRow, 2).value))
    targetSheetName = Trim(CStr(wsTransformConfig.Cells(configRow, 3).value))
    
    If Len(transformName) = 0 Or Len(sourceSheetName) = 0 Or Len(targetSheetName) = 0 Then
        MsgBox "Transform configuration incomplete. Check Transform Config sheet rows 2 (Name, Source Sheet, Target Sheet).", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Verify source sheet exists
    Dim wsSource As Worksheet
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sourceSheetName)
    On Error GoTo ErrorHandler
    
    If wsSource Is Nothing Then
        MsgBox "Source sheet '" & sourceSheetName & "' not found.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Create or get target sheet
    Dim wsTarget As Worksheet
    Set wsTarget = GetOrCreateWorksheet(targetSheetName)
    wsTarget.Cells.Clear ' Clear existing content
    
    ' Read column definitions (starting from row 5: Order, Target Column, Type, Source)
    Dim columnDefs As Collection
    Set columnDefs = New Collection
    
    Dim row As Long
    row = 5
    Do While Len(Trim(CStr(wsTransformConfig.Cells(row, 1).value))) > 0
        Dim colDef As Object
        Set colDef = CreateObject("Scripting.Dictionary")
        
        colDef("Order") = CLng(wsTransformConfig.Cells(row, 1).value)
        colDef("TargetColumn") = Trim(CStr(wsTransformConfig.Cells(row, 2).value))
        colDef("Type") = Trim(UCase(CStr(wsTransformConfig.Cells(row, 3).value)))
        
        ' Read source as text/value
        Dim sourceValue As String
        sourceValue = Trim(CStr(wsTransformConfig.Cells(row, 4).value))
        
        ' For FORMULA type, ensure it starts with =
        If colDef("Type") = "FORMULA" Then
            If left(sourceValue, 1) <> "=" Then
                sourceValue = "=" & sourceValue
            End If
        End If
        
        colDef("Source") = sourceValue
        
        columnDefs.Add colDef
        row = row + 1
    Loop
    
    If columnDefs.Count = 0 Then
        MsgBox "No column definitions found. Add column definitions starting at row 5.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Apply transformation
    ApplyTransformation wsSource, wsTarget, columnDefs
    
    MsgBox "Transformation completed successfully!" & vbCrLf & _
           "Transform: " & transformName & vbCrLf & _
           "Source: " & sourceSheetName & vbCrLf & _
           "Target: " & targetSheetName, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during transformation: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ApplyTransformation(wsSource As Worksheet, wsTarget As Worksheet, columnDefs As Collection)
    ' Write headers
    Dim colDef As Object
    Dim col As Long
    
    For Each colDef In columnDefs
        col = colDef("Order")
        wsTarget.Cells(1, col).value = colDef("TargetColumn")
    Next colDef
    
    ' Determine source data range
    Dim lastSourceRow As Long
    lastSourceRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).row
    
    If lastSourceRow < 2 Then
        MsgBox "Source sheet has no data rows.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    ' Populate all rows with EXISTING and STATIC columns only
    Dim sourceRow As Long, targetRow As Long
    Dim sourceColIndex As Long
    
    For sourceRow = 2 To lastSourceRow
        targetRow = sourceRow
        
        For Each colDef In columnDefs
            col = colDef("Order")
            
            Select Case colDef("Type")
                Case "EXISTING"
                    ' Copy from source column
                    sourceColIndex = FindColumnIndex(wsSource, colDef("Source"))
                    
                    If sourceColIndex > 0 Then
                        If wsSource.Cells(sourceRow, sourceColIndex).HasFormula Then
                            wsTarget.Cells(targetRow, col).Formula = wsSource.Cells(sourceRow, sourceColIndex).Formula
                        Else
                            wsTarget.Cells(targetRow, col).value = wsSource.Cells(sourceRow, sourceColIndex).value
                        End If
                    End If
                    
                Case "STATIC"
                    ' Static value - same for all rows
                    wsTarget.Cells(targetRow, col).value = colDef("Source")
                    
                Case "FORMULA"
                    ' Skip for now - will add after creating table
                    
            End Select
        Next colDef
    Next sourceRow
    
    ' Create table with the data we have
    Dim lastCol As Long
    lastCol = columnDefs.Count
    Dim lastRow As Long
    lastRow = lastSourceRow
    
    Dim tableRange As Range
    Set tableRange = wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(lastRow, lastCol))
    
    ' Delete existing table if present
    On Error Resume Next
    wsTarget.ListObjects(1).Delete
    On Error GoTo 0
    
    ' Create new table
    Dim tbl As ListObject
    Set tbl = wsTarget.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    tbl.TableStyle = "TableStyleMedium2"
    
    ' Now add FORMULA columns - table will auto-fill them down
    For Each colDef In columnDefs
        If colDef("Type") = "FORMULA" Then
            col = colDef("Order")
            ' Add formula to first data row - table calculated column feature will auto-fill
            wsTarget.Cells(2, col).Formula = colDef("Source")
        End If
    Next colDef
    
    ' Auto-fit columns
    wsTarget.Columns("A:" & ColLetter(lastCol)).AutoFit
End Sub

' Helper: Get worksheet by name or create it if it doesn't exist
Private Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    
    Set GetOrCreateWorksheet = ws
End Function

' Helper: Split results from All Results into Match Results and Error Results sheets
Private Sub SplitResults(wsAllResults As Worksheet, ws1 As Worksheet, ws2 As Worksheet, _
                         idCol1 As String, idCol2 As String, detailSheet As String, _
                         dictDetailRows As Object, dict1 As Object, dict2 As Object, _
                         dictAdditional1 As Object, dictAdditional2 As Object, _
                         additionalColumns1 As String, additionalColumns2 As String)
    ' Check if ReconResultsTable exists
    Dim tblAll As ListObject
    On Error Resume Next
    Set tblAll = wsAllResults.ListObjects("ReconResultsTable")
    On Error GoTo 0
    
    If tblAll Is Nothing Then Exit Sub
    If tblAll.ListRows.Count = 0 Then Exit Sub
    
    ' Find column indexes in All Results
    Dim isMatchCol As Long, diffCol As Long, idCol As Long
    Dim sheet1TotalCol As Long, sheet2TotalCol As Long
    Dim lastCol As Long
    
    isMatchCol = 1  ' Column A
    diffCol = 2     ' Column B
    idCol = 3       ' Column C
    sheet1TotalCol = 4  ' Column D
    sheet2TotalCol = 5  ' Column E
    lastCol = wsAllResults.Cells(1, wsAllResults.Columns.Count).End(xlToLeft).Column
    
    ' Get Match Results and Error Results worksheets
    Dim wsMatch As Worksheet, wsError As Worksheet
    Set wsMatch = GetOrCreateWorksheet("Match Results")
    Set wsError = GetOrCreateWorksheet("Error Results")
    
    ' Determine if we're in detail expansion mode
    Dim isDetailMode As Boolean
    isDetailMode = (detailSheet = "SHEET1" Or detailSheet = "SHEET2")
    
    If isDetailMode Then
        ' Detail expansion mode for Match Results
        WriteDetailMatchResults wsMatch, wsAllResults, ws1, ws2, idCol1, idCol2, detailSheet, _
                                dictDetailRows, dict1, dict2, dictAdditional1, dictAdditional2, _
                                additionalColumns1, additionalColumns2, idCol, sheet1TotalCol, sheet2TotalCol
    Else
        ' Aggregated mode for Match Results (current behavior)
        WriteAggregatedMatchResults wsMatch, wsAllResults, idCol, lastCol
    End If
    
    ' Error Results: always aggregated (keep current behavior)
    WriteErrorResults wsError, wsAllResults, idCol, diffCol, lastCol
End Sub

Private Sub WriteAggregatedMatchResults(wsMatch As Worksheet, wsAllResults As Worksheet, idCol As Long, lastCol As Long)
    ' Current behavior: Copy aggregated match results
    Dim tblAll As ListObject
    Set tblAll = wsAllResults.ListObjects("ReconResultsTable")
    
    ' Copy headers for Match Results (skip IsMatch and Difference)
    Dim col As Long, matchCol As Long
    matchCol = 1
    For col = idCol To lastCol
        wsMatch.Cells(1, matchCol).value = wsAllResults.Cells(1, col).value
        matchCol = matchCol + 1
    Next col
    
    ' Loop through data rows and copy matches
    Dim row As Long, matchRow As Long
    Dim isMatch As Boolean
    matchRow = 2
    
    For row = 2 To tblAll.Range.Rows.Count
        isMatch = wsAllResults.Cells(row, 1).value ' IsMatch column
        
        If isMatch Then
            ' Copy to Match Results (skip IsMatch and Difference columns)
            matchCol = 1
            For col = idCol To lastCol
                wsMatch.Cells(matchRow, matchCol).value = wsAllResults.Cells(row, col).value
                matchCol = matchCol + 1
            Next col
            matchRow = matchRow + 1
        End If
    Next row
    
    ' Create table if we have matches
    If matchRow > 2 Then
        Dim matchLastCol As Long
        matchLastCol = wsMatch.Cells(1, wsMatch.Columns.Count).End(xlToLeft).Column
        
        Dim matchTableRange As Range
        Set matchTableRange = wsMatch.Range(wsMatch.Cells(1, 1), wsMatch.Cells(matchRow - 1, matchLastCol))
        
        On Error Resume Next
        wsMatch.ListObjects("MatchResultsTable").Delete
        On Error GoTo 0
        
        wsMatch.ListObjects.Add(xlSrcRange, matchTableRange, , xlYes).Name = "MatchResultsTable"
        wsMatch.ListObjects("MatchResultsTable").TableStyle = "TableStyleMedium3"
        
        ' Format Match Results header
        With wsMatch.Range(wsMatch.Cells(1, 1), wsMatch.Cells(1, matchLastCol))
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
        
        ' Format number columns in Match Results (Sheet1 Total and Sheet2 Total)
        If matchLastCol >= 2 Then
            wsMatch.Range(wsMatch.Cells(2, 2), wsMatch.Cells(matchRow - 1, 3)).NumberFormat = "#,##0.00"
        End If
        
        wsMatch.Columns("A:" & ColLetter(matchLastCol)).AutoFit
    End If
End Sub

Private Sub WriteDetailMatchResults(wsMatch As Worksheet, wsAllResults As Worksheet, _
                                     ws1 As Worksheet, ws2 As Worksheet, _
                                     idCol1 As String, idCol2 As String, detailSheet As String, _
                                     dictDetailRows As Object, dict1 As Object, dict2 As Object, _
                                     dictAdditional1 As Object, dictAdditional2 As Object, _
                                     additionalColumns1 As String, additionalColumns2 As String, _
                                     idCol As Long, sheet1TotalCol As Long, sheet2TotalCol As Long)
    ' Detail expansion mode: Output detail rows for matched IDs
    Dim tblAll As ListObject
    Set tblAll = wsAllResults.ListObjects("ReconResultsTable")
    
    ' Parse additional column names for both sheets
    Dim additionalCols1() As String, additionalCols2() As String
    Dim hasAdditional1 As Boolean, hasAdditional2 As Boolean
    
    hasAdditional1 = (Len(Trim(additionalColumns1)) > 0)
    hasAdditional2 = (Len(Trim(additionalColumns2)) > 0)
    
    Dim i As Long
    If hasAdditional1 Then
        additionalCols1 = Split(additionalColumns1, ",")
        For i = 0 To UBound(additionalCols1)
            additionalCols1(i) = Trim(additionalCols1(i))
        Next i
    End If
    
    If hasAdditional2 Then
        additionalCols2 = Split(additionalColumns2, ",")
        For i = 0 To UBound(additionalCols2)
            additionalCols2(i) = Trim(additionalCols2(i))
        Next i
    End If
    
    ' Build header row
    Dim col As Long
    col = 1
    wsMatch.Cells(1, col).value = "ID": col = col + 1
    
    ' Add detail sheet's additional columns first
    If detailSheet = "SHEET1" And hasAdditional1 Then
        For i = 0 To UBound(additionalCols1)
            wsMatch.Cells(1, col).value = additionalCols1(i)
            col = col + 1
        Next i
    ElseIf detailSheet = "SHEET2" And hasAdditional2 Then
        For i = 0 To UBound(additionalCols2)
            wsMatch.Cells(1, col).value = additionalCols2(i)
            col = col + 1
        Next i
    End If
    
    ' Add totals
    wsMatch.Cells(1, col).value = "Sheet1 Total": col = col + 1
    wsMatch.Cells(1, col).value = "Sheet2 Total": col = col + 1
    
    ' Add summary sheet's additional columns
    If detailSheet = "SHEET1" And hasAdditional2 Then
        For i = 0 To UBound(additionalCols2)
            wsMatch.Cells(1, col).value = additionalCols2(i)
            col = col + 1
        Next i
    ElseIf detailSheet = "SHEET2" And hasAdditional1 Then
        For i = 0 To UBound(additionalCols1)
            wsMatch.Cells(1, col).value = additionalCols1(i)
            col = col + 1
        Next i
    End If
    
    Dim lastHeaderCol As Long
    lastHeaderCol = col - 1
    
    ' Format header
    With wsMatch.Range(wsMatch.Cells(1, 1), wsMatch.Cells(1, lastHeaderCol))
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Write detail rows for matched IDs
    Dim row As Long, matchRow As Long
    Dim isMatch As Boolean, id As String
    Dim total1 As Double, total2 As Double
    Dim detailRows As Collection, detailRow As Object
    Dim summaryDict As Object
    Dim j As Long
    
    matchRow = 2
    For row = 2 To tblAll.Range.Rows.Count
        isMatch = wsAllResults.Cells(row, 1).value ' IsMatch column
        
        If isMatch Then
            id = CStr(wsAllResults.Cells(row, idCol).value)
            total1 = wsAllResults.Cells(row, sheet1TotalCol).value
            total2 = wsAllResults.Cells(row, sheet2TotalCol).value
            
            ' Get detail rows for this ID
            If dictDetailRows.Exists(id) Then
                Set detailRows = dictDetailRows(id)
                
                ' Get summary sheet's additional columns
                If detailSheet = "SHEET1" Then
                    Set summaryDict = Nothing
                    If dictAdditional2.Exists(id) Then Set summaryDict = dictAdditional2(id)
                Else
                    Set summaryDict = Nothing
                    If dictAdditional1.Exists(id) Then Set summaryDict = dictAdditional1(id)
                End If
                
                ' Write each detail row
                For Each detailRow In detailRows
                    col = 1
                    wsMatch.Cells(matchRow, col).value = id: col = col + 1
                    
                    ' Write detail columns
                    If detailSheet = "SHEET1" And hasAdditional1 Then
                        For i = 0 To UBound(additionalCols1)
                            If detailRow.Exists(additionalCols1(i)) Then
                                wsMatch.Cells(matchRow, col).value = detailRow(additionalCols1(i))
                            End If
                            col = col + 1
                        Next i
                    ElseIf detailSheet = "SHEET2" And hasAdditional2 Then
                        For i = 0 To UBound(additionalCols2)
                            If detailRow.Exists(additionalCols2(i)) Then
                                wsMatch.Cells(matchRow, col).value = detailRow(additionalCols2(i))
                            End If
                            col = col + 1
                        Next i
                    End If
                    
                    ' Write totals (repeated for each detail row)
                    wsMatch.Cells(matchRow, col).value = total1: col = col + 1
                    wsMatch.Cells(matchRow, col).value = total2: col = col + 1
                    
                    ' Write summary columns
                    If Not summaryDict Is Nothing Then
                        If detailSheet = "SHEET1" And hasAdditional2 Then
                            For i = 0 To UBound(additionalCols2)
                                If summaryDict.Exists(additionalCols2(i)) Then
                                    wsMatch.Cells(matchRow, col).value = summaryDict(additionalCols2(i))
                                End If
                                col = col + 1
                            Next i
                        ElseIf detailSheet = "SHEET2" And hasAdditional1 Then
                            For i = 0 To UBound(additionalCols1)
                                If summaryDict.Exists(additionalCols1(i)) Then
                                    wsMatch.Cells(matchRow, col).value = summaryDict(additionalCols1(i))
                                End If
                                col = col + 1
                            Next i
                        End If
                    End If
                    
                    matchRow = matchRow + 1
                Next detailRow
            End If
        End If
    Next row
    
    ' Create table if we have matches
    If matchRow > 2 Then
        Dim matchTableRange As Range
        Set matchTableRange = wsMatch.Range(wsMatch.Cells(1, 1), wsMatch.Cells(matchRow - 1, lastHeaderCol))
        
        On Error Resume Next
        wsMatch.ListObjects("MatchResultsTable").Delete
        On Error GoTo 0
        
        wsMatch.ListObjects.Add(xlSrcRange, matchTableRange, , xlYes).Name = "MatchResultsTable"
        wsMatch.ListObjects("MatchResultsTable").TableStyle = "TableStyleMedium3"
        
        ' Format number columns (totals)
        Dim totalCol1 As Long, totalCol2 As Long
        If detailSheet = "SHEET1" And hasAdditional1 Then
            totalCol1 = 2 + UBound(additionalCols1) + 1
        ElseIf detailSheet = "SHEET2" And hasAdditional2 Then
            totalCol1 = 2 + UBound(additionalCols2) + 1
        Else
            totalCol1 = 2
        End If
        totalCol2 = totalCol1 + 1
        
        wsMatch.Range(wsMatch.Cells(2, totalCol1), wsMatch.Cells(matchRow - 1, totalCol2)).NumberFormat = "#,##0.00"
        wsMatch.Columns("A:" & ColLetter(lastHeaderCol)).AutoFit
    End If
End Sub

Private Sub WriteErrorResults(wsError As Worksheet, wsAllResults As Worksheet, idCol As Long, diffCol As Long, lastCol As Long)
    ' Error Results: always aggregated
    Dim tblAll As ListObject
    Set tblAll = wsAllResults.ListObjects("ReconResultsTable")
    
    ' Copy headers for Error Results (skip IsMatch only, keep Difference)
    Dim col As Long, errorCol As Long
    errorCol = 1
    wsError.Cells(1, errorCol).value = "Difference": errorCol = errorCol + 1
    For col = idCol To lastCol
        wsError.Cells(1, errorCol).value = wsAllResults.Cells(1, col).value
        errorCol = errorCol + 1
    Next col
    
    ' Loop through data rows and copy errors
    Dim row As Long, errorRow As Long
    Dim isMatch As Boolean
    errorRow = 2
    
    For row = 2 To tblAll.Range.Rows.Count
        isMatch = wsAllResults.Cells(row, 1).value ' IsMatch column
        
        If Not isMatch Then
            ' Copy to Error Results (skip IsMatch, keep Difference)
            errorCol = 1
            wsError.Cells(errorRow, errorCol).value = wsAllResults.Cells(row, diffCol).value
            errorCol = errorCol + 1
            For col = idCol To lastCol
                wsError.Cells(errorRow, errorCol).value = wsAllResults.Cells(row, col).value
                errorCol = errorCol + 1
            Next col
            
            ' Highlight error rows
            wsError.Rows(errorRow).Interior.Color = RGB(255, 200, 200)
            errorRow = errorRow + 1
        End If
    Next row
    
    ' Create table if we have errors
    If errorRow > 2 Then
        Dim errorLastCol As Long
        errorLastCol = wsError.Cells(1, wsError.Columns.Count).End(xlToLeft).Column
        
        Dim errorTableRange As Range
        Set errorTableRange = wsError.Range(wsError.Cells(1, 1), wsError.Cells(errorRow - 1, errorLastCol))
        
        On Error Resume Next
        wsError.ListObjects("ErrorResultsTable").Delete
        On Error GoTo 0
        
        wsError.ListObjects.Add(xlSrcRange, errorTableRange, , xlYes).Name = "ErrorResultsTable"
        wsError.ListObjects("ErrorResultsTable").TableStyle = "TableStyleMedium1"
        
        ' Format Error Results header
        With wsError.Range(wsError.Cells(1, 1), wsError.Cells(1, errorLastCol))
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
        
        ' Format number columns in Error Results (Difference, Sheet1 Total, Sheet2 Total)
        If errorLastCol >= 3 Then
            wsError.Range(wsError.Cells(2, 1), wsError.Cells(errorRow - 1, 3)).NumberFormat = "#,##0.00"
        End If
        
        wsError.Columns("A:" & ColLetter(errorLastCol)).AutoFit
    End If
End Sub



