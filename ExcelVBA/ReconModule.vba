Option Explicit

' Main reconciliation module
' Ported from C# WindowsForms ExcelReconApp

Public Sub RunReconciliation()
    ' Validate setup
    If Not ValidateSetup() Then Exit Sub
    
    ' Clear previous results
    ClearResults
    
    ' Get configuration
    Dim ws1 As Worksheet, ws2 As Worksheet, wsResults As Worksheet, wsConfig As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Sheet2")
    Set wsResults = ThisWorkbook.Worksheets("Results")
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    
    ' Get ID columns and value columns from config sheet
    Dim idCol1 As String, idCol2 As String
    Dim valueColumns1 As String, valueColumns2 As String
    Dim additionalColumns1 As String
    
    idCol1 = wsConfig.Range("B2").Value
    idCol2 = wsConfig.Range("B3").Value
    valueColumns1 = wsConfig.Range("B4").Value ' Comma-separated list
    valueColumns2 = wsConfig.Range("B5").Value ' Comma-separated list
    additionalColumns1 = wsConfig.Range("B6").Value ' Comma-separated list (optional)
    
    ' Build dictionaries of ID -> Total
    Dim dict1 As Object, dict2 As Object
    Set dict1 = BuildTotalsDictionary(ws1, idCol1, valueColumns1)
    Set dict2 = BuildTotalsDictionary(ws2, idCol2, valueColumns2)
    
    ' Build dictionary of ID -> first row additional column values
    Dim dictAdditional As Object
    Set dictAdditional = BuildAdditionalColumnsDictionary(ws1, idCol1, additionalColumns1)
    
    ' Get all unique IDs from both sheets
    Dim allIDs As Collection
    Set allIDs = GetUniqueIDs(dict1, dict2)
    
    ' Write results
    WriteResults wsResults, allIDs, dict1, dict2, dictAdditional, additionalColumns1
    
    MsgBox "Reconciliation completed! " & allIDs.Count & " records processed.", vbInformation, "Success"
End Sub

Private Function ValidateSetup() As Boolean
    ValidateSetup = False
    
    ' Check worksheets exist
    On Error Resume Next
    Dim ws1 As Worksheet, ws2 As Worksheet, wsConfig As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Sheet2")
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    On Error GoTo 0
    
    If ws1 Is Nothing Or ws2 Is Nothing Or wsConfig Is Nothing Then
        MsgBox "Required worksheets not found. Ensure 'Sheet1', 'Sheet2', and 'Config' exist.", vbCritical, "Error"
        Exit Function
    End If
    
    ' Check configuration is filled
    If Len(wsConfig.Range("B2").Value) = 0 Or Len(wsConfig.Range("B3").Value) = 0 Then
        MsgBox "Please specify ID columns in Config sheet.", vbCritical, "Error"
        Exit Function
    End If
    
    If Len(wsConfig.Range("B4").Value) = 0 Or Len(wsConfig.Range("B5").Value) = 0 Then
        MsgBox "Please specify value columns in Config sheet.", vbCritical, "Error"
        Exit Function
    End If
    
    ' Check data exists
    If ws1.Cells(2, 1).Value = "" Or ws2.Cells(2, 1).Value = "" Then
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
    
    ' Find value column indexes
    Dim valueColIndexes() As Long
    ReDim valueColIndexes(0 To UBound(valueColNames))
    Dim i As Long
    For i = 0 To UBound(valueColNames)
        valueColNames(i) = Trim(valueColNames(i))
        valueColIndexes(i) = FindColumnIndex(ws, valueColNames(i))
        If valueColIndexes(i) = 0 Then
            MsgBox "Column '" & valueColNames(i) & "' not found in " & ws.Name, vbCritical, "Error"
            Set BuildTotalsDictionary = dict
            Exit Function
        End If
    Next i
    
    ' Loop through data rows (starting from row 2)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, idColIndex).End(xlUp).Row
    
    Dim row As Long, id As String, total As Double, value As Double
    For row = 2 To lastRow
        id = Trim(CStr(ws.Cells(row, idColIndex).Value))
        
        ' Skip blank IDs
        If Len(id) = 0 Then GoTo NextRow
        
        ' Sum value columns for this row
        total = 0
        For i = 0 To UBound(valueColIndexes)
            If IsNumeric(ws.Cells(row, valueColIndexes(i)).Value) Then
                value = CDbl(ws.Cells(row, valueColIndexes(i)).Value)
                total = total + value
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
    lastRow = ws.Cells(ws.Rows.Count, idColIndex).End(xlUp).Row
    
    Dim row As Long, id As String
    Dim additionalDict As Object
    
    For row = 2 To lastRow
        id = Trim(CStr(ws.Cells(row, idColIndex).Value))
        
        ' Skip blank IDs
        If Len(id) = 0 Then GoTo NextRowAdditional
        
        ' Only store first occurrence of each ID
        If Not dict.Exists(id) Then
            Set additionalDict = CreateObject("Scripting.Dictionary")
            
            ' Store values from additional columns
            For i = 0 To UBound(additionalColIndexes)
                additionalDict.Add additionalColNames(i), ws.Cells(row, additionalColIndexes(i)).Value
            Next i
            
            dict.Add id, additionalDict
        End If
        
NextRowAdditional:
    Next row
    
    Set BuildAdditionalColumnsDictionary = dict
End Function

Private Function FindColumnIndex(ws As Worksheet, columnName As String) As Long
    ' Find column index by header name in row 1
    Dim col As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If Trim(CStr(ws.Cells(1, col).Value)) = Trim(columnName) Then
            FindColumnIndex = col
            Exit Function
        End If
    Next col
    
    FindColumnIndex = 0
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

Private Sub WriteResults(wsResults As Worksheet, allIDs As Collection, dict1 As Object, dict2 As Object, dictAdditional As Object, additionalColumnsStr As String)
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
    wsResults.Cells(1, col).Value = "IsMatch": col = col + 1
    wsResults.Cells(1, col).Value = "Difference": col = col + 1
    wsResults.Cells(1, col).Value = "ID": col = col + 1
    wsResults.Cells(1, col).Value = "Sheet1 Total": col = col + 1
    wsResults.Cells(1, col).Value = "Sheet2 Total": col = col + 1
    
    ' Add additional column headers
    If hasAdditionalCols Then
        For i = 0 To UBound(additionalColNames)
            wsResults.Cells(1, col).Value = additionalColNames(i)
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
    
    ' Sort IDs (convert to array first)
    Dim sortedIDs() As String
    ReDim sortedIDs(1 To allIDs.Count)
    For i = 1 To allIDs.Count
        sortedIDs(i) = allIDs(i)
    Next i
    QuickSort sortedIDs, LBound(sortedIDs), UBound(sortedIDs)
    
    ' Write data rows
    Dim row As Long, id As String
    Dim total1 As Double, total2 As Double, diff As Double
    Dim isMatch As Boolean
    Dim additionalDict As Object
    Dim j As Long
    
    row = 2
    For i = LBound(sortedIDs) To UBound(sortedIDs)
        id = sortedIDs(i)
        
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
        isMatch = (Abs(diff) < 0.01) ' Tolerance for floating point comparison
        
        ' Write to sheet in new order
        col = 1
        wsResults.Cells(row, col).Value = isMatch: col = col + 1
        wsResults.Cells(row, col).Value = diff: col = col + 1
        wsResults.Cells(row, col).Value = id: col = col + 1
        wsResults.Cells(row, col).Value = total1: col = col + 1
        wsResults.Cells(row, col).Value = total2: col = col + 1
        
        ' Write additional columns
        If hasAdditionalCols And dictAdditional.Exists(id) Then
            Set additionalDict = dictAdditional(id)
            For j = 0 To UBound(additionalColNames)
                If additionalDict.Exists(additionalColNames(j)) Then
                    wsResults.Cells(row, col).Value = additionalDict(additionalColNames(j))
                End If
                col = col + 1
            Next j
        ElseIf hasAdditionalCols Then
            ' Skip columns if no additional data
            col = col + UBound(additionalColNames) + 1
        End If
        
        ' Highlight non-matches
        If Not isMatch Then
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
    Dim wsResults As Worksheet
    Set wsResults = ThisWorkbook.Worksheets("Results")
    
    ' Delete existing table if present
    On Error Resume Next
    wsResults.ListObjects("ReconResultsTable").Delete
    On Error GoTo 0
    
    ' Clear all cells
    wsResults.Cells.Clear
End Sub

Private Sub QuickSort(arr() As String, ByVal left As Long, ByVal right As Long)
    ' Simple quicksort for string array
    Dim i As Long, j As Long
    Dim pivot As String, temp As String
    
    If left < right Then
        pivot = arr((left + right) \ 2)
        i = left
        j = right
        
        Do While i <= j
            Do While arr(i) < pivot
                i = i + 1
            Loop
            Do While arr(j) > pivot
                j = j - 1
            Loop
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If left < j Then QuickSort arr, left, j
        If i < right Then QuickSort arr, i, right
    End If
End Sub

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
