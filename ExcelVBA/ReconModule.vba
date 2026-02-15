Option Explicit

' Main reconciliation module
' Ported from C# WindowsForms ExcelReconApp

' Module-level variable to store active workflow
Public ActiveWorkflow As String

' Get workflow configuration from WorkflowsTable
Private Function GetWorkflowConfig(workflowName As String) As Object
    ' Returns Dictionary with: ReconConfigSheet, TransformConfigSheet, ResultSheetPrefix
    ' Config sheet names are now hardcoded, only ResultSheetPrefix comes from Workflows table
    Dim workflowConfig As Object
    Set workflowConfig = CreateObject("Scripting.Dictionary")
    
    ' Hardcoded configuration sheet names
    workflowConfig("ReconConfigSheet") = "Recon Config"
    workflowConfig("TransformConfigSheet") = "Transform Config"
    workflowConfig("ResultSheetPrefix") = ""
    
    ' Check if Workflows sheet exists
    Dim wsWorkflows As Worksheet
    On Error Resume Next
    Set wsWorkflows = ThisWorkbook.Worksheets("Workflows")
    On Error GoTo 0
    
    If wsWorkflows Is Nothing Then
        ' No Workflows sheet, return defaults with hardcoded sheet names
        Set GetWorkflowConfig = workflowConfig
        Exit Function
    End If
    
    ' Get WorkflowsTable (first table in the sheet)
    Dim tblWorkflows As ListObject
    On Error Resume Next
    If wsWorkflows.ListObjects.Count > 0 Then
        Set tblWorkflows = wsWorkflows.ListObjects(1)
    End If
    On Error GoTo 0
    
    If tblWorkflows Is Nothing Then
        ' No table in Workflows sheet, return defaults
        Set GetWorkflowConfig = workflowConfig
        Exit Function
    End If
    
    ' Find column indexes (only need Workflow Name and Result Sheet Prefix)
    Dim nameCol As Long, prefixCol As Long
    Dim col As Long
    
    For col = 1 To tblWorkflows.ListColumns.Count
        Select Case UCase(Trim(tblWorkflows.HeaderRowRange.Cells(1, col).value))
            Case "WORKFLOW NAME"
                nameCol = col
            Case "RESULT SHEET PREFIX", "RESULT PREFIX"
                prefixCol = col
        End Select
    Next col
    
    If nameCol = 0 Then
        ' Required column not found, return defaults
        Set GetWorkflowConfig = workflowConfig
        Exit Function
    End If
    
    ' Search for workflow by name (case-insensitive)
    Dim row As Long
    Dim workflowNameUpper As String
    workflowNameUpper = UCase(Trim(workflowName))
    
    For row = 1 To tblWorkflows.ListRows.Count
        If UCase(Trim(tblWorkflows.DataBodyRange.Cells(row, nameCol).value)) = workflowNameUpper Then
            ' Found the workflow - read result prefix if column exists
            If prefixCol > 0 Then
                workflowConfig("ResultSheetPrefix") = Trim(CStr(tblWorkflows.DataBodyRange.Cells(row, prefixCol).value))
            End If
            Exit For
        End If
    Next row
    
    Set GetWorkflowConfig = workflowConfig
End Function

Public Sub RunReconciliation(Optional workflowName As String = "")
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
    Dim sheet1Filter As String
    Dim sheet2Filter As String
    
    ' Get workflow configuration
    Dim workflowConfig As Object
    If Len(workflowName) = 0 Then
        ' Use ActiveWorkflow if no workflow specified
        If Len(ActiveWorkflow) > 0 Then
            workflowName = ActiveWorkflow
        Else
            workflowName = "Default"
        End If
    End If
    
    Set workflowConfig = GetWorkflowConfig(workflowName)
    Dim reconConfigSheetName As String
    reconConfigSheetName = workflowConfig("ReconConfigSheet")
    Dim resultPrefix As String
    resultPrefix = workflowConfig("ResultSheetPrefix")
    
    ' Validate setup
    If Not ValidateSetup(reconConfigSheetName, workflowName) Then Exit Sub
    
    ' Clear previous results
    ClearResults resultPrefix
    
    ' Get configuration
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Sheet2")
    Set wsResults = GetOrCreateWorksheet(ApplyPrefix(resultPrefix, "All Results"))
    Set wsConfig = ThisWorkbook.Worksheets(reconConfigSheetName)
    
    ' Get configuration settings from table
    idCol1 = GetConfigValue(wsConfig, "Sheet1 ID Column", workflowName)
    idCol2 = GetConfigValue(wsConfig, "Sheet2 ID Column", workflowName)
    valueColumns1 = GetConfigValue(wsConfig, "Sheet1 Value Columns", workflowName)
    valueColumns2 = GetConfigValue(wsConfig, "Sheet2 Value Columns", workflowName)
    additionalColumns1 = GetConfigValue(wsConfig, "Sheet1 Additional Columns", workflowName)
    sheet1Filter = GetConfigValue(wsConfig, "Sheet1 Filter", workflowName)
    sheet2Filter = GetConfigValue(wsConfig, "Sheet2 Filter", workflowName)
    
    ' Get tolerance (default to 0.01 if not specified or invalid)
    Dim toleranceStr As String
    toleranceStr = GetConfigValue(wsConfig, "Tolerance", workflowName)
    If IsNumeric(toleranceStr) And CDbl(toleranceStr) >= 0 Then
        tolerance = CDbl(toleranceStr)
    Else
        tolerance = 0.01
    End If
    
    ' Get detail sheet setting and additional columns from second sheet
    detailSheet = Trim(UCase(GetConfigValue(wsConfig, "Detail Sheet", workflowName)))
    additionalColumns2 = GetConfigValue(wsConfig, "Sheet2 Additional Columns", workflowName)
    
    ' Get sort order setting
    Dim sortOrder As String
    sortOrder = GetConfigValue(wsConfig, "Sort Order", workflowName)
    
    ' Build dictionaries of ID -> Total (with filtering)
    Set dict1 = BuildTotalsDictionary(ws1, idCol1, valueColumns1, sheet1Filter)
    Set dict2 = BuildTotalsDictionary(ws2, idCol2, valueColumns2, sheet2Filter)
    
    ' Build dictionaries of ID -> additional column values (from both sheets)
    Set dictAdditional1 = BuildAdditionalColumnsDictionary(ws1, idCol1, additionalColumns1)
    Set dictAdditional2 = BuildAdditionalColumnsDictionary(ws2, idCol2, additionalColumns2)
    
    ' Build detail rows dictionary if detail mode is enabled
    Set dictDetailRows = CreateObject("Scripting.Dictionary")
    If detailSheet = "SHEET1" Then
        Set dictDetailRows = BuildAllDetailRowsDictionary(ws1, idCol1, additionalColumns1, sheet1Filter)
    ElseIf detailSheet = "SHEET2" Then
        Set dictDetailRows = BuildAllDetailRowsDictionary(ws2, idCol2, additionalColumns2, sheet2Filter)
    End If
    
    ' Get all unique IDs from both sheets
    Set allIDs = GetUniqueIDs(dict1, dict2)
    
    ' Write results
    WriteResults wsResults, allIDs, dict1, dict2, dictAdditional1, additionalColumns1, tolerance, sortOrder
    
    ' Split results into Match Results and Error Results sheets
    Dim matchCount As Long, errorCount As Long
    SplitResults wsResults, ws1, ws2, idCol1, idCol2, detailSheet, dictDetailRows, dict1, dict2, dictAdditional1, dictAdditional2, additionalColumns1, additionalColumns2, resultPrefix, matchCount, errorCount
    
    ' Log the reconciliation
    Dim details As String
    details = "Config: " & reconConfigSheetName & " | Matches: " & matchCount & " | Errors: " & errorCount & " | Tolerance: " & tolerance & " | Detail Mode: " & IIf(Len(Trim(detailSheet)) > 0, detailSheet, "None")
    LogActivity "Reconciliation", workflowName, details, "Success", allIDs.Count
    
    MsgBox "Reconciliation completed! " & allIDs.Count & " records processed (" & matchCount & " matches, " & errorCount & " errors).", vbInformation, "Success"
End Sub

Private Function ValidateSetup(reconConfigSheetName As String, workflowName As String) As Boolean
    ValidateSetup = False
    
    ' Check worksheets exist
    On Error Resume Next
    Dim ws1 As Worksheet, ws2 As Worksheet, wsConfig As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Sheet2")
    Set wsConfig = ThisWorkbook.Worksheets(reconConfigSheetName)
    On Error GoTo 0
    
    If ws1 Is Nothing Or ws2 Is Nothing Or wsConfig Is Nothing Then
        MsgBox "Required worksheets not found. Ensure 'Sheet1', 'Sheet2', and '" & reconConfigSheetName & "' exist.", vbCritical, "Error"
        Exit Function
    End If
    
    ' Check that a config table exists (any table in the sheet)
    Dim tblConfig As ListObject
    On Error Resume Next
    If wsConfig.ListObjects.Count > 0 Then
        Set tblConfig = wsConfig.ListObjects(1)
    End If
    On Error GoTo 0
    
    If tblConfig Is Nothing Then
        MsgBox "No table found in '" & reconConfigSheetName & "' sheet. Please convert your configuration to a table (Ctrl+T).", vbCritical, "Error"
        Exit Function
    End If
    
    ' Check required configuration settings
    Dim idCol1 As String, idCol2 As String
    Dim valueColumns1 As String, valueColumns2 As String
    
    idCol1 = GetConfigValue(wsConfig, "Sheet1 ID Column", workflowName)
    idCol2 = GetConfigValue(wsConfig, "Sheet2 ID Column", workflowName)
    valueColumns1 = GetConfigValue(wsConfig, "Sheet1 Value Columns", workflowName)
    valueColumns2 = GetConfigValue(wsConfig, "Sheet2 Value Columns", workflowName)
    
    If Len(idCol1) = 0 Or Len(idCol2) = 0 Then
        MsgBox "Please specify both ID columns in Recon Config table.", vbCritical, "Error"
        Exit Function
    End If
    
    If Len(valueColumns1) = 0 Or Len(valueColumns2) = 0 Then
        MsgBox "Please specify value columns for both sheets in Recon Config table.", vbCritical, "Error"
        Exit Function
    End If
    
    ' Check for valid detail sheet configuration
    Dim detailSheet As String
    detailSheet = Trim(UCase(GetConfigValue(wsConfig, "Detail Sheet", workflowName)))
    If Len(detailSheet) > 0 And detailSheet <> "SHEET1" And detailSheet <> "SHEET2" Then
        MsgBox "Detail Sheet must be either 'Sheet1' or 'Sheet2' (case-insensitive) if specified.", vbCritical, "Error"
        Exit Function
    End If
    
    ' Check that both sheets are not set as detail (not supported)
    ' (In this config, only one Detail Sheet setting is possible, so this is just a future-proof check)
    ' Check data exists
    If ws1.Cells(2, 1).value = "" Or ws2.Cells(2, 1).value = "" Then
        MsgBox "Please load data into Sheet1 and Sheet2.", vbCritical, "Error"
        Exit Function
    End If
    
    ValidateSetup = True
End Function

Private Function BuildTotalsDictionary(ws As Worksheet, idColumnName As String, valueColumnsStr As String, Optional filterExpr As String = "") As Object
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
        ' Apply filter if specified
        If Len(Trim(filterExpr)) > 0 Then
            If Not EvaluateRowFilter(ws, row, filterExpr) Then
                GoTo NextRow
            End If
        End If
        
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

Private Function BuildAllDetailRowsDictionary(ws As Worksheet, idColumnName As String, additionalColumnsStr As String, Optional filterExpr As String = "") As Object
    ' Returns Dictionary with ID as key and Collection of all row dictionaries as value
    ' Each row dictionary contains column name -> value pairs
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Always build detail rows, even if no additional columns specified
    
    ' Find ID column index
    Dim idColIndex As Long
    idColIndex = FindColumnIndex(ws, idColumnName)
    If idColIndex = 0 Then
        Set BuildAllDetailRowsDictionary = dict
        Exit Function
    End If
    
    ' Parse additional column names only if specified
    Dim additionalColNames() As String
    Dim additionalColIndexes() As Long
    Dim hasAdditionalCols As Boolean
    hasAdditionalCols = (Len(Trim(additionalColumnsStr)) > 0)
    Dim i As Long
    If hasAdditionalCols Then
        additionalColNames = Split(additionalColumnsStr, ",")
        ReDim additionalColIndexes(0 To UBound(additionalColNames))
        For i = 0 To UBound(additionalColNames)
            additionalColNames(i) = Trim(additionalColNames(i))
            additionalColIndexes(i) = FindColumnIndex(ws, additionalColNames(i))
            If additionalColIndexes(i) = 0 Then
                MsgBox "Additional column '" & additionalColNames(i) & "' not found in " & ws.Name, vbCritical, "Error"
                Set BuildAllDetailRowsDictionary = dict
                Exit Function
            End If
        Next i
    End If
    
    ' Loop through data rows (starting from row 2)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, idColIndex).End(xlUp).row
    
    Dim row As Long, id As String
    Dim rowDict As Object
    Dim rowCollection As Collection
    
    For row = 2 To lastRow
        ' Apply filter if specified
        If Len(Trim(filterExpr)) > 0 Then
            If Not EvaluateRowFilter(ws, row, filterExpr) Then
                GoTo NextRowAll
            End If
        End If
        
        id = Trim(CStr(ws.Cells(row, idColIndex).value))
        
        ' Skip blank IDs
        If Len(id) = 0 Then GoTo NextRowAll
        
        ' Create dictionary for this row
        Set rowDict = CreateObject("Scripting.Dictionary")
        If hasAdditionalCols Then
            For i = 0 To UBound(additionalColIndexes)
                rowDict.Add additionalColNames(i), ws.Cells(row, additionalColIndexes(i)).value
            Next i
        End If
        
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

Private Function GetConfigValue(wsConfig As Worksheet, settingName As String, Optional workflowName As String = "") As String
    ' Get configuration value from config table by setting name and optional workflow filter (case-insensitive)
    GetConfigValue = ""
    
    ' Get the first table in the config sheet
    Dim tblConfig As ListObject
    On Error Resume Next
    If wsConfig.ListObjects.Count > 0 Then
        Set tblConfig = wsConfig.ListObjects(1)
    End If
    On Error GoTo 0
    
    If tblConfig Is Nothing Then Exit Function
    
    ' Find Setting, Value, and optional Workflow column indexes
    Dim settingColIndex As Long, valueColIndex As Long, workflowColIndex As Long
    Dim col As Long
    
    For col = 1 To tblConfig.ListColumns.Count
        Select Case UCase(Trim(tblConfig.HeaderRowRange.Cells(1, col).value))
            Case "SETTING"
                settingColIndex = col
            Case "VALUE"
                valueColIndex = col
            Case "WORKFLOW"
                workflowColIndex = col
        End Select
    Next col
    
    If settingColIndex = 0 Or valueColIndex = 0 Then Exit Function
    
    ' Search for the setting (case-insensitive) and optional workflow match
    Dim row As Long
    Dim settingNameUpper As String, workflowNameUpper As String
    settingNameUpper = UCase(Trim(settingName))
    workflowNameUpper = UCase(Trim(workflowName))
    
    For row = 1 To tblConfig.ListRows.Count
        ' Check if setting matches
        If UCase(Trim(CStr(tblConfig.DataBodyRange.Cells(row, settingColIndex).value))) = settingNameUpper Then
            ' If workflow filtering is requested and Workflow column exists
            If Len(workflowName) > 0 And workflowColIndex > 0 Then
                ' Check if workflow matches
                If UCase(Trim(CStr(tblConfig.DataBodyRange.Cells(row, workflowColIndex).value))) = workflowNameUpper Then
                    GetConfigValue = Trim(CStr(tblConfig.DataBodyRange.Cells(row, valueColIndex).value))
                    Exit Function
                End If
            Else
                ' No workflow filtering - return first match
                GetConfigValue = Trim(CStr(tblConfig.DataBodyRange.Cells(row, valueColIndex).value))
                Exit Function
            End If
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

    ' Clear previous data and formatting in wsResults
    wsResults.Cells.Clear

    ' Write header
    Dim lastHeaderCol As Long
    lastHeaderCol = WriteResultsHeader(wsResults, additionalColNames, hasAdditionalCols)

    ' Build and sort row data
    Dim rowData As Collection
    Set rowData = BuildResultsRowData(allIDs, dict1, dict2, dictAdditional, additionalColNames, hasAdditionalCols, tolerance)
    SortRowData rowData, sortOrder, wsResults

    ' Write data rows
    Dim lastRow As Long
    lastRow = WriteResultsDataRows(wsResults, rowData, additionalColNames, hasAdditionalCols)

    ' Force IsMatch column to Boolean (TRUE/FALSE)
    Dim r As Long
    For r = 2 To lastRow
        wsResults.Cells(r, 1).Value = CBool(wsResults.Cells(r, 1).Value)
    Next r

    ' Create and format table
    CreateAndFormatResultsTable wsResults, lastRow, lastHeaderCol
End Sub

Private Function WriteResultsHeader(wsResults As Worksheet, additionalColNames() As String, hasAdditionalCols As Boolean) As Long
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
        Dim i As Long
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
    
    WriteResultsHeader = lastHeaderCol
End Function

Private Function BuildResultsRowData(allIDs As Collection, dict1 As Object, dict2 As Object, _
                                     dictAdditional As Object, additionalColNames() As String, _
                                     hasAdditionalCols As Boolean, tolerance As Double) As Collection
    ' Build row data collection
    Dim rowData As Collection
    Set rowData = New Collection
    
    Dim i As Long, j As Long
    Dim id As String
    Dim total1 As Double, total2 As Double, diff As Double
    Dim isMatch As Boolean
    Dim additionalDict As Object
    Dim rowDict As Object
    
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
    
    Set BuildResultsRowData = rowData
End Function

Private Function WriteResultsDataRows(wsResults As Worksheet, rowData As Collection, _
                                      additionalColNames() As String, hasAdditionalCols As Boolean) As Long
    ' Write sorted data rows
    Dim row As Long, col As Long
    Dim i As Long, j As Long
    Dim rowDict As Object
    
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
    
    WriteResultsDataRows = row - 1  ' Return last row number
End Function

Private Sub CreateAndFormatResultsTable(wsResults As Worksheet, lastRow As Long, lastHeaderCol As Long)
    ' Format number columns (Difference, Sheet1 Total, Sheet2 Total - skip ID in column 3)
    If lastRow >= 2 Then
        ' Column 2: Difference
        wsResults.Range(wsResults.Cells(2, 2), wsResults.Cells(lastRow, 2)).NumberFormat = "#,##0.00"
        ' Columns 4-5: Sheet1 Total, Sheet2 Total
        wsResults.Range(wsResults.Cells(2, 4), wsResults.Cells(lastRow, 5)).NumberFormat = "#,##0.00"
    End If
    
    ' Convert to Table
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

Private Sub ClearResults(Optional resultPrefix As String = "")
    ' Clear All Results sheet
    Dim wsResults As Worksheet
    Set wsResults = GetOrCreateWorksheet(ApplyPrefix(resultPrefix, "All Results"))
    
    ' Delete existing table if present
    On Error Resume Next
    wsResults.ListObjects("ReconResultsTable").Delete
    On Error GoTo 0
    
    ' Clear all cells
    wsResults.Cells.Clear
    
    ' Clear Match Results sheet
    Dim wsMatch As Worksheet
    Set wsMatch = GetOrCreateWorksheet(ApplyPrefix(resultPrefix, "Match Results"))
    On Error Resume Next
    wsMatch.ListObjects("MatchResultsTable").Delete
    On Error GoTo 0
    wsMatch.Cells.Clear
    
    ' Clear Error Results sheet
    Dim wsError As Worksheet
    Set wsError = GetOrCreateWorksheet(ApplyPrefix(resultPrefix, "Error Results"))
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
    
    fd.Title = "Select Excel or CSV File to Import"
    fd.Filters.Clear
    fd.Filters.Add "Excel and CSV Files", "*.xlsx;*.xls;*.xlsm;*.csv"
    fd.Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
    fd.Filters.Add "CSV Files", "*.csv"
    
    If fd.Show = -1 Then
        Dim filePath As String
        filePath = fd.SelectedItems(1)
        
        On Error GoTo ImportError
        
        ' Clear target sheet
        targetSheet.Cells.Clear
        
        ' Open file and copy data
        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath, ReadOnly:=True)
        
        ' Copy first worksheet
        wb.Worksheets(1).UsedRange.Copy targetSheet.Range("A1")
        
        wb.Close False
        
        ' Count rows imported (excluding header)
        Dim rowCount As Long
        rowCount = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row - 1
        
        ' Extract filename from path
        Dim fileName As String
        Dim pos As Long
        pos = InStrRev(filePath, "\")
        If pos > 0 Then
            fileName = Mid(filePath, pos + 1)
        Else
            fileName = filePath
        End If
        
        ' Log the import
        Dim details As String
        details = "File: " & fileName & " | Target: " & targetSheet.Name
        LogActivity "Import", "", details, "Success", rowCount
        
        ' Update form textbox to show loaded file
        If targetSheet.Name = "Sheet1" Then
            ReconForm.txtFile1.Text = fileName
        ElseIf targetSheet.Name = "Sheet2" Then
            ReconForm.txtFile2.Text = fileName
        End If
        
        MsgBox "Successfully loaded data into " & targetSheet.Name, vbInformation, "Success"
        Exit Sub
        
ImportError:
        ' Log error
        LogActivity "Import", "", "Error loading file: " & filePath & " - " & Err.Description, "Error", 0
        MsgBox "Error loading file: " & Err.Description, vbCritical, "Error"
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
Public Sub RunTransform(Optional workflowName As String = "")
    On Error GoTo ErrorHandler
    
    Dim startTime As Double
    startTime = Timer
    
    ' Get workflow configuration
    Dim workflowConfig As Object
    If Len(workflowName) = 0 Then
        ' Use ActiveWorkflow if no workflow specified
        If Len(ActiveWorkflow) > 0 Then
            workflowName = ActiveWorkflow
        Else
            workflowName = "Default"
        End If
    End If
    
    Set workflowConfig = GetWorkflowConfig(workflowName)
    Dim transformConfigSheetName As String
    transformConfigSheetName = workflowConfig("TransformConfigSheet")
    Dim resultPrefix As String
    resultPrefix = workflowConfig("ResultSheetPrefix")
    
    ' Check if Transform Config sheet exists
    Dim wsTransformConfig As Worksheet
    On Error Resume Next
    Set wsTransformConfig = ThisWorkbook.Worksheets(transformConfigSheetName)
    On Error GoTo ErrorHandler
    
    If wsTransformConfig Is Nothing Then
        MsgBox "Transform Config sheet '" & transformConfigSheetName & "' not found. Please create it first.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Read transform settings from TransformSettingsTable (first table)
    Dim transformName As String, sourceSheetName As String, targetSheetName As String
    Dim filterExpr As String, sortOrder As String
    
    transformName = GetConfigValue(wsTransformConfig, "Name")
    sourceSheetName = GetConfigValue(wsTransformConfig, "Source Sheet")
    targetSheetName = GetConfigValue(wsTransformConfig, "Target Sheet")
    filterExpr = GetConfigValue(wsTransformConfig, "Filter")
    sortOrder = GetConfigValue(wsTransformConfig, "Sort Order")
    
    If Len(transformName) = 0 Then transformName = "Transform"
    
    If Len(sourceSheetName) = 0 Or Len(targetSheetName) = 0 Then
        MsgBox "Transform configuration incomplete. Source Sheet and Target Sheet are required in TransformSettingsTable.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Apply workflow prefix to sheet names (for reconciliation result sheets)
    sourceSheetName = ApplyPrefix(resultPrefix, sourceSheetName)
    targetSheetName = ApplyPrefix(resultPrefix, targetSheetName)
    
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
    
    ' Read column definitions from TransformColumnsTable (second table)
    Dim columnDefs As Collection
    Set columnDefs = New Collection
    
    ' Get second table in Transform Config sheet
    Dim tblColumns As ListObject
    On Error Resume Next
    If wsTransformConfig.ListObjects.Count >= 2 Then
        Set tblColumns = wsTransformConfig.ListObjects(2)
    End If
    On Error GoTo ErrorHandler
    
    If tblColumns Is Nothing Then
        MsgBox "TransformColumnsTable (second table) not found in Transform Config sheet.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Find column indexes in TransformColumnsTable
    Dim orderColIdx As Long, targetColIdx As Long, typeColIdx As Long, sourceColIdx As Long
    Dim col As Long
    
    For col = 1 To tblColumns.ListColumns.Count
        Select Case UCase(Trim(tblColumns.HeaderRowRange.Cells(1, col).value))
            Case "ORDER"
                orderColIdx = col
            Case "TARGET COLUMN"
                targetColIdx = col
            Case "TYPE"
                typeColIdx = col
            Case "SOURCE"
                sourceColIdx = col
        End Select
    Next col
    
    If orderColIdx = 0 Or targetColIdx = 0 Or typeColIdx = 0 Or sourceColIdx = 0 Then
        MsgBox "TransformColumnsTable must have columns: Order, Target Column, Type, Source", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Read column definitions from table
    Dim row As Long
    For row = 1 To tblColumns.ListRows.Count
        Dim colDef As Object
        Set colDef = CreateObject("Scripting.Dictionary")
        
        colDef("Order") = CLng(tblColumns.DataBodyRange.Cells(row, orderColIdx).value)
        colDef("TargetColumn") = Trim(CStr(tblColumns.DataBodyRange.Cells(row, targetColIdx).value))
        colDef("Type") = Trim(UCase(CStr(tblColumns.DataBodyRange.Cells(row, typeColIdx).value)))
        
        ' Read source as text/value
        Dim sourceValue As String
        sourceValue = Trim(CStr(tblColumns.DataBodyRange.Cells(row, sourceColIdx).value))
        
        ' For FORMULA type, ensure it starts with =
        If colDef("Type") = "FORMULA" Then
            If Left(sourceValue, 1) <> "=" Then
                sourceValue = "=" & sourceValue
            End If
        End If
        
        colDef("Source") = sourceValue
        
        columnDefs.Add colDef
    Next row
    
    If columnDefs.Count = 0 Then
        MsgBox "No column definitions found in TransformColumnsTable.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Apply transformation with filtering and sorting
    ApplyTransformation wsSource, wsTarget, columnDefs, filterExpr, sortOrder
    
    ' Determine row count
    Dim rowCount As Long
    rowCount = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).row - 1 ' Subtract header
    
    ' Log the transformation
    Dim details As String
    Dim duration As Double
    duration = Timer - startTime
    details = "Transform: " & transformName & " | Source: " & sourceSheetName & " | Target: " & targetSheetName & " | Duration: " & Format(duration, "0.0") & "s"
    LogActivity "Transformation", workflowName, details, "Success", rowCount
    
    MsgBox "Transformation completed successfully!" & vbCrLf & _
           "Transform: " & transformName & vbCrLf & _
           "Source: " & sourceSheetName & vbCrLf & _
           "Target: " & targetSheetName, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    ' Log error
    LogActivity "Transformation", workflowName, "Error: " & Err.Description, "Error", 0
    MsgBox "Error during transformation: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ApplyTransformation(wsSource As Worksheet, wsTarget As Worksheet, columnDefs As Collection, Optional filterExpr As String = "", Optional sortOrder As String = "")
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
    
    ' Populate rows with EXISTING and STATIC columns only (applying filter)
    Dim sourceRow As Long, targetRow As Long
    Dim sourceColIndex As Long
    targetRow = 1 ' Start at 1, will increment to 2 for first data row
    
    For sourceRow = 2 To lastSourceRow
        ' Apply filter if specified
        If Len(filterExpr) > 0 Then
            If Not EvaluateRowFilter(wsSource, sourceRow, filterExpr) Then
                ' Skip this row - doesn't match filter
                GoTo NextSourceRow
            End If
        End If
        
        ' Row passes filter - copy it
        targetRow = targetRow + 1
        
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
        
NextSourceRow:
    Next sourceRow
    
    ' Check if any rows were copied
    If targetRow < 2 Then
        MsgBox "No rows match the filter criteria.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    ' Create table with the data we have
    Dim lastCol As Long
    lastCol = columnDefs.Count
    Dim lastRow As Long
    lastRow = targetRow ' Use actual target row count after filtering
    
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
    
    ' Apply sorting if specified
    If Len(sortOrder) > 0 Then
        ApplySorting wsTarget, sortOrder
    End If
    
    ' Auto-fit columns
    wsTarget.Columns("A:" & ColLetter(lastCol)).AutoFit
End Sub

Private Sub ApplySorting(ws As Worksheet, sortOrder As String)
    ' Apply sorting to worksheet table based on sort order specification
    ' sortOrder format: "Column1 ASC, Column2 DESC, Column3" (default is ASC)
    
    If Len(Trim(sortOrder)) = 0 Then Exit Sub
    
    ' Get the table on this worksheet
    If ws.ListObjects.Count = 0 Then Exit Sub
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(1)
    
    ' Parse sort specifications
    Dim sortSpecs() As String
    sortSpecs = Split(sortOrder, ",")
    
    ' Clear existing sort
    tbl.Sort.SortFields.Clear
    
    ' Add each sort field
    Dim i As Long, spec As String, colName As String, direction As Long
    Dim colIndex As Long
    
    For i = 0 To UBound(sortSpecs)
        spec = Trim(sortSpecs(i))
        
        ' Parse column name and direction
        If UCase(Right(spec, 5)) = " DESC" Then
            colName = Trim(Left(spec, Len(spec) - 5))
            direction = xlDescending
        ElseIf UCase(Right(spec, 4)) = " ASC" Then
            colName = Trim(Left(spec, Len(spec) - 4))
            direction = xlAscending
        Else
            colName = spec
            direction = xlAscending ' Default
        End If
        
        ' Find column index in table
        colIndex = FindColumnIndex(ws, colName)
        If colIndex > 0 Then
            tbl.Sort.SortFields.Add Key:=ws.Cells(1, colIndex), _
                SortOn:=xlSortOnValues, _
                Order:=direction, _
                DataOption:=xlSortNormal
        End If
    Next i
    
    ' Apply the sort
    With tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
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
                         additionalColumns1 As String, additionalColumns2 As String, _
                         Optional resultPrefix As String = "", _
                         Optional ByRef matchCount As Long = 0, _
                         Optional ByRef errorCount As Long = 0)
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
    Set wsMatch = GetOrCreateWorksheet(ApplyPrefix(resultPrefix, "Match Results"))
    Set wsError = GetOrCreateWorksheet(ApplyPrefix(resultPrefix, "Error Results"))
    
    ' Determine if we're in detail expansion mode
    Dim isDetailMode As Boolean
    isDetailMode = (detailSheet = "SHEET1" Or detailSheet = "SHEET2")
    
    If isDetailMode Then
        ' Detail expansion mode for Match Results
        WriteDetailMatchResults wsMatch, wsAllResults, ws1, ws2, idCol1, idCol2, detailSheet, _
                                dictDetailRows, dict1, dict2, dictAdditional1, dictAdditional2, _
                                additionalColumns1, additionalColumns2, idCol, sheet1TotalCol, sheet2TotalCol, matchCount
    Else
        ' Aggregated mode for Match Results (current behavior)
        WriteAggregatedMatchResults wsMatch, wsAllResults, idCol, lastCol, matchCount
    End If
    
    ' Error Results: always aggregated (keep current behavior)
    WriteErrorResults wsError, wsAllResults, idCol, diffCol, lastCol, errorCount
End Sub

Private Sub WriteAggregatedMatchResults(wsMatch As Worksheet, wsAllResults As Worksheet, idCol As Long, lastCol As Long, ByRef matchCount As Long)
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
    matchCount = 0
    
    For row = 2 To tblAll.Range.Rows.Count
        Dim isMatchValue As Variant
        isMatchValue = wsAllResults.Cells(row, 1).Value ' IsMatch column
        If (VarType(isMatchValue) = vbBoolean And isMatchValue = True) _
            Or (VarType(isMatchValue) = vbString And UCase(Trim(isMatchValue)) = "TRUE") Then
            ' Copy to Match Results (skip IsMatch and Difference columns)
            matchCol = 1
            For col = idCol To lastCol
                wsMatch.Cells(matchRow, matchCol).value = wsAllResults.Cells(row, col).value
                matchCol = matchCol + 1
            Next col
            matchRow = matchRow + 1
            matchCount = matchCount + 1
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
                                     idCol As Long, sheet1TotalCol As Long, sheet2TotalCol As Long, _
                                     ByRef matchCount As Long)
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
    matchCount = 0
    For row = 2 To tblAll.Range.Rows.Count
        Dim isMatchValue As Variant
        isMatchValue = wsAllResults.Cells(row, 1).Value ' IsMatch column
        If (VarType(isMatchValue) = vbBoolean And isMatchValue = True) _
            Or (VarType(isMatchValue) = vbString And UCase(Trim(isMatchValue)) = "TRUE") Then
            id = CStr(wsAllResults.Cells(row, idCol).value)
            total1 = wsAllResults.Cells(row, sheet1TotalCol).value
            total2 = wsAllResults.Cells(row, sheet2TotalCol).value
            matchCount = matchCount + 1
            
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

Private Sub WriteErrorResults(wsError As Worksheet, wsAllResults As Worksheet, idCol As Long, diffCol As Long, lastCol As Long, ByRef errorCount As Long)
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
    errorCount = 0
    
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
            errorCount = errorCount + 1
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
        
        ' Format number columns in Error Results (Difference, Sheet1 Total, Sheet2 Total - skip ID)
        If errorLastCol >= 3 Then
            ' Format Difference column (column 1)
            wsError.Range(wsError.Cells(2, 1), wsError.Cells(errorRow - 1, 1)).NumberFormat = "#,##0.00"
            ' Format Sheet1 Total and Sheet2 Total (columns 3-4)
            wsError.Range(wsError.Cells(2, 3), wsError.Cells(errorRow - 1, 4)).NumberFormat = "#,##0.00"
        End If
        
        wsError.Columns("A:" & ColLetter(errorLastCol)).AutoFit
    End If
End Sub

' Helper: Apply prefix to sheet name
Private Function ApplyPrefix(prefix As String, sheetName As String) As String
    If Len(Trim(prefix)) > 0 Then
        ApplyPrefix = Trim(prefix) & " " & sheetName
    Else
        ApplyPrefix = sheetName
    End If
End Function

' Helper: Log activity to unified Activity Log sheet
Private Sub LogActivity(operationType As String, workflowName As String, _
                        details As String, status As String, recordCount As Long)
    On Error GoTo ErrorHandler
    
    ' Get or create Activity Log sheet
    Dim wsLog As Worksheet
    Set wsLog = GetOrCreateWorksheet("Activity Log")
    
    ' Check if table exists, if not create header and table
    Dim tblLog As ListObject
    On Error Resume Next
    Set tblLog = wsLog.ListObjects("ActivityLogTable")
    On Error GoTo ErrorHandler
    
    If tblLog Is Nothing Then
        ' Create header row
        wsLog.Cells(1, 1).value = "Timestamp"
        wsLog.Cells(1, 2).value = "Operation"
        wsLog.Cells(1, 3).value = "Workflow"
        wsLog.Cells(1, 4).value = "Details"
        wsLog.Cells(1, 5).value = "Status"
        wsLog.Cells(1, 6).value = "Records"
        wsLog.Cells(1, 7).value = "User"
        
        ' Format header
        With wsLog.Range("A1:G1")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
        
        ' Create table with just header row
        Dim tableRange As Range
        Set tableRange = wsLog.Range("A1:G1")
        Set tblLog = wsLog.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        tblLog.Name = "ActivityLogTable"
        tblLog.TableStyle = "TableStyleMedium2"
    End If
    
    ' Add new row to table
    Dim newRow As ListRow
    Set newRow = tblLog.ListRows.Add
    
    ' Populate row data
    With newRow
        .Range(1, 1).value = Now() ' Timestamp
        .Range(1, 2).value = operationType
        .Range(1, 3).value = IIf(Len(Trim(workflowName)) > 0, workflowName, "N/A")
        .Range(1, 4).value = details
        .Range(1, 5).value = status
        .Range(1, 6).value = recordCount
        .Range(1, 7).value = Environ("USERNAME")
    End With
    
    ' Format timestamp column
    tblLog.ListColumns("Timestamp").DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    
    ' Color code by status
    If status = "Error" Then
        newRow.Range.Interior.Color = RGB(255, 200, 200) ' Light red
    End If
    
    ' Auto-fit columns
    wsLog.Columns("A:G").AutoFit
    
    Exit Sub
    
ErrorHandler:
    ' Silent fail - don't interrupt operations if logging fails
    Debug.Print "Logging error: " & Err.Description
End Sub

' Filter evaluation function - supports multiple conditions
Private Function EvaluateRowFilter(ws As Worksheet, row As Long, filterExpr As String) As Boolean
    ' Supports multiple comma-separated conditions, e.g. Amount<>0,Category="Loan"
    ' Handles =, <>, >, <, >=, <= for numeric and text columns
    Dim conds() As String
    Dim i As Long
    Dim cond As String
    Dim colName As String, op As String, val As String
    Dim colIdx As Long
    Dim cellVal As Variant
    Dim result As Boolean
    Dim cmpVal As Variant
    
    If Len(Trim(filterExpr)) = 0 Then
        EvaluateRowFilter = True
        Exit Function
    End If
    
    conds = Split(filterExpr, ",")
    For i = 0 To UBound(conds)
        cond = Trim(conds(i))
        ' Find operator (=, <>, >=, <=, >, <)
        If InStr(cond, ">=") > 0 Then
            op = ">="
        ElseIf InStr(cond, "<=") > 0 Then
            op = "<="
        ElseIf InStr(cond, "<>") > 0 Then
            op = "<>"
        ElseIf InStr(cond, ">") > 0 Then
            op = ">"
        ElseIf InStr(cond, "<") > 0 Then
            op = "<"
        ElseIf InStr(cond, "=") > 0 Then
            op = "="
        Else
            EvaluateRowFilter = False
            Exit Function
        End If
        colName = Trim(Left(cond, InStr(cond, op) - 1))
        val = Trim(Mid(cond, InStr(cond, op) + Len(op)))
        ' Remove quotes for string comparisons
        If Left(val, 1) = Chr(34) And Right(val, 1) = Chr(34) Then
            val = Mid(val, 2, Len(val) - 2)
        End If
        colIdx = FindColumnIndex(ws, colName)
        If colIdx = 0 Then
            EvaluateRowFilter = False
            Exit Function
        End If
        cellVal = ws.Cells(row, colIdx).Value
        ' Try numeric comparison if possible
        If IsNumeric(cellVal) And IsNumeric(val) Then
            cmpVal = CDbl(val)
            Select Case op
                Case "=": result = (CDbl(cellVal) = cmpVal)
                Case "<>": result = (CDbl(cellVal) <> cmpVal)
                Case ">": result = (CDbl(cellVal) > cmpVal)
                Case "<": result = (CDbl(cellVal) < cmpVal)
                Case ">=": result = (CDbl(cellVal) >= cmpVal)
                Case "<=": result = (CDbl(cellVal) <= cmpVal)
                Case Else: result = False
            End Select
        Else
            ' String comparison
            Select Case op
                Case "=": result = (CStr(cellVal) = val)
                Case "<>": result = (CStr(cellVal) <> val)
                Case ">": result = (CStr(cellVal) > val)
                Case "<": result = (CStr(cellVal) < val)
                Case ">=": result = (CStr(cellVal) >= val)
                Case "<=": result = (CStr(cellVal) <= val)
                Case Else: result = False
            End Select
        End If
        If Not result Then
            EvaluateRowFilter = False
            Exit Function
        End If
    Next i
    EvaluateRowFilter = True
End Function



