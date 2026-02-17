Option Explicit

' Main reconciliation module
' Ported from C# WindowsForms ExcelReconApp

' Module-level variable to store active workflow
Public ActiveWorkflow As String

' Helper: Resolve workflow name with ActiveWorkflow fallback
Private Function ResolveWorkflowName(workflowName As String) As String
    If Len(workflowName) = 0 Then
        If Len(ActiveWorkflow) > 0 Then
            ResolveWorkflowName = ActiveWorkflow
        Else
            ResolveWorkflowName = "Default"
        End If
    Else
        ResolveWorkflowName = workflowName
    End If
End Function

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
    Dim workflowConfig As Object
    Dim reconConfigSheetName As String
    Dim resultPrefix As String
    Dim wsConfig As Worksheet
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim wsResults As Worksheet
    Dim allIDs As Collection
    Dim matchCount As Long
    Dim errorCount As Long
    Dim reconSettings As Object
    Dim reconData As Object
    
    ' Initialize workflow
    workflowName = ResolveWorkflowName(workflowName)
    Set workflowConfig = GetWorkflowConfig(workflowName)
    reconConfigSheetName = workflowConfig("ReconConfigSheet")
    resultPrefix = workflowConfig("ResultSheetPrefix")
    
    ' Validate and prepare
    If Not ValidateSetup(reconConfigSheetName, workflowName) Then Exit Sub
    ClearResults resultPrefix
    
    ' Load configuration and worksheets
    Set wsConfig = ThisWorkbook.Worksheets(reconConfigSheetName)
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Sheet2")
    Set wsResults = GetOrCreateWorksheet(ApplyPrefix(resultPrefix, "All Results"))
    
    ' Read all reconciliation settings
    Set reconSettings = ReadReconciliationConfig(wsConfig, workflowName)
    
    ' Build dictionaries and process data
    Set reconData = BuildReconciliationData(ws1, ws2, reconSettings)
    Set allIDs = GetUniqueIDs(reconData("Dict1"), reconData("Dict2"))
    
    ' Write and split results
    WriteResults wsResults, allIDs, reconData("Dict1"), reconData("Dict2"), reconData("DictAdditional1"), reconSettings("AdditionalColumns1"), reconSettings("Tolerance"), reconSettings("SortOrder")
    SplitResults wsResults, ws1, ws2, reconSettings("IdCol1"), reconSettings("IdCol2"), reconSettings("DetailSheet"), reconData("DictDetailRows"), reconData("Dict1"), reconData("Dict2"), reconData("DictAdditional1"), reconData("DictAdditional2"), reconSettings("AdditionalColumns1"), reconSettings("AdditionalColumns2"), resultPrefix, matchCount, errorCount
    
    ' Log and report success
    LogReconciliationSuccess workflowName, reconConfigSheetName, matchCount, errorCount, reconSettings("Tolerance"), reconSettings("DetailSheet"), allIDs.Count
End Sub

Private Function ReadReconciliationConfig(wsConfig As Worksheet, workflowName As String) As Object
    Dim settings As Object
    Dim toleranceStr As String
    
    Set settings = CreateObject("Scripting.Dictionary")
    
    ' Read ID and column configurations
    settings("IdCol1") = GetConfigValue(wsConfig, "Sheet1 ID Column", workflowName)
    settings("IdCol2") = GetConfigValue(wsConfig, "Sheet2 ID Column", workflowName)
    settings("ValueColumns1") = GetConfigValue(wsConfig, "Sheet1 Value Columns", workflowName)
    settings("ValueColumns2") = GetConfigValue(wsConfig, "Sheet2 Value Columns", workflowName)
    settings("AdditionalColumns1") = GetConfigValue(wsConfig, "Sheet1 Additional Columns", workflowName)
    settings("AdditionalColumns2") = GetConfigValue(wsConfig, "Sheet2 Additional Columns", workflowName)
    
    ' Read filter settings
    settings("Sheet1Filter") = GetConfigValue(wsConfig, "Sheet1 Filter", workflowName)
    settings("Sheet2Filter") = GetConfigValue(wsConfig, "Sheet2 Filter", workflowName)
    
    ' Read tolerance with default
    toleranceStr = GetConfigValue(wsConfig, "Tolerance", workflowName)
    If IsNumeric(toleranceStr) And CDbl(toleranceStr) >= 0 Then
        settings("Tolerance") = CDbl(toleranceStr)
    Else
        settings("Tolerance") = 0.01
    End If
    
    ' Read detail and sort settings
    settings("DetailSheet") = Trim(UCase(GetConfigValue(wsConfig, "Detail Sheet", workflowName)))
    settings("SortOrder") = GetConfigValue(wsConfig, "Sort Order", workflowName)
    
    Set ReadReconciliationConfig = settings
End Function

Private Function BuildReconciliationData(ws1 As Worksheet, ws2 As Worksheet, settings As Object) As Object
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    
    ' Build totals dictionaries with filtering
    Set data("Dict1") = BuildTotalsDictionary(ws1, settings("IdCol1"), settings("ValueColumns1"), settings("Sheet1Filter"))
    Set data("Dict2") = BuildTotalsDictionary(ws2, settings("IdCol2"), settings("ValueColumns2"), settings("Sheet2Filter"))
    
    ' Build additional columns dictionaries
    Set data("DictAdditional1") = BuildAdditionalColumnsDictionary(ws1, settings("IdCol1"), settings("AdditionalColumns1"))
    Set data("DictAdditional2") = BuildAdditionalColumnsDictionary(ws2, settings("IdCol2"), settings("AdditionalColumns2"))
    
    ' Build detail rows dictionary if needed
    Set data("DictDetailRows") = CreateObject("Scripting.Dictionary")
    If settings("DetailSheet") = "SHEET1" Then
        Set data("DictDetailRows") = BuildAllDetailRowsDictionary(ws1, settings("IdCol1"), settings("AdditionalColumns1"), settings("Sheet1Filter"))
    ElseIf settings("DetailSheet") = "SHEET2" Then
        Set data("DictDetailRows") = BuildAllDetailRowsDictionary(ws2, settings("IdCol2"), settings("AdditionalColumns2"), settings("Sheet2Filter"))
    End If
    
    Set BuildReconciliationData = data
End Function

Private Sub LogReconciliationSuccess(workflowName As String, reconConfigSheetName As String, _
                                     matchCount As Long, errorCount As Long, _
                                     tolerance As Double, detailSheet As String, totalRecords As Long)
    Dim details As String
    
    details = "Config: " & reconConfigSheetName & " | Matches: " & matchCount & " | Errors: " & errorCount & _
              " | Tolerance: " & tolerance & " | Detail Mode: " & IIf(Len(Trim(detailSheet)) > 0, detailSheet, "None")
    LogActivity "Reconciliation", workflowName, details, "Success", totalRecords
    
    MsgBox "Reconciliation completed! " & totalRecords & " records processed (" & matchCount & " matches, " & errorCount & " errors).", vbInformation, "Success"
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
    ' Declare all variables at the top
    Dim additionalColNames() As String
    Dim hasAdditionalCols As Boolean
    Dim i As Long
    Dim lastHeaderCol As Long
    Dim rowData As Collection
    Dim lastRow As Long
    Dim r As Long
    
    ' Parse additional column names
    hasAdditionalCols = (Len(Trim(additionalColumnsStr)) > 0)

    If hasAdditionalCols Then
        additionalColNames = Split(additionalColumnsStr, ",")
        For i = 0 To UBound(additionalColNames)
            additionalColNames(i) = Trim(additionalColNames(i))
        Next i
    End If

    ' Clear previous data and formatting in wsResults
    wsResults.Cells.Clear

    ' Write header
    lastHeaderCol = WriteResultsHeader(wsResults, additionalColNames, hasAdditionalCols)

    ' Build and sort row data
    Set rowData = BuildResultsRowData(allIDs, dict1, dict2, dictAdditional, additionalColNames, hasAdditionalCols, tolerance)
    SortRowData rowData, sortOrder, wsResults

    ' Write data rows
    lastRow = WriteResultsDataRows(wsResults, rowData, additionalColNames, hasAdditionalCols)

    ' Force IsMatch column to Boolean (TRUE/FALSE)
    For r = 2 To lastRow
        wsResults.Cells(r, 1).Value = CBool(wsResults.Cells(r, 1).Value)
    Next r

    ' Create and format table
    CreateAndFormatResultsTable wsResults, lastRow, lastHeaderCol
End Sub

Private Function WriteResultsHeader(wsResults As Worksheet, additionalColNames() As String, hasAdditionalCols As Boolean) As Long
    Dim col As Long
    Dim i As Long
    Dim lastHeaderCol As Long
    
    ' Write header - new order: IsMatch, Difference, ID, Sheet1 Total, Sheet2 Total, then additional columns
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
    Dim tableRange As Range
    
    ' Format number columns (Difference, Sheet1 Total, Sheet2 Total - skip ID in column 3)
    If lastRow >= 2 Then
        ' Column 2: Difference
        wsResults.Range(wsResults.Cells(2, 2), wsResults.Cells(lastRow, 2)).NumberFormat = "#,##0.00"
        ' Columns 4-5: Sheet1 Total, Sheet2 Total
        wsResults.Range(wsResults.Cells(2, 4), wsResults.Cells(lastRow, 5)).NumberFormat = "#,##0.00"
    End If
    
    ' Convert to Table
    If lastRow >= 2 Then
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
    Dim wsResults As Worksheet
    Dim wsMatch As Worksheet
    Dim wsError As Worksheet
    
    ' Clear All Results sheet
    Set wsResults = GetOrCreateWorksheet(ApplyPrefix(resultPrefix, "All Results"))
    
    ' Delete existing table if present
    On Error Resume Next
    wsResults.ListObjects("ReconResultsTable").Delete
    On Error GoTo 0
    
    ' Clear all cells
    wsResults.Cells.Clear
    
    ' Clear Match Results sheet
    Set wsMatch = GetOrCreateWorksheet(ApplyPrefix(resultPrefix, "Match Results"))
    On Error Resume Next
    wsMatch.ListObjects("MatchResultsTable").Delete
    On Error GoTo 0
    wsMatch.Cells.Clear
    
    ' Clear Error Results sheet
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
    Dim filePath As String
    Dim wb As Workbook
    Dim rowCount As Long
    Dim fileName As String
    Dim pos As Long
    Dim details As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.Title = "Select Excel or CSV File to Import"
    fd.Filters.Clear
    fd.Filters.Add "Excel and CSV Files", "*.xlsx;*.xls;*.xlsm;*.csv"
    fd.Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
    fd.Filters.Add "CSV Files", "*.csv"
    
    If fd.Show = -1 Then
        filePath = fd.SelectedItems(1)
        
        On Error GoTo ImportError
        
        ' Clear target sheet
        targetSheet.Cells.Clear
        
        ' Open file and copy data
        Set wb = Workbooks.Open(filePath, ReadOnly:=True)
        
        ' Copy first worksheet
        wb.Worksheets(1).UsedRange.Copy targetSheet.Range("A1")
        
        wb.Close False
        
        ' Count rows imported (excluding header)
        rowCount = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row - 1
        
        ' Extract filename from path
        pos = InStrRev(filePath, "\")
        If pos > 0 Then
            fileName = Mid(filePath, pos + 1)
        Else
            fileName = filePath
        End If
        
        ' Log the import
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

' Helper: Save current sheet to new workbook
' Assign this to a button on any sheet - it will save that specific sheet
Public Sub SaveSheetToFile()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim filePath As Variant
    Dim fileName As String
    Dim sheetName As String
    
    ' Get the sheet where the button is located
    On Error Resume Next
    Set ws = Application.Caller.Parent.Parent
    On Error GoTo 0
    
    ' Fallback to active sheet if we can't determine button location
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    sheetName = ws.Name
    
    ' Use GetSaveAsFilename for proper filter support
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=sheetName & ".xls", _
        FileFilter:="Excel 97-2003 Workbook (*.xls), *.xls", _
        Title:="Save " & sheetName & " as Excel 97-2003 Workbook")
    
    ' Check if user cancelled
    If filePath <> False Then
        ' Ensure .xls extension
        If LCase(Right(filePath, 4)) <> ".xls" Then
            filePath = filePath & ".xls"
        End If
        
        On Error GoTo SaveError
        
        ' Copy sheet to new workbook
        ws.Copy
        Set wb = ActiveWorkbook
        
        ' Save as Excel 97-2003 format
        wb.SaveAs fileName:=filePath, FileFormat:=xlExcel8
        
        ' Close the new workbook
        wb.Close SaveChanges:=False
        
        ' Extract filename from path
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
        
        ' Log the save
        LogActivity "Export", "", "Sheet: " & sheetName & " | File: " & fileName, "Success", 0
        
        MsgBox "Sheet '" & sheetName & "' saved successfully to:" & vbCrLf & filePath, vbInformation, "Success"
    End If
    
    Exit Sub
    
SaveError:
    LogActivity "Export", "", "Error saving sheet: " & sheetName & " - " & Err.Description, "Error", 0
    MsgBox "Error saving sheet: " & Err.Description, vbCritical, "Error"
End Sub

' ========================================
' TRANSFORMATION MODULE
' ========================================

' Main transformation routine - reads Transform Config sheet and creates transformed sheets
Public Sub RunTransform(Optional workflowName As String = "")
    On Error GoTo ErrorHandler
    
    Dim startTime As Double
    Dim workflowConfig As Object
    Dim transformConfigSheetName As String
    Dim resultPrefix As String
    Dim wsTransformConfig As Worksheet
    Dim transformSettings As Object
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim columnDefs As Collection
    Dim rowCount As Long
    
    startTime = Timer
    
    ' Initialize workflow
    workflowName = ResolveWorkflowName(workflowName)
    Set workflowConfig = GetWorkflowConfig(workflowName)
    transformConfigSheetName = workflowConfig("TransformConfigSheet")
    resultPrefix = workflowConfig("ResultSheetPrefix")
    
    ' Validate Transform Config exists
    On Error Resume Next
    Set wsTransformConfig = ThisWorkbook.Worksheets(transformConfigSheetName)
    On Error GoTo ErrorHandler
    
    If wsTransformConfig Is Nothing Then
        MsgBox "Transform Config sheet '" & transformConfigSheetName & "' not found. Please create it first.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Read transform settings and apply prefix to sheet names
    Set transformSettings = ReadTransformSettings(wsTransformConfig, resultPrefix, workflowName)
    
    ' Validate source sheet and create target sheet
    Set wsSource = ValidateSourceSheet(transformSettings("SourceSheet"))
    If wsSource Is Nothing Then Exit Sub
    
    Set wsTarget = GetOrCreateWorksheet(transformSettings("TargetSheet"))
    wsTarget.Cells.Clear
    
    ' Read column definitions
    Set columnDefs = ReadTransformColumnDefinitions(wsTransformConfig, workflowName)
    If columnDefs Is Nothing Then Exit Sub
    
    ' Execute transformation
    ApplyTransformation wsSource, wsTarget, columnDefs, transformSettings("Filter"), transformSettings("SortOrder")
    
    ' Calculate results and log success
    rowCount = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).row - 1
    LogTransformSuccess workflowName, transformSettings, rowCount, Timer - startTime
    
    ' Show results and return focus to form
    wsTarget.Activate
    ReconForm.Show vbModeless
    
    Exit Sub
    
ErrorHandler:
    LogActivity "Transformation", workflowName, "Error: " & Err.Description, "Error", 0
    MsgBox "Error during transformation: " & Err.Description, vbCritical, "Error"
End Sub

Private Function ReadTransformSettings(wsTransformConfig As Worksheet, resultPrefix As String, Optional workflowName As String = "") As Object
    Dim settings As Object
    Dim transformName As String
    Dim sourceSheetName As String
    Dim targetSheetName As String
    
    Set settings = CreateObject("Scripting.Dictionary")
    
    ' Read from TransformSettingsTable (with optional workflow filtering)
    transformName = GetConfigValue(wsTransformConfig, "Name", workflowName)
    sourceSheetName = GetConfigValue(wsTransformConfig, "Source Sheet", workflowName)
    targetSheetName = GetConfigValue(wsTransformConfig, "Target Sheet", workflowName)
    
    If Len(transformName) = 0 Then transformName = "Transform"
    
    If Len(sourceSheetName) = 0 Or Len(targetSheetName) = 0 Then
        MsgBox "Transform configuration incomplete. Source Sheet and Target Sheet are required in TransformSettingsTable." & IIf(Len(workflowName) > 0, " (Workflow: " & workflowName & ")", ""), vbCritical, "Error"
        Set ReadTransformSettings = Nothing
        Exit Function
    End If
    
    ' Apply workflow prefix to sheet names
    settings("Name") = transformName
    settings("SourceSheet") = ApplyPrefix(resultPrefix, sourceSheetName)
    settings("TargetSheet") = ApplyPrefix(resultPrefix, targetSheetName)
    settings("Filter") = GetConfigValue(wsTransformConfig, "Filter", workflowName)
    settings("SortOrder") = GetConfigValue(wsTransformConfig, "Sort Order", workflowName)
    
    Set ReadTransformSettings = settings
End Function

Private Function ValidateSourceSheet(sourceSheetName As String) As Worksheet
    Dim wsSource As Worksheet
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sourceSheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Source sheet '" & sourceSheetName & "' not found.", vbCritical, "Error"
    End If
    
    Set ValidateSourceSheet = wsSource
End Function

Private Function ReadTransformColumnDefinitions(wsTransformConfig As Worksheet, Optional workflowName As String = "") As Collection
    Dim columnDefs As Collection
    Dim tblColumns As ListObject
    Dim orderColIdx As Long
    Dim targetColIdx As Long
    Dim typeColIdx As Long
    Dim sourceColIdx As Long
    Dim workflowColIdx As Long
    Dim col As Long
    Dim row As Long
    Dim colDef As Object
    Dim sourceValue As String
    Dim workflowNameUpper As String
    
    Set columnDefs = New Collection
    workflowNameUpper = UCase(Trim(workflowName))
    
    ' Get second table (TransformColumnsTable)
    On Error Resume Next
    If wsTransformConfig.ListObjects.Count >= 2 Then
        Set tblColumns = wsTransformConfig.ListObjects(2)
    End If
    On Error GoTo 0
    
    If tblColumns Is Nothing Then
        MsgBox "TransformColumnsTable (second table) not found in Transform Config sheet.", vbCritical, "Error"
        Set ReadTransformColumnDefinitions = Nothing
        Exit Function
    End If
    
    ' Find column indexes
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
            Case "WORKFLOW"
                workflowColIdx = col
        End Select
    Next col
    
    If orderColIdx = 0 Or targetColIdx = 0 Or typeColIdx = 0 Or sourceColIdx = 0 Then
        MsgBox "TransformColumnsTable must have columns: Order, Target Column, Type, Source", vbCritical, "Error"
        Set ReadTransformColumnDefinitions = Nothing
        Exit Function
    End If
    
    ' Read column definitions
    For row = 1 To tblColumns.ListRows.Count
        ' Filter by workflow if specified and Workflow column exists
        If Len(workflowName) > 0 And workflowColIdx > 0 Then
            Dim rowWorkflow As String
            rowWorkflow = UCase(Trim(CStr(tblColumns.DataBodyRange.Cells(row, workflowColIdx).value)))
            If rowWorkflow <> workflowNameUpper Then
                GoTo NextColumnRow  ' Skip rows that don't match workflow
            End If
        End If
        
        Set colDef = CreateObject("Scripting.Dictionary")
        
        ' Validate Order value
        Dim orderValue As Variant
        orderValue = tblColumns.DataBodyRange.Cells(row, orderColIdx).value
        
        If IsEmpty(orderValue) Or Not IsNumeric(orderValue) Then
            MsgBox "Row " & row & " in TransformColumnsTable has an invalid Order value. Order must be a number between 1 and 16384." & vbCrLf & vbCrLf & _
                   "Current value: " & IIf(IsEmpty(orderValue), "(empty)", CStr(orderValue)), vbCritical, "Config Error"
            Set ReadTransformColumnDefinitions = Nothing
            Exit Function
        End If
        
        Dim orderNum As Long
        orderNum = CLng(orderValue)
        
        If orderNum < 1 Or orderNum > 16384 Then
            MsgBox "Row " & row & " in TransformColumnsTable has Order = " & orderNum & "." & vbCrLf & _
                   "Order must be between 1 and 16384.", vbCritical, "Config Error"
            Set ReadTransformColumnDefinitions = Nothing
            Exit Function
        End If
        
        colDef("Order") = orderNum
        colDef("TargetColumn") = Trim(CStr(tblColumns.DataBodyRange.Cells(row, targetColIdx).value))
        colDef("Type") = Trim(UCase(CStr(tblColumns.DataBodyRange.Cells(row, typeColIdx).value)))
        
        sourceValue = Trim(CStr(tblColumns.DataBodyRange.Cells(row, sourceColIdx).value))
        
        ' For FORMULA type, ensure it starts with =
        If colDef("Type") = "FORMULA" Then
            If Left(sourceValue, 1) <> "=" Then
                sourceValue = "=" & sourceValue
            End If
        End If
        
        colDef("Source") = sourceValue
        columnDefs.Add colDef
        
NextColumnRow:
    Next row
    
    If columnDefs.Count = 0 Then
        MsgBox "No column definitions found in TransformColumnsTable." & IIf(Len(workflowName) > 0, " (Workflow: " & workflowName & ")", ""), vbCritical, "Error"
        Set ReadTransformColumnDefinitions = Nothing
        Exit Function
    End If
    
    Set ReadTransformColumnDefinitions = columnDefs
End Function

Private Sub LogTransformSuccess(workflowName As String, transformSettings As Object, _
                                rowCount As Long, duration As Double)
    Dim details As String
    
    details = "Transform: " & transformSettings("Name") & " | Source: " & transformSettings("SourceSheet") & _
              " | Target: " & transformSettings("TargetSheet") & " | Duration: " & Format(duration, "0.0") & "s"
    LogActivity "Transformation", workflowName, details, "Success", rowCount
    
    MsgBox "Transformation completed successfully!" & vbCrLf & _
           "Transform: " & transformSettings("Name") & vbCrLf & _
           "Source: " & transformSettings("SourceSheet") & vbCrLf & _
           "Target: " & transformSettings("TargetSheet"), vbInformation, "Success"
End Sub

Private Sub ApplyTransformation(wsSource As Worksheet, wsTarget As Worksheet, columnDefs As Collection, Optional filterExpr As String = "", Optional sortOrder As String = "")
    Dim colDef As Object
    Dim col As Long
    Dim lastSourceRow As Long
    Dim sourceRow As Long
    Dim targetRow As Long
    Dim sourceColIndex As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim tableRange As Range
    Dim tbl As ListObject
    
    ' Write headers
    For Each colDef In columnDefs
        col = colDef("Order")
        wsTarget.Cells(1, col).value = colDef("TargetColumn")
    Next colDef
    
    ' Determine source data range
    lastSourceRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).row
    
    If lastSourceRow < 2 Then
        MsgBox "Source sheet has no data rows.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    ' Populate rows with EXISTING and STATIC columns only (applying filter)
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
    lastCol = columnDefs.Count
    lastRow = targetRow ' Use actual target row count after filtering
    
    Set tableRange = wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(lastRow, lastCol))
    
    ' Delete existing table if present
    On Error Resume Next
    wsTarget.ListObjects(1).Delete
    On Error GoTo 0
    
    ' Create new table
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
    Dim tbl As ListObject
    Dim sortSpecs() As String
    Dim i As Long
    Dim spec As String
    Dim colName As String
    Dim direction As Long
    Dim colIndex As Long
    
    If Len(Trim(sortOrder)) = 0 Then Exit Sub
    
    ' Get the table on this worksheet
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set tbl = ws.ListObjects(1)
    
    ' Parse sort specifications
    sortSpecs = Split(sortOrder, ",")
    
    ' Clear existing sort
    tbl.Sort.SortFields.Clear
    
    ' Add each sort field
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
    Dim tblAll As ListObject
    Dim col As Long
    Dim matchCol As Long
    Dim row As Long
    Dim matchRow As Long
    Dim isMatch As Boolean
    Dim isMatchValue As Variant
    Dim matchLastCol As Long
    Dim matchTableRange As Range
    
    ' Current behavior: Copy aggregated match results
    Set tblAll = wsAllResults.ListObjects("ReconResultsTable")
    
    ' Copy headers for Match Results (skip IsMatch and Difference)
    matchCol = 1
    For col = idCol To lastCol
        wsMatch.Cells(1, matchCol).value = wsAllResults.Cells(1, col).value
        matchCol = matchCol + 1
    Next col
    
    ' Loop through data rows and copy matches
    matchRow = 2
    matchCount = 0
    
    For row = 2 To tblAll.Range.Rows.Count
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
        matchLastCol = wsMatch.Cells(1, wsMatch.Columns.Count).End(xlToLeft).Column
        
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
    Dim additionalCols1() As String
    Dim additionalCols2() As String
    Dim hasAdditional1 As Boolean
    Dim hasAdditional2 As Boolean
    Dim lastHeaderCol As Long
    Dim row As Long
    Dim matchRow As Long
    Dim id As String
    Dim total1 As Double
    Dim total2 As Double
    Dim isMatchValue As Variant
    
    Set tblAll = wsAllResults.ListObjects("ReconResultsTable")
    
    ' Parse and prepare additional columns
    ParseAdditionalColumns additionalColumns1, additionalColumns2, additionalCols1, additionalCols2, hasAdditional1, hasAdditional2
    
    ' Build and format header
    lastHeaderCol = WriteDetailMatchHeader(wsMatch, detailSheet, additionalCols1, additionalCols2, hasAdditional1, hasAdditional2)
    
    ' Write detail rows for matched IDs
    matchRow = 2
    matchCount = 0
    For row = 2 To tblAll.Range.Rows.Count
        isMatchValue = wsAllResults.Cells(row, 1).Value
        If (VarType(isMatchValue) = vbBoolean And isMatchValue = True) _
            Or (VarType(isMatchValue) = vbString And UCase(Trim(isMatchValue)) = "TRUE") Then
            id = CStr(wsAllResults.Cells(row, idCol).value)
            total1 = wsAllResults.Cells(row, sheet1TotalCol).value
            total2 = wsAllResults.Cells(row, sheet2TotalCol).value
            matchCount = matchCount + 1
            
            ProcessMatchedDetailRows wsMatch, matchRow, id, total1, total2, detailSheet, _
                                    dictDetailRows, dictAdditional1, dictAdditional2, _
                                    additionalCols1, additionalCols2, hasAdditional1, hasAdditional2
        End If
    Next row
    
    ' Create and format table
    If matchRow > 2 Then
        FormatDetailMatchTable wsMatch, matchRow, lastHeaderCol, detailSheet, additionalCols1, additionalCols2, hasAdditional1, hasAdditional2
    End If
End Sub

Private Sub ProcessMatchedDetailRows(wsMatch As Worksheet, ByRef matchRow As Long, _
                                     id As String, total1 As Double, total2 As Double, _
                                     detailSheet As String, dictDetailRows As Object, _
                                     dictAdditional1 As Object, dictAdditional2 As Object, _
                                     additionalCols1() As String, additionalCols2() As String, _
                                     hasAdditional1 As Boolean, hasAdditional2 As Boolean)
    Dim detailRows As Collection
    Dim detailRow As Object
    Dim summaryDict As Object
    
    If Not dictDetailRows.Exists(id) Then Exit Sub
    
    Set detailRows = dictDetailRows(id)
    Set summaryDict = GetSummaryDict(detailSheet, id, dictAdditional1, dictAdditional2)
    
    For Each detailRow In detailRows
        WriteDetailMatchRow wsMatch, matchRow, id, detailRow, summaryDict, _
                           detailSheet, additionalCols1, additionalCols2, _
                           hasAdditional1, hasAdditional2, total1, total2
        matchRow = matchRow + 1
    Next detailRow
End Sub

Private Sub ParseAdditionalColumns(additionalColumns1 As String, additionalColumns2 As String, _
                                   ByRef additionalCols1() As String, ByRef additionalCols2() As String, _
                                   ByRef hasAdditional1 As Boolean, ByRef hasAdditional2 As Boolean)
    Dim i As Long
    
    hasAdditional1 = (Len(Trim(additionalColumns1)) > 0)
    hasAdditional2 = (Len(Trim(additionalColumns2)) > 0)
    
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
End Sub

Private Function WriteDetailMatchHeader(wsMatch As Worksheet, detailSheet As String, _
                                        additionalCols1() As String, additionalCols2() As String, _
                                        hasAdditional1 As Boolean, hasAdditional2 As Boolean) As Long
    Dim col As Long
    Dim i As Long
    
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
    
    WriteDetailMatchHeader = col - 1
    
    ' Format header
    With wsMatch.Range(wsMatch.Cells(1, 1), wsMatch.Cells(1, WriteDetailMatchHeader))
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
End Function

Private Function GetSummaryDict(detailSheet As String, id As String, _
                                dictAdditional1 As Object, dictAdditional2 As Object) As Object
    Set GetSummaryDict = Nothing
    
    If detailSheet = "SHEET1" Then
        If dictAdditional2.Exists(id) Then Set GetSummaryDict = dictAdditional2(id)
    Else
        If dictAdditional1.Exists(id) Then Set GetSummaryDict = dictAdditional1(id)
    End If
End Function

Private Sub WriteDetailMatchRow(wsMatch As Worksheet, matchRow As Long, id As String, _
                               detailRow As Object, summaryDict As Object, _
                               detailSheet As String, additionalCols1() As String, additionalCols2() As String, _
                               hasAdditional1 As Boolean, hasAdditional2 As Boolean, _
                               total1 As Double, total2 As Double)
    Dim col As Long
    Dim i As Long
    
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
    
    ' Write totals
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
End Sub

Private Sub FormatDetailMatchTable(wsMatch As Worksheet, matchRow As Long, lastHeaderCol As Long, _
                                   detailSheet As String, additionalCols1() As String, additionalCols2() As String, _
                                   hasAdditional1 As Boolean, hasAdditional2 As Boolean)
    Dim matchTableRange As Range
    Dim totalCol1 As Long
    Dim totalCol2 As Long
    
    Set matchTableRange = wsMatch.Range(wsMatch.Cells(1, 1), wsMatch.Cells(matchRow - 1, lastHeaderCol))
    
    On Error Resume Next
    wsMatch.ListObjects("MatchResultsTable").Delete
    On Error GoTo 0
    
    wsMatch.ListObjects.Add(xlSrcRange, matchTableRange, , xlYes).Name = "MatchResultsTable"
    wsMatch.ListObjects("MatchResultsTable").TableStyle = "TableStyleMedium3"
    
    ' Format number columns (totals)
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
End Sub

Private Sub WriteErrorResults(wsError As Worksheet, wsAllResults As Worksheet, idCol As Long, diffCol As Long, lastCol As Long, ByRef errorCount As Long)
    ' Error Results: always aggregated
    Dim tblAll As ListObject
    Dim col As Long, errorCol As Long
    Dim row As Long, errorRow As Long
    Dim isMatch As Boolean
    Dim errorLastCol As Long
    Dim errorTableRange As Range
    
    Set tblAll = wsAllResults.ListObjects("ReconResultsTable")
    
    ' Copy headers for Error Results (skip IsMatch only, keep Difference)
    errorCol = 1
    wsError.Cells(1, errorCol).value = "Difference": errorCol = errorCol + 1
    For col = idCol To lastCol
        wsError.Cells(1, errorCol).value = wsAllResults.Cells(1, col).value
        errorCol = errorCol + 1
    Next col
    
    ' Loop through data rows and copy errors
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
        errorLastCol = wsError.Cells(1, wsError.Columns.Count).End(xlToLeft).Column
        
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
    Dim tblLog As ListObject
    Dim tableRange As Range
    Dim newRow As ListRow
    
    Set wsLog = GetOrCreateWorksheet("Activity Log")
    
    ' Check if table exists, if not create header and table
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
        Set tableRange = wsLog.Range("A1:G1")
        Set tblLog = wsLog.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        tblLog.Name = "ActivityLogTable"
        tblLog.TableStyle = "TableStyleMedium2"
    End If
    
    ' Add new row to table
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

' Filter evaluation function - supports multiple conditions and functions
Private Function EvaluateRowFilter(ws As Worksheet, row As Long, filterExpr As String) As Boolean
    ' Supports multiple comma-separated conditions, e.g. Amount<>0,Category="Loan"
    ' Supports functions: ABS(columnName)>0.02
    ' Handles =, <>, >, <, >=, <= for numeric and text columns
    Dim conds() As String
    Dim i As Long
    Dim cond As String
    Dim colName As String, op As String, val As String
    Dim colIdx As Long
    Dim cellVal As Variant
    Dim result As Boolean
    Dim cmpVal As Variant
    Dim funcName As String
    Dim leftSide As String
    
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
        
        leftSide = Trim(Left(cond, InStr(cond, op) - 1))
        val = Trim(Mid(cond, InStr(cond, op) + Len(op)))
        
        ' Remove quotes for string comparisons
        If Left(val, 1) = Chr(34) And Right(val, 1) = Chr(34) Then
            val = Mid(val, 2, Len(val) - 2)
        End If
        
        ' Check for function syntax: FUNC(ColumnName)
        If InStr(leftSide, "(") > 0 And InStr(leftSide, ")") > 0 Then
            funcName = UCase(Trim(Left(leftSide, InStr(leftSide, "(") - 1)))
            colName = Trim(Mid(leftSide, InStr(leftSide, "(") + 1, InStr(leftSide, ")") - InStr(leftSide, "(") - 1))
        Else
            funcName = ""
            colName = leftSide
        End If
        
        ' Get column value
        colIdx = FindColumnIndex(ws, colName)
        If colIdx = 0 Then
            EvaluateRowFilter = False
            Exit Function
        End If
        cellVal = ws.Cells(row, colIdx).Value
        
        ' Apply function if specified
        If Len(funcName) > 0 Then
            cellVal = ApplyFilterFunction(funcName, cellVal)
        End If
        
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

Private Function ApplyFilterFunction(funcName As String, value As Variant) As Variant
    ' Apply supported functions to filter values
    Select Case UCase(Trim(funcName))
        Case "ABS"
            If IsNumeric(value) Then
                ApplyFilterFunction = Abs(CDbl(value))
            Else
                ApplyFilterFunction = value
            End If
        Case Else
            ' Unknown function - return value unchanged
            ApplyFilterFunction = value
    End Select
End Function



