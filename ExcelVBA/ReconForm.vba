Option Explicit

' UserForm code-behind for ReconForm
' Instructions:
' 1. Insert a UserForm (Insert -> UserForm)
' 2. Name it "ReconForm"
' 3. Add the following controls:
'    - ComboBox1 (Name: cboWorkflow, ListStyle: fmListStylePlain)
'    - Label2 (Name: lblWorkflow, Caption: "Workflow:")
'    - CommandButton1 (Name: btnLoadSheet1, Caption: "Load Worksheet 1")
'    - CommandButton2 (Name: btnLoadSheet2, Caption: "Load Worksheet 2")
'    - CommandButton3 (Name: btnRunRecon, Caption: "Run Reconciliation")
'    - CommandButton4 (Name: btnRunTransform, Caption: "Run Transform")
'    - Label1 (Name: lblStatus, Caption: "Ready")
' 4. Copy this code into the UserForm code window

Private Sub btnLoadSheet1_Click()
    On Error GoTo ErrorHandler
    
    ImportExcelFile ThisWorkbook.Worksheets("Sheet1")
    lblStatus.Caption = "Sheet1 loaded successfully"
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error loading Sheet1"
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnLoadSheet2_Click()
    On Error GoTo ErrorHandler
    
    ImportExcelFile ThisWorkbook.Worksheets("Sheet2")
    lblStatus.Caption = "Sheet2 loaded successfully"
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error loading Sheet2"
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnRunRecon_Click()
    On Error GoTo ErrorHandler
    
    lblStatus.Caption = "Running reconciliation..."
    DoEvents
    
    RunReconciliation
    
    lblStatus.Caption = "Reconciliation completed"
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error during reconciliation"
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnRunTransform_Click()
    On Error GoTo ErrorHandler
    
    lblStatus.Caption = "Running transformation..."
    DoEvents
    
    RunTransform
    
    lblStatus.Caption = "Transformation completed"
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error during transformation"
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Excel Reconciliation Tool"
    lblStatus.Caption = "Ready"
    
    ' Populate workflow dropdown
    PopulateWorkflows
    
    ' Fix VBA resize bug - set form size explicitly
    Me.Width = 480
    Me.Height = 360
End Sub

Private Sub PopulateWorkflows()
    ' Clear existing items
    cboWorkflow.Clear
    
    ' Check if Workflows sheet exists
    Dim wsWorkflows As Worksheet
    On Error Resume Next
    Set wsWorkflows = ThisWorkbook.Worksheets("Workflows")
    On Error GoTo 0
    
    If wsWorkflows Is Nothing Then
        ' No Workflows sheet, add Default option
        cboWorkflow.AddItem "Default"
        cboWorkflow.ListIndex = 0
        ActiveWorkflow = "Default"
        Exit Sub
    End If
    
    ' Get WorkflowsTable (first table in the sheet)
    Dim tblWorkflows As ListObject
    On Error Resume Next
    If wsWorkflows.ListObjects.Count > 0 Then
        Set tblWorkflows = wsWorkflows.ListObjects(1)
    End If
    On Error GoTo 0
    
    If tblWorkflows Is Nothing Then
        ' No table in Workflows sheet, add Default option
        cboWorkflow.AddItem "Default"
        cboWorkflow.ListIndex = 0
        ActiveWorkflow = "Default"
        Exit Sub
    End If
    
    ' Find Workflow Name and Enabled columns (support both "Workflow Name" and "WorkflowName")
    Dim nameCol As Long, enabledCol As Long, col As Long
    For col = 1 To tblWorkflows.ListColumns.Count
        Select Case UCase(Trim(tblWorkflows.HeaderRowRange.Cells(1, col).Value))
            Case "WORKFLOW NAME", "WORKFLOWNAME"
                nameCol = col
            Case "ENABLED"
                enabledCol = col
        End Select
    Next col

    If nameCol = 0 Then
        ' Workflow Name column not found
        cboWorkflow.AddItem "Default"
        cboWorkflow.ListIndex = 0
        ActiveWorkflow = "Default"
        Exit Sub
    End If

    ' Populate combo box with workflow names where Enabled is TRUE
    Dim row As Long
    For row = 1 To tblWorkflows.ListRows.Count
        Dim workflowName As String
        workflowName = Trim(CStr(tblWorkflows.DataBodyRange.Cells(row, nameCol).Value))
        Dim enabledValue As String
        If enabledCol > 0 Then
            enabledValue = UCase(Trim(CStr(tblWorkflows.DataBodyRange.Cells(row, enabledCol).Value)))
        Else
            enabledValue = "TRUE" ' If no Enabled column, default to TRUE
        End If
        If Len(workflowName) > 0 And enabledValue = "TRUE" Then
            cboWorkflow.AddItem workflowName
        End If
    Next row

    ' Select first workflow if available
    If cboWorkflow.ListCount > 0 Then
        cboWorkflow.ListIndex = 0
        ActiveWorkflow = cboWorkflow.Value
    End If
End Sub

Private Sub cboWorkflow_Change()
    ' Update ActiveWorkflow when selection changes
    If cboWorkflow.ListIndex >= 0 Then
        ActiveWorkflow = cboWorkflow.value
    End If
End Sub
