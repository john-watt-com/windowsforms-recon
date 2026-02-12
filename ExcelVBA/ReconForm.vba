Option Explicit

' UserForm code-behind for ReconForm
' Instructions:
' 1. Insert a UserForm (Insert -> UserForm)
' 2. Name it "ReconForm"
' 3. Add the following controls:
'    - CommandButton1 (Name: btnLoadSheet1, Caption: "Load Worksheet 1")
'    - CommandButton2 (Name: btnLoadSheet2, Caption: "Load Worksheet 2")
'    - CommandButton3 (Name: btnRunRecon, Caption: "Run Reconciliation")
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

Private Sub UserForm_Initialize()
    Me.Caption = "Excel Reconciliation Tool"
    lblStatus.Caption = "Ready"
End Sub
