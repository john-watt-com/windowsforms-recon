# Proposed Configuration Format
## After Phase 1 Implementation (Filtering + Sorting for Transform)

---

## Workflows Table
**Sheet: "Workflows"**

**WorkflowsTable:**
```
Workflow Name    | Result Sheet Prefix
-----------------|--------------------
Fee Recon        | Fee
Loan Recon       | Loan
Interest Recon   | Interest
```

*Note: All workflows use hardcoded sheet names ("Recon Config" and "Transform Config"). Workflow Name identifies which config rows to read.*

---

## Recon Config (All Scenarios in One Sheet)

**Sheet: "Recon Config"**

**ReconConfigTable:**
```
Workflow         | Setting                      | Value
-----------------|------------------------------|------------------------
Fee Recon        | Sheet1 ID Column            | AccountID
Fee Recon        | Sheet2 ID Column            | AccountID
Fee Recon        | Sheet1 Value Columns        | Amount
Fee Recon        | Sheet2 Value Columns        | Amount
Fee Recon        | Sheet1 Additional Columns   | AccountName, Type
Fee Recon        | Sheet2 Additional Columns   | AccountName, Type
Fee Recon        | Sheet1 Filter               | Type="Fee"
Fee Recon        | Sheet2 Filter               | Type="Fee"
Fee Recon        | Tolerance                   | 0.01
Fee Recon        | Detail Sheet                |
Fee Recon        | Sort Order                  | AccountID ASC
Loan Recon       | Sheet1 ID Column            | LoanID
Loan Recon       | Sheet2 ID Column            | LoanID
Loan Recon       | Sheet1 Value Columns        | +Principal, +Interest, -Fees
Loan Recon       | Sheet2 Value Columns        | TotalAmount
Loan Recon       | Sheet1 Additional Columns   | BorrowerName, LoanType, Status
Loan Recon       | Sheet2 Additional Columns   | BorrowerName, LoanType
Loan Recon       | Sheet1 Filter               | Type="Loan" AND Status<>"Cancelled"
Loan Recon       | Sheet2 Filter               | Type="Loan"
Loan Recon       | Tolerance                   | 0.05
Loan Recon       | Detail Sheet                | Sheet1
Loan Recon       | Sort Order                  | BorrowerName ASC, LoanID ASC
Interest Recon   | Sheet1 ID Column            | AccountID
Interest Recon   | Sheet2 ID Column            | AcctID
Interest Recon   | Sheet1 Value Columns        | InterestEarned
Interest Recon   | Sheet2 Value Columns        | InterestPaid
Interest Recon   | Sheet1 Additional Columns   | AccountType, Period
Interest Recon   | Sheet2 Additional Columns   | AccountType, Period
Interest Recon   | Sheet1 Filter               | Type="Interest" AND InterestEarned<>0
Interest Recon   | Sheet2 Filter               | Type="Interest"
Interest Recon   | Tolerance                   | 0.001
Interest Recon   | Detail Sheet                | Sheet2
Interest Recon   | Sort Order                  | Period DESC, AccountID ASC
```

*Note: All configs in one table - Workflow column distinguishes between scenarios*

---

## Transform Config (Shared for All Workflows)

**Sheet: "Transform Config"**

**TransformSettingsTable:**
```
Setting      | Value
-------------|------------------
Name         | Upload Transform
Source Sheet | Match Results
Target Sheet | Upload Ready
Filter       | Status<>"Cancelled"
Sort Order   | AccountID ASC, Date DESC
```

*Note: Filter and Sort Order are optional - leave blank if not needed*

**Important: Source Sheet and Target Sheet names are automatically prefixed by the workflow's Result Sheet Prefix.**

Example:
- Config specifies: `Source Sheet = "Match Results"`
- "Fee Recon" workflow (prefix = "Fee") will look for: **"Fee Match Results"**
- "Loan Recon" workflow (prefix = "Loan") will look for: **"Loan Match Results"**
- Workflow with no prefix will use: **"Match Results"** (as-is)

---

**TransformColumnsTable:**
```
Order | Target Column          | Type      | Source
------|------------------------|-----------|------------------
1     | Account ID             | EXISTING  | AccountID
2     | Account Name           | EXISTING  | AccountName
3     | Transaction Date       | EXISTING  | Date
4     | Transaction Type       | EXISTING  | Type
5     | Amount                 | EXISTING  | Amount
6     | Currency               | STATIC    | USD
7     | Processing Date        | STATIC    | =TODAY()
8     | Net Amount             | FORMULA   | =[@Amount]*0.98
9     | Upload Batch           | FORMULA   | =TEXT(TODAY(),"YYYYMMDD")
```

---

## Config Reading Pattern

### Both Recon and Transform now use the same pattern:

1. **Read global settings from Setting/Value table** (using GetConfigValue-style lookup with workflow filter)
2. **Read column-specific operations from tabular format** (loop through rows)

### Example Usage Flow:

```
User selects: "Fee Recon" workflow

1. Look up in WorkflowsTable:
   - Result Prefix = "Fee"

2. Load from "Recon Config" sheet (hardcoded):
   - Filter ReconConfigTable WHERE Workflow = "Fee Recon"
   - Read Setting/Value pairs for that workflow
   - Apply Sheet1 Filter: Type="Fee"
   - Apply Sheet2 Filter: Type="Fee"
   - Sort results by: AccountID ASC

3. If user runs transform after reconciliation:
   - Load from "Transform Config" sheet (hardcoded)
   - Read TransformSettingsTable (Setting/Value pairs)
   - Apply workflow prefix to sheet names:
     * Source Sheet "Match Results" → "Fee Match Results"
     * Target Sheet "Upload Ready" → "Fee Upload Ready"
   - Read TransformColumnsTable (column mappings)
   - Apply Filter: Status<>"Cancelled"
   - Sort output by: AccountID ASC, Date DESC

4. Reconciliation output goes to:
   - "Fee All Results"
   - "Fee Match Results"
   - "Fee Error Results"
   
5. Transform output goes to:
   - "Fee Upload Ready"
```

---

## Benefits of This Structure

1. **Consistency**: Both Recon and Transform use Setting/Value for global settings
2. **Flexibility**: Easy to add new settings without code changes
3. **Clarity**: Clear separation between settings and column operations
4. **Workflow Support**: Multiple scenarios sharing common sheets
5. **Optional Features**: Filter/Sort can be blank if not needed
6. **Minimal Sheet Count**: Only 5 sheets needed (Workflows, Recon Config, Transform Config, Sheet1, Sheet2) regardless of number of workflows

---

## Migration Notes

**All configs have breaking changes**:

**Workflows Table**:
- Old format: Three columns (Workflow Name | Recon Config Sheet | Transform Config Sheet | Result Sheet Prefix)
- New format: Two columns (Workflow Name | Result Sheet Prefix)
- Config sheet names are now hardcoded ("Recon Config" and "Transform Config")

**Recon Config**:
- Old format: Two columns (Setting | Value)
- New format: Three columns (Workflow | Setting | Value)
- GetConfigValue needs workflow name parameter to filter by workflow

**Transform Config**:
- Old format: Horizontal row with Name/Source/Target in row 2
- New format: Vertical Setting/Value table

**Code changes needed**:
- Hardcode config sheet names: "Recon Config" and "Transform Config"
- GetConfigValue(wsConfig, workflowName, settingName)
- Read workflow name first, then filter config table by it
