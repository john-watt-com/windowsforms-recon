# Test Data Files

## Default Workflow Test Files

**TestData_Default_Sheet1.csv** - Summary level (6 records, unique IDs)
- Columns: ID, Value, ParentID
- Each ID appears once

**TestData_Default_Sheet2.csv** - Detail level (12 records)
- Columns: ID, Value-Plus, Value-Minus, Field1, Field2, Field3
- Multiple rows per ID for records 2-7

### Expected Results (Default Workflow):

**Matches (within 0.03 tolerance):**
- ID 1: 100.00 = 100.00 (exact match)
- ID 2: 200.00 = 200.00 (2 detail rows: 150+50)
- ID 3: 150.50 ≈ 150.50 (2 detail rows: 100.50+50.02-0.02, within tolerance)
- ID 4: 75.25 = 75.25 (exact match)
- ID 5: 300.00 = 300.00 (2 detail rows: 200+100)

**Errors:**
- ID 6: 125.00 ≠ 125.00 (2 detail rows: 100+30-5, difference beyond tolerance)
- ID 7: Not in Sheet1 (only in Sheet2, 2 rows: 50+50=100)

**Match Results with Detail Expansion (Detail Sheet = Sheet2):**
- Each detail row from Sheet2 shown with Sheet1's ParentID joined
- Example for ID 2: Two rows, both showing ParentID=A2, Sheet1 Total=200, Sheet2 Total=200

---

## Summary Workflow Test Files

**TestData_Summary_Sheet1.csv** - Detail level (9 records)
- Columns: ID, Amount, Region, Department
- Multiple rows per ID for A001, A003, A004

**TestData_Summary_Sheet2.csv** - Detail level (10 records)
- Columns: ID, Debits, Credits, AccountCode, Source
- Multiple rows per ID for A001, A003, A005, A006

### Expected Results (Summary Workflow):

**Aggregated comparison (No Detail Sheet):**

**Matches (within 0.01 tolerance):**
- A001: 250.00 = 250.00 (Sheet1: 100+150, Sheet2: 300-50)
- A002: 200.00 = 200.00 (Sheet1: 200, Sheet2: 200)
- A003: 150.50 = 150.50 (Sheet1: 50+75.50+25, Sheet2: 100+50.50)

**Errors:**
- A004: 500.00 = 500.00 (Sheet1: 300+200, Sheet2: 520-20 = exact match but large amounts)
- A005: 125.00 ≠ 125.00 (Sheet1: 125, Sheet2: 100+25, exact match)
- A006: No data in Sheet1 (only in Sheet2: 75+45=120)

**Match Results:**
- Single aggregated row per ID
- Additional columns from both sheets included
- Region/Department from Sheet1, AccountCode/Source from Sheet2

---

## Testing Instructions

### For Default Workflow:
1. Select "Default" workflow from dropdown
2. Load TestData_Default_Sheet1.csv into Sheet1
3. Load TestData_Default_Sheet2.csv into Sheet2
4. Run Reconciliation
5. Check Match Results - should see detail expansion with ParentID column
6. Run Transform - should create Export sheet with ParentID populated

### For Summary Workflow:
1. Select "Summary Comparison" workflow from dropdown
2. Load TestData_Summary_Sheet1.csv into Sheet1
3. Load TestData_Summary_Sheet2.csv into Sheet2
4. Run Reconciliation
5. Check Match Results - should see aggregated summary (1 row per ID)
6. Run Transform - should create Summary Export with Region, Department, AccountCode, Source columns

---

## Key Differences to Validate

| Aspect | Default Workflow | Summary Workflow |
|--------|-----------------|------------------|
| Sheet1 | Summary (unique IDs) | Detail (multiple per ID) |
| Sheet2 | Detail (multiple per ID) | Detail (multiple per ID) |
| Match Results | Detail expanded | Aggregated summary |
| Row Count | Many rows per matched ID | One row per matched ID |
| Additional Columns | Only Sheet1 (ParentID) | Both sheets (Region, Dept, Acct, Source) |
| Detail Sheet Setting | SHEET2 | (blank) |
