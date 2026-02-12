# Recon Config Sheet Example

## Basic Example

If your Excel files look like this:

**File 1 (Customers_2024.xlsx):**
| CustomerID | Region | Sales | Tax |
|------------|--------|-------|-----|
| C001 | North | 1000 | 80 |
| C002 | South | 1500 | 120 |

**File 2 (Customers_2025.xlsx):**
| CustomerID | Department | TotalRevenue |
|------------|------------|--------------|
| C001 | A | 1080 |
| C002 | B | 1620 |

## Your Recon Config Sheet Should Look Like:

| A | B |
|---|---|
| **Setting** | **Value** |
| Sheet1 ID Column | CustomerID |
| Sheet2 ID Column | CustomerID |
| Sheet1 Value Columns | Sales,Tax |
| Sheet2 Value Columns | TotalRevenue |
| Sheet1 Additional Columns | Region |

## What This Does

- Matches records by **CustomerID** from both files
- For Sheet1: Sums **Sales + Tax** for each CustomerID
  - C001: 1000 + 80 = 1080
  - C002: 1500 + 120 = 1620
- For Sheet2: Uses **TotalRevenue** for each CustomerID
  - C001: 1080
  - C002: 1620
- Compares the totals

## Expected Results

Results are split across three worksheets:

### All Results Sheet

| IsMatch | Difference | ID | Sheet1 Total | Sheet2 Total | Region |
|---------|------------|----|--------------|--------------|--------|
| TRUE | 0 | C001 | 1080 | 1080 | North |
| TRUE | 0 | C002 | 1620 | 1620 | South |

- Complete data with all columns
- Useful for full audit trail

### Match Results Sheet

| ID | Sheet1 Total | Sheet2 Total | Region |
|----|--------------|--------------|--------|
| C001 | 1080 | 1080 | North |
| C002 | 1620 | 1620 | South |

- Only matching records (IsMatch = TRUE)
- **Excludes**: IsMatch and Difference columns
- Clean view for confirmed reconciliations

### Error Results Sheet

| Difference | ID | Sheet1 Total | Sheet2 Total | Region |
|------------|----|--------------|--------------|--------|
| _(no errors in this example)_ |

- Only mismatched records (IsMatch = FALSE)
- **Includes**: Difference column (for quick issue identification)
- **Excludes**: IsMatch column
- All rows highlighted in red
- Empty if all records match

**Column Explanation:**
- **IsMatch**: TRUE if difference is < 0.01, FALSE otherwise
- **Difference**: Sheet1 Total - Sheet2 Total
- **ID**: The unique identifier
- **Sheet1 Total**: Sum of value columns from Sheet1
- **Sheet2 Total**: Sum of value columns from Sheet2
- **Region** (and any additional columns): Values from first occurrence in Sheet1

---

## Sheet1 Additional Columns (Optional)

You can specify additional columns from Sheet1 to display in the results. This is useful for showing descriptive information like customer names, regions, or departments.

**Example:**
- Sheet1 Additional Columns: `Region,Department,CustomerName`

**Important:** 
- If an ID appears multiple times in Sheet1, only the values from the **first occurrence** are displayed
- These columns are for display only and are **not** included in the reconciliation calculation
- Leave this field blank (B6) if you don't need additional columns

---

## Multiple Value Columns

You can specify multiple value columns separated by commas:

**Example:**
- Sheet1 Value Columns: `Q1Sales,Q2Sales,Q3Sales,Q4Sales`
- Sheet2 Value Columns: `AnnualTotal`

This will sum all four quarters from Sheet1 and compare against the annual total from Sheet2.

---

## Important Notes

1. **Column names must match exactly** (case-sensitive, including spaces)
2. **No spaces after commas** in the value columns list
   - ✅ Good: `Sales,Tax,Fees`
   - ❌ Bad: `Sales, Tax, Fees`
3. **ID columns are compared as text**, so "001" ≠ "1"
4. **Non-numeric values are ignored** in value columns (treated as 0)
5. **Multiple rows with same ID** will be summed together
