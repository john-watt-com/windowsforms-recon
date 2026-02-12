# Recon Config Sheet Example

## Important: Configuration is Now Table-Based

The Recon Config sheet must contain an Excel table named **"ReconConfigTable"** with two columns:
- **Setting** - The name of the configuration setting
- **Value** - The value for that setting

Settings can be in any order. The tool looks up settings by name (case-insensitive).

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

### Aggregated Mode (Default)

Create a table named **"ReconConfigTable"** with these columns:

| Setting | Value |
|---------|-------|
| Sheet1 ID Column | CustomerID |
| Sheet2 ID Column | CustomerID |
| Sheet1 Value Columns | Sales,Tax |
| Sheet2 Value Columns | TotalRevenue |
| Sheet1 Additional Columns | Region |
| Tolerance | 0.01 |

**Notes:**
- Settings can be in any order
- Setting names are case-insensitive
- Optional settings (Tolerance, Detail Sheet, Sheet2 Additional Columns) can be omitted
- If Tolerance is omitted or invalid, defaults to 0.01

## What This Does

- Matches records by **CustomerID** from both files
- For Sheet1: Sums **Sales + Tax** for each CustomerID
  - C001: 1000 + 80 = 1080
  - C002: 1500 + 120 = 1620
- For Sheet2: Uses **TotalRevenue** for each CustomerID
  - C001: 1080
  - C002: 1620
- Compares the totals with a tolerance of **0.01** (records are considered matching if the difference is less than 0.01)

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
---

## Advanced: Detail Expansion Mode

### When to Use Detail Expansion

Sometimes one sheet has **detail-level data** (multiple rows per ID) while the other has **summary data** (one row per ID). After reconciling the aggregated totals, you may want to see the individual detail rows for matched IDs.

**Detail Expansion Mode** expands the Match Results to show:
- One row for each detail row from the detail sheet
- Totals from both sheets (repeated on every detail row)
- Additional columns from both the detail sheet and the summary sheet

### Configuration for Detail Expansion

Add these settings to your ReconConfigTable:

| Setting | Value |
|---------|-------|
| Sheet1 ID Column | OrderID |
| Sheet2 ID Column | OrderID |
| Sheet1 Value Columns | LineAmount |
| Sheet2 Value Columns | TotalAmount |
| Sheet1 Additional Columns | LineNumber,Product,Quantity |
| Tolerance | 0.01 |
| Detail Sheet | SHEET1 |
| Sheet2 Additional Columns | CustomerName,OrderDate |

**Key Settings:**
- **Detail Sheet**: Specify which sheet has the detail-level data
  - Use `SHEET1` if Sheet1 has multiple rows per ID
  - Use `SHEET2` if Sheet2 has multiple rows per ID
  - Omit this setting for standard aggregated mode
- **Sheet2 Additional Columns**: Comma-separated list of columns from Sheet2 to include in results
  - Only needed when using detail expansion
  - These columns are joined from the summary sheet to each detail row

### Example: Order Line Items vs Order Totals

**Sheet1 (Detail - Order Line Items):**
| OrderID | LineNumber | Product | Quantity | LineAmount |
|---------|------------|---------|----------|------------|
| ORD001 | 1 | Widget A | 2 | 100 |
| ORD001 | 2 | Widget B | 1 | 50 |
| ORD002 | 1 | Widget C | 3 | 150 |

**Sheet2 (Summary - Order Totals):**
| OrderID | CustomerName | OrderDate | TotalAmount |
|---------|--------------|-----------|-------------|
| ORD001 | Acme Corp | 2024-01-15 | 150 |
| ORD002 | Beta Inc | 2024-01-16 | 150 |

### Results with Detail Expansion

With `Detail Sheet = SHEET1` and `Sheet2 Additional Columns = CustomerName,OrderDate`, the **Match Results** sheet will show:

| ID | LineNumber | Product | Quantity | Sheet1 Total | Sheet2 Total | CustomerName | OrderDate |
|----|------------|---------|----------|--------------|--------------|--------------|-----------|
| ORD001 | 1 | Widget A | 2 | 150 | 150 | Acme Corp | 2024-01-15 |
| ORD001 | 2 | Widget B | 1 | 150 | 150 | Acme Corp | 2024-01-15 |
| ORD002 | 1 | Widget C | 3 | 150 | 150 | Beta Inc | 2024-01-16 |

**Notice:**
- Each line item from Sheet1 becomes a separate row
- The totals (150, 150, 150) are repeated on every detail row for the same OrderID
- Sheet2's columns (CustomerName, OrderDate) are joined to each detail row

### Mode Comparison

| Mode | Detail Sheet Setting | Match Results Output | Use Case |
|------|---------------------|----------------------|----------|
| **Aggregated** | _(blank)_ | 1 row per ID with totals | Standard reconciliation, summary comparison |
| **Detail Expansion** | `SHEET1` or `SHEET2` | 1 row per detail row with joined data | Detail-level analysis after reconciliation |

**Error Results** are always aggregated (1 row per ID) regardless of mode.