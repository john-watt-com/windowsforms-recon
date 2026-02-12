# Transform Config Guide

## Overview

The Transform feature allows you to create custom views of your reconciliation results or any other sheet by:
- Selecting specific columns
- Reordering columns
- Adding calculated columns (Excel formulas)
- Adding static values
- Looking up data from other sheets (XLOOKUP, VLOOKUP, etc.)

## Transform Config Sheet Structure

Create a worksheet named **"Transform Config"** with the following structure:

### Header Section (Rows 1-2)

| A | B | C | D |
|---|---|---|---|
| **Transform Name** | **Source Sheet** | **Target Sheet** | |
| _(your transform name)_ | _(source sheet name)_ | _(target sheet name)_ | |

### Column Definitions Section (Rows 4-onwards)

| A | B | C | D |
|---|---|---|---|
| **Order** | **Target Column** | **Type** | **Source** |
| 1 | ID | Existing | ID |
| 2 | Customer | Existing | Region |
| 3 | Status | Static | PENDING |
| 4 | Priority | Formula | IF([@Amount]>1000,"HIGH","LOW") |
| 5 | OwnerName | Formula | XLOOKUP([@ID],Lookup!$A:$A,Lookup!$B:$B,"Unknown") |

**Column Definitions:**
- **Order** (Column A): Numeric order (1, 2, 3...) for the columns in the target sheet
- **Target Column** (Column B): The header name for the new column
- **Type** (Column C): One of: `Existing`, `Static`, or `Formula`
- **Source** (Column D): Depends on Type (see below)

## Column Types

### 1. Existing
Copies data from an existing column in the source sheet.

**Source:** Column name from the source sheet

**Example:**
```
Order | Target Column | Type     | Source
1     | ID            | Existing | ID
2     | Amount        | Existing | Difference
3     | Region        | Existing | Region
```

### 2. Static
Uses the same value for all rows.

**Source:** The static value to use

**Example:**
```
Order | Target Column | Type   | Source
4     | Status        | Static | REVIEW NEEDED
5     | Department    | Static | Finance
6     | Year          | Static | 2026
```

### 3. Formula
Uses an Excel formula to calculate values. The target sheet is created as an Excel Table, so you can use **structured references** like `[@ColumnName]` to reference columns in the current row. You can also use regular cell references (like A2, B2, C2) which work too.

**Important:** Store formulas as **text** in the Transform Config sheet (prefix with single quote `'` or omit the `=`). The code will add the `=` when inserting into the target sheet.

**Source:** Excel formula as text (with or without leading `=`)

**Example (Structured References - Recommended):**
```
Order | Target Column | Type    | Source
7     | Priority      | Formula | IF([@Amount]>1000,"HIGH","LOW")
8     | IsValid       | Formula | AND([@[Sheet1 Total]]<>0,[@[Sheet2 Total]]<>0)
9     | Total         | Formula | [@[Sheet1 Total]]+[@[Sheet2 Total]]
```

Or prefix with `'` to force text entry:
```
Order | Target Column | Type    | Source
7     | Priority      | Formula | '=IF([@Amount]>1000,"HIGH","LOW")
```

**Example (Cell References - Also Works):**
```
Order | Target Column | Type    | Source
7     | Priority      | Formula | IF(C2>1000,"HIGH","LOW")
8     | IsValid       | Formula | AND(B2<>0,C2<>0)
9     | Total         | Formula | B2+C2
```

## Complete Examples

### Example 1: Simple Export

**Goal:** Export Error Results with custom status and priority

**Transform Config:**

| A | B | C |
|---|---|---|
| **Transform Name** | **Source Sheet** | **Target Sheet** |
| Error Export | Error Results | Export |

| A | B | C | D |
|---|---|---|---|
| **Order** | **Target Column** | **Type** | **Source** |
| 1 | ID | Existing | ID |
| 2 | Variance | Existing | Difference |
| 3 | Amount1 | Existing | Sheet1 Total |
| 4 | Amount2 | Existing | Sheet2 Total |
| 5 | Status | Static | NEEDS REVIEW |
| 6 | Priority | Formula | IF(ABS([@Variance])>1000,"HIGH","NORMAL") |

**Result:** Creates "Export" sheet with 6 columns, priority based on variance size

**Note:** Using `[@Variance]` (structured reference) instead of `B2` makes formulas more readable and maintainable.

---

### Example 2: XLOOKUP from Another Sheet

**Goal:** Enrich Error Results with customer names from a lookup table

**Setup:**
1. Create a sheet named **"Lookup"** with customer data:

| A | B | C |
|---|---|---|
| **ID** | **Customer Name** | **Email** |
| C001 | Acme Corp | acme@example.com |
| C002 | Beta Industries | beta@example.com |
| C003 | Gamma LLC | gamma@example.com |

2. Create Transform Config:

| A | B | C |
|---|---|---|
| **Transform Name** | **Source Sheet** | **Target Sheet** |
| Enriched Errors | Error Results | Error Report |

| A | B | C | D |
|---|---|---|---|
| **Order** | **Target Column** | **Type** | **Source** |
| 1 | Customer ID | Existing | ID |
| 2 | Customer Name | Formula | XLOOKUP([@[Customer ID]],Lookup!$A:$A,Lookup!$B:$B,"Unknown") |
| 3 | Email | Formula | XLOOKUP([@[Customer ID]],Lookup!$A:$A,Lookup!$C:$C,"") |
| 4 | Variance | Existing | Difference |
| 5 | Sheet1 Total | Existing | Sheet1 Total |
| 6 | Sheet2 Total | Existing | Sheet2 Total |
| 7 | Status | Formula | IF(ABS([@Variance])>1000,"URGENT","NORMAL") |

**Result:** Creates "Error Report" with customer names and emails looked up from the Lookup sheet

**Notes:**
- Use `$A:$A` for absolute column references if you reference other sheets
- XLOOKUP's third argument is the return value if not found

---

### Example 3: VLOOKUP Alternative

If you don't have XLOOKUP (Excel 2019 or earlier), use VLOOKUP:

```
Order | Target Column | Type    | Source
2     | Customer Name | Formula | IFERROR(VLOOKUP([@ID],Lookup!$A:$C,2,FALSE),"Unknown")
3     | Email         | Formula | IFERROR(VLOOKUP([@ID],Lookup!$A:$C,3,FALSE),"")
```

**VLOOKUP Syntax:**
- `[@[Customer ID]]` or `A2` - lookup value (ID from current row)
- `Lookup!$A:$C` - lookup table range
- `2` or `3` - column index (2 = column B, 3 = column C)
- `FALSE` - exact match

**Structured Reference with Spaces:**
When column names have spaces, use brackets: `[@[Customer ID]]`, `[@[Sheet1 Total]]`

---

### Example 4: INDEX/MATCH Pattern

For more flexibility, use INDEX+MATCH:

```
Order | Target Column  | Type    | Source
2     | Customer Name  | Formula | IFERROR(INDEX(Lookup!$B:$B,MATCH([@ID],Lookup!$A:$A,0)),"Unknown")
```

---

### Example 5: Conditional Formatting with IFS

Calculate priority based on multiple conditions:

```
Order | Target Column | Type    | Source
6     | Priority      | Formula | IFS(ABS([@Variance])>5000,"CRITICAL",ABS([@Variance])>1000,"HIGH",ABS([@Variance])>100,"MEDIUM",TRUE,"LOW")
```

---

### Example 6: Concatenation and Text Functions

```
Order | Target Column | Type    | Source
7     | Description   | Formula | "Variance of "&TEXT([@Variance],"$#,##0.00")&" for "&[@[Customer Name]]
8     | Upper ID      | Formula | UPPER([@ID])
9     | Trimmed Name  | Formula | TRIM([@[Customer Name]])
```

---

## Important Formula Notes

### Structured References (Recommended)
The target sheet is created as an Excel **Table**, so you can use structured references:

- **Current row column:** `[@ColumnName]` - e.g., `[@Amount]`, `[@ID]`
- **Columns with spaces:** `[@[Column Name]]` - e.g., `[@[Customer ID]]`, `[@[Sheet1 Total]]`
- **Benefits:** More readable, self-documenting, no need to track column letters

**Example:**
```
IF([@Amount]>1000,"HIGH","LOW")
[@[Sheet1 Total]]+[@[Sheet2 Total]]
XLOOKUP([@ID],Lookup!$A:$A,Lookup!$B:$B,"Unknown")
```

### Cell References (Alternative)
You can also use traditional cell references (A2, B2, C2, etc.):
- The table auto-fills formulas, so Excel adjusts references automatically
- Column A = first column (Order 1)
- Column B = second column (Order 2)
- Column C = third column (Order 3)

**Example:**
```
IF(C2>1000,"HIGH","LOW")
D2+E2
XLOOKUP(A2,Lookup!$A:$A,Lookup!$B:$B,"Unknown")
```

### External Sheet References
When referencing **other sheets** (like Lookup tables):
- Use full sheet references: `Lookup!A2` or `Lookup!$A:$A`
- Use `$` for absolute references to prevent column shifting
- Example: `=XLOOKUP(A2,Lookup!$A:$A,Lookup!$B:$B,"Default")`

### Error Handling
Wrap formulas in error handlers to prevent `#N/A` errors:
- `=IFERROR(formula, "default_value")`
- `=IFNA(XLOOKUP(...), "Not Found")`

### Performance Tip
For large datasets, avoid volatile functions like:
- `NOW()`, `TODAY()`, `RAND()`, `INDIRECT()`
- Use them only when necessary

---

## Running the Transform

### From the UserForm
1. Open the Recon Form (`Alt+F8` → `ShowReconForm`)
2. Click **"Run Transform"** button

### From Macros Menu
1. Press `Alt+F8`
2. Select `RunTransform`
3. Click Run

### From VBA
```vba
RunTransform
```

---

## Troubleshooting

### "Transform Config sheet not found"
- Create a worksheet named exactly **"Transform Config"** (case-sensitive)

### "Transform configuration incomplete"
- Check row 2 has values in columns A (Transform Name), B (Source Sheet), C (Target Sheet)

### "Source sheet not found"
- Verify the source sheet name in cell B2 matches exactly (case-sensitive)

### "No column definitions found"
- Add column definitions starting at row 5
- Ensure Order column (A) has numeric values

### "Column not found" errors
- For Type=Existing, verify the Source column name exists in the source sheet
- Column names are case-sensitive

### Formula errors in output
- Check your formula syntax
- Verify you're using row 2 references (A2, B2, etc.)
- Test formulas manually in Excel first
- Use IFERROR to handle lookup errors

### XLOOKUP not available
- XLOOKUP requires Excel 2021/Microsoft 365
- Use VLOOKUP or INDEX/MATCH instead (see Example 3 & 4)

---

## Tips & Best Practices

1. **Test your lookup tables first** - Make sure your Lookup sheet has the data you need

2. **Use descriptive target column names** - "Customer Name" is better than "Name"

3. **Order columns logically** - Put IDs and keys first, calculations last

4. **Create multiple transforms** - You can have different Transform Configs for different purposes (keep only one active, or modify row 2 when switching)

5. **Document your formulas** - Add comments in a separate documentation column

6. **Start simple** - Test with Existing and Static columns first, then add formulas

7. **Check your output** - Review the target sheet to ensure formulas calculated correctly

---

## Advanced: Multiple Transforms

To maintain multiple transform configurations:

**Option 1:** Change row 2 when switching transforms
- Keep multiple configuration blocks in the same sheet
- Update row 2 to point to the configuration you want to run

**Option 2:** Use separate Transform Config sheets
- Create "Transform Config 1", "Transform Config 2", etc.
- Rename the active one to "Transform Config" before running

**Option 3:** Modify the RunTransform code to accept a transform name parameter

---

## Column Type Reference

| Type | Source Value | Example |
|------|--------------|---------|
| **Existing** | Column name from source | `Region` |
| **Static** | Any static text or number | `PENDING`, `2026`, `100` |
| **Formula** | Excel formula as text (without `=`) | `IF([@Amount]>1000,"HIGH","LOW")` or `IF(C2>1000,"HIGH","LOW")` |

---

## Common Lookup Patterns

### XLOOKUP (Excel 2021+)
```excel
' Structured reference
XLOOKUP([@ID], Lookup!$A:$A, Lookup!$B:$B, "Unknown")

' Cell reference
XLOOKUP(A2, Lookup!$A:$A, Lookup!$B:$B, "Unknown")
```

### VLOOKUP (Classic)
```excel
' Structured reference
IFERROR(VLOOKUP([@ID], Lookup!$A:$C, 2, FALSE), "Unknown")

' Cell reference
IFERROR(VLOOKUP(A2, Lookup!$A:$C, 2, FALSE), "Unknown")
```

### INDEX/MATCH (Flexible)
```excel
' Structured reference
IFERROR(INDEX(Lookup!$B:$B, MATCH([@ID], Lookup!$A:$A, 0)), "Unknown")

' Cell reference
IFERROR(INDEX(Lookup!$B:$B, MATCH(A2, Lookup!$A:$A, 0)), "Unknown")
```

### Multiple Criteria Lookup
```excel
' Structured reference
XLOOKUP([@ID]&[@Region], Lookup!$A:$A&Lookup!$B:$B, Lookup!$C:$C, "Not Found")

' Cell reference
XLOOKUP(A2&B2, Lookup!$A:$A&Lookup!$B:$B, Lookup!$C:$C, "Not Found")
```

---

## Example: Complete Error Report with Lookups

**Lookup Sheet:**

| A (ID) | B (Customer) | C (Owner) | D (Region) |
|--------|--------------|-----------|------------|
| C001 | Acme Corp | John Smith | North |
| C002 | Beta Inc | Jane Doe | South |

**Transform Config:**

| A | B | C |
|---|---|---|
| **Transform Name** | **Source Sheet** | **Target Sheet** |
| Full Error Report | Error Results | Final Report |

| A | B | C | D |
|---|---|---|---|
| **Order** | **Target Column** | **Type** | **Source** |
| 1 | ID | Existing | ID |
| 2 | Customer | Formula | XLOOKUP([@ID],Lookup!$A:$A,Lookup!$B:$B,"Unknown Customer") |
| 3 | Assigned To | Formula | XLOOKUP([@ID],Lookup!$A:$A,Lookup!$C:$C,"Unassigned") |
| 4 | Region | Formula | XLOOKUP([@ID],Lookup!$A:$A,Lookup!$D:$D,"") |
| 5 | Variance | Existing | Difference |
| 6 | Amount (Sys1) | Existing | Sheet1 Total |
| 7 | Amount (Sys2) | Existing | Sheet2 Total |
| 8 | Priority | Formula | IFS(ABS([@Variance])>5000,"P1-CRITICAL",ABS([@Variance])>1000,"P2-HIGH",TRUE,"P3-NORMAL") |
| 9 | Status | Static | OPEN |
| 10 | Date Created | Formula | TODAY() |
| 11 | Notes | Formula | IF(ABS([@Variance])>5000,"Escalate immediately","Review within 48 hours") |

**Result:** Comprehensive report with customer info, priority classification, and action items!
