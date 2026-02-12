# Excel VBA Reconciliation Tool - Setup Instructions

## Overview
This is a VBA port of the C# Windows Forms Excel Reconciliation App. It performs the same reconciliation logic:
- Load two Excel files
- Select ID column and value columns for each
- Sum value columns by ID
- Compare totals between sheets with 0.01 tolerance
- Display results with match status

## Setup Steps

### 1. Create a New Excel Workbook

1. Open Excel and create a new workbook
2. Save it as a macro-enabled workbook (`.xlsm` format)
3. Name it something like `ExcelReconTool.xlsm`

### 2. Create Required Worksheets

Rename/create the following worksheets (right-click sheet tabs):

- **Sheet1** - Will hold data from first Excel file
- **Sheet2** - Will hold data from second Excel file  
- **Config** - Configuration settings
- **All Results** - Complete reconciliation output

**Note:** The following sheets will be created automatically when you run reconciliation:
- **Match Results** - Records where totals match (IsMatch = TRUE)
- **Error Results** - Records where totals don't match (IsMatch = FALSE)

### 3. Set Up the Config Sheet

In the **Config** worksheet, set up the following:

| Row | Column A | Column B |
|-----|----------|----------|
| 1   | **Setting** | **Value** |
| 2   | Sheet1 ID Column | _(enter column name, e.g., "CustomerID")_ |
| 3   | Sheet2 ID Column | _(enter column name, e.g., "CustomerID")_ |
| 4   | Sheet1 Value Columns | _(comma-separated, e.g., "Amount1,Amount2")_ |
| 5   | Sheet2 Value Columns | _(comma-separated, e.g., "TotalAmount")_ |
| 6   | Sheet1 Additional Columns | _(optional: comma-separated columns to display, e.g., "Region,Department")_ |

**Note:** Row 6 (Sheet1 Additional Columns) is optional. If specified, the first occurrence of each ID from Sheet1 will have these column values displayed in the results.

### 4. Enable Developer Tab

If you don't see the Developer tab in Excel:
1. File → Options → Customize Ribbon
2. Check "Developer" on the right side
3. Click OK

### 5. Import the VBA Code

#### Import the Module:

1. Press `Alt+F11` to open the VBA Editor
2. Insert → Module
3. Open `ReconModule.vba` in a text editor
4. Copy all the code
5. Paste it into the new module
6. (Optional) Rename the module to "ReconModule" in the Properties window (F4)

#### Create the UserForm:

1. In VBA Editor: Insert → UserForm
2. Rename the UserForm to "ReconForm" in Properties (F4)
3. Add the following controls from the Toolbox:

   **CommandButton1:**
   - Name: `btnLoadSheet1`
   - Caption: `Load Worksheet 1`
   - Position: Top left

   **CommandButton2:**
   - Name: `btnLoadSheet2`
   - Caption: `Load Worksheet 2`
   - Position: Below button 1

   **CommandButton3:**
   - Name: `btnRunRecon`
   - Caption: `Run Reconciliation`
   - Position: Below button 2

   **Label1:**
   - Name: `lblStatus`
   - Caption: `Ready`
   - Width: Fill width of form
   - Position: Bottom of form

4. Double-click the UserForm to open its code window
5. Open `ReconForm.vba` in a text editor
6. Copy all the code
7. Paste it into the UserForm code window (replacing any existing code)

### 6. Add a Quick Access Button (Optional)

**Option 1: Run from Macros Menu (Simplest)**
- Press `Alt+F8` (or Developer → Macros)
- Select `ShowReconForm`
- Click Run

**Option 2: Add a Button to the Worksheet**
1. Go to any worksheet (e.g., Config sheet)
2. Developer tab → Insert → Button (Form Control)
3. Draw a button on the sheet
4. In the "Assign Macro" dialog, select `ShowReconForm`
5. Click OK
6. Right-click button → Edit Text → Change to "Open Recon Tool"

**Option 3: Add to Quick Access Toolbar**
- Right-click the button you created → Add to Quick Access Toolbar

### 7. Enable Trust Access to VBA Project (If Needed)

If you get errors about file access:
1. File → Options → Trust Center → Trust Center Settings
2. Macro Settings → Check "Trust access to the VBA project object model"
3. Click OK

### 8. Save the Workbook

- Save the workbook (`.xlsm` format)
- Keep it in a trusted location

## Usage Instructions

### Launching the Tool:

- **Method 1**: Press `Alt+F8`, select `ShowReconForm`, click Run
- **Method 2**: Click the button if you created one in step 6
- **Method 3**: Developer → Macros → ShowReconForm → Run

### Basic Workflow:

1. **Configure** (Config sheet):
   - Enter ID column names for both sheets (B2, B3)
   - Enter value column names (B4, B5) as comma-separated lists
   - Optionally enter additional columns to display from Sheet1 (B6)

2. **Load Data**:
   - Click "Load Worksheet 1" and select first Excel file
   - Click "Load Worksheet 2" and select second Excel file
   - Check that data appears in Sheet1 and Sheet2 tabs

3. **Run Reconciliation**:
   - Click "Run Reconciliation"
   - View results in three tabs:
     - **All Results**: Complete data with all records
     - **Match Results**: Only matching records (no IsMatch/Difference columns)
     - **Error Results**: Only mismatched records (includes Difference column)

### Results Explanation:

Results are displayed across three Excel worksheets:

#### All Results Sheet
Excel Table named "ReconResultsTable" with all reconciliation data:

- **IsMatch**: TRUE if difference is < 0.01, FALSE otherwise (appears first for easy filtering)
- **Difference**: Sheet1 Total - Sheet2 Total
- **ID**: The unique identifier from both sheets
- **Sheet1 Total**: Sum of selected value columns from Sheet1
- **Sheet2 Total**: Sum of selected value columns from Sheet2  
- **Additional Columns**: Any columns specified in "Sheet1 Additional Columns" (optional)
  - Shows values from the first occurrence of each ID in Sheet1
  - Useful for displaying descriptive fields like Region, Department, Customer Name, etc.

#### Match Results Sheet
Excel Table named "MatchResultsTable" with only matching records (IsMatch = TRUE):

- Contains: **ID**, **Sheet1 Total**, **Sheet2 Total**, and any **Additional Columns**
- **Excludes**: IsMatch and Difference columns (since all records match)
- Clean view of reconciled data

#### Error Results Sheet
Excel Table named "ErrorResultsTable" with only mismatched records (IsMatch = FALSE):

- **Difference**: Sheet1 Total - Sheet2 Total (shown first for quick review)
- **ID**: The unique identifier
- **Sheet1 Total**: Sum from Sheet1
- **Sheet2 Total**: Sum from Sheet2
- **Additional Columns**: Any specified additional columns
- **Excludes**: IsMatch column (all are FALSE)
- All rows highlighted in light red for visibility

**Features:**
- All tables support filtering and sorting
- Click dropdown arrows in headers to filter
- Error Results are pre-highlighted in red
- Each sheet can be used independently for reporting

## Input Requirements

Your Excel files must have:
- **Header row in row 1** with column names
- **Data starting in row 2**
- **ID column**: Can be text or numbers (compared as text)
- **Value columns**: Must contain numeric values
  - Non-numeric values are treated as 0
  - Empty cells are treated as 0

## Differences from C# Version

| Feature | C# App | VBA Version |
|---------|--------|-------------|
| File Loading | Windows Forms OpenFileDialog | Excel FileDialog |
| UI | Separate tabs with DataGridViews | Excel worksheets |
| Configuration | ComboBoxes and CheckedListBoxes | Config sheet cells |
| Column Selection | Interactive UI controls | Manual entry in Config sheet |
| Performance | Faster for large files | Slower for 10,000+ rows |

## Troubleshooting

### "Column not found" error
- Check that column names in Config sheet exactly match headers in your data files
- Column names are case-sensitive and whitespace-sensitive

### "Required worksheets not found"
- Ensure you have Sheet1, Sheet2, and Config worksheets
- The All Results, Match Results, and Error Results sheets will be created automatically
- Names must match exactly (case-sensitive)

### Import button doesn't work
- Check that you've enabled macros
- Ensure Microsoft Office is set to "Enable all macros" (not recommended for untrusted files)
- Try File → Options → Trust Center → Macro Settings

### Reconciliation results are wrong
- Verify ID column selections
- Verify value columns are numeric
- Check for duplicate IDs in source data (they will be summed)
- Ensure value column names are comma-separated without extra spaces

## Performance Notes

- Files up to 10,000 rows: Good performance
- Files 10,000-50,000 rows: May be slow
- Files > 50,000 rows: Consider using the C# app instead

## Distribution

To share this tool:
1. Save the `.xlsm` file
2. Share the file directly
3. Recipients must enable macros when opening

Alternatively, you can export the VBA code and share the `.vba` files with setup instructions.
