# Excel VBA Reconciliation Tool

VBA port of the C# Windows Forms Excel Reconciliation App. Compares two Excel files by summing numeric columns grouped by an ID field.

## Quick Start

1. **Create workbook** with sheets: `Sheet1`, `Sheet2`, `Recon Config`, `All Results`
   - Note: `Match Results` and `Error Results` sheets are created automatically
2. **Set up Recon Config sheet** with a table named **"ReconConfigTable"** containing Setting and Value columns (see [CONFIG_EXAMPLE.md](CONFIG_EXAMPLE.md))
3. **(Optional) Set up Workflows sheet** for managing multiple configurations (see Workflows section below)
4. **Import VBA code**:
   - Module: [ReconModule.vba](ReconModule.vba)
   - UserForm: [ReconForm.vba](ReconForm.vba)
5. **Run** the UserForm to load files and reconcile

See [SETUP_INSTRUCTIONS.md](SETUP_INSTRUCTIONS.md) for detailed setup steps.

## Workflows

Manage multiple reconciliation and transformation configurations using workflows. This allows you to:
- Switch between different comparison scenarios (Orders, Invoices, GL, etc.)
- Reuse transform configs across multiple recon configs
- Keep configs organized and easily selectable via dropdown

**Setup:**
1. Create a **"Workflows"** sheet with table **"WorkflowsTable"**
2. Columns: Workflow Name | Recon Config Sheet | Transform Config Sheet | Result Sheet Prefix
3. Example:
   ```
   Default          | Recon Config           | Transform Config      | (blank)
   OrderComparison  | Recon Config - Orders  | Transform Config - Standard | Orders
   ```
4. Select workflow from dropdown in UI

If no Workflows sheet exists, the tool defaults to using "Recon Config" and "Transform Config" sheets.

## What It Does

1. Load two external Excel files into Sheet1 and Sheet2
2. For each unique ID:
   - Sum specified numeric columns from Sheet1 (supports +/- prefixes for add/subtract)
   - Sum specified numeric columns from Sheet2 (supports +/- prefixes for add/subtract)
   - Compare totals (configurable tolerance, default 0.01)
3. Output results to three sheets:
   - **All Results**: Complete reconciliation data
   - **Match Results**: Only matching records (aggregated or detail-expanded based on config)
   - **Error Results**: Only mismatched records with highlighting

## Features

- **Configurable tolerance** - Set acceptable difference threshold (default 0.01)
- **Add/Subtract operations** - Use `+Column` to add, `-Column` to subtract (e.g., `+Debits,-Credits`)
- **Additional columns** - Include descriptive columns from both sheets in results
- **Detail expansion mode** - Expand matched results to show all detail rows when one sheet is at summary level
- **Transform feature** - Create custom views with column selection, static values, and formulas
- **Automatic table formatting** - Results formatted as Excel tables with filtering/sorting

## Files

- **ReconModule.vba** - Core reconciliation logic (import as Module)
- **ReconForm.vba** - UI code (import as UserForm code-behind)
- **SETUP_INSTRUCTIONS.md** - Complete setup guide
- **CONFIG_EXAMPLE.md** - Configuration examples
- **TRANSFORM_CONFIG.md** - Guide for creating custom sheet transformations

## Input Requirements

Your Excel files must have:
- Header row in row 1
- Data starting in row 2
- Column names that match your Recon Config sheet entries
- Numeric values in your selected value columns

## Key Differences from C# Version

| C# App | VBA Version |
|--------|-------------|
| Interactive column selection UI | Table-based config in Recon Config sheet |
| Multiple worksheet tabs | Single workbook with named sheets |
| Faster (external app) | Slower (runs in Excel) |

## Transform Feature

Create custom views from any sheet using the Transform Config sheet:
- **Select & reorder columns** - Choose which columns to include and their order
- **Add static values** - Insert constant values (statuses, labels, dates)
- **Excel formulas** - Calculate new columns (IF statements, math, text functions)
- **Lookup data** - Use XLOOKUP, VLOOKUP, or INDEX/MATCH to enrich data from other sheets

See [TRANSFORM_CONFIG.md](TRANSFORM_CONFIG.md) for complete guide with examples.

## Distribution

Save the Excel workbook as `.xlsm` and share directly. Recipients must enable macros.

## Support

This tool replicates the logic from the C# app located in `../ExcelReconApp/` in this repository.
