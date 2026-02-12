# Excel VBA Reconciliation Tool

VBA port of the C# Windows Forms Excel Reconciliation App. Compares two Excel files by summing numeric columns grouped by an ID field.

## Quick Start

1. **Create workbook** with sheets: `Sheet1`, `Sheet2`, `Config`, `Results`
2. **Set up Config sheet** with ID columns and value columns (see [CONFIG_EXAMPLE.md](CONFIG_EXAMPLE.md))
3. **Import VBA code**:
   - Module: [ReconModule.vba](ReconModule.vba)
   - UserForm: [ReconForm.vba](ReconForm.vba)
4. **Run** the UserForm to load files and reconcile

See [SETUP_INSTRUCTIONS.md](SETUP_INSTRUCTIONS.md) for detailed setup steps.

## What It Does

1. Load two external Excel files into Sheet1 and Sheet2
2. For each unique ID:
   - Sum specified numeric columns from Sheet1
   - Sum specified numeric columns from Sheet2
   - Compare totals (tolerance: 0.01)
3. Output results to Results sheet with match/mismatch highlighting

## Files

- **ReconModule.vba** - Core reconciliation logic (import as Module)
- **ReconForm.vba** - UI code (import as UserForm code-behind)
- **SETUP_INSTRUCTIONS.md** - Complete setup guide
- **CONFIG_EXAMPLE.md** - Configuration examples

## Input Requirements

Your Excel files must have:
- Header row in row 1
- Data starting in row 2
- Column names that match your Config sheet entries
- Numeric values in your selected value columns

## Key Differences from C# Version

| C# App | VBA Version |
|--------|-------------|
| Interactive column selection UI | Manual config in Config sheet |
| Multiple worksheet tabs | Single workbook with named sheets |
| Faster (external app) | Slower (runs in Excel) |

## Distribution

Save the Excel workbook as `.xlsm` and share directly. Recipients must enable macros.

## Support

This tool replicates the logic from the C# app located in `../ExcelReconApp/` in this repository.
