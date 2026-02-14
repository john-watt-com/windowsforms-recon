# New Workflow & Recon Config Structure

## Workflows Config (WorkflowsTable)
- Lists all available workflows for the UI combo box.
- Columns: WorkflowName, Description, Enabled, Notes

| WorkflowName            | Description                   | Enabled | Notes                |
|-------------------------|------------------------------|---------|----------------------|
| GL Monthly              | GL monthly reconciliation     | TRUE    |                      |
| GL Daily                | GL daily reconciliation       | TRUE    |                      |
| Trade Journal Monthly   | Trade journal monthly recon   | TRUE    |                      |
| Trade Journal Daily     | Trade journal daily recon     | TRUE    |                      |

## Recon Config (ReconConfigTable)
- Stores all settings for each workflow.
- Columns: WorkflowName, Setting, Value
- Always include: Sheet1 ID Column, Sheet2 ID Column, Sheet1 Value Columns, Sheet2 Value Columns
- Add Sheet1 Filter and/or Sheet2 Filter as needed for filtering.

| WorkflowName            | Setting                | Value                                   |
|-------------------------|------------------------|-----------------------------------------|
| GL Monthly              | Sheet1 ID Column       | ID                                      |
| GL Monthly              | Sheet2 ID Column       | ID                                      |
| GL Monthly              | Sheet1 Value Columns   | Amount                                  |
| GL Monthly              | Sheet2 Value Columns   | Amount                                  |
| GL Monthly              | Sheet1 Filter          |                                         |
| GL Monthly              | Sheet2 Filter          |                                         |
| GL Monthly              | Tolerance              | 0.03                                    |
| GL Monthly              | Sort Order             | ParentID,ID                             |
| ...                     | ...                    | ...                                     |

## Notes
- Add new workflows by adding rows to both tables.
- Filtering is optional; leave filter values blank if not needed.
- Row order does not matter; settings are looked up by name.
- This structure is compatible with the existing codebase and UI.
