# Summary-Detail Reconciliation Test Scenario

## Config Table Example (ReconConfigTable)

| WorkflowName           | Setting                    | Value                    |
|------------------------|----------------------------|--------------------------|
| Trade Journal Monthly  | Sheet1 ID Column           | TradeID                  |
| Trade Journal Monthly  | Sheet2 ID Column           | TradeID                  |
| Trade Journal Monthly  | Sheet1 Value Columns       | Amount                   |
| Trade Journal Monthly  | Sheet2 Value Columns       | Amount                   |
| Trade Journal Monthly  | Sheet1 Additional Columns  | Description              |
| Trade Journal Monthly  | Sheet2 Additional Columns  | Description,Type         |
| Trade Journal Monthly  | Tolerance                  | 0.01                     |
| Trade Journal Monthly  | Detail Sheet               | Sheet2                   |
| Trade Journal Monthly  | Sheet1 Filter              | Amount<>0                |
| Trade Journal Monthly  | Sheet2 Filter              | Description="Fee"        |
| Trade Journal Monthly  | Sort Order                 | TradeID                  |

## Test Data Files

### TestData_Summary_Sheet1.csv
```
TradeID,Amount,Description
1001,500,Principal
1002,0,Interest
1003,200,Fee
1004,0,Fee
```

### TestData_Detail_Sheet2.csv
```
TradeID,Amount,Description,Type
1001,300,Principal,Base
1001,200,Principal,Adjustment
1002,0,Interest,Base
1003,100,Fee,Base
1003,100,Fee,Adjustment
1004,0,Fee,Base
```

## Filtering Example
- Sheet1 Filter: `Amount<>0` (only summary rows with nonzero Amount)
- Sheet2 Filter: `Description="Fee"` (only detail rows with Description = Fee)

## Expected Output (Sample)
| IsMatch | Difference | TradeID | Sheet1 Total | Sheet2 Total | Description (from summary) | Description,Type (from detail) |
|---------|------------|---------|--------------|--------------|----------------------------|-------------------------------|
| TRUE    | 0          | 1003    | 200          | 200          | Fee                        | Fee,Base                      |
| TRUE    | 0          | 1003    | 200          | 200          | Fee                        | Fee,Adjustment                |

## Notes
- Only rows with Description="Fee" and Amount<>0 are included.
- Summary columns are repeated for each detail row.
- You can swap which file is loaded into Sheet1 or Sheet2 and adjust "Detail Sheet" in config to test both summary-detail directions.
