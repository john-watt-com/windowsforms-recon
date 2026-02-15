# Test Data Overview

## Current Test Files (for New Config Structure)

### TestData_Summary.csv (Summary Level)
**Purpose:** Summary-level data with one row per AccountID

**Structure:**
- **AccountID**: Unique identifier for matching
- **AccountName**: Account description
- **Type**: Transaction type (Fee, Loan, Interest) - enables workflow filtering
- **Amount**: Total amount for this account
- **Status**: Transaction status (Active, Cancelled) - additional filter option

**Records:** 11 unique accounts across Fee (7), Loan (2), Interest (2) types

### TestData_Detail.csv (Detail Level)
**Purpose:** Detail-level data with multiple rows per AccountID that aggregate to Summary

**Structure:**
- **AccountID**: Identifier for matching
- **AccountName**: Account description
- **Type**: Transaction type (Fee, Loan, Interest)
- **Amount**: Line item amount
- **LineItem**: Description of detail line
- **Status**: Transaction status

**Records:** 24 detail rows (2-3 rows per account) that sum to Summary amounts

## Test Scenarios

### Fee Recon Workflow
**Filter:** Type="Fee"
**Tolerance:** 0.01
**Reconciliation Mode:** Summary vs Detail aggregation

After filtering for Fee records and aggregating Detail by AccountID:

**Matches (within tolerance):**
- 1001: Summary $150.00 = Detail $150.00 (100+50)
- 1002: Summary $225.50 = Detail $225.50 (150.50+75)
- 1006: Summary $300.00 = Detail $300.00 (200+100)
- 1010: Summary $210.00 = Detail $210.00 (150+60)

**Errors:**
- 1004: Summary $175.00 vs Detail $175.25 (100+75.25) - DIFF $0.25
- 1007: Summary $425.75 vs Detail $425.50 (300+125.50) - DIFF $0.25
- 1011: Summary $275.50 vs Detail $280.00 (200+80) - DIFF $4.50
- 1013: Only in Detail $195.00 (120+75) - MISSING FROM SUMMARY

**Filtered out:** 1003 (Loan), 1005 (Interest), 1008 (Loan), 1012 (Interest)

### Loan Recon Workflow
**Filter:** Type="Loan"
**Records:** 
- 1003: Summary $5000.00 = Detail $5000.00 (3000+2000)
- 1008: Summary $10000.00 = Detail $10000.00 (7000+3000)

### Interest Recon Workflow
**Filter:** Type="Interest"
**Records:** 
- 1005: Summary $50.00 = Detail $50.00 (30+20)
- 1012: Summary $75.00 = Detail $75.00 (45+30)

## Expected Results (Fee Recon)
- **Summary Fee records**: 7
- **Detail Fee records**: 16 rows (aggregated to 8 unique IDs)
- **Matches**: 4
- **Errors**: 4 (3 amount diffs + 1 missing from summary)

## Loading Instructions
1. Import TestData_Summary.csv as source data (Sheet1 or configured Summary Sheet)
2. Import TestData_Detail.csv as comparison data (Sheet2 or configured Detail Sheet)
3. Select workflow: "Fee Recon", "Loan Recon", or "Interest Recon"
4. Run Reconciliation
5. Verify results match expected outcomes above

---

## Legacy Test Files (Old Config Format)

### TestData_Default_Sheet1.csv / TestData_Default_Sheet2.csv
**Status:** Removed - replaced by new multi-workflow test data

**Old Structure:**
- Sheet1: ID, Value, ParentID (6 summary records)
- Sheet2: ID, Value-Plus, Value-Minus, Field1-3 (12 detail records)
- Tested detail expansion with ParentID join

### TestData_Summary_Sheet1.csv / TestData_Summary_Sheet2.csv
**Status:** Still available (if needed for legacy testing)

**Structure:**
- Sheet1: ID, Amount, Region, Department (9 detail records)
- Sheet2: ID, Debits, Credits, AccountCode, Source (10 detail records)
- Multiple rows per ID, tests aggregated summary reconciliation
