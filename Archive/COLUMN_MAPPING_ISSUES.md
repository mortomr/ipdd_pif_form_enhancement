# Column Mapping Issues and Corrections

## Summary

This document identifies critical column mapping issues in the VBA code that will cause data to be inserted into the wrong database fields.

---

## Issue #1: Duplicate Column Mappings in mod_Database.bas

**Location:** `mod_Database.bas` lines 549-568 (BulkInsertToStaging function)

### Problem

Three pairs of database fields are mapped to the **same Excel columns**, causing data loss and incorrect data assignment:

| Database Field | Assigned Column | Issue |
|---|---|---|
| `funding_project` | Column 13 (M) | ✓ Correct |
| `project_id` | Column 13 (M) | **DUPLICATE - Same as funding_project** |
| | | |
| `moving_isd_year` | Column 17 (Q) | Should be separate |
| `lcm_issue` | Column 17 (Q) | **DUPLICATE - Same as moving_isd_year** |
| | | |
| `justification` | Column 20 (T) | ✓ Correct |
| `prior_year_spend` | Column 20 (T) | **DUPLICATE - Same as justification** |

### Current Code (INCORRECT)

```vba
params(1) = dataRange.Cells(i, 13).Value ' project_id
' ... other params ...
params(10) = dataRange.Cells(i, 13).Value ' funding_project  <-- DUPLICATE!
' ... other params ...
params(14) = dataRange.Cells(i, 17).Value ' moving_isd_year
params(15) = dataRange.Cells(i, 17).Value ' lcm_issue  <-- DUPLICATE!
params(16) = dataRange.Cells(i, 20).Value ' justification
params(17) = dataRange.Cells(i, 20).Value ' prior_year_spend  <-- DUPLICATE!
```

### Impact

- `project_id` and `funding_project` will always have the same value (might be intentional if they're the same in your business logic)
- `lcm_issue` will overwrite the value meant for `moving_isd_year` (definitely wrong)
- `prior_year_spend` will contain text from `justification` instead of numeric values (type mismatch error)

### Solution Required

You need to determine the **actual Excel column layout** for your PIF sheet. Based on the enum in mod_SharedConstants.bas, the layout should be:

| Column | Letter | Field Name | Database Column |
|--------|--------|------------|-----------------|
| 3 | C | Archive Flag | archive_flag |
| 4 | D | Include Flag | include_flag |
| 5 | E | Accounting Treatment | accounting_treatment |
| 6 | F | Change Type | change_type |
| 7 | G | PIF ID | pif_id |
| 8 | H | SEG | seg |
| 9 | I | OPCO | opco |
| 10 | J | Site | site |
| 11 | K | Strategic Rank | strategic_rank |
| 12 | L | From Blanket | *(not used in insert)* |
| 13 | M | Funding Project | funding_project |
| 14 | N | Project Name | project_name |
| 15 | O | Original ISD | original_fp_isd |
| 16 | P | Revised ISD | revised_fp_isd |
| 17 | Q | LCM Issue | lcm_issue |
| 18 | R | Status | status |
| 19 | S | Category | category |
| 20 | T | Justification | justification |
| **??** | **??** | **Moving ISD Year** | **moving_isd_year** (MISSING!) |
| **??** | **??** | **Prior Year Spend** | **prior_year_spend** (MISSING!) |
| **??** | **??** | **Project ID** | **project_id** (if different from funding_project) |

**Questions to Answer:**

1. Are `project_id` and `funding_project` supposed to be the same value? If yes, keep column 13. If no, what column is `project_id`?
2. What column contains `moving_isd_year`? (Not defined in the enum)
3. What column contains `prior_year_spend`? (Not defined in the enum)

---

## Issue #2: Inconsistent Range Definition in mod_Submit.bas

**Location:** `mod_Submit.bas` lines 407-418 (UploadProjectData function)

### Problem

The dataRange starts at **C4** (column C, row 4), but the BulkInsertToStaging function references columns using absolute positions (e.g., column 7 = G, column 13 = M).

### Current Code

```vba
Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
Set dataRange = wsData.Range("C4").CurrentRegion
```

The `CurrentRegion` property expands to include all connected non-empty cells. **If the region expands left to column A**, the column references will work correctly. **If it doesn't expand left**, all column references will be off by 2 positions (C becomes position 1 instead of 3).

### Verification Needed

Check your Excel PIF sheet:
1. Do columns A and B have headers or data in row 3 or row 4?
2. If yes, CurrentRegion will expand left and include them (correct behavior)
3. If no, CurrentRegion starts at column C, making all references wrong

### Recommended Fix

Instead of relying on CurrentRegion behavior, explicitly define the range:

```vba
' Option 1: Specify exact range if you know the last column
Set dataRange = wsData.Range("A4:T" & wsData.Cells(wsData.Rows.Count, 7).End(xlUp).Row)

' Option 2: Use constants from mod_SharedConstants
Dim lastRow As Long
lastRow = wsData.Cells(wsData.Rows.Count, PIFDataColumns.colPIFID).End(xlUp).Row
Set dataRange = wsData.Range("A4:T" & lastRow)
```

---

## Issue #3: Missing Constants in mod_SharedConstants.bas

**Location:** `mod_SharedConstants.bas` lines 55-74 (PIFDataColumns enum)

### Problem

The enum is missing three fields that exist in the database schema:

1. `moving_isd_year` - Database column exists, but no enum constant
2. `prior_year_spend` - Database column exists, but no enum constant
3. `project_id` - Used in database but not in enum (if it's different from funding_project)

### Recommended Addition

Add these constants to the enum (determine actual column positions from your Excel sheet):

```vba
Public Enum PIFDataColumns
    colArchive = 3           ' C
    colInclude = 4           ' D
    colAccounting = 5        ' E
    colChangeType = 6        ' F
    colPIFID = 7            ' G
    colSEG = 8              ' H
    colOpco = 9             ' I
    colSite = 10            ' J
    colStrategicRank = 11   ' K
    colFromBlanket = 12     ' L
    colFundingProject = 13  ' M
    colProjectName = 14     ' N
    colOriginalISD = 15     ' O
    colRevisedISD = 16      ' P
    colLCMIssue = 17        ' Q
    colStatus = 18          ' R
    colCategory = 19        ' S
    colJustification = 20   ' T
    colMovingISDYear = 21   ' U (EXAMPLE - VERIFY ACTUAL COLUMN!)
    colPriorYearSpend = 22  ' V (EXAMPLE - VERIFY ACTUAL COLUMN!)
End Enum
```

---

## Issue #4: Cost Data Unpivot Logic

**Location:** `mod_Submit.bas` lines 218-299 (UnpivotCostData function)

### Current Code Analysis

The unpivot function reads:
- Line 219: `pifId = wsData.Cells(dataRow, 7).Value` → Column G (correct per enum)
- Line 220: `projectId = wsData.Cells(dataRow, 13).Value` → Column M (funding_project)

This uses **absolute worksheet column positions**, which is correct for reading the main PIF sheet.

### Cost Column Mappings

According to comments in UnpivotCostData:
- **Target Scenario Years (2025-2030):**
  - Requested: Columns U-Z (21-26)
  - Current: Columns AA-AF (27-32)
  - Variance: Columns AG-AL (33-38)

- **Closings Scenario Years (2025-2030):**
  - Requested: Columns AO-AT (41-46)
  - Current: Columns AU-AZ (47-52)
  - Variance: Columns BA-BF (53-58)

**Verify these column positions match your actual Excel layout!**

---

## Action Plan

1. **Open your PIF Excel workbook** and verify the actual column layout
2. **Document the exact column positions** for all fields (especially the missing ones)
3. **Update mod_SharedConstants.bas** with any missing column constants
4. **Fix the duplicate column mappings** in mod_Database.bas lines 549-568
5. **Test the UnpivotCostData** function to ensure cost columns are correct
6. **Consider using the enum constants** instead of hard-coded column numbers throughout the code for better maintainability

---

## Testing Recommendation

Before running the full submission process:

1. Create a **test workbook** with only 2-3 PIF records
2. Run `UploadProjectData()` separately
3. Query the staging table to verify all fields contain the expected values:
   ```sql
   SELECT TOP 3 * FROM dbo.tbl_pif_projects_staging;
   ```
4. Check specifically:
   - Are `project_id` and `funding_project` both populated correctly?
   - Does `lcm_issue` contain the right value (not moving_isd_year data)?
   - Does `prior_year_spend` contain a number (not justification text)?
5. Run `UploadCostData()` separately
6. Query the cost staging table:
   ```sql
   SELECT TOP 12 * FROM dbo.tbl_pif_cost_staging ORDER BY pif_id, scenario, year;
   ```
7. Verify the unpivoted data matches your Excel cost columns

---

## Files to Modify

1. ✅ **PIF_Missing_Stored_Procedures.sql** - Already created, needs to be executed
2. ❌ **mod_Database.bas** - Fix lines 549-568 (duplicate column mappings)
3. ❌ **mod_Submit.bas** - Fix line 410 (dataRange definition) and lines 218-299 (verify cost columns)
4. ❌ **mod_SharedConstants.bas** - Add missing enum constants (moving_isd_year, prior_year_spend)

---

**Last Updated:** 2025-11-06
