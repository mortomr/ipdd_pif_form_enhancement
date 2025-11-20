# Migration Guide: Adding line_item Field to PIF System

**Date:** 2025-11-20
**Purpose:** Allow multiple detail lines for the same PIF_ID + Project_ID combination
**Impact:** Database schema, stored procedures, VBA code, Excel workbook

---

## Table of Contents
1. [Overview](#overview)
2. [Database Migration](#database-migration)
3. [Excel Workbook Changes](#excel-workbook-changes)
4. [VBA Code Update](#vba-code-update)
5. [Testing](#testing)
6. [Rollback Plan](#rollback-plan)

---

## Overview

### Problem Statement
The original system enforced a UNIQUE constraint on `(pif_id, project_id)`, which prevented users from entering multiple rows with different justifications for the same PIF and Project combination.

**Real-world example:**
- Row 1: PIF-123 + Project-456 - "Low payroll estimate"
- Row 2: PIF-123 + Project-456 - "Emergent outage scope"

This was rejected as a duplicate, even though both represent valid, separate changes to be funded.

### Solution
Add a `line_item` field (integer) as part of the composite key:
- **Old Key:** `(pif_id, project_id)`
- **New Key:** `(pif_id, project_id, line_item)`

Users number their detail lines 1, 2, 3, etc. If only one line exists, it defaults to 1.

### Affected Components
- 6 database tables (staging, inflight, approved for projects and costs)
- 6 stored procedures
- 3 VBA modules (mod_Validation, mod_Submit, mod_Database)
- Excel PIF worksheet (insert new column G)

---

## Database Migration

### Step 1: Run Schema Migration

Execute this script on your SQL Server instance:

```sql
-- File: PIF_Database_Add_LineItem.sql
USE [IPDD];
GO

-- Add line_item column to all 6 tables
-- Creates indexes and UNIQUE constraints
-- Defaults existing data to line_item = 1
```

**Run:**
```bash
sqlcmd -S YOUR_SERVER -d IPDD -i PIF_Database_Add_LineItem.sql
```

**What it does:**
1. Adds `line_item INT NOT NULL DEFAULT 1` to all 6 tables
2. Drops old UNIQUE constraint `UQ_inflight_pif_project`
3. Creates new UNIQUE constraint `UQ_inflight_pif_project_line (pif_id, project_id, line_item)`
4. Updates 3 indexes to include `line_item`

### Step 2: Update Stored Procedures

Execute this script on your SQL Server instance:

```sql
-- File: PIF_Database_Update_SPs_LineItem.sql
USE [IPDD];
GO

-- Updates all 6 stored procedures to handle line_item parameter
-- Updates duplicate validation logic
```

**Run:**
```bash
sqlcmd -S YOUR_SERVER -d IPDD -i PIF_Database_Update_SPs_LineItem.sql
```

**What it does:**
1. Updates `usp_insert_project_staging` - adds `@line_item` parameter (defaults to 1)
2. Updates `usp_insert_cost_staging` - adds `@line_item` parameter (defaults to 1)
3. Updates `usp_validate_staging_data_secure` - changes duplicate check to include `line_item`
4. Updates `usp_commit_to_inflight` - includes `line_item` in INSERT/DELETE operations
5. Updates `usp_archive_approved_pifs` - includes `line_item` in MERGE operations
6. Updates `usp_create_staging_indexes` - includes `line_item` in index definitions

### Step 3: Verify Database Changes

```sql
-- Check that line_item column exists in all tables
SELECT
    t.name AS TableName,
    c.name AS ColumnName,
    ty.name AS DataType
FROM sys.tables t
INNER JOIN sys.columns c ON t.object_id = c.object_id
INNER JOIN sys.types ty ON c.user_type_id = ty.user_type_id
WHERE t.name LIKE 'tbl_pif%'
  AND c.name = 'line_item'
ORDER BY t.name;

-- Expected: 6 rows (all tables should have line_item column)

-- Check UNIQUE constraint
SELECT
    kc.name AS ConstraintName,
    OBJECT_NAME(kc.parent_object_id) AS TableName,
    COL_NAME(kc.parent_object_id, ic.column_id) AS ColumnName
FROM sys.key_constraints kc
INNER JOIN sys.index_columns ic ON kc.parent_object_id = ic.object_id AND kc.unique_index_id = ic.index_id
WHERE kc.name = 'UQ_inflight_pif_project_line'
ORDER BY ic.key_ordinal;

-- Expected: 3 rows (pif_id, project_id, line_item)
```

---

## Excel Workbook Changes

### Step 1: Insert New Column in PIF Worksheet

**IMPORTANT:** All users must update their Excel workbooks before submitting data.

1. Open your PIF workbook
2. Navigate to the `PIF` worksheet
3. **Insert a new column G** (before the current PIF_ID column)
4. Add header in row 3: `Line Item`
5. Add description in row 2 (optional): `Detail line number (1, 2, 3, etc.). Leave blank for single-line PIFs.`

**New Column Layout:**
```
Column F: Change Type
Column G: Line Item     <-- NEW COLUMN
Column H: PIF ID        <-- (was column G)
Column I: SEG           <-- (was column H)
...all subsequent columns shift right by 1
```

### Step 2: Update Column Headers

After inserting column G, verify that all headers are correctly aligned:

```
C: Archive
D: Include
E: Accounting Treatment
F: Change Type
G: Line Item (NEW)
H: PIF ID
I: SEG
J: OPCO
K: Site
L: Strategic Rank
M: From Blanket
N: Funding Project
O: Project Name
P: Original ISD
Q: Revised ISD
R: LCM Issue
S: Status
T: Category
U: Justification (was column T)
```

### Step 3: Format New Column

1. Select column G (Line Item)
2. Format as **Number** (integer, no decimals)
3. **Optional:** Add data validation
   - Data > Data Validation
   - Allow: Whole number
   - Minimum: 1
   - Maximum: 99

### Step 4: Add User Guidance

Add a note or comment in cell G2:
```
Enter 1, 2, 3, etc. if you have multiple detail lines for the same PIF+Project.
Leave blank if only one line (will default to 1).

Example:
  Row 4: PIF-123, Project-456, Line 1, "Low payroll estimate"
  Row 5: PIF-123, Project-456, Line 2, "Emergent outage scope"
```

---

## VBA Code Update

### Import Updated Modules

**File:** `mod_Validation.bas`
- Updated column constants (COL_LINE_ITEM = 7, all others shifted +1)
- Updated duplicate validation to check `pif_id + project_id + line_item`
- Added validation for line_item (must be integer >= 1, defaults to 1 if blank)

**File:** `mod_Submit.bas`
- Updated UnpivotCostData to include line_item in output (8 columns instead of 7)
- Updated column references to account for new column G

**File:** `mod_Database.bas`
- Updated BulkInsertToStaging to pass line_item to stored procedures
- Projects: 21 parameters (was 20)
- Costs: 8 parameters (was 7)

### Verify VBA References

After importing, check that these constants are correct:

**mod_Validation.bas:**
```vba
Private Const COL_LINE_ITEM As Integer = 7          ' G (NEW)
Private Const COL_PIF_ID As Integer = 8             ' H (was 7)
Private Const COL_FUNDING_PROJECT As Integer = 14   ' N (was 13)
Private Const COL_JUSTIFICATION As Integer = 21     ' U (was 20)
```

---

## Testing

### Test Plan

#### Test 1: Single Line Item (Backward Compatibility)
**Scenario:** User leaves line_item blank (defaults to 1)

```
PIF_ID  | Project_ID | Line_Item | Justification
--------|------------|-----------|-------------------
PIF-001 | PROJ-123   | (blank)   | Budget adjustment
```

**Expected:**
- Validation passes
- Data is submitted with `line_item = 1`
- No errors

#### Test 2: Multiple Line Items (New Functionality)
**Scenario:** User enters 2 detail lines for same PIF+Project

```
PIF_ID  | Project_ID | Line_Item | Justification
--------|------------|-----------|------------------------
PIF-002 | PROJ-456   | 1         | Low payroll estimate
PIF-002 | PROJ-456   | 2         | Emergent outage scope
```

**Expected:**
- Validation passes
- Both rows are accepted
- Cost data is correctly unpivoted with line_item preserved
- No duplicate error

#### Test 3: Duplicate Detection Still Works
**Scenario:** User enters same PIF+Project+Line twice

```
PIF_ID  | Project_ID | Line_Item | Justification
--------|------------|-----------|-------------------
PIF-003 | PROJ-789   | 1         | Scope change
PIF-003 | PROJ-789   | 1         | Scope change (duplicate)
```

**Expected:**
- Validation **FAILS**
- Error: "Duplicate Entry: PIF PIF-003 + Project PROJ-789 + Line 1"

#### Test 4: Cost Data Integrity
**Scenario:** Verify cost data does not double-count

**Setup:**
- Enter 1 row: PIF-004, PROJ-100, Line 1, $1,000 requested
- Submit to inflight

**Verify:**
```sql
-- Should return 12 rows (6 Target + 6 Closings years) for THIS LINE ITEM ONLY
SELECT * FROM tbl_pif_cost_inflight
WHERE pif_id = 'PIF-004' AND project_id = 'PROJ-100' AND line_item = 1;

-- Should return $1,000 for CY Target Requested
SELECT requested_value FROM tbl_pif_cost_inflight
WHERE pif_id = 'PIF-004' AND project_id = 'PROJ-100' AND line_item = 1
  AND scenario = 'Target' AND YEAR(year) = (SELECT CurrentYear FROM ...);
```

**Expected:**
- Exactly 12 cost rows per project line item
- Values match Excel input
- No double-counting

#### Test 5: Archive and Cleanup
**Scenario:** Archive a multi-line PIF

**Setup:**
- Submit PIF-005 with 2 line items
- Mark both as Approved
- Run Finalize Month

**Verify:**
```sql
-- Should show 2 rows in approved table
SELECT * FROM tbl_pif_projects_approved
WHERE pif_id = 'PIF-005';

-- Should show 24 cost rows (12 per line item)
SELECT COUNT(*) FROM tbl_pif_cost_approved
WHERE pif_id = 'PIF-005';
```

---

## Rollback Plan

If issues arise, you can rollback changes:

### Database Rollback

```sql
USE [IPDD];
GO

-- Restore old UNIQUE constraint on inflight table
ALTER TABLE dbo.tbl_pif_projects_inflight
DROP CONSTRAINT UQ_inflight_pif_project_line;

ALTER TABLE dbo.tbl_pif_projects_inflight
ADD CONSTRAINT UQ_inflight_pif_project UNIQUE (pif_id, project_id);

-- Remove line_item columns (WARNING: This will delete line_item data!)
ALTER TABLE dbo.tbl_pif_projects_staging DROP COLUMN line_item;
ALTER TABLE dbo.tbl_pif_projects_inflight DROP COLUMN line_item;
ALTER TABLE dbo.tbl_pif_projects_approved DROP COLUMN line_item;
ALTER TABLE dbo.tbl_pif_cost_staging DROP COLUMN line_item;
ALTER TABLE dbo.tbl_pif_cost_inflight DROP COLUMN line_item;
ALTER TABLE dbo.tbl_pif_cost_approved DROP COLUMN line_item;
```

**IMPORTANT:** Rollback will fail if any multi-line PIFs exist in the database (UNIQUE constraint violation). You must delete multi-line records first.

### VBA Rollback

1. Restore old VBA module versions from git commit before migration
2. Re-import old modules

### Excel Rollback

1. Delete column G (Line Item)
2. Revert to previous workbook version

---

## Post-Migration Checklist

- [ ] Database schema migration completed
- [ ] Stored procedures updated
- [ ] Test data submitted successfully
- [ ] Validation rules working correctly
- [ ] Cost unpivot includes line_item
- [ ] Archive process includes line_item
- [ ] All users notified of Excel workbook change
- [ ] Documentation updated (CLAUDE.md)
- [ ] Git commit with migration tag

---

## Support

If you encounter issues:

1. Check validation errors on `Validation_Report` sheet
2. Verify column G is formatted as integer
3. Check that line_item values start at 1 (not 0)
4. Review SQL error messages for constraint violations
5. Contact development team with error details

---

## FAQ

**Q: What if I forget to add a line_item value?**
A: It will default to 1. Single-line PIFs work exactly as before.

**Q: Can I use line_item values like 10, 20, 30 for sorting flexibility?**
A: Yes! Any positive integer works. Users often leave gaps for future insertions.

**Q: What happens if I submit line_item 1 and 3, but skip 2?**
A: That's fine. Line items don't have to be sequential, just unique within a PIF+Project.

**Q: Can justification be blank for multi-line PIFs?**
A: Yes for non-approved PIFs. Approved PIFs always require justification (per existing rule).

**Q: Will this break my existing data?**
A: No. All existing records get `line_item = 1` by default. The system is backward compatible.

---

**End of Migration Guide**
