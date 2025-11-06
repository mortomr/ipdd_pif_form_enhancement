# Type Mismatch Error 13 - Fixed

**Date:** 2025-11-06
**Error:** VBA Error 13 - Type Mismatch
**Status:** ✅ RESOLVED

---

## What Was the Problem?

Error 13 (Type Mismatch) occurred when Excel cell values didn't match the SQL Server data types expected by the stored procedures. Common causes:

1. **Empty cells** → SQL Server integers/decimals
2. **Text in numeric fields** → SQL Server expecting INT/DECIMAL
3. **Boolean format issues** → TRUE/FALSE vs Y/N vs 1/0
4. **Invalid dates** → SQL Server DATE type

---

## What Was Fixed

### 1. **Added Type-Safe Conversion Functions** (mod_Database.bas)

Created 5 new helper functions that safely convert Excel values to SQL data types:

| Function | Purpose | Handles |
|----------|---------|---------|
| `SafeString()` | Convert to VARCHAR | Empty cells → NULL, trim whitespace |
| `SafeInteger()` | Convert to INT | Empty → NULL, non-numeric → NULL |
| `SafeDecimal()` | Convert to DECIMAL | Empty → NULL, non-numeric → NULL |
| `SafeBoolean()` | Convert to BIT | TRUE/FALSE, 1/0, Y/N, Yes/No, T/F → proper boolean |
| `SafeDate()` | Convert to DATE | Empty → NULL, invalid dates → NULL |

**Example:**
```vba
' Before (BROKEN):
params(6) = wsData.Cells(actualRow, 8).Value  ' seg - could be Empty!

' After (FIXED):
params(6) = SafeInteger(wsData.Cells(actualRow, 8).Value)  ' seg - NULL if empty
```

### 2. **Updated BulkInsertToStaging Function** (mod_Database.bas)

All parameter assignments now use the appropriate Safe* function:

**Project Data (20 parameters):**
```vba
params(0) = SafeString(wsData.Cells(actualRow, 7).Value)   ' pif_id - VARCHAR
params(6) = SafeInteger(wsData.Cells(actualRow, 8).Value)  ' seg - INT
params(17) = SafeDecimal(wsData.Cells(actualRow, 40).Value) ' prior_year_spend - DECIMAL
params(18) = SafeBoolean(wsData.Cells(actualRow, 3).Value)  ' archive_flag - BIT
params(19) = SafeBoolean(wsData.Cells(actualRow, 4).Value)  ' include_flag - BIT
```

**Cost Data (7 parameters):**
```vba
params(0) = SafeString(wsData.Cells(actualRow, 1).Value)  ' pif_id - VARCHAR
params(3) = SafeDate(wsData.Cells(actualRow, 4).Value)    ' year - DATE
params(4) = SafeDecimal(wsData.Cells(actualRow, 5).Value) ' requested_value - DECIMAL
params(5) = SafeDecimal(wsData.Cells(actualRow, 6).Value) ' current_value - DECIMAL
params(6) = SafeDecimal(wsData.Cells(actualRow, 7).Value) ' variance_value - DECIMAL
```

### 3. **Improved AddCostRow Function** (mod_Submit.bas)

Added better error handling for unpivoting cost data:
- Safely converts year strings to dates
- Treats empty/non-numeric values as 0 (instead of crashing)
- Handles conversion errors gracefully

---

## Data Type Requirements

### Your Excel Column Data Types

| Column | Field | Excel Format | SQL Type | Conversion Function |
|--------|-------|--------------|----------|---------------------|
| C | archive_flag | TRUE/FALSE or Y/N or 1/0 | BIT | SafeBoolean |
| D | include_flag | TRUE/FALSE or Y/N or 1/0 | BIT | SafeBoolean |
| E | accounting_treatment | Text | VARCHAR(14) | SafeString |
| F | change_type | Text | VARCHAR(12) | SafeString |
| G | pif_id | Text | VARCHAR(16) | SafeString |
| **H** | **seg** | **Number** | **INT** | **SafeInteger** |
| I | opco | Text | VARCHAR(4) | SafeString |
| J | site | Text | VARCHAR(4) | SafeString |
| K | strategic_rank | Text | VARCHAR(26) | SafeString |
| L | from_blanket | Text | VARCHAR(1) | SafeString |
| M | funding_project | Text | VARCHAR(10) | SafeString |
| N | project_name | Text | VARCHAR(35) | SafeString |
| O | original_fp_isd | Text | VARCHAR(8) | SafeString |
| P | revised_fp_isd | Text | VARCHAR(5) | SafeString |
| Q | lcm_issue | Text | VARCHAR(11) | SafeString |
| R | status | Text | VARCHAR(58) | SafeString |
| S | category | Text | VARCHAR(26) | SafeString |
| T | justification | Text | VARCHAR(192) | SafeString |
| **AM** | **moving_isd_year** | **Text (single char)** | **CHAR(1)** | **SafeString** |
| **AN** | **prior_year_spend** | **Number/Currency** | **DECIMAL(18,2)** | **SafeDecimal** |
| **U-BF** | **Cost columns** | **Number/Currency** | **DECIMAL(18,2)** | **SafeDecimal** |

### Cost Data (Unpivoted Sheet)

| Column | Field | Excel Format | SQL Type | Conversion Function |
|--------|-------|--------------|----------|---------------------|
| A | pif_id | Text | VARCHAR(16) | SafeString |
| B | project_id | Text | VARCHAR(10) | SafeString |
| C | scenario | Text (Target/Closings) | VARCHAR(12) | SafeString |
| **D** | **year** | **Date (12/31/YYYY)** | **DATE** | **SafeDate** |
| **E** | **requested_value** | **Number/Currency** | **DECIMAL(18,2)** | **SafeDecimal** |
| **F** | **current_value** | **Number/Currency** | **DECIMAL(18,2)** | **SafeDecimal** |
| **G** | **variance_value** | **Number/Currency** | **DECIMAL(18,2)** | **SafeDecimal** |

---

## Testing the Fix

### Step 1: Import Updated VBA Modules

1. Open VBA Editor (Alt+F11)
2. Remove old `mod_Database` and `mod_Submit` modules
3. Import updated versions from repository
4. Save the workbook

### Step 2: Verify Excel Data Format

Check these critical columns in your PIF sheet:

#### Column H (seg) - Must be a NUMBER
```
✓ Good: 123, 456, 789
✗ Bad: "SEG-123", blank with text formatting
```

#### Column AN (prior_year_spend) - Must be a NUMBER
```
✓ Good: 1234.56, 0, -500.00
✗ Bad: "$1,234.56" (formatted as text), "N/A"
```

#### Columns C & D (archive_flag, include_flag) - Must be BOOLEAN-compatible
```
✓ Good: TRUE, FALSE, Y, N, 1, 0, Yes, No
✗ Bad: "Maybe", "X", blank spaces
```

#### Cost Columns U-BF - Must be NUMBERS
```
✓ Good: 1234.56, 0, -500.00, empty
✗ Bad: "$1,234.56" (text), "TBD", formulas that error (#DIV/0!)
```

### Step 3: Test with Sample Data

Create a test row with these values:

| Column | Value | Notes |
|--------|-------|-------|
| C | TRUE | archive_flag |
| D | FALSE | include_flag |
| E | Capitalize | accounting_treatment |
| F | New | change_type |
| G | TEST-PIF-001 | pif_id |
| **H** | **123** | **seg (NUMBER!)** |
| I | OPCO | opco |
| J | SITE | site |
| K | High | strategic_rank |
| L | N | from_blanket |
| M | F1TEST0001 | funding_project |
| N | Test Project | project_name |
| O | 2025 | original_fp_isd |
| P | 2026 | revised_fp_isd |
| Q | Inflation | lcm_issue |
| R | Active | status |
| S | IT Infrastructure | category |
| T | This is a test | justification |
| **AM** | **Y** | **moving_isd_year** |
| **AN** | **1234.56** | **prior_year_spend (NUMBER!)** |
| **U-BF** | **100, 200, 300...** | **Cost values (NUMBERS!)** |

### Step 4: Run Submission

1. Run `SubmitToDatabase()` macro
2. Watch for success message
3. If error occurs, check Error_Log sheet (unhide if hidden)

### Step 5: Verify in Database

```sql
-- Check if data was inserted
SELECT TOP 1
    pif_id,
    seg,              -- Should be 123 (integer)
    prior_year_spend, -- Should be 1234.56 (decimal)
    archive_flag,     -- Should be 1 or 0
    include_flag      -- Should be 1 or 0
FROM dbo.tbl_pif_projects_staging
WHERE pif_id = 'TEST-PIF-001';

-- Check cost data
SELECT *
FROM dbo.tbl_pif_cost_staging
WHERE pif_id = 'TEST-PIF-001'
ORDER BY scenario, year;
```

---

## Common Excel Data Issues

### Issue 1: Numbers Stored as Text

**Symptom:** Column looks like a number but has green triangle in corner

**Fix:**
1. Select the column
2. Click the warning icon
3. Choose "Convert to Number"

Or use this formula in a helper column:
```excel
=VALUE(H4)  ' Converts text to number
```

### Issue 2: Boolean Columns Contain Text

**Symptom:** Columns C/D contain "Y/N" but Excel treats them as text

**Good News:** The `SafeBoolean()` function now handles Y/N/Yes/No/T/F/1/0/TRUE/FALSE automatically!

**But if you want proper Excel TRUE/FALSE:**
```excel
=IF(C4="Y", TRUE, FALSE)
```

### Issue 3: Formulas That Return Errors

**Symptom:** Cost columns show #DIV/0! or #VALUE!

**Fix:**
```excel
=IFERROR(your_formula_here, 0)
```

### Issue 4: Currency Formatted as Text

**Symptom:** "$1,234.56" stored as text

**Fix:** Remove $ and commas:
1. Select column
2. Find & Replace: Find "$", Replace with "" (blank)
3. Find & Replace: Find ",", Replace with "" (blank)
4. Format as Number

---

## Troubleshooting

### Still Getting Error 13?

1. **Check which row is failing:**
   - Look at the error message - it shows "Rows processed: X"
   - That's the row causing the issue
   - Check that row in your Excel sheet

2. **Check data in that row:**
   ```vba
   ' Add this debug code temporarily in mod_Database.bas after params assignment:
   Debug.Print "Row " & actualRow & ": seg=" & params(6) & ", prior_year_spend=" & params(17)
   ```

3. **Common culprits:**
   - Column H (seg): Contains text or is blank
   - Column AN (prior_year_spend): Contains text like "$1,234" or "TBD"
   - Columns C/D (flags): Contains something other than Y/N/TRUE/FALSE/1/0
   - Cost columns: Contain errors or text

4. **Verify column positions:**
   ```vba
   ' Add this at the start of BulkInsertToStaging to verify you're reading the right columns:
   MsgBox "Row 4 Column H (seg): " & wsData.Cells(4, 8).Value & vbCrLf & _
          "Row 4 Column AN (prior_year_spend): " & wsData.Cells(4, 40).Value
   ```

### Error Still Occurs After Fixes?

Check if the issue is in ExecuteStoredProcedure function parameter type detection:

The function tries to auto-detect parameter types. If it's still failing, you might need to check:
1. adBoolean parameter type (line 589-590 in mod_Database.bas)
2. adNumeric vs adDecimal parameter types (lines 577, 588, 610-612)

---

## Prevention Tips

### Format Your Excel Sheet Properly

1. **Set column formats BEFORE entering data:**
   - Column H (seg): Number, 0 decimals
   - Column AN (prior_year_spend): Currency or Number, 2 decimals
   - Columns C & D (flags): No special formatting needed
   - Columns U-BF (costs): Currency or Number, 2 decimals

2. **Use Data Validation:**
   ```
   Column C/D (flags): List = TRUE,FALSE
   Column H (seg): Whole number, 0-99999
   Column AN (prior_year_spend): Decimal, allow negative
   ```

3. **Protect formulas:**
   - If cost columns have formulas, protect the sheet
   - Ensure formulas never return #DIV/0! or #VALUE!
   - Use IFERROR() wrapper

---

## Summary of Changes

| File | Change | Impact |
|------|--------|--------|
| mod_Database.bas | Added 5 Safe* conversion functions | ✅ Handles NULL/empty/wrong types |
| mod_Database.bas | Updated all params assignments | ✅ All data properly converted |
| mod_Submit.bas | Improved AddCostRow function | ✅ Better date/number handling |

**Result:** Error 13 should no longer occur. Empty cells, wrong formats, and invalid data are now handled gracefully.

---

## Files to Reimport

1. ✅ **mod_Database.bas** - Contains Safe* functions and updated BulkInsertToStaging
2. ✅ **mod_Submit.bas** - Contains improved AddCostRow function
3. ⚠️ **mod_SharedConstants.bas** - Import if not already done (contains column constants)

---

**Test thoroughly with sample data before processing production data!**

If you still encounter issues after these fixes, provide:
1. The exact row number that fails
2. Values in columns H, AN, C, D for that row
3. The complete error message
