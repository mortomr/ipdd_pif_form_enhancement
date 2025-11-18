# Diagnostic Instructions - Type Mismatch Error 13

**Status:** Need your help to diagnose

---

## Changes Made

I've made three critical fixes:

1. **Changed adBoolean → adBit** (correct SQL Server BIT type)
2. **Changed adNumeric → adCurrency** (better for DECIMAL values)
3. **Changed SafeBoolean** to return 1/0 instead of True/False
4. **Added explicit size parameter (0)** for all non-string parameters

These are fundamental ADODB parameter type issues that could cause Error 13.

---

## CRITICAL: I Need Information From You

**Please tell me what the 7 validation errors are!**

You mentioned "same 7 EXACT complaints from the validation macro". What are they? This is crucial because they might be telling us exactly what's wrong with the data.

---

## Step 1: Import Updated Modules

1. Open VBA Editor (Alt+F11)
2. **Import these updated modules:**
   - `mod_Database.bas` (UPDATED - fixed parameter types)
   - `mod_Diagnostic.bas` (NEW - diagnostic tools)
3. Save the workbook

---

## Step 2: Run Diagnostic Test

### Test A: Check Parameter Conversions

1. In VBA Editor, go to mod_Diagnostic
2. Run: `TestParameterTypes`
3. This will show you how the Safe* functions are working
4. **Take a screenshot and share it**

### Test B: Check What's Being Read from Excel

1. In VBA Editor, go to mod_Diagnostic
2. Run: `TestSingleRowInsert`
3. This will show you:
   - What values are being read from row 4
   - What they convert to
   - Whether the stored procedure call succeeds
4. **Take screenshots of ALL message boxes**

### Test C: Show Validation Errors

1. Run your normal validation first (RunValidationOnly)
2. Then in VBA Editor, run: `ShowValidationErrors`
3. **Tell me what all 7 errors are**

---

## Step 3: Check Your Excel Data

### Required Data in Row 4 (First Data Row)

Make sure row 4 has these values:

| Column | Must Contain |
|--------|--------------|
| **C (archive_flag)** | TRUE or FALSE or Y or N or 1 or 0 |
| **D (include_flag)** | TRUE or FALSE or Y or N or 1 or 0 |
| **G (pif_id)** | Some text value (required) |
| **H (seg)** | **A NUMBER** (not text, not empty) |
| **M (funding_project)** | Some text value (required) |
| **F (change_type)** | Some text value (required) |
| **AN (prior_year_spend)** | **A NUMBER** (can be 0, but must be numeric) |

### Quick Check Queries

**Column H (seg) - MUST BE NUMERIC:**
```
Select column H
Look for green triangle (number stored as text)
If present: Click warning → Convert to Number
```

**Column AN (prior_year_spend) - MUST BE NUMERIC:**
```
Select column AN
Look for "$", commas, or text
If present: Remove formatting, ensure it's a plain number
```

**Columns C & D (flags) - Check format:**
```
What's in these cells?
Options: TRUE/FALSE, Y/N, 1/0, Yes/No
If it's something else, change it to one of these
```

---

## Step 4: Check Sheet Name

**CRITICAL:** What is your main data sheet actually named?

Open your Excel workbook and check:
- Is it "PIF"?
- Or is it "PIF_Data"?
- Or something else?

If it's NOT "PIF", update this constant in mod_SharedConstants.bas:
```vba
Public Const SHEET_DATA As String = "YourActualSheetName"
```

---

## Step 5: Verify You Ran the SQL Script

Did you execute `PIF_Missing_Stored_Procedures.sql` on your database?

**Verify with this query:**
```sql
SELECT name, create_date
FROM sys.procedures
WHERE name IN ('usp_insert_project_staging', 'usp_insert_cost_staging');
```

Should show both procedures. If not, the stored procedures don't exist!

---

## What To Report Back

Please provide:

1. **Validation errors** (all 7 of them)
2. **Screenshot** from TestParameterTypes
3. **Screenshots** from TestSingleRowInsert (all message boxes)
4. **Sample data from row 4:**
   - What's in column C (archive_flag)?
   - What's in column D (include_flag)?
   - What's in column H (seg)?
   - What's in column AN (prior_year_spend)?
5. **What is your sheet actually named?**
6. **Stored procedure verification** - do both exist?

---

## Common Causes of Error 13

### Cause 1: Parameter Type Mismatch
**Symptom:** Passing string to integer, or vice versa
**Fix:** The Safe* functions I added should handle this
**Verification:** TestParameterTypes will show conversions

### Cause 2: Wrong ADODB Type
**Symptom:** Using adBoolean for SQL BIT type
**Fix:** Changed to adBit (just fixed this)
**Verification:** TestSingleRowInsert will test actual call

### Cause 3: NULL Handling
**Symptom:** VBA Null not compatible with ADODB parameter
**Fix:** Safe* functions return proper Null variant
**Verification:** TestParameterTypes shows NULL handling

### Cause 4: Data Format Issues
**Symptom:** "123" stored as text in numeric column
**Fix:** Check Excel cell formats
**Verification:** Look for green triangles in numeric columns

### Cause 5: Sheet Name Mismatch
**Symptom:** Code looks for "PIF" but sheet is "PIF_Data"
**Fix:** Update SHEET_DATA constant
**Verification:** Check actual sheet name in workbook

---

## If TestSingleRowInsert Still Fails

If you run TestSingleRowInsert and still get Error 13, the message box will tell us:
1. The exact parameter values being passed
2. The types of those parameters
3. Which specific parameter causes the error

This will let me pinpoint the exact issue.

---

## Quick Sanity Checks

### Check 1: Is ADODB Referenced?
1. VBA Editor → Tools → References
2. Look for "Microsoft ActiveX Data Objects 6.1 Library"
3. Should be CHECKED
4. If not, check it and click OK

### Check 2: Can You Connect?
1. Run: `TestConnection` (from mod_Database)
2. Should show server info
3. If fails, database connection is the problem

### Check 3: Do Tables Exist?
```sql
SELECT name
FROM sys.tables
WHERE name LIKE 'tbl_pif%staging';
```
Should show:
- tbl_pif_projects_staging
- tbl_pif_cost_staging

---

## Expected TestSingleRowInsert Output

If everything works, you should see:

**Message 1: Raw Excel Values**
```
DIAGNOSTIC: Reading data from row 4

Column G (pif_id): TEST-001
Column M (project_id): F1TEST0001
Column R (status): Active
...
Column H (seg): [123] IsNumeric=True
Column AN (prior_year_spend): [1234.56] IsNumeric=True
Column C (archive_flag): TRUE
Column D (include_flag): FALSE
```

**Message 2: Converted Parameter Values**
```
CONVERTED VALUES:

params(0) = [TEST-001] Type=String
params(1) = [F1TEST0001] Type=String
...
params(6) = [123] Type=Long
params(17) = [1234.56] Type=Double
params(18) = [1] Type=Integer
params(19) = [0] Type=Integer
```

**Message 3: Success or Failure**
```
SUCCESS! Row inserted without errors.
Check tbl_pif_projects_staging to verify data.
```

OR

```
STORED PROCEDURE FAILED!

Error Number: 13
Error Description: Type mismatch

This is the exact error occurring in BulkInsertToStaging
```

---

## Next Steps After Diagnostics

Once you run the diagnostics and report back:
1. I'll know exactly which parameter is failing
2. I'll know what value is in that parameter
3. I'll know what type it's being converted to
4. I can create a targeted fix

The validation errors are also crucial - they might tell us exactly what data is invalid.

---

**Run the diagnostics and report back with all the information above. This will let me give you a precise fix!**
