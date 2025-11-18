# ADODB Constant Fix for Error 3421

## Problem
Error 3421 occurs when ADODB constants like `adBit`, `adCurrency`, `adChar` are undefined in your VBA environment. When undefined, VBA treats them as 0, causing stored procedure calls to fail.

## Solution
Replaced all ADODB constants with their **numeric values** directly.

## Changed Values

| Old Constant | Numeric Value | Data Type | Used For |
|--------------|---------------|-----------|----------|
| `adVarChar` | 200 | VARCHAR/NVARCHAR | All text fields |
| `adChar` | 129 | CHAR | moving_isd_year (1 char) |
| `adInteger` | 3 | INT | seg column |
| `adNumeric` | 131 | DECIMAL/NUMERIC | prior_year_spend, all cost values |
| `adBit` | 11 | BIT | archive_flag, include_flag |
| `adDate` | 7 | DATE | year column in cost data |
| `adParamInput` | 1 | Input parameter | All parameters |

## Files Updated

### 1. mod_Database.bas
- Line ~579-599: `usp_insert_project_staging` call
- Line ~615-622: `usp_insert_cost_staging` call

### 2. mod_Diagnostic.bas
- Line ~107-127: Diagnostic stored procedure call

## Before (Using Constants - BROKEN)
```vba
ExecuteStoredProcedure(conn, "usp_insert_project_staging", False, _
    "@pif_id", adVarChar, adParamInput, 16, params(0), _
    "@seg", adInteger, adParamInput, 0, params(6), _
    "@prior_year_spend", adNumeric, adParamInput, 0, params(17), _
    "@archive_flag", adBit, adParamInput, 0, params(18))
```

## After (Using Numeric Values - FIXED)
```vba
ExecuteStoredProcedure(conn, "usp_insert_project_staging", False, _
    "@pif_id", 200, 1, 16, params(0), _
    "@seg", 3, 1, 0, params(6), _
    "@prior_year_spend", 131, 1, 0, params(17), _
    "@archive_flag", 11, 1, 0, params(18))
```

## Parameter Format
Each parameter now uses:
```vba
"@parameter_name", DataType, Direction, Size, Value
```

Where:
- **DataType**: Numeric value (200=VARCHAR, 3=INT, 131=DECIMAL, 11=BIT, 7=DATE, 129=CHAR)
- **Direction**: 1 (adParamInput - always 1 for input parameters)
- **Size**:
  - For strings: max length (16, 10, 58, etc.)
  - For numbers: 0 (not needed)
- **Value**: The actual parameter value from params() array

## Complete Parameter Mapping

### usp_insert_project_staging (20 parameters)

| Parameter | Type | Size | Value Source |
|-----------|------|------|--------------|
| @pif_id | 200 (VARCHAR) | 16 | params(0) - Column G |
| @project_id | 200 (VARCHAR) | 10 | params(1) - Column M |
| @status | 200 (VARCHAR) | 58 | params(2) - Column R |
| @change_type | 200 (VARCHAR) | 12 | params(3) - Column F |
| @accounting_treatment | 200 (VARCHAR) | 14 | params(4) - Column E |
| @category | 200 (VARCHAR) | 26 | params(5) - Column S |
| @seg | 3 (INT) | 0 | params(6) - Column H |
| @opco | 200 (VARCHAR) | 4 | params(7) - Column I |
| @site | 200 (VARCHAR) | 4 | params(8) - Column J |
| @strategic_rank | 200 (VARCHAR) | 26 | params(9) - Column K |
| @funding_project | 200 (VARCHAR) | 10 | params(10) - Column M |
| @project_name | 200 (VARCHAR) | 35 | params(11) - Column N |
| @original_fp_isd | 200 (VARCHAR) | 8 | params(12) - Column O |
| @revised_fp_isd | 200 (VARCHAR) | 5 | params(13) - Column P |
| @moving_isd_year | 129 (CHAR) | 1 | params(14) - Column AM |
| @lcm_issue | 200 (VARCHAR) | 11 | params(15) - Column Q |
| @justification | 200 (VARCHAR) | 192 | params(16) - Column T |
| @prior_year_spend | 131 (DECIMAL) | 0 | params(17) - Column AN |
| @archive_flag | 11 (BIT) | 0 | params(18) - Column C |
| @include_flag | 11 (BIT) | 0 | params(19) - Column D |

### usp_insert_cost_staging (7 parameters)

| Parameter | Type | Size | Value Source |
|-----------|------|------|--------------|
| @pif_id | 200 (VARCHAR) | 16 | params(0) - Column A |
| @project_id | 200 (VARCHAR) | 10 | params(1) - Column B |
| @scenario | 200 (VARCHAR) | 12 | params(2) - Column C |
| @year | 7 (DATE) | 0 | params(3) - Column D |
| @requested_value | 131 (DECIMAL) | 0 | params(4) - Column E |
| @current_value | 131 (DECIMAL) | 0 | params(5) - Column F |
| @variance_value | 131 (DECIMAL) | 0 | params(6) - Column G |

## Why This Works

Using numeric values:
1. ✅ **No dependency on ADODB constant definitions**
2. ✅ **Works across all ADODB library versions**
3. ✅ **No "Variable not defined" compile errors**
4. ✅ **Eliminates Error 3421**

## Testing

After reimporting the updated modules:

1. **Compile the VBA project:**
   - Debug → Compile VBAProject
   - Should complete with no errors

2. **Run diagnostic:**
   ```vba
   TestSingleRowInsert
   ```
   Should now succeed and insert data

3. **Verify in SQL:**
   ```sql
   SELECT * FROM dbo.tbl_pif_projects_staging;
   ```
   Should show your inserted row

## Reference: ADODB DataTypeEnum Values

For future reference:

| Constant Name | Numeric Value | SQL Server Type |
|---------------|---------------|-----------------|
| adEmpty | 0 | (invalid) |
| adTinyInt | 16 | TINYINT |
| adSmallInt | 2 | SMALLINT |
| adInteger | 3 | INT |
| adBigInt | 20 | BIGINT |
| adUnsignedTinyInt | 17 | TINYINT |
| adUnsignedSmallInt | 18 | SMALLINT |
| adUnsignedInt | 19 | INT |
| adSingle | 4 | REAL |
| adDouble | 5 | FLOAT |
| adCurrency | 6 | MONEY |
| adDecimal | 14 | DECIMAL |
| adNumeric | 131 | NUMERIC |
| adBoolean | 11 | BIT |
| adError | 10 | (invalid) |
| adUserDefined | 132 | (invalid) |
| adVariant | 12 | sql_variant |
| adIDispatch | 9 | (invalid) |
| adIUnknown | 13 | (invalid) |
| adGUID | 72 | uniqueidentifier |
| adDate | 7 | DATE |
| adDBDate | 133 | DATE |
| adDBTime | 134 | TIME |
| adDBTimeStamp | 135 | DATETIME |
| adBSTR | 8 | (invalid) |
| adChar | 129 | CHAR |
| adVarChar | 200 | VARCHAR |
| adLongVarChar | 201 | TEXT |
| adWChar | 130 | NCHAR |
| adVarWChar | 202 | NVARCHAR |
| adLongVarWChar | 203 | NTEXT |
| adBinary | 128 | BINARY |
| adVarBinary | 204 | VARBINARY |
| adLongVarBinary | 205 | IMAGE |
| adChapter | 136 | (invalid) |
| adFileTime | 64 | (invalid) |
| adPropVariant | 138 | (invalid) |
| adVarNumeric | 139 | DECIMAL |
| adArray | 0x2000 | (array flag) |

## Common ADODB Constants We Use

| Constant | Value | Purpose |
|----------|-------|---------|
| adParamInput | 1 | Input parameter |
| adParamOutput | 2 | Output parameter |
| adParamInputOutput | 3 | Input/Output parameter |
| adParamReturnValue | 4 | Return value |
| adCmdText | 1 | SQL text command |
| adCmdStoredProc | 4 | Stored procedure |

---

**After reimporting these updated files, Error 3421 should be resolved!**
