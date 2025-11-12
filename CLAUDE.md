# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Project Impact Form (PIF) reporting system that transforms an Excel-based workflow into a normalized SQL Server database with automated submission workflows. The system uses VBA macros in Excel to validate, transform, and submit PIF data to SQL Server.

## Architecture

### Three-Layer Data Flow
The system follows a staging → inflight → approved pattern:

1. **Staging Tables** (`tbl_pif_projects_staging`, `tbl_pif_cost_staging`)
   - Landing zone for Excel data
   - Truncated and reloaded each submission
   - Used for validation before commit

2. **Inflight Tables** (`tbl_pif_projects_inflight`, `tbl_pif_cost_inflight`)
   - Current working month data
   - Contains active, pending, and parking lot PIFs
   - Replaced entirely each submission

3. **Approved Tables** (`tbl_pif_projects_approved`, `tbl_pif_cost_approved`)
   - Permanent archive of approved/dispositioned PIFs
   - Grows over time, never truncated
   - Source of truth for historical analysis

### VBA Module Responsibilities

- **mod_Database.bas**: Database connectivity, SQL execution, bulk insert operations
  - Handles ADODB connections with configurable timeout settings
  - Connection constants (SQL_SERVER, SQL_DATABASE) must be updated for each environment
  - Uses Windows Authentication (SQL_TRUSTED = True)

- **mod_Validation.bas**: Data validation (Excel-side and SQL-side)
  - Column mapping constants (COL_PIF_ID, COL_STATUS, etc.) define Excel layout
  - Runs both local Excel validation and SQL stored procedure validation
  - Generates Validation_Report sheet with errors

- **mod_Submit.bas**: Orchestrates the complete submission workflow
  - Unpivots wide-format cost data (columns U-BF) into normalized rows
  - Manages transactional commit to ensure atomicity
  - Archives approved PIFs after each submission

- **mod_WorksheetQuery.bas**: Manages Archive and Inflight worksheet QueryTables
  - Creates refreshable QueryTables for viewing approved and inflight PIFs
  - CRITICAL: Properly manages WorkbookConnections to prevent abandoned connections
  - Cleanup logic (lines 264-276) deletes old OLEDB connections before creating new QueryTables
  - Without cleanup, each refresh would leave orphaned connections in SQL Server
  - Native Excel refresh works: right-click "Refresh" or "Data > Refresh All"

### Cost Data Transformation

The system transforms wide-format Excel cost columns into a normalized structure:

**Wide Format (Excel):**
- Target Scenario: Columns U-AL (requested/current/variance for years 2025-2030)
- Closings Scenario: Columns AO-BF (requested/current/variance for years 2025-2030)

**Long Format (Database):**
- Each PIF+Project+Scenario+Year becomes a single row
- Fields: pif_id, project_id, scenario, year, requested_value, current_value, variance_value
- Transformation happens in `UnpivotCostData()` function in mod_Submit.bas

## Database Setup

**Prerequisites:**
- SQL Server 2016+ (2019+ recommended)
- Execute `PIF_Database_DDL.sql` to create schema
- Update `USE [YOUR_DATABASE_NAME];` on line 17 before execution

**Objects Created:**
- 6 tables (staging, inflight, approved for projects and costs)
- 1 submission log table (`tbl_submission_log`)
- 2 views (`vw_pif_current_working`, `vw_pif_all_history`)
- 1 validation stored procedure (`usp_validate_staging_data`)
- 14 indexes for performance

## Excel Setup

**Prerequisites:**
- Excel 2016+ with macros enabled
- ADODB reference: Microsoft ActiveX Data Objects 6.1 Library

**Required Sheets:**
- `PIF_Data`: Main data entry sheet (wide format with cost columns)
- `Validation_Report`: Auto-created validation results
- `Cost_Unpivoted`: Auto-created temporary sheet (hidden)

**Configuration:**
1. Import all three .bas modules
2. Update connection constants in mod_Database.bas (lines 24-26)
3. Verify column mappings in mod_Validation.bas (lines 24-41) match your Excel layout
4. Verify cost column mappings in mod_Submit.bas (UnpivotCostData function, lines 220-302)

## Common Development Tasks

### Testing Database Connection
Run `TestConnection()` macro in mod_Database to verify connectivity and display server info.

### Running Validation Only
Use `RunValidationOnly()` macro to validate data without submitting to database.

### Full Submission Process
The `SubmitToDatabase()` macro executes these steps in order:
1. Unpivot cost data (wide → long format)
2. Create timestamped backup tables
3. Upload project data to staging
4. Upload cost data to staging
5. Run Excel validation checks
6. Run SQL validation via stored procedure
7. Commit to inflight tables (transactional)
8. Archive approved PIFs to permanent tables
9. Log submission to audit table

If any step fails, the transaction rolls back.

### Modifying Column Mappings

When Excel layout changes:
1. Update `COL_*` constants in mod_Validation.bas
2. Update column references in `UnpivotCostData()` in mod_Submit.bas
3. Test with `RunValidationOnly()` before full submission

### Adding Validation Rules

Edit `ValidateBusinessRules()` function in mod_Validation.bas:
```vba
' Add to error collection using format: "Row X|ErrorType|Description"
errors.Add "Row " & i & "|Business Rule Violation|Your error message"
```

## Important Code Locations

- Database connection config: mod_Database.bas:24-26
- Column mappings: mod_Validation.bas:24-41
- Cost unpivot logic: mod_Submit.bas:173-323
- Approval archival: mod_Submit.bas:484-533
- Business rules validation: mod_Validation.bas:294-327

## Data Maintenance

**Monthly:**
- Drop old backup tables: `DROP TABLE dbo.tbl_pif_projects_inflight_backup_YYYYMMDD`
- Review submission log for anomalies

**Quarterly:**
- Rebuild indexes on approved tables (they grow over time)
- Update statistics for query optimization

## Key Business Rules

1. PIF ID + Project ID combination must be unique per submission
2. Approved PIFs require justification text (enforced in validation)
3. Approved or Dispositioned PIFs are automatically archived to permanent tables
4. Variance values are stored from Excel but can be recalculated as requested - current
5. Each submission creates timestamped backups before truncating inflight tables
