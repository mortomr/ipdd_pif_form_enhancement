# PIF Reporting Database Setup and VBA Integration

## Overview

This package contains the complete database schema and VBA code for the Project Impact Form (PIF) reporting system. It transforms the Excel-based PIF workbook into a normalized SQL database with automated submission workflows.

## Package Contents

1. **PIF_Database_DDL.sql** - Complete SQL Server database setup
2. **mod_Database.bas** - VBA module for database connectivity
3. **mod_Validation.bas** - VBA module for data validation
4. **mod_Submit.bas** - VBA module for submission orchestration

## Prerequisites

### Database Requirements
- SQL Server 2016 or later (2019+ recommended)
- Database with appropriate permissions for:
  - CREATE TABLE
  - CREATE INDEX
  - CREATE VIEW
  - CREATE PROCEDURE
  - INSERT/UPDATE/DELETE on created tables

### Excel Requirements
- Microsoft Excel 2016 or later (Microsoft 365 recommended)
- Macros must be enabled
- Reference Library Required: **Microsoft ActiveX Data Objects 6.1**
  - In VBA Editor: Tools > References > Check "Microsoft ActiveX Data Objects 6.1 Library"

### Network Requirements
- Network connectivity to SQL Server
- Windows Authentication enabled (recommended)
- Firewall rules allowing SQL Server connection (default port 1433)

---

## Installation Instructions

### Part 1: Database Setup

1. **Open SQL Server Management Studio (SSMS)**

2. **Connect to your SQL Server instance**

3. **Create or select a database**
   ```sql
   -- Option A: Create new database
   CREATE DATABASE PIF_Reporting;
   GO
   
   USE PIF_Reporting;
   GO
   
   -- Option B: Use existing database
   USE [YOUR_EXISTING_DATABASE];
   GO
   ```

4. **Open PIF_Database_DDL.sql in SSMS**
   - Update the first line: `USE [YOUR_DATABASE_NAME];`

5. **Execute the script** (F5 or Execute button)
   - This creates:
     - 6 tables (staging, inflight, approved)
     - Multiple indexes
     - 2 views
     - 1 stored procedure
     - Audit logging table

6. **Verify installation**
   ```sql
   -- Check that all tables were created
   SELECT TABLE_NAME 
   FROM INFORMATION_SCHEMA.TABLES 
   WHERE TABLE_NAME LIKE 'tbl_pif%'
   ORDER BY TABLE_NAME;
   
   -- Should return:
   -- tbl_pif_cost_approved
   -- tbl_pif_cost_inflight
   -- tbl_pif_cost_staging
   -- tbl_pif_projects_approved
   -- tbl_pif_projects_inflight
   -- tbl_pif_projects_staging
   -- tbl_submission_log
   ```

### Part 2: Excel VBA Setup

1. **Open your PIF Excel workbook (.xlsm format required)**

2. **Open VBA Editor** (Alt + F11)

3. **Add the required reference**
   - Tools > References
   - Scroll down and check: **Microsoft ActiveX Data Objects 6.1 Library**
   - Click OK

4. **Import VBA modules**
   
   For each .bas file:
   - File > Import File
   - Select the .bas file
   - Click Open
   
   Import all three modules:
   - mod_Database.bas
   - mod_Validation.bas
   - mod_Submit.bas

5. **Configure database connection**
   
   In **mod_Database** module, update these constants (lines 21-23):
   ```vba
   Private Const SQL_SERVER As String = "YOUR_SQL_SERVER_NAME"
   Private Const SQL_DATABASE As String = "YOUR_DATABASE_NAME"
   Private Const SQL_TRUSTED As Boolean = True  ' Keep True for Windows Auth
   ```
   
   Example:
   ```vba
   Private Const SQL_SERVER As String = "SQLPROD01"
   Private Const SQL_DATABASE As String = "PIF_Reporting"
   Private Const SQL_TRUSTED As Boolean = True
   ```

6. **Test the connection**
   - In VBA Editor, click in the mod_Database module
   - Press F5 (or Run > Run Sub/UserForm)
   - Select `TestConnection` from the list
   - Click Run
   - You should see a success message with server info

7. **Create submission buttons** (optional but recommended)
   
   Add buttons to your Excel sheet:
   - Developer tab > Insert > Button (Form Control)
   - Draw button on sheet
   - Assign macro: `mod_Submit.SubmitToDatabase`
   - Label the button: "Submit to Database"
   
   Create a second button for validation only:
   - Assign macro: `mod_Submit.RunValidationOnly`
   - Label: "Validate Data"

---

## Excel Workbook Structure Requirements

Your Excel workbook must have the following sheets:

### Required Sheets

1. **PIF_Data** (main data entry sheet)
   - Contains all project and PIF metadata
   - Wide format with cost columns

2. **Validation_Report** (auto-created if missing)
   - Shows validation errors
   - Generated automatically by validation process

3. **Cost_Unpivoted** (auto-created and hidden)
   - Temporary sheet for normalized cost data
   - Created automatically during submission

### Column Mapping

The VBA code expects specific columns in the PIF_Data sheet. **Update these constants in mod_Validation.bas if your layout differs:**

```vba
' Lines 15-29 in mod_Validation.bas
Private Const COL_PIF_ID As Integer = 7            ' Column G
Private Const COL_FUNDING_PROJECT As Integer = 13  ' Column M (acts as project_id)
Private Const COL_STATUS As Integer = 18           ' Column R
Private Const COL_JUSTIFICATION As Integer = 20    ' Column T
' etc...
```

**Cost columns mapping** (in mod_Submit.bas):

Target Scenario:
- CY (2025): Columns U (requested), AA (current), AG (variance)
- CY+1 (2026): Columns V, AB, AH
- CY+2 (2027): Columns W, AC, AI
- CY+3 (2028): Columns X, AD, AJ
- CY+4 (2029): Columns Y, AE, AK
- CY+5 (2030): Columns Z, AF, AL

Closings Scenario:
- CY (2025): Columns AO, AU, BA
- CY+1 (2026): Columns AP, AV, BB
- CY+2 (2027): Columns AQ, AW, BC
- CY+3 (2028): Columns AR, AX, BD
- CY+4 (2029): Columns AS, AY, BE
- CY+5 (2030): Columns AT, AZ, BF

---

## Usage Instructions

### Monthly Submission Process

1. **Prepare the workbook**
   - Ensure all data is entered and reviewed
   - Obtain necessary approvals
   - Save a backup copy

2. **Run validation (optional but recommended)**
   - Click "Validate Data" button or run `RunValidationOnly` macro
   - Review any errors in the Validation_Report sheet
   - Fix errors and re-validate until clean

3. **Submit to database**
   - Click "Submit to Database" button or run `SubmitToDatabase` macro
   - Confirm the submission prompt
   - Wait for completion message (typically 30-60 seconds)

4. **Verify submission**
   - Check SQL Server for new records:
   ```sql
   -- Check submission log
   SELECT TOP 5 * 
   FROM dbo.tbl_submission_log 
   ORDER BY submission_date DESC;
   
   -- Check inflight records
   SELECT COUNT(*) AS TotalRecords
   FROM dbo.tbl_pif_projects_inflight;
   
   -- Check approved records
   SELECT COUNT(*) AS ApprovedRecords
   FROM dbo.tbl_pif_projects_approved;
   ```

### What Happens During Submission

The submission process executes these steps automatically:

1. **Unpivot Cost Data** - Transforms wide cost columns into normalized rows
2. **Create Backups** - Saves current inflight data with timestamp
3. **Upload to Staging** - Loads Excel data into staging tables
4. **Validate** - Runs Excel and SQL validation checks
5. **Commit** - Moves staging data to inflight tables (within transaction)
6. **Archive** - Moves approved PIFs to permanent archive
7. **Log** - Records submission in audit log

If any step fails, the entire process rolls back.

---

## Troubleshooting

### Connection Errors

**Error: "Database connection failed"**
- Verify SQL_SERVER constant matches your server name
- Test with SSMS first using same credentials
- Check firewall rules for SQL Server port
- Ensure Windows Authentication is working

**Error: "Login failed for user"**
- Verify you have permissions on the database
- Try connecting with SSMS to confirm credentials
- Contact your DBA for permission issues

### VBA Errors

**Error: "Compile error: User-defined type not defined"**
- You're missing the ADODB reference
- Go to Tools > References > Check "Microsoft ActiveX Data Objects 6.1"

**Error: "Object required" or "Nothing" errors**
- Check that all sheet names match your workbook
- Update sheet name constants in the VBA modules

**Error: "Column index out of range"**
- Your Excel layout differs from expected structure
- Update column constants in mod_Validation.bas and mod_Submit.bas

### Data Errors

**Validation fails with duplicate PIF+Project**
- Check for copy-paste errors in your data
- Each PIF+Project combination must be unique

**Approved PIFs missing justification**
- Business rule: All approved PIFs require justification text
- Fill in the justification column before submitting

### Performance Issues

**Submission takes a long time (>5 minutes)**
- Check for network latency to SQL Server
- Consider running during off-peak hours
- Check if antivirus is scanning the file during execution

---

## Database Maintenance

### Monthly Tasks

1. **Archive old backup tables**
   ```sql
   -- List backup tables
   SELECT name 
   FROM sys.tables 
   WHERE name LIKE '%backup_%'
   ORDER BY name;
   
   -- Drop old backups (adjust date as needed)
   DROP TABLE dbo.tbl_pif_projects_inflight_backup_20250101;
   DROP TABLE dbo.tbl_pif_cost_inflight_backup_20250101;
   ```

2. **Review submission log**
   ```sql
   -- View recent submissions
   SELECT 
       submission_date,
       submitted_by,
       source_file,
       record_count
   FROM dbo.tbl_submission_log
   WHERE submission_date >= DATEADD(MONTH, -3, GETDATE())
   ORDER BY submission_date DESC;
   ```

### Quarterly Tasks

1. **Index maintenance**
   ```sql
   -- Rebuild indexes on approved tables (they grow over time)
   ALTER INDEX ALL ON dbo.tbl_pif_projects_approved REBUILD;
   ALTER INDEX ALL ON dbo.tbl_pif_cost_approved REBUILD;
   ```

2. **Update statistics**
   ```sql
   UPDATE STATISTICS dbo.tbl_pif_projects_approved;
   UPDATE STATISTICS dbo.tbl_pif_cost_approved;
   ```

---

## Useful SQL Queries

### View Current Working PIFs
```sql
SELECT * 
FROM dbo.vw_pif_current_working
WHERE status NOT IN ('Approved', 'Dispositioned')
ORDER BY pif_id;
```

### View All History (Approved + Inflight)
```sql
SELECT *
FROM dbo.vw_pif_all_history
ORDER BY submission_date DESC, pif_id;
```

### Variance Analysis by Project
```sql
SELECT 
    project_id,
    project_name,
    SUM(variance_value) AS total_variance
FROM dbo.tbl_pif_cost_approved c
JOIN dbo.tbl_pif_projects_approved p
    ON c.pif_id = p.pif_id AND c.project_id = p.project_id
GROUP BY project_id, project_name
ORDER BY total_variance DESC;
```

### Monthly Approval Summary
```sql
SELECT 
    YEAR(approval_date) AS Year,
    MONTH(approval_date) AS Month,
    COUNT(DISTINCT pif_id) AS PIFs_Approved,
    SUM(variance_value) AS Total_Variance
FROM dbo.tbl_pif_cost_approved
GROUP BY YEAR(approval_date), MONTH(approval_date)
ORDER BY Year DESC, Month DESC;
```

---

## Customization Guide

### Adding New Validation Rules

Edit **mod_Validation.bas**, function `ValidateBusinessRules`:

```vba
' Example: Add rule to check that category is not empty
If Trim(wsData.Cells(i, COL_CATEGORY).Value) = "" Then
    errors.Add "Row " & i & "|Business Rule Violation|Category is required"
End If
```

### Modifying Column Mappings

If your Excel layout changes:

1. Update constants in **mod_Validation.bas** (lines 15-29)
2. Update cost column references in **mod_Submit.bas** (UnpivotCostData function)
3. Test thoroughly with validation before submitting

### Adding New Status Values

If you add new status types beyond "Approved" and "Dispositioned":

Update the WHERE clauses in **mod_Submit.bas**, function `ArchiveApprovedPIFs`:
```vba
"WHERE status IN ('Approved', 'Dispositioned', 'YourNewStatus')"
```

---

## Security Considerations

### Access Control

- Database permissions should be granted on a per-user basis
- Consider using a dedicated service account for Excel submissions
- Implement row-level security if needed for multi-tenant scenarios

### Sensitive Data

- Justification field may contain confidential information
- Consider encryption at rest if required by policy
- Implement audit logging for all data access

### Backup Strategy

- Database backups should include all tbl_pif_* tables
- Keep at least 3 months of approved data for compliance
- Test restore procedures quarterly

---

## Support and Troubleshooting

### Common Issues and Solutions

1. **"Data mismatch" errors**
   - Usually caused by Excel cell formatting
   - Ensure numeric columns contain only numbers
   - Check for hidden characters in text fields

2. **Performance degradation over time**
   - Archive old approved records to separate tables
   - Rebuild indexes quarterly
   - Consider partitioning by year if dataset exceeds 1M rows

3. **Concurrent submissions**
   - Only one person should submit at a time
   - Use a "submission lock" flag if needed
   - Coordinate via email/Teams during month-end

### Getting Help

- Database issues: Contact your DBA team
- VBA issues: Review error messages in VBA debugger (F8 to step through)
- Business logic questions: Contact PIF process owner

---

## Version History

- **v1.0** (2025-11-05): Initial release
  - Normalized database schema
  - VBA submission automation
  - Excel and SQL validation
  - Automatic archival of approved PIFs

---

## Future Enhancements

Potential improvements for consideration:

1. **Power BI integration** - Real-time dashboards from approved table
2. **Email notifications** - Automatic alerts on submission success/failure
3. **Web interface** - Browser-based submission for remote users
4. **Approval workflow** - Built-in approval routing in database
5. **Change tracking** - Detailed audit trail of who changed what

---

## License and Support

This solution was developed for internal use by the Data Architecture team. For questions, enhancements, or issues, please contact:

- **Database**: [Your DBA Team]
- **Excel/VBA**: [Your Development Team]
- **Business Process**: [PIF Process Owner]

---

**Last Updated**: November 5, 2025
