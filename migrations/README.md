# Database Migrations

This folder contains database migration scripts for applying fixes and updates to the PIF reporting database.

## How to Apply Migrations

1. **Backup your database** before running any migration script
2. Connect to your IPDD database using SQL Server Management Studio or Azure Data Studio
3. Open the migration script you need to run
4. Review the script to understand what changes it will make
5. Execute the script
6. Verify the changes using the verification queries provided in the script

## Available Migrations

### HOTFIX_re_archive_duplicate_fix.sql

**Issue**: When re-archiving a previously approved PIF with updated data, the system was creating duplicate records instead of updating the existing record.

**Fix**: Updates the `usp_archive_approved_pifs` stored procedure to use MERGE logic (UPSERT) for projects and DELETE+INSERT for costs.

**When to apply**: If you experience duplicate PIF+Project combinations in the approved table after re-archiving.

**What it does**:
- Drops and recreates `usp_archive_approved_pifs` with proper UPSERT logic
- Includes optional cleanup script for existing duplicates (commented out by default)

**Date**: 2025-11-12

## Migration Best Practices

1. Always backup before applying migrations
2. Test migrations in a non-production environment first
3. Review the script before executing
4. Keep track of which migrations have been applied
5. If a migration includes optional cleanup steps, review carefully before uncommenting

## Troubleshooting

If a migration fails:
1. Check the error message for details
2. Verify you have the necessary permissions (EXECUTE on procedures, ALTER on tables)
3. Ensure you're connected to the correct database
4. Check if the migration has already been applied (some migrations are idempotent, others are not)

For assistance, contact the Data Architecture Team.
