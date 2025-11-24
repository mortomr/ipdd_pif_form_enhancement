-- ============================================================================
-- VERIFY PROD DATABASE SETUP
-- ============================================================================
-- Purpose: Verify that all objects exist with correct schema on PROD
-- Run this on PROD to diagnose the staging upload failure
-- ============================================================================

USE [IPDD];
GO

PRINT '============================================================================';
PRINT 'PROD DATABASE VERIFICATION SCRIPT';
PRINT '============================================================================';
PRINT '';

-- ============================================================================
-- CHECK 1: Verify stored procedure exists and has correct parameters
-- ============================================================================
PRINT 'CHECK 1: Verifying usp_insert_project_staging parameters...';
PRINT '';

IF OBJECT_ID('dbo.usp_insert_project_staging', 'P') IS NULL
BEGIN
    PRINT '❌ CRITICAL: usp_insert_project_staging does NOT exist!';
    PRINT '   ACTION: Re-run PIF_Database_DDL.sql';
END
ELSE
BEGIN
    PRINT '✓ usp_insert_project_staging exists';
    PRINT '';
    PRINT 'Parameter List:';

    SELECT
        parameter_id AS [Position],
        name AS [Parameter Name],
        TYPE_NAME(user_type_id) AS [Data Type],
        max_length AS [Max Length],
        CASE
            WHEN is_output = 1 THEN 'OUTPUT'
            ELSE 'INPUT'
        END AS [Direction]
    FROM sys.parameters
    WHERE object_id = OBJECT_ID('dbo.usp_insert_project_staging')
    ORDER BY parameter_id;

    -- Check specifically for @line_item parameter
    IF EXISTS (
        SELECT 1
        FROM sys.parameters
        WHERE object_id = OBJECT_ID('dbo.usp_insert_project_staging')
        AND name = '@line_item'
    )
    BEGIN
        PRINT '';
        PRINT '✓ @line_item parameter found at position:';
        SELECT parameter_id
        FROM sys.parameters
        WHERE object_id = OBJECT_ID('dbo.usp_insert_project_staging')
        AND name = '@line_item';
    END
    ELSE
    BEGIN
        PRINT '';
        PRINT '❌ CRITICAL: @line_item parameter is MISSING!';
        PRINT '   This is the root cause of your upload failure!';
        PRINT '   ACTION: Re-run PIF_Database_DDL.sql to recreate stored procedures';
    END
END

PRINT '';
PRINT '============================================================================';
PRINT 'CHECK 2: Verifying usp_insert_cost_staging parameters...';
PRINT '';

IF OBJECT_ID('dbo.usp_insert_cost_staging', 'P') IS NULL
BEGIN
    PRINT '❌ CRITICAL: usp_insert_cost_staging does NOT exist!';
END
ELSE
BEGIN
    PRINT '✓ usp_insert_cost_staging exists';
    PRINT '';

    SELECT
        parameter_id AS [Position],
        name AS [Parameter Name],
        TYPE_NAME(user_type_id) AS [Data Type]
    FROM sys.parameters
    WHERE object_id = OBJECT_ID('dbo.usp_insert_cost_staging')
    ORDER BY parameter_id;

    IF EXISTS (
        SELECT 1
        FROM sys.parameters
        WHERE object_id = OBJECT_ID('dbo.usp_insert_cost_staging')
        AND name = '@line_item'
    )
    BEGIN
        PRINT '';
        PRINT '✓ @line_item parameter found';
    END
    ELSE
    BEGIN
        PRINT '';
        PRINT '❌ CRITICAL: @line_item parameter is MISSING!';
    END
END

PRINT '';
PRINT '============================================================================';
PRINT 'CHECK 3: Verifying table schemas have line_item column...';
PRINT '';

-- Check staging tables
IF COL_LENGTH('dbo.tbl_pif_projects_staging', 'line_item') IS NULL
BEGIN
    PRINT '❌ CRITICAL: tbl_pif_projects_staging is missing line_item column!';
END
ELSE
BEGIN
    PRINT '✓ tbl_pif_projects_staging has line_item column';
END

IF COL_LENGTH('dbo.tbl_pif_cost_staging', 'line_item') IS NULL
BEGIN
    PRINT '❌ CRITICAL: tbl_pif_cost_staging is missing line_item column!';
END
ELSE
BEGIN
    PRINT '✓ tbl_pif_cost_staging has line_item column';
END

-- Check inflight tables
IF COL_LENGTH('dbo.tbl_pif_projects_inflight', 'line_item') IS NULL
BEGIN
    PRINT '❌ CRITICAL: tbl_pif_projects_inflight is missing line_item column!';
END
ELSE
BEGIN
    PRINT '✓ tbl_pif_projects_inflight has line_item column';
END

IF COL_LENGTH('dbo.tbl_pif_cost_inflight', 'line_item') IS NULL
BEGIN
    PRINT '❌ CRITICAL: tbl_pif_cost_inflight is missing line_item column!';
END
ELSE
BEGIN
    PRINT '✓ tbl_pif_cost_inflight has line_item column';
END

PRINT '';
PRINT '============================================================================';
PRINT 'CHECK 4: Testing stored procedure execution...';
PRINT '';

BEGIN TRY
    -- Try to execute the stored procedure with a test record
    EXEC dbo.usp_insert_project_staging
        @pif_id = 'TEST123',
        @project_id = 'TESTPROJ',
        @line_item = 1,
        @status = 'Test',
        @change_type = 'New',
        @accounting_treatment = 'Cap',
        @category = 'Test Category',
        @seg = 1,
        @opco = 'TEST',
        @site = 'TEST',
        @strategic_rank = NULL,
        @funding_project = 'TESTPROJ',
        @project_name = 'Test Project',
        @original_fp_isd = NULL,
        @revised_fp_isd = NULL,
        @moving_isd_year = NULL,
        @lcm_issue = NULL,
        @justification = 'Test',
        @prior_year_spend = NULL,
        @archive_flag = 0,
        @include_flag = 1;

    PRINT '✓ Stored procedure executed successfully!';
    PRINT '';

    -- Clean up test data
    DELETE FROM dbo.tbl_pif_projects_staging WHERE pif_id = 'TEST123';
    PRINT '✓ Test record cleaned up';

END TRY
BEGIN CATCH
    PRINT '❌ CRITICAL: Stored procedure execution FAILED!';
    PRINT '';
    PRINT 'Error Details:';
    PRINT '  Error Number: ' + CAST(ERROR_NUMBER() AS VARCHAR(10));
    PRINT '  Error Message: ' + ERROR_MESSAGE();
    PRINT '  Error Line: ' + CAST(ERROR_LINE() AS VARCHAR(10));
    PRINT '';

    IF ERROR_NUMBER() = 8144
    BEGIN
        PRINT 'ERROR 8144 = Procedure has too many arguments specified';
        PRINT 'This means the stored procedure does NOT have the @line_item parameter!';
        PRINT 'ACTION: Re-run PIF_Database_DDL.sql';
    END
    ELSE IF ERROR_NUMBER() = 201
    BEGIN
        PRINT 'ERROR 201 = Parameter mismatch';
        PRINT 'This means the stored procedure parameters do not match VBA code!';
        PRINT 'ACTION: Re-run PIF_Database_DDL.sql';
    END
END CATCH

PRINT '';
PRINT '============================================================================';
PRINT 'CHECK 5: Verify database connection from VBA perspective...';
PRINT '';

PRINT 'Connection String Details:';
PRINT '  Server: ' + @@SERVERNAME;
PRINT '  Database: ' + DB_NAME();
PRINT '  Current User: ' + SYSTEM_USER;
PRINT '';

-- Check if user has EXECUTE permission on stored procedures
IF HAS_PERMS_BY_NAME('dbo.usp_insert_project_staging', 'OBJECT', 'EXECUTE') = 1
BEGIN
    PRINT '✓ Current user has EXECUTE permission on usp_insert_project_staging';
END
ELSE
BEGIN
    PRINT '❌ WARNING: Current user does NOT have EXECUTE permission!';
    PRINT '   This could cause "Access Denied" errors from VBA';
    PRINT '   ACTION: Grant EXECUTE permission to your Windows account';
END

PRINT '';
PRINT '============================================================================';
PRINT 'VERIFICATION COMPLETE';
PRINT '============================================================================';
PRINT '';
PRINT 'Next Steps:';
PRINT '  1. Review the checks above for any ❌ CRITICAL errors';
PRINT '  2. If @line_item is missing, re-run PIF_Database_DDL.sql';
PRINT '  3. If permissions are missing, run GRANT statements from DDL';
PRINT '  4. After fixing, re-test VBA upload';
PRINT '';
GO
