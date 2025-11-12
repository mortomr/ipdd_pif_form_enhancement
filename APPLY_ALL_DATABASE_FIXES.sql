-- ============================================================================
-- COMPREHENSIVE DATABASE UPDATES - All Pending Fixes
-- ============================================================================
-- Purpose: Apply all pending database schema and stored procedure fixes
-- Date: 2025-11-11
-- Version: 1.0
--
-- This script includes:
--   1. Increase lcm_issue field size from VARCHAR(11) to VARCHAR(20)
--   2. Add @site parameter to usp_archive_approved_pifs
--   3. Fix return value (0 = success instead of @ApprovedCount)
-- ============================================================================

USE [IPDD];
GO

SET NOCOUNT ON;

PRINT '';
PRINT '============================================================================';
PRINT 'COMPREHENSIVE DATABASE UPDATE - Starting...';
PRINT '============================================================================';
PRINT '';

-- ============================================================================
-- SECTION 1: INCREASE LCM_ISSUE FIELD SIZE
-- ============================================================================

PRINT '--- SECTION 1: Updating lcm_issue column size ---';
PRINT '';

PRINT 'Step 1.1: Altering tbl_pif_projects_staging...';
ALTER TABLE dbo.tbl_pif_projects_staging
    ALTER COLUMN lcm_issue VARCHAR(20) NULL;
PRINT '  ✓ tbl_pif_projects_staging.lcm_issue updated to VARCHAR(20)';

PRINT 'Step 1.2: Altering tbl_pif_projects_inflight...';
ALTER TABLE dbo.tbl_pif_projects_inflight
    ALTER COLUMN lcm_issue VARCHAR(20) NULL;
PRINT '  ✓ tbl_pif_projects_inflight.lcm_issue updated to VARCHAR(20)';

PRINT 'Step 1.3: Altering tbl_pif_projects_approved...';
ALTER TABLE dbo.tbl_pif_projects_approved
    ALTER COLUMN lcm_issue VARCHAR(20) NULL;
PRINT '  ✓ tbl_pif_projects_approved.lcm_issue updated to VARCHAR(20)';

PRINT '';
PRINT 'Step 1.4: Recreating usp_insert_project_staging with new parameter size...';

IF OBJECT_ID('dbo.usp_insert_project_staging', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_insert_project_staging;
GO

CREATE PROCEDURE dbo.usp_insert_project_staging
    @pif_id VARCHAR(16),
    @project_id VARCHAR(10),
    @status VARCHAR(58) = NULL,
    @change_type VARCHAR(12) = NULL,
    @accounting_treatment VARCHAR(14) = NULL,
    @category VARCHAR(26) = NULL,
    @seg INT = NULL,
    @opco VARCHAR(4) = NULL,
    @site VARCHAR(4) = NULL,
    @strategic_rank VARCHAR(26) = NULL,
    @funding_project VARCHAR(10) = NULL,
    @project_name VARCHAR(35) = NULL,
    @original_fp_isd VARCHAR(20) = NULL,
    @revised_fp_isd VARCHAR(20) = NULL,
    @moving_isd_year CHAR(1) = NULL,
    @lcm_issue VARCHAR(20) = NULL,  -- CHANGED FROM VARCHAR(11)
    @justification VARCHAR(192) = NULL,
    @prior_year_spend DECIMAL(18,2) = NULL,
    @archive_flag BIT = NULL,
    @include_flag BIT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    BEGIN TRY
        INSERT INTO dbo.tbl_pif_projects_staging
        (
        pif_id, project_id, status, change_type, accounting_treatment,
        category, seg, opco, site, strategic_rank, funding_project,
        project_name, original_fp_isd, revised_fp_isd, moving_isd_year,
        lcm_issue, justification, prior_year_spend, archive_flag, include_flag
        )
    VALUES
        (
            @pif_id, @project_id, @status, @change_type, @accounting_treatment,
            @category, @seg, @opco, @site, @strategic_rank, @funding_project,
            @project_name, @original_fp_isd, @revised_fp_isd, @moving_isd_year,
            @lcm_issue, @justification, @prior_year_spend, @archive_flag, @include_flag
        );

        RETURN 0;

    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        RAISERROR( @ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

PRINT '  ✓ usp_insert_project_staging recreated with lcm_issue VARCHAR(20)';
PRINT '';

-- ============================================================================
-- SECTION 2: UPDATE ARCHIVE STORED PROCEDURE
-- ============================================================================

PRINT '--- SECTION 2: Updating usp_archive_approved_pifs ---';
PRINT '';

PRINT 'Step 2.1: Dropping existing stored procedure...';

IF OBJECT_ID('dbo.usp_archive_approved_pifs', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_archive_approved_pifs;
GO

PRINT '  ✓ Old procedure dropped';
PRINT '';

PRINT 'Step 2.2: Creating new procedure with @site parameter and correct return value...';

CREATE PROCEDURE dbo.usp_archive_approved_pifs
    @site VARCHAR(4)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @ApprovedCount INT = 0;
    DECLARE @CostCount INT = 0;

    BEGIN TRY
        BEGIN TRANSACTION;

        -- Step 1: Archive approved projects for specified site
        INSERT INTO dbo.tbl_pif_projects_approved
        (
        pif_id, project_id, submission_date, approval_date, status,
        change_type, accounting_treatment, category, seg, opco, site,
        strategic_rank, funding_project, project_name, original_fp_isd,
        revised_fp_isd, moving_isd_year, lcm_issue, justification,
        prior_year_spend, archive_flag, include_flag
        )
    SELECT
        p.pif_id,
        p.project_id,
        p.submission_date,
        GETDATE(),
        p.status,
        p.change_type,
        p.accounting_treatment,
        p.category,
        p.seg,
        p.opco,
        p.site,
        p.strategic_rank,
        p.funding_project,
        p.project_name,
        p.original_fp_isd,
        p.revised_fp_isd,
        p.moving_isd_year,
        p.lcm_issue,
        p.justification,
        p.prior_year_spend,
        p.archive_flag,
        p.include_flag
    FROM dbo.tbl_pif_projects_inflight p
    WHERE p.archive_flag = 1 AND p.include_flag = 1 AND p.site = @site;

        SET @ApprovedCount = @@ROWCOUNT;

        -- Step 2: Archive approved costs for specified site
        INSERT INTO dbo.tbl_pif_cost_approved
        (
        pif_id, project_id, scenario, year,
        requested_value, current_value, variance_value, approval_date
        )
    SELECT
        c.pif_id, c.project_id, c.scenario, c.year,
        c.requested_value, c.current_value, c.variance_value, GETDATE()
    FROM dbo.tbl_pif_cost_inflight c
        INNER JOIN dbo.tbl_pif_projects_inflight p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id
    WHERE p.archive_flag = 1 AND p.include_flag = 1 AND p.site = @site;

        SET @CostCount = @@ROWCOUNT;

        -- Step 3: Remove archived records from inflight for specified site
        DELETE c
        FROM dbo.tbl_pif_cost_inflight c
        INNER JOIN dbo.tbl_pif_projects_inflight p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id
        WHERE p.archive_flag = 1 AND p.include_flag = 1 AND p.site = @site;

        DELETE FROM dbo.tbl_pif_projects_inflight
        WHERE archive_flag = 1 AND include_flag = 1 AND site = @site;

        COMMIT TRANSACTION;

        -- Return success message
        SELECT
        'SUCCESS' AS Status,
        @ApprovedCount AS ProjectsArchived,
        @CostCount AS CostRecordsArchived;

        RETURN 0;  -- Success (VBA expects 0 = success, -1 = error)

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        SELECT
        'ERROR' AS Status,
        @ErrorMessage AS ErrorMessage;

        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

PRINT '  ✓ usp_archive_approved_pifs recreated successfully';
PRINT '';

-- ============================================================================
-- VERIFICATION
-- ============================================================================

PRINT '============================================================================';
PRINT 'VERIFICATION - Checking all changes...';
PRINT '============================================================================';
PRINT '';

PRINT '--- Column Definitions (lcm_issue) ---';
SELECT
    t.name AS TableName,
    c.name AS ColumnName,
    ty.name AS DataType,
    c.max_length AS MaxLength,
    CASE WHEN c.is_nullable = 1 THEN 'NULL' ELSE 'NOT NULL' END AS Nullable
FROM sys.tables t
INNER JOIN sys.columns c ON t.object_id = c.object_id
INNER JOIN sys.types ty ON c.user_type_id = ty.user_type_id
WHERE t.name IN ('tbl_pif_projects_staging', 'tbl_pif_projects_inflight', 'tbl_pif_projects_approved')
  AND c.name = 'lcm_issue'
ORDER BY t.name;

PRINT '';
PRINT '--- Stored Procedure Parameters (usp_archive_approved_pifs) ---';
SELECT
    p.name AS ProcedureName,
    pm.name AS ParameterName,
    TYPE_NAME(pm.user_type_id) AS DataType,
    pm.max_length AS MaxLength,
    pm.is_output AS IsOutput
FROM sys.procedures p
LEFT JOIN sys.parameters pm ON p.object_id = pm.object_id
WHERE p.name = 'usp_archive_approved_pifs'
ORDER BY pm.parameter_id;

PRINT '';
PRINT '--- Stored Procedure Parameters (usp_insert_project_staging @lcm_issue) ---';
SELECT
    p.name AS ProcedureName,
    pm.name AS ParameterName,
    TYPE_NAME(pm.user_type_id) AS DataType,
    pm.max_length AS MaxLength
FROM sys.procedures p
INNER JOIN sys.parameters pm ON p.object_id = pm.object_id
WHERE p.name = 'usp_insert_project_staging'
  AND pm.name = '@lcm_issue';

PRINT '';
PRINT '============================================================================';
PRINT 'COMPREHENSIVE UPDATE COMPLETED SUCCESSFULLY!';
PRINT '============================================================================';
PRINT '';
PRINT 'SUMMARY OF CHANGES:';
PRINT '  ✓ lcm_issue column increased from VARCHAR(11) to VARCHAR(20) in 3 tables';
PRINT '  ✓ usp_insert_project_staging updated with new @lcm_issue size';
PRINT '  ✓ usp_archive_approved_pifs now accepts @site VARCHAR(4) parameter';
PRINT '  ✓ usp_archive_approved_pifs now returns 0 (success) instead of count';
PRINT '  ✓ Archival operations are now site-specific';
PRINT '';
PRINT 'FIXES APPLIED:';
PRINT '  ✓ Fix #1: LCM issue values up to 20 characters (e.g., "LCM-123-RBN25")';
PRINT '  ✓ Fix #2: Archival failure - parameter mismatch resolved';
PRINT '  ✓ Fix #3: False "Archival Failed" message - return value corrected';
PRINT '';
PRINT 'NEXT STEPS:';
PRINT '  1. VBA code is already updated and ready';
PRINT '  2. Test "Save Snapshot" function';
PRINT '  3. Test "Finalize Month" function';
PRINT '  4. Verify no error messages appear when operations succeed';
PRINT '  5. Verify Archive and Inflight worksheets can be refreshed';
PRINT '';
PRINT '============================================================================';
GO
