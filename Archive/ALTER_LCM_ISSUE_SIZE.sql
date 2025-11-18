-- ============================================================================
-- FIX: Increase lcm_issue column size from VARCHAR(11) to VARCHAR(20)
-- ============================================================================
-- Purpose: Fix "Application uses a value of the wrong type" error when
--          lcm_issue values exceed 11 characters (e.g., "LCM-123-RBN25")
-- Date: 2025-11-11
-- ============================================================================

USE [IPDD];
GO

PRINT 'Altering lcm_issue column size from VARCHAR(11) to VARCHAR(20)...';
PRINT '';

-- ============================================================================
-- STEP 1: ALTER STAGING TABLE
-- ============================================================================
PRINT 'Step 1: Altering tbl_pif_projects_staging...';

ALTER TABLE dbo.tbl_pif_projects_staging
    ALTER COLUMN lcm_issue VARCHAR(20) NULL;

PRINT '  ✓ tbl_pif_projects_staging altered successfully';
PRINT '';

-- ============================================================================
-- STEP 2: ALTER INFLIGHT TABLE
-- ============================================================================
PRINT 'Step 2: Altering tbl_pif_projects_inflight...';

ALTER TABLE dbo.tbl_pif_projects_inflight
    ALTER COLUMN lcm_issue VARCHAR(20) NULL;

PRINT '  ✓ tbl_pif_projects_inflight altered successfully';
PRINT '';

-- ============================================================================
-- STEP 3: ALTER APPROVED TABLE
-- ============================================================================
PRINT 'Step 3: Altering tbl_pif_projects_approved...';

ALTER TABLE dbo.tbl_pif_projects_approved
    ALTER COLUMN lcm_issue VARCHAR(20) NULL;

PRINT '  ✓ tbl_pif_projects_approved altered successfully';
PRINT '';

-- ============================================================================
-- STEP 4: RECREATE STORED PROCEDURE WITH NEW PARAMETER SIZE
-- ============================================================================
PRINT 'Step 4: Recreating stored procedure usp_insert_project_staging...';

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
        -- Insert into staging table
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

        -- Return error information
        RAISERROR( @ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

PRINT '  ✓ usp_insert_project_staging recreated successfully';
PRINT '';

-- ============================================================================
-- VERIFICATION
-- ============================================================================
PRINT '============================================================================';
PRINT 'VERIFICATION - Check column definitions:';
PRINT '============================================================================';

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
PRINT '============================================================================';
PRINT 'ALTER SCRIPT COMPLETED SUCCESSFULLY!';
PRINT '';
PRINT 'NEXT STEPS:';
PRINT '  1. The database schema has been updated';
PRINT '  2. The VBA code has already been updated in your files';
PRINT '  3. You can now submit PIFs with LCM issues up to 20 characters';
PRINT '============================================================================';
GO
