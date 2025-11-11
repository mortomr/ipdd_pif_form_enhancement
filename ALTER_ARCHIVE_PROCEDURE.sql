-- ============================================================================
-- FIX: Add @site parameter to usp_archive_approved_pifs
-- ============================================================================
-- Purpose: Fix "Archival Failed" error during month-end finalization
--          The VBA code was calling the procedure with a @site parameter,
--          but the procedure didn't accept it, causing a parameter mismatch.
-- Date: 2025-11-11
-- ============================================================================

USE [IPDD];
GO

PRINT 'Updating usp_archive_approved_pifs to accept @site parameter...';
PRINT '';

-- ============================================================================
-- STEP 1: DROP EXISTING PROCEDURE
-- ============================================================================
PRINT 'Step 1: Dropping existing stored procedure...';

IF OBJECT_ID('dbo.usp_archive_approved_pifs', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_archive_approved_pifs;
GO

PRINT '  ✓ Old procedure dropped successfully';
PRINT '';

-- ============================================================================
-- STEP 2: CREATE NEW PROCEDURE WITH @site PARAMETER
-- ============================================================================
PRINT 'Step 2: Creating new stored procedure with @site parameter...';

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
        GETDATE(), -- This is the value for approval_date
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

        RETURN @ApprovedCount;

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

PRINT '  ✓ New procedure created successfully';
PRINT '';

-- ============================================================================
-- VERIFICATION
-- ============================================================================
PRINT '============================================================================';
PRINT 'VERIFICATION - Check stored procedure parameters:';
PRINT '============================================================================';

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
PRINT '============================================================================';
PRINT 'ALTER SCRIPT COMPLETED SUCCESSFULLY!';
PRINT '';
PRINT 'CHANGES MADE:';
PRINT '  - Added @site VARCHAR(4) parameter to usp_archive_approved_pifs';
PRINT '  - Modified INSERT queries to filter by site (WHERE ... AND p.site = @site)';
PRINT '  - Modified DELETE queries to filter by site (WHERE ... AND site = @site)';
PRINT '';
PRINT 'IMPACT:';
PRINT '  - Archival now operates on a per-site basis';
PRINT '  - Only approved records for the specified site are archived';
PRINT '  - Prevents cross-site data mixing';
PRINT '';
PRINT 'NEXT STEPS:';
PRINT '  1. The VBA code already expects this parameter';
PRINT '  2. Test the Finalize Month function with your selected site';
PRINT '  3. Verify that only your site records are archived';
PRINT '============================================================================';
GO
