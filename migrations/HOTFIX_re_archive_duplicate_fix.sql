-- ============================================================================
-- HOTFIX: Fix Re-Archive Duplicate Bug
-- ============================================================================
-- Issue: When re-archiving a previously approved PIF with updated data,
--        the system was creating duplicate records instead of updating
--        the existing record in the approved table.
--
-- Root Cause: usp_archive_approved_pifs was using simple INSERT logic
--             instead of UPSERT (UPDATE if exists, INSERT if new)
--
-- Fix: Replace INSERT with MERGE statement for projects table
--      Use DELETE+INSERT pattern for costs table
--
-- Business Rule: PIF ID + Funding Project should be unique in approved table
--
-- Run this script on your IPDD database to apply the fix
-- ============================================================================

USE [IPDD];
GO

-- Drop and recreate the stored procedure with UPSERT logic
IF OBJECT_ID('dbo.usp_archive_approved_pifs', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_archive_approved_pifs;
GO

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

        -- Step 1: UPSERT approved projects for specified site using MERGE
        -- If PIF+Project exists, UPDATE it. If not, INSERT it.
        MERGE dbo.tbl_pif_projects_approved AS target
        USING (
            SELECT
                p.pif_id,
                p.project_id,
                p.submission_date,
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
            WHERE p.archive_flag = 1 AND p.include_flag = 1 AND p.site = @site
        ) AS source
        ON target.pif_id = source.pif_id AND target.project_id = source.project_id
        WHEN MATCHED THEN
            UPDATE SET
                submission_date = source.submission_date,
                approval_date = GETDATE(),
                status = source.status,
                change_type = source.change_type,
                accounting_treatment = source.accounting_treatment,
                category = source.category,
                seg = source.seg,
                opco = source.opco,
                site = source.site,
                strategic_rank = source.strategic_rank,
                funding_project = source.funding_project,
                project_name = source.project_name,
                original_fp_isd = source.original_fp_isd,
                revised_fp_isd = source.revised_fp_isd,
                moving_isd_year = source.moving_isd_year,
                lcm_issue = source.lcm_issue,
                justification = source.justification,
                prior_year_spend = source.prior_year_spend,
                archive_flag = source.archive_flag,
                include_flag = source.include_flag
        WHEN NOT MATCHED THEN
            INSERT (
                pif_id, project_id, submission_date, approval_date, status,
                change_type, accounting_treatment, category, seg, opco, site,
                strategic_rank, funding_project, project_name, original_fp_isd,
                revised_fp_isd, moving_isd_year, lcm_issue, justification,
                prior_year_spend, archive_flag, include_flag
            )
            VALUES (
                source.pif_id,
                source.project_id,
                source.submission_date,
                GETDATE(),
                source.status,
                source.change_type,
                source.accounting_treatment,
                source.category,
                source.seg,
                source.opco,
                source.site,
                source.strategic_rank,
                source.funding_project,
                source.project_name,
                source.original_fp_isd,
                source.revised_fp_isd,
                source.moving_isd_year,
                source.lcm_issue,
                source.justification,
                source.prior_year_spend,
                source.archive_flag,
                source.include_flag
            );

        SET @ApprovedCount = @@ROWCOUNT;

        -- Step 2: Delete and re-insert costs (simpler than MERGE for child records)
        -- First, delete existing cost records for PIFs being archived
        DELETE c
        FROM dbo.tbl_pif_cost_approved c
        WHERE EXISTS (
            SELECT 1
            FROM dbo.tbl_pif_projects_inflight p
            WHERE p.pif_id = c.pif_id
                AND p.project_id = c.project_id
                AND p.archive_flag = 1
                AND p.include_flag = 1
                AND p.site = @site
        );

        -- Then, insert new/updated cost records
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

PRINT '';
PRINT '============================================================================';
PRINT 'HOTFIX APPLIED SUCCESSFULLY';
PRINT '';
PRINT 'The usp_archive_approved_pifs stored procedure has been updated.';
PRINT '';
PRINT 'What changed:';
PRINT '  - Projects: Now uses MERGE (UPSERT) instead of INSERT';
PRINT '  - Costs: Now uses DELETE+INSERT pattern to replace existing records';
PRINT '';
PRINT 'Result:';
PRINT '  - Re-archiving a previously approved PIF will UPDATE the existing record';
PRINT '  - No more duplicate PIF+Project combinations in approved table';
PRINT '============================================================================';
PRINT '';
GO

-- ============================================================================
-- OPTIONAL: Clean up existing duplicates in approved table
-- ============================================================================
-- Uncomment this section if you have existing duplicates to clean up
-- This keeps the LATEST version of each PIF+Project combination

/*
PRINT 'Checking for existing duplicates...';
GO

WITH duplicate_check AS (
    SELECT
        pif_id,
        project_id,
        COUNT(*) as duplicate_count
    FROM dbo.tbl_pif_projects_approved
    GROUP BY pif_id, project_id
    HAVING COUNT(*) > 1
)
SELECT
    'WARNING: Found ' + CAST(COUNT(*) AS VARCHAR(10)) + ' PIF+Project combinations with duplicates' AS message
FROM duplicate_check
HAVING COUNT(*) > 0;
GO

PRINT '';
PRINT 'To remove duplicates (keeping the latest record for each PIF+Project):';
PRINT '1. Uncomment the DELETE statement below';
PRINT '2. Review the records that will be deleted';
PRINT '3. Execute the DELETE statement';
PRINT '';

-- Uncomment to delete duplicate records (keeps latest by pif_project_id)
/*
BEGIN TRANSACTION;

-- Delete costs for duplicate project records (keeping latest)
DELETE c
FROM dbo.tbl_pif_cost_approved c
WHERE EXISTS (
    SELECT 1
    FROM (
        SELECT
            pif_id,
            project_id,
            pif_project_id,
            ROW_NUMBER() OVER (PARTITION BY pif_id, project_id ORDER BY pif_project_id DESC) as rn
        FROM dbo.tbl_pif_projects_approved
    ) dups
    WHERE dups.pif_id = c.pif_id
        AND dups.project_id = c.project_id
        AND dups.rn > 1
);

-- Delete duplicate project records (keeping latest)
DELETE p
FROM dbo.tbl_pif_projects_approved p
INNER JOIN (
    SELECT
        pif_project_id,
        ROW_NUMBER() OVER (PARTITION BY pif_id, project_id ORDER BY pif_project_id DESC) as rn
    FROM dbo.tbl_pif_projects_approved
) dups ON p.pif_project_id = dups.pif_project_id
WHERE dups.rn > 1;

SELECT 'Deleted ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' duplicate project records' AS result;

COMMIT TRANSACTION;

PRINT 'Duplicates removed successfully';
*/
*/
GO
