-- ============================================================================
-- PIF DATABASE MIGRATION: Update Stored Procedures for line_item
-- ============================================================================
-- Purpose: Update all stored procedures to handle line_item field
-- Author: Data Architecture Team
-- Date: 2025-11-20
--
-- This script updates:
--   1. usp_insert_project_staging - Add @line_item parameter
--   2. usp_insert_cost_staging - Add @line_item parameter
--   3. usp_validate_staging_data_secure - Update duplicate check
--   4. usp_commit_to_inflight - Include line_item in INSERT/DELETE
--   5. usp_archive_approved_pifs - Include line_item in MERGE/DELETE
-- ============================================================================

USE [IPDD];
GO

PRINT 'Updating stored procedures for line_item support...';
PRINT '';

-- ============================================================================
-- 1. usp_insert_project_staging
-- ============================================================================

PRINT 'Updating usp_insert_project_staging...';

IF OBJECT_ID('dbo.usp_insert_project_staging', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_insert_project_staging;
GO

CREATE PROCEDURE dbo.usp_insert_project_staging
    @pif_id VARCHAR(16),
    @project_id VARCHAR(10),
    @line_item INT = 1,  -- NEW PARAMETER
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
    @lcm_issue VARCHAR(20) = NULL,
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
            pif_id, project_id, line_item, status, change_type, accounting_treatment,
            category, seg, opco, site, strategic_rank, funding_project,
            project_name, original_fp_isd, revised_fp_isd, moving_isd_year,
            lcm_issue, justification, prior_year_spend, archive_flag, include_flag
        )
        VALUES
        (
            @pif_id, @project_id, @line_item, @status, @change_type, @accounting_treatment,
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

        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

PRINT '  ✓ Updated usp_insert_project_staging';

-- ============================================================================
-- 2. usp_insert_cost_staging
-- ============================================================================

PRINT 'Updating usp_insert_cost_staging...';

IF OBJECT_ID('dbo.usp_insert_cost_staging', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_insert_cost_staging;
GO

CREATE PROCEDURE dbo.usp_insert_cost_staging
    @pif_id VARCHAR(16),
    @project_id VARCHAR(10),
    @line_item INT = 1,  -- NEW PARAMETER
    @scenario VARCHAR(12),
    @year DATE,
    @requested_value DECIMAL(18,2) = NULL,
    @current_value DECIMAL(18,2) = NULL,
    @variance_value DECIMAL(18,2) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    BEGIN TRY
        INSERT INTO dbo.tbl_pif_cost_staging
        (
            pif_id, project_id, line_item, scenario, year,
            requested_value, current_value, variance_value
        )
        VALUES
        (
            @pif_id, @project_id, @line_item, @scenario, @year,
            @requested_value, @current_value, @variance_value
        );

        RETURN 0;

    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

PRINT '  ✓ Updated usp_insert_cost_staging';

-- ============================================================================
-- 3. usp_validate_staging_data_secure
-- ============================================================================

PRINT 'Updating usp_validate_staging_data_secure...';

IF OBJECT_ID('dbo.usp_validate_staging_data_secure', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_validate_staging_data_secure;
GO

CREATE PROCEDURE dbo.usp_validate_staging_data_secure
    @ErrorCount INT OUTPUT,
    @WarningCount INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @Errors TABLE (
        error_id INT IDENTITY(1,1),
        error_severity VARCHAR(10),
        error_type VARCHAR(50),
        error_message VARCHAR(500),
        record_identifier VARCHAR(100)
    );

    -- Check 1: Missing required fields (CRITICAL)
    INSERT INTO @Errors (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Missing PIF ID',
        'Missing required field: PIF ID',
        'Row ' + CAST(pif_project_id AS VARCHAR(10))
    FROM dbo.tbl_pif_projects_staging
    WHERE pif_id IS NULL OR LTRIM(RTRIM(pif_id)) = '';

    INSERT INTO @Errors (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Missing Project ID',
        'Missing required field: Project ID',
        'Row ' + CAST(pif_project_id AS VARCHAR(10))
    FROM dbo.tbl_pif_projects_staging
    WHERE project_id IS NULL OR LTRIM(RTRIM(project_id)) = '';

    INSERT INTO @Errors (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Missing Change Type',
        'Missing required field: Change Type',
        'PIF ' + pif_id + ', Project ' + project_id + ', Line ' + CAST(line_item AS VARCHAR(10))
    FROM dbo.tbl_pif_projects_staging
    WHERE pif_id IS NOT NULL
        AND project_id IS NOT NULL
        AND (change_type IS NULL OR LTRIM(RTRIM(change_type)) = '');

    -- Check 2: Data type validation (CRITICAL)
    INSERT INTO @Errors (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Invalid Data Type',
        'SEG must be a valid positive integer',
        'PIF ' + pif_id + ', Project ' + project_id + ', Line ' + CAST(line_item AS VARCHAR(10))
    FROM dbo.tbl_pif_projects_staging
    WHERE seg IS NOT NULL
        AND (seg < 0 OR seg > 99999);

    INSERT INTO @Errors (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Invalid Line Item',
        'Line item must be a positive integer (1, 2, 3, etc.)',
        'PIF ' + pif_id + ', Project ' + project_id + ', Line ' + CAST(line_item AS VARCHAR(10))
    FROM dbo.tbl_pif_projects_staging
    WHERE line_item IS NULL OR line_item < 1;

    -- Check 3: Duplicate detection (UPDATED TO INCLUDE LINE_ITEM)
    INSERT INTO @Errors (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Duplicate',
        'Duplicate PIF+Project+Line combination (appears ' + CAST(cnt AS VARCHAR(10)) + ' times)',
        'PIF ' + pif_id + ', Project ' + project_id + ', Line ' + CAST(line_item AS VARCHAR(10))
    FROM (
        SELECT pif_id, project_id, line_item, COUNT(*) as cnt
        FROM dbo.tbl_pif_projects_staging
        WHERE pif_id IS NOT NULL AND project_id IS NOT NULL
        GROUP BY pif_id, project_id, line_item
        HAVING COUNT(*) > 1
    ) dups;

    -- Check 4: Business rule - Approved PIFs require justification (CRITICAL)
    INSERT INTO @Errors (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Missing Justification',
        'Approved or Dispositioned status requires justification',
        'PIF ' + pif_id + ', Project ' + project_id + ', Line ' + CAST(line_item AS VARCHAR(10))
    FROM dbo.tbl_pif_projects_staging
    WHERE status IN ('Approved', 'Dispositioned')
        AND (justification IS NULL OR LTRIM(RTRIM(justification)) = '');

    -- Check 5: Orphan cost records (UPDATED TO INCLUDE LINE_ITEM)
    INSERT INTO @Errors (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Orphan Cost Record',
        'Cost record exists without matching project record',
        'PIF ' + c.pif_id + ', Project ' + c.project_id + ', Line ' + CAST(c.line_item AS VARCHAR(10))
    FROM dbo.tbl_pif_cost_staging c
    LEFT JOIN dbo.tbl_pif_projects_staging p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id AND c.line_item = p.line_item
    WHERE p.pif_id IS NULL;

    -- Check 6: Invalid scenario values (CRITICAL)
    INSERT INTO @Errors (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Invalid Scenario',
        'Scenario must be ''Target'' or ''Closings'' (found: ''' + ISNULL(scenario, 'NULL') + ''')',
        'PIF ' + pif_id + ', Project ' + project_id + ', Line ' + CAST(line_item AS VARCHAR(10)) + ', Year ' + CAST(YEAR(year) AS VARCHAR(4))
    FROM dbo.tbl_pif_cost_staging
    WHERE scenario NOT IN ('Target', 'Closings');

    -- Return separate counts
    SELECT @ErrorCount = COUNT(*) FROM @Errors WHERE error_severity = 'CRITICAL';
    SELECT @WarningCount = COUNT(*) FROM @Errors WHERE error_severity = 'WARNING';

    -- Return error details
    SELECT
        error_id,
        error_severity,
        error_type,
        error_message,
        record_identifier
    FROM @Errors
    ORDER BY
        CASE error_severity
            WHEN 'CRITICAL' THEN 1
            WHEN 'WARNING' THEN 2
            ELSE 3
        END,
        error_id;

    RETURN 0;
END;
GO

PRINT '  ✓ Updated usp_validate_staging_data_secure';

-- ============================================================================
-- 4. usp_commit_to_inflight
-- ============================================================================

PRINT 'Updating usp_commit_to_inflight...';

IF OBJECT_ID('dbo.usp_commit_to_inflight', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_commit_to_inflight;
GO

CREATE PROCEDURE dbo.usp_commit_to_inflight
    @site VARCHAR(4)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @ProjectCount INT = 0;
    DECLARE @CostCount INT = 0;

    BEGIN TRY
        BEGIN TRANSACTION;

        -- Step 1: Clear inflight tables for selected site only
        DELETE c
        FROM dbo.tbl_pif_cost_inflight c
        INNER JOIN dbo.tbl_pif_projects_inflight p
            ON c.pif_id = p.pif_id AND c.project_id = p.project_id AND c.line_item = p.line_item
        WHERE p.site = @site;

        DELETE FROM dbo.tbl_pif_projects_inflight
        WHERE site = @site;

        -- Step 2: Move validated data from staging to inflight
        INSERT INTO dbo.tbl_pif_projects_inflight
        (
            pif_id, project_id, line_item, submission_date, status, change_type,
            accounting_treatment, category, seg, opco, site, strategic_rank,
            funding_project, project_name, original_fp_isd, revised_fp_isd,
            moving_isd_year, lcm_issue, justification, prior_year_spend,
            archive_flag, include_flag
        )
        SELECT
            s.pif_id,
            s.project_id,
            s.line_item,
            GETDATE(),
            s.status,
            s.change_type,
            s.accounting_treatment,
            s.category,
            s.seg,
            s.opco,
            s.site,
            s.strategic_rank,
            s.funding_project,
            s.project_name,
            s.original_fp_isd,
            s.revised_fp_isd,
            s.moving_isd_year,
            s.lcm_issue,
            s.justification,
            s.prior_year_spend,
            s.archive_flag,
            s.include_flag
        FROM dbo.tbl_pif_projects_staging s
        WHERE s.site = @site;

        SET @ProjectCount = @@ROWCOUNT;

        INSERT INTO dbo.tbl_pif_cost_inflight
        (
            pif_id, project_id, line_item, scenario, year,
            requested_value, current_value, variance_value
        )
        SELECT
            c.pif_id, c.project_id, c.line_item, c.scenario, c.year,
            c.requested_value, c.current_value, c.variance_value
        FROM dbo.tbl_pif_cost_staging c
        INNER JOIN dbo.tbl_pif_projects_staging p
            ON c.pif_id = p.pif_id AND c.project_id = p.project_id AND c.line_item = p.line_item
        WHERE p.site = @site;

        SET @CostCount = @@ROWCOUNT;

        COMMIT TRANSACTION;

        SELECT
            'SUCCESS' AS Status,
            @ProjectCount AS ProjectsCommitted,
            @CostCount AS CostRecordsCommitted;

        RETURN 0;

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        SELECT
            'ERROR' AS Status,
            @ErrorMessage AS ErrorMessage,
            @ErrorSeverity AS ErrorSeverity,
            @ErrorState AS ErrorState;

        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

PRINT '  ✓ Updated usp_commit_to_inflight';

-- ============================================================================
-- 5. usp_archive_approved_pifs
-- ============================================================================

PRINT 'Updating usp_archive_approved_pifs...';

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

        -- Step 1: UPSERT approved projects using MERGE (UPDATED TO INCLUDE LINE_ITEM)
        MERGE dbo.tbl_pif_projects_approved AS target
        USING (
            SELECT
                p.pif_id,
                p.project_id,
                p.line_item,
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
        ON target.pif_id = source.pif_id
            AND target.project_id = source.project_id
            AND target.line_item = source.line_item
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
                pif_id, project_id, line_item, submission_date, approval_date, status,
                change_type, accounting_treatment, category, seg, opco, site,
                strategic_rank, funding_project, project_name, original_fp_isd,
                revised_fp_isd, moving_isd_year, lcm_issue, justification,
                prior_year_spend, archive_flag, include_flag
            )
            VALUES (
                source.pif_id,
                source.project_id,
                source.line_item,
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

        -- Step 2: Delete and re-insert costs (UPDATED TO INCLUDE LINE_ITEM)
        DELETE c
        FROM dbo.tbl_pif_cost_approved c
        WHERE EXISTS (
            SELECT 1
            FROM dbo.tbl_pif_projects_inflight p
            WHERE p.pif_id = c.pif_id
                AND p.project_id = c.project_id
                AND p.line_item = c.line_item
                AND p.archive_flag = 1
                AND p.include_flag = 1
                AND p.site = @site
        );

        INSERT INTO dbo.tbl_pif_cost_approved
        (
            pif_id, project_id, line_item, scenario, year,
            requested_value, current_value, variance_value, approval_date
        )
        SELECT
            c.pif_id, c.project_id, c.line_item, c.scenario, c.year,
            c.requested_value, c.current_value, c.variance_value, GETDATE()
        FROM dbo.tbl_pif_cost_inflight c
        INNER JOIN dbo.tbl_pif_projects_inflight p
            ON c.pif_id = p.pif_id AND c.project_id = p.project_id AND c.line_item = p.line_item
        WHERE p.archive_flag = 1 AND p.include_flag = 1 AND p.site = @site;

        SET @CostCount = @@ROWCOUNT;

        -- Step 3: Remove archived records from inflight (UPDATED TO INCLUDE LINE_ITEM)
        DELETE c
        FROM dbo.tbl_pif_cost_inflight c
        INNER JOIN dbo.tbl_pif_projects_inflight p
            ON c.pif_id = p.pif_id AND c.project_id = p.project_id AND c.line_item = p.line_item
        WHERE p.archive_flag = 1 AND p.include_flag = 1 AND p.site = @site;

        DELETE FROM dbo.tbl_pif_projects_inflight
        WHERE archive_flag = 1 AND include_flag = 1 AND site = @site;

        COMMIT TRANSACTION;

        SELECT
            'SUCCESS' AS Status,
            @ApprovedCount AS ProjectsArchived,
            @CostCount AS CostRecordsArchived;

        RETURN 0;

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

PRINT '  ✓ Updated usp_archive_approved_pifs';

-- ============================================================================
-- 6. usp_create_staging_indexes
-- ============================================================================

PRINT 'Updating usp_create_staging_indexes...';

IF OBJECT_ID('dbo.usp_create_staging_indexes', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_create_staging_indexes;
GO

CREATE PROCEDURE dbo.usp_create_staging_indexes
AS
BEGIN
    SET NOCOUNT ON;

    -- Drop existing indexes if they exist
    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_staging_pif_proj')
        DROP INDEX IX_staging_pif_proj ON dbo.tbl_pif_projects_staging;

    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_staging_status')
        DROP INDEX IX_staging_status ON dbo.tbl_pif_projects_staging;

    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_staging_cost_lookup')
        DROP INDEX IX_staging_cost_lookup ON dbo.tbl_pif_cost_staging;

    -- Create indexes for validation queries (UPDATED TO INCLUDE LINE_ITEM)
    CREATE NONCLUSTERED INDEX IX_staging_pif_proj
        ON dbo.tbl_pif_projects_staging (pif_id, project_id, line_item);

    CREATE NONCLUSTERED INDEX IX_staging_status
        ON dbo.tbl_pif_projects_staging (status)
        INCLUDE (pif_id, project_id, line_item, justification);

    CREATE NONCLUSTERED INDEX IX_staging_cost_lookup
        ON dbo.tbl_pif_cost_staging (pif_id, project_id, line_item);

    RETURN 0;
END;
GO

PRINT '  ✓ Updated usp_create_staging_indexes';

PRINT '';
PRINT '============================================================================';
PRINT 'All stored procedures updated successfully!';
PRINT '';
PRINT 'NEXT STEPS:';
PRINT '  1. Update VBA code (mod_Validation.bas, mod_Submit.bas)';
PRINT '  2. Add line_item column to Excel PIF worksheet';
PRINT '  3. Update views to include line_item (optional - for reporting)';
PRINT '============================================================================';
GO
