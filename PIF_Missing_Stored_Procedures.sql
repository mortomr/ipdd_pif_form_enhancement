-- ============================================================================
-- PIF MISSING STORED PROCEDURES
-- ============================================================================
-- Purpose: Add missing insert stored procedures required by VBA bulk upload
-- Author: Data Architecture Team
-- Date: 2025-11-06
-- Version: 2.0.1
--
-- IMPORTANT: Execute this script AFTER running PIF_Database_DDL.sql
--
-- These procedures were referenced in mod_Database.bas but were missing
-- from the original DDL script. They provide secure, parameterized
-- bulk insert operations for the VBA application.
-- ============================================================================

USE [IPDD];
GO

-- ----------------------------------------------------------------------------
-- Procedure: usp_insert_project_staging
-- Purpose: Insert a single project record into staging table
-- SECURITY: Parameterized to prevent SQL injection
-- ----------------------------------------------------------------------------

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
    @lcm_issue VARCHAR(11) = NULL,
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
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

-- ----------------------------------------------------------------------------
-- Procedure: usp_insert_cost_staging
-- Purpose: Insert a single cost record into staging table
-- SECURITY: Parameterized to prevent SQL injection
-- ----------------------------------------------------------------------------

IF OBJECT_ID('dbo.usp_insert_cost_staging', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_insert_cost_staging;
GO

CREATE PROCEDURE dbo.usp_insert_cost_staging
    @pif_id VARCHAR(16),
    @project_id VARCHAR(10),
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
        -- Insert into cost staging table
        INSERT INTO dbo.tbl_pif_cost_staging
        (
            pif_id, project_id, scenario, year,
            requested_value, current_value, variance_value
        )
        VALUES
        (
            @pif_id, @project_id, @scenario, @year,
            @requested_value, @current_value, @variance_value
        );

        RETURN 0;

    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        -- Return error information
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

-- ============================================================================
-- GRANT PERMISSIONS
-- ============================================================================

-- IMPORTANT: Replace [DOMAIN\VBAUser] with the actual Windows account
-- that runs the VBA application. Example: [CONTOSO\john.smith]

-- Uncomment and execute these statements after replacing the placeholder:

/*
GRANT EXECUTE ON dbo.usp_insert_project_staging TO [DOMAIN\VBAUser];
GRANT EXECUTE ON dbo.usp_insert_cost_staging TO [DOMAIN\VBAUser];
*/

PRINT '';
PRINT '============================================================================';
PRINT 'Missing stored procedures created successfully!';
PRINT '';
PRINT 'NEXT STEPS:';
PRINT '  1. Uncomment the GRANT statements above';
PRINT '  2. Replace [DOMAIN\VBAUser] with your actual Windows account';
PRINT '  3. Execute the GRANT statements to allow VBA to call these procedures';
PRINT '';
PRINT 'VERIFICATION:';
PRINT '  Run this query to verify the procedures exist:';
PRINT '    SELECT name FROM sys.procedures WHERE name LIKE ''usp_insert%''';
PRINT '============================================================================';
GO
