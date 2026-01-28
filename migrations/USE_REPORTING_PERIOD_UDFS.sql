-- ============================================================================
-- MIGRATION: Replace Calendar-Based Dates with Reporting Period UDFs
-- ============================================================================
-- Date: 2025-11-19
-- Purpose: Update views to use reporting period UDFs (dbo.CY, dbo.CM) instead
--          of calendar-based date functions (YEAR(GETDATE()))
--
-- Background:
--   The database uses reporting periods that lag the calendar. Reporting periods
--   are tracked based on when actual costs are loaded into the system, which
--   differs from calendar dates. The UDFs dbo.CY() and dbo.CM() return the
--   current reporting year and month based on this special timestamp.
--
-- Prerequisites:
--   - UDFs dbo.CY() and dbo.CM() must already exist in the database
--   - These UDFs should return INT values for reporting period year and month
--
-- Impact:
--   - vw_TA_Inflight_wide: Updated to use dbo.CY()
--   - vw_pif_approved_wide: Updated to use dbo.CY()
--
-- ============================================================================

USE [IPDD];
GO

PRINT 'Starting migration: USE_REPORTING_PERIOD_UDFS';
PRINT '============================================================================';

-- ============================================================================
-- STEP 1: Verify UDFs exist
-- ============================================================================

IF OBJECT_ID('dbo.CY', 'FN') IS NULL
BEGIN
    PRINT 'ERROR: dbo.CY() function does not exist!';
    PRINT 'Please create the reporting period UDFs before running this migration.';
    RAISERROR('Missing prerequisite: dbo.CY() function', 16, 1);
    RETURN;
END

IF OBJECT_ID('dbo.CM', 'FN') IS NULL
BEGIN
    PRINT 'ERROR: dbo.CM() function does not exist!';
    PRINT 'Please create the reporting period UDFs before running this migration.';
    RAISERROR('Missing prerequisite: dbo.CM() function', 16, 1);
    RETURN;
END

PRINT 'Prerequisite check passed: UDFs exist';
PRINT '';

-- ============================================================================
-- STEP 2: Drop and recreate vw_TA_Inflight_wide
-- ============================================================================

PRINT 'Updating vw_TA_Inflight_wide...';

IF OBJECT_ID('dbo.vw_TA_Inflight_wide', 'V') IS NOT NULL
    DROP VIEW dbo.vw_TA_Inflight_wide;
GO

CREATE VIEW dbo.vw_TA_Inflight_wide
AS
    WITH
        cost_pivot
        AS
        (
            SELECT
                pif_id,
                project_id,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() THEN requested_value END) AS Target_Req_CY,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 1 THEN requested_value END) AS Target_Req_CY1,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 2 THEN requested_value END) AS Target_Req_CY2,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 3 THEN requested_value END) AS Target_Req_CY3,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 4 THEN requested_value END) AS Target_Req_CY4,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 5 THEN requested_value END) AS Target_Req_CY5,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() THEN current_value END) AS Target_Curr_CY,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 1 THEN current_value END) AS Target_Curr_CY1,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 2 THEN current_value END) AS Target_Curr_CY2,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 3 THEN current_value END) AS Target_Curr_CY3,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 4 THEN current_value END) AS Target_Curr_CY4,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 5 THEN current_value END) AS Target_Curr_CY5,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() THEN variance_value END) AS Target_Var_CY,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 1 THEN variance_value END) AS Target_Var_CY1,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 2 THEN variance_value END) AS Target_Var_CY2,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 3 THEN variance_value END) AS Target_Var_CY3,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 4 THEN variance_value END) AS Target_Var_CY4,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 5 THEN variance_value END) AS Target_Var_CY5,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() THEN requested_value END) AS Closings_Req_CY,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 1 THEN requested_value END) AS Closings_Req_CY1,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 2 THEN requested_value END) AS Closings_Req_CY2,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 3 THEN requested_value END) AS Closings_Req_CY3,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 4 THEN requested_value END) AS Closings_Req_CY4,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 5 THEN requested_value END) AS Closings_Req_CY5,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() THEN current_value END) AS Closings_Curr_CY,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 1 THEN current_value END) AS Closings_Curr_CY1,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 2 THEN current_value END) AS Closings_Curr_CY2,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 3 THEN current_value END) AS Closings_Curr_CY3,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 4 THEN current_value END) AS Closings_Curr_CY4,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 5 THEN current_value END) AS Closings_Curr_CY5,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() THEN variance_value END) AS Closings_Var_CY,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 1 THEN variance_value END) AS Closings_Var_CY1,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 2 THEN variance_value END) AS Closings_Var_CY2,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 3 THEN variance_value END) AS Closings_Var_CY3,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 4 THEN variance_value END) AS Closings_Var_CY4,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 5 THEN variance_value END) AS Closings_Var_CY5
            FROM dbo.tbl_pif_cost_inflight
            GROUP BY pif_id, project_id
        )
    SELECT
        p.archive_flag,
        p.include_flag,
        p.accounting_treatment,
        p.change_type,
        p.pif_id,
        p.seg,
        p.opco,
        p.site,
        p.strategic_rank,
        p.funding_project,
        p.project_name,
        p.original_fp_isd,
        p.revised_fp_isd,
        p.lcm_issue,
        p.status,
        p.category,
        p.justification,
        c.Target_Req_CY,
        c.Target_Req_CY1,
        c.Target_Req_CY2,
        c.Target_Req_CY3,
        c.Target_Req_CY4,
        c.Target_Req_CY5,
        c.Target_Curr_CY,
        c.Target_Curr_CY1,
        c.Target_Curr_CY2,
        c.Target_Curr_CY3,
        c.Target_Curr_CY4,
        c.Target_Curr_CY5,
        c.Target_Var_CY,
        c.Target_Var_CY1,
        c.Target_Var_CY2,
        c.Target_Var_CY3,
        c.Target_Var_CY4,
        c.Target_Var_CY5,
        p.moving_isd_year,
        p.prior_year_spend,
        c.Closings_Req_CY,
        c.Closings_Req_CY1,
        c.Closings_Req_CY2,
        c.Closings_Req_CY3,
        c.Closings_Req_CY4,
        c.Closings_Req_CY5,
        c.Closings_Curr_CY,
        c.Closings_Curr_CY1,
        c.Closings_Curr_CY2,
        c.Closings_Curr_CY3,
        c.Closings_Curr_CY4,
        c.Closings_Curr_CY5,
        c.Closings_Var_CY,
        c.Closings_Var_CY1,
        c.Closings_Var_CY2,
        c.Closings_Var_CY3,
        c.Closings_Var_CY4,
        c.Closings_Var_CY5,
        p.submission_date,
        p.project_id
    FROM dbo.tbl_pif_projects_inflight p
        LEFT JOIN cost_pivot c ON p.pif_id = c.pif_id AND p.project_id = c.project_id;
GO

PRINT 'vw_TA_Inflight_wide updated successfully';
PRINT '';

-- ============================================================================
-- STEP 3: Drop and recreate vw_pif_approved_wide
-- ============================================================================

PRINT 'Updating vw_pif_approved_wide...';

IF OBJECT_ID('dbo.vw_pif_approved_wide', 'V') IS NOT NULL
    DROP VIEW dbo.vw_pif_approved_wide;
GO

CREATE VIEW dbo.vw_pif_approved_wide
AS
    WITH
        cost_pivot
        AS
        (
            SELECT
                pif_id,
                project_id,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() THEN requested_value END) AS Target_Req_CY,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 1 THEN requested_value END) AS Target_Req_CY1,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 2 THEN requested_value END) AS Target_Req_CY2,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 3 THEN requested_value END) AS Target_Req_CY3,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 4 THEN requested_value END) AS Target_Req_CY4,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 5 THEN requested_value END) AS Target_Req_CY5,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() THEN current_value END) AS Target_Curr_CY,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 1 THEN current_value END) AS Target_Curr_CY1,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 2 THEN current_value END) AS Target_Curr_CY2,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 3 THEN current_value END) AS Target_Curr_CY3,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 4 THEN current_value END) AS Target_Curr_CY4,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 5 THEN current_value END) AS Target_Curr_CY5,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() THEN variance_value END) AS Target_Var_CY,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 1 THEN variance_value END) AS Target_Var_CY1,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 2 THEN variance_value END) AS Target_Var_CY2,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 3 THEN variance_value END) AS Target_Var_CY3,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 4 THEN variance_value END) AS Target_Var_CY4,
                MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = dbo.CY() + 5 THEN variance_value END) AS Target_Var_CY5,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() THEN requested_value END) AS Closings_Req_CY,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 1 THEN requested_value END) AS Closings_Req_CY1,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 2 THEN requested_value END) AS Closings_Req_CY2,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 3 THEN requested_value END) AS Closings_Req_CY3,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 4 THEN requested_value END) AS Closings_Req_CY4,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 5 THEN requested_value END) AS Closings_Req_CY5,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() THEN current_value END) AS Closings_Curr_CY,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 1 THEN current_value END) AS Closings_Curr_CY1,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 2 THEN current_value END) AS Closings_Curr_CY2,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 3 THEN current_value END) AS Closings_Curr_CY3,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 4 THEN current_value END) AS Closings_Curr_CY4,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 5 THEN current_value END) AS Closings_Curr_CY5,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() THEN variance_value END) AS Closings_Var_CY,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 1 THEN variance_value END) AS Closings_Var_CY1,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 2 THEN variance_value END) AS Closings_Var_CY2,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 3 THEN variance_value END) AS Closings_Var_CY3,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 4 THEN variance_value END) AS Closings_Var_CY4,
                MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = dbo.CY() + 5 THEN variance_value END) AS Closings_Var_CY5
            FROM dbo.tbl_pif_cost_approved
            GROUP BY pif_id, project_id
        )
    SELECT
        p.archive_flag,
        p.include_flag,
        p.accounting_treatment,
        p.change_type,
        p.pif_id,
        p.seg,
        p.opco,
        p.site,
        p.strategic_rank,
        p.funding_project,
        p.project_name,
        p.original_fp_isd,
        p.revised_fp_isd,
        p.lcm_issue,
        p.status,
        p.category,
        p.justification,
        c.Target_Req_CY,
        c.Target_Req_CY1,
        c.Target_Req_CY2,
        c.Target_Req_CY3,
        c.Target_Req_CY4,
        c.Target_Req_CY5,
        c.Target_Curr_CY,
        c.Target_Curr_CY1,
        c.Target_Curr_CY2,
        c.Target_Curr_CY3,
        c.Target_Curr_CY4,
        c.Target_Curr_CY5,
        c.Target_Var_CY,
        c.Target_Var_CY1,
        c.Target_Var_CY2,
        c.Target_Var_CY3,
        c.Target_Var_CY4,
        c.Target_Var_CY5,
        p.moving_isd_year,
        p.prior_year_spend,
        c.Closings_Req_CY,
        c.Closings_Req_CY1,
        c.Closings_Req_CY2,
        c.Closings_Req_CY3,
        c.Closings_Req_CY4,
        c.Closings_Req_CY5,
        c.Closings_Curr_CY,
        c.Closings_Curr_CY1,
        c.Closings_Curr_CY2,
        c.Closings_Curr_CY3,
        c.Closings_Curr_CY4,
        c.Closings_Curr_CY5,
        c.Closings_Var_CY,
        c.Closings_Var_CY1,
        c.Closings_Var_CY2,
        c.Closings_Var_CY3,
        c.Closings_Var_CY4,
        c.Closings_Var_CY5,
        p.approval_date,
        p.submission_date,
        p.project_id
    FROM dbo.tbl_pif_projects_approved p
        LEFT JOIN cost_pivot c ON p.pif_id = c.pif_id AND p.project_id = c.project_id;
GO

PRINT 'vw_pif_approved_wide updated successfully';
PRINT '';

-- ============================================================================
-- STEP 4: Verification
-- ============================================================================

PRINT 'Verifying views were created successfully...';

IF OBJECT_ID('dbo.vw_TA_Inflight_wide', 'V') IS NULL
BEGIN
    PRINT 'ERROR: vw_TA_Inflight_wide was not created!';
    RAISERROR('View creation failed: vw_TA_Inflight_wide', 16, 1);
    RETURN;
END

IF OBJECT_ID('dbo.vw_pif_approved_wide', 'V') IS NULL
BEGIN
    PRINT 'ERROR: vw_pif_approved_wide was not created!';
    RAISERROR('View creation failed: vw_pif_approved_wide', 16, 1);
    RETURN;
END

PRINT 'Verification passed: All views created successfully';
PRINT '';

-- ============================================================================
-- COMPLETION
-- ============================================================================

PRINT '============================================================================';
PRINT 'Migration completed successfully!';
PRINT '';
PRINT 'Changes applied:';
PRINT '  - vw_TA_Inflight_wide: Now uses dbo.CY() for reporting period year';
PRINT '  - vw_pif_approved_wide: Now uses dbo.CY() for reporting period year';
PRINT '';
PRINT 'The views now properly respect reporting periods instead of calendar dates.';
PRINT '============================================================================';
GO
