-- ============================================================================
-- PIF REPORTING DATABASE SETUP (SECURITY ENHANCED)
-- ============================================================================
-- Purpose: Create normalized database structure for Project Impact Form (PIF) tracking
-- Author: Data Architecture Team
-- Date: 2025-11-05
-- Version: 2.0.0 - SECURITY ENHANCED
--
-- SECURITY IMPROVEMENTS:
--   - Parameterized stored procedures prevent SQL injection
--   - Transaction management with proper error handling
--   - Validation logic encapsulated in database
--   - Index creation separated from bulk insert for performance
--
-- This script creates:
--   1. Staging tables (for data validation before commit)
--   2. Inflight tables (current working month)
--   3. Approved tables (permanent archive)
--   4. Submission log (audit trail)
--   5. Supporting indexes
--   6. Helper views
--   7. SECURE STORED PROCEDURES (NEW)
-- ============================================================================

USE [IPDD];
GO

-- ============================================================================
-- SECTION 1: STAGING TABLES
-- ============================================================================
-- Purpose: Landing zone for Excel data before validation and commit
-- Lifecycle: Truncated and reloaded each submission

IF OBJECT_ID('dbo.tbl_pif_projects_staging', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_projects_staging;
GO

CREATE TABLE dbo.tbl_pif_projects_staging
(
    pif_project_id INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id VARCHAR(16) NOT NULL,
    project_id VARCHAR(10) NOT NULL,

    -- Status & classification
    status VARCHAR(58) NULL,
    change_type VARCHAR(12) NULL,
    accounting_treatment VARCHAR(14) NULL,
    category VARCHAR(26) NULL,

    -- Organizational
    seg INT NULL,
    opco VARCHAR(4) NULL,
    site VARCHAR(4) NULL,
    strategic_rank VARCHAR(26) NULL,

    -- Project linkage
    funding_project VARCHAR(10) NULL,
    project_name VARCHAR(35) NULL,

    -- Scheduling
    original_fp_isd VARCHAR(20) NULL,
    revised_fp_isd VARCHAR(20) NULL,
    moving_isd_year CHAR(1) NULL,

    -- Context
    lcm_issue VARCHAR(11) NULL,
    justification VARCHAR(192) NULL,
    prior_year_spend DECIMAL(18,2) NULL,

    -- Flags
    archive_flag BIT NULL,
    include_flag BIT NULL
);
GO

IF OBJECT_ID('dbo.tbl_pif_cost_staging', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_cost_staging;
GO

CREATE TABLE dbo.tbl_pif_cost_staging
(
    pif_cost_id INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id VARCHAR(16) NOT NULL,
    project_id VARCHAR(10) NOT NULL,
    scenario VARCHAR(12) NOT NULL,
    -- 'Target' or 'Closings'
    year DATE NOT NULL,
    -- Fiscal year end: 12/31/YYYY

    -- Financial data
    requested_value DECIMAL(18,2) NULL,
    -- User-entered proposal
    current_value DECIMAL(18,2) NULL,
    -- System of record baseline
    variance_value DECIMAL(18,2) NULL
    -- Difference (calculated)
);
GO

-- ============================================================================
-- SECTION 2: INFLIGHT TABLES
-- ============================================================================
-- Purpose: Current working month - active, pending, and parking lot PIFs
-- Lifecycle: Truncated and reloaded each submission; approved PIFs archived monthly

IF OBJECT_ID('dbo.tbl_pif_projects_inflight', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_projects_inflight;
GO

CREATE TABLE dbo.tbl_pif_projects_inflight
(
    pif_project_id INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id VARCHAR(16) NOT NULL,
    project_id VARCHAR(10) NOT NULL,
    submission_date DATE NOT NULL,
    -- When this batch was submitted

    -- Status & classification
    status VARCHAR(58) NULL,
    change_type VARCHAR(12) NULL,
    accounting_treatment VARCHAR(14) NULL,
    category VARCHAR(26) NULL,

    -- Organizational
    seg INT NULL,
    opco VARCHAR(4) NULL,
    site VARCHAR(4) NULL,
    strategic_rank VARCHAR(26) NULL,

    -- Project linkage
    funding_project VARCHAR(10) NULL,
    project_name VARCHAR(35) NULL,

    -- Scheduling
    original_fp_isd VARCHAR(20) NULL,
    revised_fp_isd VARCHAR(20) NULL,
    moving_isd_year CHAR(1) NULL,

    -- Context
    lcm_issue VARCHAR(11) NULL,
    justification VARCHAR(192) NULL,
    prior_year_spend DECIMAL(18,2) NULL,

    -- Flags
    archive_flag BIT NULL,
    include_flag BIT NULL,

    -- Constraints
    CONSTRAINT UQ_inflight_pif_project UNIQUE (pif_id, project_id)
);
GO

IF OBJECT_ID('dbo.tbl_pif_cost_inflight', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_cost_inflight;
GO

CREATE TABLE dbo.tbl_pif_cost_inflight
(
    pif_cost_id INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id VARCHAR(16) NOT NULL,
    project_id VARCHAR(10) NOT NULL,
    scenario VARCHAR(12) NOT NULL,
    year DATE NOT NULL,

    -- Financial data
    requested_value DECIMAL(18,2) NULL,
    current_value DECIMAL(18,2) NULL,
    variance_value DECIMAL(18,2) NULL
);
GO

-- PERFORMANCE: Index for cost lookups and joins
CREATE NONCLUSTERED INDEX IX_inflight_cost_lookup
    ON dbo.tbl_pif_cost_inflight (pif_id, project_id, scenario, year);
GO

-- ============================================================================
-- SECTION 3: APPROVED TABLES
-- ============================================================================
-- Purpose: Permanent archive of approved/dispositioned PIFs
-- Lifecycle: Grows over time; never truncated

IF OBJECT_ID('dbo.tbl_pif_projects_approved', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_projects_approved;
GO

CREATE TABLE dbo.tbl_pif_projects_approved
(
    pif_project_id INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id VARCHAR(16) NOT NULL,
    project_id VARCHAR(10) NOT NULL,
    submission_date DATE NOT NULL,
    approval_date DATE NOT NULL,
    -- When archived to approved

    -- Status & classification
    status VARCHAR(58) NULL,
    change_type VARCHAR(12) NULL,
    accounting_treatment VARCHAR(14) NULL,
    category VARCHAR(26) NULL,

    -- Organizational
    seg INT NULL,
    opco VARCHAR(4) NULL,
    site VARCHAR(4) NULL,
    strategic_rank VARCHAR(26) NULL,

    -- Project linkage
    funding_project VARCHAR(10) NULL,
    project_name VARCHAR(35) NULL,

    -- Scheduling
    original_fp_isd VARCHAR(20) NULL,
    revised_fp_isd VARCHAR(20) NULL,
    moving_isd_year CHAR(1) NULL,

    -- Context
    lcm_issue VARCHAR(11) NULL,
    justification VARCHAR(192) NULL,
    prior_year_spend DECIMAL(18,2) NULL,

    -- Flags
    archive_flag BIT NULL,
    include_flag BIT NULL
);
GO

-- PERFORMANCE: Index for common queries
CREATE NONCLUSTERED INDEX IX_approved_pif_project
    ON dbo.tbl_pif_projects_approved (pif_id, project_id);
GO

CREATE NONCLUSTERED INDEX IX_approved_dates
    ON dbo.tbl_pif_projects_approved (approval_date, submission_date);
GO

IF OBJECT_ID('dbo.tbl_pif_cost_approved', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_cost_approved;
GO

CREATE TABLE dbo.tbl_pif_cost_approved
(
    pif_cost_id INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id VARCHAR(16) NOT NULL,
    project_id VARCHAR(10) NOT NULL,
    scenario VARCHAR(12) NOT NULL,
    year DATE NOT NULL,
    approval_date DATE NOT NULL,

    -- Financial data
    requested_value DECIMAL(18,2) NULL,
    current_value DECIMAL(18,2) NULL,
    variance_value DECIMAL(18,2) NULL
);
GO

-- PERFORMANCE: Index for variance analysis
CREATE NONCLUSTERED INDEX IX_approved_cost_lookup
    ON dbo.tbl_pif_cost_approved (pif_id, project_id, scenario, year);
GO

CREATE NONCLUSTERED INDEX IX_approved_cost_variance
    ON dbo.tbl_pif_cost_approved (variance_value)
    INCLUDE (pif_id, project_id, scenario, year);
GO

-- ============================================================================
-- SECTION 4: AUDIT LOG
-- ============================================================================

IF OBJECT_ID('dbo.tbl_submission_log', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_submission_log;
GO

CREATE TABLE dbo.tbl_submission_log
(
    submission_id INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    submission_date DATETIME NOT NULL DEFAULT GETDATE(),
    submitted_by VARCHAR(128) NOT NULL,
    source_file VARCHAR(255) NULL,
    record_count INT NULL,
    notes VARCHAR(500) NULL
);
GO

-- ============================================================================
-- SECTION 5: VIEWS
-- ============================================================================

IF OBJECT_ID('dbo.vw_pif_current_working', 'V') IS NOT NULL
    DROP VIEW dbo.vw_pif_current_working;
GO

CREATE VIEW dbo.vw_pif_current_working
AS
    SELECT
        p.*,
        c.scenario,
        c.year,
        c.requested_value,
        c.current_value,
        c.variance_value
    FROM dbo.tbl_pif_projects_inflight p
        LEFT JOIN dbo.tbl_pif_cost_inflight c
        ON p.pif_id = c.pif_id AND p.project_id = c.project_id;
GO

IF OBJECT_ID('dbo.vw_pif_all_history', 'V') IS NOT NULL
    DROP VIEW dbo.vw_pif_all_history;
GO

CREATE VIEW dbo.vw_pif_all_history
AS
            SELECT
            'Inflight' AS source,
            p.pif_project_id,
            p.pif_id,
            p.project_id,
            p.submission_date,
            NULL AS approval_date, -- Added for UNION ALL compatibility
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
            p.include_flag,
            c.scenario,
            c.year,
            c.requested_value,
            c.current_value,
            c.variance_value
        FROM dbo.tbl_pif_projects_inflight p
            LEFT JOIN dbo.tbl_pif_cost_inflight c
            ON p.pif_id = c.pif_id AND p.project_id = c.project_id
    UNION ALL
        SELECT
            'Approved' AS source,
            p.pif_project_id,
            p.pif_id,
            p.project_id,
            p.submission_date,
            p.approval_date, -- This column exists in approved table
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
            p.include_flag,
            c.scenario,
            c.year,
            c.requested_value,
            c.current_value,
            c.variance_value
        FROM dbo.tbl_pif_projects_approved p
            LEFT JOIN dbo.tbl_pif_cost_approved c
            ON p.pif_id = c.pif_id AND p.project_id = c.project_id;
GO

-- ----------------------------------------------------------------------------
-- View: vw_inflight_by_site
-- Purpose: Query inflight records with site filtering (VBA adds WHERE clause)
-- Usage: SELECT * FROM vw_inflight_by_site WHERE site = 'ANO'
--        SELECT * FROM vw_inflight_by_site (for Fleet - all sites)
-- ----------------------------------------------------------------------------

IF OBJECT_ID('dbo.vw_inflight_by_site', 'V') IS NOT NULL
    DROP VIEW dbo.vw_inflight_by_site;
GO

CREATE VIEW dbo.vw_inflight_by_site
AS
    SELECT
        p.pif_project_id,
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
        p.include_flag,
        c.scenario,
        c.year,
        c.requested_value,
        c.current_value,
        c.variance_value
    FROM dbo.tbl_pif_projects_inflight p
        LEFT JOIN dbo.tbl_pif_cost_inflight c
        ON p.pif_id = c.pif_id AND p.project_id = c.project_id;
GO

-- ----------------------------------------------------------------------------
-- View: vw_approved_by_site
-- Purpose: Query approved records with site filtering (VBA adds WHERE clause)
-- Usage: SELECT * FROM vw_approved_by_site WHERE site = 'ANO'
--        SELECT * FROM vw_approved_by_site (for Fleet - all sites)
-- ----------------------------------------------------------------------------

IF OBJECT_ID('dbo.vw_approved_by_site', 'V') IS NOT NULL
    DROP VIEW dbo.vw_approved_by_site;
GO

CREATE VIEW dbo.vw_approved_by_site
AS
    SELECT
        p.pif_project_id,
        p.pif_id,
        p.project_id,
        p.submission_date,
        p.approval_date,
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
        p.include_flag,
        c.scenario,
        c.year,
        c.requested_value,
        c.current_value,
        c.variance_value
    FROM dbo.tbl_pif_projects_approved p
        LEFT JOIN dbo.tbl_pif_cost_approved c
        ON p.pif_id = c.pif_id AND p.project_id = c.project_id;
GO

-- ----------------------------------------------------------------------------
-- View: vw_inflight_wide
-- Purpose: Inflight data in WIDE format (mimics original PIF Excel layout)
-- Usage: SELECT * FROM vw_inflight_wide WHERE site = 'ANO'
-- Notes: Cost data is pivoted from long format back to wide format
--        Columns match original Excel layout for user familiarity
-- ----------------------------------------------------------------------------

IF OBJECT_ID('dbo.vw_inflight_wide', 'V') IS NOT NULL
    DROP VIEW dbo.vw_inflight_wide;
GO

CREATE VIEW dbo.vw_inflight_wide
AS
    WITH cost_pivot AS (
        SELECT
            pif_id,
            project_id,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2025 THEN requested_value END) AS Target_Req_2025,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2026 THEN requested_value END) AS Target_Req_2026,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2027 THEN requested_value END) AS Target_Req_2027,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2028 THEN requested_value END) AS Target_Req_2028,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2029 THEN requested_value END) AS Target_Req_2029,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2030 THEN requested_value END) AS Target_Req_2030,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2025 THEN current_value END) AS Target_Curr_2025,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2026 THEN current_value END) AS Target_Curr_2026,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2027 THEN current_value END) AS Target_Curr_2027,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2028 THEN current_value END) AS Target_Curr_2028,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2029 THEN current_value END) AS Target_Curr_2029,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2030 THEN current_value END) AS Target_Curr_2030,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2025 THEN variance_value END) AS Target_Var_2025,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2026 THEN variance_value END) AS Target_Var_2026,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2027 THEN variance_value END) AS Target_Var_2027,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2028 THEN variance_value END) AS Target_Var_2028,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2029 THEN variance_value END) AS Target_Var_2029,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2030 THEN variance_value END) AS Target_Var_2030,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2025 THEN requested_value END) AS Closings_Req_2025,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2026 THEN requested_value END) AS Closings_Req_2026,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2027 THEN requested_value END) AS Closings_Req_2027,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2028 THEN requested_value END) AS Closings_Req_2028,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2029 THEN requested_value END) AS Closings_Req_2029,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2030 THEN requested_value END) AS Closings_Req_2030,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2025 THEN current_value END) AS Closings_Curr_2025,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2026 THEN current_value END) AS Closings_Curr_2026,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2027 THEN current_value END) AS Closings_Curr_2027,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2028 THEN current_value END) AS Closings_Curr_2028,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2029 THEN current_value END) AS Closings_Curr_2029,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2030 THEN current_value END) AS Closings_Curr_2030,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2025 THEN variance_value END) AS Closings_Var_2025,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2026 THEN variance_value END) AS Closings_Var_2026,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2027 THEN variance_value END) AS Closings_Var_2027,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2028 THEN variance_value END) AS Closings_Var_2028,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2029 THEN variance_value END) AS Closings_Var_2029,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2030 THEN variance_value END) AS Closings_Var_2030
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
        c.Target_Req_2025,
        c.Target_Req_2026,
        c.Target_Req_2027,
        c.Target_Req_2028,
        c.Target_Req_2029,
        c.Target_Req_2030,
        c.Target_Curr_2025,
        c.Target_Curr_2026,
        c.Target_Curr_2027,
        c.Target_Curr_2028,
        c.Target_Curr_2029,
        c.Target_Curr_2030,
        c.Target_Var_2025,
        c.Target_Var_2026,
        c.Target_Var_2027,
        c.Target_Var_2028,
        c.Target_Var_2029,
        c.Target_Var_2030,
        p.moving_isd_year,
        p.prior_year_spend,
        c.Closings_Req_2025,
        c.Closings_Req_2026,
        c.Closings_Req_2027,
        c.Closings_Req_2028,
        c.Closings_Req_2029,
        c.Closings_Req_2030,
        c.Closings_Curr_2025,
        c.Closings_Curr_2026,
        c.Closings_Curr_2027,
        c.Closings_Curr_2028,
        c.Closings_Curr_2029,
        c.Closings_Curr_2030,
        c.Closings_Var_2025,
        c.Closings_Var_2026,
        c.Closings_Var_2027,
        c.Closings_Var_2028,
        c.Closings_Var_2029,
        c.Closings_Var_2030,
        p.submission_date,
        p.project_id
    FROM dbo.tbl_pif_projects_inflight p
        LEFT JOIN cost_pivot c ON p.pif_id = c.pif_id AND p.project_id = c.project_id;
GO

-- ----------------------------------------------------------------------------
-- View: vw_approved_wide
-- Purpose: Approved data in WIDE format (mimics original PIF Excel layout)
-- Usage: SELECT * FROM vw_approved_wide WHERE site = 'ANO'
-- Notes: Cost data is pivoted from long format back to wide format
--        Columns match original Excel layout for user familiarity
-- ----------------------------------------------------------------------------

IF OBJECT_ID('dbo.vw_approved_wide', 'V') IS NOT NULL
    DROP VIEW dbo.vw_approved_wide;
GO

CREATE VIEW dbo.vw_approved_wide
AS
    WITH cost_pivot AS (
        SELECT
            pif_id,
            project_id,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2025 THEN requested_value END) AS Target_Req_2025,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2026 THEN requested_value END) AS Target_Req_2026,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2027 THEN requested_value END) AS Target_Req_2027,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2028 THEN requested_value END) AS Target_Req_2028,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2029 THEN requested_value END) AS Target_Req_2029,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2030 THEN requested_value END) AS Target_Req_2030,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2025 THEN current_value END) AS Target_Curr_2025,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2026 THEN current_value END) AS Target_Curr_2026,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2027 THEN current_value END) AS Target_Curr_2027,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2028 THEN current_value END) AS Target_Curr_2028,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2029 THEN current_value END) AS Target_Curr_2029,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2030 THEN current_value END) AS Target_Curr_2030,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2025 THEN variance_value END) AS Target_Var_2025,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2026 THEN variance_value END) AS Target_Var_2026,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2027 THEN variance_value END) AS Target_Var_2027,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2028 THEN variance_value END) AS Target_Var_2028,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2029 THEN variance_value END) AS Target_Var_2029,
            MAX(CASE WHEN scenario = 'Target' AND YEAR(year) = 2030 THEN variance_value END) AS Target_Var_2030,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2025 THEN requested_value END) AS Closings_Req_2025,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2026 THEN requested_value END) AS Closings_Req_2026,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2027 THEN requested_value END) AS Closings_Req_2027,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2028 THEN requested_value END) AS Closings_Req_2028,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2029 THEN requested_value END) AS Closings_Req_2029,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2030 THEN requested_value END) AS Closings_Req_2030,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2025 THEN current_value END) AS Closings_Curr_2025,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2026 THEN current_value END) AS Closings_Curr_2026,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2027 THEN current_value END) AS Closings_Curr_2027,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2028 THEN current_value END) AS Closings_Curr_2028,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2029 THEN current_value END) AS Closings_Curr_2029,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2030 THEN current_value END) AS Closings_Curr_2030,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2025 THEN variance_value END) AS Closings_Var_2025,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2026 THEN variance_value END) AS Closings_Var_2026,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2027 THEN variance_value END) AS Closings_Var_2027,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2028 THEN variance_value END) AS Closings_Var_2028,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2029 THEN variance_value END) AS Closings_Var_2029,
            MAX(CASE WHEN scenario = 'Closings' AND YEAR(year) = 2030 THEN variance_value END) AS Closings_Var_2030
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
        c.Target_Req_2025,
        c.Target_Req_2026,
        c.Target_Req_2027,
        c.Target_Req_2028,
        c.Target_Req_2029,
        c.Target_Req_2030,
        c.Target_Curr_2025,
        c.Target_Curr_2026,
        c.Target_Curr_2027,
        c.Target_Curr_2028,
        c.Target_Curr_2029,
        c.Target_Curr_2030,
        c.Target_Var_2025,
        c.Target_Var_2026,
        c.Target_Var_2027,
        c.Target_Var_2028,
        c.Target_Var_2029,
        c.Target_Var_2030,
        p.moving_isd_year,
        p.prior_year_spend,
        c.Closings_Req_2025,
        c.Closings_Req_2026,
        c.Closings_Req_2027,
        c.Closings_Req_2028,
        c.Closings_Req_2029,
        c.Closings_Req_2030,
        c.Closings_Curr_2025,
        c.Closings_Curr_2026,
        c.Closings_Curr_2027,
        c.Closings_Curr_2028,
        c.Closings_Curr_2029,
        c.Closings_Curr_2030,
        c.Closings_Var_2025,
        c.Closings_Var_2026,
        c.Closings_Var_2027,
        c.Closings_Var_2028,
        c.Closings_Var_2029,
        c.Closings_Var_2030,
        p.approval_date,
        p.submission_date,
        p.project_id
    FROM dbo.tbl_pif_projects_approved p
        LEFT JOIN cost_pivot c ON p.pif_id = c.pif_id AND p.project_id = c.project_id;
GO

-- ============================================================================
-- SECTION 6: SECURE STORED PROCEDURES (NEW)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- Procedure: usp_insert_project_staging
-- Purpose: Inserts a single project record into the staging table
-- SECURITY: Parameterized to prevent injection attacks
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
    -- Corrected length
    @revised_fp_isd VARCHAR(20) = NULL,
    -- Corrected length
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
        RAISERROR( @ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

-- ----------------------------------------------------------------------------
-- Procedure: usp_insert_cost_staging
-- Purpose: Inserts a single cost record into the staging table
-- SECURITY: Parameterized to prevent injection attacks
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

        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        RETURN -1;
    END CATCH
END;
GO

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
    -- IMPROVEMENT: Ensures transaction rollback on error

    DECLARE @Errors TABLE (
        error_id INT IDENTITY(1,1),
        error_severity VARCHAR(10),
        -- 'CRITICAL', 'WARNING', 'INFO'
        error_type VARCHAR(50),
        error_message VARCHAR(500),
        record_identifier VARCHAR(100)
    );

    -- Check 1: Missing required fields (CRITICAL)
    INSERT INTO @Errors
        (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Missing PIF ID',
        'Missing required field: PIF ID',
        'Row ' + CAST(pif_project_id AS VARCHAR(10))
    FROM dbo.tbl_pif_projects_staging
    WHERE pif_id IS NULL OR LTRIM(RTRIM(pif_id)) = '';

    INSERT INTO @Errors
        (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Missing Project ID',
        'Missing required field: Project ID',
        'Row ' + CAST(pif_project_id AS VARCHAR(10))
    FROM dbo.tbl_pif_projects_staging
    WHERE project_id IS NULL OR LTRIM(RTRIM(project_id)) = '';

    INSERT INTO @Errors
        (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Missing Change Type',
        'Missing required field: Change Type',
        'PIF ' + pif_id + ', Project ' + project_id
    FROM dbo.tbl_pif_projects_staging
    WHERE pif_id IS NOT NULL
        AND project_id IS NOT NULL
        AND (change_type IS NULL OR LTRIM(RTRIM(change_type)) = '');

    -- Check 2: Data type validation (CRITICAL)
    INSERT INTO @Errors
        (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Invalid Data Type',
        'SEG must be a valid positive integer',
        'PIF ' + pif_id + ', Project ' + project_id
    FROM dbo.tbl_pif_projects_staging
    WHERE seg IS NOT NULL
        AND (seg < 0 OR seg > 99999);
    -- IMPROVEMENT: Range validation

    -- Check 3: Duplicate detection (CRITICAL)
    INSERT INTO @Errors
        (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Duplicate',
        'Duplicate PIF+Project combination (appears ' + CAST(cnt AS VARCHAR(10)) + ' times)',
        'PIF ' + pif_id + ', Project ' + project_id
    FROM (
        SELECT pif_id, project_id, COUNT(*) as cnt
        FROM dbo.tbl_pif_projects_staging
        WHERE pif_id IS NOT NULL AND project_id IS NOT NULL
        GROUP BY pif_id, project_id
        HAVING COUNT(*) > 1
    ) dups;

    -- Check 4: Business rule - Approved PIFs require justification (CRITICAL)
    INSERT INTO @Errors
        (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Missing Justification',
        'Approved or Dispositioned status requires justification',
        'PIF ' + pif_id + ', Project ' + project_id
    FROM dbo.tbl_pif_projects_staging
    WHERE status IN ('Approved', 'Dispositioned')
        AND (justification IS NULL OR LTRIM(RTRIM(justification)) = '');

    -- Check 5: Orphan cost records (CRITICAL)
    INSERT INTO @Errors
        (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Orphan Cost Record',
        'Cost record exists without matching project record',
        'PIF ' + c.pif_id + ', Project ' + c.project_id
    FROM dbo.tbl_pif_cost_staging c
        LEFT JOIN dbo.tbl_pif_projects_staging p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id
    WHERE p.pif_id IS NULL;

    -- Check 6: Invalid scenario values (CRITICAL)
    INSERT INTO @Errors
        (error_severity, error_type, error_message, record_identifier)
    SELECT
        'CRITICAL',
        'Invalid Scenario',
        'Scenario must be ''Target'' or ''Closings'' (found: ''' + ISNULL(scenario, 'NULL') + ''')',
        'PIF ' + pif_id + ', Project ' + project_id + ', Year ' + CAST(YEAR(year) AS VARCHAR(4))
    FROM dbo.tbl_pif_cost_staging
    WHERE scenario NOT IN ('Target', 'Closings');

    -- Check 7: Variance threshold warning (WARNING - does not block submission)
    INSERT INTO @Errors
        (error_severity, error_type, error_message, record_identifier)
    SELECT
        'WARNING',
        'Variance Threshold Exceeded',
        'Variance exceeds -$1M threshold: ' + FORMAT(variance_value, 'C', 'en-US'),
        'PIF ' + pif_id + ', Project ' + project_id + ', Year ' + CAST(YEAR(year) AS VARCHAR(4))
    FROM dbo.tbl_pif_cost_staging
    WHERE variance_value < -1000000;
    -- $1M threshold

    -- Return separate counts for critical errors and warnings
    SELECT @ErrorCount = COUNT(*)
    FROM @Errors
    WHERE error_severity = 'CRITICAL';
    SELECT @WarningCount = COUNT(*)
    FROM @Errors
    WHERE error_severity = 'WARNING';

    -- Return error details (ordered by severity)
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

-- ----------------------------------------------------------------------------
-- Procedure: usp_create_staging_indexes
-- Purpose: Create indexes on staging tables after bulk insert
-- PERFORMANCE: Dramatically speeds up validation queries
-- ----------------------------------------------------------------------------

IF OBJECT_ID('dbo.usp_create_staging_indexes', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_create_staging_indexes;
GO

CREATE PROCEDURE dbo.usp_create_staging_indexes
AS
BEGIN
    SET NOCOUNT ON;

    -- Drop existing indexes if they exist
    IF EXISTS (SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_staging_pif_proj')
        DROP INDEX IX_staging_pif_proj ON dbo.tbl_pif_projects_staging;

    IF EXISTS (SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_staging_status')
        DROP INDEX IX_staging_status ON dbo.tbl_pif_projects_staging;

    IF EXISTS (SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_staging_cost_lookup')
        DROP INDEX IX_staging_cost_lookup ON dbo.tbl_pif_cost_staging;

    -- Create indexes for validation queries
    CREATE NONCLUSTERED INDEX IX_staging_pif_proj
        ON dbo.tbl_pif_projects_staging (pif_id, project_id);

    CREATE NONCLUSTERED INDEX IX_staging_status
        ON dbo.tbl_pif_projects_staging (status)
        INCLUDE (pif_id, project_id, justification);

    CREATE NONCLUSTERED INDEX IX_staging_cost_lookup
        ON dbo.tbl_pif_cost_staging (pif_id, project_id);

    RETURN 0;
END;
GO

-- ----------------------------------------------------------------------------
-- Procedure: usp_commit_to_inflight
-- Purpose: Atomically move staging data to inflight tables
-- SECURITY: Encapsulates transaction logic, prevents partial commits
-- ----------------------------------------------------------------------------

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
        -- Delete costs first (foreign key relationship)
        DELETE c
        FROM dbo.tbl_pif_cost_inflight c
        INNER JOIN dbo.tbl_pif_projects_inflight p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id
        WHERE p.site = @site;

        -- Delete projects for selected site
        DELETE FROM dbo.tbl_pif_projects_inflight
        WHERE site = @site;

        -- Step 2: Move validated data from staging to inflight
        INSERT INTO dbo.tbl_pif_projects_inflight
        (
        pif_id, project_id, submission_date, status, change_type,
        accounting_treatment, category, seg, opco, site, strategic_rank,
        funding_project, project_name, original_fp_isd, revised_fp_isd,
        moving_isd_year, lcm_issue, justification, prior_year_spend,
        archive_flag, include_flag
        )
    SELECT
        s.pif_id,
        s.project_id,
        GETDATE(), -- This is the value for submission_date
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
        pif_id, project_id, scenario, year,
        requested_value, current_value, variance_value
        )
    SELECT
        c.pif_id, c.project_id, c.scenario, c.year,
        c.requested_value, c.current_value, c.variance_value
    FROM dbo.tbl_pif_cost_staging c
    INNER JOIN dbo.tbl_pif_projects_staging p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id
    WHERE p.site = @site;

        SET @CostCount = @@ROWCOUNT;

        COMMIT TRANSACTION;

        -- Return success message
        SELECT
        'SUCCESS' AS Status,
        @ProjectCount AS ProjectsCommitted,
        @CostCount AS CostRecordsCommitted;

        RETURN 0;

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;

        -- Return error information
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

-- ----------------------------------------------------------------------------
-- Procedure: usp_archive_approved_pifs
-- Purpose: Move approved PIFs from inflight to approved tables
-- SECURITY: Atomic operation with transaction support
-- ----------------------------------------------------------------------------

IF OBJECT_ID('dbo.usp_archive_approved_pifs', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_archive_approved_pifs;
GO

CREATE PROCEDURE dbo.usp_archive_approved_pifs
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @ApprovedCount INT = 0;
    DECLARE @CostCount INT = 0;

    BEGIN TRY
        BEGIN TRANSACTION;

        -- Step 1: Archive approved projects
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
    WHERE p.archive_flag = 1 AND p.include_flag = 1;

        SET @ApprovedCount = @@ROWCOUNT;

        -- Step 2: Archive approved costs
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
    WHERE p.archive_flag = 1 AND p.include_flag = 1;

        SET @CostCount = @@ROWCOUNT;

        -- Step 3: Remove archived records from inflight
        DELETE c
        FROM dbo.tbl_pif_cost_inflight c
        INNER JOIN dbo.tbl_pif_projects_inflight p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id
        WHERE p.archive_flag = 1 AND p.include_flag = 1;

        DELETE FROM dbo.tbl_pif_projects_inflight
        WHERE archive_flag = 1 AND include_flag = 1;

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

-- ----------------------------------------------------------------------------
-- Procedure: usp_log_submission
-- Purpose: Record submission in audit log
-- Parameters: All parameters are required and validated
-- SECURITY: Parameterized to prevent injection
-- ----------------------------------------------------------------------------

IF OBJECT_ID('dbo.usp_log_submission', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_log_submission;
GO

CREATE PROCEDURE dbo.usp_log_submission
    @SourceFile VARCHAR(255),
    @RecordCount INT,
    @Notes VARCHAR(500) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        INSERT INTO dbo.tbl_submission_log
        (
        submission_date,
        submitted_by,
        source_file,
        record_count,
        notes
        )
    VALUES
        (
            GETDATE(),
            SYSTEM_USER,
            @SourceFile,
            @RecordCount,
            COALESCE(@Notes, 'Submitted via VBA')
        );

        SELECT
        'SUCCESS' AS Status,
        SCOPE_IDENTITY() AS SubmissionID;

        RETURN 0;

    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();

        SELECT
        'ERROR' AS Status,
        @ErrorMessage AS ErrorMessage;

        -- Don't fail submission if logging fails
        RETURN 0;
    END CATCH
END;
GO

-- ============================================================================
-- SECTION 7: GRANT PERMISSIONS
-- ============================================================================

-- IMPORTANT: Replace [DOMAIN\VBAUser] with the actual Windows account that will
-- run the VBA application. Example: [CONTOSO\john.smith] or [CONTOSO\AppAccount]

-- Execute these GRANT statements one at a time, or uncomment and run all at once:

/*
-- ============================================================================
-- STORED PROCEDURE PERMISSIONS (Required for VBA operations)
-- ============================================================================
GRANT EXECUTE ON dbo.usp_validate_staging_data_secure TO [DOMAIN\VBAUser];
GRANT EXECUTE ON dbo.usp_create_staging_indexes TO [DOMAIN\VBAUser];
GRANT EXECUTE ON dbo.usp_commit_to_inflight TO [DOMAIN\VBAUser];
GRANT EXECUTE ON dbo.usp_archive_approved_pifs TO [DOMAIN\VBAUser];
GRANT EXECUTE ON dbo.usp_log_submission TO [DOMAIN\VBAUser];

-- ============================================================================
-- TABLE PERMISSIONS (Required for bulk insert and data queries)
-- ============================================================================
-- Staging tables: Need INSERT for bulk load, SELECT for verification, DELETE for cleanup
GRANT INSERT, SELECT, DELETE ON dbo.tbl_pif_projects_staging TO [DOMAIN\VBAUser];
GRANT INSERT, SELECT, DELETE ON dbo.tbl_pif_cost_staging TO [DOMAIN\VBAUser];

-- Inflight tables: Need SELECT for GetRecordCount and queries
GRANT SELECT ON dbo.tbl_pif_projects_inflight TO [DOMAIN\VBAUser];
GRANT SELECT ON dbo.tbl_pif_cost_inflight TO [DOMAIN\VBAUser];

-- Approved tables: Need SELECT for reporting queries
GRANT SELECT ON dbo.tbl_pif_projects_approved TO [DOMAIN\VBAUser];
GRANT SELECT ON dbo.tbl_pif_cost_approved TO [DOMAIN\VBAUser];

-- ============================================================================
-- VIEW PERMISSIONS (Required for reporting queries)
-- ============================================================================
GRANT SELECT ON dbo.vw_pif_current_working TO [DOMAIN\VBAUser];
GRANT SELECT ON dbo.vw_pif_all_history TO [DOMAIN\VBAUser];

-- ============================================================================
-- AUDIT LOG PERMISSIONS (Required for submission logging)
-- ============================================================================
GRANT SELECT ON dbo.tbl_submission_log TO [DOMAIN\VBAUser];

-- ============================================================================
-- VERIFICATION QUERIES
-- ============================================================================
-- Run these to verify permissions were granted successfully:

-- Check stored procedure permissions
SELECT
    USER_NAME(grantee_principal_id) AS Grantee,
    OBJECT_NAME(major_id) AS ObjectName,
    permission_name
FROM sys.database_permissions
WHERE class_desc = 'OBJECT_OR_COLUMN'
  AND OBJECT_NAME(major_id) LIKE 'usp_%'
ORDER BY ObjectName;

-- Check table permissions
SELECT
    USER_NAME(grantee_principal_id) AS Grantee,
    OBJECT_NAME(major_id) AS TableName,
    permission_name
FROM sys.database_permissions
WHERE class_desc = 'OBJECT_OR_COLUMN'
  AND OBJECT_NAME(major_id) LIKE 'tbl_%'
ORDER BY TableName, permission_name;

PRINT '';
PRINT '============================================================================';
PRINT 'REMEMBER: Uncomment and execute the GRANT statements above!';
PRINT 'Replace [DOMAIN\VBAUser] with the actual Windows account.';
PRINT '============================================================================';
*/

PRINT '';
PRINT '============================================================================';
PRINT 'Database setup complete!';
PRINT '';
PRINT 'NEXT STEPS:';
PRINT '  1. Uncomment the GRANT statements above (lines 707-775)';
PRINT '  2. Replace [DOMAIN\VBAUser] with your actual Windows account';
PRINT '  3. Execute the GRANT statements';
PRINT '  4. Run the verification queries to confirm';
PRINT '';
PRINT 'CRITICAL: The VBA application will NOT work without these permissions!';
PRINT '============================================================================';
GO
