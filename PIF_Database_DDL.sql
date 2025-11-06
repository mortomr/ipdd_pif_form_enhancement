-- ============================================================================
-- PIF REPORTING DATABASE SETUP
-- ============================================================================
-- Purpose: Create normalized database structure for Project Impact Form (PIF) tracking
-- Author: Data Architecture Team
-- Date: 2025-11-05
-- 
-- This script creates:
--   1. Staging tables (for data validation before commit)
--   2. Inflight tables (current working month)
--   3. Approved tables (permanent archive)
--   4. Submission log (audit trail)
--   5. Supporting indexes
--   6. Helper views
-- ============================================================================

USE [YOUR_DATABASE_NAME];
GO

-- ============================================================================
-- SECTION 1: STAGING TABLES
-- ============================================================================
-- Purpose: Landing zone for Excel data before validation and commit
-- Lifecycle: Truncated and reloaded each submission

IF OBJECT_ID('dbo.tbl_pif_projects_staging', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_projects_staging;
GO

CREATE TABLE dbo.tbl_pif_projects_staging (
    pif_project_id     INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id             VARCHAR(16) NOT NULL,
    project_id         VARCHAR(10) NOT NULL,
    
    -- Status & classification
    status             VARCHAR(58) NULL,
    change_type        VARCHAR(12) NULL,
    accounting_treatment VARCHAR(14) NULL,
    category           VARCHAR(26) NULL,
    
    -- Organizational
    seg                INT NULL,
    opco               VARCHAR(4) NULL,
    site               VARCHAR(4) NULL,
    strategic_rank     VARCHAR(26) NULL,
    
    -- Project linkage
    funding_project    VARCHAR(10) NULL,
    project_name       VARCHAR(35) NULL,
    
    -- Scheduling
    original_fp_isd    VARCHAR(8) NULL,
    revised_fp_isd     VARCHAR(5) NULL,
    moving_isd_year    CHAR(1) NULL,
    
    -- Context
    lcm_issue          VARCHAR(11) NULL,
    justification      VARCHAR(192) NULL,
    prior_year_spend   DECIMAL(18,2) NULL,
    
    -- Flags
    archive_flag       BIT NULL,
    include_flag       BIT NULL
);
GO

IF OBJECT_ID('dbo.tbl_pif_cost_staging', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_cost_staging;
GO

CREATE TABLE dbo.tbl_pif_cost_staging (
    pif_cost_id        INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id             VARCHAR(16) NOT NULL,
    project_id         VARCHAR(10) NOT NULL,
    scenario           VARCHAR(12) NOT NULL,  -- 'Target' or 'Closings'
    year               DATE NOT NULL,         -- Fiscal year end: 12/31/YYYY
    
    -- Financial data
    requested_value    DECIMAL(18,2) NULL,    -- User-entered proposal
    current_value      DECIMAL(18,2) NULL,    -- System of record baseline
    variance_value     DECIMAL(18,2) NULL     -- Difference (will be recalculated)
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

CREATE TABLE dbo.tbl_pif_projects_inflight (
    pif_project_id     INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id             VARCHAR(16) NOT NULL,
    project_id         VARCHAR(10) NOT NULL,
    submission_date    DATE NOT NULL,         -- When this batch was submitted
    
    -- Status & classification
    status             VARCHAR(58) NULL,
    change_type        VARCHAR(12) NULL,
    accounting_treatment VARCHAR(14) NULL,
    category           VARCHAR(26) NULL,
    
    -- Organizational
    seg                INT NULL,
    opco               VARCHAR(4) NULL,
    site               VARCHAR(4) NULL,
    strategic_rank     VARCHAR(26) NULL,
    
    -- Project linkage
    funding_project    VARCHAR(10) NULL,
    project_name       VARCHAR(35) NULL,
    
    -- Scheduling
    original_fp_isd    VARCHAR(8) NULL,
    revised_fp_isd     VARCHAR(5) NULL,
    moving_isd_year    CHAR(1) NULL,
    
    -- Context
    lcm_issue          VARCHAR(11) NULL,
    justification      VARCHAR(192) NULL,
    prior_year_spend   DECIMAL(18,2) NULL,
    
    -- Flags
    archive_flag       BIT NULL,
    include_flag       BIT NULL,
    
    CONSTRAINT UQ_inflight_pif_proj UNIQUE (pif_id, project_id)
);
GO

IF OBJECT_ID('dbo.tbl_pif_cost_inflight', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_cost_inflight;
GO

CREATE TABLE dbo.tbl_pif_cost_inflight (
    pif_cost_id        INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id             VARCHAR(16) NOT NULL,
    project_id         VARCHAR(10) NOT NULL,
    scenario           VARCHAR(12) NOT NULL,
    year               DATE NOT NULL,
    
    -- Financial data
    requested_value    DECIMAL(18,2) NULL,
    current_value      DECIMAL(18,2) NULL,
    variance_value     AS (requested_value - current_value) PERSISTED,  -- Computed column
    
    load_date          DATETIME DEFAULT(GETDATE()) NULL,
    
    CONSTRAINT FK_inflight_cost_to_project
        FOREIGN KEY (pif_id, project_id)
        REFERENCES dbo.tbl_pif_projects_inflight(pif_id, project_id)
        ON DELETE CASCADE
);
GO

-- Indexes for inflight queries
CREATE INDEX IX_inflight_pif ON dbo.tbl_pif_projects_inflight (pif_id);
CREATE INDEX IX_inflight_proj ON dbo.tbl_pif_projects_inflight (project_id);
CREATE INDEX IX_inflight_status ON dbo.tbl_pif_projects_inflight (status);
CREATE INDEX IX_inflight_submission ON dbo.tbl_pif_projects_inflight (submission_date);

CREATE INDEX IX_inflight_cost_lookup ON dbo.tbl_pif_cost_inflight (pif_id, project_id)
    INCLUDE (year, scenario, variance_value);
CREATE INDEX IX_inflight_cost_year ON dbo.tbl_pif_cost_inflight (year);
GO

-- ============================================================================
-- SECTION 3: APPROVED TABLES
-- ============================================================================
-- Purpose: Permanent archive of approved/dispositioned PIFs with historical snapshot
-- Lifecycle: Append-only; data never modified once inserted

IF OBJECT_ID('dbo.tbl_pif_projects_approved', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_projects_approved;
GO

CREATE TABLE dbo.tbl_pif_projects_approved (
    pif_project_id     INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id             VARCHAR(16) NOT NULL,
    project_id         VARCHAR(10) NOT NULL,
    submission_date    DATE NOT NULL,
    approval_date      DATE NOT NULL,         -- When moved to approved
    
    -- Status & classification (will be 'Approved' or 'Dispositioned')
    status             VARCHAR(58) NULL,
    change_type        VARCHAR(12) NULL,
    accounting_treatment VARCHAR(14) NULL,
    category           VARCHAR(26) NULL,
    
    -- Organizational
    seg                INT NULL,
    opco               VARCHAR(4) NULL,
    site               VARCHAR(4) NULL,
    strategic_rank     VARCHAR(26) NULL,
    
    -- Project linkage
    funding_project    VARCHAR(10) NULL,
    project_name       VARCHAR(35) NULL,
    
    -- Scheduling
    original_fp_isd    VARCHAR(8) NULL,
    revised_fp_isd     VARCHAR(5) NULL,
    moving_isd_year    CHAR(1) NULL,
    
    -- Context
    lcm_issue          VARCHAR(11) NULL,
    justification      VARCHAR(192) NULL,
    prior_year_spend   DECIMAL(18,2) NULL,
    
    -- Flags
    archive_flag       BIT NULL,
    include_flag       BIT NULL
);
GO

IF OBJECT_ID('dbo.tbl_pif_cost_approved', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_pif_cost_approved;
GO

CREATE TABLE dbo.tbl_pif_cost_approved (
    pif_cost_id        INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    pif_id             VARCHAR(16) NOT NULL,
    project_id         VARCHAR(10) NOT NULL,
    scenario           VARCHAR(12) NOT NULL,
    year               DATE NOT NULL,
    
    -- HISTORICAL SNAPSHOT: All values stored (not computed)
    requested_value    DECIMAL(18,2) NULL,
    current_value      DECIMAL(18,2) NULL,
    variance_value     DECIMAL(18,2) NULL,    -- Stored value preserves historical truth
    
    approval_date      DATE NOT NULL,
    load_date          DATETIME DEFAULT(GETDATE()) NULL,
    
    CONSTRAINT FK_approved_cost_to_project
        FOREIGN KEY (pif_id, project_id)
        REFERENCES dbo.tbl_pif_projects_approved(pif_id, project_id)
        ON DELETE CASCADE
);
GO

-- Indexes for approved queries (analytical workload)
CREATE INDEX IX_approved_pif ON dbo.tbl_pif_projects_approved (pif_id);
CREATE INDEX IX_approved_proj ON dbo.tbl_pif_projects_approved (project_id);
CREATE INDEX IX_approved_date ON dbo.tbl_pif_projects_approved (approval_date);
CREATE INDEX IX_approved_submission ON dbo.tbl_pif_projects_approved (submission_date);

CREATE INDEX IX_approved_cost_lookup ON dbo.tbl_pif_cost_approved (pif_id, project_id)
    INCLUDE (year, scenario, variance_value);
CREATE INDEX IX_approved_cost_year ON dbo.tbl_pif_cost_approved (year)
    INCLUDE (variance_value);
CREATE INDEX IX_approved_cost_proj_year ON dbo.tbl_pif_cost_approved (project_id, year)
    INCLUDE (variance_value);

-- Optional: Filtered index for non-zero variances (saves space)
CREATE INDEX IX_approved_cost_nonzero_variance ON dbo.tbl_pif_cost_approved (pif_id, year)
    INCLUDE (variance_value)
    WHERE variance_value <> 0;
GO

-- ============================================================================
-- SECTION 4: SUBMISSION LOG
-- ============================================================================
-- Purpose: Audit trail of who submitted what when

IF OBJECT_ID('dbo.tbl_submission_log', 'U') IS NOT NULL
    DROP TABLE dbo.tbl_submission_log;
GO

CREATE TABLE dbo.tbl_submission_log (
    submission_id      INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    submission_date    DATETIME DEFAULT(GETDATE()) NOT NULL,
    submitted_by       VARCHAR(100) DEFAULT(SYSTEM_USER) NOT NULL,
    source_file        VARCHAR(255) NULL,
    record_count       INT NULL,
    notes              VARCHAR(500) NULL
);
GO

CREATE INDEX IX_submission_date ON dbo.tbl_submission_log (submission_date DESC);
GO

-- ============================================================================
-- SECTION 5: HELPER VIEWS
-- ============================================================================

-- View: Current working PIFs (what goes back to Excel next month)
IF OBJECT_ID('dbo.vw_pif_current_working', 'V') IS NOT NULL
    DROP VIEW dbo.vw_pif_current_working;
GO

CREATE VIEW dbo.vw_pif_current_working AS
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
    p.include_flag,
    c.scenario,
    c.year,
    c.requested_value,
    c.current_value,
    c.variance_value
FROM dbo.tbl_pif_projects_inflight p
LEFT JOIN dbo.tbl_pif_cost_inflight c
    ON p.pif_id = c.pif_id AND p.project_id = c.project_id
WHERE p.status NOT IN ('Approved', 'Dispositioned')
   OR p.status IS NULL;
GO

-- View: Full history (approved + inflight combined)
IF OBJECT_ID('dbo.vw_pif_all_history', 'V') IS NOT NULL
    DROP VIEW dbo.vw_pif_all_history;
GO

CREATE VIEW dbo.vw_pif_all_history AS
SELECT 
    'Approved' AS source,
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
    p.justification,
    c.scenario,
    c.year,
    c.requested_value,
    c.current_value,
    c.variance_value
FROM dbo.tbl_pif_projects_approved p
LEFT JOIN dbo.tbl_pif_cost_approved c
    ON p.pif_id = c.pif_id AND p.project_id = c.project_id

UNION ALL

SELECT 
    'Inflight' AS source,
    p.pif_id,
    p.project_id,
    p.submission_date,
    NULL AS approval_date,
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
    p.justification,
    c.scenario,
    c.year,
    c.requested_value,
    c.current_value,
    c.variance_value
FROM dbo.tbl_pif_projects_inflight p
LEFT JOIN dbo.tbl_pif_cost_inflight c
    ON p.pif_id = c.pif_id AND p.project_id = c.project_id;
GO

-- ============================================================================
-- SECTION 6: VALIDATION STORED PROCEDURES
-- ============================================================================

-- Procedure: Validate staging data before commit
IF OBJECT_ID('dbo.usp_validate_staging_data', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_validate_staging_data;
GO

CREATE PROCEDURE dbo.usp_validate_staging_data
    @ErrorCount INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @Errors TABLE (
        error_id INT IDENTITY(1,1),
        error_type VARCHAR(50),
        error_message VARCHAR(500)
    );
    
    -- Check 1: Missing required fields
    INSERT INTO @Errors (error_type, error_message)
    SELECT 
        'Missing PIF ID',
        'Row ' + CAST(pif_project_id AS VARCHAR) + ': Missing PIF ID'
    FROM dbo.tbl_pif_projects_staging
    WHERE pif_id IS NULL OR LTRIM(RTRIM(pif_id)) = '';
    
    INSERT INTO @Errors (error_type, error_message)
    SELECT 
        'Missing Project ID',
        'Row ' + CAST(pif_project_id AS VARCHAR) + ': Missing Project ID'
    FROM dbo.tbl_pif_projects_staging
    WHERE project_id IS NULL OR LTRIM(RTRIM(project_id)) = '';
    
    -- Check 2: Duplicate PIF + Project combinations
    INSERT INTO @Errors (error_type, error_message)
    SELECT 
        'Duplicate',
        'Duplicate: PIF ' + pif_id + ', Project ' + project_id + ' (appears ' + CAST(cnt AS VARCHAR) + ' times)'
    FROM (
        SELECT pif_id, project_id, COUNT(*) as cnt
        FROM dbo.tbl_pif_projects_staging
        GROUP BY pif_id, project_id
        HAVING COUNT(*) > 1
    ) dups;
    
    -- Check 3: Approved PIFs missing justification
    INSERT INTO @Errors (error_type, error_message)
    SELECT 
        'Missing Justification',
        'PIF ' + pif_id + ', Project ' + project_id + ': Approved but missing justification'
    FROM dbo.tbl_pif_projects_staging
    WHERE status = 'Approved'
      AND (justification IS NULL OR LTRIM(RTRIM(justification)) = '');
    
    -- Check 4: Cost records without matching project
    INSERT INTO @Errors (error_type, error_message)
    SELECT 
        'Orphan Cost Record',
        'Cost record for PIF ' + c.pif_id + ', Project ' + c.project_id + ' has no matching project record'
    FROM dbo.tbl_pif_cost_staging c
    LEFT JOIN dbo.tbl_pif_projects_staging p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id
    WHERE p.pif_id IS NULL;
    
    -- Return error count
    SELECT @ErrorCount = COUNT(*) FROM @Errors;
    
    -- Return error details
    SELECT 
        error_id,
        error_type,
        error_message
    FROM @Errors
    ORDER BY error_id;
    
    RETURN 0;
END;
GO

-- ============================================================================
-- SECTION 7: HELPER SCRIPTS FOR COMMON OPERATIONS
-- ============================================================================

-- Script: Create backup of inflight tables (run before each submission)
/*
DECLARE @BackupDate VARCHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112); -- YYYYMMDD

EXEC('SELECT * INTO dbo.tbl_pif_projects_inflight_backup_' + @BackupDate + ' FROM dbo.tbl_pif_projects_inflight');
EXEC('SELECT * INTO dbo.tbl_pif_cost_inflight_backup_' + @BackupDate + ' FROM dbo.tbl_pif_cost_inflight');

PRINT 'Backups created: tbl_pif_projects_inflight_backup_' + @BackupDate;
*/

-- Script: Archive approved PIFs to permanent tables
/*
BEGIN TRANSACTION;

    INSERT INTO dbo.tbl_pif_projects_approved (
        pif_id, project_id, submission_date, approval_date, status,
        change_type, accounting_treatment, category, seg, opco, site,
        strategic_rank, funding_project, project_name, original_fp_isd,
        revised_fp_isd, moving_isd_year, lcm_issue, justification,
        prior_year_spend, archive_flag, include_flag
    )
    SELECT 
        pif_id, project_id, submission_date, GETDATE(), status,
        change_type, accounting_treatment, category, seg, opco, site,
        strategic_rank, funding_project, project_name, original_fp_isd,
        revised_fp_isd, moving_isd_year, lcm_issue, justification,
        prior_year_spend, archive_flag, include_flag
    FROM dbo.tbl_pif_projects_inflight
    WHERE status IN ('Approved', 'Dispositioned');
    
    INSERT INTO dbo.tbl_pif_cost_approved (
        pif_id, project_id, scenario, year, 
        requested_value, current_value, variance_value, approval_date
    )
    SELECT 
        c.pif_id, c.project_id, c.scenario, c.year,
        c.requested_value, c.current_value, c.variance_value, GETDATE()
    FROM dbo.tbl_pif_cost_inflight c
    INNER JOIN dbo.tbl_pif_projects_inflight p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id
    WHERE p.status IN ('Approved', 'Dispositioned');
    
    -- Optional: Remove approved records from inflight
    DELETE c FROM dbo.tbl_pif_cost_inflight c
    INNER JOIN dbo.tbl_pif_projects_inflight p
        ON c.pif_id = p.pif_id AND c.project_id = p.project_id
    WHERE p.status IN ('Approved', 'Dispositioned');
    
    DELETE FROM dbo.tbl_pif_projects_inflight
    WHERE status IN ('Approved', 'Dispositioned');

COMMIT TRANSACTION;
*/

-- ============================================================================
-- SETUP COMPLETE
-- ============================================================================
PRINT 'PIF Database setup complete.';
PRINT 'Tables created:';
PRINT '  - Staging: tbl_pif_projects_staging, tbl_pif_cost_staging';
PRINT '  - Inflight: tbl_pif_projects_inflight, tbl_pif_cost_inflight';
PRINT '  - Approved: tbl_pif_projects_approved, tbl_pif_cost_approved';
PRINT '  - Audit: tbl_submission_log';
PRINT 'Views created:';
PRINT '  - vw_pif_current_working';
PRINT '  - vw_pif_all_history';
PRINT 'Stored Procedures:';
PRINT '  - usp_validate_staging_data';
GO

-- ============================================================================
-- SECTION 8: STAGING STORED PROCEDURES
-- ============================================================================

IF OBJECT_ID('dbo.usp_insert_project_staging', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_insert_project_staging;
GO

CREATE PROCEDURE dbo.usp_insert_project_staging
    @pif_id             VARCHAR(16),
    @project_id         VARCHAR(10),
    @status             VARCHAR(58),
    @change_type        VARCHAR(12),
    @accounting_treatment VARCHAR(14),
    @category           VARCHAR(26),
    @seg                INT,
    @opco               VARCHAR(4),
    @site               VARCHAR(4),
    @strategic_rank     VARCHAR(26),
    @funding_project    VARCHAR(10),
    @project_name       VARCHAR(35),
    @original_fp_isd    VARCHAR(8),
    @revised_fp_isd     VARCHAR(5),
    @moving_isd_year    CHAR(1),
    @lcm_issue          VARCHAR(11),
    @justification      VARCHAR(192),
    @prior_year_spend   DECIMAL(18,2),
    @archive_flag       BIT,
    @include_flag       BIT
AS
BEGIN
    SET NOCOUNT ON;
    
    INSERT INTO dbo.tbl_pif_projects_staging (
        pif_id,
        project_id,
        status,
        change_type,
        accounting_treatment,
        category,
        seg,
        opco,
        site,
        strategic_rank,
        funding_project,
        project_name,
        original_fp_isd,
        revised_fp_isd,
        moving_isd_year,
        lcm_issue,
        justification,
        prior_year_spend,
        archive_flag,
        include_flag
    )
    VALUES (
        @pif_id,
        @project_id,
        @status,
        @change_type,
        @accounting_treatment,
        @category,
        @seg,
        @opco,
        @site,
        @strategic_rank,
        @funding_project,
        @project_name,
        @original_fp_isd,
        @revised_fp_isd,
        @moving_isd_year,
        @lcm_issue,
        @justification,
        @prior_year_spend,
        @archive_flag,
        @include_flag
    );
END;
GO

IF OBJECT_ID('dbo.usp_insert_cost_staging', 'P') IS NOT NULL
    DROP PROCEDURE dbo.usp_insert_cost_staging;
GO

CREATE PROCEDURE dbo.usp_insert_cost_staging
    @pif_id             VARCHAR(16),
    @project_id         VARCHAR(10),
    @scenario           VARCHAR(12),
    @year               DATE,
    @requested_value    DECIMAL(18,2),
    @current_value      DECIMAL(18,2),
    @variance_value     DECIMAL(18,2)
AS
BEGIN
    SET NOCOUNT ON;
    
    INSERT INTO dbo.tbl_pif_cost_staging (
        pif_id,
        project_id,
        scenario,
        year,
        requested_value,
        current_value,
        variance_value
    )
    VALUES (
        @pif_id,
        @project_id,
        @scenario,
        @year,
        @requested_value,
        @current_value,
        @variance_value
    );
END;
GO
