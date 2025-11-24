-- ============================================================================
-- PIF DATABASE MIGRATION: Add line_item field
-- ============================================================================
-- Purpose: Allow multiple detail lines for same PIF_ID + Project_ID combination
-- Author: Data Architecture Team
-- Date: 2025-11-20
--
-- Use Case: A single PIF+Project may have multiple rows with different
--           justifications (e.g., "low payroll estimate" + "emergent outage")
--           Each row gets a line_item number (1, 2, 3, etc.)
--
-- New Composite Key: PIF_ID + Project_ID + line_item
-- ============================================================================

USE [IPDD];
GO

PRINT 'Starting migration: Adding line_item field...';
PRINT '';

-- ============================================================================
-- STEP 1: Add line_item column to all tables
-- ============================================================================

PRINT 'Adding line_item column to staging tables...';

-- Staging: Projects
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.tbl_pif_projects_staging') AND name = 'line_item')
BEGIN
    ALTER TABLE dbo.tbl_pif_projects_staging
    ADD line_item INT NOT NULL DEFAULT 1;
    PRINT '  ✓ tbl_pif_projects_staging';
END
ELSE
    PRINT '  - tbl_pif_projects_staging (already exists)';

-- Staging: Costs
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.tbl_pif_cost_staging') AND name = 'line_item')
BEGIN
    ALTER TABLE dbo.tbl_pif_cost_staging
    ADD line_item INT NOT NULL DEFAULT 1;
    PRINT '  ✓ tbl_pif_cost_staging';
END
ELSE
    PRINT '  - tbl_pif_cost_staging (already exists)';

PRINT '';
PRINT 'Adding line_item column to inflight tables...';

-- Inflight: Projects (need to drop UNIQUE constraint first)
IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE name = 'UQ_inflight_pif_project')
BEGIN
    ALTER TABLE dbo.tbl_pif_projects_inflight
    DROP CONSTRAINT UQ_inflight_pif_project;
    PRINT '  ✓ Dropped old UNIQUE constraint';
END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.tbl_pif_projects_inflight') AND name = 'line_item')
BEGIN
    ALTER TABLE dbo.tbl_pif_projects_inflight
    ADD line_item INT NOT NULL DEFAULT 1;
    PRINT '  ✓ tbl_pif_projects_inflight';
END
ELSE
    PRINT '  - tbl_pif_projects_inflight (already exists)';

-- Inflight: Costs
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.tbl_pif_cost_inflight') AND name = 'line_item')
BEGIN
    ALTER TABLE dbo.tbl_pif_cost_inflight
    ADD line_item INT NOT NULL DEFAULT 1;
    PRINT '  ✓ tbl_pif_cost_inflight';
END
ELSE
    PRINT '  - tbl_pif_cost_inflight (already exists)';

PRINT '';
PRINT 'Adding line_item column to approved tables...';

-- Approved: Projects
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.tbl_pif_projects_approved') AND name = 'line_item')
BEGIN
    ALTER TABLE dbo.tbl_pif_projects_approved
    ADD line_item INT NOT NULL DEFAULT 1;
    PRINT '  ✓ tbl_pif_projects_approved';
END
ELSE
    PRINT '  - tbl_pif_projects_approved (already exists)';

-- Approved: Costs
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.tbl_pif_cost_approved') AND name = 'line_item')
BEGIN
    ALTER TABLE dbo.tbl_pif_cost_approved
    ADD line_item INT NOT NULL DEFAULT 1;
    PRINT '  ✓ tbl_pif_cost_approved';
END
ELSE
    PRINT '  - tbl_pif_cost_approved (already exists)';

-- ============================================================================
-- STEP 2: Update UNIQUE constraint on inflight table
-- ============================================================================

PRINT '';
PRINT 'Creating new UNIQUE constraint (PIF + Project + LineItem)...';

IF NOT EXISTS (SELECT 1 FROM sys.key_constraints WHERE name = 'UQ_inflight_pif_project_line')
BEGIN
    ALTER TABLE dbo.tbl_pif_projects_inflight
    ADD CONSTRAINT UQ_inflight_pif_project_line UNIQUE (pif_id, project_id, line_item);
    PRINT '  ✓ Created UQ_inflight_pif_project_line';
END
ELSE
    PRINT '  - UQ_inflight_pif_project_line (already exists)';

-- ============================================================================
-- STEP 3: Update indexes to include line_item
-- ============================================================================

PRINT '';
PRINT 'Updating indexes to include line_item...';

-- Approved projects index
IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_approved_pif_project' AND object_id = OBJECT_ID('dbo.tbl_pif_projects_approved'))
BEGIN
    DROP INDEX IX_approved_pif_project ON dbo.tbl_pif_projects_approved;
    PRINT '  - Dropped old IX_approved_pif_project';
END

CREATE NONCLUSTERED INDEX IX_approved_pif_project
    ON dbo.tbl_pif_projects_approved (pif_id, project_id, line_item);
PRINT '  ✓ Created new IX_approved_pif_project (includes line_item)';

-- Inflight cost index
IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_inflight_cost_lookup' AND object_id = OBJECT_ID('dbo.tbl_pif_cost_inflight'))
BEGIN
    DROP INDEX IX_inflight_cost_lookup ON dbo.tbl_pif_cost_inflight;
    PRINT '  - Dropped old IX_inflight_cost_lookup';
END

CREATE NONCLUSTERED INDEX IX_inflight_cost_lookup
    ON dbo.tbl_pif_cost_inflight (pif_id, project_id, line_item, scenario, year);
PRINT '  ✓ Created new IX_inflight_cost_lookup (includes line_item)';

-- Approved cost index
IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_approved_cost_lookup' AND object_id = OBJECT_ID('dbo.tbl_pif_cost_approved'))
BEGIN
    DROP INDEX IX_approved_cost_lookup ON dbo.tbl_pif_cost_approved;
    PRINT '  - Dropped old IX_approved_cost_lookup';
END

CREATE NONCLUSTERED INDEX IX_approved_cost_lookup
    ON dbo.tbl_pif_cost_approved (pif_id, project_id, line_item, scenario, year);
PRINT '  ✓ Created new IX_approved_cost_lookup (includes line_item)';

PRINT '';
PRINT '============================================================================';
PRINT 'Migration complete!';
PRINT '';
PRINT 'NEXT STEPS:';
PRINT '  1. Update stored procedures (see PIF_Database_Update_SPs_LineItem.sql)';
PRINT '  2. Update VBA code (mod_Validation.bas, mod_Submit.bas)';
PRINT '  3. Add line_item column to Excel PIF sheet';
PRINT '';
PRINT 'IMPORTANT: Existing data now has line_item = 1 (default)';
PRINT '============================================================================';
GO
