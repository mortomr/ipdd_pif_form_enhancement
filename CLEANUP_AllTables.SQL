-- ============================================================================
-- CLEANUP: Delete all records from staging, inflight, and approved tables
-- CLEANUP_AllTables.sql
-- ============================================================================
-- WARNING: This deletes ALL data from these tables
-- Use only for testing/cleanup purposes
-- ============================================================================

USE [IPDD];
GO

PRINT 'Deleting all records from staging tables...';
DELETE FROM dbo.tbl_pif_cost_staging;
DELETE FROM dbo.tbl_pif_projects_staging;
PRINT 'Staging tables cleared.';
PRINT '';

PRINT 'Deleting all records from inflight tables...';
DELETE FROM dbo.tbl_pif_cost_inflight;
DELETE FROM dbo.tbl_pif_projects_inflight;
PRINT 'Inflight tables cleared.';
PRINT '';

PRINT 'Deleting all records from approved tables...';
DELETE FROM dbo.tbl_pif_cost_approved;
DELETE FROM dbo.tbl_pif_projects_approved;
PRINT 'Approved tables cleared.';
PRINT '';

-- Reset identity seeds (optional - for cleaner IDs after cleanup)
PRINT 'Resetting identity seeds...';
DBCC CHECKIDENT ('dbo.tbl_pif_projects_staging', RESEED, 0);
DBCC CHECKIDENT ('dbo.tbl_pif_cost_staging', RESEED, 0);
DBCC CHECKIDENT ('dbo.tbl_pif_projects_inflight', RESEED, 0);
DBCC CHECKIDENT ('dbo.tbl_pif_cost_inflight', RESEED, 0);
DBCC CHECKIDENT ('dbo.tbl_pif_projects_approved', RESEED, 0);
DBCC CHECKIDENT ('dbo.tbl_pif_cost_approved', RESEED, 0);
PRINT 'Identity seeds reset.';
PRINT '';

PRINT '============================================================';
PRINT 'CLEANUP COMPLETE - All staging/inflight/approved data deleted';
PRINT '============================================================';
GO
