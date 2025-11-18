-- ============================================================================
-- DROP NON-CONFORMING DATABASE OBJECTS
-- ============================================================================
-- Purpose: Remove database objects that don't follow the vw_pif_* naming convention
-- Author: Data Architecture Team
-- Date: 2025-11-11
-- Version: 1.0.0
--
-- This script drops views that were created without the consistent PIF prefix.
-- These views will be recreated with proper naming in the main DDL script.
--
-- IMPORTANT: Run this script BEFORE running the updated PIF_Database_DDL.sql
-- ============================================================================

USE [IPDD];
GO

PRINT '';
PRINT '============================================================================';
PRINT 'Dropping non-conforming views (missing vw_pif_* prefix)';
PRINT '============================================================================';
PRINT '';

-- ----------------------------------------------------------------------------
-- DROP VIEW: vw_inflight_by_site
-- Reason: Should be vw_pif_inflight_by_site for consistent grouping
-- ----------------------------------------------------------------------------
IF OBJECT_ID('dbo.vw_inflight_by_site', 'V') IS NOT NULL
BEGIN
    DROP VIEW dbo.vw_inflight_by_site;
    PRINT 'DROPPED: dbo.vw_inflight_by_site';
END
ELSE
BEGIN
    PRINT 'SKIPPED: dbo.vw_inflight_by_site (does not exist)';
END
GO

-- ----------------------------------------------------------------------------
-- DROP VIEW: vw_approved_by_site
-- Reason: Should be vw_pif_approved_by_site for consistent grouping
-- ----------------------------------------------------------------------------
IF OBJECT_ID('dbo.vw_approved_by_site', 'V') IS NOT NULL
BEGIN
    DROP VIEW dbo.vw_approved_by_site;
    PRINT 'DROPPED: dbo.vw_approved_by_site';
END
ELSE
BEGIN
    PRINT 'SKIPPED: dbo.vw_approved_by_site (does not exist)';
END
GO

-- ----------------------------------------------------------------------------
-- DROP VIEW: vw_inflight_wide
-- Reason: Should be vw_pif_inflight_wide for consistent grouping
-- ----------------------------------------------------------------------------
IF OBJECT_ID('dbo.vw_inflight_wide', 'V') IS NOT NULL
BEGIN
    DROP VIEW dbo.vw_inflight_wide;
    PRINT 'DROPPED: dbo.vw_inflight_wide';
END
ELSE
BEGIN
    PRINT 'SKIPPED: dbo.vw_inflight_wide (does not exist)';
END
GO

-- ----------------------------------------------------------------------------
-- DROP VIEW: vw_approved_wide
-- Reason: Should be vw_pif_approved_wide for consistent grouping
-- ----------------------------------------------------------------------------
IF OBJECT_ID('dbo.vw_approved_wide', 'V') IS NOT NULL
BEGIN
    DROP VIEW dbo.vw_approved_wide;
    PRINT 'DROPPED: dbo.vw_approved_wide';
END
ELSE
BEGIN
    PRINT 'SKIPPED: dbo.vw_approved_wide (does not exist)';
END
GO

PRINT '';
PRINT '============================================================================';
PRINT 'Non-conforming objects cleanup complete!';
PRINT '';
PRINT 'NEXT STEPS:';
PRINT '  1. Run the updated PIF_Database_DDL.sql script';
PRINT '  2. Verify all views now have the vw_pif_* prefix';
PRINT '  3. Update VBA code to reference the renamed views';
PRINT '';
PRINT 'Expected new view names:';
PRINT '  - vw_pif_approved_by_site';
PRINT '  - vw_pif_approved_wide';
PRINT '  - vw_pif_all_history';
PRINT '  - vw_pif_current_working';
PRINT '  - vw_pif_inflight_by_site';
PRINT '  - vw_pif_inflight_wide';
PRINT '============================================================================';
GO
