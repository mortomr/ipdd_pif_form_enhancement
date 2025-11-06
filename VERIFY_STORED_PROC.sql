-- ============================================================================
-- VERIFY STORED PROCEDURES EXIST
-- ============================================================================
-- Run this on your SQL Server to check if procedures exist and their parameters
-- Server: LITDBETST012\TVNN160002
-- Database: IPDD
-- ============================================================================

USE [IPDD];
GO

-- Check if procedures exist
SELECT
    name AS ProcedureName,
    create_date,
    modify_date
FROM sys.procedures
WHERE name IN ('usp_insert_project_staging', 'usp_insert_cost_staging')
ORDER BY name;

If they exist, show their parameters
SELECT
    OBJECT_NAME(object_id) AS ProcedureName,
    parameter_id,
    name AS ParameterName,
    TYPE_NAME(user_type_id) AS DataType,
    max_length,
    is_output
FROM sys.parameters
WHERE OBJECT_NAME(object_id) IN ('usp_insert_project_staging', 'usp_insert_cost_staging')
ORDER BY OBJECT_NAME(object_id), parameter_id;

-- Expected result for usp_insert_project_staging should have 20 parameters:
-- @pif_id (varchar)
-- @project_id (varchar)
-- @status (varchar)
-- @change_type (varchar)
-- @accounting_treatment (varchar)
-- @category (varchar)
-- @seg (int)
-- @opco (varchar)
-- @site (varchar)
-- @strategic_rank (varchar)
-- @funding_project (varchar)
-- @project_name (varchar)
-- @original_fp_isd (varchar)
-- @revised_fp_isd (varchar)
-- @moving_isd_year (char)
-- @lcm_issue (varchar)
-- @justification (varchar)
-- @prior_year_spend (decimal)
-- @archive_flag (bit)
-- @include_flag (bit)
