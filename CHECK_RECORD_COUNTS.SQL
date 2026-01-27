
    SELECT COUNT(*) as [record count] , '[tbl_pif_projects_staging]' as tablename
    FROM [tbl_pif_projects_staging]
UNION ALL
    SELECT COUNT(*) , '[tbl_pif_projects_inflight]'
    FROM [tbl_pif_projects_inflight]
UNION ALL
    SELECT COUNT(*) , '[tbl_pif_projects_approved]'
    FROM [tbl_pif_projects_approved]
UNION ALL
    SELECT COUNT(*) , '[tbl_pif_cost_staging]'
    FROM [tbl_pif_cost_staging]
UNION ALL
    SELECT COUNT(*) , '[tbl_pif_cost_inflight]'
    FROM [tbl_pif_cost_inflight]
UNION ALL
    SELECT COUNT(*) , '[tbl_pif_cost_approved]'
    FROM [tbl_pif_cost_approved]

SELECT *
FROM [tbl_pif_cost_inflight]
SELECT *
FROM [tbl_pif_cost_approved]
SELECT *
FROM [tbl_pif_projects_inflight]