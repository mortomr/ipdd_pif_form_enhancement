# Archive Cleanup Feature - Setup Instructions

## Overview

The Archive Cleanup feature allows site users to delete archived records from the PIF data entry worksheet after they've been successfully archived to the database. This prepares the worksheet for next month's data entry.

**Key Points:**
- **Enabled for:** Site contexts (ANO, GGN, RBS, WF3, HQN)
- **Disabled for:** Fleet context (read-only view)
- **Safety:** Uses database validation to ensure only archived records are deleted
- **Location:** Button should be placed on the `PIF_Archive` worksheet

## Workflow

1. User finalizes month and clicks **[Submit to Database]**
2. User navigates to **PIF_Archive** sheet to confirm records are captured
3. User clicks **[Delete Archived Records]** button
4. System queries database to validate which records are archived for current site
5. System displays confirmation dialog with count of records to delete
6. Upon confirmation, matching records are deleted from PIF worksheet
7. PIF worksheet is now clean and ready for next month's data entry

## Implementation Steps

### Step 1: Import the Module (if not already done)

The new module `mod_ArchiveCleanup.bas` should be imported into the Excel workbook:

1. Open the Excel workbook with VBA editor (Alt+F11)
2. Go to File > Import File
3. Select `mod_ArchiveCleanup.bas`
4. Click Open

### Step 2: Add Button to PIF_Archive Sheet

**Manual Button Creation:**

1. Open the Excel workbook
2. Navigate to the **PIF_Archive** worksheet
3. Go to **Developer** tab > **Insert** > **Button (Form Control)**
4. Draw the button on the worksheet (suggested location: top-right area, near refresh instructions)
5. In the "Assign Macro" dialog, select: **Archive_DeleteArchivedRecords**
6. Click OK
7. Right-click the button and choose **Edit Text**
8. Enter button text: **Delete Archived Records**

**Recommended Button Formatting:**
- Font: Bold, 11pt
- Size: Width ~200px, Height ~30px
- Color: Light red or orange background (to indicate caution)
- Location: Row 3, near column N-P (to the right of title)

### Step 3: Add Instructional Text (Optional)

Consider adding text near the button to explain its purpose:

**Suggested text in cell (e.g., B3):**
```
CLEANUP: After confirming archived records, click [Delete Archived Records] to remove them from PIF worksheet (SITE contexts only)
```

Format this text:
- Font Size: 9pt
- Font Color: Dark blue or gray
- Italic: Yes

### Step 4: Test the Functionality

**Test Case 1: Site Context (Should Work)**
1. Select a site (e.g., ANO) from the Instructions sheet
2. Navigate to PIF_Archive sheet
3. Click **[Delete Archived Records]** button
4. Verify:
   - Database query runs successfully
   - Confirmation dialog shows correct count
   - Records are deleted from PIF sheet after confirmation
   - Success message displays

**Test Case 2: Fleet Context (Should Be Blocked)**
1. Select "Fleet" from the Instructions sheet
2. Navigate to PIF_Archive sheet
3. Click **[Delete Archived Records]** button
4. Verify:
   - Error message appears: "Archive cleanup is disabled for Fleet context"
   - No records are deleted

**Test Case 3: No Archived Records (Should Inform)**
1. Select a site with no archived records
2. Click **[Delete Archived Records]** button
3. Verify:
   - Information message: "No archived records found in the database"

**Test Case 4: Already Cleaned Up (Should Inform)**
1. Run cleanup once successfully
2. Click **[Delete Archived Records]** button again
3. Verify:
   - Information message: "No matching archived records found in the PIF worksheet"

## Technical Details

### Module: mod_ArchiveCleanup.bas

**Public Functions:**

1. **Archive_DeleteArchivedRecords()**
   - Main entry point attached to button
   - Validates site context (blocks Fleet)
   - Queries database for archived records
   - Matches records in PIF worksheet
   - Deletes matching records with confirmation

2. **IsArchiveCleanupEnabled() As Boolean**
   - Helper function to check if cleanup is enabled
   - Returns True for site contexts, False for Fleet
   - Can be used to conditionally show/hide button (future enhancement)

### Database Query

The module queries `tbl_pif_projects_approved` table:

```sql
SELECT DISTINCT pif_id, project_id
FROM dbo.tbl_pif_projects_approved
WHERE UPPER(site) = 'SELECTED_SITE'
  AND status IN ('Approved','Dispositioned')
ORDER BY pif_id, project_id
```

### Matching Logic

Records are matched using composite key: `PIF_ID|Project_ID`

The module:
1. Builds a collection of archived keys from database
2. Scans PIF worksheet rows (Column G = PIF_ID, Column M = Project_ID)
3. Validates site matches (safety check)
4. Collects row numbers for deletion
5. Deletes from bottom to top to maintain row indices

### Safety Features

1. **Site Validation:** Only deletes records matching current site
2. **Database Validation:** Only deletes records confirmed as archived
3. **Confirmation Dialog:** Shows count and asks for explicit confirmation
4. **Fleet Protection:** Completely disabled for Fleet context
5. **Error Handling:** Graceful error messages, no partial deletions on error
6. **Table-Safe Deletion:** Handles both Excel Tables (ListObject) and regular ranges

## Troubleshooting

### Button doesn't appear or isn't clickable
- Ensure worksheet protection is off or button is not locked
- Check that Developer tab is enabled
- Verify macro security settings allow macros to run

### "Site Not Selected" error
- User must select a site from Instructions sheet before running cleanup
- Verify named range "SelectedSite" exists

### "Cannot find PIF worksheet" error
- Verify SHEET_DATA constant in mod_SharedConstants matches actual sheet name
- Default is "PIF" - update if your worksheet has a different name

### "Database connection failed" error
- Check SQL_SERVER and SQL_DATABASE constants in mod_Database
- Verify network connectivity to SQL Server
- Ensure Windows Authentication is enabled

### No records deleted (but expected some)
- Verify records have been submitted and archived (status = Approved/Dispositioned)
- Check that PIF_ID and Project_ID values match exactly between worksheet and database
- Ensure site values match (case-insensitive comparison used)
- Refresh Archive sheet to verify records are actually archived

## Future Enhancements

1. **Dynamic Button State:** Use Workbook_SheetActivate event to enable/disable button based on site context
2. **Preview Mode:** Show list of records to be deleted before confirmation
3. **Undo Capability:** Create backup sheet or log deleted records for recovery
4. **Audit Log:** Log cleanup operations to database table
5. **Batch Operations:** Allow cleanup across multiple sites (Fleet admin only)

## Related Files

- **mod_ArchiveCleanup.bas** - Main cleanup module
- **mod_SiteSetup.bas** - Site selection and context functions
- **mod_Database.bas** - Database connection functions
- **mod_SharedConstants.bas** - Column mappings and constants
- **mod_DataEntry.bas** - Row deletion utilities (referenced for table handling)

## Support

For questions or issues, contact the Data Architecture Team or submit an issue in the project repository.

---

**Document Version:** 1.0.0
**Date:** 2025-11-13
**Author:** Data Architecture Team
