# Macro Naming Convention

## Overview

All public macros now use a prefix-based naming convention for better organization and discoverability in Excel's macro list.

## Prefix Groups

| Prefix | Category | Purpose |
|--------|----------|---------|
| `DB_` | Database Operations | Submit, validate, archive, finalize |
| `Nav_` | Navigation/Queries | Refresh worksheets and views |
| `Edit_` | Data Entry | Add/delete rows, modify data |
| `Tool_` | Utilities | Highlighting, validation helpers |
| `Setup_` | Configuration | One-time setup operations |
| `Diag_` | Diagnostics | Testing and troubleshooting |

## Complete Function Reference

### Database Operations (DB_)

| New Name | Old Name | Description | Button Assignment |
|----------|----------|-------------|-------------------|
| `DB_SaveSnapshot` | `SaveSnapshot` | Save working snapshot to inflight | [Save Snapshot] |
| `DB_FinalizeMonth` | `FinalizeMonth` | Complete month-end with archival | [Finalize Month] |
| `DB_ValidateOnly` | `RunValidationOnly` | Run validation without submit | [Validate] |
| `DB_ArchiveApproved` | `ArchiveApprovedRecords` | Archive approved records | [Archive] |
| `DB_Submit` | `SubmitToDatabase` | Legacy submit (calls SaveSnapshot) | - |

### Navigation/Queries (Nav_)

| New Name | Old Name | Description | Button Assignment |
|----------|----------|-------------|-------------------|
| `Nav_RefreshArchive` | `RefreshArchiveWorksheet` | Refresh Archive worksheet | [Refresh Archive] |
| `Nav_RefreshInflight` | `RefreshInflightWorksheet` | Refresh Inflight worksheet | [Refresh Inflight] |
| `Nav_RefreshAll` | `RefreshBothWorksheets` | Refresh both worksheets | [Refresh All] |

### Data Entry (Edit_)

| New Name | Old Name | Description | Button Assignment |
|----------|----------|-------------|-------------------|
| `Edit_AddRow` | `AddRow` | Insert new data row | [Add Row] |
| `Edit_DeleteRows` | `DeleteSelectedRows` | Delete selected rows | [Delete] |

### Utilities (Tool_)

| New Name | Old Name | Description | Button Assignment |
|----------|----------|-------------|-------------------|
| `Tool_HighlightIncomplete` | `HighlightIncompleteRows` | Highlight incomplete rows | [Highlight] |
| `Tool_ClearHighlights` | `ClearRowHighlights` | Clear row highlights | [Clear] |

### Setup (Setup_)

| New Name | Old Name | Description | Button Assignment |
|----------|----------|-------------|-------------------|
| `Setup_SiteSelection` | `SetupSiteSelection` | Create site selection UI | Run once manually |

### Diagnostics (Diag_)

| New Name | Old Name | Description | Button Assignment |
|----------|----------|-------------|-------------------|
| `Diag_TestInsert` | `TestSingleRowInsert` | Test single row insert | Developer use |
| `Diag_ShowErrors` | `ShowValidationErrors` | Display validation errors | Developer use |
| `Diag_TestParameters` | `TestParameterTypes` | Test parameter conversions | Developer use |

## Backward Compatibility

All old function names remain available as wrappers that call the new functions. This ensures:
- ✅ Existing button assignments continue to work
- ✅ No need to update button macros immediately
- ✅ Gradual migration path

**Example:**
```vba
' Old name still works (calls new function internally)
Public Sub AddRow()
    Call Edit_AddRow
End Sub
```

## Migration Guide

### For Button Updates

To update a button to use the new naming:

1. Right-click the button → **Assign Macro**
2. Select the new function name (e.g., `Edit_AddRow` instead of `AddRow`)
3. Click **OK**

### Benefits of New Naming

✅ **Alphabetical grouping** - Related functions appear together in macro list
✅ **Clear categorization** - Know the function's purpose at a glance
✅ **Shorter names** - More concise and easier to type
✅ **Better organization** - Logical grouping improves maintainability

## Module Organization

| Module | Primary Prefix | Functions |
|--------|---------------|-----------|
| `mod_Submit.bas` | `DB_` | 5 functions |
| `mod_WorksheetQuery.bas` | `Nav_` | 3 functions |
| `mod_DataEntry.bas` | `Edit_`, `Tool_` | 4 functions |
| `mod_SiteSetup.bas` | `Setup_` | 1 function |
| `mod_Diagnostic.bas` | `Diag_` | 3 functions |

---

*Last updated: 2025-11-11*
