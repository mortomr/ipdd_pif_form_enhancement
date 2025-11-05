# PIF System Configuration Document

**Date Configured**: ___________________
**Configured By**: ___________________
**Version**: 1.0

---

## SQL Server Configuration

### Database Connection Details

| Setting | Value |
|---------|-------|
| SQL Server Name | _________________________ |
| SQL Server Instance | _________________________ |
| Database Name | _________________________ |
| Authentication Method | ☐ Windows Auth  ☐ SQL Auth |
| Port (if non-standard) | _________________________ |

### Database Objects Created

Verify these objects exist after running DDL script:

**Tables:**
- [ ] tbl_pif_projects_staging
- [ ] tbl_pif_cost_staging
- [ ] tbl_pif_projects_inflight
- [ ] tbl_pif_cost_inflight
- [ ] tbl_pif_projects_approved
- [ ] tbl_pif_cost_approved
- [ ] tbl_submission_log

**Views:**
- [ ] vw_pif_current_working
- [ ] vw_pif_all_history

**Stored Procedures:**
- [ ] usp_validate_staging_data

**Indexes:**
- [ ] All indexes created successfully (14 total)

### Test Queries Run

- [ ] SELECT COUNT(*) from each table (should return 0 initially)
- [ ] Test view queries
- [ ] Test stored procedure execution

---

## Excel VBA Configuration

### Workbook Details

| Setting | Value |
|---------|-------|
| Workbook Name | _________________________ |
| File Location | _________________________ |
| Template Location | _________________________ |

### VBA Modules Imported

- [ ] mod_Database.bas
- [ ] mod_Validation.bas  
- [ ] mod_Submit.bas

### ADODB Reference Added

- [ ] Microsoft ActiveX Data Objects 6.1 Library checked

### Connection Constants Updated

**In mod_Database.bas:**

```vba
Line 21: Private Const SQL_SERVER As String = "_________________"
Line 22: Private Const SQL_DATABASE As String = "_________________"
Line 23: Private Const SQL_TRUSTED As Boolean = True
```

### Sheet Names Verified

**Required Sheets:**

| Sheet Name in Code | Actual Sheet Name | Status |
|-------------------|-------------------|---------|
| PIF_Data | _________________ | ☐ Match ☐ Need to update code |
| Validation_Report | _________________ | ☐ Match ☐ Auto-created |
| Cost_Unpivoted | _________________ | ☐ Match ☐ Auto-created |

### Column Mappings Verified

**In mod_Validation.bas (lines 15-29):**

| Constant Name | Expected Column | Actual Column | Status |
|--------------|----------------|---------------|---------|
| COL_PIF_ID | G (7) | ____ | ☐ Match ☐ Updated |
| COL_FUNDING_PROJECT | M (13) | ____ | ☐ Match ☐ Updated |
| COL_STATUS | R (18) | ____ | ☐ Match ☐ Updated |
| COL_JUSTIFICATION | T (20) | ____ | ☐ Match ☐ Updated |

**In mod_Submit.bas (UnpivotCostData function):**

Target Scenario Columns:
- [ ] CY (2025): U, AA, AG
- [ ] CY+1 (2026): V, AB, AH
- [ ] CY+2 (2027): W, AC, AI
- [ ] CY+3 (2028): X, AD, AJ
- [ ] CY+4 (2029): Y, AE, AK
- [ ] CY+5 (2030): Z, AF, AL

Closings Scenario Columns:
- [ ] CY (2025): AO, AU, BA
- [ ] CY+1 (2026): AP, AV, BB
- [ ] CY+2 (2027): AQ, AW, BC
- [ ] CY+3 (2028): AR, AX, BD
- [ ] CY+4 (2029): AS, AY, BE
- [ ] CY+5 (2030): AT, AZ, BF

### Buttons Created

- [ ] "Validate Data" button → Macro: mod_Submit.RunValidationOnly
- [ ] "Submit to Database" button → Macro: mod_Submit.SubmitToDatabase

---

## Testing Results

### Connection Test

Date Tested: ___________________

- [ ] TestConnection() macro ran successfully
- [ ] Confirmed server name displayed
- [ ] Confirmed database name displayed
- [ ] Confirmed user name displayed

### Validation Test

Date Tested: ___________________

Sample Data Records: _______

- [ ] Validation ran without VBA errors
- [ ] Validation_Report sheet created
- [ ] Errors detected correctly (if test data had errors)
- [ ] No false positives

### Full Submission Test (Dev/Test Environment)

Date Tested: ___________________

- [ ] Unpivot process completed
- [ ] Backup tables created
- [ ] Data uploaded to staging
- [ ] Validation passed
- [ ] Data committed to inflight
- [ ] Submission logged
- [ ] Verified data in SQL

**Time Elapsed**: ________ seconds

**Issues Encountered**:
- ________________________________________________________________
- ________________________________________________________________
- ________________________________________________________________

**Resolution**:
- ________________________________________________________________
- ________________________________________________________________
- ________________________________________________________________

---

## Production Deployment

### Pre-Deployment Checklist

- [ ] All testing completed successfully in dev/test
- [ ] Database backup completed
- [ ] Excel template tested with real data (small sample)
- [ ] User training completed
- [ ] Documentation distributed
- [ ] Support contacts established

### Deployment Date: ___________________

**Deployed By**: ___________________

### Post-Deployment Verification

- [ ] First production submission successful
- [ ] Data verified in SQL
- [ ] No performance issues
- [ ] Users report no issues

---

## User Access

### Database Access

List users who have been granted access:

| User Name | Role | Permissions | Date Granted |
|-----------|------|-------------|--------------|
| _______________ | Submitter | Read/Write Inflight | __________ |
| _______________ | Submitter | Read/Write Inflight | __________ |
| _______________ | Analyst | Read Approved | __________ |
| _______________ | DBA | Full Access | __________ |

### Excel Access

List users who have the macro-enabled workbook:

| User Name | Site | Email | Date Trained |
|-----------|------|-------|--------------|
| _______________ | _____ | _______________ | __________ |
| _______________ | _____ | _______________ | __________ |
| _______________ | _____ | _______________ | __________ |
| _______________ | _____ | _______________ | __________ |

---

## Customizations Made

### Database Customizations

| Item | Default | Customized To | Reason |
|------|---------|---------------|--------|
| ____________ | __________ | _____________ | __________________ |
| ____________ | __________ | _____________ | __________________ |
| ____________ | __________ | _____________ | __________________ |

### VBA Customizations

| Module | Function | Change Made | Reason |
|--------|----------|-------------|--------|
| ____________ | __________ | _____________ | __________________ |
| ____________ | __________ | _____________ | __________________ |
| ____________ | __________ | _____________ | __________________ |

---

## Known Issues

### Current Issues

| Issue # | Description | Workaround | Status | Target Resolution |
|---------|-------------|------------|--------|-------------------|
| 1 | _____________ | _____________ | ☐ Open ☐ Resolved | __________ |
| 2 | _____________ | _____________ | ☐ Open ☐ Resolved | __________ |
| 3 | _____________ | _____________ | ☐ Open ☐ Resolved | __________ |

---

## Maintenance Schedule

### Regular Maintenance Tasks

| Task | Frequency | Responsible | Last Completed | Next Due |
|------|-----------|-------------|----------------|----------|
| Drop old backup tables | Monthly | DBA | __________ | __________ |
| Review submission log | Monthly | PIF Owner | __________ | __________ |
| Rebuild indexes | Quarterly | DBA | __________ | __________ |
| Update statistics | Quarterly | DBA | __________ | __________ |
| Review/purge old data | Annually | DBA + PIF Owner | __________ | __________ |

---

## Change Log

| Date | Changed By | Component | Change Description |
|------|------------|-----------|-------------------|
| __________ | __________ | __________ | _________________________ |
| __________ | __________ | __________ | _________________________ |
| __________ | __________ | __________ | _________________________ |
| __________ | __________ | __________ | _________________________ |

---

## Notes

Additional configuration notes, decisions, or context:

_______________________________________________________________________________
_______________________________________________________________________________
_______________________________________________________________________________
_______________________________________________________________________________
_______________________________________________________________________________
_______________________________________________________________________________
_______________________________________________________________________________
_______________________________________________________________________________

---

## Sign-Off

### Configuration Completed

**Configured By**: _______________________________  Date: ___________

**Verified By**: _______________________________  Date: ___________

**Approved By**: _______________________________  Date: ___________

### Ready for Production

**Production Deployment Approved**: ☐ Yes  ☐ No

**Production Go-Live Date**: ___________________

**Rollback Plan Documented**: ☐ Yes  ☐ No

---

**Document Version**: 1.0
**Last Updated**: ___________________
**Document Location**: ___________________
