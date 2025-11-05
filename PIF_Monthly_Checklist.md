# PIF Monthly Submission Checklist

## Quick Reference Guide

Use this checklist for each monthly PIF submission cycle.

---

## WEEK 1-3: DATA COLLECTION PHASE

### Preparation
- [ ] Open master workbook: `PIF_YYYY-MM_MASTER.xlsm`
- [ ] Verify workbook is in shared location (not local copy)
- [ ] Check that macros are enabled

### Data Entry
- [ ] All sites enter their PIF data
- [ ] Use data validation dropdowns for Status, Category, etc.
- [ ] Fill required fields: PIF ID, Project ID, Justification (for approved)
- [ ] Save frequently to avoid data loss

### Coordination
- [ ] Communicate updates via Teams/Email
- [ ] Flag incomplete PIFs for follow-up
- [ ] Resolve duplicate PIF IDs immediately

---

## WEEK 4 (DAY 20-25): VALIDATION PHASE

### Freeze Data Entry
- [ ] Send freeze announcement email
- [ ] Save final copy as: `PIF_YYYY-MM_FINAL.xlsm`
- [ ] Make workbook read-only (or move to different folder)

### Run Validation
1. **Open final workbook**
2. **Click "Validate Data" button** (or run macro: `RunValidationOnly`)
3. **Review Validation_Report sheet**
4. **Fix any errors found**
5. **Re-run validation until clean**

### Common Validation Errors
- Missing PIF ID or Project ID → Fill in required fields
- Duplicate PIF+Project → Remove duplicate or correct data
- Approved PIF without justification → Add justification text
- Invalid data types → Check that numbers are formatted as numbers

### Get Approval
- [ ] Share validation results with team lead
- [ ] Get email approval to submit
- [ ] Document approval in submission log or email

---

## MONTH-END (DAY 26-30): SUBMISSION PHASE

### Pre-Submission Checklist
- [ ] All validation errors resolved
- [ ] Final approval received
- [ ] Backup copy saved to safe location
- [ ] No other users have workbook open

### Submit to Database
1. **Open `PIF_YYYY-MM_FINAL.xlsm`**
2. **Click "Submit to Database" button**
   - Or run macro: `SubmitToDatabase`
3. **Read confirmation prompt carefully**
4. **Click "Yes" to proceed**
5. **Wait for completion message** (30-90 seconds)
6. **Save confirmation screenshot**

### Verify Submission
Run these SQL queries (or ask DBA to run):

```sql
-- 1. Check submission was logged
SELECT TOP 1 * 
FROM dbo.tbl_submission_log 
ORDER BY submission_date DESC;

-- 2. Check inflight record count
SELECT COUNT(*) AS Inflight_Records
FROM dbo.tbl_pif_projects_inflight;

-- 3. Check approved record count (should increase if any approved)
SELECT COUNT(*) AS Approved_Records
FROM dbo.tbl_pif_projects_approved;
```

- [ ] Submission appears in log
- [ ] Record count matches expectation
- [ ] No error messages in SQL

---

## NEW MONTH (DAY 1-5): DISTRIBUTION PHASE

### Generate Next Month's Workbook

Option A: Manual Export from SQL
```sql
SELECT * 
FROM dbo.vw_pif_current_working
WHERE status NOT IN ('Approved', 'Dispositioned')
ORDER BY pif_id;
```
- Export to Excel
- Paste into template workbook
- Save as: `PIF_YYYY-MM_MASTER.xlsm`

Option B: Use Power Query (if set up)
- Open template workbook
- Data > Refresh All
- Save as new month's file

### Distribution
- [ ] Copy new workbook to shared location
- [ ] Send email with link to all sites
- [ ] Include any notes/reminders for the month
- [ ] Archive previous month's final file

---

## TROUBLESHOOTING QUICK FIXES

### "Connection Failed" Error
1. Test connection in SQL Server Management Studio first
2. Check VPN is connected (if working remotely)
3. Verify server name in VBA code: `mod_Database` → `SQL_SERVER` constant
4. Contact DBA if still failing

### "Validation Failed" But No Errors Shown
1. Check Validation_Report sheet (may be hidden)
2. Unhide sheet: Right-click sheet tab → Unhide
3. Look for SQL validation errors at bottom of report

### Submission Hangs/Takes Forever
1. Check network connectivity
2. Close other applications using network
3. Wait up to 5 minutes before cancelling
4. If cancelled, check SQL to see if partial data was loaded
5. Contact DBA to rollback if needed

### Data Looks Wrong After Submission
1. Don't panic - backups exist
2. Contact DBA immediately
3. Provide backup table name: `tbl_pif_projects_inflight_backup_YYYYMMDD`
4. DBA can restore from backup

---

## EMERGENCY CONTACTS

**Database Issues:**
- DBA Team: [email/phone]
- After hours: [on-call number]

**Excel/VBA Issues:**
- Development Team: [email/phone]
- Alt: [backup contact]

**Process Questions:**
- PIF Owner: [name/email]
- Finance Lead: [name/email]

---

## MONTHLY METRICS TO TRACK

Track these for continuous improvement:

- [ ] Number of PIFs submitted
- [ ] Number of validation errors (first run)
- [ ] Time to complete validation
- [ ] Time to complete submission
- [ ] Issues encountered (document for next month)

---

## TIPS FOR SUCCESS

### Before You Start
✓ Save a backup copy before making major changes
✓ Don't work from a local copy - use the shared master file
✓ Communicate with other sites about timing

### During Data Entry
✓ Use dropdowns instead of typing status values
✓ Copy/paste carefully to avoid duplicating PIF IDs
✓ Review your own data before the freeze

### During Validation
✓ Run validation multiple times as you fix errors
✓ Fix errors in the order they appear
✓ Don't skip errors - they'll prevent submission

### During Submission
✓ Close all other programs to maximize performance
✓ Don't touch the keyboard while submission is running
✓ Wait for the success message before closing Excel

### After Submission
✓ Keep a copy of the submission confirmation
✓ Verify data in SQL before closing for the month
✓ Document any issues for the next cycle

---

## MONTHLY WORKFLOW AT A GLANCE

```
Week 1-3
└─ All sites enter data → Save regularly

Week 4
├─ Freeze data entry (Day 20)
├─ Run validation
├─ Fix errors
├─ Re-validate until clean
└─ Get approval (Day 25)

Month-End
├─ Submit to database (Day 26-30)
├─ Verify submission
└─ Archive final workbook

New Month
├─ Generate new workbook (Day 1-5)
├─ Distribute to sites
└─ Begin next cycle
```

---

## NEED HELP?

If this checklist doesn't cover your issue:
1. Check the full README.md for detailed troubleshooting
2. Contact your DBA or development team
3. Document the issue for future reference

---

**Remember**: The system creates automatic backups before every submission. If something goes wrong, we can recover!

Last Updated: November 5, 2025
