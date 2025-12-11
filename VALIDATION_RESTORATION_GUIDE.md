# Validation Report Restoration - Implementation Guide

## Overview
This guide walks you through restoring the validation report functionality that was lost during the line_item migration.

**Three modules provided:**
1. `mod_ValidationReport.bas` — Core validation engine
2. `mod_ValidationConfig.bas` — Configurable rules (no code changes needed)
3. `mod_ValidationIntegration.bas` — Wire into existing submit workflow

---

## Step 1: Import Modules

1. Open your PIF workbook in VBA editor (Alt+F11)
2. Right-click on project tree → **Import File**
3. Import these three files in order:
   - `mod_ValidationReport.bas`
   - `mod_ValidationConfig.bas`
   - `mod_ValidationIntegration.bas`

---

## Step 2: Update Column Constants (CRITICAL)

Open `mod_ValidationReport.bas` and update the column position constants to match your actual worksheet structure:

```vba
Private Const COL_ARCHIVE As Long = 3
Private Const COL_INCLUDE As Long = 4
Private Const COL_ACCT_TREATMENT As Long = 5
Private Const COL_CHANGE_TYPE As Long = 6
Private Const COL_LINE_ITEM As Long = 7
Private Const COL_PIF_ID As Long = 8
Private Const COL_SEG As Long = 9
Private Const COL_OPCO As Long = 10
Private Const COL_SITE As Long = 11
Private Const COL_STRATEGIC_RANK As Long = 12
Private Const COL_FUNDING_PROJECT As Long = 13
Private Const COL_PROJECT_NAME As Long = 14
Private Const COL_ORIGINAL_ISD As Long = 15
Private Const COL_REVISED_ISD As Long = 16
Private Const COL_LCM_ISSUE As Long = 17
Private Const COL_STATUS As Long = 18
Private Const COL_CATEGORY As Long = 19
Private Const COL_JUSTIFICATION As Long = 20
```

**To find your column positions:**
- Open PIF sheet
- Look at the header row (row 3)
- Count columns: A=1, B=2, C=3, etc.
- Update the constants to match

---

## Step 3: Add [Validate Before Submit] Button

1. Go to **PIF** worksheet
2. Insert a button control (Insert → Shapes → Button)
3. Name it `[Validate Before Submit]`
4. Right-click → **Assign Macro**
5. Select `mod_ValidationReport.ValidateBeforeSubmit`

---

## Step 4: Update [Submit to Database] Button (Optional but Recommended)

Your existing [Submit to Database] button currently calls your submit function directly.
Add a validation check first:

**Old code:**
```vba
Public Sub SubmitToDatabase()
    On Error GoTo ErrHandler
    
    ' ... existing code ...
    Call SaveSnapshot
    
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

**New code:**
```vba
Public Sub SubmitToDatabase()
    On Error GoTo ErrHandler
    
    ' STEP 1: Check validation
    If Not mod_ValidationIntegration.PreSubmitValidationCheck() Then
        Exit Sub  ' Validation failed; block submission
    End If
    
    ' STEP 2: Confirmation
    Dim response As VbMsgBoxResult
    response = MsgBox("Submit data to database?", vbYesNo + vbQuestion, "Confirm Submit")
    If response <> vbYes Then
        Exit Sub
    End If
    
    ' STEP 3: Proceed with submit (your existing code)
    Call SaveSnapshot
    
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

---

## Step 5: Test the Validation

1. **Enter test data in PIF sheet:**
   - Row 4: PIF-001, Project-100, Change Type="Budget Adjustment", Original ISD="01/01/2025", Revised ISD=(leave blank)
   - Row 5: PIF-002, Project-200, Change Type="Funding Increase", Revised ISD=(leave blank), Site="ANO"

2. **Click [Validate Before Submit]**

3. **Expected results:**
   - A new `Validation_Report` worksheet appears
   - Row 4, Revised_ISD field: **PASS** (not required for Budget Adjustment)
   - Row 5, Revised_ISD field: **FAIL** (required for Funding Increase)
   - Row 5, Site field: **PASS** (matches selected site)
   - Click hyperlinks in report to jump back to PIF sheet

---

## Step 6: Customize Validation Rules (Optional)

If you need to adjust when fields are required, edit `mod_ValidationConfig.bas`:

### Example: Make Justification Required for ALL PIFs (not just archived)

**Current code in mod_ValidationConfig:**
```vba
Public Function IsJustificationRequired(archiveFlag As String) As Boolean
    IsJustificationRequired = (archiveFlag = "X" Or archiveFlag = "TRUE" Or archiveFlag = 1)
End Function
```

**Modified code:**
```vba
Public Function IsJustificationRequired(archiveFlag As String) As Boolean
    IsJustificationRequired = True  ' Always required
End Function
```

### Example: Add a new required field check

In `mod_ValidationConfig.bas`, add:
```vba
Public Function IsStrategicRankRequired(category As String) As Boolean
    ' Strategic Rank required for Strategic category only
    IsStrategicRankRequired = (UCase(category) = "STRATEGIC")
End Function
```

Then in `mod_ValidationReport.bas`, in the `RunAllValidations` function, add a call:
```vba
' Around line 120, after category validation:
ValidateStrategicRankRules rowNum, category, strategicRank, results
```

And add the new validation sub:
```vba
Private Sub ValidateStrategicRankRules(rowNum As Long, category As String, strategicRank As String, ByRef results As Collection)
    If mod_ValidationConfig.IsStrategicRankRequired(category) Then
        ValidateRequired rowNum, strategicRank, "Strategic_Rank", results
    End If
End Sub
```

---

## Validation Rules Currently Implemented

| Rule | Condition | Action |
|------|-----------|--------|
| PIF_ID Required | Always | Fails if blank |
| Project_ID Required | Always | Fails if blank |
| Site Match | Always | Fails if doesn't match selected site |
| Duplicate Detection | Always | Fails if PIF+Project+LineItem seen twice |
| Revised ISD Required | Change Type = "Funding/Scope/Schedule Change" | Fails if blank |
| LCM Issue Required | Category = "Compliance" | Fails if blank |
| Justification Required | Archive_Flag = "X" (approved submission) | Fails if blank |
| Date Format Valid | Any date field | Fails if not MM/DD/YYYY format |
| Line Item Default | If blank | Defaults to 1 |

---

## Adding New Validation Rules

**To add a new rule:**

1. **Define the condition in `mod_ValidationConfig.bas`:**
   ```vba
   Public Function IsMyFieldRequired(someParameter As String) As Boolean
       IsMyFieldRequired = (someCondition)
   End Function
   ```

2. **Add validation check in `mod_ValidationReport.RunAllValidations()`:**
   ```vba
   ' Extract field value
   Dim myField As String
   myField = Trim(wsPIF.Cells(rowNum, COL_MY_FIELD).value & "")
   
   ' Add validation call
   ValidateMyFieldRule rowNum, myField, someParameter, results
   ```

3. **Add the validation sub:**
   ```vba
   Private Sub ValidateMyFieldRule(rowNum As Long, fieldValue As String, param As String, ByRef results As Collection)
       If mod_ValidationConfig.IsMyFieldRequired(param) Then
           ValidateRequired rowNum, fieldValue, "My_Field", results
       End If
   End Sub
   ```

---

## Troubleshooting

### Validation Report not appearing
- Make sure columns constants are correct (Step 2)
- Check debug output: press Ctrl+G and look for error messages
- Verify PIF sheet name is exactly "PIF"

### Hyperlinks in report not working
- Check that PIF sheet name matches the constant `WS_PIF`
- Make sure first data row is 4 (adjust `FIRST_DATA_ROW` if different)

### False positive validation failures
- Review the rule conditions in `mod_ValidationConfig.bas`
- Check cell values are exactly what you expect (trim spaces)
- Add debug prints to understand cell contents

### Performance is slow
- Reduce number of rows being validated
- Check for infinite loops in validation logic
- Consider moving cost data validation to a separate module

---

## Rollback

If you need to remove validation:
1. Delete the three imported modules from VBA editor
2. Delete the [Validate Before Submit] button
3. Revert [Submit to Database] button to original code
4. Delete the Validation_Report worksheet (it will auto-recreate on next validation run if modules still exist)

---

## Next Steps

Once validation report is restored:
1. Have users run validation on test data
2. Collect feedback on rule accuracy
3. Adjust conditional logic in `mod_ValidationConfig.bas` as needed
4. Document any custom rules for your team
5. Add error suppression to allow WARNs (non-blocking errors) if desired

---

**Questions? Debug by adding `Debug.Print` statements to understand data flow.**

Example:
```vba
Debug.Print "Row " & rowNum & ": PIF=" & pifID & ", ChangeType=" & changeType
Debug.Print "  Requires Revised ISD? " & mod_ValidationConfig.IsRevisedISDRequired(changeType)
```

Check output in Immediate Window (Ctrl+G).
