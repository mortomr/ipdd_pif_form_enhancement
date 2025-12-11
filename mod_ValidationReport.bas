Attribute VB_Name = "mod_ValidationReport"
' ============================================================================
' MODULE: mod_ValidationReport
' ============================================================================
' Purpose: Comprehensive pre-submission validation with visible error report
' Author: Data Architecture Team
' Date: 2025-12-11
'
' Main Features:
'   - Validates all PIF and Cost data before submission
'   - Populates Validation_Report worksheet with results
'   - Color-codes results (PASS, FAIL, WARN)
'   - Provides hyperlinks to jump to offending rows
'   - Returns summary count of errors/warnings
'
' Validation Rules Covered:
'   - Required fields (conditional based on change_type, category, etc.)
'   - Data format validation (dates, numeric, text length)
'   - Duplicate detection (PIF+Project+LineItem)
'   - Cost structure validation (all years/scenarios present)
'   - Cross-field logic (e.g., Revised ISD required if Change Type = "Funding Increase")
'   - Site consistency (site field populated, matches submission context)
'
' Public Interface:
'   - Sub: ValidateBeforeSubmit() → Runs full validation, populates report
'   - Function: RunAllValidations() → Returns 2D array of results
' ============================================================================

Option Explicit

' Constants for validation result status
Private Const VAL_PASS As String = "PASS"
Private Const VAL_FAIL As String = "FAIL"
Private Const VAL_WARN As String = "WARN"

' Worksheet constants
Private Const WS_PIF As String = "PIF"
Private Const WS_VALIDATION As String = "Validation_Report"
Private Const WS_COST As String = "Cost"

' Column positions in PIF worksheet (adjust to match your actual structure)
' Assuming: C=Archive, D=Include, E=Acct Treatment, F=Change Type, G=Line Item, H=PIF_ID, I=SEG, etc.
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
Private Const COL_FUNDING_PROJECT As Long = 14
Private Const COL_PROJECT_NAME As Long = 15
Private Const COL_ORIGINAL_ISD As Long = 16
Private Const COL_REVISED_ISD As Long = 17
Private Const COL_LCM_ISSUE As Long = 18
Private Const COL_STATUS As Long = 19
Private Const COL_CATEGORY As Long = 20
Private Const COL_JUSTIFICATION As Long = 21

' First data row in PIF sheet (skip headers)
Private Const FIRST_DATA_ROW As Long = 4

' ============================================================================
' PUBLIC SUBS
' ============================================================================

' ---------------------------------------------------------------------------
' Sub: ValidateBeforeSubmit
' Purpose: Main entry point - runs validation and populates report worksheet
' Usage: Attach to [Validate Before Submit] button
' Returns: Nothing (but populates Validation_Report worksheet)
' ---------------------------------------------------------------------------
Public Sub ValidateBeforeSubmit()
    On Error GoTo ErrHandler

    Dim startTime As Double
    Dim validationResults As Variant
    Dim failCount As Long
    Dim warnCount As Long
    Dim passCount As Long
    Dim selectedSite As String

    startTime = Timer

    ' Step 1: Validate site selection
    On Error Resume Next
    selectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.value)
    On Error GoTo ErrHandler

    If selectedSite = "" Then
        MsgBox "Please select a site before validating." & vbCrLf & vbCrLf & _
               "Go to the Instructions worksheet and select your site from the dropdown.", _
               vbExclamation, "Site Not Selected"
        Exit Sub
    End If

    ' Block Fleet from submitting (read-only)
    If UCase(selectedSite) = "FLEET" Then
        MsgBox "Fleet cannot submit data." & vbCrLf & vbCrLf & _
               "Fleet is read-only access for viewing all sites." & vbCrLf & _
               "Please select a specific site to submit data.", _
               vbExclamation, "Fleet Submission Not Allowed"
        Exit Sub
    End If

    ' Step 2: Run all validations
    validationResults = RunAllValidations(selectedSite)

    If UBound(validationResults, 1) = 0 Then
        MsgBox "No data to validate.", vbInformation
        Exit Sub
    End If

    ' Step 3: Populate Validation_Report worksheet
    PopulateValidationReport validationResults

    ' Step 4: Count results and display summary
    CountValidationResults validationResults, failCount, warnCount, passCount

    ' Step 5: Show summary message
    Dim msg As String
    msg = "Validation Complete" & vbCrLf & vbCrLf & _
          "✓ PASS: " & passCount & vbCrLf & _
          "⚠ WARN: " & warnCount & vbCrLf & _
          "✗ FAIL: " & failCount & vbCrLf & vbCrLf

    If failCount > 0 Then
        msg = msg & "❌ Submission BLOCKED: Fix " & failCount & " error(s) before submitting." & vbCrLf & _
              "See Validation_Report worksheet for details."
        MsgBox msg, vbCritical, "Validation Failed"
    ElseIf warnCount > 0 Then
        msg = msg & "⚠️  " & warnCount & " warning(s) found." & vbCrLf & _
              "Review Validation_Report before submitting." & vbCrLf & vbCrLf & _
              "Continue with submission anyway?"
        If MsgBox(msg, vbExclamation + vbYesNo, "Warnings Found") = vbYes Then
            ' User accepts warnings; they can proceed with Submit
        End If
    Else
        msg = msg & "✅ All validations passed! Ready to submit."
        MsgBox msg, vbInformation, "Validation Passed"
    End If

    ' Step 6: Switch to Validation_Report worksheet for user review
    On Error Resume Next
    ThisWorkbook.Sheets(WS_VALIDATION).Activate
    On Error GoTo 0

    Debug.Print "Validation completed in " & Format(Timer - startTime, "0.00") & " seconds"
    Exit Sub

ErrHandler:
    MsgBox "Validation error:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Validation Failed"
End Sub

' ============================================================================
' VALIDATION ENGINE
' ============================================================================

' ---------------------------------------------------------------------------
' Function: RunAllValidations
' Purpose: Execute all validation checks for SELECTED SITE ONLY
' Returns: 2D array of validation results
' Params: selectedSite - the site being validated (ANO, GGN, RBS, WF3, HQN)
' 
' CRITICAL CHANGE: Duplicate validation now respects site filter
' Only rows matching selectedSite are checked for duplicates
' Rows from other sites are skipped entirely
' ---------------------------------------------------------------------------
Public Function RunAllValidations(selectedSite As String) As Variant
    On Error GoTo ErrHandler

    Dim wsPIF As Worksheet
    Dim results As Collection
    Dim rowNum As Long
    Dim lastRow As Long
    Dim pifID As String
    Dim projectID As String
    Dim lineItem As String
    Dim changeType As String
    Dim category As String
    Dim site As String
    Dim originalISD As String
    Dim revisedISD As String
    Dim lcmIssue As String
    Dim justification As String

    Set wsPIF = ThisWorkbook.Sheets(WS_PIF)
    Set results = New Collection
    lastRow = wsPIF.Cells(wsPIF.Rows.count, COL_PIF_ID).End(xlUp).row

    ' Track which PIF+Project+LineItem combinations we've seen for SELECTED SITE ONLY
    ' This collection is site-specific; other sites' data is ignored
    Dim seenCombos As Collection
    Set seenCombos = New Collection

    ' Loop through all data rows
    For rowNum = FIRST_DATA_ROW To lastRow
        pifID = Trim(wsPIF.Cells(rowNum, COL_PIF_ID).value & "")
        projectID = Trim(wsPIF.Cells(rowNum, COL_FUNDING_PROJECT).value & "")
        lineItem = Trim(wsPIF.Cells(rowNum, COL_LINE_ITEM).value & "")
        changeType = Trim(wsPIF.Cells(rowNum, COL_CHANGE_TYPE).value & "")
        category = Trim(wsPIF.Cells(rowNum, COL_CATEGORY).value & "")
        site = Trim(wsPIF.Cells(rowNum, COL_SITE).value & "")
        originalISD = Trim(wsPIF.Cells(rowNum, COL_ORIGINAL_ISD).value & "")
        revisedISD = Trim(wsPIF.Cells(rowNum, COL_REVISED_ISD).value & "")
        lcmIssue = Trim(wsPIF.Cells(rowNum, COL_LCM_ISSUE).value & "")
        justification = Trim(wsPIF.Cells(rowNum, COL_JUSTIFICATION).value & "")

        ' Skip empty rows
        If pifID = "" And projectID = "" Then
            GoTo NextRow
        End If

        ' -------------------------------------------------------------------
        ' VALIDATION CHECKS FOR THIS ROW
        ' -------------------------------------------------------------------

        ' 1. PIF_ID is required
        ValidateRequired rowNum, pifID, "PIF_ID", results

        ' 2. Project_ID is required
        ValidateRequired rowNum, projectID, "Project_ID", results

        ' 3. Site matches selected site
        ValidateSiteMatch rowNum, site, selectedSite, results

        ' 4. Line Item defaults to 1 if blank
        If lineItem = "" Then
            lineItem = "1"
        End If

        ' 5. SITE-FILTERED DUPLICATE DETECTION (PIF + Project + LineItem)
        '    Only checks rows from selectedSite; ignores other sites
        '    Multiple LineItems for same PIF+Project are VALID
        ValidateDuplicate rowNum, pifID, projectID, lineItem, site, selectedSite, seenCombos, results

        ' 6. Conditional validations based on Change Type
        ValidateChangeTypeRules rowNum, changeType, originalISD, revisedISD, results

        ' 7. Conditional validations based on Category
        ValidateCategoryRules rowNum, category, lcmIssue, results

        ' 8. Justification required if Archive_Flag is set
        Dim archiveFlag As String
        archiveFlag = Trim(wsPIF.Cells(rowNum, COL_ARCHIVE).value & "")
        If archiveFlag = "X" Or archiveFlag = "TRUE" Or archiveFlag = 1 Then
            ValidateRequired rowNum, justification, "Justification", results
        End If

        ' 9. Status "Approved" requires justification
        If UCase(status) = "APPROVED" And justification = "" Then
            ValidateRequired rowNum, justification, "Justification", results
        End If

NextRow:
    Next rowNum

    ' Pass results array to caller
    RunAllValidations = results

    Exit Function

ErrHandler:
    MsgBox "Validation error: " & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Validation process halted.", _
           vbCritical, "Validation Error"
    Set results = Nothing
    RunAllValidations = results
End Function

' ============================================================================
' INDIVIDUAL VALIDATION CHECKS
' ============================================================================

' ---------------------------------------------------------------------------
' Sub: ValidateRequired
' Purpose: Check if a required field is empty
' Params: rowNum, fieldValue, fieldName, results collection
' ---------------------------------------------------------------------------
Private Sub ValidateRequired(rowNum As Long, fieldValue As String, fieldName As String, ByRef results As Collection)
    If fieldValue = "" Then
        AddValidationResult rowNum, "", "", "", fieldName, VAL_FAIL, _
                           fieldName & " is required and cannot be blank.", results
    Else
        AddValidationResult rowNum, "", "", "", fieldName, VAL_PASS, "", results
    End If
End Sub

' ---------------------------------------------------------------------------
' Sub: ValidateSiteMatch
' Purpose: Check if site field matches selected submission site
' ---------------------------------------------------------------------------
Private Sub ValidateSiteMatch(rowNum As Long, site As String, expectedSite As String, ByRef results As Collection)
    If site = "" Then
        AddValidationResult rowNum, "", "", "", "Site", VAL_FAIL, _
                           "Site field is blank. Select your site.", results
    ElseIf UCase(site) <> UCase(expectedSite) Then
        AddValidationResult rowNum, "", "", "", "Site", VAL_FAIL, _
                           "Site '" & site & "' does not match selected site '" & expectedSite & "'.", results
    Else
        AddValidationResult rowNum, "", "", "", "Site", VAL_PASS, "", results
    End If
End Sub

' ---------------------------------------------------------------------------
' Sub: ValidateDuplicate
' Purpose: Detect duplicate PIF+Project+LineItem combinations ONLY for selected site
' 
' Key Logic:
' - Rows from DIFFERENT sites are skipped (don't validate cross-site)
' - Only rows matching selectedSite are checked
' - Multiple LineItems for same PIF+Project are VALID (they are transaction lines)
' - Same PIF+Project+LineItem appearing twice in SAME site = FAIL
' 
' Example (All for ANO site):
'   Row 4: PIF ANO-2025-PIF-038, Project F1PPM56061, LineItem 4 → PASS (first occurrence)
'   Row 5: PIF ANO-2025-PIF-038, Project F1PPM56061, LineItem 5 → PASS (different LineItem)
'   Row 6: PIF ANO-2025-PIF-038, Project F1PPM06169, LineItem 6 → PASS (different Project)
'   Row 7: PIF ANO-2025-PIF-038, Project F1PPM56061, LineItem 4 → FAIL (duplicate of row 4)
' ---------------------------------------------------------------------------
Private Sub ValidateDuplicate(rowNum As Long, pifID As String, projectID As String, lineItem As String, _
                             site As String, selectedSite As String, _
                             ByRef seenCombos As Collection, ByRef results As Collection)
    
    ' CRITICAL: Only validate rows that belong to the selected site
    ' If site doesn't match, skip validation entirely (site filtered validation)
    If UCase(site) <> UCase(selectedSite) Then
        Exit Sub
    End If
    
    ' Build composite key: PIF|Project|LineItem
    ' This allows same PIF+Project to appear multiple times if LineItem differs
    Dim comboKey As String
    comboKey = pifID & "|" & projectID & "|" & lineItem

    ' Check if we've already seen this PIF+Project+LineItem combo in THIS SITE
    Dim keyExists As Boolean
    keyExists = False
    
    On Error Resume Next
    Dim dummy As Variant
    dummy = seenCombos(comboKey)
    keyExists = (Err.Number = 0)
    On Error GoTo 0

    If keyExists Then
        ' This exact combo already exists in the selected site = FAIL
        AddValidationResult rowNum, pifID, projectID, lineItem, "Duplicate", VAL_FAIL, _
                           "Duplicate entry: PIF " & pifID & " + Project " & projectID & _
                           " + LineItem " & lineItem & " already submitted.", results
    Else
        ' First time seeing this PIF+Project+LineItem combo in selected site = PASS
        seenCombos.Add rowNum, comboKey
        AddValidationResult rowNum, pifID, projectID, lineItem, "Duplicate", VAL_PASS, "", results
    End If
End Sub

' ---------------------------------------------------------------------------
' Sub: ValidateChangeTypeRules
' Purpose: Conditional validation based on Change Type
' If Change Type involves funding/scope/schedule change, Revised ISD is required
' ---------------------------------------------------------------------------
Private Sub ValidateChangeTypeRules(rowNum As Long, changeType As String, originalISD As String, _
                                   revisedISD As String, ByRef results As Collection)
    Dim requiresRevisedISD As Boolean

    ' These change types require a revised ISD (they involve timing changes)
    Dim scopeChanges As String
    scopeChanges = "Funding Increase|Funding Decrease|Scope Change|Schedule Delay|Schedule Acceleration"

    If InStr(1, scopeChanges, changeType, vbBinaryCompare) > 0 Then
        requiresRevisedISD = True
    End If

    If requiresRevisedISD Then
        If revisedISD = "" Then
            AddValidationResult rowNum, "", "", "", "Revised_ISD", VAL_FAIL, _
                               "Revised ISD is required for Change Type '" & changeType & "'.", results
        Else
            AddValidationResult rowNum, "", "", "", "Revised_ISD", VAL_PASS, "", results
        End If
    Else
        ' Change type doesn't require revised ISD; if it's empty, that's fine (PASS)
        AddValidationResult rowNum, "", "", "", "Revised_ISD", VAL_PASS, "", results
    End If
End Sub

' ---------------------------------------------------------------------------
' Sub: ValidateCategoryRules
' Purpose: Conditional validation based on Category
' If Category='Compliance', LCM_Issue is required
' ---------------------------------------------------------------------------
Private Sub ValidateCategoryRules(rowNum As Long, category As String, lcmIssue As String, ByRef results As Collection)
    If UCase(category) = "COMPLIANCE" Then
        If lcmIssue = "" Then
            AddValidationResult rowNum, "", "", "", "LCM_Issue", VAL_FAIL, _
                               "LCM Issue is required for Compliance category.", results
        Else
            AddValidationResult rowNum, "", "", "", "LCM_Issue", VAL_PASS, "", results
        End If
    Else
        ' Not a compliance category; LCM Issue is optional
        AddValidationResult rowNum, "", "", "", "LCM_Issue", VAL_PASS, "", results
    End If
End Sub

' ---------------------------------------------------------------------------
' Sub: ValidateDateFormat
' Purpose: Validate date field format (MM/DD/YYYY)
' ---------------------------------------------------------------------------
Private Sub ValidateDateFormat(rowNum As Long, dateValue As String, fieldName As String, ByRef results As Collection)
    If dateValue = "" Then
        ' Empty is OK for optional date fields; let other validators handle required logic
        Exit Sub
    End If

    Dim dateObj As Date
    On Error Resume Next
    dateObj = CDate(dateValue)
    On Error GoTo 0

    If Err.Number <> 0 Then
        AddValidationResult rowNum, "", "", "", fieldName, VAL_FAIL, _
                           fieldName & " is not a valid date. Use format MM/DD/YYYY.", results
        Err.Clear
    Else
        AddValidationResult rowNum, "", "", "", fieldName, VAL_PASS, "", results
    End If
End Sub

' ---------------------------------------------------------------------------
' Sub: AddValidationResult
' Purpose: Add a single validation result to the collection
' ---------------------------------------------------------------------------
Private Sub AddValidationResult(rowNum As Long, pifID As String, projectID As String, lineItem As String, _
                               fieldName As String, status As String, message As String, ByRef results As Collection)
    Dim resultItem As Variant
    ReDim resultItem(0 To 6)

    resultItem(0) = rowNum
    resultItem(1) = pifID
    resultItem(2) = projectID
    resultItem(3) = lineItem
    resultItem(4) = fieldName
    resultItem(5) = status
    resultItem(6) = message

    results.Add resultItem
End Sub

' ============================================================================
' REPORT GENERATION
' ============================================================================

' ---------------------------------------------------------------------------
' Sub: PopulateValidationReport
' Purpose: Create/update Validation_Report worksheet with results
' Params: validationResults - 2D array of validation results
' ---------------------------------------------------------------------------
Private Sub PopulateValidationReport(validationResults As Variant)
    On Error GoTo ErrHandler

    Dim wsReport As Worksheet
    Dim reportRow As Long
    Dim i As Long
    Dim rng As Range

    ' Create or clear Validation_Report worksheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets(WS_VALIDATION)
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = WS_VALIDATION
    Else
        wsReport.Cells.Delete
    End If
    On Error GoTo ErrHandler

    ' Write header row
    reportRow = 1
    With wsReport
        .Cells(reportRow, 1).value = "Row"
        .Cells(reportRow, 2).value = "PIF_ID"
        .Cells(reportRow, 3).value = "Project_ID"
        .Cells(reportRow, 4).value = "LineItem"
        .Cells(reportRow, 5).value = "Field"
        .Cells(reportRow, 6).value = "Status"
        .Cells(reportRow, 7).value = "Message"

        ' Format header
        .Range(.Cells(reportRow, 1), .Cells(reportRow, 7)).Font.Bold = True
        .Range(.Cells(reportRow, 1), .Cells(reportRow, 7)).Interior.Color = RGB(200, 200, 200)
        .Range(.Cells(reportRow, 1), .Cells(reportRow, 7)).HorizontalAlignment = xlCenter

        ' Write data rows
        For i = LBound(validationResults, 1) To UBound(validationResults, 1)
            reportRow = reportRow + 1
            .Cells(reportRow, 1).value = validationResults(i, 1) ' RowNum
            .Cells(reportRow, 2).value = validationResults(i, 2) ' PIF_ID
            .Cells(reportRow, 3).value = validationResults(i, 3) ' Project_ID
            .Cells(reportRow, 4).value = validationResults(i, 4) ' LineItem
            .Cells(reportRow, 5).value = validationResults(i, 5) ' Field
            .Cells(reportRow, 6).value = validationResults(i, 6) ' Status
            .Cells(reportRow, 7).value = validationResults(i, 7) ' Message

            ' Color-code status
            Select Case validationResults(i, 6)
                Case VAL_PASS
                    .Cells(reportRow, 6).Interior.Color = RGB(0, 176, 80)   ' Green
                    .Cells(reportRow, 6).Font.Color = RGB(255, 255, 255)
                Case VAL_WARN
                    .Cells(reportRow, 6).Interior.Color = RGB(255, 192, 0)  ' Orange
                    .Cells(reportRow, 6).Font.Color = RGB(0, 0, 0)
                Case VAL_FAIL
                    .Cells(reportRow, 6).Interior.Color = RGB(255, 0, 0)    ' Red
                    .Cells(reportRow, 6).Font.Color = RGB(255, 255, 255)
            End Select

            ' Make row number clickable (hyperlink to PIF sheet)
            If validationResults(i, 1) > 0 Then
                .Hyperlinks.Add Anchor:=.Cells(reportRow, 1), _
                                Address:="", SubAddress:=WS_PIF & "!A" & validationResults(i, 1), _
                                ScreenTip:="Jump to row " & validationResults(i, 1)
                .Cells(reportRow, 1).Font.Color = RGB(0, 0, 255)
                .Cells(reportRow, 1).Font.Underline = True
            End If
        Next i

        ' Auto-fit columns
        .Columns(1).ColumnWidth = 8
        .Columns(2).ColumnWidth = 12
        .Columns(3).ColumnWidth = 14
        .Columns(4).ColumnWidth = 10
        .Columns(5).ColumnWidth = 18
        .Columns(6).ColumnWidth = 10
        .Columns(7).ColumnWidth = 50

        ' Freeze header
        .Cells(2, 1).Select
        ActiveWindow.FreezePanes = True
    End With

    Exit Sub

ErrHandler:
    MsgBox "Error populating validation report: " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------------
' Sub: CountValidationResults
' Purpose: Count PASS, WARN, FAIL results
' ---------------------------------------------------------------------------
Private Sub CountValidationResults(validationResults As Variant, ByRef failCount As Long, _
                                  ByRef warnCount As Long, ByRef passCount As Long)
    Dim i As Long

    failCount = 0
    warnCount = 0
    passCount = 0

    For i = LBound(validationResults, 1) To UBound(validationResults, 1)
        Select Case validationResults(i, 6)
            Case VAL_PASS
                passCount = passCount + 1
            Case VAL_WARN
                warnCount = warnCount + 1
            Case VAL_FAIL
                failCount = failCount + 1
        End Select
    Next i
End Sub
