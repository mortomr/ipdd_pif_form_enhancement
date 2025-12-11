Attribute VB_Name = "mod_ValidationIntegration"
' ============================================================================
' MODULE: mod_ValidationIntegration
' ============================================================================
' Purpose: Wire validation report into existing Submit flow
' Usage:
'   - Add [Validate Before Submit] button → calls ValidateBeforeSubmit()
'   - Update [Submit to Database] button → check validation first, then submit
'
' Workflow:
'   1. User clicks [Validate Before Submit]
'   2. Validation report populates on Validation_Report sheet
'   3. User reviews and fixes errors
'   4. User clicks [Submit to Database]
'   5. Quick validation check (fails if FAIL status exists)
'   6. Proceeds with normal submit process
' ============================================================================

Option Explicit

' ============================================================================
' PUBLIC SUBS - INTEGRATION POINTS
' ============================================================================

' ---------------------------------------------------------------------------
' Sub: PreSubmitValidationCheck
' Purpose: Quick validation check before actual submission
' Usage: Call this from your [Submit to Database] button before calling SaveSnapshot or similar
' Returns: True if validation passes (no FAIL results), False if validation found errors
' ---------------------------------------------------------------------------
Public Function PreSubmitValidationCheck() As Boolean
    On Error GoTo ErrHandler

    Dim wsValidation As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim failCount As Long

    ' Try to access validation report
    On Error Resume Next
    Set wsValidation = ThisWorkbook.Sheets("Validation_Report")
    On Error GoTo ErrHandler

    If wsValidation Is Nothing Then
        ' No validation report exists; user hasn't validated yet
        MsgBox "Please run [Validate Before Submit] first to check for errors.", _
               vbExclamation, "Validation Required"
        PreSubmitValidationCheck = False
        Exit Function
    End If

    ' Count FAIL results in the report
    lastRow = wsValidation.Cells(wsValidation.Rows.count, 6).End(xlUp).row
    failCount = 0

    For rowNum = 2 To lastRow ' Skip header row
        If wsValidation.Cells(rowNum, 6).value = "FAIL" Then
            failCount = failCount + 1
        End If
    Next rowNum

    If failCount > 0 Then
        MsgBox "Submission blocked: " & failCount & " validation error(s) found." & vbCrLf & vbCrLf & _
               "See Validation_Report worksheet for details. Fix errors and validate again.", _
               vbCritical, "Validation Errors"
        PreSubmitValidationCheck = False
    Else
        PreSubmitValidationCheck = True
    End If

    Exit Function

ErrHandler:
    MsgBox "Error checking validation: " & Err.Description, vbCritical
    PreSubmitValidationCheck = False
End Function

' ---------------------------------------------------------------------------
' Sub: RefreshValidationReport
' Purpose: Re-run validation without showing dialog (for periodic checks)
' Usage: Call periodically if user has made changes since initial validation
' ---------------------------------------------------------------------------
Public Sub RefreshValidationReport()
    On Error GoTo ErrHandler

    Dim selectedSite As String

    ' Get selected site
    On Error Resume Next
    selectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.value)
    On Error GoTo ErrHandler

    If selectedSite = "" Then
        MsgBox "Please select a site first.", vbExclamation
        Exit Sub
    End If

    ' Run validation quietly (no dialog)
    Dim validationResults As Variant
    validationResults = RunAllValidations(selectedSite)

    ' Update the report
    mod_ValidationReport.PopulateValidationReport validationResults

    Debug.Print "Validation report refreshed"
    Exit Sub

ErrHandler:
    MsgBox "Error refreshing validation: " & Err.Description, vbCritical
End Sub

' ============================================================================
' HELPER: WIRE INTO SUBMIT BUTTON
' ============================================================================
' Example code for your [Submit to Database] button:
'
' Public Sub SubmitToDatabase()
'     On Error GoTo ErrHandler
'
'     ' Step 1: Pre-submission validation check
'     If Not mod_ValidationIntegration.PreSubmitValidationCheck() Then
'         Exit Sub  ' Validation failed; stop here
'     End If
'
'     ' Step 2: Confirmation
'     Dim response As VbMsgBoxResult
'     response = MsgBox("Submit data to database?", vbYesNo + vbQuestion)
'     If response <> vbYes Then
'         Exit Sub
'     End If
'
'     ' Step 3: Proceed with normal submit process
'     ' (call your existing SaveSnapshot, ArchiveApprovedPIFs, etc.)
'     Call SaveSnapshot  ' or whatever your existing submit function is called
'
'     Exit Sub
' ErrHandler:
'     MsgBox "Submission error: " & Err.Description, vbCritical
' End Sub

'End Module
