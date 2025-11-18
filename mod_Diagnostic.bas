<<<<<<< HEAD
Attribute VB_Name = "mod_Diagnostic"
' ============================================================================
' MODULE: mod_Diagnostic
' ============================================================================
' Purpose: Diagnose type mismatch errors in parameter passing
' Usage: Run TestSingleRowInsert() to test first data row
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Sub: TestSingleRowInsert
' Purpose: Test inserting a single row with detailed diagnostics
' Usage: Run this from VBA editor to see exactly where the error occurs
' ----------------------------------------------------------------------------
Public Sub Diag_TestInsert()
    On Error GoTo ErrHandler

    Dim wsData As Worksheet
    Dim conn As ADODB.Connection
    Dim testRow As Long
    Dim params(0 To 19) As Variant
    Dim i As Integer
    Dim msg As String

    ' Get the PIF sheet
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    testRow = 4  ' First data row

    ' Show what we're reading
    msg = "DIAGNOSTIC: Reading data from row " & testRow & vbCrLf & vbCrLf

    ' Read and display each parameter
    msg = msg & "Column G (pif_id): " & wsData.Cells(testRow, 7).Value & vbCrLf
    msg = msg & "Column M (project_id): " & wsData.Cells(testRow, 13).Value & vbCrLf
    msg = msg & "Column R (status): " & wsData.Cells(testRow, 18).Value & vbCrLf
    msg = msg & "Column F (change_type): " & wsData.Cells(testRow, 6).Value & vbCrLf
    msg = msg & "Column E (accounting_treatment): " & wsData.Cells(testRow, 5).Value & vbCrLf
    msg = msg & "Column S (category): " & wsData.Cells(testRow, 19).Value & vbCrLf
    msg = msg & "Column H (seg): [" & wsData.Cells(testRow, 8).Value & "] IsNumeric=" & IsNumeric(wsData.Cells(testRow, 8).Value) & vbCrLf
    msg = msg & "Column I (opco): " & wsData.Cells(testRow, 9).Value & vbCrLf
    msg = msg & "Column J (site): " & wsData.Cells(testRow, 10).Value & vbCrLf
    msg = msg & "Column K (strategic_rank): " & wsData.Cells(testRow, 11).Value & vbCrLf
    msg = msg & "Column M (funding_project): " & wsData.Cells(testRow, 13).Value & vbCrLf
    msg = msg & "Column N (project_name): " & wsData.Cells(testRow, 14).Value & vbCrLf
    msg = msg & "Column O (original_fp_isd): " & wsData.Cells(testRow, 15).Value & vbCrLf
    msg = msg & "Column P (revised_fp_isd): " & wsData.Cells(testRow, 16).Value & vbCrLf
    msg = msg & "Column AM (moving_isd_year): " & wsData.Cells(testRow, 39).Value & vbCrLf
    msg = msg & "Column Q (lcm_issue): " & wsData.Cells(testRow, 17).Value & vbCrLf
    msg = msg & "Column T (justification): " & wsData.Cells(testRow, 20).Value & vbCrLf
    msg = msg & "Column AN (prior_year_spend): [" & wsData.Cells(testRow, 40).Value & "] IsNumeric=" & IsNumeric(wsData.Cells(testRow, 40).Value) & vbCrLf
    msg = msg & "Column C (archive_flag): " & wsData.Cells(testRow, 3).Value & vbCrLf
    msg = msg & "Column D (include_flag): " & wsData.Cells(testRow, 4).Value & vbCrLf

    MsgBox msg, vbInformation, "Raw Excel Values"

    ' Now convert using Safe functions and show results
    params(0) = SafeString(wsData.Cells(testRow, 7).Value)
    params(1) = SafeString(wsData.Cells(testRow, 13).Value)
    params(2) = SafeString(wsData.Cells(testRow, 18).Value)
    params(3) = SafeString(wsData.Cells(testRow, 6).Value)
    params(4) = SafeString(wsData.Cells(testRow, 5).Value)
    params(5) = SafeString(wsData.Cells(testRow, 19).Value)
    params(6) = SafeInteger(wsData.Cells(testRow, 8).Value)
    params(7) = SafeString(wsData.Cells(testRow, 9).Value)
    params(8) = SafeString(wsData.Cells(testRow, 10).Value)
    params(9) = SafeString(wsData.Cells(testRow, 11).Value)
    params(10) = SafeString(wsData.Cells(testRow, 13).Value)
    params(11) = SafeString(wsData.Cells(testRow, 14).Value)
    params(12) = SafeString(wsData.Cells(testRow, 15).Value)
    params(13) = SafeString(wsData.Cells(testRow, 16).Value)
    params(14) = SafeString(wsData.Cells(testRow, 39).Value)
    params(15) = SafeString(wsData.Cells(testRow, 17).Value)
    params(16) = SafeString(wsData.Cells(testRow, 20).Value)
    params(17) = SafeDecimal(wsData.Cells(testRow, 40).Value)
    params(18) = SafeBoolean(wsData.Cells(testRow, 3).Value)
    params(19) = SafeBoolean(wsData.Cells(testRow, 4).Value)

    ' Show converted values
    msg = "CONVERTED VALUES:" & vbCrLf & vbCrLf
    For i = 0 To 19
        If IsNull(params(i)) Then
            msg = msg & "params(" & i & ") = NULL" & vbCrLf
        Else
            msg = msg & "params(" & i & ") = [" & params(i) & "] Type=" & TypeName(params(i)) & vbCrLf
        End If
    Next i

    MsgBox msg, vbInformation, "Converted Parameter Values"

    ' Try to connect and insert
    Set conn = GetDBConnection()
    If conn Is Nothing Then
        MsgBox "Failed to connect to database", vbCritical
        Exit Sub
    End If

    ' Truncate staging first
    conn.Execute "TRUNCATE TABLE dbo.tbl_pif_projects_staging"

    ' Try inserting with detailed parameter info
    MsgBox "About to call stored procedure. Click OK to continue...", vbInformation

    ' Call the stored procedure with explicit error handling
    On Error Resume Next
    Dim result As ADODB.Recordset
    Set result = ExecuteStoredProcedure(conn, "usp_insert_project_staging", False, _
        "@pif_id", 200, 1, 16, params(0), _
        "@project_id", 200, 1, 10, params(1), _
        "@status", 200, 1, 58, params(2), _
        "@change_type", 200, 1, 12, params(3), _
        "@accounting_treatment", 200, 1, 14, params(4), _
        "@category", 200, 1, 26, params(5), _
        "@seg", 3, 1, 0, params(6), _
        "@opco", 200, 1, 4, params(7), _
        "@site", 200, 1, 4, params(8), _
        "@strategic_rank", 200, 1, 26, params(9), _
        "@funding_project", 200, 1, 10, params(10), _
        "@project_name", 200, 1, 35, params(11), _
        "@original_fp_isd", 200, 1, 8, params(12), _
        "@revised_fp_isd", 200, 1, 5, params(13), _
        "@moving_isd_year", 129, 1, 1, params(14), _
        "@lcm_issue", 200, 1, 11, params(15), _
        "@justification", 200, 1, 192, params(16), _
        "@prior_year_spend", 131, 1, 0, params(17), _
        "@archive_flag", 11, 1, 0, params(18), _
        "@include_flag", 11, 1, 0, params(19))

    If Err.Number <> 0 Then
        MsgBox "STORED PROCEDURE FAILED!" & vbCrLf & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
               "This is the exact error occurring in BulkInsertToStaging", _
               vbCritical, "Diagnostic Result"
        Err.Clear
    Else
        MsgBox "SUCCESS! Row inserted without errors." & vbCrLf & vbCrLf & _
               "Check tbl_pif_projects_staging to verify data.", _
               vbInformation, "Diagnostic Result"
    End If

    conn.Close
    Set conn = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Diagnostic test failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "This error occurred BEFORE calling the stored procedure.", _
           vbCritical, "Diagnostic Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: ShowValidationErrors
' Purpose: Display all validation errors in a message box
' ----------------------------------------------------------------------------
Public Sub Diag_ShowErrors()
    Dim wsValidation As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim msg As String

    On Error Resume Next
    Set wsValidation = ThisWorkbook.Sheets(SHEET_VALIDATION_REPORT)
    On Error GoTo 0

    If wsValidation Is Nothing Then
        MsgBox "No validation report found. Run validation first.", vbExclamation
        Exit Sub
    End If

    lastRow = wsValidation.Cells(wsValidation.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No validation errors found.", vbInformation
        Exit Sub
    End If

    msg = "VALIDATION ERRORS:" & vbCrLf & vbCrLf
    For i = 2 To lastRow
        msg = msg & i - 1 & ". " & wsValidation.Cells(i, 1).Value & vbCrLf
        If Len(msg) > 1500 Then
            MsgBox msg, vbExclamation, "Validation Errors (Part 1)"
            msg = ""
        End If
    Next i

    If msg <> "" Then
        MsgBox msg, vbExclamation, "Validation Errors"
    End If
End Sub

' ----------------------------------------------------------------------------
' Sub: TestParameterTypes
' Purpose: Test each parameter type individually
' ----------------------------------------------------------------------------
Public Sub Diag_TestParameters()
    Dim msg As String
    Dim testVal As Variant

    msg = "PARAMETER TYPE TESTS:" & vbCrLf & vbCrLf

    ' Test SafeString
    testVal = SafeString("")
    msg = msg & "SafeString(empty): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    testVal = SafeString("  TEST  ")
    msg = msg & "SafeString('  TEST  '): [" & testVal & "]" & vbCrLf

    ' Test SafeInteger
    testVal = SafeInteger("")
    msg = msg & "SafeInteger(empty): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    testVal = SafeInteger(123)
    msg = msg & "SafeInteger(123): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    testVal = SafeInteger("abc")
    msg = msg & "SafeInteger('abc'): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    ' Test SafeDecimal
    testVal = SafeDecimal("")
    msg = msg & "SafeDecimal(empty): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    testVal = SafeDecimal(123.45)
    msg = msg & "SafeDecimal(123.45): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    ' Test SafeBoolean
    testVal = SafeBoolean("")
    msg = msg & "SafeBoolean(empty): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    testVal = SafeBoolean("Y")
    msg = msg & "SafeBoolean('Y'): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    testVal = SafeBoolean("N")
    msg = msg & "SafeBoolean('N'): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    testVal = SafeBoolean(True)
    msg = msg & "SafeBoolean(TRUE): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    MsgBox msg, vbInformation, "Type Conversion Tests"
End Sub

' ============================================================================
' BACKWARD COMPATIBILITY WRAPPERS
' ============================================================================

Public Sub TestSingleRowInsert()
    Call Diag_TestInsert
End Sub

Public Sub ShowValidationErrors()
    Call Diag_ShowErrors
End Sub

Public Sub TestParameterTypes()
    Call Diag_TestParameters
End Sub
=======
Attribute VB_Name = "mod_Diagnostic"
' ============================================================================
' MODULE: mod_Diagnostic
' ============================================================================
' Purpose: Diagnose type mismatch errors in parameter passing
' Usage: Run TestSingleRowInsert() to test first data row
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Sub: TestSingleRowInsert
' Purpose: Test inserting a single row with detailed diagnostics
' Usage: Run this from VBA editor to see exactly where the error occurs
' ----------------------------------------------------------------------------
Public Sub Diag_TestInsert()
    On Error GoTo ErrHandler

    Dim wsData As Worksheet
    Dim conn As ADODB.Connection
    Dim testRow As Long
    Dim params(0 To 19) As Variant
    Dim i As Integer
    Dim msg As String

    ' Get the PIF sheet
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    testRow = 4  ' First data row

    ' Show what we're reading
    msg = "DIAGNOSTIC: Reading data from row " & testRow & vbCrLf & vbCrLf

    ' Read and display each parameter
    msg = msg & "Column G (pif_id): " & wsData.Cells(testRow, 7).Value & vbCrLf
    msg = msg & "Column M (project_id): " & wsData.Cells(testRow, 13).Value & vbCrLf
    msg = msg & "Column R (status): " & wsData.Cells(testRow, 18).Value & vbCrLf
    msg = msg & "Column F (change_type): " & wsData.Cells(testRow, 6).Value & vbCrLf
    msg = msg & "Column E (accounting_treatment): " & wsData.Cells(testRow, 5).Value & vbCrLf
    msg = msg & "Column S (category): " & wsData.Cells(testRow, 19).Value & vbCrLf
    msg = msg & "Column H (seg): [" & wsData.Cells(testRow, 8).Value & "] IsNumeric=" & IsNumeric(wsData.Cells(testRow, 8).Value) & vbCrLf
    msg = msg & "Column I (opco): " & wsData.Cells(testRow, 9).Value & vbCrLf
    msg = msg & "Column J (site): " & wsData.Cells(testRow, 10).Value & vbCrLf
    msg = msg & "Column K (strategic_rank): " & wsData.Cells(testRow, 11).Value & vbCrLf
    msg = msg & "Column M (funding_project): " & wsData.Cells(testRow, 13).Value & vbCrLf
    msg = msg & "Column N (project_name): " & wsData.Cells(testRow, 14).Value & vbCrLf
    msg = msg & "Column O (original_fp_isd): " & wsData.Cells(testRow, 15).Value & vbCrLf
    msg = msg & "Column P (revised_fp_isd): " & wsData.Cells(testRow, 16).Value & vbCrLf
    msg = msg & "Column AM (moving_isd_year): " & wsData.Cells(testRow, 39).Value & vbCrLf
    msg = msg & "Column Q (lcm_issue): " & wsData.Cells(testRow, 17).Value & vbCrLf
    msg = msg & "Column T (justification): " & wsData.Cells(testRow, 20).Value & vbCrLf
    msg = msg & "Column AN (prior_year_spend): [" & wsData.Cells(testRow, 40).Value & "] IsNumeric=" & IsNumeric(wsData.Cells(testRow, 40).Value) & vbCrLf
    msg = msg & "Column C (archive_flag): " & wsData.Cells(testRow, 3).Value & vbCrLf
    msg = msg & "Column D (include_flag): " & wsData.Cells(testRow, 4).Value & vbCrLf

    MsgBox msg, vbInformation, "Raw Excel Values"

    ' Now convert using Safe functions and show results
    params(0) = SafeString(wsData.Cells(testRow, 7).Value)
    params(1) = SafeString(wsData.Cells(testRow, 13).Value)
    params(2) = SafeString(wsData.Cells(testRow, 18).Value)
    params(3) = SafeString(wsData.Cells(testRow, 6).Value)
    params(4) = SafeString(wsData.Cells(testRow, 5).Value)
    params(5) = SafeString(wsData.Cells(testRow, 19).Value)
    params(6) = SafeInteger(wsData.Cells(testRow, 8).Value)
    params(7) = SafeString(wsData.Cells(testRow, 9).Value)
    params(8) = SafeString(wsData.Cells(testRow, 10).Value)
    params(9) = SafeString(wsData.Cells(testRow, 11).Value)
    params(10) = SafeString(wsData.Cells(testRow, 13).Value)
    params(11) = SafeString(wsData.Cells(testRow, 14).Value)
    params(12) = SafeString(wsData.Cells(testRow, 15).Value)
    params(13) = SafeString(wsData.Cells(testRow, 16).Value)
    params(14) = SafeString(wsData.Cells(testRow, 39).Value)
    params(15) = SafeString(wsData.Cells(testRow, 17).Value)
    params(16) = SafeString(wsData.Cells(testRow, 20).Value)
    params(17) = SafeDecimal(wsData.Cells(testRow, 40).Value)
    params(18) = SafeBoolean(wsData.Cells(testRow, 3).Value)
    params(19) = SafeBoolean(wsData.Cells(testRow, 4).Value)

    ' Show converted values
    msg = "CONVERTED VALUES:" & vbCrLf & vbCrLf
    For i = 0 To 19
        If IsNull(params(i)) Then
            msg = msg & "params(" & i & ") = NULL" & vbCrLf
        Else
            msg = msg & "params(" & i & ") = [" & params(i) & "] Type=" & TypeName(params(i)) & vbCrLf
        End If
    Next i

    MsgBox msg, vbInformation, "Converted Parameter Values"

    ' Try to connect and insert
    Set conn = GetDBConnection()
    If conn Is Nothing Then
        MsgBox "Failed to connect to database", vbCritical
        Exit Sub
    End If

    ' Truncate staging first
    conn.Execute "TRUNCATE TABLE dbo.tbl_pif_projects_staging"

    ' Try inserting with detailed parameter info
    MsgBox "About to call stored procedure. Click OK to continue...", vbInformation

    ' Call the stored procedure with explicit error handling
    On Error Resume Next
    Dim result As ADODB.Recordset
    Set result = ExecuteStoredProcedure(conn, "usp_insert_project_staging", False, _
        "@pif_id", 200, 1, 16, params(0), _
        "@project_id", 200, 1, 10, params(1), _
        "@status", 200, 1, 58, params(2), _
        "@change_type", 200, 1, 12, params(3), _
        "@accounting_treatment", 200, 1, 14, params(4), _
        "@category", 200, 1, 26, params(5), _
        "@seg", 3, 1, 0, params(6), _
        "@opco", 200, 1, 4, params(7), _
        "@site", 200, 1, 4, params(8), _
        "@strategic_rank", 200, 1, 26, params(9), _
        "@funding_project", 200, 1, 10, params(10), _
        "@project_name", 200, 1, 35, params(11), _
        "@original_fp_isd", 200, 1, 8, params(12), _
        "@revised_fp_isd", 200, 1, 5, params(13), _
        "@moving_isd_year", 129, 1, 1, params(14), _
        "@lcm_issue", 200, 1, 11, params(15), _
        "@justification", 200, 1, 192, params(16), _
        "@prior_year_spend", 131, 1, 0, params(17), _
        "@archive_flag", 11, 1, 0, params(18), _
        "@include_flag", 11, 1, 0, params(19))

    If Err.Number <> 0 Then
        MsgBox "STORED PROCEDURE FAILED!" & vbCrLf & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
               "This is the exact error occurring in BulkInsertToStaging", _
               vbCritical, "Diagnostic Result"
        Err.Clear
    Else
        MsgBox "SUCCESS! Row inserted without errors." & vbCrLf & vbCrLf & _
               "Check tbl_pif_projects_staging to verify data.", _
               vbInformation, "Diagnostic Result"
    End If

    conn.Close
    Set conn = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Diagnostic test failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "This error occurred BEFORE calling the stored procedure.", _
           vbCritical, "Diagnostic Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: ShowValidationErrors
' Purpose: Display all validation errors in a message box
' ----------------------------------------------------------------------------
Public Sub Diag_ShowErrors()
    Dim wsValidation As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim msg As String

    On Error Resume Next
    Set wsValidation = ThisWorkbook.Sheets(SHEET_VALIDATION_REPORT)
    On Error GoTo 0

    If wsValidation Is Nothing Then
        MsgBox "No validation report found. Run validation first.", vbExclamation
        Exit Sub
    End If

    lastRow = wsValidation.Cells(wsValidation.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No validation errors found.", vbInformation
        Exit Sub
    End If

    msg = "VALIDATION ERRORS:" & vbCrLf & vbCrLf
    For i = 2 To lastRow
        msg = msg & i - 1 & ". " & wsValidation.Cells(i, 1).Value & vbCrLf
        If Len(msg) > 1500 Then
            MsgBox msg, vbExclamation, "Validation Errors (Part 1)"
            msg = ""
        End If
    Next i

    If msg <> "" Then
        MsgBox msg, vbExclamation, "Validation Errors"
    End If
End Sub

' ----------------------------------------------------------------------------
' Sub: TestParameterTypes
' Purpose: Test each parameter type individually
' ----------------------------------------------------------------------------
Public Sub Diag_TestParameters()
    Dim msg As String
    Dim testVal As Variant

    msg = "PARAMETER TYPE TESTS:" & vbCrLf & vbCrLf

    ' Test SafeString
    testVal = SafeString("")
    msg = msg & "SafeString(empty): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    testVal = SafeString("  TEST  ")
    msg = msg & "SafeString('  TEST  '): [" & testVal & "]" & vbCrLf

    ' Test SafeInteger
    testVal = SafeInteger("")
    msg = msg & "SafeInteger(empty): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    testVal = SafeInteger(123)
    msg = msg & "SafeInteger(123): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    testVal = SafeInteger("abc")
    msg = msg & "SafeInteger('abc'): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    ' Test SafeDecimal
    testVal = SafeDecimal("")
    msg = msg & "SafeDecimal(empty): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    testVal = SafeDecimal(123.45)
    msg = msg & "SafeDecimal(123.45): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    ' Test SafeBoolean
    testVal = SafeBoolean("")
    msg = msg & "SafeBoolean(empty): " & IIf(IsNull(testVal), "NULL", testVal) & vbCrLf

    testVal = SafeBoolean("Y")
    msg = msg & "SafeBoolean('Y'): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    testVal = SafeBoolean("N")
    msg = msg & "SafeBoolean('N'): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    testVal = SafeBoolean(True)
    msg = msg & "SafeBoolean(TRUE): " & testVal & " Type=" & TypeName(testVal) & vbCrLf

    MsgBox msg, vbInformation, "Type Conversion Tests"
End Sub

' ============================================================================
' BACKWARD COMPATIBILITY WRAPPERS
' ============================================================================

Public Sub TestSingleRowInsert()
    Call Diag_TestInsert
End Sub

Public Sub ShowValidationErrors()
    Call Diag_ShowErrors
End Sub

Public Sub TestParameterTypes()
    Call Diag_TestParameters
End Sub
>>>>>>> 254c5f99beb56416a0a77e8c6100e047b5b2e663
