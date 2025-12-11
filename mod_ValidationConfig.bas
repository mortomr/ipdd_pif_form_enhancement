Attribute VB_Name = "mod_ValidationConfig"
' ============================================================================
' MODULE: mod_ValidationConfig
' ============================================================================
' Purpose: Centralized validation rules configuration
' Allows rule changes without modifying the validation engine
'
' Usage:
'   - Edit the rules defined in this module to change validation behavior
'   - Add new rules by creating new functions in this module
'   - The validation engine (mod_ValidationReport) calls these rules
' ============================================================================

Option Explicit

' ============================================================================
' FIELD REQUIREMENT RULES
' ============================================================================
' These functions define which fields are required and under what conditions

' ---------------------------------------------------------------------------
' Function: IsRevisedISDRequired
' Purpose: Determine if Revised ISD field is required
' Logic: Required only if Change Type involves timing/scope changes
' ---------------------------------------------------------------------------
Public Function IsRevisedISDRequired(changeType As String) As Boolean
    ' Change types that require Revised ISD
    Dim scopeChanges As String
    scopeChanges = "Funding Increase|Funding Decrease|Scope Change|Schedule Delay|Schedule Acceleration"

    IsRevisedISDRequired = (InStr(1, scopeChanges, changeType, vbBinaryCompare) > 0)
End Function

' ---------------------------------------------------------------------------
' Function: IsLCMIssueRequired
' Purpose: Determine if LCM Issue is required
' Logic: Required only for Compliance category
' ---------------------------------------------------------------------------
Public Function IsLCMIssueRequired(category As String) As Boolean
    IsLCMIssueRequired = (UCase(category) = "COMPLIANCE")
End Function

' ---------------------------------------------------------------------------
' Function: IsJustificationRequired
' Purpose: Determine if Justification is required
' Logic: Required if Archive_Flag is set (PIF being submitted for approval)
' ---------------------------------------------------------------------------
Public Function IsJustificationRequired(archiveFlag As String) As Boolean
    ' Archive flag can be "X", "TRUE", 1, or TRUE
    IsJustificationRequired = (archiveFlag = "X" Or archiveFlag = "TRUE" Or archiveFlag = 1)
End Function

' ============================================================================
' FIELD VALIDATION RULES
' ============================================================================
' These functions validate field format/content

' ---------------------------------------------------------------------------
' Function: ValidateDateFormatRule
' Purpose: Validate date field is in MM/DD/YYYY format
' Returns: Tuple (isValid, errorMessage)
' ---------------------------------------------------------------------------
Public Function ValidateDateFormatRule(dateValue As String) As Variant
    Dim result(0 To 1) As Variant ' (isValid, errorMessage)

    If dateValue = "" Then
        result(0) = True ' Empty is valid; required check handled elsewhere
        result(1) = ""
        ValidateDateFormatRule = result
        Exit Function
    End If

    On Error Resume Next
    Dim dateObj As Date
    dateObj = CDate(dateValue)
    On Error GoTo 0

    If Err.Number <> 0 Then
        result(0) = False
        result(1) = "Invalid date format. Use MM/DD/YYYY."
        Err.Clear
    Else
        result(0) = True
        result(1) = ""
    End If

    ValidateDateFormatRule = result
End Function

' ---------------------------------------------------------------------------
' Function: ValidateTextLengthRule
' Purpose: Validate text field length
' Params: fieldValue, maxLength
' Returns: (isValid, errorMessage)
' ---------------------------------------------------------------------------
Public Function ValidateTextLengthRule(fieldValue As String, maxLength As Long) As Variant
    Dim result(0 To 1) As Variant

    If Len(fieldValue) > maxLength Then
        result(0) = False
        result(1) = "Field exceeds maximum length of " & maxLength & " characters."
    Else
        result(0) = True
        result(1) = ""
    End If

    ValidateTextLengthRule = result
End Function

' ---------------------------------------------------------------------------
' Function: ValidateNumericRule
' Purpose: Validate field is numeric
' Returns: (isValid, errorMessage)
' ---------------------------------------------------------------------------
Public Function ValidateNumericRule(fieldValue As String) As Variant
    Dim result(0 To 1) As Variant

    If fieldValue = "" Then
        result(0) = True
        result(1) = ""
        ValidateNumericRule = result
        Exit Function
    End If

    If Not IsNumeric(fieldValue) Then
        result(0) = False
        result(1) = "Field must be numeric."
    Else
        result(0) = True
        result(1) = ""
    End If

    ValidateNumericRule = result
End Function

' ---------------------------------------------------------------------------
' Function: ValidateAllowedValuesRule
' Purpose: Validate field is one of allowed values
' Params: fieldValue, allowedValuesString (pipe-delimited, e.g. "Yes|No|Maybe")
' Returns: (isValid, errorMessage)
' ---------------------------------------------------------------------------
Public Function ValidateAllowedValuesRule(fieldValue As String, allowedValuesString As String) As Variant
    Dim result(0 To 1) As Variant
    Dim allowedArray() As String
    Dim i As Long

    If fieldValue = "" Then
        result(0) = True
        result(1) = ""
        ValidateAllowedValuesRule = result
        Exit Function
    End If

    allowedArray = Split(allowedValuesString, "|")

    For i = LBound(allowedArray) To UBound(allowedArray)
        If UCase(Trim(fieldValue)) = UCase(Trim(allowedArray(i))) Then
            result(0) = True
            result(1) = ""
            ValidateAllowedValuesRule = result
            Exit Function
        End If
    Next i

    result(0) = False
    result(1) = "Invalid value. Allowed: " & allowedValuesString
    ValidateAllowedValuesRule = result
End Function

' ============================================================================
' ERROR MESSAGE RULES
' ============================================================================
' These functions return user-friendly error messages for different validation failures

' ---------------------------------------------------------------------------
' Function: GetFieldRequiredMessage
' Purpose: Return error message for missing required field
' ---------------------------------------------------------------------------
Public Function GetFieldRequiredMessage(fieldName As String, condition As String) As String
    If condition = "" Then
        GetFieldRequiredMessage = fieldName & " is required and cannot be blank."
    Else
        GetFieldRequiredMessage = fieldName & " is required when " & condition & "."
    End If
End Function

' ---------------------------------------------------------------------------
' Function: GetRevisedISDMessage
' Purpose: Return contextual error for Revised ISD
' ---------------------------------------------------------------------------
Public Function GetRevisedISDMessage(changeType As String) As String
    GetRevisedISDMessage = "Revised ISD is required for Change Type '" & changeType & "'. " & _
                          "This change type affects the project timeline and requires a new target date."
End Function

' ---------------------------------------------------------------------------
' Function: GetLCMIssueMessage
' Purpose: Return contextual error for LCM Issue
' ---------------------------------------------------------------------------
Public Function GetLCMIssueMessage() As String
    GetLCMIssueMessage = "LCM Issue is required for Compliance category. " & _
                        "Identify the relevant life-cycle management issue."
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ---------------------------------------------------------------------------
' Function: IsNumeric
' Purpose: Check if string represents a number
' ---------------------------------------------------------------------------
Private Function IsNumeric(val As String) As Boolean
    On Error Resume Next
    IsNumeric = Not IsError(val + 0)
    On Error GoTo 0
End Function
