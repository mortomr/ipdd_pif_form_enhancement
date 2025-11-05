Attribute VB_Name = "mod_Validation"
' ============================================================================
' MODULE: mod_Validation
' ============================================================================
' Purpose: Validate PIF data before submission to database
' Author: Data Architecture Team
' Date: 2025-11-05
'
' Validation Checks:
'   1. Required field presence (PIF ID, Project ID, etc.)
'   2. Data type validation
'   3. Business rule validation (e.g., approved PIFs need justification)
'   4. Duplicate detection
'   5. SQL-side validation after staging
' ============================================================================

Option Explicit

' Sheet name constants - UPDATE IF YOUR SHEETS HAVE DIFFERENT NAMES
Private Const SHEET_DATA As String = "PIF_Data"
Private Const SHEET_VALIDATION_REPORT As String = "Validation_Report"

' Column mappings - UPDATE BASED ON YOUR EXCEL LAYOUT
Private Const COL_ARCHIVE As Integer = 3           ' C
Private Const COL_INCLUDE As Integer = 4           ' D
Private Const COL_ACCOUNTING As Integer = 5        ' E
Private Const COL_CHANGE_TYPE As Integer = 6       ' F
Private Const COL_PIF_ID As Integer = 7            ' G
Private Const COL_SEG As Integer = 8               ' H
Private Const COL_OPCO As Integer = 9              ' I
Private Const COL_SITE As Integer = 10             ' J
Private Const COL_STRATEGIC_RANK As Integer = 11   ' K
Private Const COL_FROM_BLANKET As Integer = 12     ' L
Private Const COL_FUNDING_PROJECT As Integer = 13  ' M
Private Const COL_PROJECT_NAME As Integer = 14     ' N
Private Const COL_ORIGINAL_ISD As Integer = 15     ' O
Private Const COL_REVISED_ISD As Integer = 16      ' P
Private Const COL_LCM_ISSUE As Integer = 17        ' Q
Private Const COL_STATUS As Integer = 18           ' R
Private Const COL_CATEGORY As Integer = 19         ' S
Private Const COL_JUSTIFICATION As Integer = 20    ' T

' ============================================================================
' PUBLIC FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: ValidateData
' Purpose: Master validation function - runs all validation checks
' Returns: True if validation passed, False if errors found
' ----------------------------------------------------------------------------
Public Function ValidateData() As Boolean
    On Error GoTo ErrHandler
    
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim errors As Collection
    Dim startTime As Double
    
    startTime = Timer
    
    ' Initialize
    Set errors = New Collection
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    Set wsReport = ThisWorkbook.Sheets(SHEET_VALIDATION_REPORT)
    
    ' Clear previous validation report
    wsReport.Cells.Clear
    wsReport.Range("A1").Value = "PIF Validation Report"
    wsReport.Range("B1").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsReport.Range("A2").Value = "Row"
    wsReport.Range("B2").Value = "Error Type"
    wsReport.Range("C2").Value = "Error Description"
    
    ' Format headers
    With wsReport.Range("A1:C2")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    Application.StatusBar = "Starting validation..."
    Application.ScreenUpdating = False
    
    ' Run validation checks
    Call ValidateRequiredFields(wsData, errors)
    Call ValidateDataTypes(wsData, errors)
    Call ValidateBusinessRules(wsData, errors)
    Call ValidateDuplicates(wsData, errors)
    
    ' Write errors to report
    If errors.Count > 0 Then
        Call WriteErrorsToReport(wsReport, errors)
        
        MsgBox errors.Count & " validation error(s) found." & vbCrLf & vbCrLf & _
               "Please review the Validation_Report sheet and fix issues before submitting.", _
               vbExclamation, "Validation Failed"
        
        wsReport.Activate
        ValidateData = False
    Else
        wsReport.Range("A4").Value = "No errors found - data is ready for submission!"
        wsReport.Range("A4").Font.Color = RGB(0, 128, 0)
        wsReport.Range("A4").Font.Bold = True
        
        Dim elapsed As Double
        elapsed = Timer - startTime
        
        MsgBox "Validation passed successfully!" & vbCrLf & vbCrLf & _
               "Elapsed time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & _
               "Ready to submit to database.", _
               vbInformation, "Validation Passed"
        
        ValidateData = True
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Function
    
ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Validation process failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Validation Error"
    ValidateData = False
End Function

' ----------------------------------------------------------------------------
' Function: ValidateStagingData
' Purpose: Run SQL-side validation after data is staged
' Returns: True if no errors, False if errors found
' ----------------------------------------------------------------------------
Public Function ValidateStagingData() As Boolean
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim errorCount As Long
    Dim wsReport As Worksheet
    Dim rowNum As Long
    
    Set wsReport = ThisWorkbook.Sheets(SHEET_VALIDATION_REPORT)
    
    ' Clear previous SQL validation errors
    Application.StatusBar = "Running SQL validation..."
    
    ' Call stored procedure to validate staging data
    sql = "DECLARE @ErrorCount INT; " & _
          "EXEC dbo.usp_validate_staging_data @ErrorCount OUTPUT; " & _
          "SELECT @ErrorCount AS ErrorCount"
    
    Set rs = GetRecordset(sql)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            errorCount = rs.Fields("ErrorCount").Value
        End If
        rs.Close
    End If
    
    If errorCount > 0 Then
        ' Get detailed errors
        sql = "DECLARE @ErrorCount INT; EXEC dbo.usp_validate_staging_data @ErrorCount OUTPUT"
        Set rs = GetRecordset(sql)
        
        ' Write SQL errors to validation report
        If Not rs Is Nothing Then
            rowNum = wsReport.Cells(wsReport.Rows.Count, 1).End(xlUp).Row + 2
            wsReport.Cells(rowNum, 1).Value = "SQL VALIDATION ERRORS:"
            wsReport.Cells(rowNum, 1).Font.Bold = True
            rowNum = rowNum + 1
            
            Do While Not rs.EOF
                wsReport.Cells(rowNum, 1).Value = rs.Fields("error_id").Value
                wsReport.Cells(rowNum, 2).Value = rs.Fields("error_type").Value
                wsReport.Cells(rowNum, 3).Value = rs.Fields("error_message").Value
                rowNum = rowNum + 1
                rs.MoveNext
            Loop
            
            rs.Close
        End If
        
        wsReport.Activate
        MsgBox errorCount & " SQL validation error(s) found." & vbCrLf & vbCrLf & _
               "See Validation_Report sheet for details.", _
               vbExclamation, "SQL Validation Failed"
        
        ValidateStagingData = False
    Else
        ValidateStagingData = True
    End If
    
    Application.StatusBar = False
    Exit Function
    
ErrHandler:
    Application.StatusBar = False
    MsgBox "SQL validation failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "SQL Validation Error"
    ValidateStagingData = False
End Function

' ============================================================================
' PRIVATE VALIDATION FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ValidateRequiredFields
' Purpose: Check that required fields are not empty
' ----------------------------------------------------------------------------
Private Sub ValidateRequiredFields(ByVal wsData As Worksheet, ByRef errors As Collection)
    Dim lastRow As Long
    Dim i As Long
    Dim errorMsg As String
    
    lastRow = wsData.Cells(wsData.Rows.Count, COL_PIF_ID).End(xlUp).Row
    
    Application.StatusBar = "Validating required fields..."
    
    For i = 2 To lastRow
        ' Skip completely empty rows
        If WorksheetFunction.CountA(wsData.Rows(i)) = 0 Then GoTo NextRow
        
        ' Check PIF ID
        If IsEmpty(wsData.Cells(i, COL_PIF_ID).Value) Or _
           Trim(wsData.Cells(i, COL_PIF_ID).Value) = "" Then
            errors.Add "Row " & i & "|Missing Required Field|PIF ID is required"
        End If
        
        ' Check Project ID (assuming it's in column after project name for now)
        ' NOTE: Adjust this based on where project_id actually is
        If IsEmpty(wsData.Cells(i, COL_FUNDING_PROJECT).Value) Or _
           Trim(wsData.Cells(i, COL_FUNDING_PROJECT).Value) = "" Then
            errors.Add "Row " & i & "|Missing Required Field|Project ID is required"
        End If
        
        ' Check Change Type
        If IsEmpty(wsData.Cells(i, COL_CHANGE_TYPE).Value) Or _
           Trim(wsData.Cells(i, COL_CHANGE_TYPE).Value) = "" Then
            errors.Add "Row " & i & "|Missing Required Field|Change Type is required"
        End If
        
        ' Progress indicator
        If i Mod 50 = 0 Then
            Application.StatusBar = "Validating required fields... Row " & i & " of " & lastRow
        End If
        
NextRow:
    Next i
End Sub

' ----------------------------------------------------------------------------
' Sub: ValidateDataTypes
' Purpose: Check that numeric fields contain valid numbers
' ----------------------------------------------------------------------------
Private Sub ValidateDataTypes(ByVal wsData As Worksheet, ByRef errors As Collection)
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = wsData.Cells(wsData.Rows.Count, COL_PIF_ID).End(xlUp).Row
    
    Application.StatusBar = "Validating data types..."
    
    For i = 2 To lastRow
        ' Skip completely empty rows
        If WorksheetFunction.CountA(wsData.Rows(i)) = 0 Then GoTo NextRow
        
        ' Check SEG (should be numeric)
        If Not IsEmpty(wsData.Cells(i, COL_SEG).Value) Then
            If Not IsNumeric(wsData.Cells(i, COL_SEG).Value) Then
                errors.Add "Row " & i & "|Invalid Data Type|SEG must be numeric"
            End If
        End If
        
        ' Add additional data type checks as needed
        
        ' Progress indicator
        If i Mod 50 = 0 Then
            Application.StatusBar = "Validating data types... Row " & i & " of " & lastRow
        End If
        
NextRow:
    Next i
End Sub

' ----------------------------------------------------------------------------
' Sub: ValidateBusinessRules
' Purpose: Check business logic rules
' ----------------------------------------------------------------------------
Private Sub ValidateBusinessRules(ByVal wsData As Worksheet, ByRef errors As Collection)
    Dim lastRow As Long
    Dim i As Long
    Dim status As String
    Dim justification As String
    
    lastRow = wsData.Cells(wsData.Rows.Count, COL_PIF_ID).End(xlUp).Row
    
    Application.StatusBar = "Validating business rules..."
    
    For i = 2 To lastRow
        ' Skip completely empty rows
        If WorksheetFunction.CountA(wsData.Rows(i)) = 0 Then GoTo NextRow
        
        status = Trim(UCase(wsData.Cells(i, COL_STATUS).Value))
        justification = Trim(wsData.Cells(i, COL_JUSTIFICATION).Value)
        
        ' Rule 1: Approved PIFs must have justification
        If status = "APPROVED" Then
            If justification = "" Then
                errors.Add "Row " & i & "|Business Rule Violation|Approved PIFs require justification"
            End If
        End If
        
        ' Rule 2: Add more business rules as needed
        
        ' Progress indicator
        If i Mod 50 = 0 Then
            Application.StatusBar = "Validating business rules... Row " & i & " of " & lastRow
        End If
        
NextRow:
    Next i
End Sub

' ----------------------------------------------------------------------------
' Sub: ValidateDuplicates
' Purpose: Check for duplicate PIF + Project combinations
' ----------------------------------------------------------------------------
Private Sub ValidateDuplicates(ByVal wsData As Worksheet, ByRef errors As Collection)
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim pifId As String, projectId As String
    Dim pifId2 As String, projectId2 As String
    Dim dict As Object
    Dim key As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = wsData.Cells(wsData.Rows.Count, COL_PIF_ID).End(xlUp).Row
    
    Application.StatusBar = "Checking for duplicates..."
    
    For i = 2 To lastRow
        ' Skip completely empty rows
        If WorksheetFunction.CountA(wsData.Rows(i)) = 0 Then GoTo NextRow
        
        pifId = Trim(wsData.Cells(i, COL_PIF_ID).Value)
        projectId = Trim(wsData.Cells(i, COL_FUNDING_PROJECT).Value)  ' Adjust column as needed
        
        If pifId <> "" And projectId <> "" Then
            key = pifId & "|" & projectId
            
            If dict.Exists(key) Then
                errors.Add "Row " & i & "|Duplicate Entry|PIF " & pifId & " + Project " & projectId & _
                          " appears multiple times (first occurrence: Row " & dict(key) & ")"
            Else
                dict.Add key, i
            End If
        End If
        
        ' Progress indicator
        If i Mod 50 = 0 Then
            Application.StatusBar = "Checking for duplicates... Row " & i & " of " & lastRow
        End If
        
NextRow:
    Next i
End Sub

' ----------------------------------------------------------------------------
' Sub: WriteErrorsToReport
' Purpose: Write collected errors to validation report sheet
' ----------------------------------------------------------------------------
Private Sub WriteErrorsToReport(ByVal wsReport As Worksheet, ByRef errors As Collection)
    Dim i As Long
    Dim errorParts() As String
    Dim rowNum As Long
    
    rowNum = 3  ' Start below headers
    
    For i = 1 To errors.Count
        ' Error format: "Row X|ErrorType|ErrorDescription"
        errorParts = Split(errors(i), "|")
        
        If UBound(errorParts) >= 2 Then
            wsReport.Cells(rowNum, 1).Value = errorParts(0)  ' Row number
            wsReport.Cells(rowNum, 2).Value = errorParts(1)  ' Error type
            wsReport.Cells(rowNum, 3).Value = errorParts(2)  ' Description
        Else
            wsReport.Cells(rowNum, 1).Value = i
            wsReport.Cells(rowNum, 2).Value = "Unknown"
            wsReport.Cells(rowNum, 3).Value = errors(i)
        End If
        
        rowNum = rowNum + 1
    Next i
    
    ' Format report
    wsReport.Columns("A:C").AutoFit
    wsReport.Range("A2:C" & rowNum - 1).AutoFilter
End Sub

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: IsValidDate
' Purpose: Check if a value can be converted to a date
' ----------------------------------------------------------------------------
Private Function IsValidDate(ByVal value As Variant) As Boolean
    On Error Resume Next
    IsValidDate = IsDate(value)
    On Error GoTo 0
End Function
