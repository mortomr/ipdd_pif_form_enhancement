Attribute VB_Name = "mod_Validation"
' ============================================================================
' MODULE: mod_Validation (PERFORMANCE OPTIMIZED)
' ============================================================================
' Purpose: Validate PIF data before submission to database
' Author: Data Architecture Team
' Date: 2025-11-13
'
' PERFORMANCE: Array-based single-pass validation
' SPEEDUP: 4 separate loops → 1 array-based pass (4x faster)
' ============================================================================

Option Explicit

' Sheet name constants
Private Const SHEET_DATA As String = "PIF"
Private Const SHEET_VALIDATION_REPORT As String = "Validation_Report"

' Column mappings (using shared constants)
Private Const COL_ARCHIVE As Integer = 3            ' C
Private Const COL_INCLUDE As Integer = 4            ' D
Private Const COL_ACCOUNTING As Integer = 5         ' E
Private Const COL_CHANGE_TYPE As Integer = 6        ' F
Private Const COL_PIF_ID As Integer = 7             ' G
Private Const COL_SEG As Integer = 8                ' H
Private Const COL_OPCO As Integer = 9               ' I
Private Const COL_SITE As Integer = 10              ' J
Private Const COL_STRATEGIC_RANK As Integer = 11    ' K
Private Const COL_FROM_BLANKET As Integer = 12      ' L
Private Const COL_FUNDING_PROJECT As Integer = 13   ' M
Private Const COL_PROJECT_NAME As Integer = 14      ' N
Private Const COL_ORIGINAL_ISD As Integer = 15      ' O
Private Const COL_REVISED_ISD As Integer = 16       ' P
Private Const COL_LCM_ISSUE As Integer = 17         ' Q
Private Const COL_STATUS As Integer = 18            ' R
Private Const COL_CATEGORY As Integer = 19          ' S
Private Const COL_JUSTIFICATION As Integer = 20     ' T

' ============================================================================
' PUBLIC FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: ValidateData
' Purpose: Master validation function - array-based single-pass
' PERFORMANCE: Reads data once into array, validates all rules in one loop
' ----------------------------------------------------------------------------
Public Function ValidateData(Optional ByVal showSuccessMessage As Boolean = True) As Boolean
    On Error GoTo ErrHandler

    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim errors As Collection
    Dim startTime As Double
    Dim dataArray As Variant
    Dim lastRow As Long
    Dim rowNum As Long
    Dim actualRow As Long
    Dim pifId As String, projectId As String
    Dim status As String, justification As String
    Dim changeType As String
    Dim seenKeys As Object
    Dim key As String

    startTime = Timer

    Set errors = New Collection
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    Set wsReport = ThisWorkbook.Sheets(SHEET_VALIDATION_REPORT)
    Set seenKeys = CreateObject("Scripting.Dictionary")

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

    Application.StatusBar = "Running validation..."
    Application.ScreenUpdating = False

    ' Find last row
    lastRow = wsData.Cells(wsData.Rows.Count, COL_PIF_ID).End(xlUp).Row
    If lastRow < 4 Then
        ' No data to validate
        ValidateData = True
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Function
    End If

    ' Read entire data range into array (ONE READ OPERATION - FAST!)
    dataArray = wsData.Range(wsData.Cells(4, 1), wsData.Cells(lastRow, 20)).Value

    ' SINGLE-PASS VALIDATION (ALL RULES IN ONE LOOP - SUPER FAST!)
    For rowNum = 1 To UBound(dataArray, 1)
        actualRow = rowNum + 3 ' Offset for header rows

        ' Get key fields
        pifId = Trim(dataArray(rowNum, COL_PIF_ID) & "")
        projectId = Trim(dataArray(rowNum, COL_FUNDING_PROJECT) & "")

        ' Skip empty rows
        If pifId = "" Then GoTo NextRow

        ' -------------------------------------------------------------------
        ' VALIDATION RULE 1: Required Fields
        ' -------------------------------------------------------------------
        If projectId = "" Then
            errors.Add "Row " & actualRow & "|Missing Required Field|Project ID is required"
        End If

        changeType = Trim(dataArray(rowNum, COL_CHANGE_TYPE) & "")
        If changeType = "" Then
            errors.Add "Row " & actualRow & "|Missing Required Field|Change Type is required"
        End If

        ' -------------------------------------------------------------------
        ' VALIDATION RULE 2: Data Types and Field Lengths
        ' -------------------------------------------------------------------
        If Len(pifId) > 16 Then
            errors.Add "Row " & actualRow & "|Field Too Long|PIF_ID exceeds 16 characters (currently " & Len(pifId) & " chars)"
        End If

        If Len(projectId) > 10 Then
            errors.Add "Row " & actualRow & "|Field Too Long|FUNDING_PROJECT exceeds 10 characters (currently " & Len(projectId) & " chars)"
        End If

        status = Trim(dataArray(rowNum, COL_STATUS) & "")
        If Len(status) > 58 Then
            errors.Add "Row " & actualRow & "|Field Too Long|STATUS exceeds 58 characters (currently " & Len(status) & " chars)"
        End If

        Dim site As String
        site = Trim(dataArray(rowNum, COL_SITE) & "")
        If Len(site) > 4 Then
            errors.Add "Row " & actualRow & "|Field Too Long|SITE exceeds 4 characters (currently " & Len(site) & " chars)"
        End If

        ' Check SEG is numeric
        Dim segValue As Variant
        segValue = dataArray(rowNum, COL_SEG)
        If Not IsEmpty(segValue) And segValue <> "" Then
            If Not IsNumeric(segValue) Then
                errors.Add "Row " & actualRow & "|Invalid Data Type|SEG must be numeric"
            End If
        End If

        ' -------------------------------------------------------------------
        ' VALIDATION RULE 3: Business Rules
        ' -------------------------------------------------------------------
        justification = Trim(dataArray(rowNum, COL_JUSTIFICATION) & "")
        If UCase(status) = "APPROVED" And justification = "" Then
            errors.Add "Row " & actualRow & "|Business Rule Violation|Approved PIFs require justification"
        End If

        ' -------------------------------------------------------------------
        ' VALIDATION RULE 4: Duplicate Detection
        ' -------------------------------------------------------------------
        If pifId <> "" And projectId <> "" Then
            key = pifId & "|" & projectId
            If seenKeys.Exists(key) Then
                errors.Add "Row " & actualRow & "|Duplicate Entry|PIF " & pifId & " + Project " & projectId & _
                          " appears multiple times (first occurrence: Row " & seenKeys(key) & ")"
            Else
                seenKeys.Add key, actualRow
            End If
        End If

NextRow:
    Next rowNum

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

        If showSuccessMessage Then
            Dim elapsed As Double
            elapsed = Timer - startTime

            MsgBox "Validation passed successfully!" & vbCrLf & vbCrLf & _
                   "Elapsed time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & _
                   "Ready to submit to database.", _
                   vbInformation, "Validation Passed"
        End If

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

    Application.StatusBar = "Running SQL validation..."

    ' Call stored procedure to validate staging data
    sql = "DECLARE @ErrorCount INT; " & _
          "EXEC dbo.usp_validate_staging_data @ErrorCount OUTPUT; " & _
          "SELECT @ErrorCount AS ErrorCount"

    Set rs = mod_Database.GetRecordset(sql)

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            errorCount = rs.Fields("ErrorCount").Value
        End If
        rs.Close
    End If

    If errorCount > 0 Then
        ' Get detailed errors
        sql = "DECLARE @ErrorCount INT; EXEC dbo.usp_validate_staging_data @ErrorCount OUTPUT"
        Set rs = mod_Database.GetRecordset(sql)

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
' PRIVATE HELPER FUNCTIONS
' ============================================================================

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
