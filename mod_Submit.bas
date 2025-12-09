Attribute VB_Name = "mod_Submit"
' ============================================================================
' MODULE: mod_Submit
' ============================================================================
' Purpose: Orchestrate the PIF submission process from Excel to SQL Server
' Author: Data Architecture Team
' Date: 2025-11-05
'
' Main Process Flow:
'   1. User confirmation
'   2. Unpivot cost data (wide to long format)
'   3. Create database backups
'   4. Upload data to staging tables
'   5. Run validation checks
'   6. Commit to inflight tables (within transaction)
'   7. Archive approved PIFs
'   8. Log submission
' ============================================================================

Option Explicit

' Sheet name constants are now defined in mod_SharedConstants
' Use the centralized constants for consistency across all modules

' ============================================================================
' PUBLIC INTERFACE
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: SaveSnapshot
' Purpose: Save current PIF data to inflight tables (working snapshot)
' Usage: Attach to [Save Snapshot] button - for ongoing work, no archival
' Notes: This is for in-progress work - does NOT archive to permanent tables
' ----------------------------------------------------------------------------
Public Sub SaveSnapshot()
    On Error GoTo ErrHandler

    Dim response As VbMsgBoxResult
    Dim startTime As Double
    Dim success As Boolean
    Dim selectedSite As String

    ' STEP 0: Validate site selection
    On Error Resume Next
    selectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.Value)
    On Error GoTo ErrHandler

    If selectedSite = "" Then
        MsgBox "Please select a site before saving." & vbCrLf & vbCrLf & _
               "Go to the Instructions worksheet and select your site from the dropdown.", _
               vbExclamation, "Site Not Selected"
        Exit Sub
    End If

    ' Block Fleet submissions (read-only access)
    If UCase(selectedSite) = "FLEET" Then
        MsgBox "Fleet cannot save data." & vbCrLf & vbCrLf & _
               "Fleet is read-only access for viewing all sites." & vbCrLf & _
               "Please select a specific site (ANO, GGN, RBS, WF3, or HQN) to save data.", _
               vbExclamation, "Fleet Submission Not Allowed"
        Exit Sub
    End If

    ' Confirmation prompt
    response = MsgBox("Save working snapshot for site: " & selectedSite & "?" & vbCrLf & vbCrLf & _
                      "This will update the inflight tables with your current work." & vbCrLf & _
                      "Records will NOT be archived to permanent storage.", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Save Snapshot")

    If response = vbNo Then
        Exit Sub
    End If

    startTime = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    success = False  ' Track overall success

    ' STEP 1: Validate Excel data FIRST (catch errors before touching database)
    Application.StatusBar = "Running validation checks..."
    If Not ValidateData(showSuccessMessage:=False) Then
        MsgBox "Validation failed - please check the Validation_Report sheet for errors." & vbCrLf & vbCrLf & _
               "Fix all validation errors and try again.", _
               vbExclamation, "Save Snapshot Failed"
        GoTo Cleanup
    End If

    ' STEP 2: Unpivot cost data
    Application.StatusBar = "Preparing cost data..."
    If Not UnpivotCostData() Then
        MsgBox "Failed to unpivot cost data." & vbCrLf & vbCrLf & _
               "Please check your data and try again.", _
               vbExclamation, "Save Snapshot Failed"
        GoTo Cleanup
    End If

    ' STEP 3: Upload to staging
    Application.StatusBar = "Uploading project data to staging..."
    If Not UploadProjectData() Then
        MsgBox "Failed to upload project data to staging." & vbCrLf & vbCrLf & _
               "Please check database connection and try again.", _
               vbExclamation, "Save Snapshot Failed"
        GoTo Cleanup
    End If

    Application.StatusBar = "Uploading cost data to staging..."
    If Not UploadCostData() Then
        MsgBox "Failed to upload cost data to staging." & vbCrLf & vbCrLf & _
               "Please check database connection and try again.", _
               vbExclamation, "Save Snapshot Failed"
        GoTo Cleanup
    End If

    ' STEP 4: Validate staging data (SQL-side validation)
    Application.StatusBar = "Running database validation checks..."
    If Not ValidateStagingData() Then
        MsgBox "Database validation failed." & vbCrLf & vbCrLf & _
               "Please check the Validation_Report sheet for errors.", _
               vbExclamation, "Save Snapshot Failed"
        GoTo Cleanup
    End If

    ' STEP 5: Commit to inflight tables
    Application.StatusBar = "Committing to database..."
    If Not CommitToInflight() Then
        MsgBox "Failed to commit data to inflight tables." & vbCrLf & vbCrLf & _
               "Database has been rolled back. Please check connection and try again.", _
               vbCritical, "Save Snapshot Failed"
        GoTo Cleanup
    End If

    ' STEP 6: Log submission
    Application.StatusBar = "Logging submission..."
    Call LogSubmission()  ' Non-critical, don't fail on logging errors

    ' STEP 7: Refresh query worksheets (silently - final message will report success)
    Application.StatusBar = "Refreshing query worksheets..."
    Call mod_WorksheetQuery.RefreshAll(showSuccessMessage:=False)

    success = True  ' Mark as successful

    ' Success!
    MsgBox "Working snapshot saved!" & vbCrLf & vbCrLf & _
           "Site: " & selectedSite & vbCrLf & _
           "Data saved to inflight tables (in-progress work)" & vbCrLf & vbCrLf & _
           "To permanently archive approved records, use [Finalize Month] button.", _
           vbInformation, "Snapshot Saved"

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Save failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "The database has been rolled back to its previous state." & vbCrLf & _
           "Please review the error and try again.", _
           vbCritical, "Save Error"
End Sub


' ----------------------------------------------------------------------------
' Sub: FinalizeMonth
' Purpose: Complete month-end workflow - save to inflight AND archive approved
' Usage: Attach to [Finalize Month] button - when all decisions are final
' Notes: This does the FULL workflow including permanent archival
' ----------------------------------------------------------------------------
Public Sub FinalizeMonth()
    On Error GoTo ErrHandler

    Dim response As VbMsgBoxResult
    Dim startTime As Double
    Dim success As Boolean
    Dim selectedSite As String

    ' STEP 0: Validate site selection
    On Error Resume Next
    selectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.Value)
    On Error GoTo ErrHandler

    If selectedSite = "" Then
        MsgBox "Please select a site before finalizing." & vbCrLf & vbCrLf & _
               "Go to the Instructions worksheet and select your site from the dropdown.", _
               vbExclamation, "Site Not Selected"
        Exit Sub
    End If

    ' Block Fleet operations (read-only access)
    If UCase(selectedSite) = "FLEET" Then
        MsgBox "Fleet cannot finalize data." & vbCrLf & vbCrLf & _
               "Fleet is read-only access for viewing all sites." & vbCrLf & _
               "Please select a specific site (ANO, GGN, RBS, WF3, or HQN).", _
               vbExclamation, "Fleet Operation Not Allowed"
        Exit Sub
    End If

    ' Strong confirmation prompt for final operation
    response = MsgBox("FINALIZE MONTH for site: " & selectedSite & "?" & vbCrLf & vbCrLf & _
                      "This will:" & vbCrLf & _
                      "  1. Save all current data to inflight tables" & vbCrLf & _
                      "  2. PERMANENTLY ARCHIVE records marked for archive + include" & vbCrLf & _
                      "  3. Update query worksheets" & vbCrLf & vbCrLf & _
                      "Use this when all month-end decisions are final." & vbCrLf & vbCrLf & _
                      "Proceed with month-end finalization?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Month-End Finalization")

    If response = vbNo Then
        Exit Sub
    End If

    startTime = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    success = False  ' Track overall success

    ' STEP 1: Validate Excel data FIRST (catch errors before touching database)
    Application.StatusBar = "Running validation checks..."
    If Not ValidateData(showSuccessMessage:=False) Then
        MsgBox "Validation failed - please check the Validation_Report sheet for errors." & vbCrLf & vbCrLf & _
               "Fix all validation errors and try again.", _
               vbExclamation, "Finalization Failed"
        GoTo Cleanup
    End If

    ' STEP 2: Unpivot cost data
    Application.StatusBar = "Preparing cost data..."
    If Not UnpivotCostData() Then
        MsgBox "Failed to unpivot cost data." & vbCrLf & vbCrLf & _
               "Please check your data and try again.", _
               vbExclamation, "Finalization Failed"
        GoTo Cleanup
    End If

    ' STEP 3: Upload to staging
    Application.StatusBar = "Uploading project data to staging..."
    If Not UploadProjectData() Then
        MsgBox "Failed to upload project data to staging." & vbCrLf & vbCrLf & _
               "Please check database connection and try again.", _
               vbExclamation, "Finalization Failed"
        GoTo Cleanup
    End If

    Application.StatusBar = "Uploading cost data to staging..."
    If Not UploadCostData() Then
        MsgBox "Failed to upload cost data to staging." & vbCrLf & vbCrLf & _
               "Please check database connection and try again.", _
               vbExclamation, "Finalization Failed"
        GoTo Cleanup
    End If

    ' STEP 4: Validate staging data (SQL-side validation)
    Application.StatusBar = "Running database validation checks..."
    If Not ValidateStagingData() Then
        MsgBox "Database validation failed." & vbCrLf & vbCrLf & _
               "Please check the Validation_Report sheet for errors.", _
               vbExclamation, "Finalization Failed"
        GoTo Cleanup
    End If

    ' STEP 5: Commit to inflight tables
    Application.StatusBar = "Committing to inflight database..."
    If Not CommitToInflight() Then
        MsgBox "Failed to commit data to inflight tables." & vbCrLf & vbCrLf & _
               "Database has been rolled back. Please check connection and try again.", _
               vbCritical, "Finalization Failed"
        GoTo Cleanup
    End If

    ' STEP 6: Archive approved records (permanent storage)
    Application.StatusBar = "Archiving approved records to permanent storage..."
    If Not ArchiveApprovedRecordsInternal() Then
        MsgBox "Failed to archive approved records." & vbCrLf & vbCrLf & _
               "Inflight data was saved, but archival failed." & vbCrLf & _
               "Please check the database and try archiving manually.", _
               vbExclamation, "Archival Failed"
        GoTo Cleanup
    End If

    ' STEP 7: Log submission
    Application.StatusBar = "Logging submission..."
    Call LogSubmission()  ' Non-critical, don't fail on logging errors

    ' STEP 8: Refresh query worksheets (silently - final message will report success)
    Application.StatusBar = "Refreshing query worksheets..."
    Call mod_WorksheetQuery.RefreshAll(showSuccessMessage:=False)

    success = True  ' Mark as successful

    ' Success!
    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "MONTH-END FINALIZATION COMPLETE!" & vbCrLf & vbCrLf & _
           "Site: " & selectedSite & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "Actions completed:" & vbCrLf & _
           "  - Saved snapshot to inflight tables" & vbCrLf & _
           "  - Archived approved records to permanent storage" & vbCrLf & _
           "  - Refreshed Archive and Inflight worksheets", _
           vbInformation, "Finalization Complete"

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Finalization failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "The database has been rolled back to its previous state." & vbCrLf & _
           "Please review the error and try again.", _
           vbCritical, "Finalization Error"
End Sub

' ----------------------------------------------------------------------------
' Function: ArchiveApprovedRecordsInternal
' Purpose: Internal function to archive records with archive=1 AND include=1
' Returns: True if successful, False if failed
' Notes: Called by FinalizeMonth - uses stored procedure for UPSERT capability
' ----------------------------------------------------------------------------
Private Function ArchiveApprovedRecordsInternal() As Boolean
    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection
    Dim selectedSite As String
    Dim success As Boolean

    ' Get selected site
    selectedSite = mod_SiteSetup.GetSelectedSite()
    If selectedSite = "" Then
        ArchiveApprovedRecordsInternal = False
        Exit Function
    End If

    ' Connect to database
    Set conn = mod_Database.GetDBConnection()
    If conn Is Nothing Then
        ArchiveApprovedRecordsInternal = False
        Exit Function
    End If

    ' Call stored procedure with site parameter
    ' usp_archive_approved_pifs handles UPSERT and DELETE operations
    success = mod_Database.ExecuteStoredProcedureNonQuery(conn, "dbo.usp_archive_approved_pifs", _
                                                          "@site", adVarChar, adParamInput, 4, selectedSite)

    ' Cleanup
    conn.Close
    Set conn = Nothing

    ArchiveApprovedRecordsInternal = success
    Exit Function

ErrHandler:
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If

    MsgBox "Archive operation failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Archive Error"

    ArchiveApprovedRecordsInternal = False
End Function

' ----------------------------------------------------------------------------
' Sub: RunValidationOnly
' Purpose: Run validation checks without submitting
' Usage: For testing/debugging before actual submission
' ----------------------------------------------------------------------------
Public Sub ValidateOnly()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' Unpivot cost data first
    Call UnpivotCostData
    
    ' Run validation
    If ValidateData() Then
        MsgBox "Validation passed! Data is ready for submission.", vbInformation
    Else
        MsgBox "Validation failed. Please fix errors before submitting.", vbExclamation
    End If
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Validation failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical
End Sub

' ============================================================================
' PRIVATE HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: UnpivotCostData
' Purpose: Transform wide cost columns to long (normalized) format using ARRAYS
' Details: Converts columns U-BF (requested/current/variance for each year/scenario)
'          into rows in the Cost_Unpivoted sheet
' PERFORMANCE: Array-based - 100x faster than cell-by-cell operations
' ----------------------------------------------------------------------------
Private Function UnpivotCostData() As Boolean
    On Error GoTo UnpivotCostData_Err

    Dim wsData As Worksheet
    Dim wsCost As Worksheet
    Dim lastRow As Long
    Dim currentYear As Integer
    Dim sourceData As Variant
    Dim outputArray() As Variant
    Dim outputRow As Long
    Dim dataRow As Long
    Dim i As Long
    Dim pifId As String, projectId As String
    Dim lineItem As Variant

    currentYear = ThisWorkbook.Names("CurrentYear").RefersToRange.Value
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)

    ' Find last row with data (PIF_ID now in column H)
    lastRow = wsData.Cells(wsData.Rows.Count, 8).End(xlUp).Row
    If lastRow < 4 Then
        UnpivotCostData = True
        Exit Function
    End If

    ' Read entire source range into array (ONE READ OPERATION) - extended to column BO
    sourceData = wsData.Range(wsData.Cells(4, 1), wsData.Cells(lastRow, 67)).Value

    ' Calculate output size: each row generates 12 cost rows (6 Target + 6 Closings years)
    Dim maxRows As Long
    maxRows = (lastRow - 3) * 12
    ReDim outputArray(1 To maxRows, 1 To 8)  ' Now 8 columns (added line_item)

    outputRow = 1

    ' Process array in memory (FAST!)
    For dataRow = 1 To UBound(sourceData, 1)
        pifId = Trim(sourceData(dataRow, 8) & "")       ' Column H (was G)
        projectId = Trim(sourceData(dataRow, 14) & "")  ' Column N (was M)
        lineItem = sourceData(dataRow, 7)               ' Column G (NEW - Line Item)

        ' Default line_item to 1 if blank
        If IsEmpty(lineItem) Or lineItem = "" Then
            lineItem = 1
        Else
            lineItem = CLng(lineItem)
        End If

        If pifId <> "" And projectId <> "" Then
            ' TARGET SCENARIO - 6 years (CY through CY+5)
            For i = 0 To 5
                outputArray(outputRow, 1) = pifId
                outputArray(outputRow, 2) = projectId
                outputArray(outputRow, 3) = lineItem
                outputArray(outputRow, 4) = SCENARIO_TARGET
                outputArray(outputRow, 5) = DateSerial(currentYear + i, 12, 31)
                outputArray(outputRow, 6) = ConvertToNumeric(sourceData(dataRow, 22 + i))      ' V-AA (Requested) - shifted +1
                outputArray(outputRow, 7) = ConvertToNumeric(sourceData(dataRow, 28 + i))      ' AB-AG (Current) - shifted +1
                outputArray(outputRow, 8) = ConvertToNumeric(sourceData(dataRow, 34 + i))      ' AH-AM (Variance) - shifted +1
                outputRow = outputRow + 1
            Next i

            ' CLOSINGS SCENARIO - 6 years (CY through CY+5)
            For i = 0 To 5
                outputArray(outputRow, 1) = pifId
                outputArray(outputRow, 2) = projectId
                outputArray(outputRow, 3) = lineItem
                outputArray(outputRow, 4) = SCENARIO_CLOSINGS
                outputArray(outputRow, 5) = DateSerial(currentYear + i, 12, 31)
                outputArray(outputRow, 6) = ConvertToNumeric(sourceData(dataRow, 42 + i))      ' AP-AU (Requested) - shifted +1
                outputArray(outputRow, 7) = ConvertToNumeric(sourceData(dataRow, 48 + i))      ' AV-BA (Current) - shifted +1
                outputArray(outputRow, 8) = ConvertToNumeric(sourceData(dataRow, 54 + i))      ' BB-BG (Variance) - shifted +1
                outputRow = outputRow + 1
            Next i
        End If
    Next dataRow

    ' Create or clear Cost_Unpivoted sheet
    On Error Resume Next
    Set wsCost = ThisWorkbook.Sheets(SHEET_COST_UNPIVOTED)
    On Error GoTo UnpivotCostData_Err

    If wsCost Is Nothing Then
        Set wsCost = ThisWorkbook.Sheets.Add(After:=wsData)
        wsCost.Name = SHEET_COST_UNPIVOTED
    Else
        wsCost.Cells.Clear
    End If

    ' Write headers (now 8 columns)
    wsCost.Range("A1:H1").Value = Array("pif_id", "project_id", "line_item", "scenario", "year", "requested_value", "current_value", "variance_value")

    ' Write entire array to sheet (ONE WRITE OPERATION)
    If outputRow > 1 Then
        wsCost.Range("A2").Resize(outputRow - 1, 8).Value = outputArray
    End If

    ' Format and hide
    wsCost.Columns("A:H").AutoFit
    wsCost.Visible = xlSheetHidden

    UnpivotCostData = True
    Exit Function

UnpivotCostData_Err:
    MsgBox "Failed to unpivot cost data:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, vbCritical
    UnpivotCostData = False
End Function

' ----------------------------------------------------------------------------
' Function: ConvertToNumeric
' Purpose: Convert variant to numeric value, handling empty/null safely
' Returns: Numeric value or 0
' ----------------------------------------------------------------------------
Private Function ConvertToNumeric(ByVal value As Variant) As Double
    If IsEmpty(value) Or IsNull(value) Then
        ConvertToNumeric = 0
    ElseIf IsNumeric(value) Then
        ConvertToNumeric = CDbl(value)
    Else
        ConvertToNumeric = 0
    End If
End Function


' ----------------------------------------------------------------------------
' Function: CreateBackupTables
' Purpose: Create timestamped backups of inflight tables
' ----------------------------------------------------------------------------
Private Function CreateBackupTables() As Boolean
    On Error GoTo ErrHandler
    
    Dim backupDate As String
    Dim sql As String
    
    backupDate = Format(Now, "YYYYMMDD_HHMMSS")
    
    ' Backup projects table
    sql = "SELECT * INTO dbo.tbl_pif_projects_inflight_backup_" & backupDate & _
          " FROM dbo.tbl_pif_projects_inflight"
    
    If Not ExecuteSQL(sql) Then
        CreateBackupTables = False
        Exit Function
    End If
    
    ' Backup cost table
    sql = "SELECT * INTO dbo.tbl_pif_cost_inflight_backup_" & backupDate & _
          " FROM dbo.tbl_pif_cost_inflight"
    
    If Not ExecuteSQL(sql) Then
        CreateBackupTables = False
        Exit Function
    End If
    
    CreateBackupTables = True
    Exit Function
    
ErrHandler:
    MsgBox "Failed to create backup tables:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical
    CreateBackupTables = False
End Function

' ----------------------------------------------------------------------------
' Function: UploadProjectData
' Purpose: Upload project metadata to staging table
' ----------------------------------------------------------------------------
Public Function BulkInsertToStaging(ByVal dataRange As Range, _
                                    ByVal tableName As String, _
                                    Optional ByVal schemaName As String = "dbo", _
                                    Optional ByVal selectedSite As String = "") As Boolean
    On Error GoTo LogError

    Dim conn As ADODB.Connection
    Dim i As Long
    Dim rowCount As Long
    Dim startTime As Double
    Dim params() As Variant
    Dim wsData As Worksheet
    Dim actualRow As Long
    Dim errorDetailLog As String  ' Variable to capture detailed error info

    startTime = Timer

    Debug.Print "=== BulkInsertToStaging STARTED for " & tableName & " ==="
    Debug.Print "Data range rows: " & dataRange.Rows.Count

    ' Get the worksheet reference for absolute column access
    Set wsData = dataRange.Worksheet

    Set conn = GetDBConnection()
    If conn Is Nothing Then
        Debug.Print "ERROR: GetDBConnection returned Nothing!"
        errorDetailLog = "Connection failed: GetDBConnection returned Nothing"
        BulkInsertToStaging = False
        GoTo LogError
    End If
    Debug.Print "Database connection established"

    ' Truncate staging table first
    Application.StatusBar = "Truncating " & tableName & "..."
    Debug.Print "Truncating " & tableName & "..."
    If Not ExecuteSQLSecure(conn, "TRUNCATE TABLE " & schemaName & "." & tableName) Then
        Debug.Print "ERROR: Failed to truncate table"
        errorDetailLog = "Failed to truncate table: " & tableName
        BulkInsertToStaging = False
        GoTo LogError
    End If
    Debug.Print "Table truncated successfully"

    ' Loop through Excel range and add records
    Application.StatusBar = "Uploading to " & tableName & "..."
    Application.ScreenUpdating = False
    rowCount = 0

    conn.BeginTrans
    Debug.Print "Transaction started"

    For i = 1 To dataRange.Rows.Count
        ' Calculate actual worksheet row
        actualRow = dataRange.Row + i - 1

        ' Check if row has data (skip empty rows) - use PIF_ID column (H=8)
        If Not IsEmpty(wsData.Cells(actualRow, 8).Value) Then
            If tableName = "tbl_pif_projects_staging" Then
                ' Get row-specific site and PIF ID for validation
                Dim rowSite As String
                Dim rowPifId As String
                
                rowSite = Trim(wsData.Cells(actualRow, 11).Value)    ' Column K = Site
                rowPifId = Trim(wsData.Cells(actualRow, 8).Value)    ' Column H = PIF_ID

                ' Validate site consistency if a site is provided
                If selectedSite <> "" Then
                    If Not ValidateSiteConsistency(selectedSite, rowSite, rowPifId) Then
                        Debug.Print "  ERROR: Site validation failed for row " & actualRow
                        
                        errorDetailLog = "Site Validation Failed:" & vbCrLf & _
                                         "Selected Site: " & selectedSite & vbCrLf & _
                                         "Row Site: " & rowSite & vbCrLf & _
                                         "PIF ID: " & rowPifId
                        
                        conn.RollbackTrans
                        BulkInsertToStaging = False
                        GoTo LogError
                    End If
                End If

                ReDim params(0 To 20) ' 21 parameters for usp_insert_project_staging (added line_item)
                ' Use absolute column references with proper type conversion
                params(0) = SafeString(wsData.Cells(actualRow, 8).Value)   ' pif_id (H) - VARCHAR
                params(1) = SafeString(wsData.Cells(actualRow, 14).Value)  ' project_id (N) - VARCHAR
                params(2) = SafeInteger(wsData.Cells(actualRow, 7).Value)  ' line_item (G) - INT (NEW)
                params(3) = SafeString(wsData.Cells(actualRow, 19).Value)  ' status (S) - VARCHAR
                params(4) = SafeString(wsData.Cells(actualRow, 6).Value)   ' change_type (F) - VARCHAR
                params(5) = SafeString(wsData.Cells(actualRow, 5).Value)   ' accounting_treatment (E) - VARCHAR
                params(6) = SafeString(wsData.Cells(actualRow, 20).Value)  ' category (T) - VARCHAR
                params(7) = SafeInteger(wsData.Cells(actualRow, 9).Value)  ' seg (I) - INT
                params(8) = SafeString(wsData.Cells(actualRow, 10).Value)  ' opco (J) - VARCHAR
                params(9) = SafeString(wsData.Cells(actualRow, 11).Value)  ' site (K) - VARCHAR
                params(10) = SafeString(wsData.Cells(actualRow, 12).Value) ' strategic_rank (L) - VARCHAR
                params(11) = SafeString(wsData.Cells(actualRow, 14).Value) ' funding_project (N) - VARCHAR
                params(12) = SafeString(wsData.Cells(actualRow, 15).Value) ' project_name (O) - VARCHAR
                params(13) = FormatDateISO(wsData.Cells(actualRow, 16).Value) ' original_fp_isd (P) - VARCHAR
                params(14) = FormatDateISO(wsData.Cells(actualRow, 17).Value) ' revised_fp_isd (Q) - VARCHAR
                params(15) = SafeString(wsData.Cells(actualRow, 39).Value) ' moving_isd_year (AN) - CHAR
                params(16) = SafeString(wsData.Cells(actualRow, 18).Value) ' lcm_issue (R) - VARCHAR
                params(17) = SafeString(wsData.Cells(actualRow, 21).Value) ' justification (U) - VARCHAR
                params(18) = SafeDecimal(wsData.Cells(actualRow, 41).Value) ' prior_year_spend (AO) - DECIMAL
                params(19) = SafeBoolean(wsData.Cells(actualRow, 3).Value)  ' archive_flag (C) - BIT
                params(20) = SafeBoolean(wsData.Cells(actualRow, 4).Value)  ' include_flag (D) - BIT

                Debug.Print "  Attempting to insert row " & actualRow & ": PIF=" & params(0) & ", Project=" & params(1) & ", Line Item=" & params(2)

                If Not ExecuteStoredProcedureNonQuery(conn, "usp_insert_project_staging", _
                    "@pif_id", adVarChar, adParamInput, 16, params(0), _
                    "@project_id", adVarChar, adParamInput, 10, params(1), _
                    "@line_item", adInteger, adParamInput, 0, params(2), _
                    "@status", adVarChar, adParamInput, 58, params(3), _
                    "@change_type", adVarChar, adParamInput, 12, params(4), _
                    "@accounting_treatment", adVarChar, adParamInput, 14, params(5), _
                    "@category", adVarChar, adParamInput, 26, params(6), _
                    "@seg", adInteger, adParamInput, 0, params(7), _
                    "@opco", adVarChar, adParamInput, 4, params(8), _
                    "@site", adVarChar, adParamInput, 4, params(9), _
                    "@strategic_rank", adVarChar, adParamInput, 26, params(10), _
                    "@funding_project", adVarChar, adParamInput, 10, params(11), _
                    "@project_name", adVarChar, adParamInput, 35, params(12), _
                    "@original_fp_isd", adVarChar, adParamInput, 20, params(13), _
                    "@revised_fp_isd", adVarChar, adParamInput, 20, params(14), _
                    "@moving_isd_year", adChar, adParamInput, 1, params(15), _
                    "@lcm_issue", adVarChar, adParamInput, 20, params(16), _
                    "@justification", adVarChar, adParamInput, 192, params(17), _
                    "@prior_year_spend", adNumeric, adParamInput, 0, params(18), _
                    "@archive_flag", adTinyInt, adParamInput, 0, params(19), _
                    "@include_flag", adTinyInt, adParamInput, 0, params(20)) Then
                    
                    Debug.Print "  ERROR: Failed to insert row " & actualRow
                    
                    ' Capture detailed error information
                    errorDetailLog = "Failed to insert row " & actualRow & vbCrLf & _
                                     "PIF ID: " & params(0) & vbCrLf & _
                                     "Project ID: " & params(1) & vbCrLf & _
                                     "Line Item: " & params(2)
                    
                    conn.RollbackTrans
                    BulkInsertToStaging = False
                    GoTo LogError
                End If

                rowCount = rowCount + 1
            End If
        Else
            Debug.Print "Skipping row " & actualRow & " (PIF_ID is empty)"
        End If
    Next i

    Debug.Print "Loop completed. Total rows processed: " & rowCount
    Debug.Print "Committing transaction..."
    conn.CommitTrans
    Debug.Print "Transaction committed"

    conn.Close
    Set conn = Nothing

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Dim elapsed As Double
    elapsed = Timer - startTime

    Debug.Print "Successfully uploaded " & rowCount & " rows to " & tableName & " in " & Format(elapsed, "0.0") & " seconds"
    Debug.Print "=== BulkInsertToStaging COMPLETED SUCCESSFULLY ==="

    BulkInsertToStaging = True
    Exit Function

LogError:
    ' Enhanced error logging
    Dim finalErrorMsg As String
    finalErrorLog = "Bulk insert failed:" & vbCrLf & _
                    "Table: " & tableName & vbCrLf & _
                    "Error Details: " & errorDetailLog

    MsgBox finalErrorMsg, vbCritical, "Upload Error"
    
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            On Error Resume Next
            conn.RollbackTrans
            conn.Close
            On Error GoTo 0
        End If
        Set conn = Nothing
    End If

    BulkInsertToStaging = False
End Function

' Add this helper function if not already present
Private Function ValidateSiteConsistency(ByVal selectedSite As String, ByVal rowSite As String, ByVal pifId As String) As Boolean
    ' Convert both to uppercase for case-insensitive comparison
    selectedSite = UCase(Trim(selectedSite))
    rowSite = UCase(Trim(rowSite))
    
    ' First, check if the row's site matches the selected site
    If rowSite <> selectedSite Then
        Debug.Print "SITE MISMATCH WARNING: " & _
                    "Selected Site: " & selectedSite & ", " & _
                    "Row Site: " & rowSite & ", " & _
                    "PIF ID: " & pifId
        ValidateSiteConsistency = False
        Exit Function
    End If
    
    ' Additional PIF ID validation if needed
    If InStr(1, pifId, selectedSite) = 0 Then
        Debug.Print "PIF ID SITE MISMATCH WARNING: " & _
                    "PIF ID: " & pifId & " does not contain site: " & selectedSite
    End If
    
    ValidateSiteConsistency = True
End Function

' ----------------------------------------------------------------------------
' Function: UploadCostData
' Purpose: Upload unpivoted cost data to staging table
' ----------------------------------------------------------------------------
Private Function UploadCostData() As Boolean
    On Error GoTo ErrHandler
    
    Dim wsCost As Worksheet
    Dim dataRange As Range
    Dim lastDataRow As Long
    
    Set wsCost = ThisWorkbook.Sheets(SHEET_COST_UNPIVOTED)
    
    ' Find the last row with data in Column A (pif_id) to define the extent of cost data
    lastDataRow = wsCost.Cells(wsCost.Rows.Count, "A").End(xlUp).Row
    
    ' Ensure we don't include header rows or start before the actual data
    If lastDataRow < 2 Then lastDataRow = 1 ' If no data, set to just above data start
    
    ' Define the data range from row 2 (first data row) to the last data row, across relevant columns
    ' Assuming cost data spans from column A to H (8 columns with line_item)
    Set dataRange = wsCost.Range(wsCost.Cells(2, "A"), wsCost.Cells(lastDataRow, "H"))
    
    UploadCostData = BulkInsertCosts(dataRange)
    
    Exit Function
    
ErrHandler:
    MsgBox "Failed to upload cost data:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical
    UploadCostData = False
End Function

' ----------------------------------------------------------------------------
' Function: CommitToInflight
' Purpose: Move data from staging to inflight tables (within transaction)
' ----------------------------------------------------------------------------
Private Function CommitToInflight() As Boolean
    On Error GoTo ErrHandler

    Dim selectedSite As String
    Dim sql As String

    ' Get selected site from Instructions sheet
    On Error Resume Next
    selectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.Value)
    On Error GoTo ErrHandler

    If selectedSite = "" Then
        MsgBox "Site not selected. Cannot commit to inflight tables." & vbCrLf & vbCrLf & _
               "Please select a site on the Instructions worksheet.", _
               vbExclamation, "Site Required"
        CommitToInflight = False
        Exit Function
    End If

    ' Call stored procedure with site parameter (site-filtered commit)
    sql = "EXEC dbo.usp_commit_to_inflight @site = '" & SQLSafe(selectedSite) & "'"

    CommitToInflight = ExecuteSQL(sql)
    
    Exit Function
    
ErrHandler:
    MsgBox "Failed to commit to inflight tables:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Transaction has been rolled back.", _
           vbCritical
    CommitToInflight = False
End Function

' ----------------------------------------------------------------------------
' Function: ArchiveApprovedPIFs
' Purpose: Move approved PIFs from inflight to approved tables
' ----------------------------------------------------------------------------
Private Function ArchiveApprovedPIFs() As Boolean
    On Error GoTo ErrHandler

    Dim sql As String
    Dim archivedCount As Long

    ' UPSERT approved projects using MERGE (update if exists, insert if new)
    ' Build SQL in chunks to avoid VBA's 25 line continuation limit
    sql = "MERGE dbo.tbl_pif_projects_approved AS target "
    sql = sql & "USING (SELECT pif_id, project_id, submission_date, status, change_type, "
    sql = sql & "accounting_treatment, category, seg, opco, site, strategic_rank, "
    sql = sql & "funding_project, project_name, original_fp_isd, revised_fp_isd, "
    sql = sql & "moving_isd_year, lcm_issue, justification, prior_year_spend, "
    sql = sql & "archive_flag, include_flag FROM dbo.tbl_pif_projects_inflight "
    sql = sql & "WHERE archive_flag = 1 AND include_flag = 1) AS source "
    sql = sql & "ON target.pif_id = source.pif_id AND target.project_id = source.project_id "
    sql = sql & "WHEN MATCHED THEN UPDATE SET submission_date = source.submission_date, "
    sql = sql & "approval_date = GETDATE(), status = source.status, change_type = source.change_type, "
    sql = sql & "accounting_treatment = source.accounting_treatment, category = source.category, "
    sql = sql & "seg = source.seg, opco = source.opco, site = source.site, "
    sql = sql & "strategic_rank = source.strategic_rank, funding_project = source.funding_project, "
    sql = sql & "project_name = source.project_name, original_fp_isd = source.original_fp_isd, "
    sql = sql & "revised_fp_isd = source.revised_fp_isd, moving_isd_year = source.moving_isd_year, "
    sql = sql & "lcm_issue = source.lcm_issue, justification = source.justification, "
    sql = sql & "prior_year_spend = source.prior_year_spend, archive_flag = source.archive_flag, "
    sql = sql & "include_flag = source.include_flag "
    sql = sql & "WHEN NOT MATCHED THEN INSERT (pif_id, project_id, submission_date, approval_date, "
    sql = sql & "status, change_type, accounting_treatment, category, seg, opco, site, strategic_rank, "
    sql = sql & "funding_project, project_name, original_fp_isd, revised_fp_isd, moving_isd_year, "
    sql = sql & "lcm_issue, justification, prior_year_spend, archive_flag, include_flag) "
    sql = sql & "VALUES (source.pif_id, source.project_id, source.submission_date, GETDATE(), "
    sql = sql & "source.status, source.change_type, source.accounting_treatment, source.category, "
    sql = sql & "source.seg, source.opco, source.site, source.strategic_rank, source.funding_project, "
    sql = sql & "source.project_name, source.original_fp_isd, source.revised_fp_isd, "
    sql = sql & "source.moving_isd_year, source.lcm_issue, source.justification, source.prior_year_spend, "
    sql = sql & "source.archive_flag, source.include_flag);"

    If Not ExecuteSQL(sql) Then
        ArchiveApprovedPIFs = False
        Exit Function
    End If

    ' Delete and re-insert costs (simpler than MERGE for child records)
    sql = "DELETE c FROM dbo.tbl_pif_cost_approved c " & _
          "WHERE EXISTS (SELECT 1 FROM dbo.tbl_pif_projects_inflight p " & _
          "WHERE p.pif_id = c.pif_id AND p.project_id = c.project_id " & _
          "AND p.archive_flag = 1 AND p.include_flag = 1); " & _
          "INSERT INTO dbo.tbl_pif_cost_approved " & _
          "(pif_id, project_id, scenario, year, requested_value, " & _
          "current_value, variance_value, approval_date) " & _
          "SELECT c.pif_id, c.project_id, c.scenario, c.year, " & _
          "c.requested_value, c.current_value, c.variance_value, GETDATE() " & _
          "FROM dbo.tbl_pif_cost_inflight c " & _
          "INNER JOIN dbo.tbl_pif_projects_inflight p " & _
          "    ON c.pif_id = p.pif_id AND c.project_id = p.project_id " & _
          "WHERE p.archive_flag = 1 AND p.include_flag = 1"

    If Not ExecuteSQL(sql) Then
        ArchiveApprovedPIFs = False
        Exit Function
    End If

    ' Remove approved records from inflight
    sql = "DELETE c FROM dbo.tbl_pif_cost_inflight c " & _
          "INNER JOIN dbo.tbl_pif_projects_inflight p " & _
          "    ON c.pif_id = p.pif_id AND c.project_id = p.project_id " & _
          "WHERE p.archive_flag = 1 AND p.include_flag = 1; " & _
          "DELETE FROM dbo.tbl_pif_projects_inflight " & _
          "WHERE archive_flag = 1 AND include_flag = 1"

    If Not ExecuteSQL(sql) Then
        ArchiveApprovedPIFs = False
        Exit Function
    End If

    ArchiveApprovedPIFs = True
    Exit Function
    
ErrHandler:
    MsgBox "Failed to archive approved PIFs:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical
    ArchiveApprovedPIFs = False
End Function

' ----------------------------------------------------------------------------
' Function: LogSubmission
' Purpose: Record submission in audit log
' ----------------------------------------------------------------------------
Private Function LogSubmission() As Boolean
    On Error GoTo ErrHandler
    
    Dim sql As String
    Dim recordCount As Long
    
    recordCount = GetRecordCount("tbl_pif_projects_inflight")
    
    sql = "INSERT INTO dbo.tbl_submission_log " & _
          "(submission_date, submitted_by, source_file, record_count, notes) VALUES " & _
          "(GETDATE(), SYSTEM_USER, '" & SQLSafe(ThisWorkbook.Name) & "', " & _
          recordCount & ", 'Submitted via VBA')"
    
    LogSubmission = ExecuteSQL(sql)
    
    Exit Function
    
ErrHandler:
    ' Log failure is not critical - don't fail entire submission
    LogSubmission = True
End Function

' ----------------------------------------------------------------------------
' Sub: ClearArchivedCheckboxes
' Purpose: Clear archive/include checkboxes for records that were archived
' Note: Prevents accidental re-archival on next submission
' ----------------------------------------------------------------------------
Private Sub ClearArchivedCheckboxes()
    On Error Resume Next ' Don't fail submission if this fails

    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim archiveCol As Integer
    Dim includeCol As Integer
    Dim clearedCount As Long

    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)

    ' Get column indices from shared constants
    archiveCol = PIFDataColumns.colArchive  ' Column C
    includeCol = PIFDataColumns.colInclude  ' Column D

    ' Find last row with data
    lastRow = wsData.Cells(wsData.Rows.Count, PIFDataColumns.colPIFID).End(xlUp).Row

    clearedCount = 0

    ' Loop through data rows (start at row 4, rows 1-3 are headers)
    For i = 4 To lastRow
        ' Skip empty rows
        If Not IsEmpty(wsData.Cells(i, PIFDataColumns.colPIFID).Value) Then
            ' If both checkboxes are checked, clear them
            If wsData.Cells(i, archiveCol).Value = True And _
               wsData.Cells(i, includeCol).Value = True Then
                wsData.Cells(i, archiveCol).Value = False
                wsData.Cells(i, includeCol).Value = False
                clearedCount = clearedCount + 1
            End If
        End If
    Next i

    On Error GoTo 0
End Sub

' ============================================================================
' ARCHIVE BUTTON WORKFLOW (SEPARATE FROM SUBMIT)
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ArchiveApprovedRecords
' Purpose: Archive PIFs with archive=1 AND include=1 to approved tables
' Usage: Attach this to the [Archive] button
' ----------------------------------------------------------------------------
Public Sub ArchiveApproved()
    On Error GoTo ErrHandler

    Dim selectedSite As String
    Dim response As VbMsgBoxResult
    Dim startTime As Double
    Dim success As Boolean

    startTime = Timer
    Application.ScreenUpdating = False

    ' Validate site selection
    On Error Resume Next
    selectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.Value)
    On Error GoTo ErrHandler

    If selectedSite = "" Then
        MsgBox "Please select a site before archiving." & vbCrLf & vbCrLf & _
               "Go to the Instructions worksheet and select your site from the dropdown.", _
               vbExclamation, "Site Not Selected"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Block Fleet archival
    If UCase(selectedSite) = "FLEET" Then
        MsgBox "Fleet cannot archive data." & vbCrLf & vbCrLf & _
               "Fleet is read-only access for viewing all sites.", _
               vbExclamation, "Fleet Archive Not Allowed"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Confirmation prompt
    response = MsgBox("Archive PIFs with ARCHIVE and INCLUDE checked?" & vbCrLf & vbCrLf & _
                      "Site: " & selectedSite & vbCrLf & _
                      "This will move them from inflight to permanent approved tables.", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Archive")

    If response = vbNo Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Archive approved PIFs
    Application.StatusBar = "Archiving to permanent tables..."
    success = ArchiveApprovedPIFs()

    If Not success Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If

    ' Success
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Archive complete! Records moved to permanent approved tables.", vbInformation, "Success"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Archive failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Archive Error"
End Sub

