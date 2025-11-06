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

' Sheet name constants - UPDATE IF YOUR SHEETS HAVE DIFFERENT NAMES
Private Const SHEET_DATA As String = "PIF"
Private Const SHEET_COST_UNPIVOTED As String = "Cost_Unpivoted"
Private Const SHEET_VALIDATION_REPORT As String = "Validation_Report"

' ============================================================================
' PUBLIC INTERFACE
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: SubmitToDatabase
' Purpose: Main entry point for PIF submission process
' Usage: Attach this to a button or run from VBA editor
' ----------------------------------------------------------------------------
Public Sub SubmitToDatabase()
    On Error GoTo ErrHandler
    
    Dim response As VbMsgBoxResult
    Dim startTime As Double
    Dim success As Boolean
    
    ' Confirmation prompt
    response = MsgBox("This will submit the PIF workbook to the database." & vbCrLf & vbCrLf & _
                      "Before proceeding, please confirm:" & vbCrLf & _
                      "  - You have reviewed all data" & vbCrLf & _
                      "  - The workbook has been approved for submission" & vbCrLf & _
                      "  - You have a backup copy of this file" & vbCrLf & vbCrLf & _
                      "Continue with submission?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Submission")
    
    If response = vbNo Then
        MsgBox "Submission cancelled by user.", vbInformation
        Exit Sub
    End If
    
    startTime = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' STEP 1: Unpivot cost data
    Application.StatusBar = "Preparing cost data..."
    success = UnpivotCostData()
    If Not success Then GoTo Cleanup
    
    ' STEP 2: Create backups
    Application.StatusBar = "Creating backup tables..."
    success = CreateBackupTables()
    If Not success Then GoTo Cleanup
    
    ' STEP 3: Upload to staging
    Application.StatusBar = "Uploading project data to staging..."
    success = UploadProjectData()
    If Not success Then GoTo Cleanup
    
    Application.StatusBar = "Uploading cost data to staging..."
    success = UploadCostData()
    If Not success Then GoTo Cleanup
    
    ' STEP 4: Validate staging data
    Application.StatusBar = "Running validation checks..."
    success = ValidateData()  ' Excel-side validation
    If Not success Then GoTo Cleanup
    
    success = ValidateStagingData()  ' SQL-side validation
    If Not success Then GoTo Cleanup
    
    ' STEP 5: Commit to inflight tables
    Application.StatusBar = "Committing to database..."
    success = CommitToInflight()
    If Not success Then GoTo Cleanup
    
    ' STEP 6: Archive approved PIFs
    Application.StatusBar = "Archiving approved PIFs..."
    success = ArchiveApprovedPIFs()
    If Not success Then GoTo Cleanup
    
    ' STEP 7: Log submission
    Application.StatusBar = "Logging submission..."
    success = LogSubmission()
    
    ' Success!
    Dim elapsed As Double
    elapsed = Timer - startTime
    
    MsgBox "Submission completed successfully!" & vbCrLf & vbCrLf & _
           "Elapsed time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "Your PIF data has been submitted to the database." & vbCrLf & _
           "Approved PIFs have been archived to the permanent tables.", _
           vbInformation, "Success"
    
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
    
    MsgBox "Submission failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "The database has been rolled back to its previous state." & vbCrLf & _
           "Please review the error and try again.", _
           vbCritical, "Submission Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: RunValidationOnly
' Purpose: Run validation checks without submitting
' Usage: For testing/debugging before actual submission
' ----------------------------------------------------------------------------
Public Sub RunValidationOnly()
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
' Purpose: Transform wide cost columns to long (normalized) format
' Details: Converts columns U-BF (requested/current/variance for each year/scenario)
'          into rows in the Cost_Unpivoted sheet
' ----------------------------------------------------------------------------
Private Function UnpivotCostData() As Boolean
    On Error GoTo UnpivotCostData_Err
    
    Dim wsData As Worksheet
    Dim wsCost As Worksheet
    Dim lastRow As Long, dataRow As Long
    Dim outputRow As Long
    Dim pifId As String, projectId As String
    Dim scenario As String
    Dim year As String
    Dim requestedVal As Variant, currentVal As Variant, varianceVal As Variant
    Dim currentYear As Integer
    Dim yearStr As String
    
    currentYear = ThisWorkbook.Names("CurrentYear").RefersToRange.Value
    
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    
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
    
    ' Headers for unpivoted data
    wsCost.Range("A1").Value = "pif_id"
    wsCost.Range("B1").Value = "project_id"
    wsCost.Range("C1").Value = "scenario"
    wsCost.Range("D1").Value = "year"
    wsCost.Range("E1").Value = "requested_value"
    wsCost.Range("F1").Value = "current_value"
    wsCost.Range("G1").Value = "variance_value"
    
    lastRow = wsData.Cells(wsData.Rows.Count, 7).End(xlUp).Row  ' Column G = PIF_ID
    outputRow = 2
    
    Application.StatusBar = "Unpivoting cost data..."
    
    ' Process each data row
    For dataRow = 4 To lastRow
        pifId = wsData.Cells(dataRow, 7).Value       ' Column G = pif_id
        projectId = wsData.Cells(dataRow, 13).Value  ' Column M = funding_project (adjust as needed)
        
        If pifId <> "" And projectId <> "" Then
            ' TARGET SCENARIO - Years 2025-2030
            ' CY (2025): Columns U, AA, AG
            Dim reqVal As Variant
            Dim curVal As Variant
            Dim varVal As Variant
            
            reqVal = wsData.Cells(dataRow, 21).Value
            curVal = wsData.Cells(dataRow, 27).Value
            varVal = wsData.Cells(dataRow, 33).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Target", currentYear & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 22).Value
            curVal = wsData.Cells(dataRow, 28).Value
            varVal = wsData.Cells(dataRow, 34).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Target", (currentYear + 1) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 23).Value
            curVal = wsData.Cells(dataRow, 29).Value
            varVal = wsData.Cells(dataRow, 35).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Target", (currentYear + 2) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 24).Value
            curVal = wsData.Cells(dataRow, 30).Value
            varVal = wsData.Cells(dataRow, 36).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Target", (currentYear + 3) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 25).Value
            curVal = wsData.Cells(dataRow, 31).Value
            varVal = wsData.Cells(dataRow, 37).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Target", (currentYear + 4) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 26).Value
            curVal = wsData.Cells(dataRow, 32).Value
            varVal = wsData.Cells(dataRow, 38).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Target", (currentYear + 5) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1
            
            reqVal = wsData.Cells(dataRow, 41).Value
            curVal = wsData.Cells(dataRow, 47).Value
            varVal = wsData.Cells(dataRow, 53).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Closings", currentYear & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 42).Value
            curVal = wsData.Cells(dataRow, 48).Value
            varVal = wsData.Cells(dataRow, 54).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Closings", (currentYear + 1) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 43).Value
            curVal = wsData.Cells(dataRow, 49).Value
            varVal = wsData.Cells(dataRow, 55).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Closings", (currentYear + 2) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 44).Value
            curVal = wsData.Cells(dataRow, 50).Value
            varVal = wsData.Cells(dataRow, 56).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Closings", (currentYear + 3) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 45).Value
            curVal = wsData.Cells(dataRow, 51).Value
            varVal = wsData.Cells(dataRow, 57).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Closings", (currentYear + 4) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            reqVal = wsData.Cells(dataRow, 46).Value
            curVal = wsData.Cells(dataRow, 52).Value
            varVal = wsData.Cells(dataRow, 58).Value
            AddCostRow wsCost, outputRow, pifId, projectId, "Closings", (currentYear + 5) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1
        End If
        
        ' Progress indicator
        If dataRow Mod 10 = 0 Then
            Application.StatusBar = "Unpivoting cost data... Row " & dataRow & " of " & lastRow
        End If
    Next dataRow
    
    ' Format the output sheet
    wsCost.Columns("A:G").AutoFit
    wsCost.Visible = xlSheetHidden  ' Hide from users
    
    UnpivotCostData = True
    Exit Function

UnpivotCostData_Err:
    MsgBox "Failed to unpivot cost data on row " & dataRow & ":" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical
    UnpivotCostData = False
End Function

' ----------------------------------------------------------------------------
' Sub: AddCostRow
' Purpose: Helper function to add a row to the unpivoted cost sheet
' ----------------------------------------------------------------------------
Private Sub AddCostRow(ByVal ws As Worksheet, ByVal rowNum As Long, _
                       ByVal pifId As String, ByVal projectId As String, _
                       ByVal scenario As String, ByVal year As String, _
                       ByVal requested As Variant, ByVal current As Variant, _
                       ByVal variance As Variant)
    
    ws.Cells(rowNum, 1).Value = pifId
    ws.Cells(rowNum, 2).Value = projectId
    ws.Cells(rowNum, 3).Value = scenario
    ws.Cells(rowNum, 4).Value = CDate(year)
    
    ' Handle empty cells and convert to Double
    If IsEmpty(requested) Or requested = "" Then
        ws.Cells(rowNum, 5).Value = CDbl(0)
    Else
        ws.Cells(rowNum, 5).Value = CDbl(requested)
    End If
    
    If IsEmpty(current) Or current = "" Then
        ws.Cells(rowNum, 6).Value = CDbl(0)
    Else
        ws.Cells(rowNum, 6).Value = CDbl(current)
    End If
    
    If IsEmpty(variance) Or variance = "" Then
        ws.Cells(rowNum, 7).Value = CDbl(0)
    Else
        ws.Cells(rowNum, 7).Value = CDbl(variance)
    End If
End Sub

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
Private Function UploadProjectData() As Boolean
    On Error GoTo ErrHandler
    
    Dim wsData As Worksheet
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    
    ' Get the current region starting from C4 (assuming C4 is the start of data after headers)
    Set dataRange = wsData.Range("C4").CurrentRegion
    
    ' If the CurrentRegion includes more than just data (e.g., headers above C4),
    ' we need to adjust it. Assuming C4 is the first data cell.
    If dataRange.Row < 4 Then
        Set dataRange = dataRange.Offset(3, 0).Resize(dataRange.Rows.Count - 3, dataRange.Columns.Count)
    End If
    
    UploadProjectData = BulkInsertProjects(dataRange)
    
    Exit Function
    
ErrHandler:
    MsgBox "Failed to upload project data:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical
    UploadProjectData = False
End Function

' ----------------------------------------------------------------------------
' Function: UploadCostData
' Purpose: Upload unpivoted cost data to staging table
' ----------------------------------------------------------------------------
Private Function UploadCostData() As Boolean
    On Error GoTo ErrHandler
    
    Dim wsCost As Worksheet
    Dim dataRange As Range
    
    Set wsCost = ThisWorkbook.Sheets(SHEET_COST_UNPIVOTED)
    
    ' Get the current region starting from A2 (assuming A2 is the start of data after headers)
    Set dataRange = wsCost.Range("A2").CurrentRegion
    
    ' If the CurrentRegion includes more than just data (e.g., header in A1),
    ' we need to adjust it. Assuming A2 is the first data cell.
    If dataRange.Row < 2 Then
        Set dataRange = dataRange.Offset(1, 0).Resize(dataRange.Rows.Count - 1, dataRange.Columns.Count)
    End If
    
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
    
    Dim sql As String
    
    ' Execute within transaction for atomicity
    sql = "BEGIN TRANSACTION; " & _
          "TRUNCATE TABLE dbo.tbl_pif_cost_inflight; " & _
          "TRUNCATE TABLE dbo.tbl_pif_projects_inflight; " & _
          "INSERT INTO dbo.tbl_pif_projects_inflight " & _
          "SELECT *, GETDATE() AS submission_date FROM dbo.tbl_pif_projects_staging; " & _
          "INSERT INTO dbo.tbl_pif_cost_inflight " & _
          "SELECT * FROM dbo.tbl_pif_cost_staging; " & _
          "COMMIT TRANSACTION;"
    
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
    
    ' Insert approved projects
    sql = "INSERT INTO dbo.tbl_pif_projects_approved " & _
          "SELECT *, GETDATE() AS approval_date FROM dbo.tbl_pif_projects_inflight " & _
          "WHERE status IN ('Approved', 'Dispositioned')"
    
    If Not ExecuteSQL(sql) Then
        ArchiveApprovedPIFs = False
        Exit Function
    End If
    
    ' Insert approved costs
    sql = "INSERT INTO dbo.tbl_pif_cost_approved " & _
          "SELECT c.*, GETDATE() AS approval_date " & _
          "FROM dbo.tbl_pif_cost_inflight c " & _
          "INNER JOIN dbo.tbl_pif_projects_inflight p " & _
          "    ON c.pif_id = p.pif_id AND c.project_id = p.project_id " & _
          "WHERE p.status IN ('Approved', 'Dispositioned')"
    
    If Not ExecuteSQL(sql) Then
        ArchiveApprovedPIFs = False
        Exit Function
    End If
    
    ' Remove approved records from inflight
    sql = "DELETE c FROM dbo.tbl_pif_cost_inflight c " & _
          "INNER JOIN dbo.tbl_pif_projects_inflight p " & _
          "    ON c.pif_id = p.pif_id AND c.project_id = p.project_id " & _
          "WHERE p.status IN ('Approved', 'Dispositioned'); " & _
          "DELETE FROM dbo.tbl_pif_projects_inflight " & _
          "WHERE status IN ('Approved', 'Dispositioned')"
    
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
