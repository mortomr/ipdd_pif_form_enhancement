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
' Sub: SubmitToDatabase
' Purpose: Main entry point for PIF submission process
' Usage: Attach this to a button or run from VBA editor
' ----------------------------------------------------------------------------
Public Sub SubmitToDatabase()
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
        MsgBox "Please select a site before submitting." & vbCrLf & vbCrLf & _
               "Go to the Instructions worksheet and select your site from the dropdown.", _
               vbExclamation, "Site Not Selected"
        Exit Sub
    End If

    ' Block Fleet submissions (read-only access)
    If UCase(selectedSite) = "FLEET" Then
        MsgBox "Fleet cannot submit data." & vbCrLf & vbCrLf & _
               "Fleet is read-only access for viewing all sites." & vbCrLf & _
               "Please select a specific site (ANO, GGN, RBN, WF3, or HQN) to submit data.", _
               vbExclamation, "Fleet Submission Not Allowed"
        Exit Sub
    End If

    ' Confirmation prompt
    response = MsgBox("Submit PIF data for site: " & selectedSite & "?" & vbCrLf & vbCrLf & _
                      "This will update the inflight database tables.", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Submission")

    If response = vbNo Then
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

    ' STEP 2: Create backups (DISABLED - creates clutter in database)
    ' Application.StatusBar = "Creating backup tables..."
    ' success = CreateBackupTables()
    ' If Not success Then GoTo Cleanup

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

    ' STEP 6: Log submission
    Application.StatusBar = "Logging submission..."
    success = LogSubmission()

    ' Success!
    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "Submission completed successfully!" & vbCrLf & vbCrLf & _
           "Site: " & selectedSite & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "Data has been submitted to inflight tables." & vbCrLf & _
           "Use the Archive button to move approved PIFs to permanent storage.", _
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
        pifId = wsData.Cells(dataRow, PIFDataColumns.colPIFID).Value              ' Column G = pif_id
        projectId = wsData.Cells(dataRow, PIFDataColumns.colFundingProject).Value ' Column M = funding_project

        If pifId <> "" And projectId <> "" Then
            ' TARGET SCENARIO - Years CY through CY+5
            ' Columns: U-Z (requested), AA-AF (approved/current), AG-AL (variance)
            Dim reqVal As Variant
            Dim curVal As Variant
            Dim varVal As Variant

            ' CY (2025)
            reqVal = wsData.Cells(dataRow, COL_TARGET_REQ_CY).Value
            curVal = wsData.Cells(dataRow, COL_TARGET_APPR_CY).Value
            varVal = wsData.Cells(dataRow, COL_TARGET_VAR_CY).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_TARGET, currentYear & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+1 (2026)
            reqVal = wsData.Cells(dataRow, COL_TARGET_REQ_CY1).Value
            curVal = wsData.Cells(dataRow, COL_TARGET_APPR_CY1).Value
            varVal = wsData.Cells(dataRow, COL_TARGET_VAR_CY1).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_TARGET, (currentYear + 1) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+2 (2027)
            reqVal = wsData.Cells(dataRow, COL_TARGET_REQ_CY2).Value
            curVal = wsData.Cells(dataRow, COL_TARGET_APPR_CY2).Value
            varVal = wsData.Cells(dataRow, COL_TARGET_VAR_CY2).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_TARGET, (currentYear + 2) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+3 (2028)
            reqVal = wsData.Cells(dataRow, COL_TARGET_REQ_CY3).Value
            curVal = wsData.Cells(dataRow, COL_TARGET_APPR_CY3).Value
            varVal = wsData.Cells(dataRow, COL_TARGET_VAR_CY3).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_TARGET, (currentYear + 3) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+4 (2029)
            reqVal = wsData.Cells(dataRow, COL_TARGET_REQ_CY4).Value
            curVal = wsData.Cells(dataRow, COL_TARGET_APPR_CY4).Value
            varVal = wsData.Cells(dataRow, COL_TARGET_VAR_CY4).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_TARGET, (currentYear + 4) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+5 (2030)
            reqVal = wsData.Cells(dataRow, COL_TARGET_REQ_CY5).Value
            curVal = wsData.Cells(dataRow, COL_TARGET_APPR_CY5).Value
            varVal = wsData.Cells(dataRow, COL_TARGET_VAR_CY5).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_TARGET, (currentYear + 5) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CLOSINGS SCENARIO - Years CY through CY+5
            ' Columns: AO-AT (requested), AU-AZ (approved/current), BA-BF (variance)

            ' CY (2025)
            reqVal = wsData.Cells(dataRow, COL_CLOSINGS_REQ_CY).Value
            curVal = wsData.Cells(dataRow, COL_CLOSINGS_APPR_CY).Value
            varVal = wsData.Cells(dataRow, COL_CLOSINGS_VAR_CY).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_CLOSINGS, currentYear & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+1 (2026)
            reqVal = wsData.Cells(dataRow, COL_CLOSINGS_REQ_CY1).Value
            curVal = wsData.Cells(dataRow, COL_CLOSINGS_APPR_CY1).Value
            varVal = wsData.Cells(dataRow, COL_CLOSINGS_VAR_CY1).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_CLOSINGS, (currentYear + 1) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+2 (2027)
            reqVal = wsData.Cells(dataRow, COL_CLOSINGS_REQ_CY2).Value
            curVal = wsData.Cells(dataRow, COL_CLOSINGS_APPR_CY2).Value
            varVal = wsData.Cells(dataRow, COL_CLOSINGS_VAR_CY2).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_CLOSINGS, (currentYear + 2) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+3 (2028)
            reqVal = wsData.Cells(dataRow, COL_CLOSINGS_REQ_CY3).Value
            curVal = wsData.Cells(dataRow, COL_CLOSINGS_APPR_CY3).Value
            varVal = wsData.Cells(dataRow, COL_CLOSINGS_VAR_CY3).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_CLOSINGS, (currentYear + 3) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+4 (2029)
            reqVal = wsData.Cells(dataRow, COL_CLOSINGS_REQ_CY4).Value
            curVal = wsData.Cells(dataRow, COL_CLOSINGS_APPR_CY4).Value
            varVal = wsData.Cells(dataRow, COL_CLOSINGS_VAR_CY4).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_CLOSINGS, (currentYear + 4) & "-12-31", reqVal, curVal, varVal
            outputRow = outputRow + 1

            ' CY+5 (2030)
            reqVal = wsData.Cells(dataRow, COL_CLOSINGS_REQ_CY5).Value
            curVal = wsData.Cells(dataRow, COL_CLOSINGS_APPR_CY5).Value
            varVal = wsData.Cells(dataRow, COL_CLOSINGS_VAR_CY5).Value
            AddCostRow wsCost, outputRow, pifId, projectId, SCENARIO_CLOSINGS, (currentYear + 5) & "-12-31", reqVal, curVal, varVal
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
' Note: Handles NULL/empty values and type conversions safely
' ----------------------------------------------------------------------------
Private Sub AddCostRow(ByVal ws As Worksheet, ByVal rowNum As Long, _
                       ByVal pifId As String, ByVal projectId As String, _
                       ByVal scenario As String, ByVal year As String, _
                       ByVal requested As Variant, ByVal current As Variant, _
                       ByVal variance As Variant)

    On Error Resume Next  ' Handle conversion errors gracefully

    ws.Cells(rowNum, 1).Value = pifId
    ws.Cells(rowNum, 2).Value = projectId
    ws.Cells(rowNum, 3).Value = scenario

    ' Convert year string to date (e.g., "2025-12-31")
    ws.Cells(rowNum, 4).Value = CDate(year)
    If Err.Number <> 0 Then
        ws.Cells(rowNum, 4).Value = DateSerial(CInt(Left(year, 4)), 12, 31)
        Err.Clear
    End If

    ' Handle numeric values - treat empty/non-numeric as 0
    If IsEmpty(requested) Or Not IsNumeric(requested) Or Trim(CStr(requested)) = "" Then
        ws.Cells(rowNum, 5).Value = 0
    Else
        ws.Cells(rowNum, 5).Value = CDbl(requested)
    End If

    If IsEmpty(current) Or Not IsNumeric(current) Or Trim(CStr(current)) = "" Then
        ws.Cells(rowNum, 6).Value = 0
    Else
        ws.Cells(rowNum, 6).Value = CDbl(current)
    End If

    If IsEmpty(variance) Or Not IsNumeric(variance) Or Trim(CStr(variance)) = "" Then
        ws.Cells(rowNum, 7).Value = 0
    Else
        ws.Cells(rowNum, 7).Value = CDbl(variance)
    End If

    On Error GoTo 0
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
    Dim lastDataRow As Long
    
    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    
    ' Find the last row with data in Column G (PIF_ID) to define the extent of project data
    lastDataRow = wsData.Cells(wsData.Rows.Count, "G").End(xlUp).Row
    
    ' Ensure we don't include header rows or start before the actual data
    If lastDataRow < 4 Then lastDataRow = 3 ' If no data, set to just above data start
    
    ' Define the data range from row 4 (first data row) to the last data row, across relevant columns
    ' Assuming project data spans from column C to AN (40)
    Set dataRange = wsData.Range(wsData.Cells(4, "C"), wsData.Cells(lastDataRow, "AN"))
    
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
    Dim lastDataRow As Long
    
    Set wsCost = ThisWorkbook.Sheets(SHEET_COST_UNPIVOTED)
    
    ' Find the last row with data in Column A (pif_id) to define the extent of cost data
    lastDataRow = wsCost.Cells(wsCost.Rows.Count, "A").End(xlUp).Row
    
    ' Ensure we don't include header rows or start before the actual data
    If lastDataRow < 2 Then lastDataRow = 1 ' If no data, set to just above data start
    
    ' Define the data range from row 2 (first data row) to the last data row, across relevant columns
    ' Assuming cost data spans from column A to G
    Set dataRange = wsCost.Range(wsCost.Cells(2, "A"), wsCost.Cells(lastDataRow, "G"))
    
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

    Debug.Print "Cleared archive/include checkboxes for " & clearedCount & " rows"
End Sub

' ============================================================================
' ARCHIVE BUTTON WORKFLOW (SEPARATE FROM SUBMIT)
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ArchiveApprovedRecords
' Purpose: Archive PIFs with archive=1 AND include=1 to approved tables
' Usage: Attach this to the [Archive] button
' ----------------------------------------------------------------------------
Public Sub ArchiveApprovedRecords()
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
    response = MsgBox("Archive PIFs with Archive☑ and Include☑ checked?" & vbCrLf & vbCrLf & _
                      "Site: " & selectedSite & vbCrLf & _
                      "This will move them to permanent approved tables.", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Archive")

    If response = vbNo Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Archive approved PIFs
    Application.StatusBar = "Archiving approved PIFs..."
    success = ArchiveApprovedPIFs()

    If Not success Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If

    ' Clear checkboxes
    Application.StatusBar = "Clearing checkboxes..."
    Call ClearArchivedCheckboxes()

    ' Success
    Dim elapsed As Double
    elapsed = Timer - startTime

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Archive completed successfully!" & vbCrLf & vbCrLf & _
           "Site: " & selectedSite & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "Approved PIFs have been moved to permanent tables." & vbCrLf & _
           "Checkboxes have been cleared for archived records.", _
           vbInformation, "Archive Complete"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Archive failed:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Archive Error"
End Sub
