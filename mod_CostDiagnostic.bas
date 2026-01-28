' ============================================================================
' MODULE: mod_CostDiagnostic
' ============================================================================
' Purpose: Diagnose cost capture failures at each step of the pipeline
' Usage: Run DiagnoseFullCostFlow() to trace costs from Excel to database
' ============================================================================

Option Explicit

Public Sub DiagnoseFullCostFlow()
    On Error GoTo ErrHandler
    
    Dim msg As String
    Dim step As Integer
    
    step = 1
    msg = "COST FLOW DIAGNOSTIC - STEP BY STEP" & vbCrLf & vbCrLf
    
    ' STEP 1: Check PIF worksheet has cost data
    msg = msg & "STEP 1: Checking PIF worksheet cost columns..." & vbCrLf
    step = 1
    Dim pifCostCount As Long
    pifCostCount = CountCostDataInPIF()
    msg = msg & "  Result: Found " & pifCostCount & " cost cell entries in PIF sheet" & vbCrLf & vbCrLf
    
    If pifCostCount = 0 Then
        MsgBox msg & "ERROR: No cost data found in PIF worksheet!" & vbCrLf & _
               "Add cost values to columns V-BG before proceeding.", vbExclamation
        Exit Sub
    End If
    
    ' STEP 2: Run UnpivotCostData and check Cost_Unpivoted sheet
    msg = msg & "STEP 2: Running UnpivotCostData..." & vbCrLf
    step = 2
    Call UnpivotCostData
    
    Dim unpivotCount As Long
    unpivotCount = CountRowsInSheet("Cost_Unpivoted", 2)  ' Count from row 2 (skip header)
    msg = msg & "  Result: Cost_Unpivoted sheet has " & unpivotCount & " rows" & vbCrLf
    
    If unpivotCount = 0 Then
        MsgBox msg & "ERROR: Cost_Unpivoted sheet is empty!" & vbCrLf & _
               "This means UnpivotCostData failed to generate cost rows.", vbExclamation
        Exit Sub
    End If
    
    ' STEP 3: Check Cost_Unpivoted data quality
    msg = msg & "STEP 3: Validating Cost_Unpivoted data..." & vbCrLf
    step = 3
    Dim validCount As Long, nullCount As Long, zeroCount As Long
    Call ValidateCostUnpivotedData(validCount, nullCount, zeroCount)
    msg = msg & "  Result:" & vbCrLf & _
              "    Valid cost records: " & validCount & vbCrLf & _
              "    NULL requested_value: " & nullCount & vbCrLf & _
              "    Zero requested_value: " & zeroCount & vbCrLf & vbCrLf
    
    If validCount = 0 Then
        MsgBox msg & "ERROR: All cost records have NULL or zero values!" & vbCrLf & _
               "Check that you entered cost values in the PIF worksheet.", vbExclamation
        Exit Sub
    End If
    
    ' STEP 4: Try uploading to staging
    msg = msg & "STEP 4: Uploading cost data to tbl_pif_cost_staging..." & vbCrLf
    step = 4
    
    Dim uploadSuccess As Boolean
    uploadSuccess = TestCostUploadToStaging()
    
    If Not uploadSuccess Then
        MsgBox msg & "ERROR: Failed to upload costs to staging table!" & vbCrLf & _
               "Check database connection and permissions.", vbCritical
        Exit Sub
    End If
    msg = msg & "  Result: Upload succeeded" & vbCrLf & vbCrLf
    
    ' STEP 5: Check staging table has data
    msg = msg & "STEP 5: Verifying tbl_pif_cost_staging..." & vbCrLf
    step = 5
    Dim stagingCount As Long
    stagingCount = GetStagingCostCount()
    msg = msg & "  Result: Found " & stagingCount & " rows in staging table" & vbCrLf & vbCrLf
    
    If stagingCount = 0 Then
        MsgBox msg & "ERROR: Staging table is still empty after upload!" & vbCrLf & _
               "Check mod_Database.BulkInsertToStaging function.", vbCritical
        Exit Sub
    End If
    
    ' STEP 6: Check inflight table
    msg = msg & "STEP 6: Checking tbl_pif_cost_inflight..." & vbCrLf
    step = 6
    Dim inflightCount As Long
    inflightCount = GetInflightCostCount()
    msg = msg & "  Result: Found " & inflightCount & " rows in inflight table" & vbCrLf & vbCrLf
    
    ' Final summary
    msg = msg & "========================================" & vbCrLf & _
              "DIAGNOSTIC COMPLETE" & vbCrLf & _
              "========================================" & vbCrLf & vbCrLf & _
              "Cost data is flowing through the system!" & vbCrLf & _
              "PIF: " & pifCostCount & " entries" & vbCrLf & _
              "Unpivoted: " & unpivotCount & " rows" & vbCrLf & _
              "Staging: " & stagingCount & " rows" & vbCrLf & _
              "Inflight: " & inflightCount & " rows"
    
    MsgBox msg, vbInformation, "Cost Flow Diagnostic Complete"
    
    Exit Sub

ErrHandler:
    MsgBox "Diagnostic failed at step " & step & ":" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ============================================================================
' DIAGNOSTIC HELPER FUNCTIONS
' ============================================================================

Private Function CountCostDataInPIF() As Long
    On Error GoTo ErrHandler
    
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim costCount As Long
    Dim i As Long
    Dim j As Long
    Dim cellValue As Variant
    
    Set wsData = ThisWorkbook.sheets("Target Adjustment")
    lastRow = wsData.Cells(wsData.Rows.Count, 8).End(xlUp).Row
    
    ' Count non-empty cost cells in columns V-BG (columns 22-59)
    costCount = 0
    For i = 4 To lastRow
        For j = 22 To 59  ' Columns V through BG
            cellValue = wsData.Cells(i, j).Value
            If Not IsEmpty(cellValue) And cellValue <> 0 Then
                costCount = costCount + 1
            End If
        Next j
    Next i
    
    CountCostDataInPIF = costCount
    Exit Function
    
ErrHandler:
    CountCostDataInPIF = -1
End Function

Private Function CountRowsInSheet(ByVal sheetName As String, Optional ByVal startRow As Long = 1) As Long
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow >= startRow Then
        CountRowsInSheet = lastRow - startRow + 1
    Else
        CountRowsInSheet = 0
    End If
    
    Exit Function

ErrHandler:
    CountRowsInSheet = -1
End Function

Private Sub ValidateCostUnpivotedData(ByRef validCount As Long, ByRef nullCount As Long, ByRef zeroCount As Long)
    On Error GoTo ErrHandler
    
    Dim wsCost As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As Variant
    
    validCount = 0
    nullCount = 0
    zeroCount = 0
    
    Set wsCost = ThisWorkbook.Sheets("Cost_Unpivoted")
    lastRow = wsCost.Cells(wsCost.Rows.Count, 1).End(xlUp).Row
    
    ' Check column F (requested_value)
    For i = 2 To lastRow
        cellValue = wsCost.Cells(i, 6).Value
        
        If IsEmpty(cellValue) Or IsNull(cellValue) Then
            nullCount = nullCount + 1
        ElseIf cellValue = 0 Then
            zeroCount = zeroCount + 1
        Else
            validCount = validCount + 1
        End If
    Next i
    
    Exit Sub

ErrHandler:
    validCount = -1
End Sub

Private Function TestCostUploadToStaging() As Boolean
    On Error GoTo ErrHandler
    
    Dim wsCost As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long
    
    Set wsCost = ThisWorkbook.Sheets("Cost_Unpivoted")
    lastRow = wsCost.Cells(wsCost.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        TestCostUploadToStaging = False
        Exit Function
    End If
    
    Set dataRange = wsCost.Range(wsCost.Cells(2, 1), wsCost.Cells(lastRow, 8))
    
    TestCostUploadToStaging = mod_Database.BulkInsertToStaging(dataRange, "tbl_pif_cost_staging", "dbo")
    
    Exit Function

ErrHandler:
    TestCostUploadToStaging = False
End Function

Private Function GetStagingCostCount() As Long
    On Error GoTo ErrHandler
    
    GetStagingCostCount = mod_Database.GetRecordCount("tbl_pif_cost_staging", "dbo")
    
    Exit Function

ErrHandler:
    GetStagingCostCount = -1
End Function

Private Function GetInflightCostCount() As Long
    On Error GoTo ErrHandler
    
    GetInflightCostCount = mod_Database.GetRecordCount("tbl_pif_cost_inflight", "dbo")
    
    Exit Function

ErrHandler:
    GetInflightCostCount = -1
End Function

' ============================================================================
' DETAILED DIAGNOSTIC FUNCTIONS
' ============================================================================

Public Sub DetailedCostUnpivotAnalysis()
    On Error GoTo ErrHandler
    
    Dim wsCost As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim msg As String
    Dim pifId As String, projectId As String, lineItem As Long
    Dim scenario As String, year As Date
    Dim requested As Variant, current As Variant, variance As Variant
    
    Set wsCost = ThisWorkbook.Sheets("Cost_Unpivoted")
    lastRow = wsCost.Cells(wsCost.Rows.Count, 1).End(xlUp).Row
    
    msg = "COST_UNPIVOTED SHEET ANALYSIS" & vbCrLf & _
          String(50, "=") & vbCrLf & vbCrLf
    
    ' Show first 20 records
    msg = msg & "First 20 records from Cost_Unpivoted:" & vbCrLf & vbCrLf
    
    For i = 2 To Application.Min(21, lastRow)
        pifId = wsCost.Cells(i, 1).Value
        projectId = wsCost.Cells(i, 2).Value
        lineItem = wsCost.Cells(i, 3).Value
        scenario = wsCost.Cells(i, 4).Value
        year = wsCost.Cells(i, 5).Value
        requested = wsCost.Cells(i, 6).Value
        current = wsCost.Cells(i, 7).Value
        variance = wsCost.Cells(i, 8).Value
        
        msg = msg & i - 1 & ". PIF=" & pifId & " Proj=" & projectId & " Line=" & lineItem & vbCrLf & _
                      "    Scenario=" & scenario & " Year=" & Format(year, "yyyy") & vbCrLf & _
                      "    Req=" & IIf(IsEmpty(requested), "NULL", requested) & _
                      " Curr=" & IIf(IsEmpty(current), "NULL", current) & _
                      " Var=" & IIf(IsEmpty(variance), "NULL", variance) & vbCrLf & vbCrLf
    Next i
    
    MsgBox msg, vbInformation, "Cost Unpivot Analysis"
    
    Exit Sub

ErrHandler:
    MsgBox "Analysis failed: " & Err.Description, vbCritical
End Sub

Public Sub DetailedStagingAnalysis()
    On Error GoTo ErrHandler
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim msg As String
    Dim sql As String
    Dim count As Long
    
    Set conn = mod_Database.GetDBConnection()
    If conn Is Nothing Then
        MsgBox "Failed to connect to database", vbCritical
        Exit Sub
    End If
    
    ' Get staging count
    sql = "SELECT COUNT(*) AS RowCount FROM dbo.tbl_pif_cost_staging"
    Set rs = New ADODB.Recordset
    rs.Open sql, conn
    count = rs.Fields("RowCount").Value
    rs.Close
    
    msg = "STAGING TABLE ANALYSIS" & vbCrLf & _
          String(50, "=") & vbCrLf & vbCrLf & _
          "Total rows in tbl_pif_cost_staging: " & count & vbCrLf & vbCrLf
    
    If count = 0 Then
        msg = msg & "TABLE IS EMPTY - Cost upload failed!" & vbCrLf & vbCrLf & _
                   "Check:" & vbCrLf & _
                   "1. BulkInsertToStaging function execution" & vbCrLf & _
                   "2. Database connection" & vbCrLf & _
                   "3. Column references in UnpivotCostData"
    Else
        ' Show first 10 records
        msg = msg & "First 10 records:" & vbCrLf & vbCrLf
        
        sql = "SELECT TOP 10 pif_id, project_id, line_item, scenario, " & _
              "YEAR(year) AS FiscalYear, requested_value FROM dbo.tbl_pif_cost_staging " & _
              "ORDER BY pif_id, project_id, line_item, scenario, year"
        Set rs = New ADODB.Recordset
        rs.Open sql, conn
        
        Dim rowNum As Long
        rowNum = 1
        Do While Not rs.EOF
            msg = msg & rowNum & ". PIF=" & rs("pif_id").Value & " Proj=" & rs("project_id").Value & _
                        " Line=" & rs("line_item").Value & " Scenario=" & rs("scenario").Value & _
                        " Year=" & rs("FiscalYear").Value & " Requested=" & rs("requested_value").Value & vbCrLf
            rs.MoveNext
            rowNum = rowNum + 1
        Loop
        rs.Close
    End If
    
    conn.Close
    Set conn = Nothing
    
    MsgBox msg, vbInformation, "Staging Table Analysis"
    
    Exit Sub

ErrHandler:
    MsgBox "Analysis failed: " & Err.Description, vbCritical
End Sub


' ============================================================================
' CORRECTED: UploadCostData and BulkInsertCosts Functions
' ============================================================================
' Issue: These functions are incomplete and not properly calling the database upload
' Solution: Provide complete, working implementations
' ============================================================================

' Location: mod_Submit.bas

' ============================================================================
' PART 1: CORRECTED UploadCostData Function
' ============================================================================
' Purpose: Upload unpivoted cost data to the staging table
' This function:
'   1. Gets the Cost_Unpivoted worksheet
'   2. Finds the data range (excluding header)
'   3. Calls BulkInsertCosts with the range
'   4. Returns success/failure status

Private Function UploadCostData() As Boolean
    On Error GoTo ErrHandler

    Dim wsCost As Worksheet
    Dim dataRange As Range
    Dim lastDataRow As Long
    Dim success As Boolean

    ' Get the Cost_Unpivoted sheet
    On Error Resume Next
    Set wsCost = ThisWorkbook.Sheets(SHEET_COST_UNPIVOTED)
    On Error GoTo ErrHandler

    If wsCost Is Nothing Then
        Debug.Print "ERROR: Cost_Unpivoted sheet not found"
        UploadCostData = False
        Exit Function
    End If

    ' Find the last row with data (check column A for pif_id)
    lastDataRow = wsCost.Cells(wsCost.Rows.Count, 1).End(xlUp).Row

    ' If no data rows (only header or empty), return success (nothing to upload)
    If lastDataRow < 2 Then
        Debug.Print "WARNING: Cost_Unpivoted sheet has no data rows"
        UploadCostData = True  ' Not an error - just no costs to upload
        Exit Function
    End If

    ' Define the data range from row 2 (first data row) to last row, columns A-H
    Set dataRange = wsCost.Range(wsCost.Cells(2, 1), wsCost.Cells(lastDataRow, 8))

    Debug.Print "UploadCostData: Range=" & dataRange.Address & " Rows=" & dataRange.Rows.Count

    ' Upload the data
    success = BulkInsertCosts(dataRange)

    UploadCostData = success
    Exit Function

ErrHandler:
    Debug.Print "ERROR in UploadCostData: " & Err.Number & " - " & Err.Description
    MsgBox "Failed to upload cost data:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, vbCritical
    UploadCostData = False
End Function

' ============================================================================
' PART 2: CORRECTED BulkInsertCosts Function
' ============================================================================
' Purpose: Wrapper function to upload cost data from Cost_Unpivoted sheet
' Calls: mod_Database.BulkInsertToStaging with proper parameters
' Parameters:
'   dataRange - Excel range containing cost data (columns A-H)
'   tableName - Target database table (tbl_pif_cost_staging)
' Returns: True if successful, False if failed

Public Function BulkInsertCosts(ByVal dataRange As Range) As Boolean
    On Error GoTo ErrHandler

    Dim success As Boolean
    Dim rowCount As Long

    Debug.Print "=== BulkInsertCosts STARTED ==="
    Debug.Print "Data range: " & dataRange.Address
    Debug.Print "Rows: " & dataRange.Rows.Count

    ' Validate the range
    If dataRange.Rows.Count = 0 Then
        Debug.Print "ERROR: Data range is empty"
        BulkInsertCosts = True  ' Not an error - just nothing to insert
        Exit Function
    End If

    ' Call the bulk insert function from mod_Database
    ' This handles:
    '   - Database connection
    '   - Transaction management
    '   - Truncation of staging table
    '   - Bulk insert of all rows
    success = mod_Database.BulkInsertToStaging(dataRange, "tbl_pif_cost_staging", "dbo")

    If Not success Then
        Debug.Print "ERROR: BulkInsertToStaging returned False"
        BulkInsertCosts = False
        Exit Function
    End If

    Debug.Print "=== BulkInsertCosts COMPLETED SUCCESSFULLY ==="
    BulkInsertCosts = True
    Exit Function

ErrHandler:
    Debug.Print "ERROR in BulkInsertCosts: " & Err.Number & " - " & Err.Description
    MsgBox "Failed to bulk insert cost data:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, vbCritical
    BulkInsertCosts = False
End Function

' ============================================================================
' PART 3: NEW Helper Function - GetSelectedSite (For reference)
' ============================================================================
' Purpose: Extract selected site for cost filtering during upload
' Returns: Site code or empty string

Private Function GetSelectedSiteForCosts() As String
    On Error Resume Next

    Dim selectedSite As String
    selectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.Value)

    If selectedSite = "" Then
        GetSelectedSiteForCosts = ""
    Else
        GetSelectedSiteForCosts = selectedSite
    End If
End Function

' ============================================================================
' IMPORTANT NOTES
' ============================================================================
' 
' 1. The Cost_Unpivoted sheet MUST have data in columns A-H:
'    Column A: pif_id
'    Column B: project_id
'    Column C: line_item
'    Column D: scenario
'    Column E: year (as DATE format)
'    Column F: requested_value
'    Column G: current_value
'    Column H: variance_value
'
' 2. The UnpivotCostData function MUST run before UploadCostData
'    to populate the Cost_Unpivoted sheet
'
' 3. If UploadCostData returns False, check:
'    - Is Cost_Unpivoted sheet visible? (It should be hidden but accessible)
'    - Does it have data rows (row 2+)?
'    - Is the database connection working?
'    - Are tbl_pif_cost_staging permissions correct?
'
' 4. Testing: Use mod_CostDiagnostic.DiagnoseFullCostFlow() to trace the flow
'
