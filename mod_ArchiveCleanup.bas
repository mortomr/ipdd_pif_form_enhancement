Attribute VB_Name = "mod_ArchiveCleanup"
' ============================================================================
' MODULE: mod_ArchiveCleanup
' ============================================================================
' Purpose: Clean up archived records from PIF data entry worksheet
' Author: Data Architecture Team
' Date: 2025-11-13
'
' Features:
'   - Validates which PIF records are archived in the database
'   - Deletes archived records from PIF worksheet to prep for next month
'   - Enabled for SITE contexts only (disabled for Fleet read-only view)
'   - Uses database validation to ensure only archived records are deleted
'
' Workflow:
'   1. User finalizes month and clicks [Submit to Database]
'   2. User navigates to Archive sheet to confirm records are captured
'   3. User clicks [Delete Archived Records] button to remove them from PIF sheet
'   4. System queries database to identify archived records for current site
'   5. Matching records are deleted from PIF sheet with confirmation
' ============================================================================

Option Explicit

' ============================================================================
' PUBLIC FUNCTIONS - ARCHIVE CLEANUP
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: Archive_DeleteArchivedRecords
' Purpose: Delete archived records from PIF worksheet (SITE contexts only)
' Usage: Attach to [Delete Archived Records] button on PIF_Archive sheet
' Notes:
'   - Disabled for Fleet context (read-only)
'   - Queries database to validate which records are archived
'   - Provides detailed confirmation before deletion
' ----------------------------------------------------------------------------
Public Sub Archive_DeleteArchivedRecords()
    On Error GoTo ErrHandler

    Dim selectedSite As String
    Dim archivedRecords As Collection
    Dim recordsToDelete As Collection
    Dim ws As Worksheet
    Dim response As VbMsgBoxResult
    Dim deleteCount As Long
    Dim startTime As Double

    startTime = Timer

    ' Step 1: Validate site context (Fleet users cannot delete)
    selectedSite = mod_SiteSetup.GetSelectedSite()

    If selectedSite = "" Then
        MsgBox "Please select a site from the Instructions worksheet before cleaning up archived records.", _
               vbExclamation, "Site Not Selected"
        Exit Sub
    End If

    If UCase(Trim(selectedSite)) = "FLEET" Then
        MsgBox "Archive cleanup is disabled for Fleet context." & vbCrLf & vbCrLf & _
               "This feature is only available for individual sites (ANO, GGN, RBS, WF3, HQN)." & vbCrLf & vbCrLf & _
               "Reason: Fleet context is read-only and displays all sites' data.", _
               vbExclamation, "Feature Not Available"
        Exit Sub
    End If

    ' Step 2: Get PIF worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_SharedConstants.SHEET_DATA)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        MsgBox "Cannot find PIF worksheet." & vbCrLf & vbCrLf & _
               "Expected sheet name: " & mod_SharedConstants.SHEET_DATA, _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If

    Application.StatusBar = "Querying database for archived records..."
    Application.Cursor = xlWait

    ' Step 3: Query database for archived records
    Set archivedRecords = GetArchivedRecordsForSite(selectedSite)

    If archivedRecords Is Nothing Or archivedRecords.Count = 0 Then
        Application.StatusBar = False
        Application.Cursor = xlDefault

        MsgBox "No archived records found in the database for site: " & selectedSite & vbCrLf & vbCrLf & _
               "This could mean:" & vbCrLf & _
               "  - No records have been approved/dispositioned yet" & vbCrLf & _
               "  - All archived records have already been removed from the PIF sheet" & vbCrLf & vbCrLf & _
               "TIP: Refresh the Archive sheet to see all archived records.", _
               vbInformation, "No Archived Records"
        Exit Sub
    End If

    Application.StatusBar = "Matching archived records in PIF worksheet..."

    ' Step 4: Find matching records in PIF worksheet
    Set recordsToDelete = FindMatchingRecordsInWorksheet(ws, archivedRecords, selectedSite)

    Application.StatusBar = False
    Application.Cursor = xlDefault

    If recordsToDelete Is Nothing Or recordsToDelete.Count = 0 Then
        MsgBox "No matching archived records found in the PIF worksheet." & vbCrLf & vbCrLf & _
               "Database has " & archivedRecords.Count & " archived record(s) for site " & selectedSite & ", " & _
               "but none of them are currently in the PIF worksheet." & vbCrLf & vbCrLf & _
               "This is normal if you've already cleaned up this month's data.", _
               vbInformation, "No Records to Delete"
        Exit Sub
    End If

    ' Step 5: Confirmation prompt with detailed information
    deleteCount = recordsToDelete.Count

    Dim confirmMsg As String
    confirmMsg = "DELETE ARCHIVED RECORDS FROM PIF WORKSHEET" & vbCrLf & _
                 String(50, "=") & vbCrLf & vbCrLf & _
                 "Site: " & selectedSite & vbCrLf & _
                 "Records to delete: " & deleteCount & vbCrLf & vbCrLf & _
                 "These records have been APPROVED or DISPOSITIONED and are safely stored " & _
                 "in the database archive." & vbCrLf & vbCrLf & _
                 "This action will:" & vbCrLf & _
                 "  - Remove archived records from the PIF worksheet" & vbCrLf & _
                 "  - Prepare the worksheet for next month's data entry" & vbCrLf & _
                 "  - NOT affect archived records in the database" & vbCrLf & vbCrLf & _
                 "WARNING: This cannot be undone!" & vbCrLf & vbCrLf & _
                 "Do you want to proceed?"

    response = MsgBox(confirmMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Deletion")

    If response = vbNo Then
        MsgBox "Deletion cancelled. No records were removed.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Step 6: Delete the records
    Application.ScreenUpdating = False
    Application.StatusBar = "Deleting archived records..."

    Call DeleteRecordsFromWorksheet(ws, recordsToDelete)

    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim elapsed As Double
    elapsed = Timer - startTime

    ' Step 7: Success message
    MsgBox "Archive cleanup complete!" & vbCrLf & vbCrLf & _
           "Site: " & selectedSite & vbCrLf & _
           "Records deleted: " & deleteCount & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "The PIF worksheet is now ready for next month's data entry." & vbCrLf & vbCrLf & _
           "NOTE: Archived records remain safely stored in the database " & _
           "and can be viewed on the Archive sheet.", _
           vbInformation, "Cleanup Complete"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.Cursor = xlDefault

    MsgBox "Error during archive cleanup:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "No records were deleted. Please try again or contact support.", _
           vbCritical, "Cleanup Error"
End Sub

' ----------------------------------------------------------------------------
' Function: IsArchiveCleanupEnabled
' Purpose: Check if archive cleanup is enabled for current site context
' Returns: True if enabled (SITE contexts), False if disabled (Fleet)
' Usage: Can be used to enable/disable button in UI
' ----------------------------------------------------------------------------
Public Function IsArchiveCleanupEnabled() As Boolean
    On Error Resume Next

    Dim selectedSite As String
    selectedSite = mod_SiteSetup.GetSelectedSite()

    ' Enabled for specific sites, disabled for Fleet or no selection
    If selectedSite = "" Then
        IsArchiveCleanupEnabled = False
    ElseIf UCase(Trim(selectedSite)) = "FLEET" Then
        IsArchiveCleanupEnabled = False
    Else
        IsArchiveCleanupEnabled = True
    End If
End Function

' ============================================================================
' PRIVATE HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetArchivedRecordsForSite
' Purpose: Query database for archived PIF_ID + Project_ID combinations
' Parameters:
'   siteName - Site code (ANO, GGN, RBS, WF3, HQN)
' Returns: Collection of "PIF_ID|Project_ID" strings, or Nothing if error
' Notes: Queries tbl_pif_projects_approved filtered by site
' ----------------------------------------------------------------------------
Private Function GetArchivedRecordsForSite(ByVal siteName As String) As Collection
    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim recordKey As String
    Dim results As Collection

    Set results = New Collection

    ' Get database connection
    Set conn = mod_Database.GetDBConnection()
    If conn Is Nothing Then
        Set GetArchivedRecordsForSite = Nothing
        Exit Function
    End If

    ' Build SQL query - get distinct PIF_ID + Project_ID combinations
    ' Note: Records in tbl_pif_projects_approved are already approved by definition
    '       No need to filter by status - if it's in this table, it's archived
    sql = "SELECT DISTINCT pif_id, project_id " & _
          "FROM dbo.tbl_pif_projects_approved " & _
          "WHERE UPPER(site) = '" & UCase(Trim(siteName)) & "' " & _
          "ORDER BY pif_id, project_id"

    ' Execute query
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly, adLockReadOnly

    ' Build collection of record keys
    Do While Not rs.EOF
        ' Create composite key: "PIF_ID|Project_ID"
        recordKey = Trim(CStr(rs("pif_id").Value)) & "|" & Trim(CStr(rs("project_id").Value))

        ' Add to collection (use key as both item and key to prevent duplicates)
        On Error Resume Next
        results.Add recordKey, recordKey
        On Error GoTo ErrHandler

        rs.MoveNext
    Loop

    ' Cleanup
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    Set GetArchivedRecordsForSite = results
    Exit Function

ErrHandler:
    ' Cleanup on error
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set rs = Nothing
    Set conn = Nothing
    On Error GoTo 0

    Set GetArchivedRecordsForSite = Nothing
End Function

' ----------------------------------------------------------------------------
' Function: FindMatchingRecordsInWorksheet
' Purpose: Find rows in PIF worksheet that match archived record keys
' Parameters:
'   ws - PIF worksheet
'   archivedRecords - Collection of "PIF_ID|Project_ID" keys
'   siteName - Site code (for validation)
' Returns: Collection of row numbers to delete, or Nothing if none found
' ----------------------------------------------------------------------------
Private Function FindMatchingRecordsInWorksheet(ByVal ws As Worksheet, _
                                                ByVal archivedRecords As Collection, _
                                                ByVal siteName As String) As Collection
    On Error GoTo ErrHandler

    Dim results As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim pifId As String
    Dim projectId As String
    Dim rowSite As String
    Dim recordKey As String

    Set results = New Collection

    ' Find last row with data in PIF_ID column (Column G = 7)
    lastRow = ws.Cells(ws.Rows.Count, PIFDataColumns.colPIFID).End(xlUp).Row

    ' Loop through data rows (start at row 4, skip headers in rows 1-3)
    For i = 4 To lastRow
        ' Get PIF_ID, Project_ID, and Site from row
        pifId = Trim(CStr(ws.Cells(i, PIFDataColumns.colPIFID).Value))
        projectId = Trim(CStr(ws.Cells(i, PIFDataColumns.colFundingProject).Value))
        rowSite = Trim(CStr(ws.Cells(i, PIFDataColumns.colSite).Value))

        ' Skip empty rows or rows without PIF_ID
        If pifId = "" Then GoTo NextRow

        ' Validate site matches (safety check - only delete records for current site)
        If UCase(rowSite) <> UCase(siteName) Then GoTo NextRow

        ' Build composite key
        recordKey = pifId & "|" & projectId

        ' Check if this key exists in archived records collection
        On Error Resume Next
        Dim testItem As String
        testItem = archivedRecords(recordKey)

        If Err.Number = 0 Then
            ' Match found - add row number to deletion list
            results.Add i
        End If
        On Error GoTo ErrHandler

NextRow:
    Next i

    Set FindMatchingRecordsInWorksheet = results
    Exit Function

ErrHandler:
    Set FindMatchingRecordsInWorksheet = Nothing
End Function

' ----------------------------------------------------------------------------
' Sub: DeleteRecordsFromWorksheet
' Purpose: Delete specified rows from PIF worksheet
' Parameters:
'   ws - PIF worksheet
'   rowNumbers - Collection of row numbers to delete (must be sorted descending)
' Notes: Deletes from bottom to top to maintain correct row indices
'        Handles both Excel Tables (ListObject) and regular ranges
' ----------------------------------------------------------------------------
Private Sub DeleteRecordsFromWorksheet(ByVal ws As Worksheet, _
                                      ByVal rowNumbers As Collection)
    On Error GoTo ErrHandler

    Dim tbl As ListObject
    Dim i As Long
    Dim rowNum As Long
    Dim sortedRows() As Long
    Dim j As Long
    Dim temp As Long

    ' Check if PIF worksheet uses an Excel Table (ListObject)
    On Error Resume Next
    Set tbl = ws.ListObjects(1)  ' Assume first table on sheet
    On Error GoTo ErrHandler

    ' Copy collection to array for sorting
    ReDim sortedRows(1 To rowNumbers.Count)
    For i = 1 To rowNumbers.Count
        sortedRows(i) = rowNumbers(i)
    Next i

    ' Sort row numbers in DESCENDING order (delete from bottom to top)
    ' This prevents row index shifting during deletion
    For i = 1 To UBound(sortedRows) - 1
        For j = i + 1 To UBound(sortedRows)
            If sortedRows(i) < sortedRows(j) Then
                temp = sortedRows(i)
                sortedRows(i) = sortedRows(j)
                sortedRows(j) = temp
            End If
        Next j
    Next i

    If Not tbl Is Nothing Then
        ' ========================================================================
        ' EXCEL TABLE DELETION (ListObject)
        ' ========================================================================
        ' Use ListRows collection for Excel Tables

        ' Disable filters temporarily if active
        Dim filterWasOn As Boolean
        filterWasOn = False

        If tbl.ShowAutoFilter Then
            If tbl.AutoFilter.FilterMode Then
                filterWasOn = True
                tbl.AutoFilter.ShowAllData
            End If
        End If

        ' Delete rows (already sorted descending)
        For i = 1 To UBound(sortedRows)
            rowNum = sortedRows(i)

            ' Convert worksheet row to ListRow index
            ' ListRow index = worksheet row - header row - 1
            Dim listRowIndex As Long
            listRowIndex = rowNum - tbl.HeaderRowRange.Row

            ' Validate index is within table bounds
            If listRowIndex >= 1 And listRowIndex <= tbl.ListRows.Count Then
                tbl.ListRows(listRowIndex).Delete
            End If
        Next i

        ' AutoFilter is automatically restored by Excel

    Else
        ' ========================================================================
        ' REGULAR RANGE DELETION (Non-Table)
        ' ========================================================================
        ' Standard row deletion for non-table ranges

        For i = 1 To UBound(sortedRows)
            rowNum = sortedRows(i)
            ws.Rows(rowNum).Delete Shift:=xlUp
        Next i
    End If

    Exit Sub

ErrHandler:
    ' Error during deletion - log but don't stop (partial deletion is OK)
    Debug.Print "Error deleting row: " & Err.Number & " - " & Err.Description
    Resume Next
End Sub

' ============================================================================
' BACKWARD COMPATIBILITY WRAPPERS
' ============================================================================

Public Sub DeleteArchivedRecords()
    Call Archive_DeleteArchivedRecords
End Sub
