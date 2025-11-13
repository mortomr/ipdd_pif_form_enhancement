Attribute VB_Name = "mod_WorksheetQuery"
' ============================================================================
' MODULE: mod_WorksheetQuery
' ============================================================================
' Purpose: Create and refresh Archive and Inflight query worksheets
' Author: Data Architecture Team
' Date: 2025-11-11
'
' Features:
'   - Populates PIF_Archive worksheet from vw_pif_approved_wide view
'   - Populates PIF_Inflight worksheet from vw_pif_inflight_wide view
'   - Filters by selected site (respects SelectedSite named range)
'   - Displays data in original PIF table format for familiarity
'   - Auto-refreshable with button click
' ============================================================================

Option Explicit

' Sheet name constants
Private Const SHEET_ARCHIVE As String = "PIF_Archive"
Private Const SHEET_INFLIGHT As String = "PIF_Inflight"

' ============================================================================
' PUBLIC FUNCTIONS - WORKSHEET CREATION AND REFRESH
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: RefreshArchiveWorksheet
' Purpose: Create or refresh the Archive worksheet with approved records
' Usage: Call from button or manually
' Notes: Uses Excel QueryTable for native refresh capability (no VBA needed)
' ----------------------------------------------------------------------------
Public Sub Nav_RefreshArchive()
    On Error GoTo ErrHandler

    Dim selectedSite As String
    Dim ws As Worksheet
    Dim sql As String
    Dim startTime As Double
    Dim qt As QueryTable
    Dim connStr As String
    Dim tbl As ListObject

    startTime = Timer

    ' Get selected site
    selectedSite = mod_SiteSetup.GetSelectedSite()
    If selectedSite = "" Then
        MsgBox "Please select a site from the Instructions worksheet before refreshing.", _
               vbExclamation, "Site Not Selected"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Archive worksheet..."

    ' Get or create worksheet
    Set ws = GetOrCreateWorksheet(SHEET_ARCHIVE, "Archive (Read-Only)")

    ' Build SQL query with site filter
    sql = "SELECT * FROM dbo.vw_pif_approved_wide"

    ' Add site filter (Fleet sees all sites, others see only their site)
    If UCase(Trim(selectedSite)) <> "FLEET" Then
        sql = sql & " WHERE UPPER(site) = '" & UCase(selectedSite) & "'"
    End If

    sql = sql & " ORDER BY approval_date DESC, pif_id, project_id"

    ' Build connection string using Windows Authentication
    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    ' Create or refresh QueryTable
    Call CreateOrRefreshQueryTable(ws, "ArchiveQuery", sql, connStr, selectedSite)

    ws.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim elapsed As Double
    elapsed = Timer - startTime

    ' Count records (excluding header)
    Dim recordCount As Long
    recordCount = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row - 4
    If recordCount < 0 Then recordCount = 0

    MsgBox "Archive worksheet refreshed!" & vbCrLf & vbCrLf & _
           "Site: " & selectedSite & vbCrLf & _
           "Records: " & recordCount & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "NOTE: You can refresh this data anytime by:" & vbCrLf & _
           "  1. Right-clicking on the table" & vbCrLf & _
           "  2. Selecting 'Refresh' from the context menu" & vbCrLf & _
           "  3. Or using Data > Refresh All on the ribbon", _
           vbInformation, "Refresh Complete"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Failed to refresh Archive worksheet:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Refresh Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshInflightWorksheet
' Purpose: Create or refresh the Inflight worksheet with current working records
' Usage: Call from button or manually
' Notes: Uses Excel QueryTable for native refresh capability (no VBA needed)
' ----------------------------------------------------------------------------
Public Sub Nav_RefreshInflight()
    On Error GoTo ErrHandler

    Dim selectedSite As String
    Dim ws As Worksheet
    Dim sql As String
    Dim startTime As Double
    Dim qt As QueryTable
    Dim connStr As String
    Dim tbl As ListObject

    startTime = Timer

    ' Get selected site
    selectedSite = mod_SiteSetup.GetSelectedSite()
    If selectedSite = "" Then
        MsgBox "Please select a site from the Instructions worksheet before refreshing.", _
               vbExclamation, "Site Not Selected"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Inflight worksheet..."

    ' Get or create worksheet
    Set ws = GetOrCreateWorksheet(SHEET_INFLIGHT, "Inflight (Read-Only)")

    ' Build SQL query with site filter
    sql = "SELECT * FROM dbo.vw_pif_inflight_wide"

    ' Add site filter (Fleet sees all sites, others see only their site)
    If UCase(Trim(selectedSite)) <> "FLEET" Then
        sql = sql & " WHERE UPPER(site) = '" & UCase(selectedSite) & "'"
    End If

    sql = sql & " ORDER BY submission_date DESC, pif_id, project_id"

    ' Build connection string using Windows Authentication
    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    ' Create or refresh QueryTable
    Call CreateOrRefreshQueryTable(ws, "InflightQuery", sql, connStr, selectedSite)

    ws.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim elapsed As Double
    elapsed = Timer - startTime

    ' Count records (excluding header)
    Dim recordCount As Long
    recordCount = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row - 4
    If recordCount < 0 Then recordCount = 0

    MsgBox "Inflight worksheet refreshed!" & vbCrLf & vbCrLf & _
           "Site: " & selectedSite & vbCrLf & _
           "Records: " & recordCount & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "NOTE: You can refresh this data anytime by:" & vbCrLf & _
           "  1. Right-clicking on the table" & vbCrLf & _
           "  2. Selecting 'Refresh' from the context menu" & vbCrLf & _
           "  3. Or using Data > Refresh All on the ribbon", _
           vbInformation, "Refresh Complete"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Failed to refresh Inflight worksheet:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Refresh Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshBothWorksheets
' Purpose: Refresh both Archive and Inflight worksheets
' Parameters:
'   Optional showSuccessMessage - Set to False to suppress success message
' Usage: Call after submission/archival to update both views
' ----------------------------------------------------------------------------
Public Sub Nav_RefreshAll(Optional ByVal showSuccessMessage As Boolean = True)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    ' Refresh both without individual messageboxes
    Call RefreshArchiveWorksheetSilent
    Call RefreshInflightWorksheetSilent

    Application.ScreenUpdating = True

    ' Only show success message if requested (suppress during automated workflows)
    If showSuccessMessage Then
        MsgBox "Both Archive and Inflight worksheets have been refreshed.", _
               vbInformation, "Refresh Complete"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error refreshing worksheets:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Refresh Error"
End Sub

' ============================================================================
' PRIVATE HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: CreateOrRefreshQueryTable
' Purpose: Create or refresh a QueryTable connection to SQL Server
' Parameters:
'   ws - Target worksheet
'   queryName - Name for the QueryTable
'   sql - SQL query string
'   connStr - Connection string
'   siteName - Site name for table formatting
' Notes: Creates native Excel QueryTable with full refresh capability
'        Uses QueryTable-only approach (no ListObject) for reliability
'
' Connection Management:
'   - CRITICAL: Deletes orphaned WorkbookConnections BEFORE deleting QueryTables
'   - This prevents abandoned OLEDB connections in SQL Server
'   - Each refresh would otherwise create a new connection without closing old ones
'   - SAFETY: Only deletes connections specifically used by THIS worksheet's QueryTables
'   - Does NOT touch other connections in the workbook (protects other QueryTables/Power Query)
' ----------------------------------------------------------------------------
Private Sub CreateOrRefreshQueryTable(ByVal ws As Worksheet, _
                                      ByVal queryName As String, _
                                      ByVal sql As String, _
                                      ByVal connStr As String, _
                                      ByVal siteName As String)
    On Error GoTo ErrHandler

    Dim qt As QueryTable
    Dim i As Integer
    Dim conn As WorkbookConnection

    ' Clear worksheet
    ws.Cells.Clear

    ' CRITICAL FIX: Delete orphaned WorkbookConnections BEFORE deleting QueryTables
    ' This prevents abandoned database connections in SQL Server
    ' SAFETY: Only delete connections specifically created by this worksheet's QueryTables
    Dim qtConn As WorkbookConnection
    Dim qtConnName As String
    Dim connToDelete As Collection
    Set connToDelete = New Collection

    ' First pass: Identify connections used by QueryTables on THIS worksheet
    For i = ws.QueryTables.Count To 1 Step -1
        Set qt = ws.QueryTables(i)
        ' Extract connection name from QueryTable's connection string
        ' QueryTable connections typically named "Query from <source>" or custom names
        On Error Resume Next
        qtConnName = ""
        For Each qtConn In ThisWorkbook.Connections
            If qtConn.Type = xlConnectionTypeOLEDB Then
                ' Check if this connection matches the QueryTable's connection string
                If InStr(1, qt.Connection, qtConn.Name, vbTextCompare) > 0 Then
                    qtConnName = qtConn.Name
                    connToDelete.Add qtConnName, qtConnName  ' Use name as key to avoid duplicates
                    Exit For
                End If
            End If
        Next qtConn
        On Error GoTo ErrHandler
    Next i

    ' Second pass: Delete only the connections we identified (orphaned from this worksheet)
    For i = connToDelete.Count To 1 Step -1
        qtConnName = connToDelete(i)
        On Error Resume Next
        Set conn = ThisWorkbook.Connections(qtConnName)
        If Not conn Is Nothing Then
            conn.Delete
        End If
        On Error GoTo ErrHandler
    Next i

    ' Now delete existing QueryTables on this worksheet
    For i = ws.QueryTables.Count To 1 Step -1
        ws.QueryTables(i).Delete
    Next i

    ' Delete existing ListObjects (Tables) - we're using QueryTable-only approach
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i

    ' Create new QueryTable starting at B4
    Set qt = ws.QueryTables.Add( _
        Connection:=connStr, _
        Destination:=ws.Range("B4"), _
        sql:=sql)

    ' Configure QueryTable properties for full refresh capability
    With qt
        .Name = queryName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells  ' Allows data to grow/shrink on refresh
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True

        ' Refresh to populate data
        .Refresh BackgroundQuery:=False
    End With

    ' Apply table-like formatting to QueryTable result
    If Not qt.ResultRange Is Nothing Then
        Dim dataRange As Range
        Set dataRange = qt.ResultRange

        ' Format header row (first row of result)
        With dataRange.Rows(1)
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' Add borders to entire data range
        With dataRange.Borders
            .LineStyle = xlContinuous
            .Color = RGB(208, 206, 206)
            .Weight = xlThin
        End With

        ' Apply zebra striping to data rows (skip header)
        If dataRange.Rows.Count > 1 Then
            Dim row As Long
            For row = 2 To dataRange.Rows.Count Step 2
                dataRange.Rows(row).Interior.Color = RGB(242, 242, 242)
            Next row
        End If

        ' Enable AutoFilter on the data range
        If Not dataRange.Worksheet.AutoFilterMode Then
            dataRange.AutoFilter
        End If
    End If

    ' Format the worksheet
    ws.Columns("B:ZZ").AutoFit

    ' Temporarily enable screen updating to avoid Error 1004 with Select/FreezePanes
    Dim screenUpdateState As Boolean
    screenUpdateState = Application.ScreenUpdating
    On Error Resume Next  ' Handle any errors with Select/FreezePanes gracefully
    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("B5").Select  ' Freeze below header row at B4
    ActiveWindow.FreezePanes = False  ' Clear existing freeze first
    ActiveWindow.FreezePanes = True
    ws.Range("B4").Select
    Application.ScreenUpdating = screenUpdateState
    On Error GoTo ErrHandler  ' Resume normal error handling

    ' Add title header in row 1
    ws.Range("B1").Value = UCase(Replace(queryName, "Query", "")) & " - " & siteName
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14

    ' Add instructions in row 2 (updated to reflect working refresh)
    ws.Range("B2").Value = "Refresh: Right-click data range > Refresh, or Data > Refresh All (native Excel refresh enabled)"
    ws.Range("B2").Font.Italic = True
    ws.Range("B2").Font.Size = 9
    ws.Range("B2").Font.Color = RGB(0, 128, 0)  ' Green to indicate working feature

    Exit Sub

ErrHandler:
    MsgBox "Error creating QueryTable:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "QueryTable Error"
End Sub

' ----------------------------------------------------------------------------
' Function: GetOrCreateWorksheet
' Purpose: Get existing worksheet or create new one
' Parameters:
'   sheetName - Name of worksheet
'   displayName - Display name for sheet tab
' Returns: Worksheet object
' ----------------------------------------------------------------------------
Private Function GetOrCreateWorksheet(ByVal sheetName As String, _
                                     ByVal displayName As String) As Worksheet
    Dim ws As Worksheet

    ' Try to get existing worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' Create if doesn't exist
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If

    ' Unprotect if protected
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    Set GetOrCreateWorksheet = ws
End Function

' ----------------------------------------------------------------------------
' Sub: WriteRecordsetHeaders
' Purpose: Write column headers from recordset to worksheet
' Parameters:
'   ws - Target worksheet
'   rs - Source recordset
' ----------------------------------------------------------------------------
Private Sub WriteRecordsetHeaders(ByVal ws As Worksheet, ByVal rs As ADODB.Recordset)
    Dim i As Integer
    Const HEADER_ROW As Long = 4
    Const START_COL As Long = 2  ' Column B

    For i = 0 To rs.Fields.Count - 1
        ws.Cells(HEADER_ROW, START_COL + i).Value = rs.Fields(i).Name
    Next i

    ' Format header row
    With ws.Rows(HEADER_ROW)
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ----------------------------------------------------------------------------
' Function: WriteRecordsetData
' Purpose: Write recordset data to worksheet
' Parameters:
'   ws - Target worksheet
'   rs - Source recordset
'   startRow - Starting row number (usually 2, after headers)
' Returns: Number of rows written
' ----------------------------------------------------------------------------
Private Function WriteRecordsetData(ByVal ws As Worksheet, _
                                   ByVal rs As ADODB.Recordset, _
                                   ByVal startRow As Long) As Long
    Dim rowNum As Long
    Dim colNum As Integer
    Dim recordCount As Long
    Const START_COL As Long = 2  ' Column B

    rowNum = startRow
    recordCount = 0

    Do While Not rs.EOF
        For colNum = 0 To rs.Fields.Count - 1
            If Not IsNull(rs.Fields(colNum).Value) Then
                ws.Cells(rowNum, START_COL + colNum).Value = rs.Fields(colNum).Value
            End If
        Next colNum

        rowNum = rowNum + 1
        recordCount = recordCount + 1

        ' Progress indicator every 100 rows
        If recordCount Mod 100 = 0 Then
            Application.StatusBar = "Writing row " & recordCount & "..."
        End If

        rs.MoveNext
    Loop

    WriteRecordsetData = recordCount
End Function

' ----------------------------------------------------------------------------
' Sub: FormatQueryWorksheet
' Purpose: Apply formatting to query worksheet
' Parameters:
'   ws - Worksheet to format
'   title - Title for header
'   lastRow - Last row with data
' ----------------------------------------------------------------------------
Private Sub FormatQueryWorksheet(ByVal ws As Worksheet, _
                                ByVal title As String, _
                                ByVal lastRow As Long)
    Const HEADER_ROW As Long = 4
    Const START_COL As Long = 2  ' Column B
    Dim lastCol As Long

    ' Auto-fit columns
    ws.Cells.Columns.AutoFit

    ' Freeze panes at B5 (after headers in row 4)
    ws.Activate
    ws.Range("B5").Select
    ActiveWindow.FreezePanes = True
    ws.Range("B4").Select

    ' Add auto-filter - determine last column with data
    If lastRow > HEADER_ROW Then
        lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
        ws.Range(ws.Cells(HEADER_ROW, START_COL), ws.Cells(lastRow, lastCol)).AutoFilter
    End If

    ' Format as table (optional - remove if not desired)
    ' ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = Replace(title, " ", "_")

    ' Protect worksheet (allow filtering and formatting)
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingCells:=True, _
                AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowSorting:=True
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshArchiveWorksheetSilent
' Purpose: Refresh Archive worksheet without messagebox (for batch operations)
' Notes: Uses Excel QueryTable for native refresh capability
' ----------------------------------------------------------------------------
Private Sub RefreshArchiveWorksheetSilent()
    On Error Resume Next

    Dim selectedSite As String
    Dim ws As Worksheet
    Dim sql As String
    Dim connStr As String

    selectedSite = mod_SiteSetup.GetSelectedSite()
    If selectedSite = "" Then Exit Sub

    Set ws = GetOrCreateWorksheet(SHEET_ARCHIVE, "Archive (Read-Only)")

    sql = "SELECT * FROM dbo.vw_pif_approved_wide"
    If UCase(Trim(selectedSite)) <> "FLEET" Then
        sql = sql & " WHERE UPPER(site) = '" & UCase(selectedSite) & "'"
    End If
    sql = sql & " ORDER BY approval_date DESC, pif_id, project_id"

    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    Call CreateOrRefreshQueryTable(ws, "ArchiveQuery", sql, connStr, selectedSite)

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshInflightWorksheetSilent
' Purpose: Refresh Inflight worksheet without messagebox (for batch operations)
' Notes: Uses Excel QueryTable for native refresh capability
' ----------------------------------------------------------------------------
Private Sub RefreshInflightWorksheetSilent()
    On Error Resume Next

    Dim selectedSite As String
    Dim ws As Worksheet
    Dim sql As String
    Dim connStr As String

    selectedSite = mod_SiteSetup.GetSelectedSite()
    If selectedSite = "" Then Exit Sub

    Set ws = GetOrCreateWorksheet(SHEET_INFLIGHT, "Inflight (Read-Only)")

    sql = "SELECT * FROM dbo.vw_pif_inflight_wide"
    If UCase(Trim(selectedSite)) <> "FLEET" Then
        sql = sql & " WHERE UPPER(site) = '" & UCase(selectedSite) & "'"
    End If
    sql = sql & " ORDER BY submission_date DESC, pif_id, project_id"

    connStr = "OLEDB;Provider=SQLOLEDB;Data Source=" & mod_Database.SQL_SERVER & _
              ";Initial Catalog=" & mod_Database.SQL_DATABASE & _
              ";Integrated Security=SSPI;"

    Call CreateOrRefreshQueryTable(ws, "InflightQuery", sql, connStr, selectedSite)

    On Error GoTo 0
End Sub

' ============================================================================
' BACKWARD COMPATIBILITY WRAPPERS
' ============================================================================

Public Sub RefreshArchiveWorksheet()
    Call Nav_RefreshArchive
End Sub

Public Sub RefreshInflightWorksheet()
    Call Nav_RefreshInflight
End Sub

Public Sub RefreshBothWorksheets()
    Call Nav_RefreshAll
End Sub
