Attribute VB_Name = "mod_WorksheetQuery"
' ============================================================================
' MODULE: mod_WorksheetQuery
' ============================================================================
' Purpose: Create and refresh Archive and Inflight query worksheets
' Author: Data Architecture Team
' Date: 2025-11-11
'
' Features:
'   - Populates PIF_Archive worksheet from vw_approved_wide view
'   - Populates PIF_Inflight worksheet from vw_inflight_wide view
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
' ----------------------------------------------------------------------------
Public Sub RefreshArchiveWorksheet()
    On Error GoTo ErrHandler

    Dim selectedSite As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim sql As String
    Dim startTime As Double
    Dim rowCount As Long

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

    ' Clear existing data
    ws.Cells.Clear

    ' Build SQL query with site filter
    sql = "SELECT * FROM dbo.vw_approved_wide"

    ' Add site filter (Fleet sees all)
    If UCase(selectedSite) <> "FLEET" Then
        sql = sql & " WHERE site = '" & selectedSite & "'"
    End If

    sql = sql & " ORDER BY approval_date DESC, pif_id, project_id"

    ' Execute query
    Set conn = mod_Database.GetDBConnection()
    If conn Is Nothing Then Exit Sub

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly, adLockReadOnly

    ' Write headers
    Call WriteRecordsetHeaders(ws, rs)

    ' Write data
    rowCount = WriteRecordsetData(ws, rs, 2)

    ' Format worksheet
    Call FormatQueryWorksheet(ws, "ARCHIVE - " & selectedSite, rowCount + 1)

    ' Cleanup
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ws.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "Archive worksheet refreshed!" & vbCrLf & vbCrLf & _
           "Site: " & selectedSite & vbCrLf & _
           "Records: " & rowCount & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds", _
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
' ----------------------------------------------------------------------------
Public Sub RefreshInflightWorksheet()
    On Error GoTo ErrHandler

    Dim selectedSite As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim sql As String
    Dim startTime As Double
    Dim rowCount As Long

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

    ' Clear existing data
    ws.Cells.Clear

    ' Build SQL query with site filter
    sql = "SELECT * FROM dbo.vw_inflight_wide"

    ' Add site filter (Fleet sees all)
    If UCase(selectedSite) <> "FLEET" Then
        sql = sql & " WHERE site = '" & selectedSite & "'"
    End If

    sql = sql & " ORDER BY submission_date DESC, pif_id, project_id"

    ' Execute query
    Set conn = mod_Database.GetDBConnection()
    If conn Is Nothing Then Exit Sub

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly, adLockReadOnly

    ' Write headers
    Call WriteRecordsetHeaders(ws, rs)

    ' Write data
    rowCount = WriteRecordsetData(ws, rs, 2)

    ' Format worksheet
    Call FormatQueryWorksheet(ws, "INFLIGHT - " & selectedSite, rowCount + 1)

    ' Cleanup
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ws.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "Inflight worksheet refreshed!" & vbCrLf & vbCrLf & _
           "Site: " & selectedSite & vbCrLf & _
           "Records: " & rowCount & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds", _
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
' Usage: Call after submission/archival to update both views
' ----------------------------------------------------------------------------
Public Sub RefreshBothWorksheets()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    ' Refresh both without individual messageboxes
    Call RefreshArchiveWorksheetSilent
    Call RefreshInflightWorksheetSilent

    Application.ScreenUpdating = True

    MsgBox "Both Archive and Inflight worksheets have been refreshed.", _
           vbInformation, "Refresh Complete"

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

    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i

    ' Format header row
    With ws.Rows(1)
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

    rowNum = startRow
    recordCount = 0

    Do While Not rs.EOF
        For colNum = 0 To rs.Fields.Count - 1
            If Not IsNull(rs.Fields(colNum).Value) Then
                ws.Cells(rowNum, colNum + 1).Value = rs.Fields(colNum).Value
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
    ' Auto-fit columns
    ws.Cells.Columns.AutoFit

    ' Freeze top row
    ws.Activate
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    ' Add auto-filter
    If lastRow > 1 Then
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)).AutoFilter
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
' ----------------------------------------------------------------------------
Private Sub RefreshArchiveWorksheetSilent()
    On Error Resume Next

    Dim selectedSite As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim sql As String
    Dim rowCount As Long

    selectedSite = mod_SiteSetup.GetSelectedSite()
    If selectedSite = "" Then Exit Sub

    Set ws = GetOrCreateWorksheet(SHEET_ARCHIVE, "Archive (Read-Only)")
    ws.Cells.Clear

    sql = "SELECT * FROM dbo.vw_approved_wide"
    If UCase(selectedSite) <> "FLEET" Then
        sql = sql & " WHERE site = '" & selectedSite & "'"
    End If
    sql = sql & " ORDER BY approval_date DESC, pif_id, project_id"

    Set conn = mod_Database.GetDBConnection()
    If conn Is Nothing Then Exit Sub

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly, adLockReadOnly

    Call WriteRecordsetHeaders(ws, rs)
    rowCount = WriteRecordsetData(ws, rs, 2)
    Call FormatQueryWorksheet(ws, "ARCHIVE - " & selectedSite, rowCount + 1)

    rs.Close
    conn.Close

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: RefreshInflightWorksheetSilent
' Purpose: Refresh Inflight worksheet without messagebox (for batch operations)
' ----------------------------------------------------------------------------
Private Sub RefreshInflightWorksheetSilent()
    On Error Resume Next

    Dim selectedSite As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim sql As String
    Dim rowCount As Long

    selectedSite = mod_SiteSetup.GetSelectedSite()
    If selectedSite = "" Then Exit Sub

    Set ws = GetOrCreateWorksheet(SHEET_INFLIGHT, "Inflight (Read-Only)")
    ws.Cells.Clear

    sql = "SELECT * FROM dbo.vw_inflight_wide"
    If UCase(selectedSite) <> "FLEET" Then
        sql = sql & " WHERE site = '" & selectedSite & "'"
    End If
    sql = sql & " ORDER BY submission_date DESC, pif_id, project_id"

    Set conn = mod_Database.GetDBConnection()
    If conn Is Nothing Then Exit Sub

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly, adLockReadOnly

    Call WriteRecordsetHeaders(ws, rs)
    rowCount = WriteRecordsetData(ws, rs, 2)
    Call FormatQueryWorksheet(ws, "INFLIGHT - " & selectedSite, rowCount + 1)

    rs.Close
    conn.Close

    On Error GoTo 0
End Sub
