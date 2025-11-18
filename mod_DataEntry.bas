<<<<<<< HEAD
Attribute VB_Name = "mod_DataEntry"
' ============================================================================
' MODULE: mod_DataEntry
' ============================================================================
' Purpose: Helper functions for data entry on PIF worksheet
' Author: Data Architecture Team
' Date: 2025-11-11
'
' Features:
'   - Add Row button: Inserts new data row with formulas/formatting
'   - Delete Row button: Safely removes selected rows
'   - Data entry validation helpers
' ============================================================================

Option Explicit

' Sheet name constant
Private Const SHEET_DATA As String = "PIF"

' ============================================================================
' PUBLIC FUNCTIONS - USER INTERFACE BUTTONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: AddRow
' Purpose: Insert a new blank data row above the totals row
' Usage: Attach to [Add Row] button on PIF sheet
' Notes: Automatically copies formulas and formatting from previous row
' ----------------------------------------------------------------------------
Public Sub Edit_AddRow()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lastDataRow As Long
    Dim newRow As Long
    Dim sourceRow As Long

    Application.ScreenUpdating = False

    ' Get PIF worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    ' Find last row with data (look for PIF_ID in column G)
    lastDataRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row  ' Column G = PIF_ID

    ' Check if we're on a valid data area (should be row 4 or higher)
    If lastDataRow < 4 Then
        lastDataRow = 3  ' Start after header rows
    End If

    ' New row will be inserted after last data row
    newRow = lastDataRow + 1
    sourceRow = lastDataRow

    ' If source row appears to be a totals/summary row (check for SUM/SUBTOTAL formulas),
    ' use the row above it as the source
    Dim cellFormula As String
    Dim isTotalRow As Boolean
    On Error Resume Next
    cellFormula = UCase(ws.Cells(sourceRow, 21).Formula)  ' Column U (first cost column)
    On Error GoTo ErrHandler

    ' Check for total row indicators: SUM, SUBTOTAL, or AGGREGATE functions
    isTotalRow = (InStr(1, cellFormula, "SUM(") > 0) Or _
                 (InStr(1, cellFormula, "SUBTOTAL(") > 0) Or _
                 (InStr(1, cellFormula, "AGGREGATE(") > 0)

    ' Also check if PIF_ID column is empty (another indicator of total row)
    If Trim(CStr(ws.Cells(sourceRow, 7).Value)) = "" Then
        isTotalRow = True
    End If

    If isTotalRow Then
        ' This is a totals row - insert ABOVE it and copy from row above
        newRow = sourceRow
        sourceRow = sourceRow - 1

        ' Make sure source row is valid
        If sourceRow < 4 Then
            MsgBox "Cannot add row - no valid data rows found to copy formatting from.", _
                   vbExclamation, "Add Row Error"
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If

    ' Insert new row
    ws.Rows(newRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Copy formulas and formatting from source row (but not values)
    ' Copy range: Columns A through BF (1 through 58)
    Dim sourceRange As Range
    Dim targetRange As Range

    Set sourceRange = ws.Range(ws.Cells(sourceRow, 1), ws.Cells(sourceRow, 58))
    Set targetRange = ws.Range(ws.Cells(newRow, 1), ws.Cells(newRow, 58))

    ' Copy formulas (preserves formula structure)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False

    ' Copy formatting (preserves colors, borders, fonts)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Clear data from input columns (leave formulas intact)
    ' Clear columns C-T (3-20) - data entry columns
    ' Note: We don't clear calculated columns like variance columns
    Dim clearColumns As Variant
    clearColumns = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)

    Dim col As Variant
    For Each col In clearColumns
        ' Only clear if cell doesn't have a formula
        If Not ws.Cells(newRow, col).HasFormula Then
            ws.Cells(newRow, col).ClearContents
        End If
    Next col

    ' Clear Target Requested columns (U-Z, columns 21-26)
    Dim i As Integer
    For i = 21 To 26
        If Not ws.Cells(newRow, i).HasFormula Then
            ws.Cells(newRow, i).ClearContents
        End If
    Next i

    ' Clear Closings Requested columns (AO-AT, columns 41-46)
    For i = 41 To 46
        If Not ws.Cells(newRow, i).HasFormula Then
            ws.Cells(newRow, i).ClearContents
        End If
    Next i

    ' Clear Moving ISD Year (AM, column 39)
    If Not ws.Cells(newRow, 39).HasFormula Then
        ws.Cells(newRow, 39).ClearContents
    End If

    ' Clear Prior Year Spend (AN, column 40)
    If Not ws.Cells(newRow, 40).HasFormula Then
        ws.Cells(newRow, 40).ClearContents
    End If

    ' Set default values for checkboxes (Archive and Include in columns C and D)
    ws.Cells(newRow, 3).Value = False  ' Archive checkbox
    ws.Cells(newRow, 4).Value = False  ' Include checkbox

    ' Auto-populate site from Instructions sheet if available
    On Error Resume Next
    Dim selectedSite As String
    selectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.Value)
    If selectedSite <> "" And UCase(selectedSite) <> "FLEET" Then
        ws.Cells(newRow, 10).Value = selectedSite  ' Column J = Site
    End If
    On Error GoTo ErrHandler

    ' Select the PIF ID cell (Column G) for user to start entering data
    ws.Cells(newRow, 7).Select

    Application.ScreenUpdating = True

    ' Success message (optional - remove if too intrusive)
    ' MsgBox "New row added at row " & newRow & vbCrLf & vbCrLf & _
    '        "Enter PIF ID to begin.", vbInformation, "Row Added"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error adding row:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Add Row Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: DeleteSelectedRows
' Purpose: Delete selected row(s) from PIF sheet with confirmation
' Usage: Attach to [Delete Row] button or context menu
' Notes: Prevents accidental deletion with confirmation prompt
'        Handles both Excel Tables (ListObjects) and regular ranges
' ----------------------------------------------------------------------------
Public Sub Edit_DeleteRows()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim selectedRows As Range
    Dim response As VbMsgBoxResult
    Dim rowCount As Long
    Dim tbl As ListObject
    Dim firstRow As Long
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    ' Get selected rows
    On Error Resume Next
    Set selectedRows = Intersect(Selection.EntireRow, ws.UsedRange)
    On Error GoTo ErrHandler

    If selectedRows Is Nothing Then
        MsgBox "Please select one or more rows to delete.", _
               vbExclamation, "No Rows Selected"
        Exit Sub
    End If

    ' Count rows
    rowCount = selectedRows.Rows.Count
    firstRow = selectedRows.Row
    lastRow = firstRow + rowCount - 1

    ' Prevent deletion of header rows (rows 1-3)
    If firstRow <= 3 Then
        MsgBox "Cannot delete header rows (rows 1-3).", _
               vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    ' Confirmation prompt
    If rowCount = 1 Then
        response = MsgBox("Delete selected row?", _
                         vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete")
    Else
        response = MsgBox("Delete " & rowCount & " selected rows?", _
                         vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete")
    End If

    If response = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Check if selection is within an Excel Table (ListObject)
    On Error Resume Next
    Set tbl = selectedRows.Cells(1, 1).ListObject
    On Error GoTo ErrHandler

    If Not tbl Is Nothing Then
        ' ========================================================================
        ' EXCEL TABLE DELETION (ListObject)
        ' ========================================================================
        ' When working with Excel Tables, we must use the ListRows collection
        ' to delete rows. Standard row deletion causes Error 1004.

        ' Handle filtered tables: temporarily disable AutoFilter to allow deletion
        Dim filterWasOn As Boolean
        filterWasOn = False

        If tbl.ShowAutoFilter Then
            If tbl.AutoFilter.FilterMode Then
                ' Filters are active - need to clear them temporarily
                filterWasOn = True
                tbl.AutoFilter.ShowAllData  ' Clear filters
            End If
        End If

        ' Build array of ListRow indices to delete
        Dim indicesToDelete() As Long
        Dim indexCount As Long
        Dim lr As ListRow
        Dim j As Long
        Dim temp As Long

        indexCount = 0

        ' First pass: count how many rows to delete
        For Each lr In tbl.ListRows
            If Not Intersect(lr.Range, selectedRows) Is Nothing Then
                indexCount = indexCount + 1
            End If
        Next lr

        ' If no rows to delete, exit
        If indexCount = 0 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If

        ' Size the array
        ReDim indicesToDelete(1 To indexCount)

        ' Second pass: collect indices
        j = 1
        For Each lr In tbl.ListRows
            If Not Intersect(lr.Range, selectedRows) Is Nothing Then
                indicesToDelete(j) = lr.Index
                j = j + 1
            End If
        Next lr

        ' Sort indices in descending order (bubble sort - simple and reliable)
        For i = 1 To indexCount - 1
            For j = i + 1 To indexCount
                If indicesToDelete(i) < indicesToDelete(j) Then
                    temp = indicesToDelete(i)
                    indicesToDelete(i) = indicesToDelete(j)
                    indicesToDelete(j) = temp
                End If
            Next j
        Next i

        ' Delete rows from highest index to lowest to maintain correct indices
        For i = 1 To indexCount
            tbl.ListRows(indicesToDelete(i)).Delete
        Next i

        ' Note: AutoFilter is automatically reapplied by Excel after row deletion
        ' No need to manually re-enable it

    Else
        ' ========================================================================
        ' REGULAR RANGE DELETION (Non-Table)
        ' ========================================================================
        ' Standard row deletion for non-table ranges
        selectedRows.Delete Shift:=xlUp
    End If

    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error deleting rows:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Delete Error"
End Sub

' ----------------------------------------------------------------------------
' Function: ValidateDataRow
' Purpose: Quick validation check for a single row
' Parameters:
'   ws - Worksheet containing data
'   rowNum - Row number to validate
' Returns: True if row has minimum required fields, False otherwise
' Usage: Can be called before submission to check individual rows
' ----------------------------------------------------------------------------
Public Function ValidateDataRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrHandler

    Dim pifId As String
    Dim projectId As String
    Dim changeType As String
    Dim site As String

    ' Get required fields
    pifId = Trim(CStr(ws.Cells(rowNum, 7).Value))          ' Column G
    projectId = Trim(CStr(ws.Cells(rowNum, 13).Value))     ' Column M
    changeType = Trim(CStr(ws.Cells(rowNum, 6).Value))     ' Column F
    site = Trim(CStr(ws.Cells(rowNum, 10).Value))          ' Column J

    ' Check if all required fields are present
    If pifId = "" Or projectId = "" Or changeType = "" Or site = "" Then
        ValidateDataRow = False
    Else
        ValidateDataRow = True
    End If

    Exit Function

ErrHandler:
    ValidateDataRow = False
End Function

' ----------------------------------------------------------------------------
' Sub: HighlightIncompleteRows
' Purpose: Highlight rows missing required fields (visual aid)
' Usage: Run manually or attach to button for data quality check
' Notes: Highlights entire row in light red if missing required fields
' ----------------------------------------------------------------------------
Public Sub Tool_HighlightIncomplete()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim highlightCount As Long

    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row

    highlightCount = 0

    ' Loop through data rows (start at row 4, rows 1-3 are headers)
    For i = 4 To lastRow
        ' Skip completely empty rows
        If WorksheetFunction.CountA(ws.Rows(i)) = 0 Then GoTo NextRow

        ' Skip rows without PIF_ID (likely summary/total rows)
        If Trim(CStr(ws.Cells(i, 7).Value)) = "" Then GoTo NextRow

        ' Check if row is valid
        If Not ValidateDataRow(ws, i) Then
            ' Highlight row in light red
            ws.Rows(i).Interior.Color = RGB(255, 200, 200)
            highlightCount = highlightCount + 1
        Else
            ' Remove highlighting if row is now complete
            ws.Rows(i).Interior.ColorIndex = xlNone
        End If

NextRow:
    Next i

    Application.ScreenUpdating = True

    If highlightCount > 0 Then
        MsgBox highlightCount & " incomplete row(s) highlighted in light red." & vbCrLf & vbCrLf & _
               "Required fields: PIF ID, Project #, Change Type, Site", _
               vbInformation, "Validation Complete"
    Else
        MsgBox "All data rows have required fields!", _
               vbInformation, "Validation Complete"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error highlighting rows:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Highlight Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: ClearRowHighlights
' Purpose: Remove all row highlighting applied by HighlightIncompleteRows
' Usage: Run manually or attach to button
' ----------------------------------------------------------------------------
Public Sub Tool_ClearHighlights()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lastRow As Long

    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Sheets(SHEET_DATA)
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row

    ' Clear highlighting from all data rows (start at row 4)
    If lastRow >= 4 Then
        ws.Range(ws.Rows(4), ws.Rows(lastRow)).Interior.ColorIndex = xlNone
    End If

    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error clearing highlights:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Clear Error"
End Sub

' ============================================================================
' BACKWARD COMPATIBILITY WRAPPERS
' ============================================================================

Public Sub AddRow()
    Call Edit_AddRow
End Sub

Public Sub DeleteSelectedRows()
    Call Edit_DeleteRows
End Sub

Public Sub HighlightIncompleteRows()
    Call Tool_HighlightIncomplete
End Sub

Public Sub ClearRowHighlights()
    Call Tool_ClearHighlights
End Sub
=======
Attribute VB_Name = "mod_DataEntry"
' ============================================================================
' MODULE: mod_DataEntry
' ============================================================================
' Purpose: Helper functions for data entry on PIF worksheet
' Author: Data Architecture Team
' Date: 2025-11-11
'
' Features:
'   - Add Row button: Inserts new data row with formulas/formatting
'   - Delete Row button: Safely removes selected rows
'   - Data entry validation helpers
' ============================================================================

Option Explicit

' Sheet name constant
Private Const SHEET_DATA As String = "PIF"

' ============================================================================
' PUBLIC FUNCTIONS - USER INTERFACE BUTTONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: AddRow
' Purpose: Insert a new blank data row above the totals row
' Usage: Attach to [Add Row] button on PIF sheet
' Notes: Automatically copies formulas and formatting from previous row
' ----------------------------------------------------------------------------
Public Sub Edit_AddRow()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lastDataRow As Long
    Dim newRow As Long
    Dim sourceRow As Long

    Application.ScreenUpdating = False

    ' Get PIF worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    ' Find last row with data (look for PIF_ID in column G)
    lastDataRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row  ' Column G = PIF_ID

    ' Check if we're on a valid data area (should be row 4 or higher)
    If lastDataRow < 4 Then
        lastDataRow = 3  ' Start after header rows
    End If

    ' New row will be inserted after last data row
    newRow = lastDataRow + 1
    sourceRow = lastDataRow

    ' If source row appears to be a totals/summary row (check for SUM/SUBTOTAL formulas),
    ' use the row above it as the source
    Dim cellFormula As String
    Dim isTotalRow As Boolean
    On Error Resume Next
    cellFormula = UCase(ws.Cells(sourceRow, 21).Formula)  ' Column U (first cost column)
    On Error GoTo ErrHandler

    ' Check for total row indicators: SUM, SUBTOTAL, or AGGREGATE functions
    isTotalRow = (InStr(1, cellFormula, "SUM(") > 0) Or _
                 (InStr(1, cellFormula, "SUBTOTAL(") > 0) Or _
                 (InStr(1, cellFormula, "AGGREGATE(") > 0)

    ' Also check if PIF_ID column is empty (another indicator of total row)
    If Trim(CStr(ws.Cells(sourceRow, 7).Value)) = "" Then
        isTotalRow = True
    End If

    If isTotalRow Then
        ' This is a totals row - insert ABOVE it and copy from row above
        newRow = sourceRow
        sourceRow = sourceRow - 1

        ' Make sure source row is valid
        If sourceRow < 4 Then
            MsgBox "Cannot add row - no valid data rows found to copy formatting from.", _
                   vbExclamation, "Add Row Error"
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If

    ' Insert new row
    ws.Rows(newRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Copy formulas and formatting from source row (but not values)
    ' Copy range: Columns A through BF (1 through 58)
    Dim sourceRange As Range
    Dim targetRange As Range

    Set sourceRange = ws.Range(ws.Cells(sourceRow, 1), ws.Cells(sourceRow, 58))
    Set targetRange = ws.Range(ws.Cells(newRow, 1), ws.Cells(newRow, 58))

    ' Copy formulas (preserves formula structure)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False

    ' Copy formatting (preserves colors, borders, fonts)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Clear data from input columns (leave formulas intact)
    ' Clear columns C-T (3-20) - data entry columns
    ' Note: We don't clear calculated columns like variance columns
    Dim clearColumns As Variant
    clearColumns = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)

    Dim col As Variant
    For Each col In clearColumns
        ' Only clear if cell doesn't have a formula
        If Not ws.Cells(newRow, col).HasFormula Then
            ws.Cells(newRow, col).ClearContents
        End If
    Next col

    ' Clear Target Requested columns (U-Z, columns 21-26)
    Dim i As Integer
    For i = 21 To 26
        If Not ws.Cells(newRow, i).HasFormula Then
            ws.Cells(newRow, i).ClearContents
        End If
    Next i

    ' Clear Closings Requested columns (AO-AT, columns 41-46)
    For i = 41 To 46
        If Not ws.Cells(newRow, i).HasFormula Then
            ws.Cells(newRow, i).ClearContents
        End If
    Next i

    ' Clear Moving ISD Year (AM, column 39)
    If Not ws.Cells(newRow, 39).HasFormula Then
        ws.Cells(newRow, 39).ClearContents
    End If

    ' Clear Prior Year Spend (AN, column 40)
    If Not ws.Cells(newRow, 40).HasFormula Then
        ws.Cells(newRow, 40).ClearContents
    End If

    ' Set default values for checkboxes (Archive and Include in columns C and D)
    ws.Cells(newRow, 3).Value = False  ' Archive checkbox
    ws.Cells(newRow, 4).Value = False  ' Include checkbox

    ' Auto-populate site from Instructions sheet if available
    On Error Resume Next
    Dim selectedSite As String
    selectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.Value)
    If selectedSite <> "" And UCase(selectedSite) <> "FLEET" Then
        ws.Cells(newRow, 10).Value = selectedSite  ' Column J = Site
    End If
    On Error GoTo ErrHandler

    ' Select the PIF ID cell (Column G) for user to start entering data
    ws.Cells(newRow, 7).Select

    Application.ScreenUpdating = True

    ' Success message (optional - remove if too intrusive)
    ' MsgBox "New row added at row " & newRow & vbCrLf & vbCrLf & _
    '        "Enter PIF ID to begin.", vbInformation, "Row Added"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error adding row:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Add Row Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: DeleteSelectedRows
' Purpose: Delete selected row(s) from PIF sheet with confirmation
' Usage: Attach to [Delete Row] button or context menu
' Notes: Prevents accidental deletion with confirmation prompt
'        Handles both Excel Tables (ListObjects) and regular ranges
' ----------------------------------------------------------------------------
Public Sub Edit_DeleteRows()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim selectedRows As Range
    Dim response As VbMsgBoxResult
    Dim rowCount As Long
    Dim tbl As ListObject
    Dim firstRow As Long
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    ' Get selected rows
    On Error Resume Next
    Set selectedRows = Intersect(Selection.EntireRow, ws.UsedRange)
    On Error GoTo ErrHandler

    If selectedRows Is Nothing Then
        MsgBox "Please select one or more rows to delete.", _
               vbExclamation, "No Rows Selected"
        Exit Sub
    End If

    ' Count rows
    rowCount = selectedRows.Rows.Count
    firstRow = selectedRows.Row
    lastRow = firstRow + rowCount - 1

    ' Prevent deletion of header rows (rows 1-3)
    If firstRow <= 3 Then
        MsgBox "Cannot delete header rows (rows 1-3).", _
               vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    ' Confirmation prompt
    If rowCount = 1 Then
        response = MsgBox("Delete selected row?", _
                         vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete")
    Else
        response = MsgBox("Delete " & rowCount & " selected rows?", _
                         vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete")
    End If

    If response = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Check if selection is within an Excel Table (ListObject)
    On Error Resume Next
    Set tbl = selectedRows.Cells(1, 1).ListObject
    On Error GoTo ErrHandler

    If Not tbl Is Nothing Then
        ' ========================================================================
        ' EXCEL TABLE DELETION (ListObject)
        ' ========================================================================
        ' When working with Excel Tables, we must use the ListRows collection
        ' to delete rows. Standard row deletion causes Error 1004.

        ' Handle filtered tables: temporarily disable AutoFilter to allow deletion
        Dim filterWasOn As Boolean
        filterWasOn = False

        If tbl.ShowAutoFilter Then
            If tbl.AutoFilter.FilterMode Then
                ' Filters are active - need to clear them temporarily
                filterWasOn = True
                tbl.AutoFilter.ShowAllData  ' Clear filters
            End If
        End If

        ' Build array of ListRow indices to delete
        Dim indicesToDelete() As Long
        Dim indexCount As Long
        Dim lr As ListRow
        Dim j As Long
        Dim temp As Long

        indexCount = 0

        ' First pass: count how many rows to delete
        For Each lr In tbl.ListRows
            If Not Intersect(lr.Range, selectedRows) Is Nothing Then
                indexCount = indexCount + 1
            End If
        Next lr

        ' If no rows to delete, exit
        If indexCount = 0 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If

        ' Size the array
        ReDim indicesToDelete(1 To indexCount)

        ' Second pass: collect indices
        j = 1
        For Each lr In tbl.ListRows
            If Not Intersect(lr.Range, selectedRows) Is Nothing Then
                indicesToDelete(j) = lr.Index
                j = j + 1
            End If
        Next lr

        ' Sort indices in descending order (bubble sort - simple and reliable)
        For i = 1 To indexCount - 1
            For j = i + 1 To indexCount
                If indicesToDelete(i) < indicesToDelete(j) Then
                    temp = indicesToDelete(i)
                    indicesToDelete(i) = indicesToDelete(j)
                    indicesToDelete(j) = temp
                End If
            Next j
        Next i

        ' Delete rows from highest index to lowest to maintain correct indices
        For i = 1 To indexCount
            tbl.ListRows(indicesToDelete(i)).Delete
        Next i

        ' Note: AutoFilter is automatically reapplied by Excel after row deletion
        ' No need to manually re-enable it

    Else
        ' ========================================================================
        ' REGULAR RANGE DELETION (Non-Table)
        ' ========================================================================
        ' Standard row deletion for non-table ranges
        selectedRows.Delete Shift:=xlUp
    End If

    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error deleting rows:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Delete Error"
End Sub

' ----------------------------------------------------------------------------
' Function: ValidateDataRow
' Purpose: Quick validation check for a single row
' Parameters:
'   ws - Worksheet containing data
'   rowNum - Row number to validate
' Returns: True if row has minimum required fields, False otherwise
' Usage: Can be called before submission to check individual rows
' ----------------------------------------------------------------------------
Public Function ValidateDataRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrHandler

    Dim pifId As String
    Dim projectId As String
    Dim changeType As String
    Dim site As String

    ' Get required fields
    pifId = Trim(CStr(ws.Cells(rowNum, 7).Value))          ' Column G
    projectId = Trim(CStr(ws.Cells(rowNum, 13).Value))     ' Column M
    changeType = Trim(CStr(ws.Cells(rowNum, 6).Value))     ' Column F
    site = Trim(CStr(ws.Cells(rowNum, 10).Value))          ' Column J

    ' Check if all required fields are present
    If pifId = "" Or projectId = "" Or changeType = "" Or site = "" Then
        ValidateDataRow = False
    Else
        ValidateDataRow = True
    End If

    Exit Function

ErrHandler:
    ValidateDataRow = False
End Function

' ----------------------------------------------------------------------------
' Sub: HighlightIncompleteRows
' Purpose: Highlight rows missing required fields (visual aid)
' Usage: Run manually or attach to button for data quality check
' Notes: Highlights entire row in light red if missing required fields
' ----------------------------------------------------------------------------
Public Sub Tool_HighlightIncomplete()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim highlightCount As Long

    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row

    highlightCount = 0

    ' Loop through data rows (start at row 4, rows 1-3 are headers)
    For i = 4 To lastRow
        ' Skip completely empty rows
        If WorksheetFunction.CountA(ws.Rows(i)) = 0 Then GoTo NextRow

        ' Skip rows without PIF_ID (likely summary/total rows)
        If Trim(CStr(ws.Cells(i, 7).Value)) = "" Then GoTo NextRow

        ' Check if row is valid
        If Not ValidateDataRow(ws, i) Then
            ' Highlight row in light red
            ws.Rows(i).Interior.Color = RGB(255, 200, 200)
            highlightCount = highlightCount + 1
        Else
            ' Remove highlighting if row is now complete
            ws.Rows(i).Interior.ColorIndex = xlNone
        End If

NextRow:
    Next i

    Application.ScreenUpdating = True

    If highlightCount > 0 Then
        MsgBox highlightCount & " incomplete row(s) highlighted in light red." & vbCrLf & vbCrLf & _
               "Required fields: PIF ID, Project #, Change Type, Site", _
               vbInformation, "Validation Complete"
    Else
        MsgBox "All data rows have required fields!", _
               vbInformation, "Validation Complete"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error highlighting rows:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Highlight Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: ClearRowHighlights
' Purpose: Remove all row highlighting applied by HighlightIncompleteRows
' Usage: Run manually or attach to button
' ----------------------------------------------------------------------------
Public Sub Tool_ClearHighlights()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lastRow As Long

    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Sheets(SHEET_DATA)
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row

    ' Clear highlighting from all data rows (start at row 4)
    If lastRow >= 4 Then
        ws.Range(ws.Rows(4), ws.Rows(lastRow)).Interior.ColorIndex = xlNone
    End If

    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error clearing highlights:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Clear Error"
End Sub

' ============================================================================
' BACKWARD COMPATIBILITY WRAPPERS
' ============================================================================

Public Sub AddRow()
    Call Edit_AddRow
End Sub

Public Sub DeleteSelectedRows()
    Call Edit_DeleteRows
End Sub

Public Sub HighlightIncompleteRows()
    Call Tool_HighlightIncomplete
End Sub

Public Sub ClearRowHighlights()
    Call Tool_ClearHighlights
End Sub
>>>>>>> 254c5f99beb56416a0a77e8c6100e047b5b2e663
