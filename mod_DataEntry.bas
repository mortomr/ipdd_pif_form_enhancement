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
Private Const SHEET_DATA As String = "Target Adjustment"

' ============================================================================
' PUBLIC FUNCTIONS - USER INTERFACE BUTTONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: AddRow
' Purpose: Insert a new blank data row above the totals row
' Usage: Attach to [Add Row] button on PIF sheet
' Notes: Automatically copies formulas and formatting from previous row
' Rev to accomodate New Line # Column @ "G"
' ----------------------------------------------------------------------------
Public Sub Edit_AddRow()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastDataRow As Long
    Dim newRow As Long
    Dim sourceRow As Long

    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Sheets(SHEET_DATA)
    
    ' Get the table object
    Set tbl = ws.ListObjects(1)  ' Assumes first table on sheet
    
    ' Find last row with actual data (column G = Line Item)
    lastDataRow = ws.Cells(ws.Rows.count, 7).End(xlUp).row
    If lastDataRow < 4 Then lastDataRow = 3
    
    newRow = lastDataRow + 1
    sourceRow = lastDataRow

    ' Step 1: Enter dummy value in first column to trigger table auto-expand
    ws.Range("C" & newRow).value = ""  ' Archive column
    
    ' Step 2: Extend table range to include new row if not auto-expanded yet
    If newRow > tbl.Range.Rows.count + tbl.Range.row - 1 Then
        tbl.Resize tbl.Range.Offset(0, 0).Resize(newRow - tbl.Range.row + 1)
    End If

    ' Step 3: Copy formulas from source row
    ws.Range("O" & sourceRow & ":BH" & sourceRow).Copy
    ws.Range("O" & newRow).PasteSpecial xlPasteFormulas
'    ws.Range("O" & sourceRow).Copy
'    ws.Range("O" & newRow).PasteSpecial xlPasteFormulas
    
    ' Step 4: Copy/set static columns
    ws.Range("C" & newRow).value = ws.Range("C" & sourceRow).value  ' Archive
    ws.Range("D" & newRow).value = ws.Range("D" & sourceRow).value  ' Include
    ws.Range("E" & newRow).value = ws.Range("E" & sourceRow).value  ' Accounting Treatment
    ws.Range("F" & newRow).value = ws.Range("F" & sourceRow).value  ' Change Type
    ws.Range("G" & newRow).Formula = "=ROW()-3"  ' Line Item
    
    ' Step 5: Clear user-entry columns
    ws.Range("H" & newRow & ":N" & newRow).ClearContents
    ' Skip around "O"
    ws.Range("P" & newRow & ":V" & newRow).ClearContents

    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    MsgBox "Row " & newRow - 3 & " added. Fill in PIF ID and required fields.", vbInformation

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
    rowCount = selectedRows.Rows.count
    firstRow = selectedRows.row
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
    pifId = Trim(CStr(ws.Cells(rowNum, 7).value))          ' Column G
    projectId = Trim(CStr(ws.Cells(rowNum, 13).value))     ' Column M
    changeType = Trim(CStr(ws.Cells(rowNum, 6).value))     ' Column F
    site = Trim(CStr(ws.Cells(rowNum, 10).value))          ' Column J

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
    lastRow = ws.Cells(ws.Rows.count, 7).End(xlUp).row

    highlightCount = 0

    ' Loop through data rows (start at row 4, rows 1-3 are headers)
    For i = 4 To lastRow
        ' Skip completely empty rows
        If WorksheetFunction.CountA(ws.Rows(i)) = 0 Then GoTo NextRow

        ' Skip rows without PIF_ID (likely summary/total rows)
        If Trim(CStr(ws.Cells(i, 7).value)) = "" Then GoTo NextRow

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
    lastRow = ws.Cells(ws.Rows.count, 7).End(xlUp).row

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
