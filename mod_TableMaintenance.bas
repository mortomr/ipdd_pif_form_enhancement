Attribute VB_Name = "mod_TableMaintenance"
' ============================================================================
' MODULE: mod_TableMaintenance
' ============================================================================
' Purpose:  (1) Re-initialize entry table columns - restores data validation,
'               formulas, and formatting after record churn degrades them.
'           (2) Auto-populate OpCo (col J) when Site (col K) is changed.
'               Seg (col I) is formula-driven from FP lookup - not hardcoded.
'
' Public Entry Points:
'   ReinitializeEntryTable   - call from button or Workbook_Open
'   ApplySiteDefaults        - call from Worksheet_Change on col K
'
' Sheet event wiring (paste into "Target Adjustment" sheet module):
' -----------------------------------------------------------------------
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Target.Column = 11 And Target.Row >= 4 Then  ' Col K = Site
'           mod_TableMaintenance.ApplySiteDefaults Target.Row
'       End If
'   End Sub
' -----------------------------------------------------------------------
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' CONSTANTS
' ----------------------------------------------------------------------------
Private Const SHEET_DATA        As String = "Target Adjustment"
Private Const FIRST_DATA_ROW    As Long = 4

' Column numbers (1-based)
Private Const COL_LINE_ITEM     As Long = 7   ' G  - formula =ROW()-3
Private Const COL_SEG           As Long = 9   ' I  - formula from LCM_Segment lookup
Private Const COL_OPCO          As Long = 10  ' J  - auto-populated from site
Private Const COL_SITE          As Long = 11  ' K  - dropdown
Private Const COL_ACCT          As Long = 5   ' E  - dropdown
Private Const COL_CHANGE_TYPE   As Long = 6   ' F  - dropdown
Private Const COL_STRAT_RANK    As Long = 12  ' L  - dropdown
Private Const COL_REDUCTION     As Long = 19  ' S  - dropdown
Private Const COL_RISK          As Long = 20  ' T  - dropdown {High,Med,Low}
Private Const COL_STATUS        As Long = 21  ' U  - dropdown
Private Const COL_PROJ_NAME     As Long = 15  ' O  - XLOOKUP formula
Private Const COL_CURRENT_ISD   As Long = 16  ' P  - Current PowerPlan ISD

' Cost column bands
Private Const COL_TARGET_REQ_START  As Long = 23  ' W
Private Const COL_TARGET_CUR_START  As Long = 29  ' AC
Private Const COL_TARGET_VAR_START  As Long = 35  ' AI
Private Const COL_PRIOR_SPEND       As Long = 42  ' AP
Private Const COL_CLOS_REQ_START    As Long = 43  ' AQ
Private Const COL_CLOS_CUR_START    As Long = 49  ' AW
Private Const COL_CLOS_VAR_START    As Long = 55  ' BC

' Dropdown formulas - must use INDIRECT; Excel rejects table refs directly in DV
Private Const DV_ACCOUNTING     As String = "=INDIRECT(""tbl_dropdown_accounting[opt]"")"
Private Const DV_CHANGE_TYPE    As String = "=INDIRECT(""tbl_dropdown_changetype[opt]"")"
Private Const DV_SITE           As String = "=INDIRECT(""tbl_dropdown_site[opt]"")"
Private Const DV_STRAT_RANK     As String = "=INDIRECT(""tbl_dropdown_strat_rank[opt]"")"
Private Const DV_REDUCTION      As String = "=INDIRECT(""tbl_dropdown_binning[opt]"")"
Private Const DV_STATUS         As String = "=INDIRECT(""tbl_dropdown_reduction_type[opt]"")"
Private Const DV_RISK           As String = "High,Med,Low"

' ============================================================================
' PUBLIC SUBS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ReinitializeEntryTable
' Purpose: Restore all dropdowns and formulas on every data row.
'          Safe to run repeatedly - idempotent.
' ----------------------------------------------------------------------------
Public Sub ReinitializeEntryTable()
    On Error GoTo ErrHandler

    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim r       As Long
    Dim i       As Integer
    Dim rngData As Range

    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    lastRow = ws.Cells(ws.Rows.count, 14).End(xlUp).Row  ' anchor on col N (FP)

    If lastRow < FIRST_DATA_ROW Then
        MsgBox "No data rows found - nothing to reinitialize.", vbInformation, "Table Maintenance"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Reinitializing entry table..."

    ' -----------------------------------------------------------------------
    ' PASS 1: Dropdowns - apply to full column range at once
    ' -----------------------------------------------------------------------
    Set rngData = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_ACCT), ws.Cells(lastRow, COL_ACCT))
    ApplyListValidation rngData, DV_ACCOUNTING

    Set rngData = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_CHANGE_TYPE), ws.Cells(lastRow, COL_CHANGE_TYPE))
    ApplyListValidation rngData, DV_CHANGE_TYPE

    Set rngData = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_SITE), ws.Cells(lastRow, COL_SITE))
    ApplyListValidation rngData, DV_SITE

    Set rngData = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_STRAT_RANK), ws.Cells(lastRow, COL_STRAT_RANK))
    ApplyListValidation rngData, DV_STRAT_RANK

    Set rngData = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_REDUCTION), ws.Cells(lastRow, COL_REDUCTION))
    ApplyListValidation rngData, DV_REDUCTION

    Set rngData = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_STATUS), ws.Cells(lastRow, COL_STATUS))
    ApplyListValidation rngData, DV_STATUS

    Set rngData = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_RISK), ws.Cells(lastRow, COL_RISK))
    ApplyInlineListValidation rngData, DV_RISK

    ' -----------------------------------------------------------------------
    ' PASS 2: Row-by-row formulas
    ' -----------------------------------------------------------------------
    For r = FIRST_DATA_ROW To lastRow

        ' G - Line item sequence
        ws.Cells(r, COL_LINE_ITEM).formula = "=ROW()-3"

        ' I - Segment: extract single digit from LCM_Segment (e.g. "NUC-1" -> "1")
        ws.Cells(r, COL_SEG).formula = _
            "=IFERROR(MID(XLOOKUP($N" & r & ",tbl_IPDD_cost_data[FP],tbl_IPDD_cost_data[LCM_Segment]),5,1),"""")"

        ' O - Project Name
        ws.Cells(r, COL_PROJ_NAME).formula = _
            "=IFERROR(XLOOKUP($N" & r & ",tbl_IPDD_cost_data[FP],tbl_IPDD_cost_data[Row Labels]),""Not Found"")"

        ' P - Current PowerPlan In-Service Date
        ws.Cells(r, COL_CURRENT_ISD).formula = _
            "=IFERROR(XLOOKUP($N" & r & ",tbl_IPDD_cost_data[FP],tbl_IPDD_cost_data[ISD]),"""")"

        ' AP - Prior Year Spend with GGN ownership adjustment
        ws.Cells(r, COL_PRIOR_SPEND).formula = _
            "=IFERROR((XLOOKUP($N" & r & ",tbl_IPDD_cost_data[FP],tbl_IPDD_cost_data[Prior Years Spend])),0)" & _
            "*IF($K" & r & "=""GGN"",0.9164,1)"

        ' TARGET requested (W-AB): user-editable - only restore if empty
        For i = 0 To 5
            With ws.Cells(r, COL_TARGET_REQ_START + i)
                If IsEmpty(.value) Then
                    .formula = "=" & ColLetter(COL_TARGET_CUR_START + i) & r
                End If
            End With
        Next i

        ' TARGET current (AC-AH)
        ws.Cells(r, COL_TARGET_CUR_START).formula = _
            "=IFERROR(XLOOKUP($N" & r & ",tbl_IPDD_cost_data[FP],tbl_IPDD_cost_data[Target_CY]),0)"
        For i = 1 To 5
            ws.Cells(r, COL_TARGET_CUR_START + i).formula = _
                "=IFERROR(XLOOKUP($N" & r & ",tbl_IPDD_cost_data[Row Labels],tbl_IPDD_cost_data[Target_CY+" & i & "]),0)"
        Next i

        ' TARGET variance (AI-AN)
        For i = 0 To 5
            ws.Cells(r, COL_TARGET_VAR_START + i).formula = _
                "=" & ColLetter(COL_TARGET_REQ_START + i) & r & _
                "-" & ColLetter(COL_TARGET_CUR_START + i) & r
        Next i

        ' CLOSINGS requested (AQ-AV): user-editable - only restore if empty
        For i = 0 To 5
            With ws.Cells(r, COL_CLOS_REQ_START + i)
                If IsEmpty(.value) Then
                    .formula = "=" & ColLetter(COL_CLOS_CUR_START + i) & r
                End If
            End With
        Next i

        ' CLOSINGS current (AW-BB)
        ' TODO: confirm field names in tbl_IPDD_cost_data for Closings scenario
        ws.Cells(r, COL_CLOS_CUR_START).formula = _
            "=IFERROR(XLOOKUP($N" & r & ",tbl_IPDD_cost_data[FP],tbl_IPDD_cost_data[Closings_CY]),0)"
        For i = 1 To 5
            ws.Cells(r, COL_CLOS_CUR_START + i).formula = _
                "=IFERROR(XLOOKUP($N" & r & ",tbl_IPDD_cost_data[Row Labels],tbl_IPDD_cost_data[Closings_CY+" & i & "]),0)"
        Next i

        ' CLOSINGS variance (BC-BH)
        For i = 0 To 5
            ws.Cells(r, COL_CLOS_VAR_START + i).formula = _
                "=" & ColLetter(COL_CLOS_REQ_START + i) & r & _
                "-" & ColLetter(COL_CLOS_CUR_START + i) & r
        Next i

        ' Backfill OpCo if Site is set but OpCo is blank
        Dim siteVal As String
        siteVal = Trim(ws.Cells(r, COL_SITE).value & "")
        If siteVal <> "" And Trim(ws.Cells(r, COL_OPCO).value & "") = "" Then
            WriteSiteDefaults ws, r, siteVal
        End If

    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Re-initialization complete. " & (lastRow - FIRST_DATA_ROW + 1) & " rows processed.", _
           vbInformation, "Table Maintenance"
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "ReinitializeEntryTable failed on row " & r & ":" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Table Maintenance Error"
End Sub


' ----------------------------------------------------------------------------
' Sub: ApplySiteDefaults
' Purpose: Fires on Worksheet_Change for col K. Writes OpCo for the row.
'          Seg requires no action here - it recalculates via its formula.
' ----------------------------------------------------------------------------
Public Sub ApplySiteDefaults(changedRow As Long)
    On Error GoTo ErrHandler

    If changedRow < FIRST_DATA_ROW Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_DATA)

    Application.EnableEvents = False
    WriteSiteDefaults ws, changedRow, Trim(ws.Cells(changedRow, COL_SITE).value & "")
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Debug.Print "ApplySiteDefaults error row " & changedRow & ": " & Err.Description
End Sub


' ============================================================================
' PRIVATE HELPERS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: WriteSiteDefaults
' Purpose: Write OpCo based on site code. HDQ has no OpCo - cell is cleared.
'          Seg is NOT set here - it is a formula in col I.
' ----------------------------------------------------------------------------
Private Sub WriteSiteDefaults(ws As Worksheet, r As Long, siteCode As String)
    Dim opco As String

    Select Case UCase(Trim(siteCode))
        Case "ANO":    opco = "ELA"
        Case "GGN":    opco = "SERI"
        Case "RBS":    opco = "ELL"
        Case "WF3":    opco = "ELL"
        Case "HQN", "HDQ": opco = ""    ' no OpCo for HDQ
        Case "FLEET":  Exit Sub          ' read-only - do not overwrite
        Case "":       ws.Cells(r, COL_OPCO).ClearContents: Exit Sub
        Case Else:     Exit Sub          ' unknown site - leave alone
    End Select

    ws.Cells(r, COL_OPCO).value = opco
End Sub


' ----------------------------------------------------------------------------
' Sub: ApplyListValidation
' ----------------------------------------------------------------------------
Private Sub ApplyListValidation(rng As Range, formula As String)
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Operator:=xlBetween, Formula1:=formula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With
End Sub


' ----------------------------------------------------------------------------
' Sub: ApplyInlineListValidation
' ----------------------------------------------------------------------------
Private Sub ApplyInlineListValidation(rng As Range, commaList As String)
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Operator:=xlBetween, Formula1:=commaList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With
End Sub


' ----------------------------------------------------------------------------
' Function: ColLetter
' Purpose: Convert column number to letter string for formula construction.
' ----------------------------------------------------------------------------
Private Function ColLetter(colNum As Long) As String
    Dim s       As String
    Dim n       As Long
    Dim modVal  As Long
    s = ""
    n = colNum
    Do While n > 0
        modVal = (n - 1) Mod 26
        s = Chr(65 + modVal) & s
        n = (n - 1 - modVal) \ 26
    Loop
    ColLetter = s
End Function


