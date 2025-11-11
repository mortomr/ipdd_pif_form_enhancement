Attribute VB_Name = "mod_SiteSetup"
' ============================================================================
' MODULE: mod_SiteSetup
' ============================================================================
' Purpose: One-time setup for site selection UI
' Author: Data Architecture Team
' Date: 2025-11-11
'
' Usage: Run SetupSiteSelection() once to create the Instructions sheet
'        with site dropdown and named range
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Sub: SetupSiteSelection
' Purpose: Create Instructions sheet with site selection dropdown
' Usage: Run this macro once to set up site selection UI
' ----------------------------------------------------------------------------
Public Sub SetupSiteSelection()
    On Error GoTo ErrHandler

    Dim wsInstructions As Worksheet
    Dim rngSiteCell As Range
    Dim rngListSource As Range
    Dim dvValidation As Validation

    Application.ScreenUpdating = False

    ' Step 1: Create or get Instructions worksheet
    On Error Resume Next
    Set wsInstructions = ThisWorkbook.Sheets("Instructions")
    On Error GoTo ErrHandler

    If wsInstructions Is Nothing Then
        ' Create new Instructions sheet
        Set wsInstructions = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsInstructions.Name = "Instructions"
    Else
        ' Clear existing content but keep the sheet
        wsInstructions.Cells.Clear
    End If

    ' Step 2: Set up the header
    With wsInstructions
        .Range("A1").Value = "PIF SUBMISSION SYSTEM - SITE SELECTION"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' Instructions
        .Range("A3").Value = "1. Select your site from the dropdown below:"
        .Range("A4").Value = "2. Fill out the PIF worksheet with your project data"
        .Range("A5").Value = "3. Click [Submit to Database] when ready"
        .Range("A6").Value = ""
        .Range("A7").Value = "IMPORTANT: Only submit data for YOUR site!"
        .Range("A7").Font.Bold = True
        .Range("A7").Font.Color = RGB(192, 0, 0)

        ' Site selection label
        .Range("A10").Value = "Select Site:"
        .Range("A10").Font.Bold = True
        .Range("A10").Font.Size = 12

        ' Site dropdown cell
        .Range("B10").Value = ""
        .Range("B10").Interior.Color = RGB(255, 255, 200)

        ' Create list of sites in hidden area
        .Range("E1").Value = "ANO"
        .Range("E2").Value = "GGN"
        .Range("E3").Value = "RBN"
        .Range("E4").Value = "WF3"
        .Range("E5").Value = "HQN"
        .Range("E6").Value = "Fleet"

        Set rngListSource = .Range("E1:E6")
        Set rngSiteCell = .Range("B10")

        ' Add data validation dropdown
        With rngSiteCell.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=Instructions!$E$1:$E$6"
            .IgnoreBlank = False
            .InCellDropdown = True
            .ShowInput = True
            .InputTitle = "Site Selection"
            .InputMessage = "Select your site from the dropdown list." & vbCrLf & _
                           "Fleet = View all sites (read-only)"
            .ShowError = True
            .ErrorTitle = "Invalid Site"
            .ErrorMessage = "You must select a site from the dropdown list."
        End With

        ' Format dropdown cell
        With rngSiteCell
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
        End With

        ' Add site descriptions
        .Range("A12").Value = "Site Codes:"
        .Range("A12").Font.Bold = True
        .Range("A13").Value = "ANO - Arkansas Nuclear One"
        .Range("A14").Value = "GGN - Grand Gulf Nuclear"
        .Range("A15").Value = "RBN - River Bend Nuclear"
        .Range("A16").Value = "WF3 - Waterford 3"
        .Range("A17").Value = "HQN - Headquarters"
        .Range("A18").Value = "Fleet - All Sites (Read-Only)"
        .Range("A18").Font.Italic = True

        ' Hide the list source columns
        .Columns("E:E").Hidden = True

        ' Auto-fit columns
        .Columns("A:D").AutoFit

        ' Protect certain cells but allow dropdown editing
        .Range("B10").Locked = False
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                 UserInterfaceOnly:=True, AllowFormattingCells:=True
    End With

    ' Step 3: Create Named Range for site selection
    On Error Resume Next
    ThisWorkbook.Names("SelectedSite").Delete
    On Error GoTo ErrHandler

    ThisWorkbook.Names.Add Name:="SelectedSite", RefersTo:=wsInstructions.Range("B10")

    ' Step 4: Set default value
    If wsInstructions.Range("B10").Value = "" Then
        wsInstructions.Range("B10").Value = "ANO"
    End If

    ' Step 5: Activate Instructions sheet
    wsInstructions.Activate
    wsInstructions.Range("B10").Select

    Application.ScreenUpdating = True

    MsgBox "Site selection UI created successfully!" & vbCrLf & vbCrLf & _
           "Named Range 'SelectedSite' has been created at Instructions!B10" & vbCrLf & vbCrLf & _
           "Default site: ANO" & vbCrLf & vbCrLf & _
           "You can now use this dropdown to select your site before submission.", _
           vbInformation, "Setup Complete"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error setting up site selection:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Setup Error"
End Sub

' ----------------------------------------------------------------------------
' Function: GetSelectedSite
' Purpose: Helper function to retrieve selected site from named range
' Returns: Site code (ANO, GGN, RBN, WF3, HQN, Fleet) or empty string if not set
' ----------------------------------------------------------------------------
Public Function GetSelectedSite() As String
    On Error Resume Next
    GetSelectedSite = Trim(ThisWorkbook.Names("SelectedSite").RefersToRange.Value)
    If Err.Number <> 0 Then GetSelectedSite = ""
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Function: ValidateSelectedSite
' Purpose: Check if a valid site has been selected
' Returns: True if valid site selected, False otherwise
' ----------------------------------------------------------------------------
Public Function ValidateSelectedSite() As Boolean
    Dim site As String
    site = GetSelectedSite()

    If site = "" Then
        MsgBox "Please select a site from the Instructions worksheet before submitting.", _
               vbExclamation, "Site Not Selected"
        ValidateSelectedSite = False
        Exit Function
    End If

    ' Validate site is one of the allowed values
    Select Case UCase(site)
        Case "ANO", "GGN", "RBN", "WF3", "HQN", "FLEET"
            ValidateSelectedSite = True
        Case Else
            MsgBox "Invalid site selected: " & site & vbCrLf & vbCrLf & _
                   "Please select a valid site from the dropdown.", _
                   vbExclamation, "Invalid Site"
            ValidateSelectedSite = False
    End Select
End Function
