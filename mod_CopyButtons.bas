Attribute VB_Name = "mod_CopyButtons"
Option Explicit

' ===============================================================================
' MODULE: mod_CopyButtons
' PURPOSE: Provides copy functionality for Summary Cost Data images and ranges
' CREATED: 2025-11-20
' ===============================================================================

' Constants for linked picture names and source ranges
Private Const FLEET_VIEW_PICTURE_NAME As String = "FleetView"
Private Const SITE_VIEW_PICTURE_NAME As String = "SiteView"
Private Const FLEET_VIEW_RANGE As String = "'Summary Cost Data'!$A$2:$U$8"
Private Const SITE_VIEW_RANGE As String = "'Summary Cost Data'!$A$13:$U$15"
Private Const SUMMARY_SHEET_NAME As String = "Summary Cost Data"

' ===============================================================================
' PUBLIC API - Fleet View (TA_Inflight worksheet)
' ===============================================================================

Public Sub CopyFleetViewPicture()
    ' Copies the FleetView linked picture to clipboard
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim pic As Shape

    Set ws = ThisWorkbook.Worksheets("TA_Inflight")

    ' Find the FleetView picture
    On Error Resume Next
    Set pic = ws.Shapes(FLEET_VIEW_PICTURE_NAME)
    On Error GoTo ErrorHandler

    If pic Is Nothing Then
        MsgBox "Could not find picture named '" & FLEET_VIEW_PICTURE_NAME & "' on TA_Inflight sheet.", _
               vbExclamation, "Picture Not Found"
        Exit Sub
    End If

    ' Copy the picture
    pic.Copy

    MsgBox "Fleet View picture copied to clipboard. You can now paste it into another application.", _
           vbInformation, "Copy Successful"

    Exit Sub

ErrorHandler:
    MsgBox "Error copying Fleet View picture: " & Err.Description, vbCritical, "Copy Error"
End Sub

Public Sub CopyFleetViewData()
    ' Copies the Fleet View data range as values to clipboard
    On Error GoTo ErrorHandler

    Dim sourceRange As Range

    ' Reference the source range
    Set sourceRange = Range(FLEET_VIEW_RANGE)

    ' Copy the range
    sourceRange.Copy

    MsgBox "Fleet View data (A2:U8) copied to clipboard as values. " & vbCrLf & _
           "Use Paste Special > Values to paste without formatting.", _
           vbInformation, "Copy Successful"

    Exit Sub

ErrorHandler:
    MsgBox "Error copying Fleet View data: " & Err.Description, vbCritical, "Copy Error"
End Sub

' ===============================================================================
' PUBLIC API - Site View (PIF worksheet)
' ===============================================================================

Public Sub CopySiteViewPicture()
    ' Copies the SiteView linked picture to clipboard
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim pic As Shape

    Set ws = ThisWorkbook.Worksheets("Target Adjustment")

    ' Find the SiteView picture
    On Error Resume Next
    Set pic = ws.Shapes(SITE_VIEW_PICTURE_NAME)
    On Error GoTo ErrorHandler

    If pic Is Nothing Then
        MsgBox "Could not find picture named '" & SITE_VIEW_PICTURE_NAME & "' on PIF sheet.", _
               vbExclamation, "Picture Not Found"
        Exit Sub
    End If

    ' Copy the picture
    pic.Copy

    MsgBox "Site View picture copied to clipboard. You can now paste it into another application.", _
           vbInformation, "Copy Successful"

    Exit Sub

ErrorHandler:
    MsgBox "Error copying Site View picture: " & Err.Description, vbCritical, "Copy Error"
End Sub

Public Sub CopySiteViewData()
    ' Copies the Site View data range as values to clipboard
    On Error GoTo ErrorHandler

    Dim sourceRange As Range

    ' Reference the source range
    Set sourceRange = Range(SITE_VIEW_RANGE)

    ' Copy the range
    sourceRange.Copy

    MsgBox "Site View data (A13:U15) copied to clipboard as values. " & vbCrLf & _
           "Use Paste Special > Values to paste without formatting.", _
           vbInformation, "Copy Successful"

    Exit Sub

ErrorHandler:
    MsgBox "Error copying Site View data: " & Err.Description, vbCritical, "Copy Error"
End Sub

' ===============================================================================
' UTILITY - Button Setup (Run once to create buttons)
' ===============================================================================

Public Sub SetupCopyButtons()
    ' Creates the copy buttons on TA_Inflight and PIF worksheets
    ' Run this macro once to set up the buttons
    On Error GoTo ErrorHandler

    Dim wsInflight As Worksheet
    Dim wsPIF As Worksheet
    Dim btn As Shape
    Dim topPosition As Double
    Dim leftPosition As Double

    Set wsInflight = ThisWorkbook.Worksheets("TA_Inflight")
    Set wsPIF = ThisWorkbook.Worksheets("Target Adjustment")

    ' ===== TA_Inflight Worksheet Buttons =====
    ' Position buttons in upper right area (adjust as needed)
    leftPosition = 1000 ' Adjust based on your layout
    topPosition = 10

    ' Delete existing buttons if they exist
    On Error Resume Next
    wsInflight.Shapes("btnCopyFleetPicture").Delete
    wsInflight.Shapes("btnCopyFleetData").Delete
    On Error GoTo ErrorHandler

    ' Create Fleet View Picture button
    Set btn = wsInflight.Shapes.AddShape(msoShapeRectangle, leftPosition, topPosition, 180, 30)
    With btn
        .Name = "btnCopyFleetPicture"
        .TextFrame2.TextRange.Text = "Copy Fleet View Picture"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(68, 114, 196) ' Blue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text
        .OnAction = "CopyFleetViewPicture"
    End With

    ' Create Fleet View Data button (below picture button)
    Set btn = wsInflight.Shapes.AddShape(msoShapeRectangle, leftPosition, topPosition + 35, 180, 30)
    With btn
        .Name = "btnCopyFleetData"
        .TextFrame2.TextRange.Text = "Copy Fleet View Data"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(112, 173, 71) ' Green
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text
        .OnAction = "CopyFleetViewData"
    End With

    ' ===== PIF Worksheet Buttons =====
    ' Delete existing buttons if they exist
    On Error Resume Next
    wsPIF.Shapes("btnCopySitePicture").Delete
    wsPIF.Shapes("btnCopySiteData").Delete
    On Error GoTo ErrorHandler

    ' Create Site View Picture button
    Set btn = wsPIF.Shapes.AddShape(msoShapeRectangle, leftPosition, topPosition, 180, 30)
    With btn
        .Name = "btnCopySitePicture"
        .TextFrame2.TextRange.Text = "Copy Site View Picture"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(68, 114, 196) ' Blue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text
        .OnAction = "CopySiteViewPicture"
    End With

    ' Create Site View Data button (below picture button)
    Set btn = wsPIF.Shapes.AddShape(msoShapeRectangle, leftPosition, topPosition + 35, 180, 30)
    With btn
        .Name = "btnCopySiteData"
        .TextFrame2.TextRange.Text = "Copy Site View Data"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(112, 173, 71) ' Green
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text
        .OnAction = "CopySiteViewData"
    End With

    MsgBox "Copy buttons have been created successfully!" & vbCrLf & vbCrLf & _
           "TA_Inflight sheet: Copy Fleet View Picture, Copy Fleet View Data" & vbCrLf & _
           "PIF sheet: Copy Site View Picture, Copy Site View Data" & vbCrLf & vbCrLf & _
           "You can move these buttons to your preferred location.", _
           vbInformation, "Button Setup Complete"

    Exit Sub

ErrorHandler:
    MsgBox "Error setting up buttons: " & Err.Description, vbCritical, "Setup Error"
End Sub
