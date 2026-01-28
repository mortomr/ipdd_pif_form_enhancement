# Copy Buttons Setup Guide

## Overview

The copy buttons feature allows users to easily copy summary cost data from the hidden "Summary Cost Data" worksheet. Two buttons are added to each of the TA_Inflight and PIF worksheets:

- **Copy Picture Button** (Blue): Copies the linked image to clipboard
- **Copy Data Button** (Green): Copies the underlying data range as values to clipboard

## Button Locations and Functionality

### TA_Inflight Worksheet
- **Copy Fleet View Picture**: Copies the "FleetView" linked picture ('Summary Cost Data'!A2:U8)
- **Copy Fleet View Data**: Copies data range A2:U8 from Summary Cost Data as values

### PIF Worksheet
- **Copy Site View Picture**: Copies the "SiteView" linked picture ('Summary Cost Data'!A13:U15)
- **Copy Site View Data**: Copies data range A13:U15 from Summary Cost Data as values

## Initial Setup

### Step 1: Import the VBA Module
1. Open the Excel workbook with macros enabled
2. Press `Alt+F11` to open the VBA Editor
3. Go to **File > Import File...**
4. Select `mod_CopyButtons.bas`
5. Click **Open**

### Step 2: Create the Buttons
1. In the VBA Editor, press `F5` (or go to **Run > Run Sub/UserForm**)
2. Select `SetupCopyButtons` from the list
3. Click **Run**
4. A confirmation message will appear when buttons are created successfully

### Step 3: Position the Buttons (Optional)
The buttons are created at default positions (upper-right area). You can move them:
1. Go to each worksheet (TA_Inflight and PIF)
2. Click on a button to select it
3. Drag it to your preferred location
4. Repeat for all four buttons

## Prerequisites

Before using the copy buttons, ensure:
1. The "Summary Cost Data" worksheet exists and is properly configured
2. Linked pictures are created on the target worksheets:
   - **FleetView** picture on TA_Inflight worksheet (linked to 'Summary Cost Data'!A2:U8)
   - **SiteView** picture on PIF worksheet (linked to 'Summary Cost Data'!A13:U15)

### Creating Linked Pictures (If Not Present)

If the linked pictures don't exist yet:

#### For TA_Inflight (FleetView):
1. Go to "Summary Cost Data" sheet
2. Select range A2:U8
3. Copy the range (Ctrl+C)
4. Go to "TA_Inflight" sheet
5. On the Home ribbon, click the dropdown under **Paste**
6. Select **Linked Picture**
7. Right-click the picture > **Edit Alt Text** > Set name to "FleetView"

#### For PIF (SiteView):
1. Go to "Summary Cost Data" sheet
2. Select range A13:U15
3. Copy the range (Ctrl+C)
4. Go to "PIF" sheet
5. On the Home ribbon, click the dropdown under **Paste**
6. Select **Linked Picture**
7. Right-click the picture > **Edit Alt Text** > Set name to "SiteView"

## Usage

### Copying Pictures
1. Click the blue **Copy Picture** button on the desired worksheet
2. A confirmation message will appear
3. The picture is now on your clipboard
4. Paste into PowerPoint, Word, Outlook, or any application (Ctrl+V)

### Copying Data
1. Click the green **Copy Data** button on the desired worksheet
2. A confirmation message will appear
3. The data range is now on your clipboard
4. Paste into Excel, Word, or other applications
5. **Note**: Use **Paste Special > Values** to paste only the values without formatting

## Troubleshooting

### "Picture Not Found" Error
**Problem**: The macro cannot find the linked picture by name.

**Solutions**:
1. Verify the picture exists on the worksheet
2. Check the picture name:
   - Right-click picture > **Edit Alt Text** or use Selection Pane (Alt+F10)
   - Should be "FleetView" for TA_Inflight or "SiteView" for PIF
3. If the name is different, either:
   - Rename the picture to match expected name, OR
   - Edit the constants in mod_CopyButtons.bas (lines 13-14)

### "Error copying data" Message
**Problem**: The source range cannot be found.

**Solutions**:
1. Verify "Summary Cost Data" worksheet exists
2. Check that the worksheet is not protected
3. Verify the source ranges exist and contain data:
   - Fleet View: 'Summary Cost Data'!A2:U8
   - Site View: 'Summary Cost Data'!A13:U15

### Buttons Disappeared After Saving
**Problem**: Buttons are not visible when reopening the workbook.

**Solutions**:
1. Save the workbook as macro-enabled (.xlsm) format
2. If buttons were accidentally deleted, re-run `SetupCopyButtons` macro
3. Check if worksheet protection is hiding buttons

### Button Doesn't Respond
**Problem**: Clicking button does nothing.

**Solutions**:
1. Check that macros are enabled (File > Options > Trust Center)
2. Verify the macro assignment:
   - Right-click button > **Assign Macro**
   - Ensure correct macro is assigned
3. Check for VBA runtime errors (press Alt+F11 to view VBA Editor)

## Customization

### Changing Button Position
Edit `SetupCopyButtons()` in mod_CopyButtons.bas:
```vba
leftPosition = 1000  ' Change this value (pixels from left)
topPosition = 10     ' Change this value (pixels from top)
```

### Changing Button Colors
Edit the RGB values in `SetupCopyButtons()`:
```vba
.Fill.ForeColor.RGB = RGB(68, 114, 196)  ' Blue - Picture buttons
.Fill.ForeColor.RGB = RGB(112, 173, 71)  ' Green - Data buttons
```

### Changing Button Text
Edit the button creation code in `SetupCopyButtons()`:
```vba
.TextFrame2.TextRange.Text = "Your Custom Text Here"
```

### Changing Source Ranges
Edit the constants at the top of mod_CopyButtons.bas:
```vba
Private Const FLEET_VIEW_RANGE As String = "'Summary Cost Data'!$A$2:$U$8"
Private Const SITE_VIEW_RANGE As String = "'Summary Cost Data'!$A$13:$U$15"
```

## Technical Details

### Module: mod_CopyButtons.bas
- **Public API Functions**: 4 copy functions (Fleet/Site × Picture/Data)
- **Utility Function**: SetupCopyButtons() for button creation
- **Constants**: Picture names and source range references
- **Error Handling**: Comprehensive error messages for troubleshooting

### Button Specifications
- Type: Shape (msoShapeRectangle)
- Dimensions: 180px wide × 30px tall
- Colors:
  - Picture buttons: RGB(68, 114, 196) - Blue
  - Data buttons: RGB(112, 173, 71) - Green
- Font: Bold, Size 10, White text

### Dependencies
- Excel 2016+ (for TextFrame2 properties)
- Macros must be enabled
- Standard VBA object model (no external references required)

## Version History

- **2025-11-20**: Initial implementation
  - Created mod_CopyButtons.bas
  - Added 4 copy functions (Fleet/Site × Picture/Data)
  - Added SetupCopyButtons() utility
  - Comprehensive error handling
