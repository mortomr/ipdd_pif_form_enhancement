# PIF DATA ENTRY INSTRUCTIONS

## Getting Started

1. **Select Your Site**: Go to the "Instructions" worksheet and select your site from the dropdown (ANO, GGN, RBN, WF3, HQN)
2. **Enter PIF Name**: Enter the PIF name in cell A1 of the "PIF" tab (reference EN-PM-100 for naming guidance)

---

## Column Reference Guide

### Row/Data Columns (A-T)

| Column | Field Name | Required | Explanation |
|--------|------------|----------|-------------|
| **A** | (Row Label) | - | Row identifier (do not modify) |
| **B** | (Line #) | - | Sequential line number for reference |
| **C** | Archive | No | Checkbox: Mark TRUE to archive this record permanently |
| **D** | Include | No | Checkbox: Mark TRUE to include in archival (requires Archive=TRUE) |
| **E** | Accounting | No | Accounting treatment (if applicable) |
| **F** | Change Type | Yes | New, Revision, or Cancellation |
| **G** | PIF ID | Yes | Unique PIF identifier (16 char max) |
| **H** | SEG | No | Segment number |
| **I** | Opco | No | Operating company code (4 char max) |
| **J** | Site | Yes | Site code: ANO, GGN, RBN, WF3, or HQN |
| **K** | Strategic Rank | No | Strategic rank from LCM (26 char max) |
| **L** | From Blanket? | No | Yes/No - Is funding from blanket project? |
| **M** | Project # | Yes | Funding project number (10 char max) |
| **N** | Project Name | Yes | Project name (35 char max) |
| **O** | Original FP ISD | No | Current funding project in-service date (pre-change) |
| **P** | Revised FP ISD | No | New requested funding project in-service date |
| **Q** | LCM Issue | No | LCM issue associated with project (11 char max) |
| **R** | Status | No | Approved, Dispositioned, Pending, Parked, or Active (58 char max) |
| **S** | Category | No | Project category (26 char max) |
| **T** | Justification | No | Thorough explanation of change and rationale (192 char max) |

### Cost Columns - Target Scenario (U-AL)

| Columns | Field Group | Years | Explanation |
|---------|-------------|-------|-------------|
| **U-Z** | Target Requested | 2025-2030 | New requested annual project cashflow (Target scenario) |
| **AA-AF** | Target Current | 2025-2030 | Current approved cashflow (auto-calculated or manual) |
| **AG-AL** | Target Variance | 2025-2030 | Difference between Requested and Current (usually formula) |

### Additional Fields (AM-AN)

| Column | Field Name | Explanation |
|--------|------------|-------------|
| **AM** | Moving ISD Year | Year indicator if ISD is changing (1 char) |
| **AN** | Prior Year Spend | Historical spending prior to current fiscal year |

### Cost Columns - Closings Scenario (AO-BF)

| Columns | Field Group | Years | Explanation |
|---------|-------------|-------|-------------|
| **AO-AT** | Closings Requested | 2025-2030 | New requested annual cashflow (Closings scenario) |
| **AU-AZ** | Closings Current | 2025-2030 | Current approved cashflow (auto-calculated or manual) |
| **BA-BF** | Closings Variance | 2025-2030 | Difference between Requested and Current (usually formula) |

**Note**: Columns with formulas (typically variance columns and some current columns) are greyed out - do not modify these as they calculate automatically.

---

## Archival Instructions (NEW!)

Use checkboxes in columns C and D to control permanent archival:

- **Archive = TRUE, Include = TRUE**: Record will be PERMANENTLY archived when you use [Finalize Month]
- **Archive = FALSE or Include = FALSE**: Record stays in working inflight tables (not archived)
- **Important**: Only mark records for archival when decisions are final (typically month-end)

---

## Modifying Rows

### Adding Rows (Easy Method)
1. Click the **[Add Row]** button at the top of the sheet
2. A new blank row will be inserted above the totals row with formulas/formatting already applied
3. Fill in the required data fields

### Adding Rows (Manual Method)
1. Select the last row in the table (with sum values), right-click and select "Insert"
2. Copy the last row of data by highlighting columns A through BF, right-click and select "Copy"
3. Select column A in the new row, right-click and select "Paste Formulas"
4. Right-click again and select "Paste Formatting"
5. Update non-shaded cells with project information

### Deleting Rows
1. Select the row(s) you want to delete
2. Right-click and select "Delete"
3. Confirm deletion

---

## Submission Workflow (NEW!)

### During the Month (In-Progress Work)
Use the **[Save Snapshot]** button:
- Saves your current work to the inflight database tables
- Does NOT permanently archive records
- Can be used multiple times as you make changes
- Safe for ongoing work

### At Month-End (Final Decisions)
Use the **[Finalize Month]** button:
- Saves snapshot to inflight tables
- PERMANENTLY archives records marked with Archive=TRUE AND Include=TRUE
- Updates the Archive and Inflight query worksheets
- Use only when all decisions are final

---

## Query Worksheets (NEW!)

Two new worksheets show your data in familiar PIF format:

### PIF_Inflight
- Shows current working snapshot (in-progress data)
- Filtered by your selected site
- Read-only - refresh using [Refresh Inflight] button

### PIF_Archive
- Shows permanently archived records
- Filtered by your selected site
- Read-only - refresh using [Refresh Archive] button

**Fleet Users**: Can view all sites in both worksheets, but cannot save or archive data.

---

## Tips

1. **Required Fields**: PIF ID, Project #, Change Type, and Site are required
2. **Character Limits**: Watch for character limits noted above - exceeding them may cause errors
3. **Site Selection**: Always verify correct site is selected before saving/archiving
4. **Fleet Access**: Fleet is read-only for viewing all sites - use a specific site to make changes
5. **Archive Carefully**: Archival is permanent - only mark records when final
6. **Validation**: Run validation before submission to catch errors early

---

## Common Issues

**Error: "Site Not Selected"**
- Go to Instructions worksheet and select a site from the dropdown

**Error: "Fleet cannot submit data"**
- Fleet is read-only - select a specific site (ANO, GGN, RBN, WF3, HQN) to save data

**Error: "Missing Required Field"**
- Ensure PIF ID, Project #, Change Type, and Site are filled in for all data rows

**Data not appearing in Archive worksheet**
- Check that Archive=TRUE AND Include=TRUE for records you want archived
- Ensure you used [Finalize Month] button (not [Save Snapshot])
- Refresh the worksheet using [Refresh Archive] button
