# Design Observations and Questions

**Document Purpose**: Track design decisions, architectural concerns, and open questions about the PIF system implementation.

**Last Updated**: 2025-11-12

---

## Critical Observations

### 1. Abandoned Database Connections (RESOLVED)

**Issue**: PIF_Archive and PIF_Inflight worksheet refreshes left abandoned OLEDB connections in SQL Server.

**Root Cause**:
- `CreateOrRefreshQueryTable` function deleted QueryTables without closing their WorkbookConnections
- Each refresh created a NEW OLEDB connection to SQL Server
- Old connections remained open indefinitely, consuming SQL Server resources
- After multiple refreshes, connection pool could be exhausted

**Technical Details**:
- QueryTables created with `ws.QueryTables.Add(Connection:=connStr, ...)` automatically create WorkbookConnection objects
- Deleting QueryTables (line 270) did NOT automatically close their connections
- Excel maintained orphaned connections in `ThisWorkbook.Connections` collection
- These connections held open sessions in SQL Server

**Solution Implemented** (2025-11-12):
- Added connection cleanup logic BEFORE deleting QueryTables (lines 258-270)
- Iterates through `ThisWorkbook.Connections` collection in reverse
- Identifies OLEDB connections matching `SQL_SERVER` and `SQL_DATABASE`
- Explicitly calls `conn.Delete` to close and remove connections
- Ensures clean slate before creating new QueryTable connections

**Code Location**: `mod_WorksheetQuery.bas:258-270`

**Impact**:
- ✅ Prevents connection pool exhaustion in SQL Server
- ✅ Reduces server memory usage
- ✅ Eliminates "max pool size reached" errors
- ✅ Improves refresh reliability

**Status**: **RESOLVED** - Fix committed 2025-11-12

---

### 2. QueryTable Refresh Functionality Fixed

**Previous Issue**: The PIF_Archive and PIF_Inflight worksheets had non-functional native refresh capabilities (see QUERYTABLE_REFRESH_FIX_OPTIONS.md for details).

**Solution Implemented**: Option 1 (QueryTable-Only Approach)
- Uses QueryTable without ListObject conversion
- Native Excel refresh now works perfectly
- Right-click "Refresh" and "Data > Refresh All" are fully functional

**Status**: **RESOLVED** - See mod_WorksheetQuery.bas for implementation

---

### 3. Fleet vs Site-Specific View Disconnect

**Issue**: The system was designed from a fleet perspective but is less useful in site-specific deployments.

**Technical Details**:
- SQL queries filter by site when `SelectedSite <> "FLEET"` (lines 64-66, 148-150)
- Site-specific users cannot access fleet-wide data
- No mechanism to temporarily switch to fleet view without changing site selection
- Slicers on PIF_Data sheet are optimized for fleet analysis, not project-specific work

**Observations**:

#### Slicers Are Fleet-Optimized
- **Appropriate for Fleet**: Slicers excel at cross-OPCO impact analysis
  - Filter by status across all sites
  - Compare project types across regions
  - Identify trends and patterns fleet-wide

- **Less Useful for Site-Specific**: Project managers working on individual PIFs
  - Slicers add UI complexity for single-project focus
  - Fleet-wide filtering is irrelevant when only working on local projects
  - Cognitive overhead for features that won't be used

#### Architecture Tension
The system exhibits two conflicting use cases:

**Fleet Perspective** (Original Design Intent):
- Central office analyzing impacts across all OPCOs
- Strategic decision-making requires broad visibility
- Slicers, filters, and aggregate views are essential
- Needs "helicopter view" of all projects

**Site-Specific Perspective** (Current Deployment Model):
- Individual sites managing their own PIFs
- Focus on local project details, not comparisons
- Limited interest in other sites' data
- Needs "ground-level view" of local work

**The Disconnect**:
- Fleet-optimized UI in a site-specific deployment **reduces utility**
- Features designed for comparative analysis don't serve single-site users
- PIF_Archive and PIF_Inflight sheets filter by site, but PIF_Data sheet has fleet-wide slicers

**Current Limitations**:
1. Site-specific users cannot see fleet data (security/intentional?)
2. No "peek at fleet" option for benchmarking or context
3. Fleet users would see all data, but sites are isolated
4. Archive/Inflight worksheets respect site boundaries, but slicers don't

**Questions**:
1. **Is site isolation intentional for security/privacy?**
   - Should sites be prevented from seeing other sites' data?
   - Or is this just an artifact of the SelectedSite mechanism?

2. **Should there be a "View Fleet Data" toggle?**
   - Allow site users to temporarily view fleet-wide for context
   - Useful for benchmarking or understanding corporate priorities
   - Could be role-based (some users see fleet, others don't)

3. **Should slicers be conditionally shown based on SelectedSite?**
   - Hide slicers when SelectedSite != "FLEET"
   - Reduce UI complexity for site-specific users
   - Slicers only appear for fleet analysts

4. **Is this a single workbook with multi-site deployment?**
   - Each site gets their own copy with SelectedSite hardcoded?
   - Or one shared workbook with role-based site selection?

**Status**: **ARCHITECTURAL QUESTION** - Requires product owner input

---

### 3. Refresh Style and Table Overlap Issues

**Issue**: Comment on line 326 indicates previous conflicts: "Changed from xlInsertDeleteCells to avoid overlap"

**Technical Details**:
- `xlInsertDeleteCells` causes "overlap error" when working with Excel Tables
- Changed to `xlOverwriteCells` to avoid conflict
- But this change doesn't fix the underlying refresh problem

**Hypothesis**:
- QueryTables and ListObjects (Excel Tables) have conflicting refresh mechanisms
- Trying to have both on the same range causes Excel internal conflicts
- The delete/recreate pattern was an attempt to work around this

**Root Cause**:
- Attempting to get "best of both worlds":
  - QueryTable for native refresh capability
  - ListObject for beautiful table formatting
- Excel doesn't fully support this hybrid approach

**Status**: **TECHNICAL LIMITATION** - Choose one or the other

---

## Design Questions Requiring Input

### Question 1: QueryTable vs ListObject Trade-off

**Should we prioritize:**
- **A) Native refresh capability** (keep QueryTable, lose table formatting)
- **B) Beautiful table UI** (keep ListObject, accept VBA-only refresh)
- **C) Hybrid approach** (current broken state)
- **D) Migrate to Power Query** (modern solution, requires Excel 2016+)

### Question 2: Site Isolation Policy

**Is site-specific data isolation:**
- **A) Intentional security requirement** (sites must not see each other's data)
- **B) Incidental implementation detail** (can be changed if useful)
- **C) Configurable** (some users see fleet, others see site-only)

### Question 3: Slicer Utility in Site Deployments

**For site-specific users, should slicers:**
- **A) Be hidden** (reduce UI complexity, show only for fleet)
- **B) Remain visible** (users might find local filtering useful)
- **C) Be simplified** (show only relevant filters for single-site work)

### Question 4: Fleet View Access

**Should site-specific users have a "View Fleet" option?**
- **A) Yes** - Useful for benchmarking and context
- **B) No** - Keep sites isolated for security/simplicity
- **C) Role-based** - Some users have fleet access, others don't

---

## Recommendations

### Short-term (Low Effort)
1. **Document current behavior**: Users understand refresh is VBA-only
2. **Add visual indicator**: "Click 'Refresh Archive' button to update data"
3. **Hide native refresh UI**: Prevent confusion with disabled native refresh

### Medium-term (Moderate Effort)
1. **Fix QueryTable refresh**: Choose QueryTable OR ListObject, not both
2. **Add "View Fleet" toggle**: Allow temporary fleet view for context
3. **Conditionally show slicers**: Hide when SelectedSite != "FLEET"

### Long-term (Strategic)
1. **Migrate to Power Query**: Modern, native refresh, better performance
2. **Implement role-based access**: RBAC for fleet vs site visibility
3. **Redesign UI for dual-mode**: Optimize for both fleet and site use cases

---

## Open Questions Log

| # | Question | Asked By | Date | Status |
|---|----------|----------|------|--------|
| 1 | Is site data isolation intentional for security? | User | 2025-11-12 | Open |
| 2 | Should slicers be hidden in site-specific deployments? | User | 2025-11-12 | Open |
| 3 | Can we fix QueryTable refresh without losing table formatting? | User | 2025-11-12 | Open |
| 4 | Should site users have "peek at fleet" capability? | User | 2025-11-12 | Open |
| 5 | Is this one workbook (multi-site) or per-site copies? | User | 2025-11-12 | Open |

---

## Change History

| Date | Author | Change |
|------|--------|--------|
| 2025-11-12 | User | Initial document creation - identified QueryTable refresh issue and fleet/site disconnect |
