# FTC DECODE 2025-2026 Match Review Scoring Spreadsheet

## Overview

A Google Apps Script that builds a complete match review scoring system in Google Sheets for FTC DECODE 2025-2026. Designed for 2-6 referees to independently score solo matches (one match per team, no opponents), with automatic score calculation, disagreement detection, and a Final Scores sheet for the head referee. Supports up to 500 teams (MAX_TEAMS).

The script is parameterized for annual reuse — see [Updating for a New Season](#updating-for-a-new-season).

## Setup Instructions

1. Create a new Google Spreadsheet (or open the target one)
2. Go to **Extensions > Apps Script**
3. Paste the entire contents of `DECODE_Scoring_Spreadsheet.gs`
4. Click **Run > buildAll()** and grant permissions when prompted
5. Follow the on-screen setup steps:
   - Enter team data in Config columns A-C (number, name, video link) starting row 4
   - Enter referee names in Config row 2 (columns D-I)
   - Enter referee emails in Config row 3 (for per-referee protection)
   - Use **DECODE Scoring > Randomize Team Orders** menu (also renames sheets)
   - Use **DECODE Scoring > Apply Sheet Protection** menu (this also hides Config and unused referee sheets)

## Sheet Structure

### Config Sheet
- **Column A** (row 4+): Team numbers
- **Column B** (row 4+): Team names
- **Column C** (row 4+): Video links (YouTube URLs)
- **Columns D-I** (row 2): Referee names (used as sheet tab names)
- **Columns D-I** (row 3): Referee emails (for protection)
- **Columns J-O** (row 4+): Randomized team orders per referee (auto-generated)

### Referee Sheets (named by referee)
Each referee gets an independent sheet named from Config (e.g., "Paul", "Jeff"). Sheets are initially named "Referee 1" through "Referee 6" and renamed when Randomize or Rename is run. Unused referee sheets (still named "Referee N") are hidden automatically.

**Row 1**: Title bar (split merge: A1:C1 = progress counter in frozen zone, D1:X1 = title in scrollable zone)
**Row 2**: Point values as a quick reference (hidden by default)
**Row 3**: Column headers (color-coded by section)
**Row 4+**: Data rows

| Column | Field | Type | Notes |
|--------|-------|------|-------|
| A | Team # | Auto-filled | From randomization |
| B | Team Name | **Auto (VLOOKUP)** | From Config |
| C | Video | **Auto (VLOOKUP)** | From Config |
| D | Notes | Free text | ~200px wide |
| E | TOTAL SCORE | **Calculated** | MAX(0, Score w/o Fouls - Foul Deduction) |
| F | Score w/o Fouls | **Calculated** | Auto + TeleOp (before foul deduction) |
| G | Auto Score | **Calculated** | LEAVE(3) + CLS×3 + OVF×1 + PAT×2 |
| H | TeleOp Score | **Calculated** | CLS×3 + OVF×1 + DEPOT×1 + PAT×2 + BASE(0/5/10) |
| I | Foul Deduction | **Calculated** | Minor×5 + Major×15 |
| J | Minor Fouls | Whole number | |
| K | Major Fouls | Whole number | |
| L | G Rules | Multiselect | Codes of violated game rules (see [G Rules](#g-rules-multiselect)) |
| M | MOTIF | Dropdown (GPP/PGP/PPG/Not Shown) | Input field; "Not Shown" or blank → PATTERN=0 |
| N | LEAVE | Dropdown (Yes/No) | Did robot leave LAUNCH LINE during AUTO? |
| O | Auto CLASSIFIED | Whole number | Artifacts through SQUARE to RAMP during AUTO |
| P | Auto OVERFLOW | Whole number | Artifacts through SQUARE not to RAMP during AUTO |
| Q | Auto RAMP Colors | Text (G/P chars) | Artifact colors on RAMP at end of AUTO, GATE→SQUARE order |
| R | Auto PATTERN Count | **Calculated** (hidden) | Character-by-character match of Q vs REPT(MOTIF,3) |
| S | TeleOp CLASSIFIED | Whole number | Same as O but during TELEOP |
| T | TeleOp OVERFLOW | Whole number | Same as P but during TELEOP |
| U | TeleOp DEPOT | Whole number | Artifacts over DEPOT tape at end of TELEOP |
| V | TeleOp RAMP Colors | Text (G/P chars) | Artifact colors on RAMP at end of TELEOP |
| W | TeleOp PATTERN Count | **Calculated** (hidden) | Character-by-character match of V vs REPT(MOTIF,3) |
| X | BASE | Dropdown (None/Partial/Full) | Robot position on BASE TILE at end of TELEOP |

Frozen rows: 3 (title, points, headers). Frozen columns: 3 (Team #, Team Name, Video).

### FinalScores Sheet
Aggregation and score breakdown sheet for the head referee. **First tab** in the spreadsheet. Uses a 3-row header matching last year's layout.

**Row 1**: Merged category group headers (Teams | Referee | Scores | Fouls | G Rules | Autonomous Period | TeleOp Period)
**Row 2**: Point values per scoring element (hidden by default)
**Row 3**: Column names
**Row 4+**: Data

| Column | Field | Notes |
|--------|-------|-------|
| A | Team # | From Config |
| B | Team Name | From Config |
| C | Video | From Config |
| D | Scored By | Referee names who scored this team (composite check: MOTIF/LEAVE/AUTO_CLS) |
| E | Official Referee | **Editable dropdown** - select which referee's score to display |
| F | Refs Agree? | Yes/No/N/A - do all scoring referees agree on every input element? |
| G | Notes | Two-mode: plain text if Official Ref set; "RefName: text" format otherwise |
| H | Final Score | Per-field agreement or from effective referee |
| I | Score w/o Fouls | Per-field agreement or from effective referee |
| J | Auto Score | Per-field agreement or from effective referee |
| K | TeleOp Score | Per-field agreement or from effective referee |
| L | Foul Deduction | Per-field agreement or from effective referee |
| M | Minor Fouls | Per-field agreement or from effective referee |
| N | Major Fouls | Per-field agreement or from effective referee |
| O | G Rules | Per-field agreement or from effective referee |
| P | LEAVE | Per-field agreement or from effective referee |
| Q | Auto CLASSIFIED | Per-field agreement or from effective referee |
| R | Auto OVERFLOW | Per-field agreement or from effective referee |
| S | Auto RAMP Colors | Per-field agreement or from effective referee |
| T | Auto PATTERN Count | Per-field agreement or from effective referee (hidden) |
| U | TeleOp CLASSIFIED | Per-field agreement or from effective referee |
| V | TeleOp OVERFLOW | Per-field agreement or from effective referee |
| W | TeleOp DEPOT | Per-field agreement or from effective referee |
| X | TeleOp RAMP Colors | Per-field agreement or from effective referee |
| Y | TeleOp PATTERN Count | Per-field agreement or from effective referee (hidden) |
| Z | BASE | Per-field agreement or from effective referee |
| AA | effectiveRef | **Hidden** helper column — computes effective referee once per row |

Frozen rows: 3 (category headers, point values, column names). Frozen columns: 3 (Team #, Name, Video).

**Per-field agreement mode** (no Official Referee selected, 2+ refs scored): Each scoring field (H-Z) independently checks whether all referees agree on that field's value. Agreed values are displayed normally. Disagreed fields show empty with a **red background**, making it easy to identify exactly which fields need attention. The row also gets a yellow background (lower priority than per-field red).

**Effective referee mode** (Official Referee selected or single ref): That referee's data is shown directly for all fields. When only one referee has scored a team, that referee's data is auto-displayed without needing to select them.

## Scoring Rules (DECODE Section 10.5)

### Point Values

| Element | Points |
|---------|--------|
| LEAVE | 3 |
| CLASSIFIED | 3 per artifact |
| OVERFLOW | 1 per artifact |
| PATTERN match | 2 per matching position |
| DEPOT | 1 per artifact |
| BASE Full | 10 |
| BASE Partial | 5 |
| Minor Foul | -5 (from own score) |
| Major Foul | -15 (from own score) |

### PATTERN Scoring (RAMP Color Entry)
The OBELISK displays one of three MOTIFs: **GPP**, **PGP**, or **PPG** (G=green, P=purple).

The MOTIF repeats 3x across the 9 RAMP positions: e.g., GPP -> GPPGPPGPP.

The referee types the actual artifact colors on the RAMP in order from GATE to SQUARE. The spreadsheet auto-calculates how many positions match. Each match = 2 points.

**Example**: MOTIF = GPP, RAMP has 5 artifacts -> referee types `GPPGP`
- Position 1: G vs G (match)
- Position 2: P vs P (match)
- Position 3: P vs P (match)
- Position 4: G vs G (match)
- Position 5: P vs P (match)
- Result: 5 matches x 2 = 10 points

### Key Deviations from Official Rules
- **Foul points subtracted from own score** (official rules add to opponent's score, but these are solo matches with no opponent)
- **Total score floored at 0** (cannot go negative)

## Updating for a New Season

The script is designed for annual reuse. All game-specific values are in the **GAME CONFIGURATION** section at the top of the file. Refer to the [competition manual](https://ftc-resources.firstinspires.org/ftc/archive/2026/game/manual) for scoring rules and point values. To update for a new FTC season:

1. **`GAME_NAME`** / **`SEASON`** — Change the game name and year (used in titles, menu name)
2. **Point values** (`PTS_*`) — Update from the competition manual's scoring table
3. **`MOTIFS`** — Update the allowed gate-field values (the dropdown options for the field that activates scoring)
4. **`LEAVE_OPTIONS`** / **`BASE_OPTIONS`** — Update dropdown choices if the new game changes these
5. **`RAMP_REGEX`** / **`RAMP_MAX_CHARS`** — Update if the text-entry format changes; remove RAMP logic if the new game has no equivalent
6. **`G_RULES`** — Replace the full-text rules array with the new season's game rules (Section 11 of the competition manual)
7. **Column headers** — In `_buildRefereeSheet()`, update the `headers` array for new scoring elements
8. **Scoring formulas** — In `_buildRefereeSheet()`, update the formula comments and expressions in the DATA ROWS section
9. **`RC` column indices** — If columns are added, removed, or reordered, update the `RC` object
10. **FinalScores mapping** — In `_buildFinalScoresSheet()`, update `vlookupMap`, `elemCols`, `headers`, and category groups
11. **Help text** — In the DATA VALIDATION section, update the text shown on invalid input
12. **Documentation** — Update this README and the Scoring Guide
12. Run `buildAll()` to generate all sheets from scratch

## Authorization

All user-callable functions (`buildAll`, `confirmRebuild`, `randomizeTeamOrders`, `renameRefSheets`, `applyProtection`, `updateSheets`) require the current user's email to match one of the authorized SHA-256 hashes stored in `AUTHORIZED_HASHES`. The user's email is hashed at runtime via `Utilities.computeDigest()` and compared against the stored hashes — no cleartext emails appear in the source code.

Unauthorized users see an alert and the function exits immediately.

## Protection Model

### With Referee Emails (Full Isolation)
- **Sheet-level protection**: All cells locked except designated input ranges
- **Range-level protection**: Input cells restricted to specific referee + owner
- Each referee can ONLY edit input columns on their own sheet (D, J:L, M:Q, S:V, X)
- FinalScores column E (Override Name selection) restricted to owner only
- Config restricted to owner only (except team data and referee info)
- Config sheet hidden after protection is applied (unhide via tab right-click > Unhide)
- Unused referee sheets hidden automatically

### Without Emails (Advisory Mode)
- Formula cells protected with warnings
- No per-referee enforcement (anyone can edit any input cell)
- Config sheet hidden after protection is applied
- Unused referee sheets hidden automatically

## Named Referee Sheets

Referee sheets are named from Config row 2 (columns D-I). Initially "Referee 1" through "Referee 6", they are renamed when:
- **Randomize Team Orders** is run (auto-renames)
- **Rename Referee Sheets from Config** is run manually

Sheets that still have default "Referee N" names (unused referee slots) are hidden automatically after renaming or applying protection.

FinalScores uses `INDIRECT` formulas referencing Config names, so sheet name changes are reflected automatically.

### Conditional Formatting

**Referee sheets** (applied in priority order — first match wins for background):
1. **Red empty cells** on non-newest unfinished rows — highlights specific required input cells that are still empty
2. **Yellow row** for non-newest unfinished rows — when 2+ rows are unfinished (have team# but missing required fields), all except the newest (highest row number, presumed actively being scored) are highlighted yellow
3. **Red fouls** — cells in Minor/Major columns highlighted when > 0
4. **Orange unscored rows** — team# present but no data entered at all
5. **Zebra striping** on even rows

Required fields (12 total): MOTIF, Minor, Major, LEAVE, Auto CLS, Auto OVF, Auto RAMP, Tel CLS, Tel OVF, Tel DEPOT, Tel RAMP, BASE. "Not Shown" counts as a valid MOTIF value (not empty).

**FinalScores** (applied in priority order):
1. **Agree column formatting** — green (Yes), red (No), gray (N/A) for Refs Agree? (col F)
2. **Per-field disagreement red** — individual scoring cells (H-Z) highlighted red when refs disagree on that specific field (cell shows empty). Uses CHAR(8203) zero-width space marker with EXACT() detection. Only active when no Official Referee is selected and 2+ refs scored.
3. **Yellow row disagreement** — entire row highlighted yellow when Refs Agree? = "No" (lower priority than per-field red, so agreed fields in a disagreement row show yellow, disagreed fields show red)
4. **Orange missing Official Referee** — col E highlighted when team present but no selection made
5. **Zebra striping** on even rows

## Key Technical Details

### Non-Destructive Update

The **Update Sheets (Non-Destructive)** menu item (`updateSheets()`) rebuilds all referee sheets and FinalScores with the current template while preserving:
- Team orders (column A on referee sheets)
- All referee scoring inputs (MOTIF, Notes, Minor, Major, G Rules, LEAVE, Auto CLS/OVF/RAMP, Tel CLS/OVF/DEPOT/RAMP, BASE)
- Official Referee selections (column E on FinalScores)
- **Auto-appends missing teams**: Any teams in Config that are not already on a referee sheet are appended after the existing teams (no re-randomization needed)
- **Extends Config**: Validation and formatting are extended to cover the full MAX_TEAMS range

This is used to apply template changes (formulas, formatting, validation) to an existing spreadsheet without losing referee work. Sheet protection is NOT reapplied — run "Apply Sheet Protection" afterward if needed.

The function auto-detects the current sheet layout (old 23-column, v2 24-column with MOTIF at D, or current 24-column with MOTIF at M) when reading data to ensure compatibility during template transitions. Old-layout data is mapped to new column positions automatically; the G Rules column will be empty after migration from the original 23-column layout (it didn't exist).

### G Rules Multiselect
The G Rules column (L on referee sheets, O on FinalScores) allows referees to record which game rules were violated during a match. All 53 rules from the DECODE competition manual (Section 11) are available as dropdown options.

**How it works:**
- A hidden "Rules" sheet stores the full text of all 53 rules (used as the dropdown data source via `requireValueInRange`)
- The dropdown shows the full rule text (e.g., "G401 - Let the ROBOT do its thing. ...")
- When a referee selects a rule, the `onEdit()` trigger extracts the 4-character code (e.g., "G401") and appends it to the cell
- Selecting a rule that's already present toggles it off (removes it)
- The cell displays compact comma-separated codes: "G401, G418"
- G Rules are **excluded** from the Refs Agree? comparison (referees may cite different rules)

The `G_RULES` constant in the script contains the full text of all rules. To update for a new season, replace the array contents with the new season's game rules.

### Formula: PATTERN Match Calculation
```
=IF($A4="","",IF(OR($M4="",$M4="Not Shown"),0,IF(LEN(Q4)=0,0,
  SUMPRODUCT((MID(UPPER(Q4),SEQUENCE(MIN(LEN(Q4),9)),1)=
  MID(REPT($M4,3),SEQUENCE(MIN(LEN(Q4),9)),1))*1))))
```
Uses `SUMPRODUCT` with `MID`/`SEQUENCE` for character-by-character comparison. `REPT(M4,3)` repeats the 3-char MOTIF to cover all 9 RAMP positions. Returns 0 when MOTIF is blank or "Not Shown".

### Formula: INDIRECT Cross-Sheet References
FinalScores uses INDIRECT to reference referee sheets by name from Config:
```
=VLOOKUP($A4, INDIRECT("'"&Config!D$2&"'!$A:$X"), 6, FALSE)
```
This allows sheet tab names to change without breaking formulas. All score VLOOKUPs reference the effectiveRef helper column (AA) to avoid recomputing the effective referee in every formula.

### Formula: Agreement Check
Uses `FILTER`/`UNIQUE`/`ROWS` to compare values across all referees who scored a team. Each referee's value is fetched via INDIRECT VLOOKUP. All input columns except G Rules are checked: MOTIF, LEAVE, Auto CLASSIFIED/OVERFLOW/RAMP, TeleOp CLASSIFIED/OVERFLOW/DEPOT/RAMP, BASE, Minor Fouls, and Major Fouls. G Rules are excluded because referees may legitimately cite different rules. Blank fields on scored rows are treated as "(blank)" — distinct from an explicit 0 or any entered value. Text fields are normalized to uppercase for comparison.

### Formula: Scored By
Uses TEXTJOIN to list referee names who have scored a team. A referee is considered to have scored when any of MOTIF, LEAVE, or Auto CLASSIFIED is non-empty (composite check). Wrapped in IFERROR for edge cases.

### Blank vs Zero
A blank cell on a scored row means the referee forgot to enter a value. An explicit 0 means they intentionally scored zero:
- 0 vs 0 -> Agreement
- 0 vs blank -> Disagreement
- blank vs blank -> Agreement (both forgot)
- 5 vs blank -> Disagreement

### MOTIF Behavior
MOTIF is a regular input field — it does **not** gate formulas. Calculated columns populate whenever a Team # is present (even with blank inputs, showing 0 scores). PATTERN calculations return 0 when MOTIF is blank or "Not Shown".

The "Not Shown" option is for cases where the OBELISK/MOTIF was not visible in the match video.

### Scored Detection
A referee is considered to have "scored" a team when any of MOTIF, LEAVE, or Auto CLASSIFIED is non-empty. This composite check is used by:
- Progress counter (A1 on referee sheets)
- "Scored By" on FinalScores
- `_hasAnyScoringData()` for migration detection

### Required Fields
All input columns should be filled for a complete review:
- Dropdowns: Select a value (MOTIF, LEAVE, BASE)
- Numeric fields: Enter 0 explicitly if the count is zero
- RAMP Colors: Leave blank only if no artifacts on the RAMP
- G Rules: Optional (only when rules were violated)
- Notes: Optional
Blank required cells on non-newest unfinished rows are highlighted red.

### Sheet Identification
Each referee sheet stores its index in a note on cell A1 (`ref_index:N`). This allows the script to find sheets even after they've been renamed, if the Config name no longer matches a fallback.

### Randomization
Fisher-Yates shuffle ensures each referee sees teams in a different random order. A guard checks for existing scoring data before allowing re-randomization to prevent data corruption.

### Two-Phase Rename
Referee sheet renaming uses a two-phase approach: first rename all sheets to temporary names (`_temp_rename_N`), then rename to final desired names. This prevents collisions when referee names are swapped (e.g., swapping "Alice" and "Bob").

### Hidden Unnamed Sheets
After renaming or applying protection, referee sheets that still have default "Referee N" names are automatically hidden. This keeps the tab bar clean when fewer than 6 referees are used. Hidden sheets can be unhidden via tab right-click > Unhide.

### Merge / Freeze Constraint
Google Sheets does not allow frozen rows or columns to split a merged cell. All merges are split at frozen boundaries: referee sheet title uses A1:C1 (frozen) + D1:X1 (scrollable); FinalScores row 1 groups naturally align with the 3-column freeze (A1:C1 = "Teams").

### Sheet Tab Order
`buildAll()` moves FinalScores to the first tab position after creating all sheets, so the head referee's view is front and center. Tab order: FinalScores, Config, Rules (hidden), Referee 1-6.

### Batch Operations
All formula writes use batch `setFormulas()` and `setValues()` instead of individual cell writes for significantly better performance. Referee sheets write formulas in batch groups for non-contiguous calculated columns (F-J, R, W). FinalScores writes columns in batch operations covering A-D, F, G-Z, and the hidden helper column AA.

### Data Validation
- **Config**: Team numbers validated for uniqueness; referee names validated for uniqueness and restricted to characters safe for sheet names and INDIRECT references (letters, numbers, spaces, hyphens, periods, underscores; no leading spaces)
- **Referee sheets**: MOTIF, LEAVE, BASE use dropdown lists from configuration constants; numeric fields validated as whole numbers >= 0; RAMP Colors validated against RAMP_REGEX; G Rules uses `requireValueInRange` pointing to a hidden "Rules" sheet (avoids `requireValueInList` character limit) with `setAllowInvalid(true)` for multiselect combined values
- **FinalScores**: Official Referee dropdown populated from Config referee names range
- **Rules sheet**: Hidden sheet with all 53 G rule texts in column A, used as the data source for G Rules dropdown validation

## Adding / Removing Teams

### Adding Teams
1. Open the **Config** sheet (unhide it first if protected: right-click any tab > Unhide > Config)
2. Enter the new team number, name, and video link in the next empty row of columns A-C (starting at row 4)
3. Run **DECODE Scoring > Update Sheets (Non-Destructive)**

The update will automatically append the new team(s) to every referee sheet (after any existing teams) and rebuild FinalScores — without disturbing any scoring data already entered. There is no need to re-randomize.

If you want the new teams shuffled into the existing order instead of appended at the end, run **DECODE Scoring > Randomize Team Orders**. This will only work if no scoring data has been entered yet (the script guards against re-randomizing after scoring starts).

### Removing Teams
Teams cannot be removed through the menu — this is intentional to prevent accidental data loss. To remove a team:

1. Open the **Config** sheet
2. Delete the team's number, name, and video link from their row in columns A-C (clear the cells, don't delete the row)
3. On each referee sheet, clear the team number from column A for that team's row (the VLOOKUP-based Name and Video will clear automatically, and all calculated columns will blank out)
4. Optionally run **Update Sheets (Non-Destructive)** to refresh formulas

The team's row will remain as an empty row on each referee sheet. This is harmless — empty rows are skipped by FinalScores (no Team # = no output). Deleting entire rows from referee sheets is not recommended as it can shift formula ranges.

### Replacing a Team
To swap one team for another (e.g., a team drops out and a replacement joins):

1. On the **Config** sheet, overwrite the old team's number, name, and video link with the new team's info
2. On each referee sheet, clear any scoring data in the old team's row (the team number, Name, and Video will update automatically via Config lookup)
3. Referees can now score the new team in that same row

No script action is required — Config VLOOKUPs update the Name and Video columns on referee sheets automatically.

## Menu Items

| Menu Item | Function | Description |
|-----------|----------|-------------|
| Randomize Team Orders | `randomizeTeamOrders()` | Shuffles teams for each referee, renames sheets, hides unused sheets |
| Rename Referee Sheets from Config | `renameRefSheets()` | Renames sheet tabs to match Config names, hides unused sheets |
| Apply Sheet Protection | `applyProtection()` | Sets up sheet/range protections, hides Config and unused sheets |
| Update Sheets (Non-Destructive) | `updateSheets()` | Rebuilds layouts/formulas preserving scoring data |
| Rebuild All Sheets (DESTRUCTIVE) | `confirmRebuild()` -> `buildAll()` | Deletes and recreates everything |

## File Inventory

| File | Purpose |
|------|---------|
| `DECODE_Scoring_Spreadsheet.gs` | Main Google Apps Script (paste into Apps Script editor) |
| `DECODE_Scoring_Guide.md` | User guide for referees and head referee |
| `README.md` | This documentation (admin/developer reference) |
