# FTC DECODE 2025-2026 Match Review Scoring Spreadsheet

## Overview

A Google Apps Script that builds a complete match review scoring system in Google Sheets for FTC DECODE 2025-2026. Designed for 2-6 referees to independently score solo matches (one match per team, no opponents), with automatic score calculation, disagreement detection, and a Final Scores sheet for the head referee.

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
   - Use **DECODE Scoring > Apply Sheet Protection** menu (this also hides the Config sheet)

## Sheet Structure

### Config Sheet
- **Column A** (row 4+): Team numbers
- **Column B** (row 4+): Team names
- **Column C** (row 4+): Video links (YouTube URLs)
- **Columns D-I** (row 2): Referee names (used as sheet tab names)
- **Columns D-I** (row 3): Referee emails (for protection)
- **Columns J-O** (row 4+): Randomized team orders per referee (auto-generated)

### Referee Sheets (named by referee)
Each referee gets an independent sheet named from Config (e.g., "Paul", "Jeff"). Sheets are initially named "Referee 1" through "Referee 6" and renamed when Randomize or Rename is run.

**Row 1**: Title bar (split merge: A1:B1 in frozen zone, C1:W1 in scrollable zone)
**Row 2**: Referee name (auto from Config), inline instructions, and progress counter ("Scored: X / Y")
**Row 3**: Point values as a quick reference (×3, ×1, ×2 ea, 5/10, etc.)
**Row 4**: Column headers (color-coded by section)
**Row 5+**: Data rows

| Column | Field | Type | Notes |
|--------|-------|------|-------|
| A | Team # | Auto-filled | From randomization |
| B | Team Name | **Auto (VLOOKUP)** | From Config |
| C | Video | **Auto (VLOOKUP)** | From Config |
| D | MOTIF | Dropdown (GPP/PGP/PPG) | **Required** - gates all formulas |
| E | LEAVE | Dropdown (Yes/No) | Did robot leave LAUNCH LINE during AUTO? |
| F | Auto CLASSIFIED | Whole number | Artifacts through SQUARE to RAMP during AUTO |
| G | Auto OVERFLOW | Whole number | Artifacts through SQUARE not to RAMP during AUTO |
| H | Auto RAMP Colors | Text (G/P chars) | Artifact colors on RAMP at end of AUTO, GATE→SQUARE order |
| I | TeleOp CLASSIFIED | Whole number | Same as F but during TELEOP |
| J | TeleOp OVERFLOW | Whole number | Same as G but during TELEOP |
| K | TeleOp DEPOT | Whole number | Artifacts over DEPOT tape at end of TELEOP |
| L | TeleOp RAMP Colors | Text (G/P chars) | Artifact colors on RAMP at end of TELEOP |
| M | BASE | Dropdown (None/Partial/Full) | Robot position on BASE TILE at end of TELEOP |
| N | Minor Fouls | Whole number | |
| O | Major Fouls | Whole number | |
| P | Auto PATTERN Count | **Calculated** | Character-by-character match of H vs REPT(MOTIF,3) |
| Q | Auto Score | **Calculated** | LEAVE(3) + CLS×3 + OVF×1 + PAT×2 |
| R | TeleOp PATTERN Count | **Calculated** | Character-by-character match of L vs REPT(MOTIF,3) |
| S | TeleOp Score | **Calculated** | CLS×3 + OVF×1 + DEPOT×1 + PAT×2 + BASE(0/5/10) |
| T | Foul Deduction | **Calculated** | Minor×5 + Major×15 |
| U | Score w/o Fouls | **Calculated** | Auto + TeleOp (before foul deduction) |
| V | TOTAL SCORE | **Calculated** | MAX(0, U - T) |
| W | Notes | Free text | |

Frozen rows: 4 (title, instructions, points, headers). Frozen columns: 2 (Team #, Team Name).

### FinalScores Sheet
Aggregation and score breakdown sheet for the head referee. **First tab** in the spreadsheet. Uses a 3-row header matching last year's layout.

**Row 1**: Merged category group headers (Teams | Referee | Total Scores | Fouls | Autonomous Period | TeleOp Period)
**Row 2**: Instructions (cols D-F, split merge at frozen boundary) + Point values per scoring element (cols L-X)
**Row 3**: Column names
**Row 4+**: Data

| Column | Field | Notes |
|--------|-------|-------|
| A | Team # | From Config |
| B | Team Name | From Config |
| C | Video | From Config |
| D | Scored By | Multiline list of referee names who scored this team |
| E | Official Referee | **Editable dropdown** - select which referee's score to display |
| F | Refs Agree? | Yes/No/N/A - do all scoring referees agree on every input element (including fouls)? |
| G | Final Score | From Official Referee (TOTAL SCORE) |
| H | Score w/o Fouls | From Official Referee |
| I | Auto Score | From Official Referee |
| J | TeleOp Score | From Official Referee |
| K | Foul Deduction | From Official Referee |
| L | Minor Fouls | From Official Referee |
| M | Major Fouls | From Official Referee |
| N | LEAVE | From Official Referee |
| O | Auto CLASSIFIED | From Official Referee |
| P | Auto OVERFLOW | From Official Referee |
| Q | Auto RAMP Colors | From Official Referee |
| R | Auto PATTERN Count | From Official Referee |
| S | TeleOp CLASSIFIED | From Official Referee |
| T | TeleOp OVERFLOW | From Official Referee |
| U | TeleOp DEPOT | From Official Referee |
| V | TeleOp RAMP Colors | From Official Referee |
| W | TeleOp PATTERN Count | From Official Referee |
| X | BASE | From Official Referee |

Frozen rows: 3 (category headers, point values, column names). Frozen columns: 3 (Team #, Name, Video).

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

## Authorization

All user-callable functions (`buildAll`, `confirmRebuild`, `randomizeTeamOrders`, `renameRefSheets`, `applyProtection`) require the current user's email to match one of the authorized SHA-256 hashes stored in `AUTHORIZED_HASHES`. The user's email is hashed at runtime via `Utilities.computeDigest()` and compared against the stored hashes — no cleartext emails appear in the source code.

Unauthorized users see an alert and the function exits immediately.

## Protection Model

### With Referee Emails (Full Isolation)
- **Sheet-level protection**: All cells locked except designated input ranges
- **Range-level protection**: Input cells restricted to specific referee + owner
- Each referee can ONLY edit columns D-O and W on their own sheet
- FinalScores column E (Override Name selection) restricted to owner only
- Config restricted to owner only (except team data and referee info)
- Config sheet hidden after protection is applied (unhide via tab right-click > Unhide)

### Without Emails (Advisory Mode)
- Formula cells protected with warnings
- No per-referee enforcement (anyone can edit any input cell)
- Config sheet hidden after protection is applied

## Named Referee Sheets

Referee sheets are named from Config row 2 (columns D-I). Initially "Referee 1" through "Referee 6", they are renamed when:
- **Randomize Team Orders** is run (auto-renames)
- **Rename Referee Sheets from Config** is run manually

FinalScores uses `INDIRECT` formulas referencing Config names, so sheet name changes are reflected automatically.

### Conditional Formatting
- **Referee sheets**: Pink highlight (#FFCCCC) for required fields left blank on scored rows; orange background (#FDE9D9) for unscored rows (team# but no MOTIF); red highlight for fouls > 0; zebra striping on even rows
- **FinalScores**: Green/red/gray for Refs Agree? (Yes/No/N/A); orange for missing Official Referee selection; zebra striping on even rows

## Key Technical Details

### Formula: PATTERN Match Calculation
```
=IF(OR($A5="",$D5=""),"",IF(LEN(H5)=0,0,
  SUMPRODUCT((MID(UPPER(H5),SEQUENCE(LEN(H5)),1)=
  MID(REPT(D5,3),SEQUENCE(LEN(H5)),1))*1)))
```
Uses `SUMPRODUCT` with `MID`/`SEQUENCE` for character-by-character comparison. `REPT(D5,3)` repeats the 3-char MOTIF to cover all 9 RAMP positions.

### Formula: INDIRECT Cross-Sheet References
FinalScores uses INDIRECT to reference referee sheets by name from Config:
```
=VLOOKUP($A5, INDIRECT("'"&Config!D$2&"'!$A:$W"), 22, FALSE)
```
This allows sheet tab names to change without breaking formulas.

### Formula: Agreement Check
Uses `FILTER`/`UNIQUE`/`ROWS` to compare values across all referees who scored a team. Each referee's value is fetched via INDIRECT VLOOKUP. All 12 input columns are checked: MOTIF, LEAVE, Auto CLASSIFIED/OVERFLOW/RAMP, TeleOp CLASSIFIED/OVERFLOW/DEPOT/RAMP, BASE, Minor Fouls, and Major Fouls. Blank fields on scored rows (MOTIF filled) are treated as "(blank)" - distinct from an explicit 0 or any entered value. Text fields are normalized to uppercase for comparison.

### Blank vs Zero
A blank cell on a scored row means the referee forgot to enter a value. An explicit 0 means they intentionally scored zero:
- 0 vs 0 -> Agreement
- 0 vs blank -> Disagreement
- blank vs blank -> Agreement (both forgot)
- 5 vs blank -> Disagreement

### MOTIF as Gate Field
All calculated columns are gated on `OR($A="",$D="")` - formulas only produce output when both Team # AND MOTIF are non-empty. This prevents unscored rows from showing 0 scores.

### Required Fields
Once MOTIF is selected (row is scored), ALL input columns should be filled:
- Dropdowns: Select a value (LEAVE=Yes/No, BASE=None/Partial/Full)
- Numeric fields: Enter 0 explicitly if the count is zero
- RAMP Colors: Leave blank only if no artifacts on the RAMP
- Notes: Optional
Blank cells on scored rows are highlighted pink and will trigger disagreements in FinalScores.

### Sheet Identification
Each referee sheet stores its index in a note on cell A1 (`ref_index:N`). This allows the script to find sheets even after they've been renamed, if the Config name no longer matches a fallback.

### Randomization
Fisher-Yates shuffle ensures each referee sees teams in a different random order. A guard checks for existing MOTIF data before allowing re-randomization to prevent data corruption.

### Two-Phase Rename
Referee sheet renaming uses a two-phase approach: first rename all sheets to temporary names (`_temp_rename_N`), then rename to final desired names. This prevents collisions when referee names are swapped (e.g., swapping "Alice" and "Bob").

### Merge / Freeze Constraint
Google Sheets does not allow frozen rows or columns to split a merged cell. All merges are split at frozen boundaries: referee sheet title uses A1:B1 (frozen) + C1:W1 (scrollable); FinalScores row 2 uses A2:C2 (frozen) + D2:F2 (scrollable). FinalScores row 1 groups naturally align with the 3-column freeze (A1:C1 = "Teams").

### Sheet Tab Order
`buildAll()` moves FinalScores to the first tab position after creating all sheets, so the head referee's view is front and center. Tab order: FinalScores, Config, Referee 1-6.

### Batch Operations
All formula writes use batch `setFormulas()` and `setValues()` instead of individual cell writes for significantly better performance. Referee sheets write formulas for columns B, C, and P-V in three batch operations. FinalScores writes columns A-D, F, and G-X in three batch operations.

### Data Validation
- **Config**: Team numbers validated for uniqueness; referee names validated for uniqueness and restricted to characters safe for sheet names and INDIRECT references (letters, numbers, spaces, hyphens, periods, underscores; no leading spaces)
- **Referee sheets**: MOTIF, LEAVE, BASE use dropdown lists; numeric fields validated as whole numbers >= 0; RAMP Colors validated as 1-9 G/P characters via regex
- **FinalScores**: Official Referee dropdown populated from Config referee names range

## Menu Items

| Menu Item | Function | Description |
|-----------|----------|-------------|
| Randomize Team Orders | `randomizeTeamOrders()` | Shuffles teams for each referee, renames sheets |
| Rename Referee Sheets from Config | `renameRefSheets()` | Renames sheet tabs to match Config names |
| Apply Sheet Protection | `applyProtection()` | Sets up sheet/range protections, hides Config |
| Rebuild All Sheets (DESTRUCTIVE) | `confirmRebuild()` -> `buildAll()` | Deletes and recreates everything |

## File Inventory

| File | Purpose |
|------|---------|
| `DECODE_Scoring_Spreadsheet.gs` | Main Google Apps Script (paste into Apps Script editor) |
| `DECODE_Scoring_Guide.md` | User guide for referees and head referee |
| `README.md` | This documentation (admin/developer reference) |
