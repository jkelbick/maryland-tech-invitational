# DECODE 2025-2026 Match Review Scoring Guide

This guide explains how to use the DECODE Match Review Scoring Spreadsheet. Each team plays a **solo match** (no opponent) that is video-recorded and scored independently by multiple referees. The head referee reviews all scores on the FinalScores sheet.

For game rules, scoring definitions, and field element descriptions, refer to the [DECODE Competition Manual](https://ftc-resources.firstinspires.org/ftc/archive/2026/game/manual) (Section 10.5 covers scoring). This guide covers **spreadsheet usage only** — two scoring rules differ from standard competition; see [Scoring Differences](#scoring-differences-from-official-ftc-rules) at the end.

**Contents**:
- [For Referees](#for-referees) — Finding your sheet, scoring a match, RAMP color entry, point values, cell highlights
- [For the Head Referee](#for-the-head-referee) — Using FinalScores, resolving disagreements
- [Scoring Differences from Official FTC Rules](#scoring-differences-from-official-ftc-rules)

---

## For Referees

### Finding Your Sheet

Open the spreadsheet and click the tab at the bottom with your name. Each referee has their own sheet with a unique randomized team order. You can only edit your own sheet's scoring cells.

The top 4 rows (title, instructions, point values, headers) and the first 2 columns (Team #, Team Name) stay visible as you scroll. Your progress counter (**Scored: X / Y**) is at the right end of row 2 — scroll right if you don't see it.

### Scoring a Match

For each team row, watch the match video (column C), then fill in the scoring columns left to right.

#### Step 1: Select the MOTIF (column D) — REQUIRED FIRST

Select the MOTIF displayed on the OBELISK: **GPP**, **PGP**, or **PPG**. This marks the row as scored and activates all score calculations. Nothing will calculate until you set this.

#### Step 2: Fill in ALL remaining columns

| Column | What to Enter | How to Fill It In |
|--------|--------------|-------------------|
| **LEAVE** (E) | Yes or No | Select from the dropdown. |
| **Auto CLASSIFIED** (F) | Whole number | Enter the count. Enter **0** if none. |
| **Auto OVERFLOW** (G) | Whole number | Enter the count. Enter **0** if none. |
| **Auto RAMP Colors** (H) | G and P characters | See [RAMP Color Entry](#ramp-color-entry) below. Leave blank if no artifacts on the RAMP. |
| **TeleOp CLASSIFIED** (I) | Whole number | Enter the count. Enter **0** if none. |
| **TeleOp OVERFLOW** (J) | Whole number | Enter the count. Enter **0** if none. |
| **TeleOp DEPOT** (K) | Whole number | Enter the count. Enter **0** if none. |
| **TeleOp RAMP Colors** (L) | G and P characters | See [RAMP Color Entry](#ramp-color-entry) below. Leave blank if no artifacts on the RAMP. |
| **BASE** (M) | None, Partial, or Full | Select from the dropdown. |
| **Minor Fouls** (N) | Whole number | Enter the count. Enter **0** if none. |
| **Major Fouls** (O) | Whole number | Enter the count. Enter **0** if none. |
| **Notes** (W) | Free text | Optional. Use for anything you want to flag for the head referee. |

**Important**: Enter **0** explicitly for zero counts. Do not leave numeric fields blank — a blank cell means you forgot to score it, not that the count is zero.

#### Step 3: Review

Columns P through V calculate automatically. Check that the total (column V) looks reasonable. If something seems off, double-check your MOTIF selection and RAMP color entries.

### RAMP Color Entry

The RAMP holds up to 9 artifacts. The OBELISK MOTIF repeats 3 times to define the expected color at each position:

| MOTIF | Full RAMP Pattern (GATE to SQUARE) |
|-------|------------------------------------|
| GPP | G P P G P P G P P |
| PGP | P G P P G P P G P |
| PPG | P P G P P G P P G |

**How to enter**: Type the actual artifact colors on the RAMP in order from **GATE to SQUARE**, using **G** for green and **P** for purple. Case does not matter.

**Example**: MOTIF is GPP. There are 5 artifacts on the RAMP. From GATE to SQUARE they are: green, purple, purple, green, green. You type: `GPPGG`

The spreadsheet compares each position against the expected pattern and counts matches:

| Position | You Entered | Expected (GPP repeated) | Match? |
|----------|------------|-------------------------|--------|
| 1 | G | G | Yes |
| 2 | P | P | Yes |
| 3 | P | P | Yes |
| 4 | G | G | Yes |
| 5 | G | P | No |

Result: 4 matches × 2 points = 8 PATTERN points.

If there are no artifacts on the RAMP, leave the cell blank.

### Point Values

These are shown in row 3 of your sheet for quick reference. See the [DECODE Competition Manual](https://ftc-resources.firstinspires.org/ftc/archive/2026/game/manual) Section 10.5 (Table 10-2) for the official point values.

| Scoring Element | Points |
|----------------|--------|
| LEAVE | 3 |
| CLASSIFIED | 3 per artifact |
| OVERFLOW | 1 per artifact |
| PATTERN match | 2 per matching position |
| DEPOT | 1 per artifact |
| BASE Partial | 5 |
| BASE Full | 10 |
| Minor Foul | −5 (deducted from score) |
| Major Foul | −15 (deducted from score) |

**Score cannot go below zero.** If foul deductions exceed the earned score, the total is 0.

### Cell Highlights

These highlights help you spot what needs attention. You can identify each state by the condition described, not just the color.

| State | Visual Indicator | What to Do |
|-------|-----------------|------------|
| Unscored team (team # present, no MOTIF) | Orange-highlighted row | Select a MOTIF in column D to begin scoring |
| Missing required field on a scored row | Pink cell | A required field is empty — fill it in (enter 0 for zero counts) |
| Foul recorded | Red-highlighted foul cell | Informational — confirms foul count is above zero |
| Auto-calculated field (columns B-C, P-V) | Light blue background | Do not edit — these are formulas |
| Input field (columns D-O, W) | Light yellow, green, peach, or purple background | This is where you enter data |

### Important Reminders

- Score each team independently — do not discuss scores with other referees until all scoring is complete.
- You cannot edit other referees' sheets or the FinalScores sheet. However, you may be able to see other referees' tabs — please do not look at them until all scoring is complete, to maintain independence.
- Do not edit columns A-C (team info) or P-V (calculated scores).
- If a cell rejects your input, read the pop-up message for guidance on what values are accepted.

---

## For the Head Referee

### Using FinalScores

The **FinalScores** sheet is the first tab in the spreadsheet. It aggregates scores from all referee sheets. The top 3 rows (category headers, point values, column names) and the first 3 columns (Team #, Name, Video) stay visible as you scroll.

#### Key Columns

| Column | Purpose |
|--------|---------|
| **Scored By** (D) | Lists which referees have scored this team |
| **Official Referee** (E) | Select which referee's scores to use as the official record |
| **Refs Agree?** (F) | Whether all referees who scored this team entered identical values |

#### Workflow

1. Wait for referees to finish scoring (check each referee's progress counter on their sheet).
2. For each team, check the **Refs Agree?** column:
   - **Yes** (green): All referees agree on every scoring element. Select any referee as Official.
   - **No** (red): At least one scoring element differs between referees. Review their individual sheets, determine the correct score, and select the most accurate referee as Official.
   - **N/A** (gray): Fewer than two referees scored this team — no comparison possible.
3. Select an **Official Referee** from the dropdown in column E for every team.
4. Columns G through X populate automatically from the selected referee's sheet.

#### What Triggers a Disagreement?

The agreement check compares **all 12 input fields** across every referee who scored the team:

- MOTIF, LEAVE, BASE (dropdowns)
- Auto/TeleOp CLASSIFIED, OVERFLOW (numbers)
- Auto/TeleOp RAMP Colors (text)
- TeleOp DEPOT (number)
- Minor Fouls, Major Fouls (numbers)

A blank cell and a zero are treated as **different values**. If one referee entered 0 and another left the field blank, that counts as a disagreement. Text comparisons are case-insensitive (e.g., "gpp" matches "GPP").

#### Resolving Disagreements

When **Refs Agree?** shows **No**:

1. Note which team has the disagreement.
2. Open each referee's sheet (tabs at the bottom) and find that team's row.
3. Compare the input values side-by-side to identify what differs.
4. Re-watch the match video if needed.
5. Select the referee whose scoring is most accurate as the Official Referee.

### FinalScores Columns G-X

All values in columns G through X come directly from the Official Referee's sheet. They update automatically when you change the Official Referee selection.

| Columns | Content |
|---------|---------|
| **G** | Final Score (total after foul deductions, minimum 0) |
| **H** | Score without Fouls |
| **I** | Auto Score |
| **J** | TeleOp Score |
| **K** | Foul Deduction |
| **L-M** | Minor Fouls, Major Fouls |
| **N** | LEAVE |
| **O-R** | Auto scoring elements (CLASSIFIED, OVERFLOW, RAMP Colors, PATTERN Count) |
| **S-X** | TeleOp scoring elements (CLASSIFIED, OVERFLOW, DEPOT, RAMP Colors, PATTERN Count, BASE) |

---

## Scoring Differences from Official FTC Rules

This spreadsheet is for **solo match reviews** (one team per match, no opponent). Two rules differ from standard FTC competition scoring:

1. **Foul points are subtracted from the team's own score** instead of being added to an opponent's score (there is no opponent).
2. **The 2-robot BASE bonus (10 points) and Ranking Points are not included** (single robot, non-competitive context).

For the full official scoring rules, see the [DECODE Competition Manual](https://ftc-resources.firstinspires.org/ftc/archive/2026/game/manual), Section 10.5.
