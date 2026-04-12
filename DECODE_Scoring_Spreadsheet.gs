/**
 * FTC Match Review Scoring Spreadsheet
 * Google Apps Script to build all sheets, formulas, validation, formatting, and protection.
 *
 * USAGE:
 *   1. Open your Google Sheet
 *   2. Extensions > Apps Script
 *   3. Paste this entire file
 *   4. Run buildAll()
 *   5. Grant permissions when prompted
 *
 * This creates:
 *   - "Config" sheet: team list (number/name/video), referee names/emails, randomized orders
 *   - "Rules" sheet (hidden): G rules text for dropdown validation
 *   - Named referee sheets (one per referee, named from Config)
 *   - "FinalScores" sheet: aggregation, comparison, disagreement highlighting, official score selection
 *
 * PROTECTION MODEL:
 *   Sheet-level protection locks formulas/headers/team numbers (owner-only).
 *   Range-level protection on input cells restricts each referee to their own sheet.
 *   Requires referee emails in Config row 3 for enforcement; advisory-only without emails.
 *   NOTE: Google Sheets protection prevents editing via the UI only. Users with Editor access
 *   can still VIEW all sheets including other referees' scores. If scoring independence is
 *   critical, use separate spreadsheets per referee.
 *
 * SCORING MODEL:
 *   - CLASSIFIED/OVERFLOW: Counted throughout the match as artifacts pass through the SQUARE. No cap.
 *   - PATTERN: Assessed at end of AUTO and end of TELEOP based on RAMP snapshot. Referee enters
 *     artifact colors on the RAMP in order (G/P), and the spreadsheet auto-calculates matches
 *     against the MOTIF. Max 9 characters (RAMP capacity). PATTERN = 0 when MOTIF blank or "Not Shown".
 *   - Fouls: Subtracted from team score (deviation from official rules for solo match context).
 *   - 2-robot BASE bonus (10 pts) and Ranking Points: Excluded (solo match, single robot).
 *
 * HOW TO UPDATE FOR A NEW SEASON:
 * ─────────────────────────────────────────────────────────
 *  1. GAME_NAME / SEASON — Update the game name and year.
 *  2. Point values (PTS_*) — See the competition manual scoring table.
 *  3. MOTIFS — The allowed gate-field values (e.g., OBELISK patterns).
 *  4. LEAVE_OPTIONS / BASE_OPTIONS — Dropdown choices for Yes/No and base status.
 *  5. RAMP_REGEX / RAMP_MAX_CHARS — Update if the text-entry format changes.
 *     Remove RAMP logic entirely if the new game has no equivalent.
 *  6. G_RULES — Update from the competition manual Section 11.
 *  7. Column headers — In _buildRefereeSheet(), update the 'headers' array.
 *  8. Scoring formulas — In _buildRefereeSheet(), update the formulas in the
 *     DATA ROWS section for Auto Score, TeleOp Score, etc.
 *  9. RC column indices — If columns are added/removed/reordered, update the
 *     RC object and layout constants as needed.
 * 10. FinalScores mapping — In _buildFinalScoresSheet(), update vlookupMap,
 *     elemCols, headers, and category groups.
 * 11. Help text — In the DATA VALIDATION section of _buildRefereeSheet(),
 *     update the text shown to referees on invalid input.
 * 12. Documentation — Update README.md and the Scoring Guide.
 * 13. Run buildAll() to generate all sheets from scratch.
 *
 * After updating an already-in-use spreadsheet, run "Update Sheets" from the
 * menu instead of buildAll — this preserves existing referee scoring data.
 * ─────────────────────────────────────────────────────────
 */

// ============================================================
// GAME CONFIGURATION — Update these values each season
// ============================================================

// --- Season identity ---
const GAME_NAME = "DECODE";
const SEASON = "2025-2026";

// --- General ---
const NUM_REFEREES = 6;
const MAX_TEAMS = 500;
const FORMULA_BUFFER = 50;  // Extra formula rows beyond current team count for headroom

// --- Scoring point values (see competition manual scoring table) ---
const PTS_LEAVE = 3;
const PTS_CLASSIFIED = 3;
const PTS_OVERFLOW = 1;
const PTS_PATTERN = 2;
const PTS_DEPOT = 1;
const PTS_BASE_PARTIAL = 5;
const PTS_BASE_FULL = 10;
const PTS_MINOR_FOUL = 5;
const PTS_MAJOR_FOUL = 15;

// --- Dropdown option values ---
const MOTIFS = ["GPP", "PGP", "PPG", "Not Shown"];
const LEAVE_OPTIONS = ["Yes", "No"];
const BASE_OPTIONS = ["None", "Partial", "Full"];

// --- Text-entry validation (RAMP color entry) ---
const RAMP_REGEX = "^[GP]{1,9}$";
const RAMP_MAX_CHARS = 9;

// --- G Rules (from DECODE Competition Manual Section 11) ---
// Each entry: "G{num} - {short title}" (truncated to avoid row-height expansion in dropdown).
// Used for dropdown validation on referee sheets and FinalScores.
// The onEdit trigger extracts the first 4 characters as the rule code.
const G_RULES = [
  'G101 - Humans, stay off the FIELD during the MATCH',
  'G102 - Be careful when interacting with ARENA elements',
  'G201 - Be a good person',
  'G202 - DRIVE TEAM Interactions',
  'G203 - Asking other teams to throw a MATCH',
  'G204 - Letting someone coerce you into throwing a MATCH',
  'G205 - Throwing your own MATCH is bad',
  'G206 - Don\'t violate rules for RPs',
  'G207 - Do not abuse ARENA access',
  'G208 - Show up to your MATCHES',
  'G209 - Keep your ROBOT together',
  'G210 - Do not expect to gain by doing others harm',
  'G211 - Egregious or exceptional violations',
  'G212 - All teams can play',
  'G301 - Be prompt',
  'G302 - Limit what you bring to the FIELD',
  'G303 - ROBOTS on the FIELD must come ready to play a MATCH',
  'G304 - ROBOTS must be set up correctly on the FIELD',
  'G305 - Teams must select an OpMode',
  'G401 - Let the ROBOT do its thing',
  'G402 - No AUTO opponent interference',
  'G403 - ROBOTS are motionless between AUTO and TELEOP',
  'G404 - ROBOTS are motionless at the end of TELEOP',
  'G405 - ROBOTS use SCORING ELEMENTS as directed',
  'G406 - Keep SCORING ELEMENTS in bounds',
  'G407 - Do not damage SCORING ELEMENTS',
  'G408 - No more than 3 at a time',
  'G409 - ROBOTS must be under control',
  'G410 - ROBOTS must stop when instructed',
  'G411 - ROBOTS must be identifiable',
  'G412 - Don\'t damage the FIELD',
  'G413 - Watch your ARENA interaction',
  'G414 - ROBOTS have horizontal expansion limits',
  'G415 - ROBOTS have vertical expansion limits, with exceptions',
  'G416 - LAUNCHING in the LAUNCH ZONE only',
  'G417 - ROBOTS only operate GATES as directed',
  'G418 - ROBOTS may not meddle with ARTIFACTS on RAMPS',
  'G419 - ROBOTS LAUNCH into their own GOAL',
  'G420 - This is not combat robotics',
  'G421 - Do not tip or entangle',
  'G422 - There is a 3-count on PINS',
  'G423 - Do not use strategies intended to shut down major parts of gameplay',
  'G424 - GATE ZONE is OFF LIMITS',
  'G425 - Keep out of opponent\'s SECRET TUNNEL',
  'G426 - LOADING ZONE protection',
  'G427 - BASE ZONE protection',
  'G428 - No wandering',
  'G429 - DRIVE COACHES and other teams: hands off the controls',
  'G430 - DRIVE COACHES, SCORING ELEMENTS are off limits',
  'G431 - DRIVE TEAMS, watch your reach',
  'G432 - Humans, only meddle with ARTIFACTS in the LOADING ZONE',
  'G433 - Humans may only enter SCORING ELEMENTS',
  'G434 - The ALLIANCE AREA has a storage limit'
];

// --- Layout constants ---
// Referee sheets: Row 1=Title, Row 2=Point values (hidden), Row 3=Headers, Row 4+=Data
const REF_DATA_START = 4;
const REF_DATA_END = MAX_TEAMS + 3;
// FinalScores: Row 1=Category headers, Row 2=Point values (hidden), Row 3=Headers, Row 4+=Data
const FS_DATA_START = 4;
const FS_DATA_END = MAX_TEAMS + 3;

// Referee sheet column layout (A-V, 22 columns):
//   A=Team#(auto)  B=Name(auto)  C=Video(auto)
//   D=Notes
//   E=TOTAL(calc)  F=Score w/o Fouls(calc)  G=Auto Score(calc)  H=TeleOp Score(calc)  I=Foul Deduction(calc)
//   J=Minor Fouls  K=Major Fouls  L=G Rules(multiselect)
//   M=MOTIF  N=LEAVE  O=Auto CLS  P=Auto OVF  Q=Auto RAMP Colors
//   R=Tel CLS  S=Tel OVF  T=Tel DEPOT  U=Tel RAMP Colors
//   V=BASE
const RC = {
  TEAM: 1, NAME: 2, VIDEO: 3, NOTES: 4,
  TOTAL: 5, SCORE_NO_FOULS: 6, AUTO_SCORE: 7, TEL_SCORE: 8, FOUL_DED: 9,
  MINOR: 10, MAJOR: 11, G_RULES: 12,
  MOTIF: 13, LEAVE: 14, AUTO_CLS: 15, AUTO_OVF: 16, AUTO_RAMP: 17,
  TEL_CLS: 18, TEL_OVF: 19, TEL_DEPOT: 20, TEL_RAMP: 21,
  BASE: 22
};

// ============================================================
// AUTHORIZATION
// ============================================================
// SHA-256 hashes of authorized emails (lowercase)
const AUTHORIZED_HASHES = [
  "c05ddb09d44266bc0b82e5b8322d000861b378638de771a24cce341c3859d7cc",
  "2ee4d1fd7155caaddb53b63a305413733616070f7f4802b6a13792167a4f1d88"
];

/** Constant-time string comparison to prevent timing attacks on hash matching. */
function _constantTimeEquals(a, b) {
  let len = Math.max(a.length, b.length);
  let result = a.length ^ b.length; // non-zero if lengths differ
  for (let i = 0; i < len; i++) {
    result |= (i < a.length ? a.charCodeAt(i) : 0) ^ (i < b.length ? b.charCodeAt(i) : 0);
  }
  return result === 0;
}

function _hashEmail(email) {
  let raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email);
  let hex = "";
  for (let i = 0; i < raw.length; i++) {
    let b = (raw[i] + 256) % 256; // convert signed byte to unsigned
    hex += ("0" + b.toString(16)).slice(-2);
  }
  return hex;
}

function checkAuthorization() {
  let userEmail = (Session.getEffectiveUser().getEmail() || "").toLowerCase();
  if (!userEmail) {
    try {
      SpreadsheetApp.getUi().alert(
        "Authorization Error",
        "Could not determine your identity. Please ensure you have authorized this script.\n" +
        "Try: Extensions > Apps Script > Run any function > Grant permissions when prompted.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch(e) {
      Logger.log("Authorization check failed: empty email");
    }
    return false;
  }
  let emailHash = _hashEmail(userEmail);
  for (let i = 0; i < AUTHORIZED_HASHES.length; i++) {
    if (_constantTimeEquals(emailHash, AUTHORIZED_HASHES[i])) return true;
  }
  try {
    SpreadsheetApp.getUi().alert(
      "Unauthorized",
      "You are not authorized to run this script.\nContact the spreadsheet owner for access.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) {
    Logger.log("Unauthorized user attempted access");
  }
  return false;
}

// ============================================================
// HELPERS
// ============================================================

/** Converts 1-based column number to column letter(s). 1→A, 26→Z, 27→AA, etc. */
function _colLetter(n) {
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

/** Writes a 1D array as a single column starting at startRow, colNum. */
function _writeColumn(sheet, startRow, colNum, values1D) {
  let data = [];
  for (let i = 0; i < values1D.length; i++) {
    data.push([values1D[i]]);
  }
  sheet.getRange(startRow, colNum, data.length, 1).setValues(data);
}

/**
 * Writes multiple contiguous columns in a single API call.
 * @param {Sheet} sheet
 * @param {number} startRow - First row to write (1-based)
 * @param {number} startCol - First column to write (1-based)
 * @param {Array<Array>} columns - Array of 1D arrays, one per column (must all be same length)
 */
function _writeColumns(sheet, startRow, startCol, columns) {
  if (columns.length === 0) return;
  let numRows = columns[0].length;
  let numCols = columns.length;
  let data = [];
  for (let i = 0; i < numRows; i++) {
    let row = [];
    for (let c = 0; c < numCols; c++) {
      row.push(columns[c][i]);
    }
    data.push(row);
  }
  sheet.getRange(startRow, startCol, numRows, numCols).setValues(data);
}

/** Extracts a 0-indexed column from a 2D values array into a 1D array. */
function _extractColumn(data2D, colIndex0) {
  let result = [];
  for (let i = 0; i < data2D.length; i++) {
    result.push(data2D[i][colIndex0]);
  }
  return result;
}

/**
 * Detect layout version by checking row 3 col 5 header.
 * New layout: col 5 = "Notes". Old layout: col 5 = "LEAVE".
 */
function _detectLayoutVersion(sheet) {
  let val4 = sheet.getRange(3, 4).getValue().toString().toLowerCase();
  if (val4.indexOf("notes") !== -1) {
    // Could be v3 (24-col MOTIF@M with PAT columns) or new (22-col without PAT)
    let val18 = sheet.getRange(3, 18).getValue().toString().toLowerCase();
    if (val18.indexOf("pat") !== -1) return "v3"; // 24-col MOTIF at M, with Auto PAT / Tel PAT
    return "new"; // current 22-col: MOTIF at M, no PAT columns
  }
  let val5 = sheet.getRange(3, 5).getValue().toString().toLowerCase();
  if (val5.indexOf("notes") !== -1) return "v2"; // prior 24-col: MOTIF at col D, Notes at col E
  return "old"; // original 23-col layout
}

/** Returns the Config column letter for a given referee number (1-6 -> D-I). */
function _refConfigCol(refNum) {
  return String.fromCharCode(67 + refNum);
}

function getRefSheetName(config, refNum) {
  if (!config) return "Referee " + refNum;
  let name = config.getRange(_refConfigCol(refNum) + "2").getValue();
  return (name && name.toString().trim() !== "") ? name.toString().trim() : "Referee " + refNum;
}

function findRefSheet(ss, config, refNum, noteMap) {
  let sheet;
  if (config) {
    let name = getRefSheetName(config, refNum);
    sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  sheet = ss.getSheetByName("Referee " + refNum);
  if (sheet) return sheet;
  sheet = ss.getSheetByName("Referee" + refNum);
  if (sheet) return sheet;
  // Fallback: check ref_index note on A1 (use pre-built noteMap if available)
  if (noteMap) return noteMap["ref_index:" + refNum] || null;
  let allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    try {
      if (allSheets[i].getRange("A1").getNote() === "ref_index:" + refNum) return allSheets[i];
    } catch(e) {}
  }
  return null;
}

/** Pre-builds a map of A1 notes to sheets for batch findRefSheet lookups. */
function _buildNoteMap(ss) {
  let map = {};
  let sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    try { let n = sheets[i].getRange("A1").getNote(); if (n) map[n] = sheets[i]; } catch(e) {}
  }
  return map;
}

/** Remove all editors except the specified allowed emails from a protection object. */
function _restrictEditors(protection, allowedEmails) {
  let allowed = {};
  for (let i = 0; i < allowedEmails.length; i++) {
    allowed[allowedEmails[i].toLowerCase()] = true;
  }
  protection.getEditors().forEach(function(editor) {
    if (!allowed[editor.getEmail().toLowerCase()]) {
      try { protection.removeEditor(editor); } catch(e) {}
    }
  });
}

/** Hide referee sheets that still have default "Referee N" names (unused slots). */
function _hideUnnamedRefSheets(ss, config) {
  if (!config) return;
  let noteMap = _buildNoteMap(ss);
  for (let r = 1; r <= NUM_REFEREES; r++) {
    let configName = config.getRange(_refConfigCol(r) + "2").getValue();
    let isDefault = !configName || configName.toString().trim() === "" ||
      configName.toString().trim() === "Referee " + r;
    if (isDefault) {
      let sheet = findRefSheet(ss, config, r, noteMap);
      if (sheet) {
        try { sheet.hideSheet(); } catch(e) {
          Logger.log("Could not hide sheet for Referee " + r + ": " + e);
        }
      }
    }
  }
}

/**
 * Detect whether a referee sheet uses the old layout (data at row 5) or new (data at row 4).
 * Old layout had 4 frozen rows (title, instructions, points, headers).
 * New layout has 3 frozen rows (title, points, headers).
 */
function _detectRefDataStart(sheet) {
  let frozen = sheet.getFrozenRows();
  return frozen >= 4 ? 5 : 4;
}

// ============================================================
// CUSTOM MENU
// ============================================================
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu(GAME_NAME + " Scoring")
      .addItem("Randomize Team Orders", "randomizeTeamOrders")
      .addItem("Reorder Active Sheet to Config Order", "reorderToConfigOrder")
      .addItem("Rename Referee Sheets from Config", "renameRefSheets")
      .addSeparator()
      .addItem("Apply Sheet Protection", "applyProtection")
      .addSeparator()
      .addItem("Update Sheets (Non-Destructive)", "updateSheets")
      .addItem("Rebuild All Sheets (DESTRUCTIVE)", "confirmRebuild")
      .addToUi();
  } catch(e) {
    Logger.log("onOpen: UI not available (script editor context): " + e);
  }
}

// ============================================================
// MAIN ENTRY POINT
// ============================================================
function confirmRebuild() {
  if (!checkAuthorization()) return;
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("confirmRebuild: UI not available. Rebuild requires interactive confirmation.");
    return;
  }
  let response = ui.alert(
    "Rebuild All Sheets",
    "This will DELETE and recreate ALL sheets (Config, referee sheets, FinalScores).\n" +
    "All existing data will be LOST.\n\nAre you sure?",
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) buildAll();
}

function buildAll() {
  if (!checkAuthorization()) return;
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  // Reuse leftover temp sheet from a previous failed build, or create new
  let temp = ss.getSheetByName("_temp_build") || ss.insertSheet("_temp_build");
  let allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName() !== "_temp_build") {
      try { ss.deleteSheet(allSheets[i]); } catch(e) {}
    }
  }

  _buildConfigSheet(ss);
  _buildRulesSheet(ss);
  let config = ss.getSheetByName("Config");
  for (let r = 1; r <= NUM_REFEREES; r++) {
    _buildRefereeSheet(ss, config, r);
  }
  _buildFinalScoresSheet(ss);

  try { ss.deleteSheet(temp); } catch(e) {}

  // Move FinalScores to first tab position
  let finalSheet = ss.getSheetByName("FinalScores");
  if (finalSheet) {
    ss.setActiveSheet(finalSheet);
    ss.moveActiveSheet(1);
  }

  if (config) config.activate();
  SpreadsheetApp.flush();

  try {
    SpreadsheetApp.getUi().alert(
      "Setup Complete",
      GAME_NAME + " Scoring Spreadsheet built successfully!\n\n" +
      "Next steps:\n" +
      "1. Enter team data in Config columns A-C (row 4+): number, name, video link\n" +
      "2. Enter referee names in Config row 2 (columns D-I)\n" +
      "3. Enter referee emails in Config row 3 (row is hidden \u2014 right-click row 2/4 border to unhide)\n" +
      "4. " + GAME_NAME + " Scoring > Randomize Team Orders\n" +
      "5. " + GAME_NAME + " Scoring > Apply Sheet Protection (this also hides Config)\n" +
      "6. Referees score on their individual sheets\n" +
      "7. Use FinalScores to compare scores and select an Official Referee per team",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) {
    Logger.log(GAME_NAME + " Scoring Spreadsheet built successfully.");
  }
}

// ============================================================
// UPDATE SHEETS (non-destructive — preserves scoring data)
// ============================================================
/**
 * Rebuilds all referee sheets, Rules sheet, and FinalScores with the current
 * template while preserving referee scoring inputs and Official Referee selections.
 * Layout-aware: detects old (23-col) vs new (24-col) layout and migrates data.
 */
function updateSheets() {
  if (!checkAuthorization()) return;
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("updateSheets must be run from the " + GAME_NAME + " Scoring menu.");
    return;
  }

  let response = ui.alert(
    "Update All Sheets",
    "This updates all sheet layouts, formulas, formatting, and validation " +
    "to the current template without erasing scoring data.\n\n" +
    "Preserved: team orders, referee scoring inputs, Official Referee selections.\n" +
    "Updated: formulas, formatting, validation, headers, conditional formatting.\n\n" +
    "Continue?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let config = ss.getSheetByName("Config");
  if (!config) {
    ui.alert("Error", "Config sheet not found. Run buildAll() first.", ui.ButtonSet.OK);
    return;
  }

  // Layout column indices for migration from older layouts
  let OLD_RC = {
    TEAM: 1, NAME: 2, VIDEO: 3, MOTIF: 4, LEAVE: 5,
    AUTO_CLS: 6, AUTO_OVF: 7, AUTO_RAMP: 8,
    TEL_CLS: 9, TEL_OVF: 10, TEL_DEPOT: 11, TEL_RAMP: 12,
    BASE: 13, MINOR: 14, MAJOR: 15,
    AUTO_PAT: 16, AUTO_SCORE: 17, TEL_PAT: 18, TEL_SCORE: 19,
    FOUL_DED: 20, SCORE_NO_FOULS: 21, TOTAL: 22, NOTES: 23
  };
  let V2_RC = {
    TEAM: 1, NAME: 2, VIDEO: 3, MOTIF: 4, NOTES: 5,
    TOTAL: 6, SCORE_NO_FOULS: 7, AUTO_SCORE: 8, TEL_SCORE: 9, FOUL_DED: 10,
    MINOR: 11, MAJOR: 12, G_RULES: 13,
    LEAVE: 14, AUTO_CLS: 15, AUTO_OVF: 16, AUTO_RAMP: 17, AUTO_PAT: 18,
    TEL_CLS: 19, TEL_OVF: 20, TEL_DEPOT: 21, TEL_RAMP: 22, TEL_PAT: 23,
    BASE: 24
  };
  let V3_RC = {
    TEAM: 1, NAME: 2, VIDEO: 3, NOTES: 4,
    TOTAL: 5, SCORE_NO_FOULS: 6, AUTO_SCORE: 7, TEL_SCORE: 8, FOUL_DED: 9,
    MINOR: 10, MAJOR: 11, G_RULES: 12,
    MOTIF: 13, LEAVE: 14, AUTO_CLS: 15, AUTO_OVF: 16, AUTO_RAMP: 17, AUTO_PAT: 18,
    TEL_CLS: 19, TEL_OVF: 20, TEL_DEPOT: 21, TEL_RAMP: 22, TEL_PAT: 23,
    BASE: 24
  };

  // --- Save referee scoring data (layout-aware per-field extraction) ---
  let savedRefData = {};
  let noteMap = _buildNoteMap(ss);
  for (let r = 1; r <= NUM_REFEREES; r++) {
    let sheet = findRefSheet(ss, config, r, noteMap);
    if (!sheet) continue;

    let dataStart = _detectRefDataStart(sheet);
    let dataEnd = dataStart + MAX_TEAMS - 1;
    let layoutVer = _detectLayoutVersion(sheet);
    let src = (layoutVer === "old") ? OLD_RC : (layoutVer === "v2") ? V2_RC : (layoutVer === "v3") ? V3_RC : RC;
    let numCols = (layoutVer === "old") ? 23 : (layoutVer === "v2" || layoutVer === "v3") ? 24 : 22;

    let readRows = Math.min(MAX_TEAMS, Math.max(0, sheet.getMaxRows() - dataStart + 1));
    let allData = readRows > 0 ? sheet.getRange(dataStart, 1, readRows, numCols).getValues() : [];
    while (allData.length < MAX_TEAMS) allData.push(new Array(numCols).fill(""));

    savedRefData[r] = {
      teams:    _extractColumn(allData, src.TEAM - 1),
      motif:    _extractColumn(allData, src.MOTIF - 1),
      notes:    _extractColumn(allData, src.NOTES - 1),
      leave:    _extractColumn(allData, src.LEAVE - 1),
      autoCls:  _extractColumn(allData, src.AUTO_CLS - 1),
      autoOvf:  _extractColumn(allData, src.AUTO_OVF - 1),
      autoRamp: _extractColumn(allData, src.AUTO_RAMP - 1),
      telCls:   _extractColumn(allData, src.TEL_CLS - 1),
      telOvf:   _extractColumn(allData, src.TEL_OVF - 1),
      telDepot: _extractColumn(allData, src.TEL_DEPOT - 1),
      telRamp:  _extractColumn(allData, src.TEL_RAMP - 1),
      base:     _extractColumn(allData, src.BASE - 1),
      minor:    _extractColumn(allData, src.MINOR - 1),
      major:    _extractColumn(allData, src.MAJOR - 1),
      gRules:   (layoutVer !== "old" && src.G_RULES) ? _extractColumn(allData, src.G_RULES - 1) : new Array(MAX_TEAMS).fill("")
    };
  }

  // --- Save FinalScores Official Referee selections ---
  let fsSheet = ss.getSheetByName("FinalScores");
  let savedOfficialRefs = null;
  if (fsSheet) {
    savedOfficialRefs = fsSheet.getRange("E" + FS_DATA_START + ":E" + FS_DATA_END).getValues();
  }

  // --- Create temp sheet to avoid last-sheet deletion errors ---
  let temp = ss.getSheetByName("_temp_update") || ss.insertSheet("_temp_update");

  // --- Delete old referee sheets ---
  for (let r = 1; r <= NUM_REFEREES; r++) {
    let sheet = findRefSheet(ss, config, r);
    if (sheet) {
      try { ss.deleteSheet(sheet); } catch(e) {
        Logger.log("Could not delete sheet for referee " + r + ": " + e);
      }
    }
  }

  // --- Delete old FinalScores and Rules ---
  fsSheet = ss.getSheetByName("FinalScores");
  if (fsSheet) { try { ss.deleteSheet(fsSheet); } catch(e) {} }
  let rulesSheet = ss.getSheetByName("Rules");
  if (rulesSheet) { try { ss.deleteSheet(rulesSheet); } catch(e) {} }

  // --- Rebuild Rules sheet ---
  _buildRulesSheet(ss);

  // --- Compute formulaRows for optimized rebuild ---
  // Must cover both Config teams AND any saved referee data (in case Config shrank)
  let configTeamRange = config.getRange("A4:A" + (MAX_TEAMS + 3)).getValues();
  let configTeamCount = 0;
  for (let i = 0; i < configTeamRange.length; i++) {
    if (configTeamRange[i][0] !== "" && configTeamRange[i][0] !== null) configTeamCount++;
  }
  let maxSavedTeamCount = 0;
  for (let r = 1; r <= NUM_REFEREES; r++) {
    if (savedRefData[r]) {
      let cnt = 0;
      for (let i = 0; i < savedRefData[r].teams.length; i++) {
        if (savedRefData[r].teams[i] !== "" && savedRefData[r].teams[i] !== null) cnt++;
      }
      if (cnt > maxSavedTeamCount) maxSavedTeamCount = cnt;
    }
  }
  let formulaRows = Math.min(MAX_TEAMS, Math.max(configTeamCount, maxSavedTeamCount) + FORMULA_BUFFER);

  // --- Rebuild all referee sheets and restore data ---
  // Note: simple triggers (onEdit) do not fire on programmatic setValue() calls,
  // so no guard is needed during data restoration.
  for (let r = 1; r <= NUM_REFEREES; r++) {
    _buildRefereeSheet(ss, config, r, formulaRows);

    if (savedRefData[r]) {
      let sheet = findRefSheet(ss, config, r);
      if (sheet) {
        let d = savedRefData[r];
        // Batch restore: 6 contiguous group writes instead of 15 individual calls
        _writeColumn(sheet, REF_DATA_START, RC.TEAM, d.teams);           // col 1
        _writeColumn(sheet, REF_DATA_START, RC.NOTES, d.notes);          // col 4
        _writeColumns(sheet, REF_DATA_START, RC.MINOR, [d.minor, d.major, d.gRules]); // cols 10-12
        _writeColumns(sheet, REF_DATA_START, RC.MOTIF, [d.motif, d.leave, d.autoCls, d.autoOvf, d.autoRamp]); // cols 13-17
        _writeColumns(sheet, REF_DATA_START, RC.TEL_CLS, [d.telCls, d.telOvf, d.telDepot, d.telRamp]); // cols 18-21
        _writeColumn(sheet, REF_DATA_START, RC.BASE, d.base);            // col 22
      }
    }
  }

  // --- Extend Config validation/formatting for new MAX_TEAMS range ---
  config.getRange("A4:A" + (MAX_TEAMS + 3)).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(A4="",AND(ISNUMBER(A4),A4=INT(A4),A4>0,COUNTIF($A$4:$A$' + (MAX_TEAMS + 3) + ',A4)<=1))')
      .setAllowInvalid(false)
      .setHelpText("Enter a unique positive integer team number. Duplicates are not allowed.")
      .build()
  );
  config.getRange("A4:C" + (MAX_TEAMS + 3)).setBackground("#FFF2CC")
    .setBorder(true, true, true, true, true, true);

  // --- Append missing Config teams to referee sheets ---
  // Reuse configTeamRange from formula count computation above
  let allConfigTeams = [];
  for (let i = 0; i < configTeamRange.length; i++) {
    if (configTeamRange[i][0] !== "" && configTeamRange[i][0] !== null) {
      allConfigTeams.push(configTeamRange[i][0]);
    }
  }

  let noteMap3 = _buildNoteMap(ss);
  let teamsAdded = 0;
  for (let r = 1; r <= NUM_REFEREES; r++) {
    let sheet = findRefSheet(ss, config, r, noteMap3);
    if (!sheet) continue;
    let currentTeams = sheet.getRange(REF_DATA_START, 1, MAX_TEAMS, 1).getValues();
    let onSheet = {};
    let lastFilledIdx = -1;
    for (let i = 0; i < currentTeams.length; i++) {
      if (currentTeams[i][0] !== "" && currentTeams[i][0] !== null) {
        onSheet[currentTeams[i][0]] = true;
        lastFilledIdx = i;
      }
    }
    let missingTeams = [];
    for (let i = 0; i < allConfigTeams.length; i++) {
      if (!onSheet[allConfigTeams[i]]) missingTeams.push(allConfigTeams[i]);
    }
    if (missingTeams.length > 0) {
      let writeIdx = lastFilledIdx + 1;
      let toWrite = [];
      for (let m = 0; m < missingTeams.length && writeIdx + m < MAX_TEAMS; m++) {
        toWrite.push([missingTeams[m]]);
      }
      if (toWrite.length > 0) {
        sheet.getRange(REF_DATA_START + writeIdx, 1, toWrite.length, 1).setValues(toWrite);
        teamsAdded = Math.max(teamsAdded, toWrite.length);
      }
    }
  }

  // --- Rebuild FinalScores and restore selections ---
  _buildFinalScoresSheet(ss, formulaRows);
  if (savedOfficialRefs) {
    fsSheet = ss.getSheetByName("FinalScores");
    if (fsSheet) {
      // Restore only the rows that existed in the old FinalScores
      let restoreRows = Math.min(savedOfficialRefs.length, MAX_TEAMS);
      if (restoreRows > 0) {
        fsSheet.getRange("E" + FS_DATA_START + ":E" + (FS_DATA_START + restoreRows - 1))
          .setValues(savedOfficialRefs.slice(0, restoreRows));
      }
    }
  }

  // Move FinalScores to first tab
  fsSheet = ss.getSheetByName("FinalScores");
  if (fsSheet) {
    ss.setActiveSheet(fsSheet);
    ss.moveActiveSheet(1);
  }

  // Hide unnamed referee sheets
  _hideUnnamedRefSheets(ss, config);

  // Clean up temp sheet
  try { ss.deleteSheet(temp); } catch(e) {}

  if (config) config.activate();
  SpreadsheetApp.flush();

  let addedMsg = teamsAdded > 0 ? "\n\n" + teamsAdded + " missing team(s) from Config were appended to each referee sheet." : "";
  ui.alert(
    "Update Complete",
    "All sheets updated to the latest template.\n" +
    "Referee scoring data and Official Referee selections have been preserved." + addedMsg + "\n\n" +
    "Note: Sheet protection has NOT been reapplied.\n" +
    "Run '" + GAME_NAME + " Scoring > Apply Sheet Protection' if needed.",
    ui.ButtonSet.OK
  );
}

// ============================================================
// CONFIG SHEET (internal — called by buildAll)
// ============================================================
function _buildConfigSheet(ss) {
  let oldSheet = ss.getSheetByName("Config");
  let sheet = ss.insertSheet("Config" + (oldSheet ? "_new" : ""));
  if (oldSheet) ss.deleteSheet(oldSheet);
  sheet.setName("Config");

  // Rows 1-3: Headers, referee names, and labels (batch write)
  let row1 = ["Team #", "Team Name", "Video"];
  let row2 = ["Name \u2192", "", ""];
  let row3 = ["Email \u2192", "", ""];
  for (let r = 1; r <= NUM_REFEREES; r++) {
    row1.push("Referee " + r);
    row2.push("Referee " + r);
    row3.push("");
  }
  for (let r = 1; r <= NUM_REFEREES; r++) {
    row1.push("Ref " + r + " Order");
    row2.push("(randomized)");
    row3.push("(do not edit)");
  }
  sheet.getRange(1, 1, 3, row1.length).setValues([row1, row2, row3]);

  // Row 3: Referee email validation
  sheet.getRange("D3:I3").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(D3="",REGEXMATCH(D3,"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$"))')
      .setAllowInvalid(false)
      .setHelpText("Enter a valid Google account email address for this referee.")
      .build()
  );


  // ---- FORMATTING ----
  sheet.getRange("A1:O1").setFontWeight("bold").setBackground("#4472C4").setFontColor("white");
  sheet.getRange("A2:A3").setFontWeight("bold").setBackground("#D6E4F0").setFontColor("#1F4E79");
  sheet.getRange("D2:I2").setBackground("#FFF2CC").setFontWeight("bold");
  sheet.getRange("D3:I3").setBackground("#FFF2CC");
  sheet.getRange("J1:O3").setBackground("#F2F2F2").setFontColor("#5A5A5A");

  sheet.setColumnWidth(1, 85);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 300);
  for (let c = 4; c <= 9; c++) sheet.setColumnWidth(c, 120);
  for (let c = 10; c <= 15; c++) sheet.setColumnWidth(c, 100);

  sheet.getRange("A4:C" + (MAX_TEAMS + 3)).setBackground("#FFF2CC")
    .setBorder(true, true, true, true, true, true);

  // Prevent duplicate team numbers; enforce positive integers
  sheet.getRange("A4:A" + (MAX_TEAMS + 3)).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(A4="",AND(ISNUMBER(A4),A4=INT(A4),A4>0,COUNTIF($A$4:$A$' + (MAX_TEAMS + 3) + ',A4)<=1))')
      .setAllowInvalid(false)
      .setHelpText("Enter a unique positive integer team number. Duplicates are not allowed.")
      .build()
  );

  // Prevent duplicate referee names; block characters that break INDIRECT or sheet tabs
  sheet.getRange("D2:I2").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(D2="",AND(COUNTIF($D$2:$I$2,D2)<=1,REGEXMATCH(D2,"^[A-Za-z0-9._-][A-Za-z0-9 ._-]*$"),NOT(REGEXMATCH(D2,"^(Config|Rules|FinalScores|_temp_.*)$"))))')
      .setAllowInvalid(false)
      .setHelpText("Enter a unique referee name. Allowed: letters, numbers, spaces, hyphens, periods, underscores. Must not start with a space. Reserved names (Config, Rules, FinalScores) are not allowed.")
      .build()
  );

  sheet.setFrozenRows(3);
  // Hide row 3 (referee emails — admin-only)
  sheet.hideRows(3);
}

// ============================================================
// RULES SHEET (hidden — G Rules text for dropdown validation)
// ============================================================
function _buildRulesSheet(ss) {
  let oldSheet = ss.getSheetByName("Rules");
  if (oldSheet) { try { ss.deleteSheet(oldSheet); } catch(e) {} }
  let sheet = ss.insertSheet("Rules");
  let data = [];
  for (let i = 0; i < G_RULES.length; i++) {
    data.push([G_RULES[i]]);
  }
  sheet.getRange(1, 1, data.length, 1).setValues(data);
  sheet.setColumnWidth(1, 400);
  sheet.hideSheet();
}

// ============================================================
// REFEREE SHEET (internal — called by buildAll)
// ============================================================
function _buildRefereeSheet(ss, config, refNum, formulaRows) {
  let sheetName = getRefSheetName(config, refNum);
  let oldSheet = ss.getSheetByName(sheetName);
  let sheet = ss.insertSheet(sheetName + (oldSheet ? "_new" : ""));
  if (oldSheet) ss.deleteSheet(oldSheet);
  sheet.setName(sheetName);

  sheet.getRange("A1").setNote("ref_index:" + refNum);

  // Column letter shortcuts
  let cA = _colLetter(RC.TEAM), cD = _colLetter(RC.MOTIF), cE = _colLetter(RC.NOTES);
  let cF = _colLetter(RC.TOTAL), cG = _colLetter(RC.SCORE_NO_FOULS);
  let cH = _colLetter(RC.AUTO_SCORE), cI = _colLetter(RC.TEL_SCORE), cJ = _colLetter(RC.FOUL_DED);
  let cK = _colLetter(RC.MINOR), cL = _colLetter(RC.MAJOR), cM = _colLetter(RC.G_RULES);
  let cN = _colLetter(RC.LEAVE), cO = _colLetter(RC.AUTO_CLS), cP = _colLetter(RC.AUTO_OVF);
  let cQ = _colLetter(RC.AUTO_RAMP);
  let cS = _colLetter(RC.TEL_CLS), cT = _colLetter(RC.TEL_OVF), cU = _colLetter(RC.TEL_DEPOT);
  let cV = _colLetter(RC.TEL_RAMP), cX = _colLetter(RC.BASE);
  let ds = REF_DATA_START, de = REF_DATA_END;
  let lastCol = cX; // last column letter

  // ---- ROW 1: Title + progress counter (split merge at frozen column boundary) ----
  // A1:C1 (frozen) = progress counter; D1:V1 (scrollable) = title
  sheet.getRange("A1:C1").merge();
  // Composite "scored" check: MOTIF OR LEAVE OR AUTO_CLS non-empty
  // Shows "✓ All N scored" when complete, "Scored: X / Y" otherwise
  let totalTeams = 'COUNTA(' + cA + ds + ':' + cA + de + ')';
  let scoredTeams = 'SUMPRODUCT((' + cA + ds + ':' + cA + de + '<>"")*((' + cD + ds + ':' + cD + de + '<>"")+(' + cN + ds + ':' + cN + de + '<>"")+(' + cO + ds + ':' + cO + de + '<>"")>0))';
  sheet.getRange("A1").setFormula(
    '=IF(' + totalTeams + '=0,"No teams",' +
    'IF(' + scoredTeams + '=' + totalTeams + ',"\u2713 All "&' + totalTeams + '&" scored",' +
    '"Scored: "&' + scoredTeams + '&" / "&' + totalTeams + '))'
  );
  sheet.getRange("A1:C1").setFontSize(11).setFontWeight("bold")
    .setBackground("#1F4E79").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("D1").setValue(GAME_NAME + " " + SEASON + " Match Review \u2014 " + sheetName);
  sheet.getRange("D1:" + lastCol + "1").merge().setFontSize(14).setFontWeight("bold")
    .setBackground("#1F4E79").setFontColor("white").setHorizontalAlignment("center");

  // ---- ROW 2: Point values (hidden quick reference) ----
  let pointRow = new Array(RC.BASE).fill("");
  pointRow[RC.LEAVE - 1] = String(PTS_LEAVE);
  pointRow[RC.AUTO_CLS - 1] = "\u00d7" + PTS_CLASSIFIED;
  pointRow[RC.AUTO_OVF - 1] = "\u00d7" + PTS_OVERFLOW;
  pointRow[RC.AUTO_RAMP - 1] = "\u00d7" + PTS_PATTERN + " ea";
  pointRow[RC.TEL_CLS - 1] = "\u00d7" + PTS_CLASSIFIED;
  pointRow[RC.TEL_OVF - 1] = "\u00d7" + PTS_OVERFLOW;
  pointRow[RC.TEL_DEPOT - 1] = "\u00d7" + PTS_DEPOT;
  pointRow[RC.TEL_RAMP - 1] = "\u00d7" + PTS_PATTERN + " ea";
  pointRow[RC.BASE - 1] = PTS_BASE_PARTIAL + " / " + PTS_BASE_FULL;
  pointRow[RC.MINOR - 1] = "\u00d7(\u2212" + PTS_MINOR_FOUL + ")";
  pointRow[RC.MAJOR - 1] = "\u00d7(\u2212" + PTS_MAJOR_FOUL + ")";
  sheet.getRange(2, 1, 1, RC.BASE).setValues([pointRow]);
  sheet.getRange("A2:" + lastCol + "2").setFontStyle("italic").setFontSize(10)
    .setHorizontalAlignment("center").setBackground("#E8E8E8").setFontColor("#505050");
  sheet.hideRows(2);

  // ---- ROW 3: Column Headers ----
  let headers = [
    "Team #", "Name", "Video",                             // A-C
    "Notes",                                               // D
    "TOTAL", "w/o\nFouls", "Auto\nScore",                  // E-G
    "Tel\nScore", "Foul\nDed",                              // H-I
    "Minor", "Major", "G Rules",                            // J-L
    "MOTIF", "LEAVE", "Auto\nCLASS", "Auto\nOFLOW",        // M-P
    "Auto\nRAMP",                                           // Q
    "Tel\nCLASS", "Tel\nOFLOW", "Tel\nDEPOT",              // R-T
    "Tel\nRAMP",                                            // U
    "BASE"                                                  // V
  ];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange("A3:" + lastCol + "3").setFontWeight("bold").setWrap(true)
    .setVerticalAlignment("middle").setHorizontalAlignment("center").setFontColor("white");
  sheet.setRowHeight(3, 60);

  // Color-code header groups
  sheet.getRange("A3:C3").setBackground("#2F5496");   // Team
  sheet.getRange(cD + "3").setBackground("#7030A0");   // MOTIF
  sheet.getRange(cE + "3").setBackground("#696969");   // Notes
  sheet.getRange(cF + "3:" + cJ + "3").setBackground("#2F5496"); // Scores
  sheet.getRange(cK + "3:" + cL + "3").setBackground("#C00000"); // Fouls
  sheet.getRange(cM + "3").setBackground("#8B6914");   // G Rules
  sheet.getRange(cN + "3:" + cQ + "3").setBackground("#548235"); // Auto
  sheet.getRange(cS + "3:" + lastCol + "3").setBackground("#C55A11"); // TeleOp

  // ---- DATA ROWS (batch formula writes) ----
  // Gate: team# only (no MOTIF gate)
  let formulaCount = formulaRows || MAX_TEAMS;
  let fe = ds + formulaCount - 1;
  let formulasB = [], formulasC = [], formulasFJ = [];
  for (let row = ds; row <= fe; row++) {
    let gate = '$' + cA + row + '=""';

    formulasB.push(['=IF(' + cA + row + '="","",IFERROR(VLOOKUP(' + cA + row + ',Config!$A:$' + _colLetter(3) + ',2,FALSE),""))']);
    formulasC.push(['=IF(' + cA + row + '="","",IFERROR(VLOOKUP(' + cA + row + ',Config!$A:$' + _colLetter(3) + ',3,FALSE),""))']);

    // Inline PATTERN count: SUMPRODUCT/MID/SEQUENCE matching RAMP colors against MOTIF
    // Returns 0 when MOTIF blank or "Not Shown", 0 when RAMP empty
    let autoPatInline = 'IF(OR(' + cD + row + '="",' + cD + row + '="Not Shown"),0,' +
      'IF(LEN(' + cQ + row + ')=0,0,SUMPRODUCT((MID(UPPER(' + cQ + row + '),SEQUENCE(MIN(LEN(' + cQ + row + '),' + RAMP_MAX_CHARS + ')),1)=' +
      'MID(REPT(' + cD + row + ',3),SEQUENCE(MIN(LEN(' + cQ + row + '),' + RAMP_MAX_CHARS + ')),1))*1)))';
    let telPatInline = 'IF(OR(' + cD + row + '="",' + cD + row + '="Not Shown"),0,' +
      'IF(LEN(' + cV + row + ')=0,0,SUMPRODUCT((MID(UPPER(' + cV + row + '),SEQUENCE(MIN(LEN(' + cV + row + '),' + RAMP_MAX_CHARS + ')),1)=' +
      'MID(REPT(' + cD + row + ',3),SEQUENCE(MIN(LEN(' + cV + row + '),' + RAMP_MAX_CHARS + ')),1))*1)))';

    // F: TOTAL = max(0, score - fouls)
    let fF = '=IF(' + gate + ',"",MAX(0,' + cG + row + '-' + cJ + row + '))';
    // G: SCORE_NO_FOULS = Auto + TeleOp
    let fG = '=IF(' + gate + ',"",'+cH + row + '+' + cI + row + ')';
    // H: AUTO_SCORE = LEAVE + CLS*pts + OVF*pts + PATTERN*pts (inline)
    let fH = '=IF(' + gate + ',"",IF(' + cN + row + '="Yes",' + PTS_LEAVE + ',0)+' +
      cO + row + '*' + PTS_CLASSIFIED + '+' + cP + row + '*' + PTS_OVERFLOW + '+' + autoPatInline + '*' + PTS_PATTERN + ')';
    // I: TEL_SCORE = CLS*pts + OVF*pts + DEPOT*pts + PATTERN*pts (inline) + BASE
    let fI = '=IF(' + gate + ',"",'+cS + row + '*' + PTS_CLASSIFIED + '+' + cT + row + '*' + PTS_OVERFLOW + '+' +
      cU + row + '*' + PTS_DEPOT + '+' + telPatInline + '*' + PTS_PATTERN + '+' +
      'IF(' + cX + row + '="Full",' + PTS_BASE_FULL + ',IF(' + cX + row + '="Partial",' + PTS_BASE_PARTIAL + ',0)))';
    // J: FOUL_DED = Minor*pts + Major*pts
    let fJ = '=IF(' + gate + ',"",'+cK + row + '*' + PTS_MINOR_FOUL + '+' + cL + row + '*' + PTS_MAJOR_FOUL + ')';
    formulasFJ.push([fF, fG, fH, fI, fJ]);
  }
  sheet.getRange(ds, RC.NAME, formulaCount, 1).setFormulas(formulasB);
  sheet.getRange(ds, RC.VIDEO, formulaCount, 1).setFormulas(formulasC);
  sheet.getRange(ds, RC.TOTAL, formulaCount, 5).setFormulas(formulasFJ);

  // ---- DATA VALIDATION ----
  // MOTIF dropdown (including "Not Shown")
  sheet.getRange(cD + ds + ":" + cD + de).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(MOTIFS, true)
      .setAllowInvalid(false)
      .setHelpText("Select the MOTIF shown on the OBELISK (" + MOTIFS.join(", ") + "). Select 'Not Shown' if the OBELISK was not visible.")
      .build()
  );

  // LEAVE dropdown
  sheet.getRange(cN + ds + ":" + cN + de).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(LEAVE_OPTIONS, true)
      .setAllowInvalid(false)
      .setHelpText("Did the robot LEAVE? Robot must no longer be over any LAUNCH LINE at the end of AUTO.")
      .build()
  );

  // BASE dropdown
  sheet.getRange(cX + ds + ":" + cX + de).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(BASE_OPTIONS, true)
      .setAllowInvalid(false)
      .setHelpText("Robot position on BASE TILE at end of TELEOP: " + BASE_OPTIONS.join(", ") + ".")
      .build()
  );

  // G Rules dropdown (from hidden Rules sheet, allow invalid for multiselect)
  let rulesSheet = ss.getSheetByName("Rules");
  if (rulesSheet) {
    sheet.getRange(cM + ds + ":" + cM + de).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(rulesSheet.getRange("A1:A" + G_RULES.length), true)
        .setAllowInvalid(true)
        .setHelpText("Select a G rule from the dropdown. Multiple selections are accumulated as comma-separated codes. To remove a rule, re-open the dropdown and select it again.")
        .build()
    );
  }

  // Integer validation for count fields
  function intRule(colLetter, helpText) {
    return SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(' + colLetter + ds + '="",AND(ISNUMBER(' + colLetter + ds + '),' + colLetter + ds + '>=0,INT(' + colLetter + ds + ')=' + colLetter + ds + '))')
      .setAllowInvalid(false)
      .setHelpText(helpText)
      .build();
  }
  let intFields = [
    [cO, "Count of CLASSIFIED artifacts during AUTO (passed through SQUARE directly to RAMP). Whole number \u2265 0."],
    [cP, "Count of OVERFLOW artifacts during AUTO (passed through SQUARE but not directly to RAMP). Whole number \u2265 0."],
    [cS, "Count of CLASSIFIED artifacts during TELEOP (passed through SQUARE directly to RAMP). Whole number \u2265 0."],
    [cT, "Count of OVERFLOW artifacts during TELEOP (passed through SQUARE but not directly to RAMP). Whole number \u2265 0."],
    [cU, "Count of DEPOT artifacts at end of TELEOP (whole number \u2265 0)."],
    [cK, "Number of MINOR fouls (whole number \u2265 0). Each = " + PTS_MINOR_FOUL + " point deduction."],
    [cL, "Number of MAJOR fouls (whole number \u2265 0). Each = " + PTS_MAJOR_FOUL + " point deduction."]
  ];
  for (let f = 0; f < intFields.length; f++) {
    sheet.getRange(intFields[f][0] + ds + ":" + intFields[f][0] + de)
      .setDataValidation(intRule(intFields[f][0], intFields[f][1]));
  }

  // RAMP Colors (G/P characters, max 9)
  let rampCols = [
    [cQ, "Enter artifact colors on the RAMP at end of AUTO, in order from GATE to SQUARE (position 1 = nearest GATE, e.g., for GPP motif the full pattern is GPPGPPGPP). Use G (green) or P (purple). Max " + RAMP_MAX_CHARS + " characters."],
    [cV, "Enter artifact colors on the RAMP at end of TELEOP, in order from GATE to SQUARE (position 1 = nearest GATE, e.g., for GPP motif the full pattern is GPPGPPGPP). Use G (green) or P (purple). Max " + RAMP_MAX_CHARS + " characters."]
  ];
  for (let rc = 0; rc < rampCols.length; rc++) {
    let col = rampCols[rc][0];
    sheet.getRange(col + ds + ":" + col + de).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied('=OR(' + col + ds + '="",REGEXMATCH(UPPER(' + col + ds + '),"' + RAMP_REGEX + '"))')
        .setAllowInvalid(false)
        .setHelpText(rampCols[rc][1])
        .build()
    );
  }

  // ---- FORMATTING (merged adjacent same-color ranges to reduce API calls) ----
  sheet.getRange(cA + ds + ":" + cA + de).setBackground("#F2F2F2");
  sheet.getRange("B" + ds + ":C" + de).setBackground("#E8F0FE");
  sheet.getRange(cD + ds + ":" + cD + de).setBackground("#E2D9F3");
  sheet.getRange(cE + ds + ":" + cE + de).setBackground("#FFF2CC");
  sheet.getRange(cF + ds + ":" + cJ + de).setBackground("#D6E4F0");
  sheet.getRange(cK + ds + ":" + cM + de).setBackground("#FFF2CC");     // K-L-M all #FFF2CC
  sheet.getRange(cN + ds + ":" + cQ + de).setBackground("#E2EFDA");     // N-P-Q all #E2EFDA
  sheet.getRange(cS + ds + ":" + cV + de).setBackground("#FDF2E9");     // S-U-V all #FDF2E9
  sheet.getRange(cX + ds + ":" + cX + de).setBackground("#FFF2CC");

  sheet.getRange(cF + ds + ":" + cF + de).setFontWeight("bold").setFontSize(11);
  sheet.getRange(cQ + ds + ":" + cQ + de).setFontFamily("Courier New").setFontWeight("bold");
  sheet.getRange(cV + ds + ":" + cV + de).setFontFamily("Courier New").setFontWeight("bold");

  sheet.getRange(cA + ds + ":" + lastCol + de).setHorizontalAlignment("center");
  sheet.getRange("B" + ds + ":C" + de).setHorizontalAlignment("left");
  sheet.getRange(cE + ds + ":" + cE + de).setHorizontalAlignment("left");
  sheet.getRange(cM + ds + ":" + cM + de).setHorizontalAlignment("left").setWrap(true);

  sheet.getRange("A3:" + lastCol + de).setBorder(true, true, true, true, true, true,
    "#B4B4B4", SpreadsheetApp.BorderStyle.SOLID);

  // Column widths: A=Team#, B=Name, C=Video, D=Notes, E=TOTAL, F=ScoreNoFouls,
  // G=AutoScore, H=TelScore, I=FoulDed, J=Minor, K=Major, L=GRules, M=MOTIF,
  // N=LEAVE, O=AutoCLS, P=AutoOVF, Q=AutoRAMP, R=TelCLS, S=TelOVF,
  // T=TelDEPOT, U=TelRAMP, V=BASE
  let colWidths = [55,120,200,150,55,55,55,55,50,50,50,55,60,50,55,55,75,55,55,50,75,55];
  for (let c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(3);


  // ---- CONDITIONAL FORMATTING ----
  let rules = [];

  // Required input fields for "unfinished" detection (12 fields, includes RAMP cols)
  let reqCols = [cD, cK, cL, cN, cO, cP, cQ, cS, cT, cU, cV, cX];

  // Build unfinished detection expressions for CF formulas
  let unfinishedParts = [];
  for (let ci = 0; ci < reqCols.length; ci++) {
    unfinishedParts.push('$' + reqCols[ci] + ds + '=""');
  }
  let isUnfinished = 'AND($' + cA + ds + '<>"",OR(' + unfinishedParts.join(',') + '))';

  // Count unfinished rows (SUMPRODUCT-based)
  let countParts = [];
  for (let ci = 0; ci < reqCols.length; ci++) {
    countParts.push('($' + reqCols[ci] + '$' + ds + ':$' + reqCols[ci] + '$' + de + '="")');
  }
  let countUnfinished = 'SUMPRODUCT(($' + cA + '$' + ds + ':$' + cA + '$' + de + '<>"")*(' + countParts.join('+') + '>0))';

  // Newest unfinished row (MAX IF)
  let newestParts = [];
  for (let ci = 0; ci < reqCols.length; ci++) {
    newestParts.push('($' + reqCols[ci] + '$' + ds + ':$' + reqCols[ci] + '$' + de + '="")');
  }
  let newestUnfinished = 'MAX(IF(($' + cA + '$' + ds + ':$' + cA + '$' + de + '<>"")*(' + newestParts.join('+') + '>0),ROW($' + cA + '$' + ds + ':$' + cA + '$' + de + '),0))';

  // 1. Red empty cells in non-newest unfinished rows (on required input columns only)
  for (let ci = 0; ci < reqCols.length; ci++) {
    let col = reqCols[ci];
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(' + isUnfinished + ',' + countUnfinished + '>1,ROW()<' + newestUnfinished + ',$' + col + ds + '="")')
      .setBackground("#FF9999")
      .setRanges([sheet.getRange(col + ds + ":" + col + de)])
      .build());
  }

  // 2. Yellow non-newest unfinished rows (entire row)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(' + isUnfinished + ',' + countUnfinished + '>1,ROW()<' + newestUnfinished + ')')
    .setBackground("#FFFF00")
    .setRanges([sheet.getRange(cA + ds + ":" + lastCol + de)])
    .build());

  // 3. Fouls > 0
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#FFC7CE").setFontColor("#9C0006")
    .setRanges([sheet.getRange(cK + ds + ":" + cL + de)])
    .build());

  // 4. Unscored row (team# present, no scoring data at all) — orange
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + cA + ds + '<>"",$' + cD + ds + '="",$' + cN + ds + '="",$' + cO + ds + '="")')
    .setBackground("#FDE9D9")
    .setRanges([sheet.getRange(cA + ds + ":" + lastCol + de)])
    .build());

  // 5. Zebra striping
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + cA + ds + '<>"",ISEVEN(ROW()))')
    .setBackground("#F0F4FA")
    .setRanges([sheet.getRange(cA + ds + ":" + lastCol + de)])
    .build());

  sheet.setConditionalFormatRules(rules);
}

// ============================================================
// onEdit TRIGGER — G Rules multiselect
// ============================================================
function onEdit(e) {
  let range = e.range;
  let sheet = range.getSheet();
  let col = range.getColumn();
  let row = range.getRow();
  if (row < REF_DATA_START || row > REF_DATA_END) return;
  // Only process referee sheets (identified by ref_index note on A1)
  let note = "";
  try { note = sheet.getRange("A1").getNote() || ""; } catch(ex) { return; }
  if (note.indexOf("ref_index:") !== 0) return;

  // --- RAMP auto-uppercase (case-insensitive input) ---
  if (col === RC.AUTO_RAMP || col === RC.TEL_RAMP) {
    let val = range.getValue();
    if (val && typeof val === "string") {
      let upper = val.toUpperCase();
      if (upper !== val) range.setValue(upper);
    }
    return;
  }

  // --- G Rules multiselect ---
  if (col !== RC.G_RULES) return;

  let newValue = e.value;
  if (!newValue) {
    // Cell cleared — restore dropdown validation from Rules sheet
    let rulesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rules");
    if (rulesSheet) {
      range.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(rulesSheet.getRange("A1:A" + G_RULES.length), true)
          .setAllowInvalid(true)
          .setHelpText("Select a G rule from the dropdown. Multiple selections are accumulated as comma-separated codes. To remove a rule, re-open the dropdown and select it again.")
          .build()
      );
    }
    return;
  }

  // Extract and validate the 4-character rule code against known G_RULES
  let code = newValue.substring(0, 4);
  let validCode = false;
  for (let i = 0; i < G_RULES.length; i++) {
    if (G_RULES[i].substring(0, 4) === code) { validCode = true; break; }
  }
  if (!validCode) { range.setValue(""); return; }

  let oldValue = e.oldValue || "";

  let codes = oldValue ? oldValue.split(", ").filter(function(c) { return c; }) : [];
  let idx = codes.indexOf(code);

  if (idx > -1) {
    codes.splice(idx, 1); // Toggle off
  } else {
    codes.push(code); // Add
  }

  let result = codes.join(", ");
  range.clearDataValidation();
  range.setValue(result);

  // If all codes removed, restore dropdown for next selection
  if (!result) {
    let rulesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rules");
    if (rulesSheet) {
      range.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(rulesSheet.getRange("A1:A" + G_RULES.length), true)
          .setAllowInvalid(true)
          .setHelpText("Select a G rule from the dropdown. Multiple selections are accumulated as comma-separated codes. To remove a rule, re-open the dropdown and select it again.")
          .build()
      );
    }
  }
}

// ============================================================
// FINAL SCORES SHEET (internal — called by buildAll)
// ============================================================
function _buildFinalScoresSheet(ss, formulaRows) {
  let oldSheet = ss.getSheetByName("FinalScores");
  let sheet = ss.insertSheet("FinalScores" + (oldSheet ? "_new" : ""));
  if (oldSheet) ss.deleteSheet(oldSheet);
  sheet.setName("FinalScores");

  // FinalScores column constants (24 cols, A-X, no hidden columns)
  let FS = {
    TEAM:1, NAME:2, VIDEO:3, SCORED_BY:4, OFFICIAL_REF:5, AGREE:6, NOTES:7,
    FINAL_SCORE:8, SCORE_NO_FOULS:9, AUTO_SCORE:10, TEL_SCORE:11, FOUL_DED:12,
    MINOR:13, MAJOR:14, G_RULES:15,
    LEAVE:16, AUTO_CLS:17, AUTO_OVF:18, AUTO_RAMP:19,
    TEL_CLS:20, TEL_OVF:21, TEL_DEPOT:22, TEL_RAMP:23, BASE:24
  };
  let fsLastVisCol = _colLetter(24); // X

  // Pre-compute INDIRECT referee range strings (one per referee, reused across all rows)
  let indRefStrs = [];
  for (let r = 1; r <= NUM_REFEREES; r++) {
    indRefStrs[r] = 'INDIRECT("\'"&Config!' + _refConfigCol(r) + '$2&"\'!$A:$' + _colLetter(RC.BASE) + '")';
  }
  function indRef(r) { return indRefStrs[r]; }

  // "hasScored" expression for a given referee — composite check
  function hasScored(r, rowRef) {
    return 'OR(IFERROR(VLOOKUP($A' + rowRef + ',' + indRef(r) + ',' + RC.MOTIF + ',FALSE),"")<>"",' +
           'IFERROR(VLOOKUP($A' + rowRef + ',' + indRef(r) + ',' + RC.LEAVE + ',FALSE),"")<>"",' +
           'IFERROR(VLOOKUP($A' + rowRef + ',' + indRef(r) + ',' + RC.AUTO_CLS + ',FALSE),"")<>"")';
  }

  // ---- ROW 1: Category group headers (merged) ----
  let groups = [
    {range: "A1:C1", label: "Teams",             bg: "#2F5496"},
    {range: _colLetter(FS.SCORED_BY)+"1:"+_colLetter(FS.NOTES)+"1", label: "Referee", bg: "#2E75B6"},
    {range: _colLetter(FS.FINAL_SCORE)+"1:"+_colLetter(FS.FOUL_DED)+"1", label: "Total Scores", bg: "#7030A0"},
    {range: _colLetter(FS.MINOR)+"1:"+_colLetter(FS.MAJOR)+"1", label: "Fouls", bg: "#C00000"},
    {range: _colLetter(FS.G_RULES)+"1", label: "G Rules", bg: "#8B6914"},
    {range: _colLetter(FS.LEAVE)+"1:"+_colLetter(FS.AUTO_RAMP)+"1", label: "Autonomous Period", bg: "#548235"},
    {range: _colLetter(FS.TEL_CLS)+"1:"+_colLetter(FS.BASE)+"1", label: "TeleOp Period", bg: "#C55A11"}
  ];
  for (let g = 0; g < groups.length; g++) {
    sheet.getRange(groups[g].range).merge()
      .setValue(groups[g].label)
      .setFontWeight("bold").setFontColor("white").setHorizontalAlignment("center")
      .setBackground(groups[g].bg).setFontSize(11);
  }

  // ---- ROW 2: Point values (hidden) ----
  sheet.getRange("A2:C2").merge().setBackground("#FFF5D6");
  sheet.getRange(_colLetter(FS.SCORED_BY) + "2:" + _colLetter(FS.NOTES) + "2").merge().setValue(
    "Select an Official Referee for each team. " +
    "\"Refs Agree?\" shows Yes when all referees match on every scoring element. " +
    "Notes show the selected referee's notes when an Official Referee is chosen; " +
    "otherwise all referees' notes are shown with name prefixes."
  ).setFontStyle("italic").setFontColor("#6B4400").setBackground("#FFF5D6")
   .setHorizontalAlignment("left").setWrap(true);

  let pvRow = new Array(FS.BASE - FS.FINAL_SCORE + 1).fill("");
  pvRow[FS.MINOR - FS.FINAL_SCORE] = "\u00d7(\u2212" + PTS_MINOR_FOUL + ")";
  pvRow[FS.MAJOR - FS.FINAL_SCORE] = "\u00d7(\u2212" + PTS_MAJOR_FOUL + ")";
  pvRow[FS.LEAVE - FS.FINAL_SCORE] = String(PTS_LEAVE);
  pvRow[FS.AUTO_CLS - FS.FINAL_SCORE] = "\u00d7" + PTS_CLASSIFIED;
  pvRow[FS.AUTO_OVF - FS.FINAL_SCORE] = "\u00d7" + PTS_OVERFLOW;
  pvRow[FS.TEL_CLS - FS.FINAL_SCORE] = "\u00d7" + PTS_CLASSIFIED;
  pvRow[FS.TEL_OVF - FS.FINAL_SCORE] = "\u00d7" + PTS_OVERFLOW;
  pvRow[FS.TEL_DEPOT - FS.FINAL_SCORE] = "\u00d7" + PTS_DEPOT;
  pvRow[FS.BASE - FS.FINAL_SCORE] = PTS_BASE_PARTIAL + "/" + PTS_BASE_FULL;
  sheet.getRange(2, FS.FINAL_SCORE, 1, pvRow.length).setValues([pvRow]);
  sheet.getRange(_colLetter(FS.FINAL_SCORE) + "2:" + fsLastVisCol + "2")
    .setFontWeight("bold").setHorizontalAlignment("center")
    .setFontSize(10).setFontColor("#505050").setBackground("#E8E8E8");
  sheet.hideRows(2);

  // ---- ROW 3: Column headers ----
  let headers = [
    "Team #", "Name", "Video",
    "Scored By", "Official\nRef", "Refs\nAgree?", "Notes",
    "Final\nScore", "w/o\nFouls", "Auto\nScore", "Tel\nScore", "Foul\nDed",
    "Minor", "Major", "G Rules",
    "LEAVE", "Auto\nCLASS", "Auto\nOFLOW", "Auto\nRAMP",
    "Tel\nCLASS", "Tel\nOFLOW", "Tel\nDEPOT", "Tel\nRAMP", "BASE"
  ];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange("A3:" + fsLastVisCol + "3").setFontWeight("bold").setWrap(true)
    .setVerticalAlignment("middle").setHorizontalAlignment("center").setFontColor("white");
  sheet.setRowHeight(3, 60);

  sheet.getRange("A3:C3").setBackground("#2F5496");
  sheet.getRange(_colLetter(FS.SCORED_BY) + "3:" + _colLetter(FS.NOTES) + "3").setBackground("#2E75B6");
  sheet.getRange(_colLetter(FS.FINAL_SCORE) + "3:" + _colLetter(FS.FOUL_DED) + "3").setBackground("#7030A0");
  sheet.getRange(_colLetter(FS.MINOR) + "3:" + _colLetter(FS.MAJOR) + "3").setBackground("#C00000");
  sheet.getRange(_colLetter(FS.G_RULES) + "3").setBackground("#8B6914");
  sheet.getRange(_colLetter(FS.LEAVE) + "3:" + _colLetter(FS.AUTO_RAMP) + "3").setBackground("#548235");
  sheet.getRange(_colLetter(FS.TEL_CLS) + "3:" + _colLetter(FS.BASE) + "3").setBackground("#C55A11");

  // ---- DATA ROWS ----
  // Mapping: FS column -> RC column for VLOOKUP (no PATTERN Count — those stay on referee sheets only)
  let vlookupMap = [
    [FS.FINAL_SCORE,   RC.TOTAL],
    [FS.SCORE_NO_FOULS, RC.SCORE_NO_FOULS],
    [FS.AUTO_SCORE,    RC.AUTO_SCORE],
    [FS.TEL_SCORE,     RC.TEL_SCORE],
    [FS.FOUL_DED,      RC.FOUL_DED],
    [FS.MINOR,         RC.MINOR],
    [FS.MAJOR,         RC.MAJOR],
    [FS.G_RULES,       RC.G_RULES],
    [FS.LEAVE,         RC.LEAVE],
    [FS.AUTO_CLS,      RC.AUTO_CLS],
    [FS.AUTO_OVF,      RC.AUTO_OVF],
    [FS.AUTO_RAMP,     RC.AUTO_RAMP],
    [FS.TEL_CLS,       RC.TEL_CLS],
    [FS.TEL_OVF,       RC.TEL_OVF],
    [FS.TEL_DEPOT,     RC.TEL_DEPOT],
    [FS.TEL_RAMP,      RC.TEL_RAMP],
    [FS.BASE,          RC.BASE]
  ];

  // Agreement check: all input columns except G_RULES
  let elemCols = [RC.MOTIF, RC.LEAVE, RC.AUTO_CLS, RC.AUTO_OVF, RC.AUTO_RAMP,
                  RC.TEL_CLS, RC.TEL_OVF, RC.TEL_DEPOT, RC.TEL_RAMP, RC.BASE,
                  RC.MINOR, RC.MAJOR];

  let ds = FS_DATA_START, de = FS_DATA_END;
  let formulaCount = formulaRows || MAX_TEAMS;
  let fe = ds + formulaCount - 1;

  // Build all formulas as arrays
  let formulasAD = [];  // A-D
  let formulasF = [];   // F (Agree?)
  let formulasG = [];   // G (Notes)
  let formulasHX = [];  // H-X (scores with inlined effectiveRef via LET)

  for (let row = ds; row <= fe; row++) {
    // A: Team #
    let fA = '=IF(Config!A' + row + '="","",Config!A' + row + ')';
    // B: Name
    let fB = '=IF(A' + row + '="","",IFERROR(VLOOKUP(A' + row + ',Config!$A:$' + _colLetter(3) + ',2,FALSE),""))';
    // C: Video
    let fC = '=IF(A' + row + '="","",IFERROR(VLOOKUP(A' + row + ',Config!$A:$' + _colLetter(3) + ',3,FALSE),""))';

    // D: Scored By — referee names who have scored (composite check)
    let refNameParts = [];
    for (let r = 1; r <= NUM_REFEREES; r++) {
      refNameParts.push('IF(' + hasScored(r, row) + ',Config!' + _refConfigCol(r) + '$2,"")');
    }
    let fD = '=IF(A' + row + '="","",IFERROR(TEXTJOIN(CHAR(10),TRUE,' + refNameParts.join(',') + '),""))';
    formulasAD.push([fA, fB, fC, fD]);

    // refCount expression (used in F, G, and score columns)
    let refCountParts = [];
    for (let r = 1; r <= NUM_REFEREES; r++) {
      refCountParts.push('IF(' + hasScored(r, row) + ',1,0)');
    }
    let refCountExpr = '(' + refCountParts.join('+') + ')';

    // F: Refs Agree? — concatenate all elemCol values per referee into one string, then compare.
    // This reduces 12 separate UNIQUE/FILTER checks to 1, keeping formula under 50,000 char limit.
    let concatParts = [];
    for (let r = 1; r <= NUM_REFEREES; r++) {
      let fieldParts = [];
      for (let e = 0; e < elemCols.length; e++) {
        fieldParts.push('IFERROR(VLOOKUP($A' + row + ',' + indRef(r) + ',' + elemCols[e] + ',FALSE),"")');
      }
      concatParts.push(
        'IF(' + hasScored(r, row) + ',UPPER(' + fieldParts.join('&"|"&') + '),"")'
      );
    }
    let concatJoined = concatParts.join(';');
    formulasF.push([
      '=IF(OR($A' + row + '="",' + refCountExpr + '=0),"",IFERROR(IF(' + refCountExpr + '=1,"N/A",' +
      'LET(cv,{' + concatJoined + '},IF(IFERROR(ROWS(UNIQUE(FILTER(cv,cv<>"")))=1,TRUE),"Yes","No"))),"N/A"))'
    ]);

    // effectiveRef expression — inlined via LET in Notes and score formulas
    // Official Ref if set and valid (MATCH defense-in-depth), else auto-select single ref, else ""
    let offRefCol = '$' + _colLetter(FS.OFFICIAL_REF) + row;
    let singleRefParts = [];
    for (let r = 1; r <= NUM_REFEREES; r++) {
      singleRefParts.push('IF(' + hasScored(r, row) + ',Config!' + _refConfigCol(r) + '$2,"")');
    }
    let effRefExpr = 'IF(AND(' + offRefCol + '<>"",ISNUMBER(MATCH(' + offRefCol + ',Config!$D$2:$I$2,0))),' +
      offRefCol + ',IF(rc=1,TEXTJOIN("",TRUE,' + singleRefParts.join(',') + '),""))';

    // G: Notes — two-mode (effectiveRef set: plain; not set: all refs with "Name: text")
    // Uses LET to compute rc (refCount) and er (effectiveRef) once per cell
    let notesEffRef = 'IFERROR(VLOOKUP($A' + row + ',INDIRECT("\'"&er&"\'!$A:$' + _colLetter(RC.BASE) + '"),' + RC.NOTES + ',FALSE),"")';
    let notesAllParts = [];
    for (let r = 1; r <= NUM_REFEREES; r++) {
      let noteVal = 'IFERROR(VLOOKUP($A' + row + ',' + indRef(r) + ',' + RC.NOTES + ',FALSE),"")';
      notesAllParts.push('IF(AND(' + hasScored(r, row) + ',' + noteVal + '<>""),Config!' + _refConfigCol(r) + '$2&": "&' + noteVal + ',"")');
    }
    formulasG.push([
      '=IF($A' + row + '="","",LET(rc,' + refCountExpr + ',er,' + effRefExpr + ',' +
      'IF(er<>"",' + notesEffRef + ',' +
      'IFERROR(TEXTJOIN(CHAR(10),TRUE,' + notesAllParts.join(',') + '),""))' +
      '))'
    ]);

    // H-X: Score columns — per-field agreement with inlined effectiveRef
    // Each formula uses LET(rc, ..., er, ..., <body>) to compute effectiveRef once.
    // When er is set (official ref selected or single ref), show that ref's value.
    // When multiple refs scored and no official ref, show value only if all refs agree;
    // otherwise show each ref's value ("RefName: val" per line) prefixed with CHAR(8203)
    // (zero-width space) so CF can detect disagreement via LEFT().
    let overrideRef = 'INDIRECT("\'"&er&"\'!$A:$' + _colLetter(RC.BASE) + '")';
    let rowFormulas = [];

    // Pre-build per-ref scored booleans for FILTER criteria (reused across all fields)
    let scoredBools = [];
    for (let r = 1; r <= NUM_REFEREES; r++) {
      scoredBools.push('IF(' + hasScored(r, row) + ',TRUE,FALSE)');
    }
    let filterCriteria = '{' + scoredBools.join(';') + '}';

    for (let v = 0; v < vlookupMap.length; v++) {
      let srcCol = vlookupMap[v][1];
      let effRefVal = 'IFERROR(VLOOKUP($A' + row + ',' + overrideRef + ',' + srcCol + ',FALSE),"")';

      // Per-ref values for this field
      let valParts = [];
      for (let r = 1; r <= NUM_REFEREES; r++) {
        valParts.push('IFERROR(VLOOKUP($A' + row + ',' + indRef(r) + ',' + srcCol + ',FALSE),"")');
      }
      let filterVals = '{' + valParts.join(';') + '}';

      // Per-ref "Name: value" strings for disagreement display
      let disagreeTextParts = [];
      for (let r = 1; r <= NUM_REFEREES; r++) {
        let refName = 'Config!' + _refConfigCol(r) + '$2';
        let refVal = valParts[r - 1]; // reuse already-built VLOOKUP expression
        disagreeTextParts.push('IF(' + hasScored(r, row) + ',' + refName + '&": "&' + refVal + ',"")');
      }
      let disagreeText = 'CHAR(8203)&TEXTJOIN(CHAR(10),TRUE,' + disagreeTextParts.join(',') + ')';

      let agreeCheck = 'LET(v,FILTER(' + filterVals + ',' + filterCriteria + '),' +
        'IF(ROWS(UNIQUE(v))=1,INDEX(v,1),' + disagreeText + '))';

      rowFormulas.push(
        '=IF($A' + row + '="","",LET(rc,' + refCountExpr + ',er,' + effRefExpr + ',' +
        'IF(er<>"",' + effRefVal + ',IF(rc<2,"",' + agreeCheck + '))))'
      );
    }
    formulasHX.push(rowFormulas);
  }

  // Batch write all formulas
  sheet.getRange(ds, 1, formulaCount, 4).setFormulas(formulasAD);
  sheet.getRange(ds, FS.AGREE, formulaCount, 1).setFormulas(formulasF);
  sheet.getRange(ds, FS.NOTES, formulaCount, 1).setFormulas(formulasG);
  sheet.getRange(ds, FS.FINAL_SCORE, formulaCount, vlookupMap.length).setFormulas(formulasHX);

  // ---- DATA VALIDATION ----
  let configSheet = ss.getSheetByName("Config");
  if (configSheet) {
    sheet.getRange(_colLetter(FS.OFFICIAL_REF) + ds + ":" + _colLetter(FS.OFFICIAL_REF) + de).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(configSheet.getRange("D2:I2"), true)
        .setAllowInvalid(false)
        .setHelpText("Select the referee whose scores should be used as the official record for this team.")
        .build()
    );
  }

  // ---- FORMATTING ----
  let cFS = function(n) { return _colLetter(n); };
  sheet.getRange("A" + ds + ":A" + de).setBackground("#F2F2F2");
  sheet.getRange("B" + ds + ":C" + de).setBackground("#E8F0FE");
  sheet.getRange(cFS(FS.SCORED_BY) + ds + ":" + cFS(FS.SCORED_BY) + de).setBackground("#E2EFDA");
  sheet.getRange(cFS(FS.OFFICIAL_REF) + ds + ":" + cFS(FS.OFFICIAL_REF) + de).setBackground("#FFF2CC");
  sheet.getRange(cFS(FS.NOTES) + ds + ":" + cFS(FS.NOTES) + de).setBackground("#FFF2CC").setWrap(true);
  sheet.getRange(cFS(FS.FINAL_SCORE) + ds + ":" + cFS(FS.FINAL_SCORE) + de).setFontWeight("bold").setFontSize(11).setBackground("#E2D9F3");
  sheet.getRange(cFS(FS.SCORE_NO_FOULS) + ds + ":" + cFS(FS.FOUL_DED) + de).setBackground("#F3EDF9");
  sheet.getRange(cFS(FS.MINOR) + ds + ":" + cFS(FS.MAJOR) + de).setBackground("#FCE4EC");
  sheet.getRange(cFS(FS.G_RULES) + ds + ":" + cFS(FS.G_RULES) + de).setBackground("#FFF8E1").setWrap(true);
  sheet.getRange(cFS(FS.LEAVE) + ds + ":" + cFS(FS.AUTO_RAMP) + de).setBackground("#E2EFDA");
  sheet.getRange(cFS(FS.TEL_CLS) + ds + ":" + cFS(FS.BASE) + de).setBackground("#FDF2E9");
  // Enable wrapping on all score/input columns so multi-line disagreement text is visible
  sheet.getRange(cFS(FS.FINAL_SCORE) + ds + ":" + fsLastVisCol + de).setWrap(true);

  sheet.getRange("A" + ds + ":" + fsLastVisCol + de).setHorizontalAlignment("center");
  sheet.getRange("B" + ds + ":C" + de).setHorizontalAlignment("left");
  sheet.getRange(cFS(FS.SCORED_BY) + ds + ":" + cFS(FS.SCORED_BY) + de).setHorizontalAlignment("left").setWrap(true);
  sheet.getRange(cFS(FS.NOTES) + ds + ":" + cFS(FS.NOTES) + de).setHorizontalAlignment("left");
  sheet.getRange(cFS(FS.G_RULES) + ds + ":" + cFS(FS.G_RULES) + de).setHorizontalAlignment("left");

  sheet.getRange("A3:" + fsLastVisCol + de).setBorder(true, true, true, true, true, true,
    "#B4B4B4", SpreadsheetApp.BorderStyle.SOLID);

  // Column widths: A-X (24 cols, no hidden columns)
  let colWidths = [
    55, 120, 200,                      // A-C: Team#, Name, Video
    120, 85, 55, 170,                  // D-G: ScoredBy, OfficialRef, Agree, Notes
    55, 55, 55, 55, 50,               // H-L: Scores
    50, 50, 55,                        // M-O: Minor, Major, GRules
    50, 55, 55, 65,                    // P-S: LEAVE, AutoCLS, AutoOVF, AutoRAMP
    55, 55, 50, 65, 50                 // T-X: TelCLS, TelOVF, TelDEPOT, TelRAMP, BASE
  ];
  for (let c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // ---- CONDITIONAL FORMATTING ----
  let rules = [];
  let agreeCol = cFS(FS.AGREE);
  let agreeRange = [sheet.getRange(agreeCol + ds + ":" + agreeCol + de)];

  // 1. Agree column Yes/No/N/A formatting
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Yes").setBackground("#C6EFCE").setFontColor("#006100").setBold(true)
    .setRanges(agreeRange).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("No").setBackground("#FFC7CE").setFontColor("#9C0006").setBold(true)
    .setRanges(agreeRange).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("N/A").setBackground("#F2F2F2").setFontColor("#5A5A5A").setBold(true)
    .setRanges(agreeRange).build());

  // 2. Per-field disagreement — red background on cells with CHAR(8203) prefix (shows ref values)
  let scoreDataRange = [sheet.getRange(cFS(FS.FINAL_SCORE) + ds + ":" + fsLastVisCol + de)];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEFT(' + cFS(FS.FINAL_SCORE) + ds + ',1)=CHAR(8203)')
    .setBackground("#FF9999")
    .setRanges(scoreDataRange)
    .build());

  // 3. Missing Official Ref — orange
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + ds + '<>"",$' + _colLetter(FS.OFFICIAL_REF) + ds + '="")')
    .setBackground("#FDE9D9")
    .setRanges([sheet.getRange(_colLetter(FS.OFFICIAL_REF) + ds + ":" + _colLetter(FS.OFFICIAL_REF) + de)])
    .build());

  // 4. Zebra striping
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + ds + '<>"",ISEVEN(ROW()))')
    .setBackground("#F0F4FA")
    .setRanges([sheet.getRange("A" + ds + ":" + fsLastVisCol + de)])
    .build());

  sheet.setConditionalFormatRules(rules);
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(3);
}

// ============================================================
// RANDOMIZE TEAM ORDERS
// ============================================================
function randomizeTeamOrders() {
  if (!checkAuthorization()) return;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("randomizeTeamOrders must be run from the " + GAME_NAME + " Scoring menu, not the script editor.");
    return;
  }
  let config = ss.getSheetByName("Config");
  if (!config) {
    ui.alert("Error", "Config sheet not found. Run 'Rebuild All Sheets' first.", ui.ButtonSet.OK);
    return;
  }

  if (_hasAnyScoringData(ss)) {
    let response = ui.alert(
      "Warning: Scoring Data Exists",
      "One or more referee sheets already contain scoring data.\n" +
      "Re-randomizing will break team-order alignment and corrupt scores.\n\n" +
      "Are you SURE you want to re-randomize?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  _doRenameRefSheets(ss, config);
  _hideUnnamedRefSheets(ss, config);

  let teamRange = config.getRange("A4:A" + (MAX_TEAMS + 3));
  let teamValues = teamRange.getValues();
  let teams = [];
  for (let i = 0; i < teamValues.length; i++) {
    if (teamValues[i][0] !== "" && teamValues[i][0] !== null) {
      teams.push(teamValues[i][0]);
    }
  }

  if (teams.length === 0) {
    ui.alert("No Teams Found", "No team numbers found in Config column A (starting row 4).\nPlease enter team numbers first.", ui.ButtonSet.OK);
    return;
  }

  let noteMap2 = _buildNoteMap(ss);
  for (let r = 1; r <= NUM_REFEREES; r++) {
    let shuffled = teams.slice();
    for (let i = shuffled.length - 1; i > 0; i--) {
      let j = Math.floor(Math.random() * (i + 1));
      let temp = shuffled[i];
      shuffled[i] = shuffled[j];
      shuffled[j] = temp;
    }

    let orderCol = r + 9;
    let orderValues = [];
    for (let i = 0; i < MAX_TEAMS; i++) {
      orderValues.push([i < shuffled.length ? shuffled[i] : ""]);
    }
    config.getRange(4, orderCol, MAX_TEAMS, 1).setValues(orderValues);

    let refSheet = findRefSheet(ss, config, r, noteMap2);
    if (refSheet) {
      refSheet.getRange(REF_DATA_START, 1, MAX_TEAMS, 1).setValues(orderValues);
    }
  }

  SpreadsheetApp.flush();
  ui.alert(
    "Randomization Complete",
    "Team orders randomized for all " + NUM_REFEREES + " referees.\n" +
    "Orders are saved in Config columns J-O and on each referee sheet.\n\n" +
    "Do NOT re-randomize after referees begin scoring!\n" +
    "If you re-randomized by mistake, run 'Update Sheets' to realign data.",
    ui.ButtonSet.OK
  );
}

/**
 * Reorders the active referee sheet so teams match Config/FinalScores order.
 * All scoring data moves with its team. Orphaned teams (on sheet but not in Config)
 * are appended after the Config-ordered teams.
 */
function reorderToConfigOrder() {
  if (!checkAuthorization()) return;
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) { return; }

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();

  let note = "";
  try { note = sheet.getRange("A1").getNote() || ""; } catch(e) {}
  if (!note.startsWith("ref_index:")) {
    ui.alert("Error", "Navigate to a referee sheet first.\nThis reorders the active sheet's teams to match Config/FinalScores order.", ui.ButtonSet.OK);
    return;
  }

  let config = ss.getSheetByName("Config");
  if (!config) {
    ui.alert("Error", "Config sheet not found.", ui.ButtonSet.OK);
    return;
  }

  let response = ui.alert(
    "Reorder to Config Order",
    "This will reorder \"" + sheet.getName() + "\" so teams match Config/FinalScores order.\n\n" +
    "All scoring data moves with its team \u2014 nothing is lost.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  // Input columns to preserve (all user-entered data)
  let inputCols = [RC.TEAM, RC.NOTES, RC.MINOR, RC.MAJOR, RC.G_RULES,
                   RC.MOTIF, RC.LEAVE, RC.AUTO_CLS, RC.AUTO_OVF, RC.AUTO_RAMP,
                   RC.TEL_CLS, RC.TEL_OVF, RC.TEL_DEPOT, RC.TEL_RAMP, RC.BASE];

  // Read all current data
  let allData = sheet.getRange(REF_DATA_START, 1, MAX_TEAMS, RC.BASE).getValues();

  // Build map: team# → input values
  let teamDataMap = {};
  for (let i = 0; i < allData.length; i++) {
    let team = allData[i][0];
    if (team !== "" && team !== null) {
      let rowInputs = {};
      for (let c = 0; c < inputCols.length; c++) {
        rowInputs[inputCols[c]] = allData[i][inputCols[c] - 1];
      }
      teamDataMap[team] = rowInputs;
    }
  }

  // Read Config order
  let configData = config.getRange("A4:A" + (MAX_TEAMS + 3)).getValues();
  let configOrder = [];
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] !== "" && configData[i][0] !== null) {
      configOrder.push(configData[i][0]);
    }
  }

  // Build reordered arrays
  let reordered = {};
  for (let c = 0; c < inputCols.length; c++) {
    reordered[inputCols[c]] = new Array(MAX_TEAMS).fill("");
  }

  let placed = {};
  let idx = 0;

  // Config teams in Config order
  for (let i = 0; i < configOrder.length && idx < MAX_TEAMS; i++) {
    let team = configOrder[i];
    if (teamDataMap[team]) {
      for (let c = 0; c < inputCols.length; c++) {
        reordered[inputCols[c]][idx] = teamDataMap[team][inputCols[c]];
      }
    } else {
      reordered[RC.TEAM][idx] = team;
    }
    placed[team] = true;
    idx++;
  }

  // Orphaned teams (on sheet but not in Config) appended after
  for (let i = 0; i < allData.length && idx < MAX_TEAMS; i++) {
    let team = allData[i][0];
    if (team !== "" && team !== null && !placed[team]) {
      for (let c = 0; c < inputCols.length; c++) {
        reordered[inputCols[c]][idx] = teamDataMap[team][inputCols[c]];
      }
      placed[team] = true;
      idx++;
    }
  }

  // Write back each input column
  for (let c = 0; c < inputCols.length; c++) {
    _writeColumn(sheet, REF_DATA_START, inputCols[c], reordered[inputCols[c]]);
  }

  ui.alert("Reorder Complete", "\"" + sheet.getName() + "\" teams now match Config/FinalScores order.\nAll scoring data has been preserved.", ui.ButtonSet.OK);
}

/**
 * Check whether any referee sheet contains actual scoring data.
 * Uses composite check: MOTIF, LEAVE, or AUTO_CLS non-empty.
 */
function _hasAnyScoringData(ss) {
  let config = ss.getSheetByName("Config");
  let validMotifs = {};
  for (let i = 0; i < MOTIFS.length; i++) validMotifs[MOTIFS[i]] = true;

  let noteMap = _buildNoteMap(ss);
  for (let r = 1; r <= NUM_REFEREES; r++) {
    let refSheet = findRefSheet(ss, config, r, noteMap);
    if (!refSheet) continue;
    let dataStart = _detectRefDataStart(refSheet);
    let layoutVer = _detectLayoutVersion(refSheet);
    let src = (layoutVer === "old") ? {MOTIF: 4, LEAVE: 5, AUTO_CLS: 6} :
              (layoutVer === "v2") ? {MOTIF: 4, LEAVE: 14, AUTO_CLS: 15} : RC;
    let numCols = (layoutVer === "old") ? 23 : (layoutVer === "v2" || layoutVer === "v3") ? 24 : RC.BASE;

    // Single batch read for all columns (defensive: clamp to actual sheet size)
    let readRows = Math.min(MAX_TEAMS, Math.max(0, refSheet.getMaxRows() - dataStart + 1));
    if (readRows <= 0) continue;
    let allData = refSheet.getRange(dataStart, 1, readRows, numCols).getValues();
    for (let i = 0; i < allData.length; i++) {
      if (validMotifs[allData[i][src.MOTIF - 1]]) return true;
      let leaveVal = allData[i][src.LEAVE - 1];
      if (leaveVal === "Yes" || leaveVal === "No") return true;
      let clsVal = allData[i][src.AUTO_CLS - 1];
      if (clsVal !== "" && clsVal !== null) return true;
    }
  }
  return false;
}

// ============================================================
// RENAME REFEREE SHEETS
// ============================================================
function renameRefSheets() {
  if (!checkAuthorization()) return;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("renameRefSheets must be run from the " + GAME_NAME + " Scoring menu, not the script editor.");
    return;
  }
  let config = ss.getSheetByName("Config");
  if (!config) {
    ui.alert("Error", "Config sheet not found. Run 'Rebuild All Sheets' first.", ui.ButtonSet.OK);
    return;
  }

  if (_hasAnyScoringData(ss)) {
    let response = ui.alert(
      "Warning: Scoring Data Exists",
      "One or more referee sheets already contain scoring data.\n" +
      "Renaming sheets will update tab names and titles but preserve all data.\n\n" +
      "Are you sure you want to rename?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  let renamed = _doRenameRefSheets(ss, config);
  _hideUnnamedRefSheets(ss, config);
  let renameMsg = renamed === 0
    ? "All referee sheets already match the names in Config. No changes needed."
    : renamed + " sheet(s) renamed to match Config names.";
  ui.alert("Rename Complete", renameMsg, ui.ButtonSet.OK);
}

/**
 * Two-phase rename to handle swapped names correctly.
 * Phase 1: Rename all sheets that need changing to temporary names.
 * Phase 2: Rename from temporary names to final desired names.
 */
function _doRenameRefSheets(ss, config) {
  let renamed = 0;
  let renames = [];

  let noteMap = _buildNoteMap(ss);
  for (let r = 1; r <= NUM_REFEREES; r++) {
    let desiredName = getRefSheetName(config, r);
    let sheet = findRefSheet(ss, config, r, noteMap);
    if (sheet && sheet.getName() !== desiredName) {
      renames.push({sheet: sheet, desiredName: desiredName, refNum: r});
    }
  }

  if (renames.length === 0) return 0;

  // Phase 1: Rename to temporary names
  for (let i = 0; i < renames.length; i++) {
    let tempName = "_temp_rename_" + renames[i].refNum;
    try { renames[i].sheet.setName(tempName); } catch(e) {
      Logger.log("Failed to rename " + renames[i].sheet.getName() + " to temp: " + e);
    }
  }

  // Phase 2: Rename to final names
  for (let i = 0; i < renames.length; i++) {
    try {
      renames[i].sheet.setName(renames[i].desiredName);
      // Title cell is D1 in new layout
      renames[i].sheet.getRange("D1:" + _colLetter(RC.BASE) + "1").merge()
        .setValue(GAME_NAME + " " + SEASON + " Match Review \u2014 " + renames[i].desiredName);
      renamed++;
    } catch(e) {
      Logger.log("Failed to rename to " + renames[i].desiredName + ": " + e);
      try { renames[i].sheet.setName("Referee " + renames[i].refNum); } catch(e2) {}
    }
  }

  return renamed;
}

// ============================================================
// PROTECTION
// ============================================================
function applyProtection() {
  if (!checkAuthorization()) return;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("applyProtection must be run from the " + GAME_NAME + " Scoring menu, not the script editor.");
    return;
  }
  let me = Session.getEffectiveUser();
  let meEmail = (me.getEmail() || "").toLowerCase();
  let config = ss.getSheetByName("Config");

  let refEmails = [];
  if (config) {
    for (let r = 1; r <= NUM_REFEREES; r++) {
      let email = config.getRange(_refConfigCol(r) + "3").getValue();
      refEmails.push(email ? email.toString().trim().toLowerCase() : "");
    }
  }

  let hasEmails = false;
  for (let i = 0; i < refEmails.length; i++) {
    if (refEmails[i] !== "" && refEmails[i].indexOf("@") !== -1) { hasEmails = true; break; }
  }

  let failedEmails = [];
  let cE = _colLetter(RC.NOTES);
  let cK = _colLetter(RC.MINOR);
  let cX = _colLetter(RC.BASE);

  // ---- Referee sheets ----
  let noteMap = _buildNoteMap(ss);
  for (let r = 1; r <= NUM_REFEREES; r++) {
    let sheet = findRefSheet(ss, config, r, noteMap);
    if (!sheet) continue;

    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { try { p.remove(); } catch(e) {} });
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { try { p.remove(); } catch(e) {} });

    let sheetProt = sheet.protect().setDescription(sheet.getName() + " Sheet");
    sheetProt.addEditor(me);
    _restrictEditors(sheetProt, [meEmail]);

    // Input ranges: Notes (D) and all scoring inputs (J:V are contiguous)
    let inputRange1 = sheet.getRange(cE + REF_DATA_START + ":" + cE + REF_DATA_END);  // Notes (col D)
    let inputRange2 = sheet.getRange(cK + REF_DATA_START + ":" + cX + REF_DATA_END);  // Minor..BASE (cols J-V)
    sheetProt.setUnprotectedRanges([inputRange1, inputRange2]);

    if (hasEmails && refEmails[r - 1] !== "" && refEmails[r - 1].indexOf("@") !== -1) {
      let refEmail = refEmails[r - 1];

      let inputRanges = [inputRange1, inputRange2];
      let rangeNames = ["Notes", "Scoring"];
      for (let ri = 0; ri < inputRanges.length; ri++) {
        let rangeProt = inputRanges[ri].protect().setDescription(sheet.getName() + " " + rangeNames[ri]);
        rangeProt.addEditor(me);
        try { rangeProt.addEditor(refEmail); } catch(e) {
          if (ri === 0) failedEmails.push(refEmail + " (" + sheet.getName() + ")");
        }
        _restrictEditors(rangeProt, [meEmail, refEmail]);
      }
    } else {
      sheetProt.setWarningOnly(true);
    }
  }

  // ---- FinalScores ----
  let finalSheet = ss.getSheetByName("FinalScores");
  if (finalSheet) {
    finalSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { try { p.remove(); } catch(e) {} });
    finalSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { try { p.remove(); } catch(e) {} });

    let protection = finalSheet.protect().setDescription("FinalScores - Protected");
    protection.addEditor(me);
    _restrictEditors(protection, [meEmail]);

    let overrideNameRange = finalSheet.getRange("E" + FS_DATA_START + ":E" + FS_DATA_END);
    protection.setUnprotectedRanges([overrideNameRange]);

    let rangeProt = overrideNameRange.protect().setDescription("Official Referee Selection");
    rangeProt.addEditor(me);
    _restrictEditors(rangeProt, [meEmail]);

  }

  // ---- Config ----
  if (config) {
    config.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { try { p.remove(); } catch(e) {} });
    config.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { try { p.remove(); } catch(e) {} });

    let protection = config.protect().setDescription("Config - Protected");
    protection.addEditor(me);
    _restrictEditors(protection, [meEmail]);

    let teamDataRange = config.getRange("A4:C" + (MAX_TEAMS + 3));
    let refInfoRange = config.getRange("D2:I3");
    protection.setUnprotectedRanges([teamDataRange, refInfoRange]);

    let teamProt = teamDataRange.protect().setDescription("Config - Team Data");
    teamProt.addEditor(me);
    _restrictEditors(teamProt, [meEmail]);

    let refInfoProt = refInfoRange.protect().setDescription("Config - Referee Info");
    refInfoProt.addEditor(me);
    _restrictEditors(refInfoProt, [meEmail]);
  }

  // Hide Config sheet
  if (config) {
    try { config.hideSheet(); } catch(e) {
      Logger.log("Could not hide Config sheet: " + e);
    }
  }

  // ---- Rules sheet ----
  let rulesSheet = ss.getSheetByName("Rules");
  if (rulesSheet) {
    rulesSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .forEach(function(p) { try { p.remove(); } catch(e) {} });
    rulesSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach(function(p) { try { p.remove(); } catch(e) {} });
    let rulesProt = rulesSheet.protect().setDescription("Rules - Protected");
    rulesProt.addEditor(me);
    _restrictEditors(rulesProt, [meEmail]);
  }

  // Hide unnamed referee sheets
  _hideUnnamedRefSheets(ss, config);

  let msg;
  if (hasEmails) {
    msg = "Protection applied with per-referee isolation!\n\n" +
      "- Each referee can ONLY edit their own sheet's scoring cells\n" +
      "- Formula cells, team info, and headers are locked\n" +
      "- FinalScores 'Official Referee' column is restricted to the owner\n" +
      "- Config sheet is now hidden (right-click any tab > Unhide to access it)\n" +
      "- Unused referee sheets are hidden\n\n" +
      "Make sure each referee has been shared on the spreadsheet.";
    if (failedEmails.length > 0) {
      msg += "\n\nWARNING: Could not grant access for:\n" + failedEmails.join("\n") +
        "\nCheck that these are valid Google account emails.";
    }
  } else {
    msg = "Protection applied (advisory mode).\n\n" +
      "- Formula cells are protected on all sheets\n" +
      "- Scoring input cells show a warning but are NOT restricted per-referee\n" +
      "- Config sheet is now hidden (right-click any tab > Unhide to access it)\n" +
      "- Unused referee sheets are hidden\n\n" +
      "To enable per-referee isolation:\n" +
      "1. Unhide Config sheet (right-click any tab > Unhide)\n" +
      "2. Unhide row 3 in Config (right-click row 2/4 border > Unhide rows)\n" +
      "3. Enter referee emails in row 3 (columns D\u2013I)\n" +
      "4. Re-run " + GAME_NAME + " Scoring > Apply Sheet Protection";
  }

  ui.alert("Protection Applied", msg, ui.ButtonSet.OK);
}
