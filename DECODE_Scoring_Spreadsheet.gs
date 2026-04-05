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
 *     against the MOTIF. Max 9 characters (RAMP capacity).
 *   - Fouls: Subtracted from team score (deviation from official rules for solo match context).
 *   - 2-robot BASE bonus (10 pts) and Ranking Points: Excluded (solo match, single robot).
 *
 * HOW TO UPDATE FOR A NEW SEASON:
 * ─────────────────────────────────────────────────────────
 *  1. GAME_NAME / SEASON — Update the game name and year.
 *  2. Point values (PTS_*) — See the competition manual scoring table.
 *  3. MOTIFS — The allowed gate-field values (e.g., obelisk patterns).
 *  4. LEAVE_OPTIONS / BASE_OPTIONS — Dropdown choices for Yes/No and base status.
 *  5. RAMP_REGEX / RAMP_MAX_CHARS — Update if the text-entry format changes.
 *     Remove RAMP logic entirely if the new game has no equivalent.
 *  6. Column headers — In _buildRefereeSheet(), update the 'headers' array.
 *  7. Scoring formulas — In _buildRefereeSheet(), update the formulas in the
 *     DATA ROWS section for Auto Score, TeleOp Score, etc.
 *  8. RC column indices — If columns are added/removed/reordered, update the
 *     RC object and layout constants as needed.
 *  9. FinalScores mapping — In _buildFinalScoresSheet(), update vlookupMap,
 *     elemCols, headers, and category groups.
 * 10. Help text — In the DATA VALIDATION section of _buildRefereeSheet(),
 *     update the text shown to referees on invalid input.
 * 11. Documentation — Update README.md and the Scoring Guide.
 * 12. Run buildAll() to generate all sheets from scratch.
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
const MAX_TEAMS = 50;

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
const MOTIFS = ["GPP", "PGP", "PPG"];
const LEAVE_OPTIONS = ["Yes", "No"];
const BASE_OPTIONS = ["None", "Partial", "Full"];

// --- Text-entry validation (RAMP color entry) ---
const RAMP_REGEX = "^[GP]{1,9}$";
const RAMP_MAX_CHARS = 9;

// --- Layout constants ---
// Referee sheets: Row 1=Title, Row 2=Point values, Row 3=Headers, Row 4+=Data
const REF_DATA_START = 4;
const REF_DATA_END = MAX_TEAMS + 3;
// FinalScores: Row 1=Category headers, Row 2=Point values, Row 3=Headers, Row 4+=Data
const FS_DATA_START = 4;
const FS_DATA_END = MAX_TEAMS + 3;

// Referee sheet column layout (update if columns change):
//   A=Team#(auto)  B=Name(auto)  C=Video(auto)
//   D=MOTIF(gate)  E=LEAVE  F=Auto CLS  G=Auto OVF  H=Auto RAMP Colors
//   I=Tel CLS  J=Tel OVF  K=Tel DEPOT  L=Tel RAMP Colors  M=BASE
//   N=Minor Fouls  O=Major Fouls
//   P=Auto PAT(calc)  Q=Auto Score(calc)  R=Tel PAT(calc)  S=Tel Score(calc)
//   T=Foul Deduction(calc)  U=Score w/o Fouls(calc)  V=TOTAL(calc)  W=Notes
const RC = {
  TEAM: 1, NAME: 2, VIDEO: 3, MOTIF: 4, LEAVE: 5,
  AUTO_CLS: 6, AUTO_OVF: 7, AUTO_RAMP: 8,
  TEL_CLS: 9, TEL_OVF: 10, TEL_DEPOT: 11, TEL_RAMP: 12,
  BASE: 13, MINOR: 14, MAJOR: 15,
  AUTO_PAT: 16, AUTO_SCORE: 17, TEL_PAT: 18, TEL_SCORE: 19,
  FOUL_DED: 20, SCORE_NO_FOULS: 21, TOTAL: 22, NOTES: 23
};

// ============================================================
// AUTHORIZATION
// ============================================================
// SHA-256 hashes of authorized emails (lowercase)
const AUTHORIZED_HASHES = [
  "c05ddb09d44266bc0b82e5b8322d000861b378638de771a24cce341c3859d7cc",
  "2ee4d1fd7155caaddb53b63a305413733616070f7f4802b6a13792167a4f1d88"
];

function _hashEmail(email) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email);
  var hex = "";
  for (var i = 0; i < raw.length; i++) {
    var b = (raw[i] + 256) % 256; // convert signed byte to unsigned
    hex += ("0" + b.toString(16)).slice(-2);
  }
  return hex;
}

function checkAuthorization() {
  var userEmail = (Session.getEffectiveUser().getEmail() || "").toLowerCase();
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
  var emailHash = _hashEmail(userEmail);
  for (var i = 0; i < AUTHORIZED_HASHES.length; i++) {
    if (emailHash === AUTHORIZED_HASHES[i]) return true;
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

/** Returns the Config column letter for a given referee number (1-6 -> D-I). */
function _refConfigCol(refNum) {
  return String.fromCharCode(67 + refNum);
}

function getRefSheetName(config, refNum) {
  if (!config) return "Referee " + refNum;
  var name = config.getRange(_refConfigCol(refNum) + "2").getValue();
  return (name && name.toString().trim() !== "") ? name.toString().trim() : "Referee " + refNum;
}

function findRefSheet(ss, config, refNum) {
  var sheet;
  if (config) {
    var name = getRefSheetName(config, refNum);
    sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  sheet = ss.getSheetByName("Referee " + refNum);
  if (sheet) return sheet;
  sheet = ss.getSheetByName("Referee" + refNum);
  if (sheet) return sheet;
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    try {
      if (allSheets[i].getRange("A1").getNote() === "ref_index:" + refNum) return allSheets[i];
    } catch(e) {}
  }
  return null;
}

/** Remove all editors except the specified allowed emails from a protection object. */
function _restrictEditors(protection, allowedEmails) {
  var allowed = {};
  for (var i = 0; i < allowedEmails.length; i++) {
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
  for (var r = 1; r <= NUM_REFEREES; r++) {
    var configName = config.getRange(_refConfigCol(r) + "2").getValue();
    var isDefault = !configName || configName.toString().trim() === "" ||
      configName.toString().trim() === "Referee " + r;
    if (isDefault) {
      var sheet = findRefSheet(ss, config, r);
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
  var frozen = sheet.getFrozenRows();
  return frozen >= 4 ? 5 : 4;
}

// ============================================================
// CUSTOM MENU
// ============================================================
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu(GAME_NAME + " Scoring")
      .addItem("Randomize Team Orders", "randomizeTeamOrders")
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
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("confirmRebuild: UI not available. Rebuild requires interactive confirmation.");
    return;
  }
  var response = ui.alert(
    "Rebuild All Sheets",
    "This will DELETE and recreate ALL sheets (Config, referee sheets, FinalScores).\n" +
    "All existing data will be LOST.\n\nAre you sure?",
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) buildAll();
}

function buildAll() {
  if (!checkAuthorization()) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Reuse leftover temp sheet from a previous failed build, or create new
  var temp = ss.getSheetByName("_temp_build") || ss.insertSheet("_temp_build");
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName() !== "_temp_build") {
      try { ss.deleteSheet(allSheets[i]); } catch(e) {}
    }
  }

  _buildConfigSheet(ss);
  var config = ss.getSheetByName("Config");
  for (var r = 1; r <= NUM_REFEREES; r++) {
    _buildRefereeSheet(ss, config, r);
  }
  _buildFinalScoresSheet(ss);

  try { ss.deleteSheet(temp); } catch(e) {}

  // Move FinalScores to first tab position
  var finalSheet = ss.getSheetByName("FinalScores");
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
      "3. Enter referee emails in Config row 3 (for per-referee protection)\n" +
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
 * Rebuilds all referee sheets and FinalScores with the current template
 * (formulas, formatting, validation, headers) while preserving:
 *   - Team orders (column A on referee sheets)
 *   - Referee scoring inputs (columns D-O, W on referee sheets)
 *   - Official Referee selections (column E on FinalScores)
 *
 * Use this after modifying the template code to apply changes to an
 * existing spreadsheet without losing referee work. Sheet protection
 * is NOT reapplied — run "Apply Sheet Protection" afterward if needed.
 */
function updateSheets() {
  if (!checkAuthorization()) return;
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("updateSheets must be run from the " + GAME_NAME + " Scoring menu.");
    return;
  }

  var response = ui.alert(
    "Update All Sheets",
    "This updates all sheet layouts, formulas, formatting, and validation " +
    "to the current template without erasing scoring data.\n\n" +
    "Preserved: team orders, referee scoring inputs, Official Referee selections.\n" +
    "Updated: formulas, formatting, validation, headers, conditional formatting.\n\n" +
    "Continue?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = ss.getSheetByName("Config");
  if (!config) {
    ui.alert("Error", "Config sheet not found. Run buildAll() first.", ui.ButtonSet.OK);
    return;
  }

  // --- Save referee scoring data ---
  var savedRefData = {};
  for (var r = 1; r <= NUM_REFEREES; r++) {
    var sheet = findRefSheet(ss, config, r);
    if (!sheet) continue;

    var dataStart = _detectRefDataStart(sheet);
    var dataEnd = dataStart + MAX_TEAMS - 1;

    savedRefData[r] = {
      teams: sheet.getRange("A" + dataStart + ":A" + dataEnd).getValues(),
      scoring: sheet.getRange("D" + dataStart + ":O" + dataEnd).getValues(),
      notes: sheet.getRange("W" + dataStart + ":W" + dataEnd).getValues()
    };
  }

  // --- Save FinalScores Official Referee selections ---
  var fsSheet = ss.getSheetByName("FinalScores");
  var savedOfficialRefs = null;
  if (fsSheet) {
    savedOfficialRefs = fsSheet.getRange("E" + FS_DATA_START + ":E" + FS_DATA_END).getValues();
  }

  // --- Create temp sheet to avoid last-sheet deletion errors ---
  var temp = ss.getSheetByName("_temp_update") || ss.insertSheet("_temp_update");

  // --- Delete old referee sheets ---
  for (var r = 1; r <= NUM_REFEREES; r++) {
    var sheet = findRefSheet(ss, config, r);
    if (sheet) {
      try { ss.deleteSheet(sheet); } catch(e) {
        Logger.log("Could not delete sheet for referee " + r + ": " + e);
      }
    }
  }

  // --- Delete old FinalScores ---
  fsSheet = ss.getSheetByName("FinalScores");
  if (fsSheet) {
    try { ss.deleteSheet(fsSheet); } catch(e) {}
  }

  // --- Rebuild all referee sheets and restore data ---
  for (var r = 1; r <= NUM_REFEREES; r++) {
    _buildRefereeSheet(ss, config, r);

    if (savedRefData[r]) {
      var sheet = findRefSheet(ss, config, r);
      if (sheet) {
        sheet.getRange(REF_DATA_START, 1, MAX_TEAMS, 1).setValues(savedRefData[r].teams);
        sheet.getRange("D" + REF_DATA_START + ":O" + REF_DATA_END).setValues(savedRefData[r].scoring);
        sheet.getRange("W" + REF_DATA_START + ":W" + REF_DATA_END).setValues(savedRefData[r].notes);
      }
    }
  }

  // --- Rebuild FinalScores and restore selections ---
  _buildFinalScoresSheet(ss);
  if (savedOfficialRefs) {
    fsSheet = ss.getSheetByName("FinalScores");
    if (fsSheet) {
      fsSheet.getRange("E" + FS_DATA_START + ":E" + FS_DATA_END).setValues(savedOfficialRefs);
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

  ui.alert(
    "Update Complete",
    "All sheets updated to the latest template.\n" +
    "Referee scoring data and Official Referee selections have been preserved.\n\n" +
    "Note: Sheet protection has NOT been reapplied.\n" +
    "Run '" + GAME_NAME + " Scoring > Apply Sheet Protection' if needed.",
    ui.ButtonSet.OK
  );
}

// ============================================================
// CONFIG SHEET (internal — called by buildAll)
// ============================================================
function _buildConfigSheet(ss) {
  var oldSheet = ss.getSheetByName("Config");
  var sheet = ss.insertSheet("Config" + (oldSheet ? "_new" : ""));
  if (oldSheet) ss.deleteSheet(oldSheet);
  sheet.setName("Config");

  // Row 1: Headers
  sheet.getRange("A1").setValue("Team #");
  sheet.getRange("B1").setValue("Team Name");
  sheet.getRange("C1").setValue("Video");
  for (var r = 1; r <= NUM_REFEREES; r++) {
    sheet.getRange(1, r + 3).setValue("Referee " + r);
    sheet.getRange(1, r + 9).setValue("Ref " + r + " Order");
  }

  // Row 2: Referee names
  sheet.getRange("A2").setValue("Name \u2192");
  for (var r = 1; r <= NUM_REFEREES; r++) {
    sheet.getRange(_refConfigCol(r) + "2").setValue("Referee " + r);
  }

  // Row 3: Referee emails
  sheet.getRange("A3").setValue("Email \u2192");

  // Randomized order labels
  for (var r = 1; r <= NUM_REFEREES; r++) {
    var col = String.fromCharCode(73 + r); // J-O
    sheet.getRange(col + "2").setValue("(randomized)");
    sheet.getRange(col + "3").setValue("(do not edit)");
  }

  // ---- FORMATTING ----
  sheet.getRange("A1:O1").setFontWeight("bold").setBackground("#4472C4").setFontColor("white");
  sheet.getRange("A2:A3").setFontWeight("bold").setBackground("#D6E4F0").setFontColor("#1F4E79");
  sheet.getRange("D2:I2").setBackground("#FFF2CC").setFontWeight("bold");
  sheet.getRange("D3:I3").setBackground("#FFF2CC");
  sheet.getRange("J1:O3").setBackground("#F2F2F2").setFontColor("#5A5A5A");

  sheet.setColumnWidth(1, 85);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 300);
  for (var c = 4; c <= 9; c++) sheet.setColumnWidth(c, 120);
  for (var c = 10; c <= 15; c++) sheet.setColumnWidth(c, 100);

  sheet.getRange("A4:C" + (MAX_TEAMS + 3)).setBackground("#FFF2CC")
    .setBorder(true, true, true, true, true, true);

  // Prevent duplicate team numbers
  sheet.getRange("A4:A" + (MAX_TEAMS + 3)).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(A4="",COUNTIF($A$4:$A$' + (MAX_TEAMS + 3) + ',A4)<=1)')
      .setAllowInvalid(false)
      .setHelpText("Enter a unique team number. Duplicates are not allowed.")
      .build()
  );

  // Prevent duplicate referee names; block characters that break INDIRECT or sheet tabs;
  // disallow leading/trailing spaces to prevent INDIRECT mismatch with trimmed sheet names
  sheet.getRange("D2:I2").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(D2="",AND(COUNTIF($D$2:$I$2,D2)<=1,REGEXMATCH(D2,"^[A-Za-z0-9._-][A-Za-z0-9 ._-]*$")))')
      .setAllowInvalid(false)
      .setHelpText("Enter a unique referee name. Allowed: letters, numbers, spaces, hyphens, periods, underscores. Must not start with a space.")
      .build()
  );

  sheet.setFrozenRows(3);
}

// ============================================================
// REFEREE SHEET (internal — called by buildAll)
// ============================================================
function _buildRefereeSheet(ss, config, refNum) {
  var sheetName = getRefSheetName(config, refNum);
  var oldSheet = ss.getSheetByName(sheetName);
  var sheet = ss.insertSheet(sheetName + (oldSheet ? "_new" : ""));
  if (oldSheet) ss.deleteSheet(oldSheet);
  sheet.setName(sheetName);

  sheet.getRange("A1").setNote("ref_index:" + refNum);

  // ---- ROW 1: Title + progress counter (split merge at frozen column boundary) ----
  // A1:B1 (frozen) = progress counter; C1:W1 (scrollable) = title
  sheet.getRange("A1:B1").merge();
  sheet.getRange("A1").setFormula(
    '=IF(COUNTA(A' + REF_DATA_START + ':A' + REF_DATA_END + ')=0,"No teams",' +
    '"Scored: "&COUNTA(D' + REF_DATA_START + ':D' + REF_DATA_END + ')&" / "&COUNTA(A' + REF_DATA_START + ':A' + REF_DATA_END + '))'
  );
  sheet.getRange("A1:B1").setFontSize(11).setFontWeight("bold")
    .setBackground("#1F4E79").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("C1").setValue(GAME_NAME + " " + SEASON + " Match Review \u2014 " + sheetName);
  sheet.getRange("C1:W1").merge().setFontSize(14).setFontWeight("bold")
    .setBackground("#1F4E79").setFontColor("white").setHorizontalAlignment("center");

  // ---- ROW 2: Point values (quick reference for referees) ----
  var pointLabels = [
    [RC.LEAVE, String(PTS_LEAVE)],
    [RC.AUTO_CLS, "\u00d7" + PTS_CLASSIFIED],
    [RC.AUTO_OVF, "\u00d7" + PTS_OVERFLOW],
    [RC.AUTO_RAMP, "\u00d7" + PTS_PATTERN + " ea"],
    [RC.TEL_CLS, "\u00d7" + PTS_CLASSIFIED],
    [RC.TEL_OVF, "\u00d7" + PTS_OVERFLOW],
    [RC.TEL_DEPOT, "\u00d7" + PTS_DEPOT],
    [RC.TEL_RAMP, "\u00d7" + PTS_PATTERN + " ea"],
    [RC.BASE, PTS_BASE_PARTIAL + " / " + PTS_BASE_FULL],
    [RC.MINOR, "\u00d7(\u2212" + PTS_MINOR_FOUL + ")"],
    [RC.MAJOR, "\u00d7(\u2212" + PTS_MAJOR_FOUL + ")"]
  ];
  for (var i = 0; i < pointLabels.length; i++) {
    sheet.getRange(2, pointLabels[i][0]).setValue(pointLabels[i][1]);
  }
  sheet.getRange("A2:W2").setFontStyle("italic").setFontSize(10)
    .setHorizontalAlignment("center").setBackground("#E8E8E8").setFontColor("#505050");

  // ---- ROW 3: Column Headers ----
  var headers = [
    "Team #",                     // A
    "Team Name",                  // B
    "Video",                      // C
    "MOTIF\n(required)",          // D
    "LEAVE\n(Yes/No)",            // E
    "Auto\nCLASSIFIED",           // F
    "Auto\nOVERFLOW",             // G
    "Auto RAMP\nColors\n(G/P)",   // H
    "TeleOp\nCLASSIFIED",         // I
    "TeleOp\nOVERFLOW",           // J
    "TeleOp\nDEPOT",              // K
    "TeleOp RAMP\nColors\n(G/P)", // L
    "BASE\n(None/Partial/Full)",  // M
    "Minor\nFouls",               // N
    "Major\nFouls",               // O
    "Auto PATTERN\nCount",        // P
    "Auto\nScore",                // Q
    "TeleOp PATTERN\nCount",      // R
    "TeleOp\nScore",              // S
    "Foul\nDeduction",            // T
    "Score w/o\nFouls",           // U
    "TOTAL\nSCORE",               // V
    "Notes"                       // W
  ];
  for (var c = 0; c < headers.length; c++) {
    sheet.getRange(3, c + 1).setValue(headers[c]);
  }
  sheet.getRange("A3:W3").setFontWeight("bold").setWrap(true).setVerticalAlignment("middle")
    .setHorizontalAlignment("center").setFontColor("white");
  sheet.setRowHeight(3, 60);

  // Color-code header groups (TeleOp uses darker gold for better contrast with white text)
  sheet.getRange("A3:C3").setBackground("#2F5496");
  sheet.getRange("D3").setBackground("#7030A0");
  sheet.getRange("E3:H3").setBackground("#548235");
  sheet.getRange("I3:M3").setBackground("#8B6914");
  sheet.getRange("N3:O3").setBackground("#C00000");
  sheet.getRange("P3:V3").setBackground("#2F5496");
  sheet.getRange("W3").setBackground("#696969");

  // ---- DATA ROWS (batch formula writes) ----
  var formulasB = [], formulasC = [], formulasPV = [];
  for (var row = REF_DATA_START; row <= REF_DATA_END; row++) {
    var gate = 'OR($A' + row + '="",$D' + row + '="")';

    formulasB.push(['=IF(A' + row + '="","",IFERROR(VLOOKUP(A' + row + ',Config!$A:$C,2,FALSE),""))']);
    formulasC.push(['=IF(A' + row + '="","",IFERROR(VLOOKUP(A' + row + ',Config!$A:$C,3,FALSE),""))']);

    formulasPV.push([
      // P: Auto PATTERN Count — character-by-character match of RAMP colors vs MOTIF repeated 3x
      '=IF(' + gate + ',"",IF(LEN(H' + row + ')=0,0,' +
        'SUMPRODUCT((MID(UPPER(H' + row + '),SEQUENCE(LEN(H' + row + ')),1)=' +
        'MID(REPT(D' + row + ',3),SEQUENCE(LEN(H' + row + ')),1))*1)))',
      // Q: Auto Score = LEAVE + CLASSIFIED×pts + OVERFLOW×pts + PATTERN×pts
      '=IF(' + gate + ',"",IF(E' + row + '="Yes",' + PTS_LEAVE + ',0)+F' + row + '*' + PTS_CLASSIFIED + '+G' + row + '*' + PTS_OVERFLOW + '+P' + row + '*' + PTS_PATTERN + ')',
      // R: TeleOp PATTERN Count — same logic as Auto but with TeleOp RAMP colors
      '=IF(' + gate + ',"",IF(LEN(L' + row + ')=0,0,' +
        'SUMPRODUCT((MID(UPPER(L' + row + '),SEQUENCE(LEN(L' + row + ')),1)=' +
        'MID(REPT(D' + row + ',3),SEQUENCE(LEN(L' + row + ')),1))*1)))',
      // S: TeleOp Score = CLASSIFIED×pts + OVERFLOW×pts + DEPOT×pts + PATTERN×pts + BASE
      '=IF(' + gate + ',"",I' + row + '*' + PTS_CLASSIFIED + '+J' + row + '*' + PTS_OVERFLOW + '+K' + row + '*' + PTS_DEPOT + '+R' + row + '*' + PTS_PATTERN + '+IF(M' + row + '="Full",' + PTS_BASE_FULL + ',IF(M' + row + '="Partial",' + PTS_BASE_PARTIAL + ',0)))',
      // T: Foul Deduction = Minor×pts + Major×pts
      '=IF(' + gate + ',"",N' + row + '*' + PTS_MINOR_FOUL + '+O' + row + '*' + PTS_MAJOR_FOUL + ')',
      // U: Score without Fouls = Auto + TeleOp
      '=IF(' + gate + ',"",Q' + row + '+S' + row + ')',
      // V: TOTAL SCORE = max(0, score - fouls)
      '=IF(' + gate + ',"",MAX(0,U' + row + '-T' + row + '))'
    ]);
  }
  sheet.getRange(REF_DATA_START, 2, MAX_TEAMS, 1).setFormulas(formulasB);
  sheet.getRange(REF_DATA_START, 3, MAX_TEAMS, 1).setFormulas(formulasC);
  sheet.getRange(REF_DATA_START, RC.AUTO_PAT, MAX_TEAMS, 7).setFormulas(formulasPV);

  // ---- DATA VALIDATION ----
  sheet.getRange("D" + REF_DATA_START + ":D" + REF_DATA_END).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(MOTIFS, true)
      .setAllowInvalid(false)
      .setHelpText("REQUIRED: Select the MOTIF shown on the OBELISK (" + MOTIFS.join(", ") + "). Selecting a value marks this row as scored.")
      .build()
  );

  sheet.getRange("E" + REF_DATA_START + ":E" + REF_DATA_END).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(LEAVE_OPTIONS, true)
      .setAllowInvalid(false)
      .setHelpText("Did the robot LEAVE? Robot must no longer be over any LAUNCH LINE at the end of AUTO.")
      .build()
  );

  sheet.getRange("M" + REF_DATA_START + ":M" + REF_DATA_END).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(BASE_OPTIONS, true)
      .setAllowInvalid(false)
      .setHelpText("Robot position on BASE TILE at end of TELEOP: " + BASE_OPTIONS.join(", ") + ".")
      .build()
  );

  function intRule(colLetter, helpText) {
    return SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(' + colLetter + REF_DATA_START + '="",AND(ISNUMBER(' + colLetter + REF_DATA_START + '),' + colLetter + REF_DATA_START + '>=0,INT(' + colLetter + REF_DATA_START + ')=' + colLetter + REF_DATA_START + '))')
      .setAllowInvalid(false)
      .setHelpText(helpText)
      .build();
  }

  var intFields = [
    ["F", "Count of CLASSIFIED artifacts during AUTO (whole number \u2265 0). Artifacts that passed through the SQUARE directly to the RAMP."],
    ["I", "Count of CLASSIFIED artifacts during TELEOP (whole number \u2265 0). Artifacts that passed through the SQUARE directly to the RAMP."],
    ["G", "Count of OVERFLOW artifacts during AUTO (whole number \u2265 0). Artifacts that passed through the SQUARE but did not go directly to the RAMP."],
    ["J", "Count of OVERFLOW artifacts during TELEOP (whole number \u2265 0). Artifacts that passed through the SQUARE but did not go directly to the RAMP."],
    ["K", "Count of DEPOT artifacts at end of TELEOP (whole number \u2265 0). Artifacts over the DEPOT tape."],
    ["N", "Number of MINOR fouls (whole number \u2265 0). Each MINOR foul = " + PTS_MINOR_FOUL + " point deduction."],
    ["O", "Number of MAJOR fouls (whole number \u2265 0). Each MAJOR foul = " + PTS_MAJOR_FOUL + " point deduction."]
  ];
  for (var f = 0; f < intFields.length; f++) {
    sheet.getRange(intFields[f][0] + REF_DATA_START + ":" + intFields[f][0] + REF_DATA_END)
      .setDataValidation(intRule(intFields[f][0], intFields[f][1]));
  }

  // RAMP Colors (G/P characters, max 9)
  var rampCols = [
    ["H", "Enter artifact colors on the RAMP at end of AUTO, in order from GATE to SQUARE. Use G (green) or P (purple). Max " + RAMP_MAX_CHARS + " characters. Case-insensitive."],
    ["L", "Enter artifact colors on the RAMP at end of TELEOP, in order from GATE to SQUARE. Use G (green) or P (purple). Max " + RAMP_MAX_CHARS + " characters. Case-insensitive."]
  ];
  for (var rc = 0; rc < rampCols.length; rc++) {
    var col = rampCols[rc][0];
    sheet.getRange(col + REF_DATA_START + ":" + col + REF_DATA_END).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied('=OR(' + col + REF_DATA_START + '="",REGEXMATCH(UPPER(' + col + REF_DATA_START + '),"' + RAMP_REGEX + '"))')
        .setAllowInvalid(false)
        .setHelpText(rampCols[rc][1])
        .build()
    );
  }

  // ---- FORMATTING ----
  sheet.getRange("A" + REF_DATA_START + ":A" + REF_DATA_END).setBackground("#F2F2F2");
  sheet.getRange("B" + REF_DATA_START + ":C" + REF_DATA_END).setBackground("#E8F0FE");
  sheet.getRange("D" + REF_DATA_START + ":D" + REF_DATA_END).setBackground("#E2D9F3");
  sheet.getRange("E" + REF_DATA_START + ":G" + REF_DATA_END).setBackground("#FFF2CC");
  sheet.getRange("H" + REF_DATA_START + ":H" + REF_DATA_END).setBackground("#E2EFDA");
  sheet.getRange("I" + REF_DATA_START + ":K" + REF_DATA_END).setBackground("#FFF2CC");
  sheet.getRange("L" + REF_DATA_START + ":L" + REF_DATA_END).setBackground("#FDF2E9");
  sheet.getRange("M" + REF_DATA_START + ":M" + REF_DATA_END).setBackground("#FFF2CC");
  sheet.getRange("N" + REF_DATA_START + ":O" + REF_DATA_END).setBackground("#FFF2CC");
  sheet.getRange("P" + REF_DATA_START + ":V" + REF_DATA_END).setBackground("#D6E4F0");
  sheet.getRange("W" + REF_DATA_START + ":W" + REF_DATA_END).setBackground("#FFF2CC");

  sheet.getRange("V" + REF_DATA_START + ":V" + REF_DATA_END).setFontWeight("bold").setFontSize(11);
  sheet.getRange("H" + REF_DATA_START + ":H" + REF_DATA_END).setFontFamily("Courier New").setFontWeight("bold");
  sheet.getRange("L" + REF_DATA_START + ":L" + REF_DATA_END).setFontFamily("Courier New").setFontWeight("bold");

  sheet.getRange("A" + REF_DATA_START + ":V" + REF_DATA_END).setHorizontalAlignment("center");
  sheet.getRange("B" + REF_DATA_START + ":B" + REF_DATA_END).setHorizontalAlignment("left");
  sheet.getRange("C" + REF_DATA_START + ":C" + REF_DATA_END).setHorizontalAlignment("left");
  sheet.getRange("W" + REF_DATA_START + ":W" + REF_DATA_END).setHorizontalAlignment("left");

  sheet.getRange("A3:W" + REF_DATA_END).setBorder(true, true, true, true, true, true,
    "#B4B4B4", SpreadsheetApp.BorderStyle.SOLID);

  var colWidths = [
    85, 150, 80, 80, 75, 90, 90, 100, 100, 90, 80, 100, 110, 75, 75, 85, 80, 85, 85, 80, 85, 90, 200
  ];
  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(2);

  // ---- CONDITIONAL FORMATTING ----
  var rules = [];

  // Fouls > 0
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#FFC7CE").setFontColor("#9C0006")
    .setRanges([sheet.getRange("N" + REF_DATA_START + ":O" + REF_DATA_END)])
    .build());

  // Incomplete entry (MOTIF set but required input empty) — stronger pink
  var reqCols = ["E","F","G","I","J","K","M","N","O"];
  for (var ci = 0; ci < reqCols.length; ci++) {
    var col = reqCols[ci];
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($D' + REF_DATA_START + '<>"",' + col + REF_DATA_START + '="")')
      .setBackground("#FFCCCC")
      .setRanges([sheet.getRange(col + REF_DATA_START + ":" + col + REF_DATA_END)])
      .build());
  }

  // Not yet scored (team# but no MOTIF)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + REF_DATA_START + '<>"",$D' + REF_DATA_START + '="")')
    .setBackground("#FDE9D9")
    .setRanges([sheet.getRange("A" + REF_DATA_START + ":W" + REF_DATA_END)])
    .build());

  // Zebra striping
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + REF_DATA_START + '<>"",ISEVEN(ROW()))')
    .setBackground("#F0F4FA")
    .setRanges([sheet.getRange("A" + REF_DATA_START + ":W" + REF_DATA_END)])
    .build());

  sheet.setConditionalFormatRules(rules);
}

// ============================================================
// FINAL SCORES SHEET (internal — called by buildAll)
// ============================================================
function _buildFinalScoresSheet(ss) {
  var oldSheet = ss.getSheetByName("FinalScores");
  var sheet = ss.insertSheet("FinalScores" + (oldSheet ? "_new" : ""));
  if (oldSheet) ss.deleteSheet(oldSheet);
  sheet.setName("FinalScores");

  // Column layout (A-X, 24 columns):
  // A=Team#, B=Name, C=Video,
  // D=Ref Names, E=Official Referee, F=Refs Agree?,
  // G=Final Score, H=Score w/o Fouls, I=Auto Score, J=TeleOp Score, K=Foul Deduction,
  // L=Minor Fouls, M=Major Fouls,
  // N=LEAVE, O=Auto CLS, P=Auto OVF, Q=Auto RAMP Colors, R=Auto PATTERN,
  // S=Tel CLS, T=Tel OVF, U=Tel DEPOT, V=Tel RAMP Colors, W=Tel PATTERN, X=BASE

  // ---- ROW 1: Category group headers (merged) ----
  var groups = [
    {range: "A1:C1", label: "Teams",            bg: "#2F5496"},
    {range: "D1:F1", label: "Referee",           bg: "#548235"},
    {range: "G1:K1", label: "Total Scores",      bg: "#7030A0"},
    {range: "L1:M1", label: "Fouls",             bg: "#C00000"},
    {range: "N1:R1", label: "Autonomous Period",  bg: "#8B6914"},
    {range: "S1:X1", label: "TeleOp Period",      bg: "#8B6914"}
  ];
  for (var g = 0; g < groups.length; g++) {
    sheet.getRange(groups[g].range).merge()
      .setValue(groups[g].label)
      .setFontWeight("bold").setFontColor("white").setHorizontalAlignment("center")
      .setBackground(groups[g].bg).setFontSize(11);
  }

  // ---- ROW 2: Instructions (split merge at frozen column boundary) + Point values (L-X) ----
  sheet.getRange("A2:C2").merge().setBackground("#FFF5D6");
  sheet.getRange("D2:F2").merge().setValue(
    "Select an Official Referee for each team. " +
    "\"Refs Agree?\" shows Yes when all referees match on every scoring element."
  ).setFontStyle("italic").setFontColor("#6B4400").setBackground("#FFF5D6")
   .setHorizontalAlignment("left").setWrap(true);

  var pointValues = [
    "", "", "",
    "", "", "",
    "", "", "", "", "",
    "\u00d7(\u2212" + PTS_MINOR_FOUL + ")", "\u00d7(\u2212" + PTS_MAJOR_FOUL + ")",
    String(PTS_LEAVE), "\u00d7" + PTS_CLASSIFIED, "\u00d7" + PTS_OVERFLOW, "", "\u00d7" + PTS_PATTERN + " ea",
    "\u00d7" + PTS_CLASSIFIED, "\u00d7" + PTS_OVERFLOW, "\u00d7" + PTS_DEPOT, "", "\u00d7" + PTS_PATTERN + " ea",
    PTS_BASE_PARTIAL + "/" + PTS_BASE_FULL
  ];
  for (var c = 6; c < pointValues.length; c++) {
    if (pointValues[c] !== "") sheet.getRange(2, c + 1).setValue(pointValues[c]);
  }
  sheet.getRange("G2:X2").setFontWeight("bold").setHorizontalAlignment("center")
    .setFontSize(10).setFontColor("#505050").setBackground("#E8E8E8");

  // ---- ROW 3: Column headers ----
  var headers = [
    "Number", "Name", "Video",
    "Scored By", "Official\nReferee", "Refs\nAgree?",
    "Final\nScore", "Score w/o\nFouls", "Auto\nScore", "TeleOp\nScore", "Foul\nDeduction",
    "Minor", "Major",
    "LEAVE", "CLASSIFIED", "OVERFLOW", "RAMP\nColors", "PATTERN\nCount",
    "CLASSIFIED", "OVERFLOW", "DEPOT", "RAMP\nColors", "PATTERN\nCount", "BASE"
  ];
  for (var c = 0; c < headers.length; c++) {
    sheet.getRange(3, c + 1).setValue(headers[c]);
  }
  sheet.getRange("A3:X3").setFontWeight("bold").setWrap(true).setVerticalAlignment("middle")
    .setHorizontalAlignment("center").setFontColor("white");
  sheet.setRowHeight(3, 40);

  sheet.getRange("A3:C3").setBackground("#2F5496");
  sheet.getRange("D3:F3").setBackground("#548235");
  sheet.getRange("G3:K3").setBackground("#7030A0");
  sheet.getRange("L3:M3").setBackground("#C00000");
  sheet.getRange("N3:R3").setBackground("#8B6914");
  sheet.getRange("S3:X3").setBackground("#8B6914");

  // ---- DATA ROWS ----
  function indRef(r) {
    return "INDIRECT(\"'\"&Config!" + _refConfigCol(r) + "$2&\"'!$A:$W\")";
  }

  // Mapping: FinalScores destination column -> referee sheet source column index
  var vlookupMap = [
    [7,  RC.TOTAL],          // G: Final Score
    [8,  RC.SCORE_NO_FOULS], // H: Score w/o Fouls
    [9,  RC.AUTO_SCORE],     // I: Auto Score
    [10, RC.TEL_SCORE],      // J: TeleOp Score
    [11, RC.FOUL_DED],       // K: Foul Deduction
    [12, RC.MINOR],          // L: Minor Fouls
    [13, RC.MAJOR],          // M: Major Fouls
    [14, RC.LEAVE],          // N: LEAVE
    [15, RC.AUTO_CLS],       // O: Auto CLASSIFIED
    [16, RC.AUTO_OVF],       // P: Auto OVERFLOW
    [17, RC.AUTO_RAMP],      // Q: Auto RAMP Colors
    [18, RC.AUTO_PAT],       // R: Auto PATTERN
    [19, RC.TEL_CLS],        // S: TeleOp CLASSIFIED
    [20, RC.TEL_OVF],        // T: TeleOp OVERFLOW
    [21, RC.TEL_DEPOT],      // U: TeleOp DEPOT
    [22, RC.TEL_RAMP],       // V: TeleOp RAMP Colors
    [23, RC.TEL_PAT],        // W: TeleOp PATTERN
    [24, RC.BASE]            // X: BASE
  ];

  // Agreement check elements: all 12 input columns (MOTIF through BASE, plus fouls)
  var elemCols = [RC.MOTIF, RC.LEAVE, RC.AUTO_CLS, RC.AUTO_OVF, RC.AUTO_RAMP,
                  RC.TEL_CLS, RC.TEL_OVF, RC.TEL_DEPOT, RC.TEL_RAMP, RC.BASE,
                  RC.MINOR, RC.MAJOR];
  var numericCols = [RC.AUTO_CLS, RC.AUTO_OVF, RC.TEL_CLS, RC.TEL_OVF, RC.TEL_DEPOT,
                     RC.MINOR, RC.MAJOR];

  // Build all formulas as arrays for batch write
  var formulasAD = []; // cols A-D
  var formulasF = [];  // col F (Refs Agree?)
  var formulasGX = []; // cols G-X

  for (var row = FS_DATA_START; row <= FS_DATA_END; row++) {
    // A: Team # from Config
    var fA = '=IF(Config!A' + row + '="","",Config!A' + row + ')';
    // B: Team Name
    var fB = '=IF(A' + row + '="","",IFERROR(VLOOKUP(A' + row + ',Config!$A:$C,2,FALSE),""))';
    // C: Video
    var fC = '=IF(A' + row + '="","",IFERROR(VLOOKUP(A' + row + ',Config!$A:$C,3,FALSE),""))';
    // D: Referee names who scored (multiline) — wrapped in IFERROR for empty-data edge cases
    var refNameParts = [];
    for (var r = 1; r <= NUM_REFEREES; r++) {
      refNameParts.push(
        'IF(IFERROR(VLOOKUP($A' + row + ',' + indRef(r) + ',' + RC.MOTIF + ',FALSE),"")="",' +
        '"",Config!' + _refConfigCol(r) + '$2)'
      );
    }
    var fD = '=IF(A' + row + '="","",IFERROR(TEXTJOIN(CHAR(10),TRUE,' + refNameParts.join(',') + '),""))';
    formulasAD.push([fA, fB, fC, fD]);

    // F: Refs Agree? — count refs and check agreement on all input elements including fouls
    //    Wrapped in IFERROR to handle edge cases when no data exists yet
    var refCountParts = [];
    for (var r = 1; r <= NUM_REFEREES; r++) {
      refCountParts.push(
        'IF(IFERROR(VLOOKUP($A' + row + ',' + indRef(r) + ',' + RC.MOTIF + ',FALSE),"")="",0,1)'
      );
    }
    var refCountExpr = '(' + refCountParts.join('+') + ')';

    var matchParts = [];
    for (var e = 0; e < elemCols.length; e++) {
      var ec = elemCols[e];
      var isNum = numericCols.indexOf(ec) !== -1;
      var vParts = [];
      for (var r = 1; r <= NUM_REFEREES; r++) {
        var valExpr = 'IFERROR(VLOOKUP($A' + row + ',' + indRef(r) + ',' + ec + ',FALSE),"")';
        var motExpr = 'IFERROR(VLOOKUP($A' + row + ',' + indRef(r) + ',' + RC.MOTIF + ',FALSE),"")';
        if (isNum) {
          vParts.push('IF(' + motExpr + '="","",IF(' + valExpr + '="","(blank)",' + valExpr + '))');
        } else {
          vParts.push('IF(' + motExpr + '="","",IF(' + valExpr + '="","(blank)",UPPER(' + valExpr + ')))');
        }
      }
      matchParts.push(
        'IFERROR(ROWS(UNIQUE(FILTER({' + vParts.join(';') + '},{' + vParts.join(';') + '}<>"")))=1,TRUE)'
      );
    }

    formulasF.push([
      '=IF($A' + row + '="","",IFERROR(IF(' + refCountExpr + '<2,"N/A",' +
      'IF(AND(' + matchParts.join(',') + '),"Yes","No")),"N/A"))'
    ]);

    // G-X: Override referee VLOOKUP (loop over mapping)
    var overrideRef = 'INDIRECT("\'"&$E' + row + '&"\'!$A:$W")';
    var rowFormulas = [];
    for (var v = 0; v < vlookupMap.length; v++) {
      var srcCol = vlookupMap[v][1];
      rowFormulas.push(
        '=IF(OR($A' + row + '="",$E' + row + '=""),"",IFERROR(VLOOKUP($A' + row + ',' + overrideRef + ',' + srcCol + ',FALSE),""))'
      );
    }
    formulasGX.push(rowFormulas);
  }

  // Batch write all formulas
  sheet.getRange(FS_DATA_START, 1, MAX_TEAMS, 4).setFormulas(formulasAD);
  sheet.getRange(FS_DATA_START, 6, MAX_TEAMS, 1).setFormulas(formulasF);
  sheet.getRange(FS_DATA_START, 7, MAX_TEAMS, 18).setFormulas(formulasGX);

  // ---- DATA VALIDATION ----
  var configSheet = ss.getSheetByName("Config");
  if (configSheet) {
    sheet.getRange("E" + FS_DATA_START + ":E" + FS_DATA_END).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(configSheet.getRange("D2:I2"), true)
        .setAllowInvalid(false)
        .setHelpText("Select the referee whose scores should be used as the official record for this team. When referees disagree, review their individual sheets and pick the most accurate.")
        .build()
    );
  }

  // ---- FORMATTING ----
  sheet.getRange("A" + FS_DATA_START + ":A" + FS_DATA_END).setBackground("#F2F2F2");
  sheet.getRange("B" + FS_DATA_START + ":C" + FS_DATA_END).setBackground("#E8F0FE");
  sheet.getRange("D" + FS_DATA_START + ":D" + FS_DATA_END).setBackground("#E2EFDA");
  sheet.getRange("E" + FS_DATA_START + ":E" + FS_DATA_END).setBackground("#FFF2CC");
  sheet.getRange("G" + FS_DATA_START + ":G" + FS_DATA_END).setFontWeight("bold").setFontSize(11).setBackground("#E2D9F3");
  sheet.getRange("H" + FS_DATA_START + ":K" + FS_DATA_END).setBackground("#F3EDF9");
  sheet.getRange("L" + FS_DATA_START + ":M" + FS_DATA_END).setBackground("#FCE4EC");
  sheet.getRange("N" + FS_DATA_START + ":R" + FS_DATA_END).setBackground("#FFF8E1");
  sheet.getRange("S" + FS_DATA_START + ":X" + FS_DATA_END).setBackground("#FFF8E1");

  sheet.getRange("A" + FS_DATA_START + ":X" + FS_DATA_END).setHorizontalAlignment("center");
  sheet.getRange("B" + FS_DATA_START + ":B" + FS_DATA_END).setHorizontalAlignment("left");
  sheet.getRange("C" + FS_DATA_START + ":C" + FS_DATA_END).setHorizontalAlignment("left");
  sheet.getRange("D" + FS_DATA_START + ":D" + FS_DATA_END).setHorizontalAlignment("left").setWrap(true);

  sheet.getRange("A3:X" + FS_DATA_END).setBorder(true, true, true, true, true, true,
    "#B4B4B4", SpreadsheetApp.BorderStyle.SOLID);

  var colWidths = [
    85, 150, 80, 120, 115, 85, 85, 85, 75, 80, 80, 65, 65, 65, 85, 85, 85, 80, 85, 85, 65, 85, 80, 65
  ];
  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // ---- CONDITIONAL FORMATTING ----
  var rules = [];

  var matchRange = [sheet.getRange("F" + FS_DATA_START + ":F" + FS_DATA_END)];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Yes").setBackground("#C6EFCE").setFontColor("#006100").setBold(true)
    .setRanges(matchRange).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("No").setBackground("#FFC7CE").setFontColor("#9C0006").setBold(true)
    .setRanges(matchRange).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("N/A").setBackground("#F2F2F2").setFontColor("#5A5A5A").setBold(true)
    .setRanges(matchRange).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + FS_DATA_START + '<>"",$E' + FS_DATA_START + '="")')
    .setBackground("#FDE9D9")
    .setRanges([sheet.getRange("E" + FS_DATA_START + ":E" + FS_DATA_END)])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + FS_DATA_START + '<>"",ISEVEN(ROW()))')
    .setBackground("#F0F4FA")
    .setRanges([sheet.getRange("A" + FS_DATA_START + ":X" + FS_DATA_END)])
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("randomizeTeamOrders must be run from the " + GAME_NAME + " Scoring menu, not the script editor.");
    return;
  }
  var config = ss.getSheetByName("Config");
  if (!config) {
    ui.alert("Config sheet not found. Run 'Rebuild All Sheets' first.");
    return;
  }

  if (_hasAnyScoringData(ss)) {
    var response = ui.alert(
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

  var teamRange = config.getRange("A4:A" + (MAX_TEAMS + 3));
  var teamValues = teamRange.getValues();
  var teams = [];
  for (var i = 0; i < teamValues.length; i++) {
    if (teamValues[i][0] !== "" && teamValues[i][0] !== null) {
      teams.push(teamValues[i][0]);
    }
  }

  if (teams.length === 0) {
    ui.alert("No team numbers found in Config column A (starting row 4).\nPlease enter team numbers first.");
    return;
  }

  for (var r = 1; r <= NUM_REFEREES; r++) {
    // Fisher-Yates shuffle
    var shuffled = teams.slice();
    for (var i = shuffled.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var temp = shuffled[i];
      shuffled[i] = shuffled[j];
      shuffled[j] = temp;
    }

    // Batch write to Config randomized order column (J-O = cols 10-15)
    var orderCol = r + 9;
    var orderValues = [];
    for (var i = 0; i < MAX_TEAMS; i++) {
      orderValues.push([i < shuffled.length ? shuffled[i] : ""]);
    }
    config.getRange(4, orderCol, MAX_TEAMS, 1).setValues(orderValues);

    // Batch write to referee sheet (reuse orderValues — same data)
    var refSheet = findRefSheet(ss, config, r);
    if (refSheet) {
      refSheet.getRange(REF_DATA_START, 1, MAX_TEAMS, 1).setValues(orderValues);
    }
  }

  SpreadsheetApp.flush();
  ui.alert(
    "Randomization Complete",
    "Team orders randomized for all " + NUM_REFEREES + " referees.\n" +
    "Orders are saved in Config columns J-O and on each referee sheet.\n\n" +
    "Do NOT re-randomize after referees begin scoring!",
    ui.ButtonSet.OK
  );
}

/** Check whether any referee sheet contains actual scoring data (valid MOTIF values). */
function _hasAnyScoringData(ss) {
  var config = ss.getSheetByName("Config");
  var validMotifs = {};
  for (var i = 0; i < MOTIFS.length; i++) validMotifs[MOTIFS[i]] = true;

  for (var r = 1; r <= NUM_REFEREES; r++) {
    var refSheet = findRefSheet(ss, config, r);
    if (!refSheet) continue;
    // Check rows 4 through MAX_TEAMS+4 to handle both old and new layout
    var motifData = refSheet.getRange("D4:D" + (MAX_TEAMS + 4)).getValues();
    for (var i = 0; i < motifData.length; i++) {
      if (validMotifs[motifData[i][0]]) return true;
    }
  }
  return false;
}

// ============================================================
// RENAME REFEREE SHEETS
// ============================================================
function renameRefSheets() {
  if (!checkAuthorization()) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("renameRefSheets must be run from the " + GAME_NAME + " Scoring menu, not the script editor.");
    return;
  }
  var config = ss.getSheetByName("Config");
  if (!config) {
    ui.alert("Config sheet not found. Run 'Rebuild All Sheets' first.");
    return;
  }
  var renamed = _doRenameRefSheets(ss, config);
  _hideUnnamedRefSheets(ss, config);
  ui.alert("Rename Complete", renamed + " sheet(s) renamed to match Config names.", ui.ButtonSet.OK);
}

/**
 * Two-phase rename to handle swapped names correctly.
 * Phase 1: Rename all sheets that need changing to temporary names.
 * Phase 2: Rename from temporary names to final desired names.
 */
function _doRenameRefSheets(ss, config) {
  var renamed = 0;
  var renames = []; // [{sheet, desiredName}]

  // Collect sheets that need renaming
  for (var r = 1; r <= NUM_REFEREES; r++) {
    var desiredName = getRefSheetName(config, r);
    var sheet = findRefSheet(ss, config, r);
    if (sheet && sheet.getName() !== desiredName) {
      renames.push({sheet: sheet, desiredName: desiredName, refNum: r});
    }
  }

  if (renames.length === 0) return 0;

  // Phase 1: Rename to temporary names to avoid collisions
  for (var i = 0; i < renames.length; i++) {
    var tempName = "_temp_rename_" + renames[i].refNum;
    try {
      renames[i].sheet.setName(tempName);
    } catch(e) {
      Logger.log("Failed to rename " + renames[i].sheet.getName() + " to temp: " + e);
    }
  }

  // Phase 2: Rename to final names
  for (var i = 0; i < renames.length; i++) {
    try {
      renames[i].sheet.setName(renames[i].desiredName);
      renames[i].sheet.getRange("C1").setValue(GAME_NAME + " " + SEASON + " Match Review \u2014 " + renames[i].desiredName);
      renamed++;
    } catch(e) {
      Logger.log("Failed to rename to " + renames[i].desiredName + ": " + e);
      // Restore to a safe fallback name
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {
    Logger.log("applyProtection must be run from the " + GAME_NAME + " Scoring menu, not the script editor.");
    return;
  }
  var me = Session.getEffectiveUser();
  var meEmail = (me.getEmail() || "").toLowerCase();
  var config = ss.getSheetByName("Config");

  var refEmails = [];
  if (config) {
    for (var r = 1; r <= NUM_REFEREES; r++) {
      var email = config.getRange(_refConfigCol(r) + "3").getValue();
      refEmails.push(email ? email.toString().trim().toLowerCase() : "");
    }
  }

  var hasEmails = false;
  for (var i = 0; i < refEmails.length; i++) {
    if (refEmails[i] !== "" && refEmails[i].indexOf("@") !== -1) { hasEmails = true; break; }
  }

  var failedEmails = [];

  // ---- Referee sheets ----
  for (var r = 1; r <= NUM_REFEREES; r++) {
    var sheet = findRefSheet(ss, config, r);
    if (!sheet) continue;

    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); });
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });

    var sheetProt = sheet.protect().setDescription(sheet.getName() + " Sheet");
    sheetProt.addEditor(me);
    _restrictEditors(sheetProt, [meEmail]);

    var inputRange1 = sheet.getRange("D" + REF_DATA_START + ":O" + REF_DATA_END);
    var inputRange2 = sheet.getRange("W" + REF_DATA_START + ":W" + REF_DATA_END);
    sheetProt.setUnprotectedRanges([inputRange1, inputRange2]);

    if (hasEmails && refEmails[r - 1] !== "" && refEmails[r - 1].indexOf("@") !== -1) {
      var refEmail = refEmails[r - 1];

      var rangeProt1 = inputRange1.protect().setDescription(sheet.getName() + " Scoring Input");
      rangeProt1.addEditor(me);
      try { rangeProt1.addEditor(refEmail); } catch(e) {
        failedEmails.push(refEmail + " (" + sheet.getName() + ")");
      }
      _restrictEditors(rangeProt1, [meEmail, refEmail]);

      var rangeProt2 = inputRange2.protect().setDescription(sheet.getName() + " Notes");
      rangeProt2.addEditor(me);
      try { rangeProt2.addEditor(refEmail); } catch(e) {}
      _restrictEditors(rangeProt2, [meEmail, refEmail]);
    } else {
      sheetProt.setWarningOnly(true);
    }
  }

  // ---- FinalScores ----
  var finalSheet = ss.getSheetByName("FinalScores");
  if (finalSheet) {
    finalSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); });
    finalSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });

    var protection = finalSheet.protect().setDescription("FinalScores - Protected");
    protection.addEditor(me);
    _restrictEditors(protection, [meEmail]);

    var overrideNameRange = finalSheet.getRange("E" + FS_DATA_START + ":E" + FS_DATA_END);
    protection.setUnprotectedRanges([overrideNameRange]);

    var rangeProt = overrideNameRange.protect().setDescription("Official Referee Selection");
    rangeProt.addEditor(me);
    _restrictEditors(rangeProt, [meEmail]);
  }

  // ---- Config ----
  if (config) {
    config.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); });
    config.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });

    var protection = config.protect().setDescription("Config - Protected");
    protection.addEditor(me);
    _restrictEditors(protection, [meEmail]);

    var teamDataRange = config.getRange("A4:C" + (MAX_TEAMS + 3));
    var refInfoRange = config.getRange("D2:I3");
    protection.setUnprotectedRanges([teamDataRange, refInfoRange]);

    var teamProt = teamDataRange.protect().setDescription("Config - Team Data");
    teamProt.addEditor(me);
    _restrictEditors(teamProt, [meEmail]);

    var refInfoProt = refInfoRange.protect().setDescription("Config - Referee Info");
    refInfoProt.addEditor(me);
    _restrictEditors(refInfoProt, [meEmail]);
  }

  // Hide Config sheet — referees don't need it; admin can unhide via sheet tab right-click
  if (config) {
    try { config.hideSheet(); } catch(e) {
      Logger.log("Could not hide Config sheet: " + e);
    }
  }

  // Hide unnamed referee sheets (unused referee slots)
  _hideUnnamedRefSheets(ss, config);

  var msg;
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
      "1. Unhide Config, enter referee emails in row 3\n" +
      "2. Re-run " + GAME_NAME + " Scoring > Apply Sheet Protection";
  }

  ui.alert("Protection Applied", msg, ui.ButtonSet.OK);
}
