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
// Each entry: "G{num} - {title}. {full text including violation details}"
// Used for dropdown validation on referee sheets and FinalScores.
// The onEdit trigger extracts the first 4 characters as the rule code.
const G_RULES = [
  'G101 - Humans, stay off the FIELD during the MATCH. Other than actions explicitly allowed in section 11.4.6 Human, a DRIVE TEAM member may only enter the FIELD during the following times: A. pre-MATCH set-up in order to place their ROBOT and pre-loaded SCORING ELEMENTS per G301, G303, and G304, and B. after a MATCH is over to stop and collect their ROBOT in a reasonable amount of time when instructed to do so by the Head REFEREE or their designee. Violation: VERBAL WARNING. A team may not delay the FIELD reset process through an excessively lengthy process to remove the ROBOT from the FIELD. It is not a violation of this rule if DRIVE TEAM members contribute to FIELD reset by placing SCORING ELEMENTS that they inadvertently move while setting up their ROBOT or placing removed SCORING ELEMENTS on the FIELD. Egregious violations of this rule, such as entering the FIELD during a MATCH, are covered by G211.',
  'G102 - Be careful when interacting with ARENA elements. A team member is prohibited from the following actions with regards to interaction with ARENA elements: A. climbing on, B. hanging from, C. manipulating such that it does not return to its original shape without human intervention, and D. damaging. Violation: VERBAL WARNING. YELLOW CARD if subsequent violations occur during the event. DRIVE TEAM members may brace the FIELD perimeter at any point during the MATCH. Moving the FIELD perimeter out of position is considered a violation of G102.C. 11.2 Conduct',
  'G201 - Be a good person. All teams must be civil toward everyone and respectful of team and event equipment while at a FIRST Tech Challenge event. Please review the FIRST Code of Conduct and Core Values for more information. Violation: VERBAL WARNING. YELLOW CARD if subsequent violations occur during the event. Examples of inappropriate behavior include, but are not limited to, the use of offensive language or other uncivil conduct. Examples of particularly contemptible behavior that is likely to result in ARENA ejection include, but are not limited to, the following: A. assault, e.g., throwing something that hits another person (even if unintended), B. threat, e.g., saying something like \'if you don\'t reverse that call, I\'ll make you regret it,\' C. harassment, e.g., badgering someone with no new information after a decision has been made or a question has been answered, D. bullying, e.g., using body or verbal language to cause another person to feel inadequate, E. insulting, e.g., telling someone they don\'t deserve to be on a DRIVE TEAM, F. swearing at another person (versus swearing under one\'s breath or at oneself), and G. yelling at another person(s) in anger or frustration.',
  'G202 - DRIVE TEAM Interactions. DRIVE TEAM members cannot distract/interfere with the opposing ALLIANCE. This includes taunting or other disruptive behavior. Violation: VERBAL WARNING. YELLOW CARD if subsequent violations occur during the event.',
  'G203 - Asking other teams to throw a MATCH \' not cool. A team may not encourage an ALLIANCE of which it is not a member to play beneath its ability. NOTE: This rule is not intended to prevent an ALLIANCE from planning and/or executing its own strategy in a specific MATCH in which all the teams are members of the ALLIANCE. Violation: VERBAL WARNING. RED CARD if subsequent violations occur during the event. Example 1: A MATCH is being played by Teams A and B. Team C requests Team A to open the GATE at the end of the MATCH in order resulting in teams A and B not earning the PATTERN RP. Team A accepts this request from team C. Team C\'s motivation for this behavior is to prevent Team B from rising in the Tournament rankings and negatively affect Team C\'s ranking. Team C has violated this rule. Example 2: A MATCH is being played by teams A and B, in which team A is assigned to participate as a SURROGATE. Team D encourages team A not to participate in the MATCH so that team D gains ranking position over team B. Team D has violated this rule. FIRST considers the action of a team influencing another team to throw a MATCH, to deliberately miss RANKING POINTS, etc. incompatible with FIRST values and not a strategy any team should employ.',
  'G204 - Letting someone coerce you into throwing a MATCH \' also not cool. A team, as the result of encouragement by a team not on their ALLIANCE, may not play beneath its ability. NOTE: This rule is not intended to prevent an ALLIANCE from planning and/or executing its own strategy in a specific MATCH in which all the ALLIANCE members are participants. Violation: VERBAL WARNING. RED CARD if subsequent violations occur during the event. Example 1: A MATCH is being played by Teams A and B. Team C requests Team A to open the GATE at the end of the MATCH in order resulting in teams A and B not earning the PATTERN RP. Team A accepts this request from team C. Team C\'s motivation for this behavior is to prevent Team B from rising in the Tournament rankings and negatively affect Team C\'s ranking. Team A has violated this rule. Example 2: A MATCH is being played by Teams A and B, in which Team A is assigned to participate as a SURROGATE. Team A accepts Team D\'s request to not participate in the MATCH so that Team D gains ranking position over Team B. Team A has violated this rule. FIRST considers the action of a team influencing another team to throw a MATCH, to deliberately miss RANKING POINTS, etc. incompatible with FIRST values and not a strategy any team should employ.',
  'G205 - Throwing your own MATCH is bad. A team may not intentionally lose a MATCH or sacrifice RANKING POINTS in an effort to lower their own ranking and/or manipulate the rankings of other teams. Violation: VERBAL WARNING. RED CARD if subsequent violations occur during the event. The intent of this rule is not to punish teams who are employing alternate strategies, but rather to ensure that it is clear that throwing MATCHES to negatively affect your own rankings, or to manipulate the rankings of other teams (e.g., throw a MATCH to lower a partner\'s ranking, and/or increase the ranking of another team not in the MATCH) is incompatible with FIRST values and not a strategy any team should employ.',
  'G206 - Don\'t violate rules for RPs. A team or ALLIANCE may not collude with another team to each purposefully violate a rule in an attempt to influence RANKING POINTS. Violation: YELLOW CARD and the ALLIANCE is ineligible for PATTERN and GOAL RPs For example, if Team A on the blue ALLIANCE agrees with Team D on the red ALLIANCE to disrupt each other\'s GATE in violation of G417 resulting in both ALLIANCES being awarded the PATTERN RP.',
  'G207 - Do not abuse ARENA access. A team member (except those DRIVE TEAM members on the DRIVE TEAM for the MATCH) granted access to restricted areas in and around the ARENA (e.g., via event issued media badges) may not assist, coach, or use signaling devices during the MATCH. Exceptions will be granted for inconsequential infractions and in cases concerning safety. Violation: VERBAL WARNING. YELLOW CARD if subsequent violations occur during the event. Team members in open-access spectator seating areas are not considered to be in a restricted area and are not prevented from assisting or using signaling devices. See E102 for related details.',
  'G208 - Show up to your MATCHES. If a ROBOT has passed initial, complete inspection, at least 1 member of its DRIVE TEAM must report to the ARENA and participate in each of their assigned Qualification MATCHES. Violation: DISQUALIFIED from the current MATCH. The team should inform the Lead Queuer if the team\'s ROBOT is not able to participate.',
  'G209 - Keep your ROBOT together. A ROBOT may not intentionally detach or leave a part on the FIELD. Violation: RED CARD.',
  'G210 - Do not expect to gain by doing others harm. Actions clearly aimed at forcing the opponent ALLIANCE to violate a rule are not in the spirit of FIRST Tech Challenge and not allowed. Rule violations forced in this manner will not result in an assignment of a penalty to the targeted ALLIANCE. Violation: MINOR FOUL. MAJOR FOUL if REPEATED. The ALLIANCE that was forced to break a rule will not be assessed a penalty. This rule does not apply for strategies consistent with standard gameplay, for example: A. a red ROBOT attempting to access its GATE pushes a blue ROBOT into an ARTIFACT on the red RAMP. This rule requires an intentional act with limited or no opportunity for the team being acted on to avoid the penalty, such as: B. a blue ALLIANCE ROBOT pushing a red ALLIANCE ROBOT from \'far away\' (more than one TILE distance away) into the blue ALLIANCE LOADING ZONE. \' C. Placing an ARTIFACT into an opponent ROBOT such that it is in violation of G408.',
  'G211 - Egregious or exceptional violations. Egregious behavior beyond what is listed in the rules or subsequent violations of any rule or procedure during the event is prohibited. In addition to rule violations explicitly listed in this manual and witnessed by a REFEREE, the Head REFEREE may assign a YELLOW or RED CARD for egregious ROBOT actions or team member behavior at any time during the event. Continued violations will be brought to FIRST Headquarters\' attention. FIRST Headquarters will work with event staff to determine if further escalations are necessary, which can include removal from award consideration and removal from the event. Please see section 10.6.1 YELLOW and RED CARDS for additional detail. Violation: YELLOW or RED CARD. The intent of this rule is to provide the Head REFEREES with the flexibility necessary to keep the event running smoothly, as well as keep the safety of all the participants as the highest priority. There are certain behaviors that automatically result in a YELLOW or RED CARD because this behavior puts the FIRST community at risk. Those behaviors include, but are not limited to the list below: A. inappropriate behavior as outlined in the orange box of G201, B. reaching into the FIELD and grabbing a ROBOT during a MATCH, C. a single PIN in excess of 15 seconds, D. descoring SCORING ELEMENTS strategically or REPEATEDLY The Head REFEREE may assign a YELLOW or RED CARD for a single instance of a rule violation such as the examples given in items above, or for multiple instances of any single rule violation. Teams should be aware that any rule in this manual could escalate to a YELLOW or RED CARD. The Head REFEREE has final authority on all rules and violations at an event.',
  'G212 - All teams can play. A team may not encourage another team to exclude their ROBOT or be DISQUALIFIED from a Qualification MATCH for any reason. Violation: YELLOW CARD. RED CARD if the ROBOT does not participate in the MATCH 11.3 Pre-MATCH',
  'G301 - Be prompt. A DRIVE TEAM member may not cause significant delays to the start of their MATCH. Causing a significant delay requires both of the following to be true: A. The expected MATCH start time has passed, and During Qualification MATCHES, the expected start time of the MATCH is the time indicated on the MATCH schedule or ~3 minutes from the end of the previous MATCH on the same FIELD, whichever is later. If T206 is in effect, the expected MATCH start time is the later of the end of the T206 time or the time indicated on the schedule. During Playoff MATCHES, the expected start time of the MATCH is the time indicated on the MATCH schedule or 8 minutes from either ALLIANCE\'S previous MATCH, whichever is later. B. The DRIVE TEAM has access to the ARENA and is neither MATCH ready nor making a good faith effort, as perceived by the Head REFEREE, to quickly become MATCH ready. Teams that have violated G208 or have 1 DRIVE TEAM member present and have informed event staff that their ROBOT will not be participating in the MATCH are considered MATCH ready and not in violation of this rule. Violation: If a Qualification MATCH: VERBAL WARNING. MAJOR FOUL for the upcoming MATCH if a subsequent violation occurs within the tournament phase. If the DRIVE TEAM is not MATCH ready within 2 minutes of the expected MATCH start time, the ROBOT will be DISABLED.',
  'G302 - Limit what you bring to the FIELD. Items brought to the FIELD to be used for a MATCH, in addition to the ROBOT, OPERATOR CONSOLE, must fit in the team\'s designated ALLIANCE AREA, be worn or held by members of the DRIVE TEAM, or be an item used as an accommodation (e.g., single-step stools that do not roll/fold, crutches, cushion, kneeling mat,). Regardless of if the equipment fits the criteria above, it may not: A. be employed in a way that introduces a safety hazard, B. extend more than 6 ft. 6 in. (~198 cm) above the TILES, C. communicate with anything or anyone outside of the ARENA with the exception of medically required equipment, D. block visibility for FIELD STAFF or audience members, or E. jam or interfere with anything in the ARENA. Violation: MATCH will not start until the situation is remedied. YELLOW CARD, if discovered or used inappropriately during a MATCH. It is not a violation of this rule to bring an alignment device to the FIELD to aid pre-MATCH ROBOT set-up and alignment. The use of any alignment devices should not delay MATCH start in violation of G301. Examples of equipment that may be considered a safety hazard in the confined space of the ALLIANCE AREA include but are not limited to, a folding step stool, ladder, or a large signaling device. Using an item that has wireless communications disabled complies with G302.C above. Examples of jamming or interfering with remote sensing capabilities include, but are not limited to, mimicking the FIELD AprilTags and shining bright lighting or laser pointers onto the FIELD.',
  'G303 - ROBOTS on the FIELD must come ready to play a MATCH. A ROBOT must meet all following MATCH-start requirements: A. does not pose a hazard to humans, FIELD elements, or other ROBOTS. B. has passed inspection, i.e., it is compliant with all ROBOT rules. C. if modified after initial Inspection, it is compliant with I305. D. is the only team-provided item left in the FIELD. E. ROBOT SIGNS must indicate the correct ALLIANCE color (see R402). F. ROBOT must be motionless following completion of OpMode initialization. If a ROBOT is DISABLED prior to the start of the MATCH, the DRIVE TEAM may not remove the ROBOT from the FIELD without permission from the Head REFEREE or the FTA. For assessment of many of the items listed above, the Head REFEREE is likely to consult with the LRI. Violation: The MATCH will not start until all requirements are met if there is a quick remedy. DISABLED if it is not a quick remedy, and, at the discretion of the Head REFEREE, ROBOT must be re-inspected. RED CARD if a team\'s ROBOT is not compliant with part B or C participates.',
  'G304 - ROBOTS must be set up correctly on the FIELD. A ROBOT must be positioned on the FIELD such that it meets all of the following requirements: \' A. is over a LAUNCH LINE, B. is either touching its own ALLIANCE\'s GOAL or the FIELD perimeter, C. is fully contained on its own ALLIANCE\'s side of the FIELD (FIELD columns A, B, C for blue, or FIELD columns D, E, F for red) (Figure 9-4), D. not attached to, entangled with, or suspended from any FIELD element, E. confined to its STARTING CONFIGURATION (see R101 and R102), and F. in contact with no more than the allowed pre-load possession limit as described in section 10.3.1 SCORING ELEMENTS. Violation: The MATCH will not start until all requirements are met if there is a quick remedy. DISABLED if it is not a quick remedy. G304.C requires the ROBOT to be fully contained within the FIELD perimeter and not overhang the FIELD perimeter wall. Figure 11-1 shows examples of several possible legal ROBOT starting locations. Figure 11-1: Examples of allowed ROBOT starting locations',
  'G305 - Teams must select an OpMode. An OpMode must be selected on the DRIVER STATION app and initialized by pressing the INIT button. If this OpMode is an AUTO OpMode, the 30 second AUTO timer must be enabled. Violation: MATCH will not start until the situation is remedied. DISABLED if ROBOT cannot initialize an OpMode or the situation cannot be remedied quickly. This rule requires all teams to select and INIT an OpMode regardless of whether or not an AUTO OpMode is planned to be used during AUTO. FIELD STAFF will use this as an indication that a team is ready to start the MATCH. Teams without an AUTO OpMode should consider creating a default AUTO OpMode using the BasicOpMode sample and use the auto-loading feature to automatically queue up their TELEOP OpMode. 11.4 In-MATCH Rules in this section pertain to gameplay once a MATCH begins. 11.4.1 AUTO AUTO is the first 30 seconds of the MATCH, during which DRIVERS may not provide input to their ROBOTS, so ROBOTS operate with only their pre-programmed instructions.\'',
  'G401 - Let the ROBOT do its thing. As soon as FIELD STAFF begins the randomization process and until the end of AUTO, DRIVE TEAM members may not directly or indirectly interact with a ROBOT or an OPERATOR CONSOLE, with the following exceptions: A. to press the (▶) start button within a MOMENTARY reaction of the start of the MATCH, B. to press the (■) stop button either at the team\'s discretion or instruction of the Head REFEREE per T202, or C. for personal safety or OPERATOR CONSOLE safety. Violation: MAJOR FOUL plus the ALLIANCE is not eligible for PATTERN points in AUTO if the ROBOT LAUNCHES an ARTIFACT such that it enters the open top of the GOAL after the interaction and before the end of AUTO. FIELD STAFF will not re-randomize the OBELISK due to violations of this rule prior to MATCH start. Teams do not have to start an OpMode if they choose not to run an AUTO OpMode. The intent of G401.A is for teams to start AUTO on time, accounting for the variability in human factors. Strategic violations of G401.A will be considered egregious behavior under G211.',
  'G402 - No AUTO opponent interference. During AUTO, FIELD columns A, B, C constitute the blue side of the FIELD, and columns D, E, F (Figure 9-5) constitute the red side of the FIELD. During AUTO, a ROBOT may not: A. contact an opposing ALLIANCE\'S ROBOT which is completely within the opposing ALLIANCE\'S side of the FIELD either directly or transitively through an ARTIFACT, or B. disrupt an ARTIFACT from its pre-staged location on the opposing ALLIANCE\'S side of the FIELD either directly or transitively through contact with an ARTIFACT, or by LAUNCHING an ARTIFACT directly into it. Violation: MAJOR FOUL per instance of ROBOT contact in G402.A and MAJOR FOUL per ARTIFACT in G402.B. Navigating into the opposing ALLIANCE\'S side of the FIELD during AUTO is a risky gameplay strategy. LAUNCHED ARTIFACTS which happen to enter the other side of the FIELD after being deflected by another object in the FIELD (e.g., FIELD element, ROBOT) will not be penalized. Example 1: A red ROBOT LAUNCHES 1 ARTIFACT onto the opponent side of the FIELD. The LAUNCHED ARTIFACT disrupts 2 pre-staged ARTIFACTS on the blue side of the FIELD. Red is assessed 2 MAJOR FOULS under G402. Example 2: A red ROBOT LAUNCHES 1 ARTIFACT at their GOAL in an attempt to score, but the ARTIFACT misses the open top of the GOAL, deflects off the GOAL structure and rolls into the blue side of the FIELD, disrupting 2 pre-staged ARTIFACTS. No G402 penalties are assessed. 11.4.2 TELEOP',
  'G403 - ROBOTS are motionless between AUTO and TELEOP. Any powered movement of the ROBOT or any of its MECHANISMS is not allowed during the transition period between AUTO and TELEOP. Violation: MAJOR FOUL. Movement that occurs following the conclusion of an AUTO OpMode (due to inertia, gravity, or de-energizing of actuators, etc.) is not a violation of this rule. Teams may press buttons on their DRIVER STATION app to stop the AUTO OpMode, initialize or start a TELEOP OpMode during the AUTO to TELEOP transition period. If the INIT portion of the OpMode causes the ROBOT to violate this rule (actuators moving or twitching in any way) then the team should wait until TELEOP begins before pressing INIT. A ROBOT LAUNCHING an ARTIFACT during the transition period is considered a violation of this rule. Strategic violations of this rule will be considered egregious behavior under G211. Strategic violations include, but are not limited to: - LAUNCHING multiple SCORING ELEMENTS, - operating the GATE, and - moving the ROBOT a substantial distance in a preferred direction.',
  'G404 - ROBOTS are motionless at the end of TELEOP. ROBOTS must no longer have powered movement after the end of TELEOP until the Head REFEREE or their designee signals that teams may retrieve their ROBOTS. \' Violation: MINOR FOUL. MAJOR FOUL per ARTIFACT if ROBOT LAUNCHES an ARTIFACT such that it enters the open top of a GOAL after the end of TELEOP. MAJOR FOUL if ROBOT contacts a GATE after the end of TELEOP. DRIVE TEAMS should make it obvious that the ROBOTS are no longer being controlled by pressing the (■) stop button on the DRIVER STATION app or by discontinuing any operation of the ROBOT by the end of the MATCH period and setting down their controllers. Movement due to inertia, gravity, or de-energizing of actuators, etc. is not considered powered movement. 11.4.3 SCORING ELEMENT',
  'G405 - ROBOTS use SCORING ELEMENTS as directed. A ROBOT may not deliberately use a SCORING ELEMENT in an attempt to ease or amplify a challenge associated with a FIELD element other than as intended. Violation: MAJOR FOUL per SCORING ELEMENT. Examples include, but are not limited to: A. Intentionally positioning SCORING ELEMENTS to impede opponent access to FIELD elements B. Intentionally placing SCORING ELEMENTS into inaccessible locations on the FIELD such as under the RAMP or GOAL C. Intentionally using a SCORING ELEMENT to hold open the GATE',
  'G406 - Keep SCORING ELEMENTS in bounds. A ROBOT may not intentionally eject a SCORING ELEMENT from the FIELD (either directly or by bouncing off a FIELD element or another ROBOT). Violation: MAJOR FOUL per SCORING ELEMENT. SCORING ELEMENTS that leave the FIELD during scoring attempts are not considered intentional ejections.',
  'G407 - Do not damage SCORING ELEMENTS. Neither a ROBOT nor a DRIVE TEAM member may damage a SCORING ELEMENT. Violation: VERBAL WARNING. MAJOR FOUL if REPEATED. DISABLED if the damage is caused by a ROBOT, and the Head REFEREE determines that further damage is likely to occur. Corrective action (such as eliminating sharp edges, removing the damaging MECHANISM, and/or reinspection) may be required before the ROBOT may compete in subsequent MATCHES. SCORING ELEMENTS are expected to undergo a reasonable amount of wear and tear as they are handled by ROBOTS and humans, such as scratching, marking, and eventually damage due to fatigue. Routinely gouging, tearing off pieces, or marking SCORING ELEMENTS are violations of this rule.',
  'G408 - No more than 3 at a time. A ROBOT may not simultaneously CONTROL more than 3 ARTIFACTS. Violation: MINOR FOUL per SCORING ELEMENT over the limit. YELLOW CARD if excessive. Examples of interaction with a SCORING ELEMENT that are not \'CONTROL\' include, but are not limited to: A. \'bulldozing\' (inadvertent contact with a SCORING ELEMENT while in the path of the ROBOT moving about the FIELD) B. \'deflecting\' (being hit by a SCORING ELEMENT that bounces into or off a ROBOT) C. inadvertent contact with a SCORING ELEMENT while attempting to acquire a SCORING ELEMENT from the LOADING ZONE. D. SCORING ELEMENTS that have been LAUNCHED by a ROBOT that are no longer in contact with the ROBOT. It is important to design your ROBOT so that it is impossible to inadvertently or unintentionally CONTROL more than the limit. Excessive violations of CONTROL limits include, but are not limited to: A. simultaneous CONTROL of 5 or more ARTIFACTS, or B. frequent (i.e., 3 or more separate violations in a MATCH), greater-than-MOMENTARY CONTROL of 4 or more ARTIFACTS. REPEATED excessive violations of this rule do not result in additional YELLOW CARDS unless the violation reaches the level of egregious to trigger a G211 violation. 11.4.4 ROBOT',
  'G409 - ROBOTS must be under control. A ROBOT must not pose an undue hazard to a human or an ARENA element during a MATCH in the following ways: A. the ROBOT or anything it CONTROLS, i.e., a SCORING ELEMENT, disrupts anything outside the FIELD or contacts a human that is outside the FIELD. B. the ROBOT operation is dangerous. Violation: DISABLED and VERBAL WARNING. YELLOW CARD if REPEATED or if subsequent violations occur during the event. Please be conscious of REFEREES and FIELD STAFF working around the ARENA who may be in close proximity to your ROBOT. Examples of violations include, but are not limited to: A. Wildly flailing outside the FIELD B. Knocking over a DRIVER STATION stand C. Moving/damaging the FIELD timer display D. Contacting FIELD STAFF or a DRIVE TEAM member outside the FIELD ROBOT contact with ARENA elements outside the FIELD, such as a DRIVER STATION stand, the floor outside the FIELD, or the FIELD wall perimeter outside of the FIELD is not a violation of this rule. Disrupting the OBELISK is not a violation of this rule.',
  'G410 - ROBOTS must stop when instructed. If a team is instructed to DISABLE their ROBOT by a REFEREE per T202, a DRIVE TEAM member must press the (■) stop button on the DRIVER STATION app. Violation: MAJOR FOUL if greater-than-MOMENTARY delay plus RED CARD if CONTINUOUS.',
  'G411 - ROBOTS must be identifiable. A ROBOT\'S team number and ALLIANCE color must not become indeterminate by determination of the Head REFEREE. Violation: VERBAL WARNING. MINOR FOUL if subsequent violations occur during the event. Teams are encouraged to robustly affix their ROBOT SIGNS to their ROBOT in highly visible locations such that they do not easily fall off or become obscured during normal gameplay.',
  'G412 - Don\'t damage the FIELD. A ROBOT may not damage FIELD elements. Violation: VERBAL WARNING. DISABLED if the Head REFEREE infers that additional damage is likely. YELLOW CARD for any subsequent damage during the event. Corrective action (such as eliminating sharp edges, removing the damaging MECHANISM, and/or re-inspection) may be required before the ROBOT will be allowed to compete in subsequent MATCHES. SCORING ELEMENT damage is specifically covered in G407. G407 and G412 do not stack. G412 does not apply to damage caused by normal gameplay actions. FIELD damage includes, but is not limited to: - contaminating the FIELD with a liquid or fine solid as in R205, - damaging TILE in R201, - causing the GATE to bend or break off FIELD damage does not include: - normal GATE interaction resulting in a GATE that \'sticks\' open - normal interaction with the GOAL that causes it to lift off the TILES',
  'G413 - Watch your ARENA interaction. A ROBOT is prohibited from the following interactions with an ARENA element, except for SCORING ELEMENTS (per G407): A. grabbing, B. grasping, C. attaching to, D. becoming entangled with, or E. suspending from. Violation: MAJOR FOUL plus YELLOW CARD if REPEATED or if greater-than-MOMENTARY. DISABLED if the Head REFEREE infers that damage is likely. Corrective action (such as removing the offending MECHANISM, and/or re-inspection) may be required before the ROBOT will be allowed to compete in subsequent MATCHES. ROBOTS operating the GATE should make it clear that they do not violate this rule. ROBOTS are expected to push the GATE lever down to open, but no closing force (e.g., pulling) should be applied.',
  'G414 - ROBOTS have horizontal expansion limits. ROBOTS must comply with the horizontal expansion limits outlined in R105.A during the MATCH. Exceptions: A. If the over-expansion is due to damage and not used for strategic benefit. Violation: MINOR FOUL. MAJOR FOUL if the over-expansion is used for strategic benefit, including if it impedes or enables a scoring action. ROBOTS are allowed to have moving parts that extend outside its STARTING CONFIGURATION, but these extensions must stay within the expansion limit as described in R105.',
  'G415 - ROBOTS have vertical expansion limits, with exceptions. ROBOTS must comply with the vertical expansion limits outlined in R105. ROBOTS may only expand above 18 in. (45.70 cm) up to 38 in. (96.50 cm) if both of the following conditions are true: A. during the final 20 seconds of the MATCH, and B. when not in any LAUNCH ZONES. Violation: MINOR FOUL. MAJOR FOUL if the over-expansion is used for strategic benefit, including if it impedes or enables a scoring action. ROBOTS are allowed to have moving parts that extend outside its STARTING CONFIGURATION, but these extensions must stay within the expansion limit as described in R105.',
  'G416 - LAUNCHING in the LAUNCH ZONE only. ROBOTS may only LAUNCH SCORING ELEMENTS when inside a LAUNCH ZONE or overlapping a LAUNCH LINE. Violation: MINOR FOUL per LAUNCHED SCORING ELEMENT. MAJOR FOUL per LAUNCHED SCORING ELEMENT if the SCORING ELEMENT enters the open top of the GOAL. A SCORING ELEMENT is considered LAUNCHED if it is shot into the air, propelled across the floor to a desired location or in a preferred direction, or thrown in a forceful way. \'Bulldozing\' (inadvertent contact with a SCORING ELEMENT while in the path of the ROBOT moving about the FIELD) is not considered LAUNCHING This is not intended to penalize teams with active manipulators which are expelling SCORING ELEMENTS through normal operation, such as: A. Running an intake in reverse causing a SCORING ELEMENT to travel a short distance from the ROBOT. B. A ROBOT pushing a SCORING ELEMENT a short distance away in the process of herding it across the FIELD.',
  'G417 - ROBOTS only operate GATES as directed. ROBOTS may not: A. contact, either directly or transitively through a SCORING ELEMENT, an opposing ALLIANCE\'S GATE, or B. apply, either directly or transitively through a SCORING ELEMENT, any closing force to either GATE. Violation: MAJOR FOUL and the opposing ALLIANCE is awarded the PATTERN RP if G417.A. Closing force includes any force applied to the GATE in the direction that closes the GATE, even if the GATE is already closed. A ROBOT bumping into a GATE handle which is stuck open to try to get it to close is not considered a closing force.',
  'G418 - ROBOTS may not meddle with ARTIFACTS on RAMPS. ROBOTS may not contact, either directly or transitively through a SCORING ELEMENT CONTROLLED by the ROBOT, ARTIFACTS on a RAMP, including their own RAMP. Additionally, ROBOTS may not: A. remove an ARTIFACT from their own RAMP except by operating the GATE, or B. remove an ARTIFACT from the opponent\'s RAMP by any means. Violation: MAJOR FOUL per ARTIFACT, and the ALLIANCE is ineligible for the PATTERN RP if G418.A, or the opposing ALLIANCE is awarded the PATTERN RP if G418.B. Exceptions are granted for inconsequential and inadvertent contact made by a ROBOT while operating a GATE. Example 1: A red ROBOT that contacts an ARTIFACT on the blue RAMP is in violation of this rule and is assessed 1 MAJOR FOUL under G418. Example 2: A red ROBOT that LAUNCHES an ARTIFACT at an ARTIFACT on the red RAMP, removing it from the RAMP is in violation of this rule. The red ALLIANCE is assessed 1 MAJOR FOUL and is ineligible for the PATTERN RP under G418.A. Example 3: A red ROBOT contacts and opens the blue GATE, causing 5 ARTIFACTS that were on the blue RAMP to leave the RAMP and return to the FIELD. Red is assessed a total of 6 MAJOR FOULS \' 1 under G417.A and 5 under G418.B \' in addition to blue being awarded PATTERN RP under G417.A/G418.B.',
  'G419 - ROBOTS LAUNCH into their own GOAL. ROBOTS may not: A. intentionally place or LAUNCH ARTIFACTS directly onto their own RAMP, or B. place or LAUNCH ARTIFACTS into the opponent\'s GOAL or onto the opponent\'s RAMP. Violation: MAJOR FOUL per ARTIFACT and the opposing ALLIANCE is awarded the PATTERN RP if G419.B. The intent is for ROBOTS to score by LAUNCHING into the open top of their own GOAL. Attempts to intentionally score points with actions that enter the ARTIFACT further down on the RAMP are considered violations of this rule. Attempts to score points for the opponent either through the opponent GOAL or with actions that enter an ARTIFACT further down on the opponent RAMP are also considered violations of this rule. There is no violation for scoring in an opponent\'s DEPOT. 11.4.5 Opponent Interaction Note, G420 and G421 are mutually exclusive. A single ROBOT to ROBOT interaction which violates more than 1 of these rules results in the most punitive penalty, and only the most punitive penalty, being assessed.',
  'G420 - This is not combat robotics. A ROBOT may not deliberately functionally impair an opponent ROBOT. Damage or functional impairment because of contact with a tipped-over or DISABLED opponent ROBOT, which is not perceived by a REFEREE to be deliberate, is not a violation of this rule. Violation: MAJOR FOUL and YELLOW CARD. MAJOR FOUL and RED CARD if opponent ROBOT is unable to drive. FIRST Tech Challenge can be a high-contact competition and may include rigorous gameplay. While this rule aims to limit severe damage to ROBOTS, teams should design their ROBOTS to be robust. Teams are expected to act responsibly. An example of a violation of this rule includes, but is not limited to: A. A ROBOT high-speed rams and/or REPEATEDLY smashes an opponent ROBOT and causes damage. The REFEREE infers that the ROBOT was deliberately trying to damage the opponent\'s ROBOT. Examples of functionally impairing another ROBOT include, but are not limited to: B. disconnecting wires for operation of a component inside the ROBOT CHASSIS. C. disconnecting the opponent ROBOT\'S battery (this example also clearly results in a RED CARD because the ROBOT is no longer able to drive). D. powering off an opponent\'s ROBOT using their reasonably well-protected power switch (This example also clearly results in a RED CARD because the ROBOT is no longer able to drive). Teams should mount their main power switch so it is protected per R609. A team that mounts their ROBOT\'S power switch in an exposed location puts themselves at high risk of incidental contact. Powering off an opponent\'s ROBOT by their exposed power switch during normal interactive gameplay will be considered incidental and not deliberate. At the conclusion of the MATCH, the Head REFEREE may elect to visually inspect a ROBOT to confirm violations of this rule made during a MATCH and remove the violation if the damage cannot be verified. "Unable to drive" means that because of the incident, the DRIVER can no longer drive to a desired location in a reasonable time (generally). For example, if a ROBOT can only move in circles, or can only move extremely slowly, the ROBOT is considered unable to drive.',
  'G421 - Do not tip or entangle. A ROBOT may not deliberately, as perceived by a REFEREE, attach to, tip, or entangle an opponent ROBOT. Violation: MAJOR FOUL and YELLOW CARD. MAJOR FOUL and RED CARD if CONTINUOUS or opponent ROBOT is unable to drive. Examples of violations of this rule include, but are not limited to: A. using a wedge-like MECHANISM to tip over an opponent ROBOT B. making frame-to-frame contact with an opponent ROBOT that is attempting to right itself after previously falling over and causing them to fall over. C. causing an opponent ROBOT to tip over by contacting the ROBOT after it starts to tip if, in the judgement of the REFEREE, that contact could have been avoided. Tipping as an unintended consequence of normal ROBOT-to-ROBOT interaction, including single frame-to-frame hits that result in a ROBOT tipping, as perceived by the REFEREE, is not a violation of this rule. "Unable to drive" means that because of the incident, the DRIVER can no longer drive to a desired location in a reasonable time (generally). For example, if a ROBOT can only move in circles, or can only move extremely slowly, the ROBOT is considered unable to drive.',
  'G422 - There is a 3-count on PINS. A ROBOT may not PIN an opponent\'s ROBOT for more than 3 seconds. A ROBOT is PINNING if it is preventing the movement of an opponent ROBOT by contact, either direct or transitive (such as against a FIELD element) and the opponent ROBOT is attempting to move. A PIN count ends once any of the following criteria below are met: A. the ROBOTS have separated by at least 2 ft. (~61 cm) from each other for more than 3 seconds, B. either ROBOT has moved 2 ft. from where the PIN initiated for more than 3 seconds, or C. the PINNING ROBOT gets PINNED. For criteria A, the PIN count pauses once ROBOTS are separated by 2 ft. until either the PIN ends or the PINNING ROBOT moves back within 2 ft., at which point the PIN count is resumed. For criteria B, the PIN count pauses once either ROBOT has moved 2ft from where the PIN initiated until the PIN ends or until both ROBOTS move back within 2ft., at which point the PIN count is resumed. Violation: MINOR FOUL and an additional MINOR FOUL for every 3 seconds in which the situation is not corrected.',
  'G423 - Do not use strategies intended to shut down major parts of gameplay. \'A ROBOT or ROBOTS may not, in the judgment of a REFEREE, isolate or close off any major element of MATCH play for a greater-than-MOMENTARY duration. Violation: MINOR FOUL and an additional MINOR FOUL for every 3 seconds in which the situation is not corrected. Examples of violations of this rule include, but are not limited to: A. shutting down access to all SCORING ELEMENTS,\' B. quarantining an opponent to a small area of the FIELD, C. quarantining SCORING ELEMENTS out of the opposing ALLIANCE\'S reach, or D. completely blocking access to the opponent\'s GATE.',
  'G424 - GATE ZONE is OFF LIMITS. A ROBOT may not contact, directly or transitively though a SCORING ELEMENT, an opponent ROBOT if either ROBOT is in the opponent\'s GATE ZONE, regardless of who initiates contact. Exceptions: A. A ROBOT in their own ALLIANCE\'S GATE ZONE and in their opponent\'s SECRET TUNNEL ZONE is not protected under G424. Violation: MINOR FOUL. For the exception in G424.A, G425 would apply instead. Figure 11-2 shows some examples of typically protected and non-protected contact in the GATE ZONE. The intent of this rule is to ensure an ALLIANCE has access to their own GATE. Some of the actions shown below may also fall under other penalties including G423 or escalate to G211. Figure 11-2:',
  'G425 - Keep out of opponent\'s SECRET TUNNEL A ROBOT in the opponent\'s SECRET TUNNEL ZONE may not contact, directly or transitively though a SCORING ELEMENT, an opponent ROBOT regardless of who initiates contact. Violation: MINOR FOUL. Figure 11-3 shows some examples of typically protected and non-protected contact in the SECRET TUNNEL ZONE. The intent of this rule is to ensure an ALLIANCE has access to ARTIFACTS exiting from the opponent\'s GATE, but still allow the opponent the opportunity to also access ARTIFACTS if there is no defender present. Figure 11-3:',
  'G426 - LOADING ZONE protection. A ROBOT may not contact, directly or transitively through a SCORING ELEMENT, an opponent ROBOT while either ROBOT is in the opponent\'s LOADING ZONE, regardless of who initiates contact. Violation: MINOR FOUL. Figure 11-4 shows some examples of typically protected and non-protected contact in the LOADING ZONE. The intent of this rule is to ensure an ALLIANCE has access to ARTIFACTS exiting from the opponent\'s GATE but still allows the opponent the opportunity to also access ARTIFACTS if there is no defender present. Some of the actions shown below may also fall under other penalties including G423. Figure 11-4:',
  'G427 - BASE ZONE protection. During the last 20 seconds of the MATCH, a ROBOT may not contact, directly or transitively through a SCORING ELEMENT, an opponent ROBOT while either ROBOT is in the opponent\'s BASE ZONE, regardless of who initiates contact. Violation: MAJOR FOUL and opponent ROBOT and any ROBOT fully supported by the contacted ROBOT are awarded fully returned to BASE points. 11.4.6 Human',
  'G428 - No wandering. DRIVE TEAM members must remain in their designated ALLIANCE AREA. A. DRIVE TEAMS may be anywhere in their respective ALLIANCE AREA during a MATCH. B. DRIVE TEAM members must be staged inside their respective ALLIANCE AREA prior to MATCH start. Violation: VERBAL WARNING. MINOR FOUL if subsequent violations occur during the event. The intent of this rule is to prevent DRIVE TEAM members from leaving their assigned AREA during a MATCH to gain a competitive advantage. For example, moving to another part of the FIELD for better viewing or reaching into the FIELD. Simply breaking the plane of the AREA during normal MATCH play is not a FOUL. DRIVE TEAM members may retrieve SCORING ELEMENTS that have left the FIELD if they are able to do so without violating G428, G430, and G434. Reintroduction of SCORING ELEMENTS must follow rule G432. Exceptions are granted in cases concerning safety and for actions that are inadvertent, MOMENTARY, and inconsequential.',
  'G429 - DRIVE COACHES and other teams: hands off the controls. A ROBOT shall be operated only by the DRIVERS of that team; DRIVE COACHES may not handle the gamepads. DRIVE COACHES, if desired, may help the DRIVERS in the following ways: A. holding the DRIVER STATION device, B. troubleshooting the DRIVER STATION device, C. selecting OpModes on the DRIVER STATION app, D. pressing the INIT button on the DRIVER STATION app, E. pressing the (▶) start button on the DRIVER STATION app, or F. pressing the (■) stop button on the DRIVER STATION app. Violation: MAJOR FOUL. YELLOW CARD if greater-than-MOMENTARY. Exceptions may be made before a MATCH for major conflicts, e.g., religious holidays, major testing, transportation issues.',
  'G430 - DRIVE COACHES, SCORING ELEMENTS are off limits. DRIVE COACHES may not contact SCORING ELEMENTS, unless for safety purposes. Violation: MINOR FOUL.',
  'G431 - DRIVE TEAMS, watch your reach. Once a MATCH starts, a DRIVE TEAM member inside the FIELD may not: A. directly contact a ROBOT, B. contact a SCORING ELEMENT in contact with a ROBOT, C. disrupt SCORING ELEMENT scoring, or D. contact a FIELD element. Violation: MAJOR FOUL plus YELLOW CARD if G431.A. RED CARD and the opposing ALLIANCE is awarded the PATTERN RP if G431.C. Exceptions are granted in cases concerning safety and for actions that are inadvertent, MOMENTARY, and inconsequential. For G431.A and G431.B, the penalty is applied to the DRIVE TEAM member regardless of whether the DRIVE TEAM member or ROBOT initiates contact. Impacting ARTIFACT scoring includes, but is not limited to: A. Contacting an ARTIFACT LAUNCHED by the opponent within the FIELD B. Contacting an ARTIFACT in the opponent\'s GOAL C. Disrupting the scoring of an ARTIFACT on the opponent\'s RAMP or by operating the opponent\'s GATE',
  'G432 - Humans, only meddle with ARTIFACTS in the LOADING ZONE. DRIVE TEAM members may only introduce ARTIFACTS to, remove ARTIFACTS from, or move ARTIFACTS within the LOADING ZONE and only the LOADING ZONE. Actions must occur: A. only during TELEOP, B. without using a tool, C. without causing an ARTIFACT to enter into the LOADING ZONE from elsewhere on the FIELD, and D. without causing an ARTIFACT to leave the LOADING ZONE and enter the rest of the FIELD unless the ARTIFACT is CONTROLLED by a ROBOT as follows: i. ARTIFACT CONTROL begins when the ROBOT is in the LOADING ZONE, and ii. ARTIFACT is still CONTROLLED by the ROBOT when the ROBOT leaves the LOADING ZONE. Violation: MINOR FOUL per ARTIFACT. MAJOR FOUL per ARTIFACT that enters the open top of the GOAL. DRIVE TEAM members may load SCORING ELEMENTS into a ROBOT that is partially or fully in the LOADING ZONE. ARTIFACTS that are unintentionally deflected, e.g., a DRIVE TEAM member protecting themselves from a LAUNCHED ARTIFACT, are an exception to this rule. DECODE is a fast-paced game and teams should practice coordination and communication between the DRIVE TEAM members to avoid unintentional contact between the ROBOT and any humans in violation of G431.A.',
  'G433 - Humans may only enter SCORING ELEMENTS. DRIVE TEAM members may only enter ARTIFACTS onto the FIELD. Violation: MINOR FOUL per non-ARTIFACT item entered onto the FIELD.',
  'G434 - The ALLIANCE AREA has a storage limit. During TELEOP, each ALLIANCE may not store more than 6 ARTIFACTS out of play. DRIVE TEAM members making a good-faith effort to immediately enter additional ARTIFACTS back into play is an exception to this rule. Violation: MINOR FOUL per ARTIFACT over the limit and an additional MINOR FOUL per ARTIFACT over the limit for every 3 seconds in which the situation is not corrected. The intent of this rule is to prevent teams from hoarding ARTIFACTS out of play in the ALLIANCE AREA.'
];

// --- Layout constants ---
// Referee sheets: Row 1=Title, Row 2=Point values (hidden), Row 3=Headers, Row 4+=Data
const REF_DATA_START = 4;
const REF_DATA_END = MAX_TEAMS + 3;
// FinalScores: Row 1=Category headers, Row 2=Point values (hidden), Row 3=Headers, Row 4+=Data
const FS_DATA_START = 4;
const FS_DATA_END = MAX_TEAMS + 3;

// Referee sheet column layout (A-X, 24 columns):
//   A=Team#(auto)  B=Name(auto)  C=Video(auto)
//   D=Notes
//   E=TOTAL(calc)  F=Score w/o Fouls(calc)  G=Auto Score(calc)  H=TeleOp Score(calc)  I=Foul Deduction(calc)
//   J=Minor Fouls  K=Major Fouls  L=G Rules(multiselect)
//   M=MOTIF  N=LEAVE  O=Auto CLS  P=Auto OVF  Q=Auto RAMP Colors  R=Auto PAT(calc)
//   S=Tel CLS  T=Tel OVF  U=Tel DEPOT  V=Tel RAMP Colors  W=Tel PAT(calc)
//   X=BASE
const RC = {
  TEAM: 1, NAME: 2, VIDEO: 3, NOTES: 4,
  TOTAL: 5, SCORE_NO_FOULS: 6, AUTO_SCORE: 7, TEL_SCORE: 8, FOUL_DED: 9,
  MINOR: 10, MAJOR: 11, G_RULES: 12,
  MOTIF: 13, LEAVE: 14, AUTO_CLS: 15, AUTO_OVF: 16, AUTO_RAMP: 17, AUTO_PAT: 18,
  TEL_CLS: 19, TEL_OVF: 20, TEL_DEPOT: 21, TEL_RAMP: 22, TEL_PAT: 23,
  BASE: 24
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
  if (val4.indexOf("notes") !== -1) return "new"; // current: MOTIF at col M, Notes at col D
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

  // --- Save referee scoring data (layout-aware per-field extraction) ---
  let savedRefData = {};
  let noteMap = _buildNoteMap(ss);
  for (let r = 1; r <= NUM_REFEREES; r++) {
    let sheet = findRefSheet(ss, config, r, noteMap);
    if (!sheet) continue;

    let dataStart = _detectRefDataStart(sheet);
    let dataEnd = dataStart + MAX_TEAMS - 1;
    let layoutVer = _detectLayoutVersion(sheet);
    let src = (layoutVer === "old") ? OLD_RC : (layoutVer === "v2") ? V2_RC : RC;
    let numCols = (layoutVer === "old") ? 23 : 24;

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

  // --- Rebuild all referee sheets and restore data ---
  // Note: simple triggers (onEdit) do not fire on programmatic setValue() calls,
  // so no guard is needed during data restoration.
  for (let r = 1; r <= NUM_REFEREES; r++) {
    _buildRefereeSheet(ss, config, r);

    if (savedRefData[r]) {
      let sheet = findRefSheet(ss, config, r);
      if (sheet) {
        _writeColumn(sheet, REF_DATA_START, RC.TEAM, savedRefData[r].teams);
        _writeColumn(sheet, REF_DATA_START, RC.MOTIF, savedRefData[r].motif);
        _writeColumn(sheet, REF_DATA_START, RC.NOTES, savedRefData[r].notes);
        _writeColumn(sheet, REF_DATA_START, RC.LEAVE, savedRefData[r].leave);
        _writeColumn(sheet, REF_DATA_START, RC.AUTO_CLS, savedRefData[r].autoCls);
        _writeColumn(sheet, REF_DATA_START, RC.AUTO_OVF, savedRefData[r].autoOvf);
        _writeColumn(sheet, REF_DATA_START, RC.AUTO_RAMP, savedRefData[r].autoRamp);
        _writeColumn(sheet, REF_DATA_START, RC.TEL_CLS, savedRefData[r].telCls);
        _writeColumn(sheet, REF_DATA_START, RC.TEL_OVF, savedRefData[r].telOvf);
        _writeColumn(sheet, REF_DATA_START, RC.TEL_DEPOT, savedRefData[r].telDepot);
        _writeColumn(sheet, REF_DATA_START, RC.TEL_RAMP, savedRefData[r].telRamp);
        _writeColumn(sheet, REF_DATA_START, RC.BASE, savedRefData[r].base);
        _writeColumn(sheet, REF_DATA_START, RC.MINOR, savedRefData[r].minor);
        _writeColumn(sheet, REF_DATA_START, RC.MAJOR, savedRefData[r].major);
        _writeColumn(sheet, REF_DATA_START, RC.G_RULES, savedRefData[r].gRules);
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
  let configTeamRange = config.getRange("A4:A" + (MAX_TEAMS + 3)).getValues();
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
  _buildFinalScoresSheet(ss);
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
function _buildRefereeSheet(ss, config, refNum) {
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
  let cQ = _colLetter(RC.AUTO_RAMP), cR = _colLetter(RC.AUTO_PAT);
  let cS = _colLetter(RC.TEL_CLS), cT = _colLetter(RC.TEL_OVF), cU = _colLetter(RC.TEL_DEPOT);
  let cV = _colLetter(RC.TEL_RAMP), cW = _colLetter(RC.TEL_PAT), cX = _colLetter(RC.BASE);
  let ds = REF_DATA_START, de = REF_DATA_END;
  let lastCol = cX; // last column letter

  // ---- ROW 1: Title + progress counter (split merge at frozen column boundary) ----
  // A1:C1 (frozen) = progress counter; D1:X1 (scrollable) = title
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
  let pointRow = new Array(24).fill("");
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
  sheet.getRange(2, 1, 1, 24).setValues([pointRow]);
  sheet.getRange("A2:" + lastCol + "2").setFontStyle("italic").setFontSize(10)
    .setHorizontalAlignment("center").setBackground("#E8E8E8").setFontColor("#505050");
  sheet.hideRows(2);

  // ---- ROW 3: Column Headers ----
  let headers = [
    "Team #", "Team Name", "Video",                       // A-C
    "Notes",                                               // D
    "TOTAL\nSCORE", "Score w/o\nFouls", "Auto\nScore",    // E-G
    "TeleOp\nScore", "Foul\nDeduction",                    // H-I
    "Minor\nFouls", "Major\nFouls", "G Rules",             // J-L
    "MOTIF", "LEAVE\n(Yes/No)", "Auto\nCLASSIFIED", "Auto\nOVERFLOW", // M-P
    "Auto RAMP\nColors\n(G/P)", "Auto PATTERN\nCount",     // Q-R
    "TeleOp\nCLASSIFIED", "TeleOp\nOVERFLOW", "TeleOp\nDEPOT", // S-U
    "TeleOp RAMP\nColors\n(G/P)", "TeleOp PATTERN\nCount", // V-W
    "BASE\n(None/Partial/Full)"                             // X
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
  sheet.getRange(cN + "3:" + cR + "3").setBackground("#548235"); // Auto
  sheet.getRange(cS + "3:" + lastCol + "3").setBackground("#C55A11"); // TeleOp

  // ---- DATA ROWS (batch formula writes) ----
  // Gate: team# only (no MOTIF gate)
  let formulasB = [], formulasC = [], formulasFJ = [], formulasR = [], formulasW = [];
  for (let row = ds; row <= de; row++) {
    let gate = '$' + cA + row + '=""';

    formulasB.push(['=IF(' + cA + row + '="","",IFERROR(VLOOKUP(' + cA + row + ',Config!$A:$' + _colLetter(3) + ',2,FALSE),""))']);
    formulasC.push(['=IF(' + cA + row + '="","",IFERROR(VLOOKUP(' + cA + row + ',Config!$A:$' + _colLetter(3) + ',3,FALSE),""))']);

    // F: TOTAL = max(0, score - fouls)
    let fF = '=IF(' + gate + ',"",MAX(0,' + cG + row + '-' + cJ + row + '))';
    // G: SCORE_NO_FOULS = Auto + TeleOp
    let fG = '=IF(' + gate + ',"",'+cH + row + '+' + cI + row + ')';
    // H: AUTO_SCORE = LEAVE + CLS*pts + OVF*pts + PAT*pts
    let fH = '=IF(' + gate + ',"",IF(' + cN + row + '="Yes",' + PTS_LEAVE + ',0)+' +
      cO + row + '*' + PTS_CLASSIFIED + '+' + cP + row + '*' + PTS_OVERFLOW + '+' + cR + row + '*' + PTS_PATTERN + ')';
    // I: TEL_SCORE = CLS*pts + OVF*pts + DEPOT*pts + PAT*pts + BASE
    let fI = '=IF(' + gate + ',"",'+cS + row + '*' + PTS_CLASSIFIED + '+' + cT + row + '*' + PTS_OVERFLOW + '+' +
      cU + row + '*' + PTS_DEPOT + '+' + cW + row + '*' + PTS_PATTERN + '+' +
      'IF(' + cX + row + '="Full",' + PTS_BASE_FULL + ',IF(' + cX + row + '="Partial",' + PTS_BASE_PARTIAL + ',0)))';
    // J: FOUL_DED = Minor*pts + Major*pts
    let fJ = '=IF(' + gate + ',"",'+cK + row + '*' + PTS_MINOR_FOUL + '+' + cL + row + '*' + PTS_MAJOR_FOUL + ')';
    formulasFJ.push([fF, fG, fH, fI, fJ]);

    // R: AUTO_PAT — PATTERN count (0 when MOTIF blank or "Not Shown")
    // MIN(LEN,RAMP_MAX_CHARS) caps comparison length defensively
    let fR = '=IF(' + gate + ',"",IF(OR(' + cD + row + '="",' + cD + row + '="Not Shown"),0,' +
      'IF(LEN(' + cQ + row + ')=0,0,SUMPRODUCT((MID(UPPER(' + cQ + row + '),SEQUENCE(MIN(LEN(' + cQ + row + '),' + RAMP_MAX_CHARS + ')),1)=' +
      'MID(REPT(' + cD + row + ',3),SEQUENCE(MIN(LEN(' + cQ + row + '),' + RAMP_MAX_CHARS + ')),1))*1))))';
    formulasR.push([fR]);

    // W: TEL_PAT — same logic with TeleOp RAMP colors
    let fW = '=IF(' + gate + ',"",IF(OR(' + cD + row + '="",' + cD + row + '="Not Shown"),0,' +
      'IF(LEN(' + cV + row + ')=0,0,SUMPRODUCT((MID(UPPER(' + cV + row + '),SEQUENCE(MIN(LEN(' + cV + row + '),' + RAMP_MAX_CHARS + ')),1)=' +
      'MID(REPT(' + cD + row + ',3),SEQUENCE(MIN(LEN(' + cV + row + '),' + RAMP_MAX_CHARS + ')),1))*1))))';
    formulasW.push([fW]);
  }
  sheet.getRange(ds, RC.NAME, MAX_TEAMS, 1).setFormulas(formulasB);
  sheet.getRange(ds, RC.VIDEO, MAX_TEAMS, 1).setFormulas(formulasC);
  sheet.getRange(ds, RC.TOTAL, MAX_TEAMS, 5).setFormulas(formulasFJ);
  sheet.getRange(ds, RC.AUTO_PAT, MAX_TEAMS, 1).setFormulas(formulasR);
  sheet.getRange(ds, RC.TEL_PAT, MAX_TEAMS, 1).setFormulas(formulasW);

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

  // ---- FORMATTING ----
  sheet.getRange(cA + ds + ":" + cA + de).setBackground("#F2F2F2");
  sheet.getRange("B" + ds + ":C" + de).setBackground("#E8F0FE");
  sheet.getRange(cD + ds + ":" + cD + de).setBackground("#E2D9F3");
  sheet.getRange(cE + ds + ":" + cE + de).setBackground("#FFF2CC");
  sheet.getRange(cF + ds + ":" + cJ + de).setBackground("#D6E4F0");
  sheet.getRange(cK + ds + ":" + cL + de).setBackground("#FFF2CC");
  sheet.getRange(cM + ds + ":" + cM + de).setBackground("#FFF2CC");
  sheet.getRange(cN + ds + ":" + cP + de).setBackground("#E2EFDA");
  sheet.getRange(cQ + ds + ":" + cQ + de).setBackground("#E2EFDA");
  sheet.getRange(cS + ds + ":" + cU + de).setBackground("#FDF2E9");
  sheet.getRange(cV + ds + ":" + cV + de).setBackground("#FDF2E9");
  sheet.getRange(cX + ds + ":" + cX + de).setBackground("#FFF2CC");

  sheet.getRange(cF + ds + ":" + cF + de).setFontWeight("bold").setFontSize(11);
  sheet.getRange(cQ + ds + ":" + cQ + de).setFontFamily("Courier New").setFontWeight("bold");
  sheet.getRange(cV + ds + ":" + cV + de).setFontFamily("Courier New").setFontWeight("bold");

  sheet.getRange(cA + ds + ":" + lastCol + de).setHorizontalAlignment("center");
  sheet.getRange("B" + ds + ":B" + de).setHorizontalAlignment("left");
  sheet.getRange("C" + ds + ":C" + de).setHorizontalAlignment("left");
  sheet.getRange(cE + ds + ":" + cE + de).setHorizontalAlignment("left");
  sheet.getRange(cM + ds + ":" + cM + de).setHorizontalAlignment("left").setWrap(true);

  sheet.getRange("A3:" + lastCol + de).setBorder(true, true, true, true, true, true,
    "#B4B4B4", SpreadsheetApp.BorderStyle.SOLID);

  // Column widths: A=Team#, B=Name, C=Video, D=Notes, E=TOTAL, F=ScoreNoFouls,
  // G=AutoScore, H=TelScore, I=FoulDed, J=Minor, K=Major, L=GRules, M=MOTIF,
  // N=LEAVE, O=AutoCLS, P=AutoOVF, Q=AutoRAMP, R=AutoPAT, S=TelCLS, T=TelOVF,
  // U=TelDEPOT, V=TelRAMP, W=TelPAT, X=BASE
  let colWidths = [85,150,250,200,90,85,80,85,80,75,75,80,80,75,90,90,120,85,100,90,80,120,85,110];
  for (let c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(3);

  // Hide calculated PATTERN columns (formulas still compute, just not visible)
  sheet.hideColumns(RC.AUTO_PAT);
  sheet.hideColumns(RC.TEL_PAT);

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
  if (range.getColumn() !== RC.G_RULES) return;
  if (range.getRow() < REF_DATA_START || range.getRow() > REF_DATA_END) return;
  // Only process referee sheets (identified by ref_index note on A1)
  let note = "";
  try { note = sheet.getRange("A1").getNote() || ""; } catch(ex) { return; }
  if (note.indexOf("ref_index:") !== 0) return;

  let newValue = e.value;
  if (!newValue) return;

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

  range.setValue(codes.join(", "));
}

// ============================================================
// FINAL SCORES SHEET (internal — called by buildAll)
// ============================================================
function _buildFinalScoresSheet(ss) {
  let oldSheet = ss.getSheetByName("FinalScores");
  let sheet = ss.insertSheet("FinalScores" + (oldSheet ? "_new" : ""));
  if (oldSheet) ss.deleteSheet(oldSheet);
  sheet.setName("FinalScores");

  // FinalScores column constants
  let FS = {
    TEAM:1, NAME:2, VIDEO:3, SCORED_BY:4, OFFICIAL_REF:5, AGREE:6, NOTES:7,
    FINAL_SCORE:8, SCORE_NO_FOULS:9, AUTO_SCORE:10, TEL_SCORE:11, FOUL_DED:12,
    MINOR:13, MAJOR:14, G_RULES:15,
    LEAVE:16, AUTO_CLS:17, AUTO_OVF:18, AUTO_RAMP:19, AUTO_PAT:20,
    TEL_CLS:21, TEL_OVF:22, TEL_DEPOT:23, TEL_RAMP:24, TEL_PAT:25, BASE:26,
    EFF_REF: 27
  };
  let fsLastVisCol = _colLetter(26); // Z
  let fsEffRef = _colLetter(FS.EFF_REF); // AA

  // Pre-compute INDIRECT referee range strings (one per referee, reused across all rows)
  let indRefStrs = [];
  for (let r = 1; r <= NUM_REFEREES; r++) {
    indRefStrs[r] = 'INDIRECT("\'"&Config!' + _refConfigCol(r) + '$2&"\'!$A:$' + _colLetter(24) + '")';
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
    {range: _colLetter(FS.LEAVE)+"1:"+_colLetter(FS.AUTO_PAT)+"1", label: "Autonomous Period", bg: "#548235"},
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

  let pvRow = new Array(26 - FS.FINAL_SCORE + 1).fill("");
  pvRow[FS.MINOR - FS.FINAL_SCORE] = "\u00d7(\u2212" + PTS_MINOR_FOUL + ")";
  pvRow[FS.MAJOR - FS.FINAL_SCORE] = "\u00d7(\u2212" + PTS_MAJOR_FOUL + ")";
  pvRow[FS.LEAVE - FS.FINAL_SCORE] = String(PTS_LEAVE);
  pvRow[FS.AUTO_CLS - FS.FINAL_SCORE] = "\u00d7" + PTS_CLASSIFIED;
  pvRow[FS.AUTO_OVF - FS.FINAL_SCORE] = "\u00d7" + PTS_OVERFLOW;
  pvRow[FS.AUTO_PAT - FS.FINAL_SCORE] = "\u00d7" + PTS_PATTERN + " ea";
  pvRow[FS.TEL_CLS - FS.FINAL_SCORE] = "\u00d7" + PTS_CLASSIFIED;
  pvRow[FS.TEL_OVF - FS.FINAL_SCORE] = "\u00d7" + PTS_OVERFLOW;
  pvRow[FS.TEL_DEPOT - FS.FINAL_SCORE] = "\u00d7" + PTS_DEPOT;
  pvRow[FS.TEL_PAT - FS.FINAL_SCORE] = "\u00d7" + PTS_PATTERN + " ea";
  pvRow[FS.BASE - FS.FINAL_SCORE] = PTS_BASE_PARTIAL + "/" + PTS_BASE_FULL;
  sheet.getRange(2, FS.FINAL_SCORE, 1, pvRow.length).setValues([pvRow]);
  sheet.getRange(_colLetter(FS.FINAL_SCORE) + "2:" + fsLastVisCol + "2")
    .setFontWeight("bold").setHorizontalAlignment("center")
    .setFontSize(10).setFontColor("#505050").setBackground("#E8E8E8");
  sheet.hideRows(2);

  // ---- ROW 3: Column headers ----
  let headers = [
    "Number", "Name", "Video",
    "Scored By", "Official\nReferee", "Refs\nAgree?", "Notes",
    "Final\nScore", "Score w/o\nFouls", "Auto\nScore", "TeleOp\nScore", "Foul\nDeduction",
    "Minor", "Major", "G Rules",
    "LEAVE", "Auto\nCLASSIFIED", "Auto\nOVERFLOW", "Auto RAMP\nColors", "Auto\nPATTERN",
    "Tel\nCLASSIFIED", "Tel\nOVERFLOW", "Tel\nDEPOT", "Tel RAMP\nColors", "Tel\nPATTERN", "BASE"
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
  sheet.getRange(_colLetter(FS.LEAVE) + "3:" + _colLetter(FS.AUTO_PAT) + "3").setBackground("#548235");
  sheet.getRange(_colLetter(FS.TEL_CLS) + "3:" + _colLetter(FS.BASE) + "3").setBackground("#C55A11");

  // Hidden helper column AA header
  sheet.getRange(3, FS.EFF_REF).setValue("effectiveRef");

  // ---- DATA ROWS ----
  // Mapping: FS column -> RC column for VLOOKUP
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
    [FS.AUTO_PAT,      RC.AUTO_PAT],
    [FS.TEL_CLS,       RC.TEL_CLS],
    [FS.TEL_OVF,       RC.TEL_OVF],
    [FS.TEL_DEPOT,     RC.TEL_DEPOT],
    [FS.TEL_RAMP,      RC.TEL_RAMP],
    [FS.TEL_PAT,       RC.TEL_PAT],
    [FS.BASE,          RC.BASE]
  ];

  // Agreement check: all input columns except G_RULES
  let elemCols = [RC.MOTIF, RC.LEAVE, RC.AUTO_CLS, RC.AUTO_OVF, RC.AUTO_RAMP,
                  RC.TEL_CLS, RC.TEL_OVF, RC.TEL_DEPOT, RC.TEL_RAMP, RC.BASE,
                  RC.MINOR, RC.MAJOR];
  let numericCols = [RC.AUTO_CLS, RC.AUTO_OVF, RC.TEL_CLS, RC.TEL_OVF, RC.TEL_DEPOT,
                     RC.MINOR, RC.MAJOR];

  let ds = FS_DATA_START, de = FS_DATA_END;

  // Build all formulas as arrays
  let formulasAD = [];  // A-D
  let formulasF = [];   // F (Agree?)
  let formulasG = [];   // G (Notes)
  let formulasHZ = [];  // H-Z (scores from effectiveRef)
  let formulasAA = [];  // AA (effectiveRef helper)

  for (let row = ds; row <= de; row++) {
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

    // refCount expression (used in F and AA)
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
      'IF(IFERROR(ROWS(UNIQUE(FILTER({' + concatJoined + '},{' + concatJoined + '}<>"")))=1,TRUE),"Yes","No")),"N/A"))'
    ]);

    // AA: effectiveRef — Official Ref if set, else auto-select if exactly 1 ref scored
    let singleRefParts = [];
    for (let r = 1; r <= NUM_REFEREES; r++) {
      singleRefParts.push('IF(' + hasScored(r, row) + ',Config!' + _refConfigCol(r) + '$2,"")');
    }
    // Validate Official Referee against Config names (defense-in-depth: prevents INDIRECT from
    // referencing arbitrary sheets even if data validation is bypassed via API)
    let offRefCol = '$' + _colLetter(FS.OFFICIAL_REF) + row;
    formulasAA.push([
      '=IF($A' + row + '="","",IF(AND(' + offRefCol + '<>"",ISNUMBER(MATCH(' + offRefCol + ',Config!$D$2:$I$2,0))),' +
      offRefCol + ',' +
      'IF(' + refCountExpr + '=1,TEXTJOIN("",TRUE,' + singleRefParts.join(',') + '),"")))'
    ]);

    // G: Notes — two-mode (effectiveRef set: plain; not set: all refs with "Name: text")
    let notesEffRef = 'IFERROR(VLOOKUP($A' + row + ',INDIRECT("\'"&$' + fsEffRef + row + '&"\'!$A:$' + _colLetter(24) + '"),' + RC.NOTES + ',FALSE),"")';
    let notesAllParts = [];
    for (let r = 1; r <= NUM_REFEREES; r++) {
      let noteVal = 'IFERROR(VLOOKUP($A' + row + ',' + indRef(r) + ',' + RC.NOTES + ',FALSE),"")';
      notesAllParts.push('IF(AND(' + hasScored(r, row) + ',' + noteVal + '<>""),Config!' + _refConfigCol(r) + '$2&": "&' + noteVal + ',"")');
    }
    formulasG.push([
      '=IF($A' + row + '="","",IF($' + fsEffRef + row + '<>"",' + notesEffRef + ',' +
      'IFERROR(TEXTJOIN(CHAR(10),TRUE,' + notesAllParts.join(',') + '),"")))'
    ]);

    // H-Z: Score columns — per-field agreement with effectiveRef override
    // When effectiveRef is set (official ref selected or single ref), show that ref's value.
    // When multiple refs scored and no official ref, show value only if all refs agree;
    // otherwise CHAR(8203) (zero-width space) marks disagreement for CF highlighting.
    let overrideRef = 'INDIRECT("\'"&$' + fsEffRef + row + '&"\'!$A:$' + _colLetter(24) + '")';
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

      let agreeCheck = 'LET(v,FILTER(' + filterVals + ',' + filterCriteria + '),' +
        'IF(ROWS(UNIQUE(v))=1,INDEX(v,1),CHAR(8203)))';

      rowFormulas.push(
        '=IF($A' + row + '="","",IF($' + fsEffRef + row + '<>"",' + effRefVal + ',' +
        'IF(' + refCountExpr + '<2,"",' + agreeCheck + ')))'
      );
    }
    formulasHZ.push(rowFormulas);
  }

  // Batch write all formulas
  sheet.getRange(ds, 1, MAX_TEAMS, 4).setFormulas(formulasAD);
  sheet.getRange(ds, FS.AGREE, MAX_TEAMS, 1).setFormulas(formulasF);
  sheet.getRange(ds, FS.NOTES, MAX_TEAMS, 1).setFormulas(formulasG);
  sheet.getRange(ds, FS.FINAL_SCORE, MAX_TEAMS, vlookupMap.length).setFormulas(formulasHZ);
  sheet.getRange(ds, FS.EFF_REF, MAX_TEAMS, 1).setFormulas(formulasAA);

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
  sheet.getRange(cFS(FS.LEAVE) + ds + ":" + cFS(FS.AUTO_PAT) + de).setBackground("#E2EFDA");
  sheet.getRange(cFS(FS.TEL_CLS) + ds + ":" + cFS(FS.BASE) + de).setBackground("#FDF2E9");

  sheet.getRange("A" + ds + ":" + fsLastVisCol + de).setHorizontalAlignment("center");
  sheet.getRange("B" + ds + ":B" + de).setHorizontalAlignment("left");
  sheet.getRange("C" + ds + ":C" + de).setHorizontalAlignment("left");
  sheet.getRange(cFS(FS.SCORED_BY) + ds + ":" + cFS(FS.SCORED_BY) + de).setHorizontalAlignment("left").setWrap(true);
  sheet.getRange(cFS(FS.NOTES) + ds + ":" + cFS(FS.NOTES) + de).setHorizontalAlignment("left");
  sheet.getRange(cFS(FS.G_RULES) + ds + ":" + cFS(FS.G_RULES) + de).setHorizontalAlignment("left");

  sheet.getRange("A3:" + fsLastVisCol + de).setBorder(true, true, true, true, true, true,
    "#B4B4B4", SpreadsheetApp.BorderStyle.SOLID);

  // Column widths: A-Z (26 cols)
  let colWidths = [
    85, 150, 250,                      // A-C: Team#, Name, Video
    150, 115, 85, 200,                 // D-G: ScoredBy, OfficialRef, Agree, Notes
    85, 85, 75, 80, 80,               // H-L: Scores
    65, 65, 80,                        // M-O: Minor, Major, GRules
    65, 85, 85, 85, 80,               // P-T: LEAVE, AutoCLS, AutoOVF, AutoRAMP, AutoPAT
    85, 85, 65, 85, 80, 65            // U-Z: TelCLS, TelOVF, TelDEPOT, TelRAMP, TelPAT, BASE
  ];
  for (let c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // Hide helper column AA and calculated PATTERN columns
  sheet.hideColumns(FS.EFF_REF);
  sheet.hideColumns(FS.AUTO_PAT);
  sheet.hideColumns(FS.TEL_PAT);

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

  // 2. Per-field disagreement — red background on scoring cells containing CHAR(8203) marker
  let scoreDataRange = [sheet.getRange(cFS(FS.FINAL_SCORE) + ds + ":" + fsLastVisCol + de)];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=EXACT(' + cFS(FS.FINAL_SCORE) + ds + ',CHAR(8203))')
    .setBackground("#FF9999")
    .setRanges(scoreDataRange)
    .build());

  // 3. Yellow row disagreement (Refs Agree? = "No") — lower priority than per-field red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + ds + '<>"",$' + agreeCol + ds + '="No")')
    .setBackground("#FFFF00")
    .setRanges([sheet.getRange("A" + ds + ":" + fsLastVisCol + de)])
    .build());

  // 4. Missing Official Ref orange
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + ds + '<>"",$' + _colLetter(FS.OFFICIAL_REF) + ds + '="")')
    .setBackground("#FDE9D9")
    .setRanges([sheet.getRange(_colLetter(FS.OFFICIAL_REF) + ds + ":" + _colLetter(FS.OFFICIAL_REF) + de)])
    .build());

  // 5. Zebra striping
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
    let numCols = (layoutVer === "old") ? 23 : 24;

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
  let cD = _colLetter(RC.MOTIF), cE = _colLetter(RC.NOTES);
  let cK = _colLetter(RC.MINOR), cM = _colLetter(RC.G_RULES);
  let cN = _colLetter(RC.LEAVE), cQ = _colLetter(RC.AUTO_RAMP);
  let cS = _colLetter(RC.TEL_CLS), cV = _colLetter(RC.TEL_RAMP);
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

    // Non-contiguous input ranges (5 blocks)
    let inputRange1 = sheet.getRange(cE + REF_DATA_START + ":" + cE + REF_DATA_END);  // Notes
    let inputRange2 = sheet.getRange(cK + REF_DATA_START + ":" + cM + REF_DATA_END);  // Minor, Major, G Rules
    let inputRange3 = sheet.getRange(cD + REF_DATA_START + ":" + cQ + REF_DATA_END);  // MOTIF, LEAVE, Auto CLS/OVF/RAMP
    let inputRange4 = sheet.getRange(cS + REF_DATA_START + ":" + cV + REF_DATA_END);  // Tel CLS/OVF/DEPOT/RAMP
    let inputRange5 = sheet.getRange(cX + REF_DATA_START + ":" + cX + REF_DATA_END);  // BASE
    sheetProt.setUnprotectedRanges([inputRange1, inputRange2, inputRange3, inputRange4, inputRange5]);

    if (hasEmails && refEmails[r - 1] !== "" && refEmails[r - 1].indexOf("@") !== -1) {
      let refEmail = refEmails[r - 1];

      let inputRanges = [inputRange1, inputRange2, inputRange3, inputRange4, inputRange5];
      let rangeNames = ["Scoring 1", "Scoring 2", "Scoring 3", "Scoring 4", "Scoring 5"];
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

    // Explicit range protection on hidden effectiveRef helper column (belt-and-suspenders;
    // already locked by sheet-level protection, but this makes the intent explicit and
    // prevents the INDIRECT trust chain from being tampered with via the unprotected ranges list)
    let effRefRange = finalSheet.getRange("AA" + FS_DATA_START + ":AA" + FS_DATA_END);
    let effRefProt = effRefRange.protect().setDescription("FinalScores - effectiveRef Helper");
    effRefProt.addEditor(me);
    _restrictEditors(effRefProt, [meEmail]);
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
