/*
 * SPDX-FileCopyrightText: 2025 Mahmut Berkan Çetin <m.berkancetin@gmail.com> & Selim Gündüz <sgunduz@atu.edu.tr>
 *
 * SPDX-License-Identifier: MIT
 */


/*
Creators: Mahmut Berkan Çetin and Selim Gündüz
Title: Conducting Multilingual Customizable Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script.
Description: This script automates the mCSEB method in Google Sheets.
How to Cite:

Çetin, Mahmut Berkan & Gündüz, Selim (2025). Protocol for Conducting Multilingual Customizable Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script. protocols.io, https://dx.doi.org/10.17504/protocols.io.x54v9r3xzv3e/v1

*/


const GLOBAL_VARIABLES = {
  SESSION_FOCUS: {{SESSION_FOCUS_JS_STRING}},
  LANDING_SHEET: {{LANDING_SHEET_JS_STRING}},
  PARTICIPANT: {{PARTICIPANT_LIST_AS_JS_ARRAY}},
  PARTICIPANT_SHEET_LANGUAGE: {{PARTICIPANT_SHEET_LANGUAGE_LIST_AS_JS_ARRAY}},
  MODERATOR_SHEET: {{MODERATOR_SHEET_JS_STRING}},
  SESSION_LANGUAGE: {{SESSION_LANGUAGE_JS_STRING}},
  IDEA_SWAP_ALGORITHM: {{IDEA_SWAP_ALGORITHM_JS_STRING}},
  TIME_LEFT: {{TIME_LEFT_JS_STRING}},
  MINS_LEFT: {{MINS_LEFT_JS_STRING}},
  ONE_MIN_LEFT: {{ONE_MIN_LEFT_JS_STRING}},
  TIME_IS_UP: {{TIME_IS_UP_JS_STRING}},
  MINUTES: {{MINUTES_INT}},
  FOCUS: {{FOCUS_LIST_AS_JS_ARRAY}},
  ROUNDS: {{ROUNDS_LIST_AS_JS_ARRAY}},
  IDEAS: {{IDEAS_LIST_AS_JS_ARRAY}},
  CHECK_IDEAS: {{CHECK_IDEAS_JS_STRING}},
  ROUND_CHANGE: {{ROUND_CHANGE_LIST_AS_JS_ARRAY}},
  SESSION_COMPLETE: {{SESSION_COMPLETE_JS_STRING}},
  STARTING: {{STARTING_JS_STRING}},
  STOPPED: {{STOPPED_JS_STRING}}
};

const harmonicShift = {{ALGORITHM_SHIFT_JS_STRING}}
const roundCount = GLOBAL_VARIABLES.ROUNDS.length;
const ideasCount = GLOBAL_VARIABLES.IDEAS.length;
const participantCount = GLOBAL_VARIABLES.PARTICIPANT.length;
const totalIdeas = participantCount * ideasCount * roundCount;
const totalTime = roundCount * GLOBAL_VARIABLES.MINUTES;
const languageColumnLetter = String.fromCharCode(67 + roundCount); // 'C' is 67 in ASCII

const MODERATOR_VARIABLES = {
  MENU: {{MENU_OBJECT_AS_JS_OBJECT}},
  SESSION_START: {{MOD_SESSION_START_TEMPLATE_STRING_PART1}} + ideasCount + {{MOD_SESSION_START_TEMPLATE_STRING_PART2}},
  IDEAS_TABLE : GLOBAL_VARIABLES.PARTICIPANT.flatMap(x => GLOBAL_VARIABLES.IDEAS.map(y => x + "-" + y)),
  ROUND_END_PHRASE: {{ROUND_END_PHRASE_JS_STRING}},
  CURRENT_ROUND: {{CURRENT_ROUND_JS_STRING}},
  DATA_PREP: {
    SheetName: {{DATA_PREP_SheetName_JS_STRING}},
    SessionLanguage: {{DATA_PREP_SessionLanguage_JS_STRING}},
    IdeaRawColumn: {{DATA_PREP_IdeaRawColumn_JS_STRING}},
    TranslatedLanguage: {{DATA_PREP_TranslatedLanguage_JS_STRING}},
    IdeaID: {{DATA_PREP_IdeaID_JS_STRING}},
    IdeaTimestamp: {{DATA_PREP_IdeaTimestamp_JS_STRING}},
    IdeaRoundStartTimestamp: {{DATA_PREP_IdeaRoundStartTimestamp_JS_STRING}},
    TranslateColumn: {{DATA_PREP_TranslateColumn_JS_STRING}},
    ManualCategorization: {{DATA_PREP_ManualCategorization_JS_STRING}}
  },
  COLORS: {{COLORS_OBJECT_AS_JS_OBJECT}},
  START_TIMER: {{START_TIMER_OBJECT_AS_JS_OBJECT}},
};

const imageUrl = {{IMAGE_URL_JS_STRING}};
const colabGitHubUrl = {{COLAB_GITHUB_URL_JS_STRING}};
const ideas_prefix = {{LP_IDEAS_PREFIX_JS_STRING}}.replace("{ideasCount}", ideasCount);
const time_prefix = {{LP_TIME_PREFIX_JS_STRING}}.replace("{GLOBAL_VARIABLES.MINUTES}", GLOBAL_VARIABLES.MINUTES);
const totals_prefix = {{LP_TOTALS_TEXT_INFIX_JS_STRING}}.replace("{totalIdeas}", totalIdeas).replace("{totalTime}", totalTime);

const landing_page_alert_header = {{LP_ALERT_HEADER_JS_STRING}};
const landing_page_alert = {{LP_ALERT_TEXT_JS_STRING}}.replace("{Welcome}", GLOBAL_VARIABLES.LANDING_SHEET);
const reset_session_flags_header = {{RESET_SESSION_FLAGS_HEADER_JS_STRING}};
const reset_session_alert = {{RESET_SESSION_ALERT_TEXT_JS_STRING}}.replace("{CreateLandingPage}", MODERATOR_VARIABLES.MENU.LandingPage).replace("{CreateSession}", MODERATOR_VARIABLES.MENU.SessionElements);
const create_session_alert_header = {{CREATE_SESSION_ALERT_HEADER_JS_STRING}};
const create_session_alert_text = {{CREATE_SESSION_ALERT_TEXT_JS_STRING}};

const LANDINGPAGE_VARIABLES = {
  GREETING_MESSAGE: {{LP_GREETING_MESSAGE_JS_STRING}},
  SESSION_TITLE: participantCount + "-" + roundCount + "-" + ideasCount + "-" + GLOBAL_VARIABLES.MINUTES + " " + {{LP_SESSION_TITLE_SUFFIX_JS_STRING}},
  PARTICIPANTS: {{LP_PARTICIPANTS_PREFIX_JS_STRING}} + participantCount,
  ROUNDS: {{LP_ROUNDS_PREFIX_JS_STRING}} + roundCount,
  IDEAS: ideas_prefix,
  TIME: time_prefix,
  TOTALS: totals_prefix
};

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const sheetId = spreadsheet.getId();

const ANALYSIS_VARIABLES = {
  POPUP_TITLE : {{ANALYSIS_POPUP_TITLE_JS_STRING}},
  COLAB_POPUP: `
    <p><b>{{ANALYSIS_COLAB_POPUP_NEXT_STEPS_JS_STRING}}</b></p>
    <ol>
      <li>{{ANALYSIS_COLAB_POPUP_STEP1_PART1_JS_STRING}}
          <br><a href="${colabGitHubUrl}" target="_blank">{{ANALYSIS_COLAB_POPUP_STEP1_LINK_TEXT_JS_STRING}}</a>
          <br><small>({{ANALYSIS_COLAB_POPUP_STEP1_SMALL_TEXT_JS_STRING}})</small>
      </li>
      <li>
        {{ANALYSIS_COLAB_POPUP_STEP2_JS_STRING}}
      </li>
      <li>
        {{ANALYSIS_COLAB_POPUP_STEP3_JS_STRING}}
      </li>
        <code><br><b>spreadsheet_id = "${sheetId}"</b>
        <br><b>sheet_name = "${MODERATOR_VARIABLES.DATA_PREP.SheetName}"</b>
        <br><b>ideaID = "${MODERATOR_VARIABLES.DATA_PREP.IdeaID}"</b>
        <br><b>idea_timestamp = "${MODERATOR_VARIABLES.DATA_PREP.IdeaTimestamp}"</b>
        <br><b>round_timestamp = "${MODERATOR_VARIABLES.DATA_PREP.IdeaRoundStartTimestamp}"</b>
        <br><b>original_column = "${MODERATOR_VARIABLES.DATA_PREP.IdeaRawColumn}"</b>
        <br><b>translate_column = "${MODERATOR_VARIABLES.DATA_PREP.TranslateColumn}"</b>
        <br><b>category_manual = "${MODERATOR_VARIABLES.DATA_PREP.ManualCategorization}"</b>
        </code>
    </ol>
  `
};

const TIMER_STR = String(GLOBAL_VARIABLES.MINUTES).padStart(2, '0') + ':00';


/*
The code begins here.
Feel free to modify, but please review carefully to maintain functionality.
*/



function getFormulaSeparatorFromSheet() {
  // Reads if seperator is comma: min(1,0) returns 1
  // or if seperator is semicolon: min(1;0) returns 0
  // in the sheet to detect the separator
  const ss = spreadsheet.getActiveSheet();
  const range = ss.getRange((600), 25);

  range.setFormula('=MIN(1,0)');
  SpreadsheetApp.flush();

  const formula = range.getFormula();

  range.clearContent();

  // Detect which separator is used
  return formula.includes(1) ? ';' : ',';
}

const sep = getFormulaSeparatorFromSheet();


/*
----- FUNCTION TO CREATE LANDING PAGE -----
*/
function createBrainwritingLandingPage() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getDocumentProperties();

  if (properties.getProperty('landingPageCreated') === 'true') {
    const response = ui.alert(
      landing_page_alert_header,
      landing_page_alert,
      ui.ButtonSet.YES_NO);

    if (response !== ui.Button.YES) {
      return;
    }
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let landingSheet;
  const allSheets = spreadsheet.getSheets();
  if (allSheets.length > 0) {
    const firstSheet = allSheets[0];
    landingSheet = firstSheet.setName(GLOBAL_VARIABLES.LANDING_SHEET);
  } else {
    landingSheet = spreadsheet.insertSheet(GLOBAL_VARIABLES.LANDING_SHEET);
  }
  const clearRange = landingSheet.getRange("A1:R36");
  clearRange.clearContent();
  clearRange.clearFormat();
  clearRange.breakApart();
  landingSheet.setHiddenGridlines(true);

  // Set Column Widths
  landingSheet.setColumnWidth(1, 10); // Column A (Spacer)
  landingSheet.setColumnWidths(2, 5, 140); // B, C, D, E, F
  landingSheet.setColumnWidth(7, 10); // G (Spacer)
  landingSheet.setColumnWidths(8, 5, 140); // H, I, J, K, L
  landingSheet.setColumnWidth(10, 20) // J (Spacer)
  landingSheet.setColumnWidth(13, 10) // M (Spacer)

  // Set Row Heights
  landingSheet.setRowHeights(1, 20, 25); // Adjust heights for rows 1-20

  // Greeting Message
  landingSheet.getRange("B2:F19").merge()
    .setValue(LANDINGPAGE_VARIABLES.GREETING_MESSAGE)
    .setFontSize(13).setWrap(true).setVerticalAlignment("middle");

  // Title
  landingSheet.getRange("H2:L2").merge()
    .setValue(LANDINGPAGE_VARIABLES.SESSION_TITLE)
    .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("center");

  // Participant Explanation
  landingSheet.getRange("H5:I6").merge()
    .setValue(LANDINGPAGE_VARIABLES.PARTICIPANTS)
    .setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  // Rounds Explanation
  landingSheet.getRange("K5:L6").merge()
    .setValue(LANDINGPAGE_VARIABLES.ROUNDS)
    .setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  // Ideas Explanation
  landingSheet.getRange("H8:I9").merge()
    .setValue(LANDINGPAGE_VARIABLES.IDEAS)
    .setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  // Minutes Explanation
  landingSheet.getRange("K8:L9").merge()
    .setValue(LANDINGPAGE_VARIABLES.TIME)
    .setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  // Total Ideas Explanation
  landingSheet.getRange("I11:K15").merge()
    .setValue(LANDINGPAGE_VARIABLES.TOTALS)
    .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  // Add background colors
  landingSheet.getRange("A1:G20").setBackground("#f0f0f0"); // Light grey for left
  landingSheet.getRange("H1:M20").setBackground("#e0e8f0"); // Light blue for right

  // Add institution logo
  if (imageUrl) {
  landingSheet.getRange("K18:L19").merge().setHorizontalAlignment("right").setFormula(`=IMAGE("${imageUrl}"${sep} 1)`)
  }
  SpreadsheetApp.flush(); // Apply all changes
  properties.setProperty('landingPageCreated', 'true');
}


/*
----- FUNCTION TO CREATE PROGRESS TRACKING AND PARTICIPANT SHEETS -----
*/
function createMultipleWorksheets() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getDocumentProperties();

  if (properties.getProperty('sessionCreated') === 'true') {
    const response = ui.alert(
      create_session_alert_header,
      create_session_alert_text,
      ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) {
      return;
    }
  }
  // Try to get existing tracking sheet or create new one
  try {
    trackingSheet = spreadsheet.insertSheet(GLOBAL_VARIABLES.MODERATOR_SHEET);
  } catch (e) {
    // If sheet already exists, get it and clear it
    trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
    trackingSheet.clear();
  }

  // Set up headers
  trackingSheet.getRange(1, 1, 1, (participantCount * ideasCount)).setValues([MODERATOR_VARIABLES.IDEAS_TABLE]);

  // Format headers
  const headerRange = trackingSheet.getRange(1, 1, 1, (participantCount * ideasCount));
  headerRange.setBackground(MODERATOR_VARIABLES.COLORS.LightSteelBlue);
  headerRange.setFontWeight("bold");

  // for timestamps
  trackingSheet.getRange((roundCount + 26), 1, 1, (participantCount * ideasCount)).setValues([MODERATOR_VARIABLES.IDEAS_TABLE]);
  const timestampHeaderRange = trackingSheet.getRange((roundCount + 26), 1, 1, (participantCount * ideasCount));
  timestampHeaderRange.setBackground(MODERATOR_VARIABLES.COLORS.Grey);
  timestampHeaderRange.setFontWeight("bold");

  // Set up progress tracking table
  const roundsRange = trackingSheet.getRange((roundCount + 4), 8, 1, roundCount);
  roundsRange.setValues([GLOBAL_VARIABLES.ROUNDS]);

  // Add participants
  const participantRange = trackingSheet.getRange((roundCount + 5), 5, participantCount, 1);
  const participantValues = GLOBAL_VARIABLES.PARTICIPANT.map(name => [name]);
  participantRange.setValues(participantValues);


  // Create checkboxes for each round
  const checkboxRange = trackingSheet.getRange((roundCount + 5), 8, participantCount, roundCount);
  checkboxRange.insertCheckboxes();

  // Format tracking table
  const headerRow = trackingSheet.getRange((roundCount + 4), 6, 1, roundCount+2);
  headerRow.setBackground(MODERATOR_VARIABLES.COLORS.Gold);
  headerRow.setFontWeight("bold");

  // Add borders to the tracking table
  const tableRange = trackingSheet.getRange((roundCount + 4), 6, (participantCount+1), roundCount+2);
  tableRange.setBorder(true, true, true, true, true, true);

  trackingSheet.getRange((roundCount + 4), 1, 1, 1).setValue(MODERATOR_VARIABLES.CURRENT_ROUND);
  trackingSheet.getRange((roundCount + 4), 1, 1, 1).setBackground(MODERATOR_VARIABLES.COLORS.Gold);
  trackingSheet.getRange((roundCount + 4), 2, 1, 1).setValue(1);

  trackingSheet.getRange((roundCount + 6), 1, 1, 1).setValue(MODERATOR_VARIABLES.ROUND_END_PHRASE);
  trackingSheet.getRange((roundCount + 6), 2, 1, 1).setValue(GLOBAL_VARIABLES.TIME_IS_UP);
  trackingSheet.getRange((roundCount + 6), 1, 4, 1).setBackground(MODERATOR_VARIABLES.COLORS.Gold);

  trackingSheet.getRange((roundCount + 7), 1, 1, 1).setValue(GLOBAL_VARIABLES.STARTING);
  trackingSheet.getRange((roundCount + 7), 2, 1, 1).setValue(GLOBAL_VARIABLES.STARTING);

  trackingSheet.getRange((roundCount + 8), 1, 1, 1).setValue(GLOBAL_VARIABLES.STOPPED);
  trackingSheet.getRange((roundCount + 8), 2, 1, 1).setValue(GLOBAL_VARIABLES.STOPPED);

  trackingSheet.getRange((roundCount + 9), 1, 1, 1).setValue(GLOBAL_VARIABLES.TIME_LEFT);
  trackingSheet.getRange((roundCount + 9), 2, 1, 1).setValue(GLOBAL_VARIABLES.MINUTES);

  trackingSheet.getRange((roundCount + 10), 1, 1, 1).setValue(GLOBAL_VARIABLES.TIME_LEFT);
  trackingSheet.getRange((roundCount + 10), 1, 1, 1).setBackground(MODERATOR_VARIABLES.COLORS.Green);
  trackingSheet.getRange((roundCount + 10), 2, 1, 1).setValue(TIMER_STR)

  trackingSheet.getRange((roundCount + 12), 1, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS);
  trackingSheet.getRange((roundCount + 12), 1, 1, 1).setBackground(MODERATOR_VARIABLES.COLORS.Green);
  trackingSheet.getRange((roundCount + 12), 2, 1, 1).setValue(GLOBAL_VARIABLES.SESSION_FOCUS);

  trackingSheet.getRange((roundCount + 13), 1, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS);
  trackingSheet.getRange((roundCount + 13), 2, 1, 1).setValue(GLOBAL_VARIABLES.SESSION_FOCUS);
  trackingSheet.getRange((roundCount + 13), 1, 4, 1).setBackground(MODERATOR_VARIABLES.COLORS.Gold);

  trackingSheet.getRange((roundCount + 14), 1, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS[1]);
  trackingSheet.getRange((roundCount + 14), 2, 1, 1).setValue(MODERATOR_VARIABLES.SESSION_START);
  trackingSheet.getRange((roundCount + 15), 1, 1, 1).setValue(GLOBAL_VARIABLES.ROUND_CHANGE[0]);
  trackingSheet.getRange((roundCount + 15), 2, 1, 1).setValue(GLOBAL_VARIABLES.ROUND_CHANGE[1]);
  trackingSheet.getRange((roundCount + 16), 1, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS[2]);
  trackingSheet.getRange((roundCount + 16), 2, 1, 1).setValue(GLOBAL_VARIABLES.SESSION_COMPLETE);

  // Create participant sheets
  const templateSheetName = "Template_DoNotDelete";
  let templateSheet = spreadsheet.getSheetByName(templateSheetName);
  if (templateSheet) {
    templateSheet.clear(); // Clear if it exists from a previous run
  } else {
    templateSheet = spreadsheet.insertSheet(templateSheetName);
  }

  // Set Column Widths
  templateSheet.setColumnWidth(1, 10); // Column A (empty)

  // Round Column Widths
  for (let col = 2; col <= 2 + roundCount; col++) {
    templateSheet.setColumnWidth(col, col === 2 ? 100 : 180); // B is narrower
  }

  // Set Row Heights
  templateSheet.setRowHeight(1, 10); // Row 1 (empty)
  templateSheet.setRowHeight(2, 20); // Row 2 Round Timer
  templateSheet.setRowHeight(3, 20); // Row 3
  templateSheet.setRowHeight(4, 10); // Row 4 (empty)
  templateSheet.setRowHeight(5, 20); // Row 5 Round Headers

  // Idea Rows Heights set to 70
  for (let row = 6; row < 6 + ideasCount; row++) {
    templateSheet.setRowHeight(row, 70);
  }

  templateSheet.setRowHeight(6 + ideasCount, 40); // GLOBAL_VARIABLES.CHECK_IDEAS
  templateSheet.setRowHeight(7 + ideasCount, 20); // Checkbox


  // Enable text wrapping
  const formatRange = templateSheet.getRange(2, 2, (ideasCount + 6), (roundCount + 1));
  formatRange.setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Set headers
  const languageColumnLetter = String.fromCharCode(67 + roundCount); // 'C' is 67 in ASCII
  templateSheet.getRange("B2").setFormula(`=GOOGLETRANSLATE(${GLOBAL_VARIABLES.MODERATOR_SHEET}!A${roundCount + 10}${sep} "${GLOBAL_VARIABLES.SESSION_LANGUAGE}"${sep} ${languageColumnLetter}2)`);
  templateSheet.getRange("C2").setFormula(`=GOOGLETRANSLATE(${GLOBAL_VARIABLES.MODERATOR_SHEET}!A${roundCount + 12}${sep} "${GLOBAL_VARIABLES.SESSION_LANGUAGE}"${sep} ${languageColumnLetter}2)`);

  // Set round timer and focus question
  templateSheet.getRange("B3").setFormula(`=GOOGLETRANSLATE(${GLOBAL_VARIABLES.MODERATOR_SHEET}!B${roundCount + 10}${sep} "${GLOBAL_VARIABLES.SESSION_LANGUAGE}"${sep} ${languageColumnLetter}2)`);
  templateSheet.getRange("C3").setFormula(`=GOOGLETRANSLATE(${GLOBAL_VARIABLES.MODERATOR_SHEET}!B${roundCount + 12}${sep} "${GLOBAL_VARIABLES.SESSION_LANGUAGE}"${sep} ${languageColumnLetter}2)`);

  // Set round headers
  const roundRange = templateSheet.getRange(5, 3, 1, roundCount);
  roundRange.setValues([GLOBAL_VARIABLES.ROUNDS.map(round => `=GOOGLETRANSLATE("${round}"${sep} "${GLOBAL_VARIABLES.SESSION_LANGUAGE}"${sep} ${languageColumnLetter}2)`)]);

  // Set idea labels
  templateSheet.getRange(6, 2, ideasCount, 1).setValues(GLOBAL_VARIABLES.IDEAS.map(idea => [`=GOOGLETRANSLATE("${idea}"${sep} "${GLOBAL_VARIABLES.SESSION_LANGUAGE}"${sep} ${languageColumnLetter}2)`]));

  // Add submission text and checkbox
  templateSheet.getRange((ideasCount + 6), 3, 1, 1).setFormula(`=GOOGLETRANSLATE("${GLOBAL_VARIABLES.CHECK_IDEAS}"${sep} "${GLOBAL_VARIABLES.SESSION_LANGUAGE}"${sep} ${languageColumnLetter}2)`);
  templateSheet.getRange((ideasCount + 7), 3, 1, 1).insertCheckboxes();

  // Merge focus cell
  templateSheet.getRange(3, 3, 1, roundCount).merge();
  templateSheet.getRange(2, 3, 1, roundCount).merge();

  // Header styling
  templateSheet.getRangeList(["B2:C2", `B5:${String.fromCharCode(66 + roundCount)}5`, `B6:B${6 + ideasCount - 1}`]).setBackground(MODERATOR_VARIABLES.COLORS.LightBlue);

  templateSheet.getRange(6, 3, (ideasCount + 1), 1).setBackground(MODERATOR_VARIABLES.COLORS.LightGreen);
  templateSheet.getRange(5, 2, (ideasCount + 1), (roundCount + 1)).setBorder(true, true, true, true, true, true, MODERATOR_VARIABLES.COLORS.Grey, SpreadsheetApp.BorderStyle.SOLID);

  templateSheet.getRange(2, 2, 2, (roundCount + 1)).setBorder(true, true, true, true, true, true, MODERATOR_VARIABLES.COLORS.Grey, SpreadsheetApp.BorderStyle.SOLID); // Border focus and round timer

  templateSheet.setHiddenGridlines(true) // Hide gridlines

  function addConditionalFormatting() {
    // Define the cell to apply conditional formatting
    var range = templateSheet.getRange(2, 3, 1, roundCount);
    var rules = templateSheet.getConditionalFormatRules();
    const blue_formula = `=INDIRECT("${GLOBAL_VARIABLES.MODERATOR_SHEET}!B${roundCount + 12}")=INDIRECT("${GLOBAL_VARIABLES.MODERATOR_SHEET}!B${roundCount + 13}")`;
    const green_formula = `=INDIRECT("${GLOBAL_VARIABLES.MODERATOR_SHEET}!B${roundCount + 12}")=INDIRECT("${GLOBAL_VARIABLES.MODERATOR_SHEET}!B${roundCount + 16}")`;
    const otherwiseFormula = '=TRUE';

    // Define the conditional formatting rules
    var newRules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(blue_formula) // Apply rule when the text is "Blue"
        .setBackground(MODERATOR_VARIABLES.COLORS.LightBlue)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(green_formula) // Apply rule when the text is "Green"
        .setBackground(MODERATOR_VARIABLES.COLORS.Green)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(otherwiseFormula)
        .setBackground(MODERATOR_VARIABLES.COLORS.Yellow)
        .setRanges([range])
        .build(),
    ];
    // Add the new rules to the sheet
    templateSheet.setConditionalFormatRules(rules.concat(newRules));
  }
  addConditionalFormatting();

  for (let i = 0; i < participantCount; i++) {
    const sheetName = GLOBAL_VARIABLES.PARTICIPANT[i];

    // If a sheet with this name already exists, delete it before copying.
    let existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    // This one call replaces dozens of formatting calls inside the loop.
    const newSheet = templateSheet.copyTo(spreadsheet).setName(sheetName);

    // Set sheet language
    var languages = [
    'af', 'sq', 'am', 'ar', 'hy', 'as', 'ay', 'az', 'bm', 'eu', 'be', 'bn', 'bho', 'bs', 'bg', 'ca', 'ceb', 'ny', 'zh', 'zh-TW', 'co', 'hr', 'cs', 'da', 'dv', 'doi', 'nl', 'en', 'eo', 'et', 'ee', 'tl', 'fi', 'fr', 'fy', 'gl', 'lg', 'ka', 'de', 'el', 'gn', 'gu', 'ht', 'ha', 'haw', 'iw', 'hi', 'hmn', 'hu', 'is', 'ig', 'ilo', 'id', 'ga', 'it', 'ja', 'jw', 'kn', 'kk', 'km', 'rw', 'gom', 'ko', 'kri', 'ku', 'ckb', 'ky', 'lo', 'la', 'lv', 'ln', 'lt', 'lb', 'mk', 'mg', 'ms', 'ml', 'mt', 'mi', 'mr', 'mni-Mtei', 'lus', 'mn', 'my', 'ne', 'no', 'or', 'om', 'ps', 'fa', 'pl', 'pt', 'pa', 'qu', 'ro', 'ru', 'sm', 'sa', 'gd', 'nso', 'sr', 'st', 'sn', 'sd', 'si', 'sk', 'sl', 'so', 'es', 'su', 'sw', 'sv', 'tg', 'ta', 'tt', 'te', 'th', 'ti', 'ts', 'tr', 'tk', 'ak', 'uk', 'ur', 'ug', 'uz', 'vi', 'cy', 'xh', 'yi', 'yo', 'zu'
    ];
    var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(languages)
        .setAllowInvalid(false) // This prevents typing a language not in the list
        .build();
    var lan_cell = newSheet.getRange(2, (roundCount + 3), 1, 1)
    const participantLanguageValue = GLOBAL_VARIABLES.PARTICIPANT_SHEET_LANGUAGE[i];
    lan_cell.setValue(participantLanguageValue);
    lan_cell.setDataValidation(rule);
  }

  templateSheet.hideSheet();

  // Set formulas for participant tracking
  for (let i = 0; i < participantCount; i++) {
    let row = roundCount + 5 + i;
    let sheetName = GLOBAL_VARIABLES.PARTICIPANT[i]; // Participant sheet name

    for (let j = 0; j < roundCount; j++) {
      // Columns C to H in participant sheet → columns 6 to 6+roundCount-1 in tracking sheet
      let participantColumnLetter = String.fromCharCode(67 + j); // 'C' is 67 in ASCII
      let formula = `=${sheetName}!${participantColumnLetter}${ideasCount+7}`;
      trackingSheet.getRange(row, 8 + j).setFormula(formula);

      let languageColumnLetter = String.fromCharCode(67 + roundCount); // 'C' is 67 in ASCII
      let participantLangugageRange = trackingSheet.getRange(row, 6, 1, 1);
      participantLangugageRange.setFormula(`=${sheetName}!${languageColumnLetter}2`);

      trackingSheet.getRange(row, 7).setFormula(`=GOOGLETRANSLATE(B${roundCount + 12}${sep} "${GLOBAL_VARIABLES.SESSION_LANGUAGE}"${sep} F${row})`);
    }
  }
  // Get the total number of sheets
  var sheetCount = spreadsheet.getSheets().length;

  // Move "Moderator ProgerssTracking sheet" to the last position (index = sheetCount - 1)
  spreadsheet.setActiveSheet(trackingSheet);
  spreadsheet.moveActiveSheet(sheetCount);
  spreadsheet.deleteSheet(templateSheet);
  properties.setProperty('sessionCreated', 'true');
}


/*
----- FUNCTION TO ADD UI MENU -----
*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(MODERATOR_VARIABLES.MENU.Tools)
      .addItem(MODERATOR_VARIABLES.MENU.LandingPage, "createBrainwritingLandingPage")
      .addItem(MODERATOR_VARIABLES.MENU.SessionElements, "createMultipleWorksheets")
      .addSeparator()
      .addItem(MODERATOR_VARIABLES.MENU.Start, "StartSession")
      .addItem(MODERATOR_VARIABLES.MENU.SubmitNext, "SubmitData")
      .addSeparator()
      .addItem(MODERATOR_VARIABLES.MENU.PrepareData, "PrepData")
      .addItem(MODERATOR_VARIABLES.MENU.ColabEnvironment, "completeAutomationAndGuideToColab")
      .addSeparator()
      .addItem(MODERATOR_VARIABLES.MENU.ResetSessionSetup, "resetSessionFlags")
      .addToUi();
}


// Function to reset setup flags and allow re-creation
function resetSessionFlags() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    reset_session_flags_header,
    reset_session_alert,
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    const properties = PropertiesService.getDocumentProperties();
    properties.deleteProperty('landingPageCreated');
    properties.deleteProperty('sessionCreated');
  }
}

// Round countdown
function updateCountdownStatus(secondsLeft) {
  const trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  let message = '';

  if (secondsLeft <= 0) {
    var time_is_up_text = trackingSheet.getRange((roundCount + 6), 2, 1, 1).getValue();
    message = time_is_up_text;
  } else if (secondsLeft <= 60) {
    message = GLOBAL_VARIABLES.ONE_MIN_LEFT;
  } else {
    message = Math.ceil(secondsLeft / 60) + GLOBAL_VARIABLES.MINS_LEFT;
  }

  trackingSheet.getRange(roundCount + 10, 2).setValue(message);
}

function roundStartTimestamp() {
  const trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  const round_num = trackingSheet.getRange((roundCount + 4), 2, 1, 1).getValue();
  const columnRound = round_num + 7;
  const rowRound = roundCount + participantCount + 5;
  const timestampNow = new Date();
  trackingSheet.getRange(rowRound, columnRound).setValue(timestampNow);
}

function getRoundMinutes() {
  const trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  // trackingSheet.getRange(roundCount + 8, 2).setFormula("=NOW()"); // redundant

  let TIMER_INTERVAL = 10;
  var round_min = trackingSheet.getRange((roundCount + 9), 2, 1, 1).getValue();
  let TIMER_DURATION = round_min * 60;
  let TIMER_STR_IF_CHANGED = String(round_min).padStart(2, '0') + ':00';

  function showTimerSidebar() {
    const html = HtmlService.createHtmlOutput(`
      <html>
        <head>
          <base target="_top">
          <style>
            body {
              margin: 0;
              padding: 5px;
              font-family: Arial, sans-serif;
              width: 120px;
              height: 60px;
              background: #ffffff;
            }
            #timerContainer {
              width: 105px;
              height: 42.5px;
              background: #f1f1f1;
              text-align: center;
              line-height: 42.5px;
              font-size: 1.6em;
              border-radius: 6px;
              margin: 5px auto;
            }
            #startBtn {
              display: block;
              margin: 5px auto;
              padding: 4px 10px;
              font-size: 0.9em;
              cursor: pointer;
            }
          </style>
        </head>
        <body>
          <div id="timerContainer">${TIMER_STR_IF_CHANGED}</div>
          <button id="startBtn" onclick="startTimer()">${MODERATOR_VARIABLES.START_TIMER}</button>
          <script>
            let duration = ${TIMER_DURATION};
            let interval = ${TIMER_INTERVAL};

            function startTimer() {
              document.getElementById('startBtn').style.display = 'none';
              const display = document.getElementById('timerContainer');
              google.script.run.roundStartTimestamp();

              interval = setInterval(function() {
                let minutes = Math.floor(duration / 60);
                let seconds = duration % 60;
                const timeStr =
                  String(minutes).padStart(2, '0') + ':' + String(seconds).padStart(2, '0');
                display.textContent = timeStr;

                // Report to Sheets every 10 seconds
                if (duration % 10 === 0) {
                  google.script.run.updateCountdownStatus(duration);
                }

                if (--duration < 0) {
                  clearInterval(interval);
                  display.textContent = '00:00';
                  google.script.run.updateCountdownStatus(0);
                }
              }, 1000);
            }
          </script>
        </body>
      </html>
    `).setTitle(GLOBAL_VARIABLES.TIME_LEFT).setWidth(130);

    SpreadsheetApp.getUi().showSidebar(html);
  }

  showTimerSidebar();
}


// Start session with a note in yellow background
function StartSession() {
  var trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  var session_started = trackingSheet.getRange((roundCount + 14), 1, 1, 2).getValues();
  var starting_text = trackingSheet.getRange((roundCount + 7), 2, 1, 1).getValue();

  trackingSheet.getRange((roundCount + 12), 1, 1, 2).setValues(session_started);
  trackingSheet.getRange(roundCount + 10, 2).setValue(starting_text);

  SpreadsheetApp.flush();
  getRoundMinutes();

  // wait 10 seconds for participants to read the text
  Utilities.sleep(10000); // 10 seconds waiting time
  var focus_if_changed = trackingSheet.getRange((roundCount + 13), 1, 1, 2).getValues();
  trackingSheet.getRange((roundCount + 12), 1, 1, 2).setValues(focus_if_changed);
  SpreadsheetApp.flush();
}


// End session with a note in green background
function SessionEnd() {
  var trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  var session_end = trackingSheet.getRange((roundCount + 16), 1, 1, 2).getValues();

  trackingSheet.getRange((roundCount + 12), 1, 1, 2).setValues(session_end);
}


/*
----- FUNCTION TO SUBMIT DATA AND CHANGE ROUND (with Interleaved Logic) -----
*/
function SubmitData() {
  const trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  const round_num = trackingSheet.getRange((roundCount + 4), 2, 1, 1).getValue();
  var col_num = round_num + 2

  const round_change_text = trackingSheet.getRange((roundCount + 15), 1, 1, 2).getValues();
  const stopped_text = trackingSheet.getRange((roundCount + 8), 2, 1, 1).getValue();
  trackingSheet.getRange((roundCount + 12), 1, 1, 2).setValues(round_change_text);
  trackingSheet.getRange(roundCount + 10, 2).setValue(stopped_text);
  SpreadsheetApp.flush();
  getRoundMinutes();

  const allIdeasData = trackingSheet.getRange(2, 1, roundCount, participantCount * ideasCount).getValues();

  let allNewIdeas = [];
  for (let k = 0; k < participantCount; k++) {
    const sheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.PARTICIPANT[k]);
    const participantIdeas = sheet.getRange(6, col_num, ideasCount, 1).getValues();
    allNewIdeas.push(participantIdeas.flat());

    for (var i=0; i<(ideasCount+2); i++) {
      sheet.getRange(6+i, col_num).clearContent();
      sheet.getRange(6+i, col_num).setBackground(MODERATOR_VARIABLES.COLORS.LightGrey);
    }
    sheet.getRange((ideasCount+6), col_num, 2, 1).clearFormat().removeCheckboxes();
  }

  const newIdeasForThisRound = allNewIdeas.flat();

  if (newIdeasForThisRound.length > 0) {
      trackingSheet.getRange(round_num + 1, 1, 1, newIdeasForThisRound.length).setValues([newIdeasForThisRound]);
  }

  allIdeasData[round_num - 1] = newIdeasForThisRound;

  // Perform the "Paper Swap" using the selected algorithm
  const positiveModulo = (a, n) => ((a % n) + n) % n;

  // Dynamic Harmonic Sweep Pre-calculation
  const cumulativePointer = (round_num > 1) ? positiveModulo((round_num - 1) * harmonicShift, (participantCount - 1)) : 0;

  const universalContributors = [participantCount - 1].concat(Array.from({ length: participantCount - 2 }, (_, idx) => idx + 1));
  const universalIdeaChains = [];
  for (let k = 0; k < ideasCount; k++) {
      const chain = [];
      for(let n = 0; n < participantCount - 1; n++) {
          chain.push(universalContributors[positiveModulo(n - k, participantCount - 1)]);
      }
      universalIdeaChains.push(chain);
  }

  for (let i = 0; i < participantCount; i++) {
    const sheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.PARTICIPANT[i]);
    const languageColumnLetter = String.fromCharCode(67 + roundCount);
    let formulasForSheet = Array(ideasCount).fill(null).map(() => Array(round_num));

    for (let j = 0; j < round_num; j++) {
      for (let k = 0; k < ideasCount; k++) {
        let finalSourceIndex;

        if (GLOBAL_VARIABLES.IDEA_SWAP_ALGORITHM === "DynamicHarmonicSweep") {
            const viewPointer = positiveModulo(cumulativePointer + j, (participantCount - 1));
            const baseSourceIndex = universalIdeaChains[k][viewPointer];
            finalSourceIndex = positiveModulo(baseSourceIndex + i, participantCount);
        } else { // Fallback to "CascadingRoundRobin"
          const shift = round_num - j;
          finalSourceIndex = positiveModulo(i - shift, participantCount);
        }

        const idea_row_index = j;
        const idea_col_index = (finalSourceIndex * ideasCount) + k;
        const ideaValue = allIdeasData[idea_row_index][idea_col_index];

        const cleanIdea = ideaValue ? String(ideaValue).replace(/"/g, '""') : "";
        const formula = cleanIdea ? `=GOOGLETRANSLATE("${cleanIdea}"${sep} "auto"${sep} ${languageColumnLetter}2)` : "";
        formulasForSheet[k][j] = formula;
      }
    }

    if(round_num > 0) {
      const rangeToUpdate = sheet.getRange(6, 3, ideasCount, round_num);
      rangeToUpdate.clearContent();
      rangeToUpdate.setFormulas(formulasForSheet);
    }
  }

  const next_input_col = col_num + 1;

  if (round_num == roundCount) {
    trackingSheet.getRange((roundCount + 4), 2, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS[2]);
    SessionEnd();
    for (let k = 0; k < participantCount; k++) {
      const sheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.PARTICIPANT[k]);
      const lastInputRange = sheet.getRange(6, col_num, ideasCount + 2, 1);
      lastInputRange.setBackground(MODERATOR_VARIABLES.COLORS.LightGrey);
      sheet.getRange(ideasCount + 6, col_num, 2, 1).clearFormat().removeCheckboxes();
    }
    return;
  }

  for (let i = 0; i < participantCount; i++) {
    const sheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.PARTICIPANT[i]);
    const languageColumnLetter = String.fromCharCode(67 + roundCount);
    sheet.getRange(6, next_input_col, (ideasCount + 1), 1).setBackground(MODERATOR_VARIABLES.COLORS.LightGreen);
    sheet.getRange((ideasCount + 6), next_input_col).setFormula(`=GOOGLETRANSLATE("${GLOBAL_VARIABLES.CHECK_IDEAS}"${sep} "${GLOBAL_VARIABLES.SESSION_LANGUAGE}"${sep} ${languageColumnLetter}2)`);
    sheet.getRange((ideasCount + 7), next_input_col).insertCheckboxes();
  }

  trackingSheet.getRange((roundCount + 4), 2, 1, 1).setValue(round_num + 1);
  const focus_if_changed = trackingSheet.getRange((roundCount + 13), 1, 1, 2).getValues();
  trackingSheet.getRange((roundCount + 12), 1, 1, 2).setValues(focus_if_changed);
  SpreadsheetApp.flush();
}


/*
----- FUNCTION TO PREPARE DATA FOR ANALYSIS -----
*/
function PrepData() {

  try {
    prepSheet = spreadsheet.insertSheet(MODERATOR_VARIABLES.DATA_PREP.SheetName);
  } catch (e) {
    // If sheet already exists, get it and clear it
    prepSheet = spreadsheet.getSheetByName(MODERATOR_VARIABLES.DATA_PREP.SheetName);
    prepSheet.clear();
  }
  prepSheet.getRange(1, 1).setValue(MODERATOR_VARIABLES.DATA_PREP.IdeaID);
  prepSheet.getRange(1, 2).setValue(MODERATOR_VARIABLES.DATA_PREP.IdeaTimestamp);
  prepSheet.getRange(1, 3).setValue(MODERATOR_VARIABLES.DATA_PREP.IdeaRoundStartTimestamp);
  prepSheet.getRange(1, 4).setValue(MODERATOR_VARIABLES.DATA_PREP.IdeaRawColumn);
  prepSheet.getRange(1, 5).setValue(MODERATOR_VARIABLES.DATA_PREP.TranslateColumn);
  prepSheet.getRange(1, 6).setValue(MODERATOR_VARIABLES.DATA_PREP.ManualCategorization);

  trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);

  var ideasTableRange = trackingSheet.getRange(2, 1, roundCount, participantCount*ideasCount);
  var timestampTableRange = trackingSheet.getRange((roundCount + 27), 1, roundCount, participantCount*ideasCount);
  var roundStartTimestampTableRange = trackingSheet.getRange((roundCount + participantCount + 5), 8, 1, roundCount);
  var ideasTableValues = ideasTableRange.getValues();
  var timestampTableValues = timestampTableRange.getValues();
  var roundStartTimestampValues = roundStartTimestampTableRange.getValues();

  // Create an array to hold the values for the single column
  var ideasSingleColumn = [];
  var timestampSingleColumn = [];
  var participantSingleColumn = [];
  var roundTimestampSingleColumn = [];

  // Loop through rows and columns to extract values
  for (var row = 0; row < ideasTableValues.length; row++) {
    for (var col = 0; col < ideasTableValues[row].length; col++) {
      participantSingleColumn.push([`${MODERATOR_VARIABLES.IDEAS_TABLE[col]}-Round${row + 1}`]);
      ideasSingleColumn.push([ideasTableValues[row][col]]); // Wrap value in array for single-column format
      timestampSingleColumn.push([timestampTableValues[row][col]]);
      roundTimestampSingleColumn.push([roundStartTimestampValues[0][row]]);

    }
  }

  // Output the single column starting at D2 cell
  prepSheet.getRange(2, 4).offset(0, 0, ideasSingleColumn.length, 1).setValues(ideasSingleColumn);
  prepSheet.getRange(2, 3).offset(0, 0, roundTimestampSingleColumn.length, 1).setValues(roundTimestampSingleColumn);
  prepSheet.getRange(2, 2).offset(0, 0, timestampSingleColumn.length, 1).setValues(timestampSingleColumn);
  prepSheet.getRange(2, 1).offset(0, 0, participantSingleColumn.length, 1).setValues(participantSingleColumn);

  // Translate the data
  for (var row = 2; row < (participantCount*ideasCount*roundCount + 2); row++) { // Start from row 2 to skip the header
    var sourceCell = "D" + row; // Reference the source cell in column A
    var targetCell = "E" + row; // Reference the target cell in column B
    var formula = `=GOOGLETRANSLATE(${sourceCell}${sep}"${MODERATOR_VARIABLES.DATA_PREP.SessionLanguage}"${sep}"${MODERATOR_VARIABLES.DATA_PREP.TranslatedLanguage}")`;
    prepSheet.getRange(targetCell).setFormula(formula);
  };
}


// Pop-up that guides the moderator for Colab environment
function completeAutomationAndGuideToColab() {
  SpreadsheetApp.flush(); // Ensure all data is written to the sheet

  let ui = HtmlService.createHtmlOutput(ANALYSIS_VARIABLES.COLAB_POPUP)
      .setWidth(800)
      .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(ui, ANALYSIS_VARIABLES.POPUP_TITLE);
}


// Captures the exact time a participant writes an idea
function onEdit(e) {
  const ideaTimestamp = new Date()
  const trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  const round_num = trackingSheet.getRange((roundCount + 4), 2, 1, 1).getValue();
  const range = e.range;
  const sheet = range.getSheet();
  const name = sheet.getName();
  const parts = GLOBAL_VARIABLES.PARTICIPANT;
  const pIndex = parts.indexOf(name);
  if (pIndex < 0) return;
  const ideaCoordinateCol = range.getColumn()
  const ideaCoordinateRow = range.getRow()

  if (ideaCoordinateRow >= 6 + ideasCount ||
      ideaCoordinateRow < 6 ||
      ideaCoordinateCol < round_num + 2
      ) {
    return;
  }

  const targetRow = roundCount + 26 + (ideaCoordinateCol-2);
  const startCol = pIndex * ideasCount + (ideaCoordinateRow-5);

  trackingSheet.getRange(targetRow, startCol, 1, 1).setValue(ideaTimestamp);
}
