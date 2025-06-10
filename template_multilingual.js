/*
MIT License

Copyright (c) 2025 Mahmut Berkan Çetin & Selim Gündüz

This software is licensed under the MIT License.
See the LICENSE file or visit https://opensource.org/licenses/MIT for details.
*/

/*
Creator: Mahmut Berkan Çetin and Selim Gündüz
Title: Conducting Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script.
Description: This script automates the customizable structured electronic brainwriting method in Google Sheets.
How to Cite:

Çetin, Mahmut Berkan & Gündüz, Selim (2025). Protocol for Conducting Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script. protocols.io, https://dx.doi.org/10.17504/protocols.io.x54v9r3xzv3e/v1


*/

const GLOBAL_VARIABLES = {
  SESSION_FOCUS: {{SESSION_FOCUS_JS_STRING}},
  LANDING_SHEET: {{LANDING_SHEET_JS_STRING}},
  PARTICIPANT: {{PARTICIPANT_LIST_AS_JS_ARRAY}},
  MODERATOR_SHEET: {{MODERATOR_SHEET_JS_STRING}},
  TIME_LEFT: {{TIME_LEFT_JS_STRING}},
  TIME_IS_UP: {{TIME_IS_UP_JS_STRING}},
  MINUTES: {{MINUTES_INT}},
  FOCUS: {{FOCUS_LIST_AS_JS_ARRAY}},
  ROUNDS: {{ROUNDS_LIST_AS_JS_ARRAY}},
  IDEAS: {{IDEAS_LIST_AS_JS_ARRAY}},
  CHECK_IDEAS: {{CHECK_IDEAS_JS_STRING}},
  ROUND_CHANGE: {{ROUND_CHANGE_LIST_AS_JS_ARRAY}},
  SESSION_COMPLETE: {{SESSION_COMPLETE_JS_STRING}}
};

const roundCount = GLOBAL_VARIABLES.ROUNDS.length;
const ideasCount = GLOBAL_VARIABLES.IDEAS.length;
const participantCount = GLOBAL_VARIABLES.PARTICIPANT.length;
const totalIdeas = participantCount * ideasCount * roundCount;
const totalTime = roundCount * GLOBAL_VARIABLES.MINUTES;

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
    TranslateColumn: {{DATA_PREP_TranslateColumn_JS_STRING}},
    ManualCategorization: {{DATA_PREP_ManualCategorization_JS_STRING}}
  },
  COLORS : {{COLORS_OBJECT_AS_JS_OBJECT}}
};

const imageUrl = {{IMAGE_URL_JS_STRING}};
const colabGitHubUrl = {{COLAB_GITHUB_URL_JS_STRING}};
const ideas_prefix = {{LP_IDEAS_PREFIX_JS_STRING}}.replace("{ideasCount}", ideasCount);
const time_prefix = {{LP_TIME_PREFIX_JS_STRING}}.replace("{GLOBAL_VARIABLES.MINUTES}", GLOBAL_VARIABLES.MINUTES);
const totals_prefix = {{LP_TOTALS_TEXT_INFIX_JS_STRING}}.replace("{totalIdeas}", totalIdeas).replace("{totalTime}", totalTime);

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
        <br><b>original_column = "${MODERATOR_VARIABLES.DATA_PREP.IdeaRawColumn}"</b>
        <br><b>translate_column = "${MODERATOR_VARIABLES.DATA_PREP.TranslateColumn}"</b></code>
    </ol>
  `
};

/*
The code begins here.
Feel free to modify, but please review carefully to maintain functionality.
*/

function createBrainwritingLandingPage() {
  const allSheets = spreadsheet.getSheets();

  if (allSheets.length > 0) {
    const firstSheet = allSheets[0];
    landingSheet = firstSheet.setName(GLOBAL_VARIABLES.LANDING_SHEET);
  } else {
    landingSheet = spreadsheet.insertSheet(GLOBAL_VARIABLES.LANDING_SHEET);
  }

  // Clear previous content and formatting in the target area
  const clearRange = landingSheet.getRange("A1:R36");
  clearRange.clearContent();
  clearRange.clearFormat();
  clearRange.breakApart(); // Remove any previous merges
  landingSheet.setHiddenGridlines(true); // Hide gridlines for a cleaner look

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
  landingSheet.getRange("K18:L19").merge().setHorizontalAlignment("right").setFormula(`=IMAGE("${imageUrl}"; 1)`)
  }

  SpreadsheetApp.flush(); // Apply all changes
}

/*
----- FUNCTION TO CREATE PROGRESS TRACKING AND PARTICIPANT SHEETS -----
*/

function createMultipleWorksheets() {
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

  // Set up progress tracking table
  const roundsRange = trackingSheet.getRange((roundCount + 4), 6, 1, roundCount);
  roundsRange.setValues([GLOBAL_VARIABLES.ROUNDS]);

  // Add participants
  const participantRange = trackingSheet.getRange((roundCount + 5), 5, participantCount, 1);
  const participantValues = GLOBAL_VARIABLES.PARTICIPANT.map(name => [name]);
  participantRange.setValues(participantValues);

  // Create checkboxes for each round
  const checkboxRange = trackingSheet.getRange((roundCount + 5), 6, participantCount, roundCount);
  checkboxRange.insertCheckboxes();

  // Format tracking table
  const headerRow = trackingSheet.getRange((roundCount + 4), 6, 1, roundCount);
  headerRow.setBackground(MODERATOR_VARIABLES.COLORS.Gold);
  headerRow.setFontWeight("bold");

  // Add borders to the tracking table
  const tableRange = trackingSheet.getRange((roundCount + 4), 6, (participantCount+1), roundCount);
  tableRange.setBorder(true, true, true, true, true, true);

  trackingSheet.getRange((roundCount + 4), 1, 1, 1).setValue(MODERATOR_VARIABLES.CURRENT_ROUND);
  trackingSheet.getRange((roundCount + 4), 1, 1, 1).setBackground(MODERATOR_VARIABLES.COLORS.Gold);
  trackingSheet.getRange((roundCount + 4), 2, 1, 1).setValue(1);

  trackingSheet.getRange((roundCount + 5), 1, 1, 1).setValue(MODERATOR_VARIABLES.ROUND_END_PHRASE);
  trackingSheet.getRange((roundCount + 5), 1, 1, 1).setBackground(MODERATOR_VARIABLES.COLORS.Gold);
  trackingSheet.getRange((roundCount + 5), 2, 1, 1).setValue(GLOBAL_VARIABLES.TIME_IS_UP);

  trackingSheet.getRange((roundCount + 10), 1, 1, 1).setValue(GLOBAL_VARIABLES.TIME_LEFT);
  trackingSheet.getRange((roundCount + 10), 1, 1, 1).setBackground(MODERATOR_VARIABLES.COLORS.Gold);
  trackingSheet.getRange((roundCount + 12), 1, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS);
  trackingSheet.getRange((roundCount + 12), 1, 1, 1).setBackground(MODERATOR_VARIABLES.COLORS.Gold);

  trackingSheet.getRange((roundCount + 10), 2, 1, 1).setValue("0" + GLOBAL_VARIABLES.MINUTES + ":00")
  trackingSheet.getRange((roundCount + 12), 2, 1, 1).setValue(GLOBAL_VARIABLES.SESSION_FOCUS);

  // Create 6 sheets
  for (let i = 0; i < participantCount; i++) {

    // Create new sheet
    const sheetName = GLOBAL_VARIABLES.PARTICIPANT[i];
    let sheet;

    // Try to get existing sheet or create new one
    try {
      sheet = spreadsheet.insertSheet(sheetName);
    } catch (e) {
      // If sheet already exists, get it and clear it
      sheet = spreadsheet.getSheetByName(sheetName);
      sheet.clear();
    }

    // Set Column Widths
    sheet.setColumnWidth(1, 10); // Column A (empty)

    // Round Column Widths
    for (let col = 2; col <= 2 + roundCount; col++) {
      sheet.setColumnWidth(col, col === 2 ? 100 : 180); // B is narrower
    }

    // Set Row Heights
    sheet.setRowHeight(1, 10); // Row 1 (empty)
    sheet.setRowHeight(2, 20); // Row 2 Round Timer
    sheet.setRowHeight(3, 20); // Row 3
    sheet.setRowHeight(4, 10); // Row 4 (empty)
    sheet.setRowHeight(5, 20); // Row 5 Round Headers

    // Idea Rows Heights set to 70
    for (let row = 6; row < 6 + ideasCount; row++) {
      sheet.setRowHeight(row, 70);
    }

    sheet.setRowHeight(6 + ideasCount, 40); // GLOBAL_VARIABLES.CHECK_IDEAS
    sheet.setRowHeight(7 + ideasCount, 20); // Checkbox


    // Enable text wrapping
    sheet.getRange(2, 2, (ideasCount + 6), (roundCount + 1)).setWrap(true);
    sheet.getRange(2, 2, (ideasCount + 6), (roundCount + 1)).setHorizontalAlignment("center");
    sheet.getRange(2, 2, (ideasCount + 6), (roundCount + 1)).setVerticalAlignment("middle");

    // Set headers
    sheet.getRange("B2").setFormula(`=${GLOBAL_VARIABLES.MODERATOR_SHEET}!A${roundCount + 10}`);
    sheet.getRange("C2").setFormula(`=${GLOBAL_VARIABLES.MODERATOR_SHEET}!A${roundCount + 12}`);

    // Set round timer and focus question
    sheet.getRange("B3").setFormula(`=${GLOBAL_VARIABLES.MODERATOR_SHEET}!B${roundCount + 10}`);
    sheet.getRange("C3").setFormula(`=${GLOBAL_VARIABLES.MODERATOR_SHEET}!B${roundCount + 12}`);

    // Set round headers
    const roundsRange = sheet.getRange(5, 3, 1, roundCount);
    roundsRange.setValues([GLOBAL_VARIABLES.ROUNDS]);

    // Set idea labels
    sheet.getRange(6, 2, ideasCount, 1).setValues(GLOBAL_VARIABLES.IDEAS.map(idea => [idea]));

    // Add submission text and checkbox
    sheet.getRange((ideasCount + 6), 3, 1, 1).setValue(GLOBAL_VARIABLES.CHECK_IDEAS);
    sheet.getRange((ideasCount + 7), 3, 1, 1).insertCheckboxes();

    // Header styling
    const headerRanges = [
      sheet.getRange(2, 2),
      sheet.getRange(2, 3, 1, roundCount),
      sheet.getRange(5, 2, 1, roundCount + 1),
      sheet.getRange(6, 2, ideasCount, 1)
    ];

    headerRanges.forEach(range => {
      range.setBackground(MODERATOR_VARIABLES.COLORS.LightBlue);
    });

    sheet.getRange(6, 3, (ideasCount + 1), 1).setBackground(MODERATOR_VARIABLES.COLORS.LightGreen);
    sheet.getRange(5, 2, (ideasCount + 1), (roundCount + 1)).setBorder(true, true, true, true, true, true, MODERATOR_VARIABLES.COLORS.Grey, SpreadsheetApp.BorderStyle.SOLID); // Border the main table

    // Merge focus cell
    sheet.getRange(3, 3, 1, roundCount).merge();
    sheet.getRange(2, 3, 1, roundCount).merge();

    sheet.getRange(2, 2, 2, (roundCount + 1)).setBorder(true, true, true, true, true, true, MODERATOR_VARIABLES.COLORS.Grey, SpreadsheetApp.BorderStyle.SOLID); // Border focus and round timer

    sheet.setHiddenGridlines(true) // Hide gridlines

    function addConditionalFormatting() {
      // Define the cell to apply conditional formatting
      var range = sheet.getRange(2, 3, 1, roundCount);
      var rules = sheet.getConditionalFormatRules();

      // Define the conditional formatting rules
      var newRules = [
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(GLOBAL_VARIABLES.FOCUS[0]) // Apply rule when the text is "Blue"
          .setBackground(MODERATOR_VARIABLES.COLORS.LightBlue)
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(GLOBAL_VARIABLES.FOCUS[2]) // Apply rule when the text is "Green"
          .setBackground(MODERATOR_VARIABLES.COLORS.Green)
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextDoesNotContain(GLOBAL_VARIABLES.FOCUS[0]) // Apply rule otherwise
          .whenTextDoesNotContain(GLOBAL_VARIABLES.FOCUS[2])
          .setBackground(MODERATOR_VARIABLES.COLORS.Yellow)
          .setRanges([range])
          .build(),
      ];
      // Add the new rules to the sheet
      sheet.setConditionalFormatRules(rules.concat(newRules));
    }
    addConditionalFormatting();
  }

  // Set formulas for participant tracking
  for (let i = 0; i < participantCount; i++) {
    let row = roundCount + 5 + i;
    let sheetName = GLOBAL_VARIABLES.PARTICIPANT[i]; // Participant sheet name

    for (let j = 0; j < roundCount; j++) {
      // Columns C to H in participant sheet → columns 6 to 6+roundCount-1 in tracking sheet
      let participantColumnLetter = String.fromCharCode(67 + j); // 'C' is 67 in ASCII
      let formula = `=${sheetName}!${participantColumnLetter}${ideasCount+7}`;
      trackingSheet.getRange(row, 6 + j).setFormula(formula);
    }
  }
  // Get the total number of sheets
  var sheetCount = spreadsheet.getSheets().length;

  // Move "Moderator ProgerssTracking sheet" to the last position (index = sheetCount - 1)
  spreadsheet.setActiveSheet(trackingSheet);
  spreadsheet.moveActiveSheet(sheetCount);
}

/*
----- FUNCTION TO ADD UI MENU -----
*/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(MODERATOR_VARIABLES.MENU.Tools)
      .addItem(MODERATOR_VARIABLES.MENU.LandingPage, "createBrainwritingLandingPage")
      .addItem(MODERATOR_VARIABLES.MENU.SessionElements, "createMultipleWorksheets")
      .addItem(MODERATOR_VARIABLES.MENU.Start, "StartSession")
      .addItem(MODERATOR_VARIABLES.MENU.SubmitNext, "SubmitData")
      .addItem(MODERATOR_VARIABLES.MENU.PrepareData, "PrepData")
      .addItem(MODERATOR_VARIABLES.MENU.ColabEnvironment, "completeAutomationAndGuideToColab")
      .addToUi();
}

// Round countdown
function calculateRoundLength() {
  var minutesAdded = new Date();
  minutesAdded.setMinutes(minutesAdded.getMinutes() + GLOBAL_VARIABLES.MINUTES);
  return minutesAdded;
}

function getFormulaSeparatorFromSheet() {
  // Reads if seperator is comma: min(1,0) returns 1
  // or if seperator is semicolon: min(1;0) returns 0
  // in the sheet to detect the separator
  const ss = spreadsheet.getActiveSheet();
  const range = ss.getRange(25, 1);

  range.setFormula('=MIN(1,0)');
  SpreadsheetApp.flush();

  const formula = range.getFormula();

  range.clearContent();

  // Detect which separator is used
  return formula.includes(1) ? ';' : ',';
}

const sep = getFormulaSeparatorFromSheet();

function getRoundMinutes() {
  const trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);


  trackingSheet.getRange(roundCount + 8, 2).setFormula("=NOW()");
  trackingSheet.getRange(roundCount + 9, 2).setValue(calculateRoundLength());

  const rA = roundCount + 9;
  const rB = roundCount + 8;
  const rC = roundCount + 5;

  const timerFormula = `=IF((B${rA} <= B${rB})${sep} B${rC}${sep} TEXT(QUOTIENT(ROUND((B${rA} - B${rB}) * 86400)${sep} 60)${sep} "00") & ":" & TEXT(MOD(ROUND((B${rA} - B${rB}) * 86400)${sep} 60)${sep} "00"))`;

  trackingSheet.getRange(roundCount + 10, 2).setFormula(timerFormula);
}

// Start session with a note in yellow background
function StartSession() {
  var trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  trackingSheet.getRange((roundCount + 12), 1, 1, 2).setValues([[GLOBAL_VARIABLES.FOCUS[1], MODERATOR_VARIABLES.SESSION_START]]);

  SpreadsheetApp.flush();

  // wait 10 seconds for participants to read the text
  Utilities.sleep(10000); // 10 seconds waiting time
  trackingSheet.getRange((roundCount + 12), 1, 1, 2).setValues([[GLOBAL_VARIABLES.FOCUS[0], GLOBAL_VARIABLES.SESSION_FOCUS]]);
  getRoundMinutes();
}

// End session with a note in green background
function SessionEnd() {
  var trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  trackingSheet.getRange((roundCount + 12), 1).setValue(GLOBAL_VARIABLES.FOCUS[2]);
  trackingSheet.getRange((roundCount + 12), 2).setValue(GLOBAL_VARIABLES.SESSION_COMPLETE);
}

/*
----- FUNCTION TO SUBMIT DATA AND CHANGE ROUND -----
*/

function SubmitData() {
  var trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  var round_num = trackingSheet.getRange((roundCount + 4), 2, 1, 1).getValue();
  // var x = [1, 4, 7, 10, 13, 16]
  var x = Array.from({ length: participantCount }, (_, i) => 1 + i * ideasCount);

  var round_robin = x.slice(-1*round_num).concat(x);
  var col_num = round_num + 2

  trackingSheet.getRange((roundCount + 12), 1, 1, 2).setValues([[GLOBAL_VARIABLES.ROUND_CHANGE[0], GLOBAL_VARIABLES.ROUND_CHANGE[1]]]);
  SpreadsheetApp.flush();

  for (var k=0; k<participantCount; k++) {
    var sheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.PARTICIPANT[k]);
    var values = [
                  Array.from({ length: ideasCount }, (_, i) => sheet.getRange(6 + i, col_num).getValue())
                ];
    trackingSheet.getRange(round_num+1, 1+(ideasCount*k), 1, ideasCount).setValues(values);

    for (var i=0; i<(ideasCount+2); i++) {
      sheet.getRange(6+i, col_num).clearContent();
      sheet.getRange(6+i, col_num).setBackground(MODERATOR_VARIABLES.COLORS.LightGrey);
    }
    sheet.getRange((ideasCount+6), col_num).clearFormat();
    sheet.getRange((ideasCount+7), col_num).clearFormat();
    sheet.getRange((ideasCount+7), col_num).removeCheckboxes();
  }

  for (var i=0; i<participantCount; i++) {
    var sheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.PARTICIPANT[i]);
    for (var j = 0; j < round_num; j++) {
      var sourceRange = trackingSheet.getRange(2 + j, round_robin[i + j], 1, ideasCount).getValues(); // Calculate the source range dynamically
      var columnValues = sourceRange[0].map(v => [v]);

      // Set the values in the corresponding target column
      sheet.getRange(6, 3 + j, ideasCount, 1).setValues(columnValues);
    }
    if (round_num == roundCount) {
      trackingSheet.getRange((roundCount+4), 2, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS[2]);
      SessionEnd();
    } else {
      sheet.getRange((ideasCount+6), col_num+1).setValue(GLOBAL_VARIABLES.CHECK_IDEAS);
      sheet.getRange((ideasCount+7), col_num+1).insertCheckboxes();
      sheet.getRange(6, col_num+1, (ideasCount+1), 1).setBackground(MODERATOR_VARIABLES.COLORS.LightGreen);
    }
  }
  trackingSheet.getRange((roundCount+4), 2, 1, 1).setValue(trackingSheet.getRange((roundCount+4), 2, 1, 1).getValue() + 1);   // increment round
  trackingSheet.getRange((roundCount+12), 1, 1, 2).setValues([[GLOBAL_VARIABLES.FOCUS[0], GLOBAL_VARIABLES.SESSION_FOCUS]]);
  getRoundMinutes();
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
  prepSheet.getRange("A1").setValue(MODERATOR_VARIABLES.DATA_PREP.IdeaRawColumn);

  trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);

  var tableRange = trackingSheet.getRange(2, 1, roundCount, participantCount*ideasCount);
  var tableValues = tableRange.getValues();

  // Create an array to hold the values for the single column
  var singleColumn = [];

  // Loop through rows and columns to extract values
  for (var row = 0; row < tableValues.length; row++) {
    for (var col = 0; col < tableValues[row].length; col++) {
      singleColumn.push([tableValues[row][col]]); // Wrap value in array for single-column format
    }
  }

  // Output the single column starting at A2 cell
  var outputRange = prepSheet.getRange("A2");
  outputRange.offset(0, 0, singleColumn.length, 1).setValues(singleColumn);

  if (MODERATOR_VARIABLES.DATA_PREP.SessionLanguage == "auto") {
    prepSheet.getRange("B1").setValue(MODERATOR_VARIABLES.DATA_PREP.TranslateColumn);
    prepSheet.getRange("C1").setValue(MODERATOR_VARIABLES.DATA_PREP.ManualCategorization);

    // Translate the data
    for (var row = 2; row < (participantCount*ideasCount*roundCount + 2); row++) { // Start from row 2 to skip the header
      var sourceCell = "A" + row; // Reference the source cell in column A
      var targetCell = "B" + row; // Reference the target cell in column B
      var formula = `=GOOGLETRANSLATE(${sourceCell}${sep}"${MODERATOR_VARIABLES.DATA_PREP.SessionLanguage}"${sep}"${MODERATOR_VARIABLES.DATA_PREP.TranslatedLanguage}")`;
      prepSheet.getRange(targetCell).setFormula(formula);
    };
  } else {
      prepSheet.getRange("B1").setValue(MODERATOR_VARIABLES.DATA_PREP.ManualCategorization);
  }
}

// Pop-up that guides the moderator for Colab environment
function completeAutomationAndGuideToColab() {
  SpreadsheetApp.flush(); // Ensure all data is written to the sheet

  let ui = HtmlService.createHtmlOutput(ANALYSIS_VARIABLES.COLAB_POPUP)
      .setWidth(800)
      .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(ui, ANALYSIS_VARIABLES.POPUP_TITLE);
}
