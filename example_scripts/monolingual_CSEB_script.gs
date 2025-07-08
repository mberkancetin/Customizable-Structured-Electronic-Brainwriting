/*
 * SPDX-FileCopyrightText: 2025 Mahmut Berkan Çetin <m.berkancetin@gmail.com> & Selim Gündüz <sgunduz@atu.edu.tr>
 *
 * SPDX-License-Identifier: MIT
 */


/*
Creators: Mahmut Berkan Çetin and Selim Gündüz
Title: Conducting Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script.
Description: This script automates the customizable structured electronic brainwriting method in Google Sheets.
How to Cite:

Çetin, Mahmut Berkan & Gündüz, Selim (2025). Protocol for Conducting Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script. protocols.io, https://dx.doi.org/10.17504/protocols.io.x54v9r3xzv3e/v1

*/


const GLOBAL_VARIABLES = {
  SESSION_FOCUS: "The effects of local textile companies on your daily life and your community.",
  LANDING_SHEET: "Welcome",
  PARTICIPANT: ["Participant1", "Participant2", "Participant3", "Participant4", "Participant5", "Participant6"],
  MODERATOR_SHEET: "ProgressTracking",
  TIME_LEFT: "Timer",
  MINS_LEFT: " minutes left",
  ONE_MIN_LEFT: "1 minute left",
  TIME_IS_UP: "Time's up!",
  MINUTES: 5,
  FOCUS: ["Focus", "Session Started", "Session Completed"],
  ROUNDS: ["Round 1", "Round 2", "Round 3", "Round 4", "Round 5", "Round 6"],
  IDEAS: ["Idea 1", "Idea 2", "Idea 3"],
  CHECK_IDEAS: "Check the box to submit your ideas",
  ROUND_CHANGE: ["Round-Robin Sheet Update in Progress", "Please wait for the script to finish before proceeding to avoid errors."],
  SESSION_COMPLETE: "Thank you for your valuable contributions! Your answers have been successfully saved."
};

const roundCount = GLOBAL_VARIABLES.ROUNDS.length;
const ideasCount = GLOBAL_VARIABLES.IDEAS.length;
const participantCount = GLOBAL_VARIABLES.PARTICIPANT.length;
const totalIdeas = participantCount * ideasCount * roundCount;
const totalTime = roundCount * GLOBAL_VARIABLES.MINUTES;

const MODERATOR_VARIABLES = {
  MENU: {"Tools": "Workshop Tools", "LandingPage": "Create Landing Page", "SessionElements": "Create Session", "Start": "Start Session", "SubmitNext": "Submit & Next", "PrepareData": "Prepare Data", "ColabEnvironment": "Create Colab Environment"},
  SESSION_START: "Please submit " + ideasCount + " ideas for Round 1. Check the box when done.",
  IDEAS_TABLE : GLOBAL_VARIABLES.PARTICIPANT.flatMap(x => GLOBAL_VARIABLES.IDEAS.map(y => x + "-" + y)),
  ROUND_END_PHRASE: "Round End Phrase:",
  CURRENT_ROUND: "Current Round:",
  DATA_PREP: {
    SheetName: "PrepData",
    SessionLanguage: "auto",
    IdeaRawColumn: "RawIdea",
    TranslatedLanguage: "en",
    TranslateColumn: "Translation",
    ManualCategorization: "ManualCategories"
  },
  COLORS : {"LightSteelBlue": "LightSteelBlue", "LightBlue": "#cfe2f3", "Gold": "#ffd700", "LightGreen": "#d9ead3", "Green": "#90ee90", "Yellow": "yellow", "Grey": "grey", "LightGrey": "#efefef"},
  START_TIMER: ""
};

const imageUrl = "";
const colabGitHubUrl = "https://colab.research.google.com/github/mberkancetin/Customizable-Structured-Electronic-Brainwriting/blob/main/DataAnalysis.ipynb";
const ideas_prefix = "💡 Ideas per Round: \n{ideasCount} per person".replace("{ideasCount}", ideasCount);
const time_prefix = "⏱️ Time per Round: \n{GLOBAL_VARIABLES.MINUTES} minutes".replace("{GLOBAL_VARIABLES.MINUTES}", GLOBAL_VARIABLES.MINUTES);
const totals_prefix = "At the end of the session, group can produce total \n{totalIdeas} ideas in just {totalTime} minutes".replace("{totalIdeas}", totalIdeas).replace("{totalTime}", totalTime);

const LANDINGPAGE_VARIABLES = {
  GREETING_MESSAGE: "Dear Participants, \n\nWelcome and thank you for your interest in this structured electronic brainwriting session! Our goal today is to collaboratively generate innovative ideas regarding: The effects of local textile companies on your daily life and your community.. \n\nAt the core of this session, there is a no-judgment policy. We encourage you to force your imagination beyond its limits, regardless of the common beliefs or possibilities. When ideas are exchanged each round, it may inspire another participant to come up with something innovative! \n\nYou will be guided through rounds of generating ideas and building upon the ideas of others. Full instructions and details were provided in your email. Please review them if you haven't already. \n\nPlease navigate to the worksheet tab assigned to you. You'll find detailed instructions in the email sent earlier. Let's innovate together!",
  SESSION_TITLE: participantCount + "-" + roundCount + "-" + ideasCount + "-" + GLOBAL_VARIABLES.MINUTES + " " + "Structured Electronic Brainwriting",
  PARTICIPANTS: "👥 Participants: \n" + participantCount,
  ROUNDS: "🔄 Total Rounds: \n" + roundCount,
  IDEAS: ideas_prefix,
  TIME: time_prefix,
  TOTALS: totals_prefix
};

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const sheetId = spreadsheet.getId();

const ANALYSIS_VARIABLES = {
  POPUP_TITLE : "Proceed to Data Analysis in Colab",
  COLAB_POPUP: `
    <p><b>"Next Steps for Analysis:"</b></p>
    <ol>
      <li>"Click this link to open the analysis notebook in Google Colab:"
          <br><a href="${colabGitHubUrl}" target="_blank">"Open Analysis Notebook in Colab"</a>
          <br><small>("(This will open the notebook directly from GitHub. If you want to save your work in the notebook, use \"File > Save a copy in Drive\" in Colab.)")</small>
      </li>
      <li>
        "Once the Colab notebook is open, find the first cell labeled \"SETUP\"."
      </li>
      <li>
        "Copy the following variables and paste it into SETUP field in Colab:"
      </li>
        <code><br><b>spreadsheet_id = "${sheetId}"</b>
        <br><b>sheet_name = "${MODERATOR_VARIABLES.DATA_PREP.SheetName}"</b>
        <br><b>original_column = "${MODERATOR_VARIABLES.DATA_PREP.IdeaRawColumn}"</b>
        <br><b>translate_column = "${MODERATOR_VARIABLES.DATA_PREP.TranslateColumn}"</b></code>
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
  const range = ss.getRange(25, 1);

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
  landingSheet.getRange("K18:L19").merge().setHorizontalAlignment("right").setFormula(`=IMAGE("${imageUrl}"${sep} 1)`)
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

  trackingSheet.getRange((roundCount + 14), 1, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS);
  trackingSheet.getRange((roundCount + 14), 2, 1, 1).setValue(GLOBAL_VARIABLES.SESSION_FOCUS);
  trackingSheet.getRange((roundCount + 14), 1, 4, 1).setBackground(MODERATOR_VARIABLES.COLORS.Gold);
  trackingSheet.getRange((roundCount + 15), 1, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS[1]);
  trackingSheet.getRange((roundCount + 15), 2, 1, 1).setValue(MODERATOR_VARIABLES.SESSION_START);
  trackingSheet.getRange((roundCount + 16), 1, 1, 2).setValue(GLOBAL_VARIABLES.ROUND_CHANGE[0]);
  trackingSheet.getRange((roundCount + 17), 1, 1, 1).setValue(GLOBAL_VARIABLES.FOCUS[2]);
  trackingSheet.getRange((roundCount + 17), 2, 1, 1).setValue(GLOBAL_VARIABLES.SESSION_COMPLETE);

  trackingSheet.getRange((roundCount + 10), 2, 1, 1).setValue(TIMER_STR)
  trackingSheet.getRange((roundCount + 9), 2, 1, 1).setValue(GLOBAL_VARIABLES.MINUTES)
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
// in progress

function updateCountdownStatus(secondsLeft) {
  const trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  let message = '';

  if (secondsLeft <= 0) {
    message = GLOBAL_VARIABLES.TIME_IS_UP;
  } else if (secondsLeft <= 60) {
    message = GLOBAL_VARIABLES.ONE_MIN_LEFT;
  } else {
    message = Math.ceil(secondsLeft / 60) + GLOBAL_VARIABLES.MINS_LEFT;
  }

  trackingSheet.getRange(roundCount + 10, 2).setValue(message);
}

function getRoundMinutes() {
  const trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  trackingSheet.getRange(roundCount + 8, 2).setFormula("=NOW()");

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
// in progress


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
