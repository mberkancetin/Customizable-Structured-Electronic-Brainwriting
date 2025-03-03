// @ts-nocheck

/*
MIT License

Copyright (c) 2025 Mahmut Berkan Çetin & Selim Gündüz

This software is licensed under the MIT License. 
See the LICENSE file or visit https://opensource.org/licenses/MIT for details.
*/

/*
Creator: Mahmut Berkan Çetin and Selim Gündüz
Title: Online 6-3-5 Brainwriting Session: Utilizing Google Sheets with Google Apps Script
Description: This script automates the 6-3-5 brainwriting method in Google Sheets.
How to Cite: Cetin, Mahmut Berkan & Gunduz, Selim (2025). Online 6-3-5 Brainwriting Session: Utilizing Google Sheets with Google Apps Script. protocols.io, https://dx.doi.org/10.17504/protocols.io.x54v9r3xzv3e/v1
*/

const GLOBAL_VARIABLES = {
  SESSION_FOCUS: "Ex. How can companies improve the efficiency of Corporate Responsibility policies?",
  PARTICIPANT: ["Participant1", "Participant2", "Participant3", "Participant4", "Participant5", "Participant6"], // please use under_score or CamelCase
  MODERATOR_SHEET: "ProgressTracking", // please use under_score or CamelCase 
  DATE: "Date",
  FOCUS: ["Focus", "Session Started", "Session Completed"],
  ROUNDS: ["Round 1", "Round 2", "Round 3", "Round 4", "Round 5", "Round 6"],
  IDEAS: ["Idea 1", "Idea 2", "Idea 3"],
  SESSION_START: "Please submit 3 ideas for Round 1. Check the box when done. Time limit: 5 minutes per round.",
  CHECK_IDEAS: ["Check the box to submit your ideas for the first round!",
                "Check the box to submit your ideas for the second round!",
                "Check the box to submit your ideas for the third round!",
                "Check the box to submit your ideas for the fourth round!",
                "Check the box to submit your ideas for the fifth round!",
                "Check the box to submit your ideas for the last round!"],
  ROUND_CHANGE: ["Round-Robin Sheet Update in Progress", 
                "Please wait for the script to finish before proceeding to avoid errors."],
  SESSION_COMPLETE: "Thank you for your valuable contributions! Your answers have been successfully saved."
};

const MODERATOR_VARIABLES = {
  MENU: {
    Tools: "Workshop Tools",
    Start: "Start Session",
    SubmitNext: "Submit & Next"
  },
  IDEAS_TABLE : [
    `P1-idea1`,`P1-idea2`,`P1-idea3`,
    `P2-idea1`,`P2-idea2`,`P2-idea3`,
    `P3-idea1`,`P3-idea2`,`P3-idea3`,
    `P4-idea1`,`P4-idea2`,`P4-idea3`,
    `P5-idea1`,`P5-idea2`,`P5-idea3`,
    `P6-idea1`,`P6-idea2`,`P6-idea3`
  ],
  COLORS : {
    LightSteelBlue: "LightSteelBlue",
    LightBlue: "#cfe2f3",
    Gold: "#ffd700",
    LightGreen: "#d9ead3",
    Green: "#90ee90",
    Yellow: "yellow",
    Grey: "grey",
    LightGrey: "#efefef"
  }
}

function createMultipleWorksheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Try to get existing tracking sheet or create new one
  try {
    trackingSheet = spreadsheet.insertSheet(GLOBAL_VARIABLES.MODERATOR_SHEET);
  } catch (e) {
    // If sheet already exists, get it and clear it
    trackingSheet = spreadsheet.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
    trackingSheet.clear();
  }

  // Set up headers
  trackingSheet.getRange("A1:R1").setValues([MODERATOR_VARIABLES.IDEAS_TABLE]);
  
  // Format headers
  const headerRange = trackingSheet.getRange("A1:R1");
  headerRange.setBackground(MODERATOR_VARIABLES.COLORS.LightSteelBlue); 
  headerRange.setFontWeight("bold");
  
  // Set up progress tracking table
  const roundsRange = trackingSheet.getRange("F10:K10");
  roundsRange.setValues([GLOBAL_VARIABLES.ROUNDS]);
  
   // Add participants
  const participantRange = trackingSheet.getRange("E11:E16");
  participantRange.setValues([
    [GLOBAL_VARIABLES.PARTICIPANT[0]],
    [GLOBAL_VARIABLES.PARTICIPANT[1]],
    [GLOBAL_VARIABLES.PARTICIPANT[2]],
    [GLOBAL_VARIABLES.PARTICIPANT[3]],
    [GLOBAL_VARIABLES.PARTICIPANT[4]],
    [GLOBAL_VARIABLES.PARTICIPANT[5]]
  ]);
  
  // Create checkboxes for each round
  const checkboxRange = trackingSheet.getRange("F11:K16");
  checkboxRange.insertCheckboxes();
  
  // Format tracking table
  const headerRow = trackingSheet.getRange("F10:K10");
  headerRow.setBackground(MODERATOR_VARIABLES.COLORS.Gold); 
  headerRow.setFontWeight("bold");
  
  // Add borders to the tracking table
  const tableRange = trackingSheet.getRange("F10:K16");
  tableRange.setBorder(true, true, true, true, true, true);

  trackingSheet.getRange("A10").setValue("Current Round:");
  trackingSheet.getRange("A10").setBackground(MODERATOR_VARIABLES.COLORS.Gold); 
  trackingSheet.getRange("B10").setValue(1);

  trackingSheet.getRange("A16").setValue(GLOBAL_VARIABLES.DATE);
  trackingSheet.getRange("A16").setBackground(MODERATOR_VARIABLES.COLORS.Gold); 
  trackingSheet.getRange("A18").setValue(GLOBAL_VARIABLES.FOCUS);
  trackingSheet.getRange("A18").setBackground(MODERATOR_VARIABLES.COLORS.Gold);  

  trackingSheet.getRange("B16").setValue(new Date().toLocaleDateString());
  trackingSheet.getRange("B18").setValue(GLOBAL_VARIABLES.SESSION_FOCUS);

  // Create 6 sheets
  for (let i = 0; i <= 5; i++) {
   
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
    
    // Set column widths
    sheet.setColumnWidth(1, 10);  // Column A
    sheet.setColumnWidth(2, 80);  // Column B
    sheet.setColumnWidth(3, 180);  // Column C
    sheet.setColumnWidth(4, 180);  // Column D
    sheet.setColumnWidth(5, 180);  // Column E
    sheet.setColumnWidth(6, 180);  // Column F
    sheet.setColumnWidth(7, 180);  // Column G
    sheet.setColumnWidth(8, 180);  // Column H

    // Set row heights
    sheet.setRowHeight(1, 10);  // Row 1
    sheet.setRowHeight(2, 20);  // Row 2
    sheet.setRowHeight(3, 20);  // Row 3
    sheet.setRowHeight(4, 10);  // Row 4
    sheet.setRowHeight(5, 20);  // Row 5
    sheet.setRowHeight(6, 70);  // Row 6
    sheet.setRowHeight(7, 70);  // Row 7
    sheet.setRowHeight(8, 70);  // Row 8
    sheet.setRowHeight(9, 40);  // Row 9
    sheet.setRowHeight(10, 20);  // Row 10

    // Enable text wrapping
    sheet.getRange("B2:H10").setWrap(true);
    sheet.getRange("B2:H10").setHorizontalAlignment("center");
    sheet.getRange("B2:H10").setVerticalAlignment("middle");

    // Set headers
    sheet.getRange("B2").setFormula("=" + GLOBAL_VARIABLES.MODERATOR_SHEET + "!A16"); 
    sheet.getRange("C2").setFormula("=" + GLOBAL_VARIABLES.MODERATOR_SHEET + "!A18"); 
    
    // Set date and focus question
    sheet.getRange("B3").setFormula("=" + GLOBAL_VARIABLES.MODERATOR_SHEET + "!B16");
    sheet.getRange("C3").setFormula("=" + GLOBAL_VARIABLES.MODERATOR_SHEET + "!B18");
    
    // Set round headers
    const roundsRange = sheet.getRange("C5:H5");
    roundsRange.setValues([GLOBAL_VARIABLES.ROUNDS]);
    
    // Set idea labels
    sheet.getRange("B6:B8").setValues(GLOBAL_VARIABLES.IDEAS.map(idea => [idea]));
    
    // Add submission text and checkbox
    sheet.getRange("C9").setValue(GLOBAL_VARIABLES.CHECK_IDEAS[0]);
    sheet.getRange("C10").insertCheckboxes();
        
    // Header styling
    const headerRanges = ["B2", "C2:H2", "B5:H5", "B6:B8"];
    headerRanges.forEach(range => {
      sheet.getRange(range).setBackground(MODERATOR_VARIABLES.COLORS.LightBlue);  
    });
    
    sheet.getRange("C6:C9").setBackground(MODERATOR_VARIABLES.COLORS.LightGreen); 
    sheet.getRange("B5:H8").setBorder(true, true, true, true, true, true, MODERATOR_VARIABLES.COLORS.Grey, SpreadsheetApp.BorderStyle.SOLID); // Border the main table
    
    // Merge focus cell
    sheet.getRange("C3:H3").merge();
    sheet.getRange("C2:H2").merge();

    sheet.getRange("B2:H3").setBorder(true, true, true, true, true, true, MODERATOR_VARIABLES.COLORS.Grey, SpreadsheetApp.BorderStyle.SOLID); // Border focus and date

    sheet.setHiddenGridlines(true) // Hide gridlines

    function addConditionalFormatting() {
      // Define the cell to apply conditional formatting
      var range = sheet.getRange("C2:H2");
      var rules = sheet.getConditionalFormatRules();

      // Define the conditional formatting rules
      var newRules = [
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(GLOBAL_VARIABLES.FOCUS[0]) // Apply rule when the text is "Blue"
          .setBackground(MODERATOR_VARIABLES.COLORS.LightBlue)
          .setRanges([range]) // Apply to cell B18
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(GLOBAL_VARIABLES.FOCUS[2]) // Apply rule when the text is "Green"
          .setBackground(MODERATOR_VARIABLES.COLORS.Green) 
          .setRanges([range]) // Apply to cell B18
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextDoesNotContain(GLOBAL_VARIABLES.FOCUS[0]) // Apply rule otherwise
          .whenTextDoesNotContain(GLOBAL_VARIABLES.FOCUS[2])
          .setBackground(MODERATOR_VARIABLES.COLORS.Yellow) 
          .setRanges([range]) // Apply to cell B18
          .build(),        
      ];
      // Add the new rules to the sheet
      sheet.setConditionalFormatRules(rules.concat(newRules));
    }
    addConditionalFormatting();
  }

  // Set formulas to participant tracking
  for (let i = 0; i <= 5; i++) {
    var cell_no = 11 + i
    trackingSheet.getRange("F" + cell_no).setFormula("=" + GLOBAL_VARIABLES.PARTICIPANT[i] + "!C10"); 
    trackingSheet.getRange("G" + cell_no).setFormula("=" + GLOBAL_VARIABLES.PARTICIPANT[i] + "!D10"); 
    trackingSheet.getRange("H" + cell_no).setFormula("=" + GLOBAL_VARIABLES.PARTICIPANT[i] + "!E10"); 
    trackingSheet.getRange("I" + cell_no).setFormula("=" + GLOBAL_VARIABLES.PARTICIPANT[i] + "!F10"); 
    trackingSheet.getRange("J" + cell_no).setFormula("=" + GLOBAL_VARIABLES.PARTICIPANT[i] + "!G10"); 
    trackingSheet.getRange("K" + cell_no).setFormula("=" + GLOBAL_VARIABLES.PARTICIPANT[i] + "!H10");  
  };
  // Get the total number of sheets
  var sheetCount = spreadsheet.getSheets().length;

  // Move "Sheet1" to the last position (index = sheetCount - 1)
  spreadsheet.setActiveSheet(trackingSheet);
  spreadsheet.moveActiveSheet(sheetCount);
}

// Function to add UI menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(MODERATOR_VARIABLES.MENU.Tools)
      .addItem(MODERATOR_VARIABLES.MENU.Start, "StartSession")
      .addItem(MODERATOR_VARIABLES.MENU.SubmitNext, "SubmitData")
      .addToUi();
}

// Start session with a note in yellow background
function StartSession() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var form = ss.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  form.getRange(18, 1, 1, 2).setValues([[GLOBAL_VARIABLES.FOCUS[1], GLOBAL_VARIABLES.SESSION_START]]);

  SpreadsheetApp.flush();

  // wait 10 seconds for participants to read the text
  Utilities.sleep(10000); // 10 seconds waiting time
  form.getRange(18, 1, 1, 2).setValues([[GLOBAL_VARIABLES.FOCUS[0], GLOBAL_VARIABLES.SESSION_FOCUS]]);
}

// End session with a note in green background
function SessionEnd() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var form = ss.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  form.getRange(18, 1).setValue(GLOBAL_VARIABLES.FOCUS[2]);
  form.getRange(18, 2).setValue(GLOBAL_VARIABLES.SESSION_COMPLETE);
}

// Submit Data & Next
function SubmitData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var form = ss.getSheetByName(GLOBAL_VARIABLES.MODERATOR_SHEET);
  var round_num = form.getRange("B10").getValue();
  var x = [1, 4, 7, 10, 13, 16]
  var round_robin = x.slice(-1*round_num).concat(x);  
  var col_num = round_num + 2

  form.getRange(18, 1, 1, 2).setValues([[GLOBAL_VARIABLES.ROUND_CHANGE[0], GLOBAL_VARIABLES.ROUND_CHANGE[1]]]);
  SpreadsheetApp.flush();

  for (var k=0; k<6; k++) {
    var data = ss.getSheetByName(GLOBAL_VARIABLES.PARTICIPANT[k]);
    var values = [[data.getRange(6, col_num).getValue(),
                data.getRange(7, col_num).getValue(),
                data.getRange(8, col_num).getValue(),]];       
    form.getRange(round_num+1, 1+(3*k), 1, 3).setValues(values);
    
    for (var i=0; i<5; i++) {
      data.getRange(6+i, col_num).clearContent();
      data.getRange(6+i, col_num).setBackground(MODERATOR_VARIABLES.COLORS.LightGrey); 
    }
    data.getRange(9, col_num).clearFormat(); 
    data.getRange(10, col_num).clearFormat(); 
    data.getRange(10, col_num).removeCheckboxes();
  }

  for (var i=0; i<6; i++) {
    var data = ss.getSheetByName(GLOBAL_VARIABLES.PARTICIPANT[i]);
    for (var j = 0; j < round_num; j++) {
      var sourceRange = form.getRange(2 + j, round_robin[i + j], 1, 3).getValues(); // Calculate the source range dynamically
      var columnValues = sourceRange[0].map(v => [v]); // Map the 1x3 row array to a 3x1 column array

      // Set the values in the corresponding target column
      data.getRange(6, 3 + j, 3, 1).setValues(columnValues);
    }
    if (round_num > 5) {
      form.getRange("B10").setValue(GLOBAL_VARIABLES.FOCUS[2]);
      SessionEnd();
    } else {
      data.getRange(9, col_num+1).setValue(GLOBAL_VARIABLES.CHECK_IDEAS[round_num]);
      data.getRange(10, col_num+1).insertCheckboxes();
      data.getRange(6, col_num+1, 4, 1).setBackground(MODERATOR_VARIABLES.COLORS.LightGreen); 
      form.getRange(18, 1, 1, 2).setValues([[GLOBAL_VARIABLES.FOCUS[0], GLOBAL_VARIABLES.SESSION_FOCUS]]);
    }     
  }
  form.getRange("B10").setValue(form.getRange("B10").getValue() + 1);   // increment round  
}

