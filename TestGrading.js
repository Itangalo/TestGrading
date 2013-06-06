/**
 * Some global variables.
 */
var TOTAL_COLUMN = 3; // Column used for grand total.
var GRADE_COLUMN = 2; // Column used for grade.
var FIRST_STUDENT_ROW = 5; // Row with first student entry.
var LAST_STUDENT_ROW = 5; // Row with last student entry.
var SCORE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Poäng"); // The scoring sheet, built by this script.

/**
 * Adds the custom menu when opening the spreadsheet.
 */
function onOpen() {
  buildMenu();
};

/**
 * Adds a custom menu.
 */
function buildMenu() {
  var menuEntries = [];
  if (sheetExists("Poäng")) {
    menuEntries.push({name : "Ta bort kalkylblad för poäng", functionName : "removeScoreSheet"});
  }
  else {
    menuEntries.push({name : "Bygg kalkylblad för poäng", functionName : "buildScoreSheet"});
  }
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Poängmall", menuEntries);
}

/**
 * Checks whether a given sheet exists in the active spreadsheet.
 */
function sheetExists(sheetName) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getName();
    return true;
  }
  catch (err) {
    return false;
  }
}

/**
 * Helper function to check whether an array of strings has only empty values or not.
 */
function arrayHasValues(arrayToCheck) {
  for (var i in arrayToCheck) {
    if (arrayToCheck[i] !== "") {
      return true;
    }
  }
  return false;
}

/**
 * Helper function for adding a formula for point averages.
 */
function addAverageFormula(column) {
  var columnA1 = SCORE_SHEET.getRange(2, column).getA1Notation().slice(0, -1);
  var rangeA1 = columnA1 + FIRST_STUDENT_ROW + ":" + columnA1 + LAST_STUDENT_ROW;
  SCORE_SHEET.getRange(4, column).setFormula("=if(count(" + rangeA1 + ")>0;average(" + rangeA1 + ")/" + columnA1 + "3;0)").setNumberFormat("0%");
}

/**
 * Menu callback for removing the score sheet.
 */
function removeScoreSheet() {
  var confirm = Browser.msgBox("Är du säker på att du vill radera poängbladet?", Browser.Buttons.OK_CANCEL);
  if (confirm == "ok") {
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SCORE_SHEET);
    SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();
  }
  buildMenu();
}

/**
 * Sets up the sheet for entering test scores. This is the big function.
 */
function buildScoreSheet() {
  if (sheetExists("Poäng")) {
    Browser.msgBox("Poängblad finns redan.");
    return;
  }

  // Get the number of students to create scoring sheet for. To save some manual copy-paste work.
  var numberOfStudents = parseInt(Browser.inputBox("Antal elever"));
  if (numberOfStudents < 1) {
    numberOfStudents = 1;
  }
  LAST_STUDENT_ROW = 4 + numberOfStudents;

  // Get the data for building the scoring sheet, and parse it a bit.
  var buildInfoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Maxpoäng");
  var buildInfo = buildInfoSheet.getSheetValues(1, 1, buildInfoSheet.getLastRow(), buildInfoSheet.getLastColumn());
  var bgColors = buildInfoSheet.getRange(1, 1, buildInfoSheet.getLastRow(), 1).getBackgrounds();
  var pointCategories = buildInfo.shift();
  bgColors.shift();
  pointCategories.shift();

  // Create the sheet, and update the global variable keeping trach of this sheet.
  var scoreSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Poäng");
  SCORE_SHEET = scoreSheet;

  // The variable scoreSheetColumn will keep track on where in the scoring sheet we are.
  // We start just before the column for total points, and start adding things.
  var scoreSheetColumn = TOTAL_COLUMN - 1;

  // First, add columns for keeping track of the point totals, both grand total and
  // total for each point category.
  SCORE_SHEET.getRange(2, scoreSheetColumn, 3, 1).setValues([["Alla delar"], ["Max"], ["Medel"]]);
  scoreSheetColumn = TOTAL_COLUMN;
  SCORE_SHEET.getRange(2, scoreSheetColumn).setValue("Totalt");
  SCORE_SHEET.getRange(3, scoreSheetColumn).setFormula("=0");
  addAverageFormula(scoreSheetColumn);
  for (var i in pointCategories) {
    scoreSheetColumn++;
    SCORE_SHEET.getRange(2, scoreSheetColumn).setValue(pointCategories[i]);
    SCORE_SHEET.getRange(3, scoreSheetColumn).setFormula("=0");
    addAverageFormula(scoreSheetColumn);
  }
  var sumRangeA1 = SCORE_SHEET.getRange(3, TOTAL_COLUMN + 1, 1, scoreSheetColumn - TOTAL_COLUMN).getA1Notation();
  SCORE_SHEET.getRange(3, TOTAL_COLUMN, 2, 1).setFormulas([["=sum(" + sumRangeA1 + ")"], ["=0"]]);
  // Copy the formula for grand total to the row for the first student. This row will
  // be used as a template for all students.
  SCORE_SHEET.getRange(3, TOTAL_COLUMN).copyTo(SCORE_SHEET.getRange(FIRST_STUDENT_ROW, TOTAL_COLUMN));

  // Freeze columns and rows, to make the sheet more easy to read.
  SCORE_SHEET.setFrozenColumns(scoreSheetColumn);
  SCORE_SHEET.setFrozenRows(FIRST_STUDENT_ROW - 1);

  // Iterate through the declaration of questions + maximum scores, and test sections.
  for (var row in buildInfo) {
    var questionName = buildInfo[row].shift();
    // If a row in the build info sheet has values for maximum scores, it means it is a question.
    // Let's add a column for the question.
    if (arrayHasValues(buildInfo[row])) {
      var columnStart = scoreSheetColumn + 1;
      // Check for each possible point category, to see if we should create a column for it.
      for (var i in buildInfo[row]) {
        if (buildInfo[row][i] !== "") {
          // Tick the column counter up one step, add information about point category
          // and maximum score, and set the relevant background color.
          scoreSheetColumn++;
          SCORE_SHEET.getRange(1, scoreSheetColumn, 3, 1).setValues([[questionName], [pointCategories[i]], [buildInfo[row][i]]]);
          addAverageFormula(scoreSheetColumn);
          SCORE_SHEET.getRange(1, scoreSheetColumn, LAST_STUDENT_ROW).setBackground(bgColors[row]);
        }
      }
      // If a question has more than one point category, it is represented by multiple
      // columns. Merge the name cells, to make look like one question a bit more.
      SCORE_SHEET.getRange(1, columnStart, 1, scoreSheetColumn - columnStart + 1).merge().setHorizontalAlignment("center");
    }

    // If a row *doesn't* have any maximum scores, it is a new test section. Let's mark
    // this.
    else {
      // We want to create sums for the section before this one, before creating the new
      // test section. But we can't do this for the very first section -- thus this check.
      if (scoreSheetColumn > pointCategories.length + TOTAL_COLUMN) {
        // Tick the column counter, add some headers, and then call a function that creates
        // sums for each point category.
        scoreSheetColumn++;
        SCORE_SHEET.getRange(2, scoreSheetColumn, 3, 1).setValues([["Totalt"], ["Max"], ["Medel"]]);
        buildSumColumns(sectionColumnStart + 1, scoreSheetColumn - 1, pointCategories);
        scoreSheetColumn = scoreSheetColumn + pointCategories.length + 1;
      }

      // Next, let's start a new section by adding a blank column + some headers. We make
      // the blank column narrow *after* the headers are created, to avoid the case where
      // we try to change the width of a column beyond the end of the spreadsheet.
      scoreSheetColumn++;
      SCORE_SHEET.getRange(2, scoreSheetColumn, 3, 1).setValues([[questionName], ["Max"], ["Medel"]]);
      SCORE_SHEET.getRange(1, scoreSheetColumn, LAST_STUDENT_ROW).setBackground(bgColors[row]);
      SCORE_SHEET.setColumnWidth(scoreSheetColumn - 1, 20);
      // The variable sectionColumnStart is used to keep track of where the current section
      // started, to be able to create sums at the end of the section.
      var sectionColumnStart = scoreSheetColumn;
    }
  }

  // After iterating through each row in the build info, we need to create sum columns
  // for the last section as well. (This doesn't have any new section that triggers the
  // sum procedures.)
  scoreSheetColumn++;
  SCORE_SHEET.getRange(2, scoreSheetColumn, 3, 1).setValues([["Totalt"], ["Max"], ["Medel"]]);
  buildSumColumns(sectionColumnStart + 1, scoreSheetColumn - 1, pointCategories);

  // Copy the row for the first student to as many rows as needed.
  if (numberOfStudents > 1) {
    var targetRange = SCORE_SHEET.getRange(FIRST_STUDENT_ROW + 1, 1, LAST_STUDENT_ROW - FIRST_STUDENT_ROW, SCORE_SHEET.getLastColumn());
    SCORE_SHEET.getRange(FIRST_STUDENT_ROW, 1, 1, SCORE_SHEET.getLastColumn()).copyTo(targetRange);
  }
  for (var row = FIRST_STUDENT_ROW; row <= LAST_STUDENT_ROW; row++) {
    SCORE_SHEET.getRange(row, 1).setValue("Elev " + parseInt(row - FIRST_STUDENT_ROW + 1));
  }

  // Prettify the spreadsheet by reducing the columns widths.
  for (var column = 1; column <= SCORE_SHEET.getLastColumn(); column++) {
    SCORE_SHEET.autoResizeColumn(column);
  }

  // Rebuild the menu, to hide the option to create a new scoring sheet.
  buildMenu();
}

/**
 * Builds columns containing point sums for a section of the test.
 *
 * This function creates summary columns for a section, summing up all points
 * for each category. It also updates the cells containing formulas for point
 * totals.
 */
function buildSumColumns(columnStart, columnEnd, categories) {
  // We need the two A1 column names.
  var startA1 = SCORE_SHEET.getRange(2, columnStart).getA1Notation().slice(0, -1);
  var endA1 = SCORE_SHEET.getRange(2, columnEnd).getA1Notation().slice(0, -1);

  // For each point category: Create a sum column, and add it to the cell with
  // the totals.
  for (var i in categories) {
    var scoreSheetColumn = TOTAL_COLUMN - 1 + parseInt(columnEnd) + parseInt(i);
    addAverageFormula(scoreSheetColumn);
    // Get the A1 name for the column we are creating. We need it for formulas.
    var currentA1 = SCORE_SHEET.getRange(2, scoreSheetColumn).getA1Notation().slice(0, -1);
    // Create sum column, using the SUMIF() spreadsheet function.
    SCORE_SHEET.getRange(2, scoreSheetColumn, 1).setValue(categories[i]);
    SCORE_SHEET.getRange(3, scoreSheetColumn, 1).setFormula("=sumif($" + startA1 + "$2:$" + endA1 + "$2;" + currentA1 + "$2;$" + startA1 + "3:$" + endA1 + "3)")
      .copyTo(SCORE_SHEET.getRange(FIRST_STUDENT_ROW, scoreSheetColumn, 1));
    // Append a reference to the new column in the point totals.
    var formula = SCORE_SHEET.getRange(3, TOTAL_COLUMN + 1 + parseInt(i), 1).getFormula();
    SCORE_SHEET.getRange(3, TOTAL_COLUMN + 1 + parseInt(i), 1).setFormula(formula + "+" + currentA1 + "3")
      .copyTo(SCORE_SHEET.getRange(FIRST_STUDENT_ROW, TOTAL_COLUMN + 1 + parseInt(i), 1));
  }
}
