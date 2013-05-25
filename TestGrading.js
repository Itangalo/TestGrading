/**
 * Adds a custom menu to the active spreadsheet on opening the spreadsheet.
 */
function onOpen() {
  var menuEntries = [];
  menuEntries.push({name : "Bygg kalkylblad för poäng", functionName : "buildScoreSheet"});
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Poängmall", menuEntries);
};

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
 * Sets up the sheet for entering test scores.
 */
function buildScoreSheet() {
  if (!sheetExists("Poäng")) {
    var scoreSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Poäng");
    scoreSheet.setFrozenColumns(1);
    scoreSheet.setFrozenRows(4);
  }
  var scoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Poäng");
  
  var buildInfoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Maxpoäng");
  var buildInfo = buildInfoSheet.getSheetValues(1, 1, buildInfoSheet.getLastRow(), buildInfoSheet.getLastColumn());
  var bgColors = buildInfoSheet.getRange(1, 1, buildInfoSheet.getLastRow(), 1).getBackgrounds();
  
  var pointCategories = buildInfo.shift();
  bgColors.shift();
  pointCategories.shift();

  var scoreSheetColumn = 1;
  for (var row in buildInfo) {
    var questionName = buildInfo[row].shift();
    // If we don't have any points on a row, assume it is a new section of the test.
    if (!arrayHasValues(buildInfo[row])) {
      if (scoreSheetColumn > 1) {
        scoreSheetColumn++;
        scoreSheet.getRange(2, scoreSheetColumn, 3, 1).setValues([["Totalt"], ["Max"], ["Medel"]]);
        buildSumColumns(sectionColumnStart + 1, scoreSheetColumn - 1, pointCategories);
        scoreSheetColumn = scoreSheetColumn + pointCategories.length + 1;
      }
      scoreSheetColumn++;
      scoreSheet.getRange(2, scoreSheetColumn, 3, 1).setValues([[questionName], ["Max"], ["Medel"]]);
      scoreSheet.getRange(1, scoreSheetColumn, 40).setBackground(bgColors[row]);
      var sectionColumnStart = scoreSheetColumn;
    }
    // If we DO have points on a row, add column(s) for this question.
    else {
      var columnStart = scoreSheetColumn + 1;
      for (var i in buildInfo[row]) {
        if (buildInfo[row][i] !== "") {
          scoreSheetColumn++;
          scoreSheet.getRange(1, scoreSheetColumn, 3, 1).setValues([[questionName], [pointCategories[i]], [buildInfo[row][i]]]);
          scoreSheet.getRange(1, scoreSheetColumn, 40).setBackground(bgColors[row]);
        }
      }
      scoreSheet.getRange(1, columnStart, 1, scoreSheetColumn - columnStart + 1).merge().setHorizontalAlignment("center");
    }
  }

  // Build sum columns for the last section as well.  
  if (scoreSheetColumn > 1) {
    scoreSheetColumn++;
    scoreSheet.getRange(2, scoreSheetColumn, 3, 1).setValues([["Totalt"], ["Max"], ["Medel"]]);
    buildSumColumns(sectionColumnStart + 1, scoreSheetColumn - 1, pointCategories);
    scoreSheetColumn = scoreSheetColumn + pointCategories.length + 1;
  }

  
  for (var column = 1; column <= scoreSheet.getLastColumn(); column++) {
    scoreSheet.autoResizeColumn(column);
  }
}

/**
 * Builds columns containing point sums for a section of the test.
 */
function buildSumColumns(columnStart, columnEnd, categories) {
  var scoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Poäng");
  var startA1 = scoreSheet.getRange(2, columnStart).getA1Notation().slice(0, -1);
  var endA1 = scoreSheet.getRange(2, columnEnd).getA1Notation().slice(0, -1);
  for (var i in categories) {
    var currentA1 = scoreSheet.getRange(2, parseInt(columnEnd) + 2 + parseInt(i)).getA1Notation().slice(0, -1);
    scoreSheet.getRange(2, parseInt(columnEnd) + 2 + parseInt(i), 1).setValue(categories[i]);
    scoreSheet.getRange(3, parseInt(columnEnd) + 2 + parseInt(i), 1).setFormula("=sumif($" + startA1 + "$2:$" + endA1 + "$2;" + currentA1 + "$2;$" + startA1 + "3:$" + endA1 + "3)");
    scoreSheet.getRange(5, parseInt(columnEnd) + 2 + parseInt(i), 1).setFormula("=sumif($" + startA1 + "$2:$" + endA1 + "$2;" + currentA1 + "$2;$" + startA1 + "5:$" + endA1 + "5)");
  }
}
