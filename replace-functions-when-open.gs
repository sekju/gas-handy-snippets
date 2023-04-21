// Function replacing value of all cells with =GPT() formula from formulas into text during opening workbook
function replaceGPTFunctions() {
  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = sheet.getActiveSheet();
  
  // Get the dimensions of the sheet
  var numRows = activeSheet.getMaxRows();
  var numCols = activeSheet.getMaxColumns();
  
  // Get the formulas for the entire sheet
  var formulas = activeSheet.getRange(1, 1, numRows, numCols).getFormulas();
  
  // Get the values calculated by the formulas
  var values = activeSheet.getRange(1, 1, numRows, numCols).getValues();
  
  // Search for cells with GPT() function and replace them with the calculated value as text
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[0].length; j++) {
      if (formulas[i][j].indexOf("GPT(") !== -1) {
        activeSheet.getRange(i + 1, j + 1).setValue(values[i][j]);
      }
    }
  }
}
