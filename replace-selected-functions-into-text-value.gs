// Function replacing value of cells from formulas into text in selected cells 
function replaceFunctions() {
  // Download the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = sheet.getActiveSheet();
  
  // Get the selected cells
  var selection = activeSheet.getActiveRange();
  
  // Get the formula of the selected cells
  var formulas = selection.getFormulas();
  
  // Get the values calculated by the formulas
  var values = selection.getValues();
  
  // Replace formulas with text of calculated values
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[0].length; j++) {
      if (formulas[i][j]) {
        activeSheet.getRange(selection.getRow() + i, selection.getColumn() + j)
        .setValue(values[i][j]);
      }
    }
  }
}
