// custom menu function
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Add Rows','addRows')
      .addToUi();
}
 
// function to add rows
function addRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const ui = SpreadsheetApp.getUi();

  // first prompt (select location)
  var result = ui.prompt(
      'Select first row after target',
      'Input:',
      ui.ButtonSet.OK_CANCEL);

  // prompt yard code and find last instance

  // second prompt (select # of rows)
  var result2 = ui.prompt(
      'Select # of rows to add',
      'Input:',
      ui.ButtonSet.OK_CANCEL);

  // retrieve answers and convert to integers
  var button = result.getSelectedButton();
  var targetRow = parseInt(result.getResponseText());
  var button2 = result2.getSelectedButton();
  var numRows = parseInt(result2.getResponseText());

  // identify row to copy down values from 
  var copyRow = targetRow - 1;

  // check if row number and quantity are provided
  if (button == ui.Button.OK & button2 == ui.Button.OK) {

    // insert blank rows
    sheet.insertRows(targetRow, numRows);
    
    // for each new (blank) row ...
    for (var counter = targetRow; counter <= targetRow + numRows - 1; counter = counter + 1) {

      // copy down formatting for columns A through U
      var src = sheet.getRange('A' + copyRow + ':U' + copyRow);
      var target = sheet.getRange('A' + counter + ':U' + counter);
      src.copyTo(target, {formatOnly:true})

      // copy down formula in column A
      var src1 = sheet.getRange('A' + copyRow);
      var target1 = sheet.getRange('A' + counter);
      target1.setValues(src1.getFormulasR1C1());

      // copy down values in column B
      var src2 = sheet.getRange('B' + copyRow);
      var target2 = sheet.getRange('B' + counter);
      target2.setValues(src2.getValues());

      // copy down formulas in columns C through F
      var src3 = sheet.getRange('C' + copyRow + ':F' + copyRow);
      var target3 = sheet.getRange('C' + counter + ':F' + counter);
      target3.setValues(src3.getFormulasR1C1());

      // copy down values in column H
      var src4 = sheet.getRange('H' + copyRow);
      var target4 = sheet.getRange('H' + counter);
      target4.setValues(src4.getValues());

      // copy down formula in column I
      var src5 = sheet.getRange('I' + copyRow);
      var target5 = sheet.getRange('I' + counter);
      target5.setValues(src5.getFormulasR1C1());

      // copy down values in column J
      var src6 = sheet.getRange('J' + copyRow);
      var target6 = sheet.getRange('J' + counter);
      target6.setValues(src6.getValues());
    }

  }

}
