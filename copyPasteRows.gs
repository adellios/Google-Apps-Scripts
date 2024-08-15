function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('MyFunction','myFunction')
      .addToUi();
}

function myFunction() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('MT Returned');
  const ui = SpreadsheetApp.getUi();
  const sheet2 = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1EXDZyzYA1AaRRzTu3Vn3UHGFTsOMM1gaSc-CHaSYOKA/edit#gid=1013228974');

  var result = ui.prompt('How many new entries?','Input:',ui.ButtonSet.OK_CANCEL);
  
  var button = result.getSelectedButton();
  var numLines = parseInt(result.getResponseText());

  if (numLines < 1){
    Browser.msgBox('Please enter a positive integer.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const lastRow1 = data.length;
  const data2 = sheet2.getDataRange().getValues();
  const lastRow2 = data2.length + 1;

  new_val = lastRow1 - numLines + 1;
  new_val2 = lastRow2 - numLines;
  variable1 = lastRow2 + numLines - 1;

  copy_list = ['C','D','E','G','I','J','K','L'];
  paste_list = ['G','F','K','L','S','M','P','R'];

  for (i=0; i<8; i++){
    var src1 = sheet.getRange(copy_list[i] + new_val + ':' + copy_list[i] + lastRow1);
    var target_range = sheet2.getRange(paste_list[i] + lastRow2 + ':' + paste_list[i] + variable1);
    target_range.setValues(src1.getValues());
  }

  var target_range = sheet2.getRange('C' + lastRow2 + ':C' + variable1);
  twod_array = [];

  for (i=0; i<numLines; i++){
    twod_array.push([Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy")]);
  }
  target_range.setValues(twod_array);

}
