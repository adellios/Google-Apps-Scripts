function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Cognos SQL','saveData')
      .addItem('Test','myFunction')
      .addItem('Archive', 'makeCopy')
      .addToUi();
}




 
function saveData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var copySheet = ss.getSheetByName("Copy");
  var pasteSheet = ss.getSheetByName("Paste");


  // get source range
  var source = copySheet.getRange(1,2,pasteSheet.getLastRow(),2);
  // get destination range
  var destination = pasteSheet.getRange(1,2,12,2);

  // copy values to destination range
  source.copyTo(destination);

  var count = 0;

  for (var counter = 0; counter <= pasteSheet.getLastRow(); counter = counter + 1) {
    if("Paste!B" + counter == "DRYU9500946"){
      count = 1 + count;
    }
  }
  Browser.msgBox(count);
  Browser.msgBox(SpreadsheetApp.getActive().getRange("Paste!B1:B1" + pasteSheet.getLastRow()).getValues());
  Browser.msgBox("Paste!B2:B3");

  // clear source values
 // Browser.msgBox("Paste!A1");
  Browser.msgBox( SpreadsheetApp.getActive().getRange("Paste!D1:D" + pasteSheet.getLastRow()).getValues());
  //Browser.msgBox(pasteSheet.getLastRow());
 // source.clearContent();
  //pasteSheet.getLastRow()
}




function makeCopy() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = "";
  const ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Archive Week?', 'Tab name:', ui.ButtonSet.YES_NO);

  if (response.getSelectedButton() == ui.Button.YES) {
    if (ss.getSheetByName(response.getResponseText()) == null){
      Browser.msgBox('Tab not found.');
      } else {
        var formattedDate = Utilities.formatDate(new Date(), "GMT", "(MM-dd-yyyy)");
        var saveAs = response.getResponseText() + ' ' + formattedDate;
        var destinationFolder = DriveApp.getFolderById("10UTNNmU0AmO5JjQnY_zZ8i3MTiOR7uqW");
        DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).makeCopy(saveAs, destinationFolder);

        ss.setActiveSheet(ss.getSheetByName(response.getResponseText()));//Grabs tab to be deleted
        ss.deleteActiveSheet();
    
        Browser.msgBox(response.getResponseText() + ' was archived.');
      }
  } else if (response.getSelectedButton() == ui.Button.NO){
    Browser.msgBox('Tab was not archived.');
  }
}
