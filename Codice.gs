function sendEmails1() {
   var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Alert");
var sheetDest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quotazioni");
  var NumRows=sheet.getRange("G1").getValue();
  var startRow = 1; // First row of data to process
 //var MaxR=sheet.getRange("1").getValue();
//  var numRows = 2; // Number of rows to process
  // Fetch the range of cells A2:B3
  if (NumRows>0 ) {
    var dataRange = sheet.getRange(startRow, 1, NumRows, 4);
  // Fetch values for each row in the Rang1e.
    var data = dataRange.getValues();
    var emailAddress = "silvio.cilloco@gmail.com";
  }  
  for (var i = 0; i < NumRows; ++i) {
    var row = data[i];
     var message = row[2]; // Second column
    var subject=row[2];
    Logger.log("oggetto "+subject+ " valore"+row[3]);
    MailApp.sendEmail(emailAddress, subject, message);
      sheetDest.getRange(row[0],row[1]).setValue(row[3]);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
   
  
}


function sendEmails() {
   var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Alert");
var sheetDest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quotazioni");
  var NumRows=sheet.getRange("G1").getValue();
 
 var emailAddress = "profmath1961@gmail.com";
  for (var i = 1; i <= NumRows; ++i) {
   
    var NumR=sheet.getRange(i,1).getValue();
       var NumC=sheet.getRange(i,2).getValue();
     var message=sheet.getRange(i,3).getValue();
    var subject=message;
    var NewVal = sheet.getRange(i,4).getValue();
 //   var subject=row[2];
   // Logger.log("oggetto "+subject+ " valore"+row[3]);
  //  MailApp.sendEmail(emailAddress, subject, message);
      sheetDest.getRange(NumR,NumC).setValue(NewVal);
      // Make sure the cell is updated right away in case the script is interrupted
   //   SpreadsheetApp.flush();
    }
   
  
}

