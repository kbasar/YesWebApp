var SS= SpreadsheetApp.openByUrl(database);

function doPost(e) {

    Logger.log(JSON.stringify(e));
    var number = e.parameters.mobile;
    var textarea = e.parameters.FormControlTextarea1;
    
    if(number != '' && textarea != ''){
        AddRecordMsg(number,textarea);

        // var sheet = SpreadsheetApp.openById('1xnq6kMaQSsK7imGGgzOt_x4tYB5GFSlMpOcuwyEoGjI').getSheetByName('Message');

        // var numRows = sheet.getLastRow()+1; 

        // sheet.getRange("A"+numRows).setValue(new Date());
        // sheet.getRange("B"+numRows).setValue(e.parameter.mobile);
        // sheet.getRange("C"+numRows).setValue(e.parameter.FormControlTextarea1);

        
        var htmlOutput = HtmlService
        .createTemplateFromFile(e.parameter['page'])

        htmlOutput.message = 'Thank you. We have received your message';
        htmlOutput.display_error = '';
        var html = htmlOutput.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        html.addMetaTag("viewport", "width=device-width, initial-scale=1.0");
        return html;
         }
         else{
          var htmlOutput =  HtmlService.createTemplateFromFile(e.parameter['page']);
          htmlOutput.message = '';
          htmlOutput.display_error = 'Please Enter All Information!';
          var html = htmlOutput.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
          html.addMetaTag("viewport", "width=device-width, initial-scale=1.0");
          return html; 
         }


}

function getFirstEmptyRow() {
    var ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19uu3W5ZC1pnwXB7WGtEg05i6n-8V85PR4FPsc4x2O10/edit#gid=151695352");
    var dataSheet = ss.getSheetByName("Message");
    var column = dataSheet.getRange('A:A');
    var FirstEmptyRow = dataSheet.getRange('A:A').getLastRow()+1;
    var values = column.getValues(); // get all data in one call
    var ct = 0;
    while ( values[ct][0] != "" ) {
      ct++;
    }
    return ct+1;
  }

function AddRecordMsg(formData) { 
    //URL OF GOOGLE SHEET;
    var ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19uu3W5ZC1pnwXB7WGtEg05i6n-8V85PR4FPsc4x2O10/edit#gid=151695352");
    var dataSheet = ss.getSheetByName("Form Responses");
    // var newDate = new Date(); // Get new date object
    var timezone = Session.getScriptTimeZone(); // See Apps Script documentation or like "GMT+5"
    var format = "MM-dd-yyyy HH:mm:ss"; // Set date format for output
    var timeStamp = Utilities.formatDate(new Date(), timezone, format); // This returns a string
    // dataSheet.appendRow([new Date(), '','','','Message to the CEO',formData.FormControlTextarea1]);
    var numRows = getFirstEmptyRow(); 
    // var numRows = getColumnHeight(1,dataSheet,ss);
    dataSheet.getRange("A"+numRows).setValue(timeStamp);
    dataSheet.getRange("E"+numRows).setValue("Message to the CEO");
    dataSheet.getRange("F"+numRows).setValue(formData.FormControlTextarea1);
    return true;
  }

  function AddServiceRequest(formData) { 
    //URL OF GOOGLE SHEET;
    var ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19uu3W5ZC1pnwXB7WGtEg05i6n-8V85PR4FPsc4x2O10/edit#gid=151695352");
    var dataSheet = ss.getSheetByName("Form Responses");
    // var newDate = new Date(); // Get new date object
    var timezone = Session.getScriptTimeZone(); // See Apps Script documentation or like "GMT+5"
    var format = "MM-dd-yyyy HH:mm:ss"; // Set date format for output
    var timeStamp = Utilities.formatDate(new Date(), timezone, format); // This returns a string
    // dataSheet.appendRow([new Date(), '','','','Message to the CEO',formData.FormControlTextarea1]);
    var numRows = getFirstEmptyRow(); 
    // var numRows = getColumnHeight(1,dataSheet,ss);
    //A,U,Q,H,G,R
    dataSheet.getRange("A"+numRows).setValue(timeStamp);
    dataSheet.getRange("E"+numRows).setValue("Request for service");
    dataSheet.getRange("U"+numRows).setValue(formData.COMPANY);
    dataSheet.getRange("Q"+numRows).setValue(formData.MACHINETYPE);
    dataSheet.getRange("H"+numRows).setValue(formData.MACHINENO);
    dataSheet.getRange("G"+numRows).setValue(formData.PROBLEM);
    dataSheet.getRange("R"+numRows).setValue(formData.COMMENT);
    return true;
  }

  function AddOrderInks(formData) { 
    //URL OF GOOGLE SHEET;
    var ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19uu3W5ZC1pnwXB7WGtEg05i6n-8V85PR4FPsc4x2O10/edit#gid=151695352");
    var dataSheet = ss.getSheetByName("Form Responses");
    // var newDate = new Date(); // Get new date object
    var timezone = Session.getScriptTimeZone(); // See Apps Script documentation or like "GMT+5"
    var format = "MM-dd-yyyy HH:mm:ss"; // Set date format for output
    var timeStamp = Utilities.formatDate(new Date(), timezone, format); // This returns a string
    // dataSheet.appendRow([new Date(), '','','','Message to the CEO',formData.FormControlTextarea1]);
    var numRows = getFirstEmptyRow(); 
    // var numRows = getColumnHeight(1,dataSheet,ss);
    //A,U,Q,H,G,R
    dataSheet.getRange("A"+numRows).setValue(timeStamp);
    dataSheet.getRange("E"+numRows).setValue("Order Inks");
    dataSheet.getRange("I"+numRows).setValue(formData.TYPE);
    dataSheet.getRange("J"+numRows).setValue(formData.CYAN);
    dataSheet.getRange("K"+numRows).setValue(formData.MAGENTA);
    dataSheet.getRange("L"+numRows).setValue(formData.YELLOW);
    dataSheet.getRange("M"+numRows).setValue(formData.BLACK);
    dataSheet.getRange("N"+numRows).setValue(formData.CLEANSOL);
    dataSheet.getRange("O"+numRows).setValue(formData.CLEANWAT);
    return true;
  }


  function sendGmail() {
    var now = new Date();
    GmailApp.sendEmail("khairul.basar@gmail.com", "Time Message tester", "The time is: " + now.toString());
}

function logConsole(item=1){
    var ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19uu3W5ZC1pnwXB7WGtEg05i6n-8V85PR4FPsc4x2O10/edit#gid=151695352");
    var dataSheet = ss.getSheetByName("Form Responses");
    Logger.log(item);
    var v= getFirstEmptyRow();
    Logger.log("first empty row= "+v);
}

function getFirstEmptyRow() {
    var ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19uu3W5ZC1pnwXB7WGtEg05i6n-8V85PR4FPsc4x2O10/edit#gid=151695352");
    var dataSheet = ss.getSheetByName("Form Responses");
    var column = dataSheet.getRange('A:A');
    var FirstEmptyRow = dataSheet.getRange('A:A').getLastRow()+1;
    var values = column.getValues(); // get all data in one call
    var ct = 0;
    while (!values[ct][0].toString().length == 0 || ct==1) {
      ct++;
    }
    
    return (ct+1);
  }

  function getColumnHeight(col=1, sh, ss) {
    var ss = ss || SpreadsheetApp.getActive();
    var sh = sh || ss.getActiveSheet();
    var col = col || sh.getActiveCell().getColumn();
    var rcA = [];
    if (sh.getLastRow()){ rcA = sh.getRange(1, col, sh.getLastRow(), 1).getValues().flat().reverse(); }
    let s = 0;
    for (let i = 0; i < rcA.length; i++) {
      if (rcA[i].toString().length == 0) {
        s++;
      } else {
        break;
      }
    }
    return rcA.length - s;
    //const h = Utilities.formatString('col: %s len: %s', col, rcA.length - s);
    //Logger.log(h);
    //SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(h).setWidth(150).setHeight(100), 'Col Length')
  }