/*
 * Login GOOGLE SCRIPT
 * ***********************************************************************************
 * Code by : khairul|basar
 * contact me : khairul.basar@gmail.com
 * Date : July 2021
 * License :  
 * 
 */
 
//spreadsheet url
var database = "https://docs.google.com/spreadsheets/d/1xnq6kMaQSsK7imGGgzOt_x4tYB5GFSlMpOcuwyEoGjI/edit#gid=0";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Send WhatsApp', 'whatsApp')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Future item', 'menuItem2'))
      .addToUi();
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the future item!');
}

/**
 * Get "home/Index page", or a requested page.
 * Expects a 'page' parameter in querystring.
 *
 * @param {event} e Event passed to doGet, with querystring
 * @returns {String/html} Html to be served
 */
function doGet(e){
  Logger.log( Utilities.jsonStringify(e) );
  if (!e.parameter.page){
    // When no specific page requested, return "home/Index page"
    var htmlOutput = HtmlService
    .createTemplateFromFile('Index')

    htmlOutput.message = '';
    htmlOutput.display_error = '';

    return htmlOutput.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  }
  else{
    var htmlOutput = HtmlService
    .createTemplateFromFile(e.parameter['page'])
   
    htmlOutput.message = '';
    htmlOutput.display_error = '';

    return htmlOutput.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
 var url = ScriptApp.getService().getUrl();
 return url;
}

function getPageURL(qs) {

  var url = ScriptApp.getService().getUrl();
  Logger.log(url + qs);
  return url + qs ;
}

function submitAbsensi(data){
  // Buka Spreadsheet
  let ss  = SpreadsheetApp.openByUrl(database);
  let ws  = ss.getSheetByName('UserPassword');
  let list    = ws.getRange(2,1,ws.getRange("A2").getDataRegion().getLastRow() - 1, 5).getValues();
  let userId  = list.map(function(r){ return r[1].toString(); });
  let pass    = list.map(function(r){ return r[3].toString(); });

  // Verifikasi data yang dikirim
  data.idPegawai  = data['idPegawai'] !== undefined?data.idPegawai.toString():'';
  data.password   = data['password'] !== undefined?data.password.toString():'';
  data.position   = data['position'] !== undefined?data.position:[0,0];

  // Verifikasi user dan password
  let indexData = userId.indexOf(data.idPegawai);
  if(indexData > -1 && pass[indexData] == data.password){
    let nama      = list.map(function(r){ return r[2]});

    // Get Alamat dari koordinat
    let koordinat = data.position[0] +', '+ data.position[1]; 
    let response  = Maps.newGeocoder().reverseGeocode(data.position[0], data.position[1]);
    let lokasi    = response.results[0].formatted_address;

    //buka sheet absensi dan simpan data
    ws = ss.getSheetByName("Location");
    ws.appendRow([data.idPegawai, nama[indexData], new Date(),koordinat, lokasi]);
    return nama[indexData];
  }
  else 
    return false;
}

function submitCEOMsg(number=123,textarea="abc"){
 //URL OF GOOGLE SHEET;
 var ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19uu3W5ZC1pnwXB7WGtEg05i6n-8V85PR4FPsc4x2O10/edit#gid=1158463931");
 var dataSheet = ss.getSheetByName("Message");
 dataSheet.appendRow([new Date(), number, textarea]);
}

function whatsApp() {
  var wsheet = SpreadsheetApp.openByUrl(database);
  var sheet = wsheet.getSheetByName("WhataApp");
  var startRow = 2;  // First row of data to process - 2 exempts my header row
  var numRows = sheet.getLastRow();   // Number of rows to process
  var numColumns = sheet.getLastColumn();
  var dataRange = sheet.getRange(startRow, 1, numRows-1, numColumns); 
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var fName =     row[0]; 
    var lName =     row[1];  
    var mobileNo =  row[2]; 
    var mess =      row[4];
    Logger.log(mess);   
    var my_apikey = "c210e7758877fa3b56094fcfdae537a7";
    var destination = mobileNo; 
    var message = mess; 
    var api_url = "http://wapi.targetad.in/api/v2/send"; 
    api_url += "WAMessage?token="+ my_apikey; 
    api_url += "&to="+ destination; 
    api_url += "&type=text&text="+ encodeURIComponent (message); 
    var response = UrlFetchApp.fetch(api_url);
    Logger.log(response.getContentText());   
    Logger.log(data.flag);
    Logger.log(data.message);
   
  }
}