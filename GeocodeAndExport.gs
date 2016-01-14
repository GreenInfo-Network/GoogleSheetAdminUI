/**
 * The onOpen() function, when defined, is automatically invoked whenever the spreadsheet is opened.
 * Use this to start creating the cuswtom menus.
 *
 * For more information on using the Spreadsheet API, see https://developers.google.com/apps-script/service_spreadsheet
 *
 * The programming is broken into sections and families such as Geocode_XXX and ExportJSON_XXX
 * As far as I can tell, the API doesn't support placing functions inside larger objects, e.g. Geocode = { Help:function( {...} };
 * so this model of grouping seems to be how we need to go.
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Geocode', [
    { name: "Geocode Address to Lat and Lng", functionName: "Geocode_geocodeTrio" },
    { name: "Help", functionName: "Geocode_Help" }
  ]);

  SpreadsheetApp.getActiveSpreadsheet().addMenu("Export JSON", [
    { name: "Export as JSON", functionName: "ExportJSON_exportSheet"},
    { name: "Help", functionName: "ExportJSON_Help" }
  ]);
};

/*
 *
 * The Geocode functions
 *
 */

var GEOCODE_REGION = 'us';

function Geocode_Help () {
    var html = HtmlService.createHtmlOutputFromFile('Geocode_Help').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Help :: Geocoding Addresses');
}
function Geocode_geocodeTrio () {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();
  
  // Must have selected 3 columns (Address, Lat, Lng).
  if (cells.getNumColumns() != 3) return SpreadsheetApp.getUi().alert("Select 3 columns: Address, Lat, Lng");

  var addressColumn = 1;
  var addressRow;
  
  var latColumn = addressColumn + 1;
  var lngColumn = addressColumn + 2;
  
  var geocoder = Maps.newGeocoder().setRegion(GEOCODE_REGION);
  var location;
  
  for (addressRow = 1; addressRow <= cells.getNumRows(); ++addressRow) {
    var address = cells.getCell(addressRow, addressColumn).getValue();
    
    // Geocode the address and plug the lat, lng pair into the 
    // 2nd and 3rd elements of the current range row.
    location = geocoder.geocode(address);
   
    // Only change cells if geocoder seems to have gotten a 
    // valid response.
    if (location.status == 'OK') {
      lat = location["results"][0]["geometry"]["location"]["lat"];
      lng = location["results"][0]["geometry"]["location"]["lng"];
      
      cells.getCell(addressRow, latColumn).setValue(lat);
      cells.getCell(addressRow, lngColumn).setValue(lng);
    }
  }
}

/*
 *
 * The ExportJSON functions
 *
 */

function ExportJSON_Help () {
    var html = HtmlService.createHtmlOutputFromFile('ExportJSON_Help').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Help :: Exporting JSON to the Web Map');
}
function ExportJSON_exportSheet () {
  var sheet    = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowsData = ExportJSON_getRowsData(sheet);
  var json     = JSON.stringify(rowsData, null, 4);

  // templating! assign data, then exchange the template for a HtmlOutput suitable for use in a modal
  var html = HtmlService.createTemplateFromFile('ExportJSON_Output');
  html.data = { json:json };
  html = html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Paste This JSON Data into the Map Admin UI');  
}

function ExportJSON_getRowsData(sheet) {
  var headersRange = sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getMaxColumns());
  var headers      = headersRange.getValues()[0];
  var dataRange    = sheet.getRange(sheet.getFrozenRows()+1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  var dataRows     = ExportJSON_getObjects(dataRange.getValues(), ExportJSON_normalizeHeaders(headers));
  return dataRows;
}

function ExportJSON_getObjects(data, keys) {
  // loop over rows, collecting assoc objects
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false; // make sure we have at least one non-empty cell in this row

    // iterate over headers/keys to load the object but skip the pesky "empty" columkns that spreadsheets love to add
    for (var j = 0; j < data[i].length; ++j) {
      var key = keys[j];
      if (! key || key == 'undefined') continue;

      var cellData = data[i][j];
      if (ExportJSON_isCellEmpty(cellData)) cellData = null;
      else hasData = true;

      object[key] = cellData;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

function ExportJSON_normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = ExportJSON_normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

function ExportJSON_normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!ExportJSON_isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && ExportJSON_isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

function ExportJSON_isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}
function ExportJSON_isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    ExportJSON_isDigit(char);
}
function ExportJSON_isDigit(char) {
  return char >= '0' && char <= '9';
}
