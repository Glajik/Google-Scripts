/* Filter list of URLs by domain.
 * Script is need for work four sheets with 
 * names "input", "output" and "reference".
 *
 * "input" - sheet for input list of URLs
 *
 * "output" - sheet for result
 *
 * "Filter Out" - sheet for URLs which were discarded
 *
 * "reference" - sheet with domain name list, 
 *      where every row is item of list,
 *      and domain name start from dot: ".es", ".ru", ".nl"
 */ 
function urlFilter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var extensionsSheet = ss.getSheetByName('reference');
  var inputSheet = ss.getSheetByName('input');
  
  var extensions = getValuesFrom(extensionsSheet);
  var urls = getValuesFrom(inputSheet);
 
  var searchString = extensions.map(
    function(ext) {
      return '\\' + ext;
    })
  .join('|');
  
  var regexp = new RegExp('(' + searchString + ')(?=[/.]|$)', '');
  
  var filtered = urls.filter(
    function(u) {
      return !regexp.test(u);
    });
  
  var filterOut = urls.filter(
    function(u) {
      return regexp.test(u);
    }); 
  
  var filterOutSheet = ss.getSheetByName('Filter Out');
  filterOutSheet.clear();
  setValuesTo(filterOutSheet, filterOut);
  
  var outputSheet = ss.getSheetByName('output');
  outputSheet.clear();
  setValuesTo(outputSheet, filtered);
  outputSheet.activate();
}

function getValuesFrom(sheet) {
  var numRows = sheet.getLastRow();
  var inputValues = sheet.getRange(1, 1, numRows, 1).getValues(); // [[wer@erw.e], [wer@rf.rr]] 
  return inputValues.map(
    function(v) {
      return v[0];
    });
}

function setValuesTo(sheet, values) {
  var result = values.map(
    function(v) {
      return [v];
    });
  
  sheet.getRange(1, 1, values.length, 1).setValues(result);
}

function onOpen(e) {
const ui = SpreadsheetApp.getUi();
ui.createMenu('Url Filter')
.addItem('Run', 'urlFilter')
.addToUi();
}