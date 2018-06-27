
//myFunction
/*
Read url from column 2 and
- format URL to anchor with NoFollow _Blank attributes (column 3)
- create Shortcodes Ultimate button (column 4)
*/
function myFunction() {
  // reference SpreadSheet is Active Spreadsheet
  const reference = SpreadsheetApp.getActive();
  
  // get array of values in SpreadSheet
  const sheet = reference.getSheets()[0];
  var lastRow = sheet.getLastRow();
  const input = sheet.getSheetValues(1, 1, lastRow, 2);

  const output = input.reduce(modify, new Array());

  // Save links to ShpreadSheet
  const myRange = sheet.getRange(1, 1, lastRow, 4);
  myRange.setValues(output);

  // modify is callback function for reduce.
  // It call for each element in values array,
  function modify(acc, element) {
    const productName = element[0];
    const url = element[1];
    const anchor = '<a href="' + url + '" target="_blank" rel="nofollow">' + productName + '</a>';
    const button = '[su_button url="' + url + '" target="blank" background="#FF5722" size="6" center="no" icon_color="#c81616" rel="nofollow" class="redButtonInTable"] Check Price[/su_button]';
    acc.push([productName, url, anchor, button]);
    return acc; 
  }
}
// Add simple menu in Addons to run script
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu() // Or DocumentApp.
  .addItem('Format URL', 'myFunction')
  .addToUi();
}