/*
* Script generate tasks for copywriters from template (Dog food)
* and using values from table for replace special tags.
* 
* myFunction is entry point of script.
* It open SpreadSheet by URL, 
* get values from first Sheet into array,
* and call function createNewDoc each element.
*/
function myFunction() {

  // Get document properties from active spreadsheet
  const properties = PropertiesService.getDocumentProperties();
  
  // reference SpreadSheet is Active Spreadsheet
  const reference = SpreadsheetApp.getActive();
  
  // Set URL of template Document
  const docURL = properties.getProperty('template');
  if (!docURL) {
    SpreadsheetApp.getUi().alert('Template is not set. Please set it first, and try run script again.');
    setTemplate();
    return 0;
  }
  const template = DocumentApp.openByUrl(docURL);
  
  // Set custom folder name for output
  const customFolderName = properties.getProperty('customFolderName');
  if (!customFolderName) {
    SpreadsheetApp.getUi().alert('Output folder is not set. Please set it first, and try run script again.');
    setOutput();
    return 0;
  }
  
  // If custom folder not exist, create folder
  const customFolder = makeFolderAtRoot(customFolderName);

/// MAIN ALGORITHM
    
  // get array of values in SpreadSheet
  const sheet = reference.getSheets()[0];
  var lastRow = sheet.getLastRow();
  const values = sheet.getSheetValues(2, 2, lastRow - 1, 17);

  const newValues = values.reduce(createNewDoc, new Array());

  // Save links to ShpreadSheet
  const newRange = sheet.getRange(2, 3, lastRow - 1, 1);
  newRange.setValues(newValues);
    
  // createNewDoc is callback function.
  // It call for each element in values array,
  // create new document from template file,
  // and replace {VALUE} in the text to 
  // string from the each element of array.  
  function createNewDoc(acc, element, index) {
    const SerialNumber = element[0];
    const ProductName = element[1];
    const ProductNameWithoutSpaces = removeSpaces(ProductName);
    const WhyIsThisProductHere = element[2];
    const Proteins = element[3];
    const Fats = element[4];
    const Carbs = element[5];
    const VitaminsMinerals = element[6];
    const PreservativesBadStuff = element[7];
    const CheckCarefully = element[8];
    const MustKnowFacts = element[9];
    const watchOutFor = element[10];
    const CrucialTips = '<ol>' + tagLI(element[11]) + '<ol>';
    const Pros = tagLI(element[12]);
    const Cons = tagLI(element[13]);
    const Conclusion = element[14];
    const ProductURL = element[15];
    const LinkToAmazon = addNofollow(element[16]);

    // Name template of output Documents
    const num = index + 1;
    const newDocumentName = 'Task №' + num + ' (S/N: ' + SerialNumber + ', ProductName: ' + ProductName + ' )';

    // Create new doc
    const newDocument = DocumentApp.create(newDocumentName);
    const file = DriveApp.getFileById(newDocument.getId());
    customFolder.addFile(file);
    
    //Add text from template
    const text = template.getBody().getText();
    newDocument.getBody().setText(text);
    
    // Replace {VALUE}
    newDocument.getBody().replaceText('{SerialNumber}', SerialNumber);
    newDocument.getBody().replaceText('{ProductName}', ProductName);
    newDocument.getBody().replaceText('{ProductNameWithoutSpaces}', ProductNameWithoutSpaces);
    newDocument.getBody().replaceText('{WhyIsThisProductHere}', WhyIsThisProductHere);
    newDocument.getBody().replaceText('{Proteins}', Proteins);
    newDocument.getBody().replaceText('{Fats}', Fats);
    newDocument.getBody().replaceText('{Carbs}', Carbs);
    newDocument.getBody().replaceText('{VitaminsMinerals}', VitaminsMinerals);
    newDocument.getBody().replaceText('{PreservativesBadStuff}', PreservativesBadStuff);
    newDocument.getBody().replaceText('{CheckCarefully}', CheckCarefully);
    newDocument.getBody().replaceText('{MustKnowFacts}', MustKnowFacts);
    newDocument.getBody().replaceText('{watchOutFor}', watchOutFor);
    newDocument.getBody().replaceText('{CrucialTips}', CrucialTips);
    newDocument.getBody().replaceText('{Pros}', Pros);
    newDocument.getBody().replaceText('{Cons}', Cons);
    newDocument.getBody().replaceText('{Conclusion}', Conclusion);
    newDocument.getBody().replaceText('{ProductURL}', ProductURL);
    newDocument.getBody().replaceText('{LinkToAmazon}', LinkToAmazon);
    
    
  
    // add to array link on new document
    const newDocumentUrl = newDocument.getUrl();
    const link = '=HYPERLINK("' + newDocumentUrl + '","' + ProductName + '")';
    acc.push([link]);
    return acc;
  }

  /// ADDITIONAL FUNCTIONS
  // function return string without space characters
  function removeSpaces(text) {
    return text.replace(/ /g, '');
  }
  
  // function make HTML list from multiline text
  function tagLI(text) {
    var result = '\n<li>';
    result += text.replace(/\n/g, '</li>\n<li>');
    result += '</li>\n';
    return result.replace(/<li> *<\/li>\n/g, '');
  }
  
  // adding nofollow attribute
  function addNofollow(text) {
    return text.replace(/target="_blank"/g, 'target="_blank" rel="nofollow"');
  }
};

// make Folder at root of Google Drive
function makeFolderAtRoot(name) {
  const isFolder = findMyFolder(name);
  if (!isFolder) {
    const root = DriveApp.getRootFolder();
    const newFolder = DriveApp.createFolder(name);
    root.addFolder(newFolder);
    return newFolder;
  } else { 
    return isFolder;
  }
};

// function for find folder and check if it in the root of Google Drive
// return ID of folder
// return false if not find
function findMyFolder(name) {
  const folders = DriveApp.getFolders();
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() === name) {
      parents = folder.getParents();
      while(parents.hasNext()) {
        var parent = parents.next();
        if (parent.getId() === DriveApp.getRootFolder().getId()) return folder;
      }
    }
  }
  return false;
};

// Add simple menu in Addons to run script
function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('Dog Food') // Or DocumentApp.
  .addItem('Run script', 'myFunction')  
  .addItem('Set template', 'setTemplate')
  .addItem('Set output folder', 'setOutput')
  .addToUi();
};

// Set Url to template and store into Spreadsheet property
function setTemplate() {
  
  // Access to user interface
  const ui = SpreadsheetApp.getUi(); 
  
  // Access to properies of spreadsheet
  const properties = PropertiesService.getDocumentProperties(); 
  
  // Check if property is set, and allow to change it
  const property = properties.getProperty('template');
  if (property) {
    message = 'Template URL: ' + property;
  } else {
    message = 'Template URL is not set';
  }
  
  // show message
  const result = ui.prompt(message, 'New URL:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.CANCEL) {
    return 0;
  }
  
  // Check URL
  try {
    const url = result.getResponseText();
    DocumentApp.openByUrl(url);
  }
  catch (e) {
    ui.alert('Can\'t access to file. Check URL and try again. (' + e + ')');
    return 0;
  }
  
  // Save URL
  properties.setProperty('template', result.getResponseText());
};

// Set folder name for output and store it
function setOutput() {
  // Access to user interface
  const ui = SpreadsheetApp.getUi(); 
  
  // Access to properies of spreadsheet
  const properties = PropertiesService.getDocumentProperties();   
  
  // Check if property is set, and allow to change it
  const property = properties.getProperty('customFolderName');
  if (property) {
    message = 'Output is set to: ' + property;
  } else {
    message = 'Output folder is not set';
  }
  
  // show message
  const result = ui.prompt(message, 'New folder name:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.CANCEL) {
    return 0;
  }
  
  // check folder name
  try {
    const folderName = result.getResponseText();
    if (folderName === '') {
      ui.alert('Folder name must contain one or more symbols');
      return 0;
    }
    makeFolderAtRoot(folderName);
  }
  catch (e) {
    ui.alert('Can\'t make folder. Check name and try again. (' + e + ')');
    return 0;
  }
  
  properties.setProperty('customFolderName', result.getResponseText());
};
