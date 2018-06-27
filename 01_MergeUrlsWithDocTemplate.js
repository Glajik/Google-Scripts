// Script is merge list of URLs (in Google Sheet) with 
// temlate (Google Doc)
//
// myFunction is entry point of script.
// It open SpreadSheet by URL, 
// get values from first Sheet into array,
// and call function createNewDoc each element.
function myFunction() {
  
  // Set URL of reference SpreadSheet
  const reference = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1nJuf_t3VGQ4ldfG0sQt66hcaGiLYyAp8qM8aY94XtPE/edit');
  
  // Set URL of teplate Document
  const template = DocumentApp.openByUrl('https://docs.google.com/document/d/1z7snJEK_KSCnI8NUx2qJvE399nHRgO6Z0kgCPND548k/edit');
  
  // Set custom folder name for output
  const customFolderName = 'Writer Template {Individual product dog food}';
  
  // If custom folder not exist, create folder
  const customFolder = makeFolderAtRoot(customFolderName);
  
  // get array of values in SpreadSheet
  const sheet = reference.getSheets()[0];
  var lastRow = sheet.getLastRow();
  const values = sheet.getSheetValues(1, 1, lastRow,1);

  const newValues = values.reduce(createNewDoc, new Array());
  
  // Save links to ShpreadSheet
  const newRange = sheet.getRange(1, 1, lastRow, 2);
  newRange.setValues(newValues);
  
  
  // createNewDoc is callback function.
  // It call for each element in values array,
  // create new document from template file,
  // and replace {VALUE} in the text to 
  // string from the each element of array.  
  function createNewDoc(acc, element, index) {
    const valueToPaste = element[0];

    // Name template of output Documents
    const num = index + 1;
    const newDocumentName = 'Task №' + num + ' for copywriter {' + valueToPaste + '}';

    // Create new doc
    const newDocument = DocumentApp.create(newDocumentName);
    const file = DriveApp.getFileById(newDocument.getId());
    customFolder.addFile(file);
    
    //Add text from template
    const text = template.getBody().getText();
    newDocument.getBody().setText(text);
    
    // Replace {VALUE}
    newDocument.getBody().replaceText('{VALUE}', valueToPaste);
    
    // add to array link on new document
    const newDocumentUrl = newDocument.getUrl();
    acc.push([valueToPaste, newDocumentUrl]);
    return acc; 
  }

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
  }
  
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
  }
};



