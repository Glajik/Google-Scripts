/*
 * Search sent messages in Gmail.
 * Filter it by emails (domain or some text too) in column A and retrieve
 * subjects, bodies and dates of last email in lists,
 * put info to column B, C, and D
 */
 
function myFunction() {
  function _toConsumableArray(arr) { if (Array.isArray(arr)) { for (var i = 0, arr2 = Array(arr.length); i < arr.length; i++) { arr2[i] = arr[i]; } return arr2; } else { return Array.from(arr); } };
  
  var ss = SpreadsheetApp.getActiveSheet();
  // get email list
  var numRows = ss.getDataRange().getLastRow();
  var inputValues = ss.getRange(1, 1, numRows, 1).getValues(); // [[wer@erw.e], [wer@rf.rr]]
  var output = ss.getRange(1, 2, numRows, 3);
  output.setValues(inputValues.map(function (item) {
    return item[0];
  }).reduce(function (acc, email) {
    var treads = GmailApp.search('is:sent ' + email);
    if (!treads || treads.length < 1) {
      return [].concat(_toConsumableArray(acc), [['None', '', '']]);
    }
    var messages = treads.map(function (tread) {
      return tread.getMessages();
    });
    if (!messages || messages.length < 1) {
      return [].concat(_toConsumableArray(acc), [['Nothing', '', '']]);
    }
    var lastMessage = messages.slice().shift().pop();
    /*
    const values = messages.map(
      (message, id) => `'№: ${id} Subject: ${message[0].getSubject()} Body: ${message[0].getPlainBody()}`
    ).join('\n\n');
    */
    return [].concat(_toConsumableArray(acc), [[lastMessage.getSubject(), lastMessage.getPlainBody(), lastMessage.getDate()]]);
  }, []));
}

function onOpen(e) {
 const ui = SpreadsheetApp.getUi();
 ui.createMenu('Email Search')
  .addItem('search in sent', 'myFunction')
  .addToUi();
}
  

/* ES 2015
function myFunction() {
  const ss = SpreadsheetApp.getActiveSheet();
  // get email list
  const numRows = ss.getDataRange().getLastRow();
  const inputValues = ss.getRange(1, 1, numRows, 1).getValues(); // [[wer@erw.e], [wer@rf.rr]]
  const output = ss.getRange(1, 2, numRows, 3);
  output.setValues(inputValues
    .map(item => item[0])
    .reduce((acc, email) => {
        const treads = GmailApp.search('in:sent to:'+ email);
        if (!treads || treads.length < 1) {
          return [...acc, ['None', '', '']];
        }
        const messages = treads.map(tread => tread.getMessages());
        if (!messages || messages.length < 1) {
          return [...acc, ['Nothing', '', '']];
        }
        const lastMessage = messages.pop().slice();

//        const values = messages.map(
//          (message, id) => `'№: ${id} Subject: ${message[0].getSubject()} Body: ${message[0].getPlainBody()}`
//        ).join('\n\n');

        return [...acc, [lastMessage[0].getSubject(), lastMessage[0].getPlainBody(), lastMessage[0].getDate()]];
      },
      []
    )
  );
}
*/


