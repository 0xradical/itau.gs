// based on http://ctrlq.org/code/19053-send-to-google-drive
function sendToGoogleDrive() {

  var gmailLabels           = 'inbox';
  var driveFolder           = 'Itaú Notifications';
  var spreadsheetName       = 'itau';
  var archiveLabel          = 'itau.processed';
  var itauNotificationEmail = 'comunicacaodigital@itau-unibanco.com.br';
  var filter                = "from: " +
                              itauNotificationEmail +
                              " -label:" +
                              archiveLabel +
                              " label:" +
                              gmailLabels;

  // Create label for 'itau.processed' if it doesn't exist
  var moveToLabel =  GmailApp.getUserLabelByName(archiveLabel);

  if (!moveToLabel) {
    moveToLabel = GmailApp.createLabel(archiveLabel);
  }

  // Create folder 'Itaú Notifications' if it doesn't exist
  var folders = DriveApp.getFoldersByName(driveFolder);
  var folder;

  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(driveFolder);
  }

  // Create spreadsheet file 'itau' if it doesn't exist
  var files = folder.getFilesByName(spreadsheetName);

  // File is in DriveApp
  // Doc is in SpreadsheetApp
  // They are not interchangeable
  var file, doc;

  // Confusing :\
  // As per: https://code.google.com/p/google-apps-script-issues/issues/detail?id=3578
  if (files.hasNext()){
    file = files.next();
    doc = SpreadsheetApp.openById(file.getId());
  } else {
    doc = SpreadsheetApp.create(spreadsheetName);
    file = DriveApp.getFileById(doc.getId());
    folder.addFile(file);
    DriveApp.removeFile(file);
  }

  var sheet = doc.getSheets()[0];

  // Append header if first line
  if(sheet.getLastRow() == 0){
    sheet.appendRow(['Conta', 'Operação', 'Valor', 'Data', 'Hora']);
  }

  var message, account, operation, value, date, hour, plainBody;
  var accountRegex = /Conta: (XXX[0-9\-]+)/;
  var operationRegex = /Tipo de operação: ([A-Z]+)/;
  var valueRegex = /Valor: R\$ ([0-9\,]+)/;
  var dateRegex = /Data: ([0-9\/]+)/;
  var hourRegex = /Hora: ([0-9\:]+)/;

  var threads = GmailApp.search(filter, 0, 5);

  for (var x = 0; x < threads.length; x++) {
    message = threads[x].getMessages()[0];

    plainBody = message.getPlainBody();

    if(accountRegex.test(plainBody)) {
      account = RegExp.$1;
    }

    if(operationRegex.test(plainBody)) {
      operation = RegExp.$1;
    }

    if(valueRegex.test(plainBody)) {
      value = RegExp.$1;
    }

    if(dateRegex.test(plainBody)) {
      date = RegExp.$1;
    }

    if(hourRegex.test(plainBody)) {
      hour = RegExp.$1;
    }

    sheet.appendRow([account, operation, value, date, hour]);

    // Logger.log(account);
    // Logger.log(operation);
    // Logger.log(value);
    // Logger.log(date);
    // Logger.log(hour);

    threads[x].addLabel(moveToLabel);
  }

}
