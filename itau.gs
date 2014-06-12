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
  var folder = DriveApp.getFoldersByName(driveFolder);

  if (folder.hasNext()) {
    folder = folder.next();
  } else {
    folder = DriveApp.createFolder(driveFolder);
  }

  // Create spreadsheet file 'itau' if it doesn't exist
  var file = folder.getFilesByName(spreadsheetName);

  if (file.hasNext()){
    file = file.next();
  } else {
    // https://developers.google.com/apps-script/reference/drive/folder?hl=pt-br#createFile(String,String,String)
    // https://developers.google.com/apps-script/reference/base/mime-type?hl=pt-BR
    // file = folder.createFile(spreadsheetName, '', MimeType.GOOGLE_SHEETS);
    file = folder.createFile(spreadsheetName, null, MimeType.GOOGLE_SHEETS);
  }

  // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app?hl=pt-br#open(File)
  var spreadsheet = SpreadsheetApp.open(file);

  var message, account, operation, value, date, hour, plainBody;
  var accountRegex = /Conta: (XXX[0-9\-]+)/;
  var operationRegex = /Tipo de operação: ([A-Z]+)/;
  var valueRegex = /Valor: R\$ ([0-9\,]+)/;
  var dateRegex = /Data: ([0-9\/]+)/;
  var hourRegex = /Hora: ([0-9\:]+)/;

  var threads = GmailApp.search(filter, 0, 5);

  // for (var x = 0; x < threads.length; x++) {
  for (var x = 0; x < 1; x++) {
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

    Logger.log(account);
    Logger.log(operation);
    Logger.log(value);
    Logger.log(date);
    Logger.log(hour);

    // Logger.log(plainBody);

    // var desc   = message.getSubject() + " #" + message.getId();
    // var att    = message.getAttachments();

    // for (var z = 0; z < att.length; z++) {
    //   try {
    //     file = folder.createFile(att[z]);
    //     file.setDescription(desc);
    //   }
    //   catch (e) {
    //     Logger.log(e.toString());
    //   }
    // }

    // threads[x].addLabel(moveToLabel);
  }

}
