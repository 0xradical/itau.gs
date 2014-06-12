// based on http://ctrlq.org/code/19053-send-to-google-drive
function sendToGoogleDrive() {

  var sheet = SpreadsheetApp.getActiveSheet();

  var gmailLabels           = 'inbox';
  var driveFolder           = 'Itaú Notifications';
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

  if ( ! moveToLabel ) {
    moveToLabel = GmailApp.createLabel(archiveLabel);
  }

  // Create folder 'Itaú Notifications' if it doesn't exist
  var folder = DriveApp.getFoldersByName(driveFolder);

  if (folder.hasNext()) {
    folder = folder.next();
  } else {
    folder = DriveApp.createFolder(driveFolder);
  }

  var threads = GmailApp.search(filter, 0, 5);

  for (var x = 0; x < threads.length; x++) {

    var message = threads[x].getMessages()[0];

    var plainBody = message.getPlainBody();

    Logger.log(e.toString());

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


function configure() {
  reset();
  ScriptApp.newTrigger("sendToGoogleDrive").timeBased().everyMinutes(5).create();
  Browser.msgBox("Initialized", "The program is now running.", Browser.Buttons.OK);
}

function onOpen() {
  var menu = [
    { name: "Step 1: Authorize",   functionName: "configure" },
    { name: "Step 2: Run Program", functionName: "configure" },
    { name: "Uninstall (Stop)",    functionName: "reset"     }
  ];

  SpreadsheetApp.getActiveSpreadsheet().addMenu("Itaú Notifications", menu);
}

function reset() {

  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

}
