function pushNotificationToGoogleSpace() {
  var sheet = SpreadsheetApp.openById('').getSheetByName('');
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return; // Exit if no data
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
  var data = dataRange.getValues();

  var spreadsheetUrl = "";
  var groupedData = {};

  // Group data by email ID
  for (var i = 0; i < data.length; i++) {
    var centerName = data[i][0];
    var batchCode = data[i][1];
    var subject = data[i][2];
    var lecStartTime = data[i][3];
    var emailId = data[i][4];
    var errorMsg = data[i][5];
    var sheetName = data[i][6];

    if (emailId) {
      if (!groupedData[emailId]) {
        groupedData[emailId] = [];
      }
      groupedData[emailId].push({
        centerName: centerName,
        batchCode: batchCode,
        subject: subject,
        lecStartTime: lecStartTime,
        errorMsg: errorMsg,
        sheetName: sheetName
      });
    }
  }

  // Send grouped notifications
  for (var emailId in groupedData) {
    sendGroupedNotificationToGoogleSpace(emailId, groupedData[emailId], spreadsheetUrl);
  }
}

function sendGroupedNotificationToGoogleSpace(emailId, entries, spreadsheetUrl) {
  var webhookUrl = '';

  // Construct the grouped message text
  var messageText = `<b>Notifications for: ${emailId}</b><br><br>`;
  entries.forEach(function (entry, index) {
    messageText += `
<b>Entry ${index + 1}:</b><br>
<b>Center:</b> ${entry.centerName}<br>
<b>Batch:</b> ${entry.batchCode}<br>
<b>Subject:</b> ${entry.subject}<br>
<b>Lecture Date:</b> ${entry.lecStartTime}<br>
<b>Error:</b> ${entry.errorMsg}<br>
<b>Sheet Name:</b> ${entry.sheetName}<br><br>`;
  });

  messageText += `<i>You can view the full details <a href="${spreadsheetUrl}">here</a>.</i>`;

  // Create the payload
  var message = {
    "cards": [
      {
        "header": {
          "title": `Grouped Notification for ${emailId}`,
          "subtitle": "Pending Batch Entries in Audit Sheet",
          "imageUrl": "https://lh3.googleusercontent.com/proxy/QdmwVz21UmInZ72wt3b6WK0ND4316O6nRJs7o12gX--cLsz-VuWyPXPT568TxRXq-41RKLo_7eKrEg33gfKYDNJN_gAbP3f7milXKiqua0a--3Tttb9G5lpJ-DIJkfJnqajz8y-JDXcVunKpVBULOM3vfkjYQaftBrKxUwy4CLEK_PKt_CNOuGUCS7g8"
        },
        "sections": [
          {
            "widgets": [
              {
                "textParagraph": {
                  "text": messageText
                }
              }
            ]
          }
        ]
      }
    ]
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(message)
  };

  try {
    UrlFetchApp.fetch(webhookUrl, options);
  } catch (e) {
    if (e.message.includes("429")) { 
      Utilities.sleep(60000);
      sendGroupedNotificationToGoogleSpace(emailId, entries, spreadsheetUrl);
    } else {
      Logger.log(`Error sending notification to ${emailId}: ${e.message}`);
    }
  }
}

function deleteOldTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function createHourlyTrigger() {
  deleteOldTriggers(); 
  ScriptApp.newTrigger('pushNotificationToGoogleSpace')
    .timeBased()
    .everyHours(1)
    .create();
}
