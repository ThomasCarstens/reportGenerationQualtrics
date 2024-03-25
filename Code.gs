


function doPost(e) {

  var dataContents = JSON.parse(e.postData.contents);

  // Parameters from Qualtrics:
  var startup = dataContents.startup
  var email = dataContents.email2
  var token = dataContents.token

  
  var triggerReport = dataContents.triggerReport

  // Logger is sent by email to debug the app - change to yours if you need a debugger.
  // This was a workaround since Google Cloud refuses to display the app when Qualtrics API triggers the app.
  Logger.log(e)
  Logger.log(dataContents)  
  Logger.log(startup)
  Logger.log(email.toString())

  MailApp.sendEmail({ to: 'thomas.carstens@outlook.com', 
                      subject: 'Log from Startup ',           
                      body: Logger.getLog()
                      })


  // create Report IF reportTrigger activated.
  if (triggerReport == 1){

      generateReport(email, startup, triggerReport, token)
      MailApp.sendEmail({ to: 'thomas.carstens@outlook.com', 
                      subject: 'Report generated and sent to' + email.toString(),           
                      body: Logger.getLog()
                      })

  } else {
          MailApp.sendEmail({ to: 'thomas.carstens@outlook.com', 
                      subject: 'No report requested ',           
                      body: Logger.getLog()
                      })
  }
  
  // In comments: Methods for Hiding Sheets / Removing Sheets

  // Hide relevant sheets in order to send only 1 as pdf. New code: only takes the selected Sheet.
  // active_spreadsheet.getSheetByName("Pivot-Tabelle 1").hideSheet();
  // active_spreadsheet.getSheetByName("ALL").hideSheet();
  // active_spreadsheet.getSheetByName("Template").hideSheet();
  

  // Removes the Radar Graph Sheet permanently. @18.03 Meeting: DO NOT remove a sheet. Keep it.
  // sheetId = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(... Add the custom name here based on startup and UTC formatted date)
  // SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetId);


  // POST Requests must always send info back to client (Qualtrics), but we don't use this info
  var params = JSON.stringify(e);
  return HtmlService.createHtmlOutput(params);

};

function generateReport(email, startup, apiTriggered=0, token) {

  if (apiTriggered == 1){
    /* API Code. */
    sheetName = generateRadarGraph(token, startup) 
    sendExportedSheetAsPDFAttachment(email.toString(), sheetName, startup)
  } else {
  /* Test purposes (without API) */
    sheetName = generateRadarGraph('VLSC003', 'startup X') 
    sendExportedSheetAsPDFAttachment('alex.stamler@unternehmertum.de', sheetName, 'startup X')
  }
}


function sendExportedSheetAsPDFAttachment(email, sheetName, startupName) {
  var spreadsheetId = "1xr4f3sH_19781VnWFZS36aV3CzNGDx6xMj5cvJvXt3U"; // DEPENDENCY on the Google Sheet
  var ss = SpreadsheetApp.openById(spreadsheetId);

  var sheetRef = ss.getSheetByName(sheetName)

  
  sheetId = sheetRef.getSheetId()
  console.log('sheet id', sheetId.toString())

  // Define PDF export URL with desired options
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?";
  var exportOptions = 'exportFormat=pdf&format=pdf' + // Export format
                      '&size=letter' + // Paper size
                      '&portrait=false' + // Orientation, false for landscape
                      '&fitw=true&source=labnol' + // Fit to width, true
                      '&sheetnames=false&printtitle=false' + // Hide sheet names and document title
                      '&pagenumbers=false&gridlines=false' + // Hide page numbers and gridlines
                      '&fzr=false' + // Repeat frozen rows in each page
                      '&gid=' + sheetId; // Sheet ID : only takes the generated sheet from Google Sheets

  var headers = {
    "Authorization": "Bearer " + ScriptApp.getOAuthToken()
  };

  var options = {
    headers: headers,
    method: "get"
  };

  var response = UrlFetchApp.fetch(url + exportOptions, options);
  var blob = response.getBlob().setName(startupName+" Entry Gate Report.pdf");

  var message = {
    to: email.toString(),
    subject: "New Coach Report "+startupName+" - IRL Entry Gate in Qualtrics",
    body: "Dear Coach,\n\nPlease find attached the IRL Entry Gate report for "+startupName+".\n\nThank you,\nTUM Venture Labs - Startup Assessment Taskforce ",
    name: "TUM Venture Labs",
    attachments: [blob]
  };
  
  MailApp.sendEmail(message);
}


