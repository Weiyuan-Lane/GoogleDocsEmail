/* Using templating, connecting CSS, JS and HTML files to main HTML file.
 *
 * https://developers.google.com/apps-script/guides/support/best-practices
 * https://yagisanatode.com/2018/04/15/google-apps-script-how-to-create-javascript-and-css-files-for-a-sidebar-project-in-google-apps-script/
 */

function sendEmail(emailPayload, changeset, fileId){
  console.log(emailPayload, changeset, fileId)
  var to = emailPayload['to'].replace(';', ',');
  var cc = emailPayload['cc'].replace(';', ',');
  var bcc = emailPayload['bcc'].replace(';', ',');
  
  MailApp.sendEmail(to, emailPayload['subject'], '', {
   cc: cc,
   bcc: bcc,
   htmlBody: getDocAsHtmlFormat(changeset, fileId)
  });

  alert('Email has been sent!');
  // return(getDocAsHtmlFormat(changeset, fileId))
}

/*
 * used to dynamically include files in html template
 * allows separation of stylesheet and javascript 
 * 
 * use as:  `<?!= include('STYLESHEET'); ?>` in the html template
 */  
function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
}

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  SpreadsheetApp.getUi()
             .createAddonMenu()
             .addItem('Email with template', 'showSidebar')
             .addToUi();
}

function showSidebar() {
  var html= HtmlService
      .createTemplateFromFile('sidebar')
      .evaluate();
  html.setTitle('Send Email');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getFilesInCurrentFolder() {
  var thisFileId = SpreadsheetApp.getActive().getId();
  var thisFile = DriveApp.getFileById(thisFileId);
  var parentFolder = thisFile.getParents();
  
  // Loop through all the files
  var folder = parentFolder.next();
  var files = folder.getFiles();
  var fileList = []
  while(files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
      fileList.push({ name: file.getName(), id: file.getId() })
    }
  }
  
  return fileList
}

function getFile(fileId) {
  return DocumentApp.openById(fileId).getBody().getText();
}

function getTemplateVars(s, fileId) {
  var bodyText = getFile(fileId)
  var regex = new RegExp(s + '[A-Z_]+' + s, "g");
  var matches = {};
  var key = null;
  
  while ((key = regex.exec(bodyText)) != null) {
    if (matches[key]) {
      matches[key] += 1;
    } else { 
      matches[key] = 1;
    }
  }
  
  return matches;
}

function previewCurrentEmail(changeset, fileId) {
  var ui = SpreadsheetApp.getUi();
  var html = getDocAsHtmlFormat(changeset, fileId);
  
  var htmlOutput = HtmlService
    .createHtmlOutput(html)
    .setWidth(800)
    .setHeight(600);
  ui.showModalDialog(htmlOutput, 'Email Preview');
}

function getColumnHeaders() {
  var sheet = SpreadsheetApp.getActive()
  var headings = sheet.getDataRange()
                    .offset(0, 0, 1)
                    .getValues()[0];
  return headings
}

function getColumnValues() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var headers = getColumnHeaders();
  var results = {};
  for(var index in headers) {
    var key = headers[index];
    var rows = sheet.getDataRange().getValues()
    var values = rows.slice(1).map(function(row) {
      return row[index]
    });
    
    results[key]  = values
  }
  return results
}

// UTILITY METHODS
function alert(msg){
  var ui = SpreadsheetApp.getUi();
  ui.alert(msg, ui.ButtonSet.OK);
}

function putToCache(key, value) {
  var cache = CacheService.getScriptCache();
  cache.put(key, value, 31540000);
}

function getFromCache(key) {
  var cache = CacheService.getScriptCache();
  return cache.get(key);
}

function batchToCache(keyObj) {
  Object.keys(keyObj).forEach(function(key){
    putToCache(key, keyObj[key]);
  });
}

function listFromCache(keyList) {
  var cache = CacheService.getScriptCache();
  var result = {};
  
  keyList.forEach(function(key){
    result[key] = cache.get(key);
  });
  
  return result;
}

function getDocAsHtmlFormat(changeset, fileId) {
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + fileId + "&exportFormat=html";
  var forDriveScope = DriveApp.getStorageUsed();
  var param = {
    method: "get",
    headers: {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true,
  };
  
  html = UrlFetchApp.fetch(url, param).getContentText();
  // Swap values as and when required
  for(var key in changeset) {
      var regex = new RegExp(key, 'g');
      html = html.replace(regex, changeset[key]); 
  }
  
  return html;
}