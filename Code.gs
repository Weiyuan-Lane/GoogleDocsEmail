function sendEmail(emailPayload, changeset){
  var to = emailPayload['to'].replace(';', ',');
  var cc = emailPayload['cc'].replace(';', ',');
  var bcc = emailPayload['bcc'].replace(';', ',');
  
  MailApp.sendEmail(to, emailPayload['subject'], '', {
    cc: cc,
    bcc: bcc,
    htmlBody: getDocAsHtmlFormat(changeset)
  });
  alert('Email has been sent!');
}


function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  DocumentApp.getUi()
             .createAddonMenu()
             .addItem('Send with template', 'showSidebar')
             .addToUi();
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
                      .setTitle('Send Email');
  DocumentApp.getUi().showSidebar(ui);
}

function getTemplateVars(s) {
  var bodyText = DocumentApp.getActiveDocument().getBody().getText();
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

function previewCurrentEmail(changeset) {
  var ui = DocumentApp.getUi();
  var html = getDocAsHtmlFormat(changeset);
  
  var htmlOutput = HtmlService
    .createHtmlOutput(html)
    .setWidth(800)
    .setHeight(600);
  ui.showModalDialog(htmlOutput, 'Email Preview');
}

// UTILITY METHODS
function alert(msg){
  var ui = DocumentApp.getUi();
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

function getDocAsHtmlFormat(changeset) {
  var id = DocumentApp.getActiveDocument().getId();
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + id + "&exportFormat=html";
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
  changeset.forEach(function(change) {
    if (change.key && change.value) {
      var regex = new RegExp(change.key, 'g');
      html = html.replace(regex, change.value); 
    }
  });
  
  return html;
}