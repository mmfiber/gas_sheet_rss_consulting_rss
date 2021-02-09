// triggers
function setTriggers() {
  delTriggers();
  setOnEditRSS();
  setOnEditStudent();
  setSendRemidMail();
}

function setOnEditRSS() {
  ScriptApp.newTrigger('onEditRss').forSpreadsheet(ssRss).onEdit().create();
}

function setOnEditStudent() {
  ScriptApp.newTrigger('onEditStudent')
    .forSpreadsheet(ssStudent)
    .onEdit()
    .create();
}

function setSendRemidMail() {
  ScriptApp.newTrigger('sendRemidMail')
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .create();
}

function delTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}
