// triggers
function setTriggers() {
  delTriggers();
  setOnEditRSS();
  setOnEditStudent();
  setSendRemidMail();
  setCreateMemberList();
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

function setCreateMemberList() {
  ScriptApp.newTrigger('createMemberList')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();
}

function delTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}
