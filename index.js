const DATE_TO_ADD = {
  year: 2021,
  month: 2,
  days: [8, 10, 12],
};
function addBlankData() {
  INDUSTRIES.forEach((name) => {
    DATE_TO_ADD.days.forEach((day) => {
      const sheetRss = ssRss.getSheetByName(name);
      setSheetRss(sheetRss, DATE_TO_ADD.year, DATE_TO_ADD.month, day);
      const sheetStudent = ssStudent.getSheetByName(name);
      setSheetStudent(sheetStudent, DATE_TO_ADD.year, DATE_TO_ADD.month, day);
    });
  });
}

// create sheets
function createSheets() {
  recreateSheets();
  const sheetsRss = ssRss.getSheets();
  sheetsRss.forEach((sheet) => setSheetCommon(sheet));
  const sheetsStudent = ssStudent.getSheets();
  sheetsStudent.forEach((sheet) => setSheetCommon(sheet));
  const sheetsData = ssData.getSheets();
  sheetsData.forEach((sheet) => setSheetData(sheet));
}

function recreateSheets() {
  const sheetNamesRss = ssRss.getSheets().map((s) => s.getSheetName());
  const sheetNamesStudent = ssStudent.getSheets().map((s) => s.getSheetName());
  const sheetNamesData = ssData.getSheets().map((s) => s.getSheetName());
  INDUSTRIES.forEach((name) => {
    if (sheetNamesRss.some((s) => s === name))
      ssRss.deleteSheet(ssRss.getSheetByName(name));
    ssRss.insertSheet(name);
    // if (!sheetNamesRss.some((s) => s === name)) ssRss.insertSheet(name)

    if (sheetNamesStudent.some((s) => s === name))
      ssStudent.deleteSheet(ssStudent.getSheetByName(name));
    ssStudent.insertSheet(name);
    // if (!sheetNamesStudent.some((s) => s === name)) ssStudent.insertSheet(name)

    if (sheetNamesData.some((s) => s === name)) return;
    ssData.insertSheet(name);
  });
}
