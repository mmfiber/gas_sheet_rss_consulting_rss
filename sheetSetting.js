function setSheetCommon(sheet) {
  const initCol = OFFSET_COL + 1;
  let lastCol = initCol + 1;
  sheet
    .getRange(OFFSET_ROW + 1, initCol, LABELS_TT_ROW.length)
    .setValues(LABELS_TT_ROW)
    .setHorizontalAlignment('center');
  TIMETABLE.forEach((t) => {
    sheet
      .getRange(OFFSET_ROW + 1, lastCol, 1, MAX_SESSION_NUM)
      .merge()
      .setValue(t)
      .setHorizontalAlignment('center');
    for (let i = 0; i < MAX_SESSION_NUM; i++) {
      sheet
        .getRange(OFFSET_ROW + 2, lastCol + i)
        .setValue(i + 1)
        .setHorizontalAlignment('center');
    }
    lastCol += MAX_SESSION_NUM;
  });
  sheet
    .getRange(
      OFFSET_ROW + LABELS_TT_ROW.length,
      initCol,
      1,
      TIMETABLE.length * MAX_SESSION_NUM + 1
    )
    .setBackground(LABEL_COLOR_SESSION);
}

function setSheetRss(sheet, year, month, day) {
  const initRow = sheet.getLastRow() + 1;
  const initCol = OFFSET_COL + 1;

  // set date
  setDate(sheet, initRow, year, month, day, 'rss');

  // set color
  sheet
    .getRange(initRow, initCol, 2, TIMETABLE.length * MAX_SESSION_NUM + 1)
    .setBackground(LABEL_COLOR_BASE);

  // set labels
  sheet
    .getRange(initRow, initCol, LABELS_INPUT_ROW_RSS.length)
    .setValues(LABELS_INPUT_ROW_RSS)
    .setHorizontalAlignment('center');
}

function setSheetStudent(sheet, year, month, day) {
  const initCol = OFFSET_COL + 1;
  const initRow = sheet.getLastRow() + 1;

  // set date
  setDate(sheet, initRow, year, month, day, 'student');

  // set color
  sheet.getRange(initRow, initCol, 2, 1).setBackground(LABEL_COLOR_BASE);
  sheet
    .getRange(
      initRow,
      initCol + 1,
      LABELS_INPUT_ROW_STUDENT.length,
      TIMETABLE.length * MAX_SESSION_NUM
    )
    .setBackground(COLOR_UNAVAILABLE);

  // set labels and default data
  sheet.setColumnWidth(initCol, 300);
  sheet
    .getRange(initRow, initCol, LABELS_INPUT_ROW_STUDENT.length)
    .setValues(LABELS_INPUT_ROW_STUDENT)
    .setHorizontalAlignment('center');
  LABELS_INPUT_ROW_STUDENT.forEach((_, idx) => {
    if (idx !== 2) return;
    const cell = sheet.getRange(
      initRow + idx,
      initCol + 1,
      1,
      TIMETABLE.length * MAX_SESSION_NUM
    );
    cell.setDataValidation(CONSULTING_CONTENT_RULE);
  });
}

function setSheetData(sheet) {
  sheet.getRange(1, 1, 1, DATA_COLMUNS.length).setValues([DATA_COLMUNS]);
}

function setDate(sheet, initRow, year, month, day, type) {
  const labelsLen =
    type === 'rss'
      ? LABELS_INPUT_ROW_RSS.length
      : LABELS_INPUT_ROW_STUDENT.length;
  const date = year + '/' + month + '/' + day;
  sheet
    .getRange(initRow, 1, labelsLen, 1)
    .merge()
    .setValue(date)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
}
