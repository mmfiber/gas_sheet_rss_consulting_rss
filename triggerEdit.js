function onEditRss(e) {
  const debug = { range: e.range.getA1Notation(), value: e.value };
  logger('onEditRss()', debug);

  const range = e.range;
  const rowIdx = range.getRow();
  const colIdx = range.getColumn();
  for (let i = 0; i < range.getNumColumns(); i++) {
    for (let j = 0; j < range.getNumRows(); j++) {
      if (
        rowIdx + j < CHECK_RANGE.from.rowIdx ||
        colIdx + i < CHECK_RANGE.from.colIdx
      )
        continue;
      _onEditRss(e, rowIdx + j, colIdx + i);
    }
  }
}

function _onEditRss(e, rowIdx, colIdx) {
  const sheetRss = e.source.getActiveSheet();
  const sheetStudent = ssStudent.getSheetByName(sheetRss.getSheetName());
  const sheetData = ssData.getSheetByName(sheetRss.getSheetName());

  const value = sheetRss.getRange(rowIdx, colIdx).getValue();

  const debug = { value, rowIdx, colIdx, sheetName: sheetRss.getSheetName() };
  logger('func: editRss()', debug);

  const sectionNum = calcSectionNum(rowIdx, 'rss');
  const questionNum = calcQuestionNum(rowIdx, 'rss');

  const rowIdxStartFromRss = calcSectionRowStartFrom(sectionNum, 'rss');
  const rowIdxStartFromStudent = calcSectionRowStartFrom(sectionNum, 'student');
  const rowIdxStudentIdRss = rowIdxStartFromRss + 1;

  // validate student ID and return if invalid the ID is invalid
  if (rowIdx === rowIdxStudentIdRss && !validateStudentId(value))
    return showStudentIdAlert(sheetRss, rowIdx, colIdx, value);

  const date = sheetRss
    .getRange(rowIdxStartFromRss, 1, LABELS_INPUT_ROW_RSS.length, 1)
    .getValue(); // suppose to be Date object
  const dataId = calcDataId(date, colIdx);
  const rowIdxData = getRowIdxData(sheetData, dataId);

  // set data sheet
  let colIdxData = null;
  if (questionNum === 1) colIdxData = 5;
  if (questionNum === 2) colIdxData = 4;
  if (colIdxData) {
    const period = calcPeriod(colIdx);
    const session = calcSession(colIdx);
    sheetData.getRange(rowIdxData, colIdxData).setValue(value);
    sheetData.getRange(rowIdxData, 6).setValue(date);
    sheetData.getRange(rowIdxData, 7).setValue(period);
    sheetData.getRange(rowIdxData, 8).setValue(session);
    sheetData.getRange(rowIdxData, 10).setValue(new Date()); // updated_at
  }

  // set student sheet
  const isAvailable = isAllValuesSet(sheetRss, rowIdxStartFromRss, colIdx);
  updateReservationStatus(
    sheetRss,
    sheetStudent,
    rowIdxStartFromRss,
    rowIdxStartFromStudent,
    colIdx,
    isAvailable
  );
}

function onEditStudent(e) {
  const debug = { range: e.range.getA1Notation(), value: e.value };
  logger('onEditStudent()', debug);

  const range = e.range;
  const rowIdx = range.getRow();
  const colIdx = range.getColumn();
  for (let i = 0; i < range.getNumColumns(); i++) {
    for (let j = 0; j < range.getNumRows(); j++) {
      if (
        rowIdx + j < CHECK_RANGE.from.rowIdx ||
        colIdx + i < CHECK_RANGE.from.colIdx
      )
        continue;
      _onEditStudent(e, rowIdx + j, colIdx + i);
    }
  }
}

function _onEditStudent(e, rowIdx, colIdx) {
  const sheetStudent = e.source.getActiveSheet();
  const sheetData = ssData.getSheetByName(sheetStudent.getSheetName());

  const value = sheetStudent.getRange(rowIdx, colIdx).getValue();

  const debug = {
    value,
    rowIdx,
    colIdx,
    sheetName: sheetStudent.getSheetName(),
  };
  logger('_onEditStudent()', debug);

  const sectionNum = calcSectionNum(rowIdx, 'student');
  const questionNum = calcQuestionNum(rowIdx, 'student');

  const rowIdxStartFromStudent = calcSectionRowStartFrom(sectionNum, 'student');
  const rowIdxStudentId = rowIdxStartFromStudent + 1;

  if (rowIdx === rowIdxStudentId && !validateStudentId(value))
    return showStudentIdAlert(sheetStudent, rowIdx, colIdx, value);

  const date = sheetStudent
    .getRange(rowIdxStartFromStudent, 1, LABELS_INPUT_ROW_STUDENT.length, 1)
    .getValue(); // suppose to be Date object
  const dataId = calcDataId(date, colIdx);
  const rowIdxData = getRowIdxData(sheetData, dataId);

  // set data
  let colIdxData = null;
  if (questionNum === 1) colIdxData = 3;
  if (questionNum === 2) colIdxData = 2;
  if (questionNum === 3) colIdxData = 11;
  if (!colIdxData) return;

  const period = calcPeriod(colIdx);
  const session = calcSession(colIdx);
  sheetData.getRange(rowIdxData, colIdxData).setValue(value);
  sheetData.getRange(rowIdxData, 6).setValue(date);
  sheetData.getRange(rowIdxData, 7).setValue(period);
  sheetData.getRange(rowIdxData, 8).setValue(session);
  sheetData.getRange(rowIdxData, 10).setValue(new Date()); // updated_at

  // hide student id
  if (rowIdx === rowIdxStudentId && value)
    sheetStudent.getRange(rowIdx, colIdx).setValue('*******');
}

function isAllValuesSet(sheet, initRowIdx, initColIdx) {
  let isSet = true;
  const values = sheet
    .getRange(initRowIdx, initColIdx, LABELS_INPUT_ROW_RSS.length, 1)
    .getValues();
  values.forEach((v) => {
    if (v.length === 1 && v[0] === '') isSet = false;
  });
  return isSet;
}

function getRowIdxData(sheetData, dataId) {
  const dataIdx = getDataIdx(sheetData, dataId);
  const existDataId = dataIdx !== -1;

  if (existDataId) return dataIdx + 2;

  const rowIdxData = sheetData.getLastRow() + 1;
  sheetData.getRange(rowIdxData, 1).setValue(dataId); // id
  sheetData.getRange(rowIdxData, 9).setValue(new Date()); // created_at
  return rowIdxData;
}

function getDataIdx(sheetData, dataId) {
  const rowIdxLastData = sheetData.getLastRow();
  if (rowIdxLastData < 2) return -1;
  const dataIds = sheetData.getRange(2, 1, rowIdxLastData - 1, 1).getValues();
  let concated = [];
  dataIds.forEach((id) => (concated = concated.concat(id)));
  return concated.indexOf(dataId);
}

function showStudentIdAlert(sheet, rowIdx, colIdx, value) {
  const ui = SpreadsheetApp.getUi();
  const msg = studentIdHelpMsg(value);
  ui.alert('学籍番号が不適切です', msg, ui.ButtonSet.OK);
  sheet.getRange(rowIdx, colIdx).setValue('');
}

function studentIdHelpMsg(value) {
  return `
    ${value}は正しい学籍番号ではありません。
    以下をご確認いただき再度入力をお願いします。

    ・学籍番号が半角英数字7桁になっている
    ・院生の方は'm'が小文字になっている
  `;
}

function updateReservationStatus(
  sheetRss,
  sheetStudent,
  rowIdxStartFromRss,
  rowIdxStartFromStudent,
  colIdx,
  isAvailable
) {
  if (isAvailable) {
    // for "担当RSSの詳しい業界"
    const outputRowIdx =
      rowIdxStartFromStudent + LABELS_INPUT_ROW_STUDENT.length - 1;
    const value = sheetRss
      .getRange(rowIdxStartFromRss + LABELS_INPUT_ROW_RSS.length - 1, colIdx)
      .getValues();
    sheetStudent.getRange(outputRowIdx, colIdx).setValues(value);
    sheetStudent
      .getRange(rowIdxStartFromStudent, colIdx, 2, 1)
      .setFontColor('black')
      .setBackground(LABEL_COLOR_BASE);
    sheetStudent
      .getRange(rowIdxStartFromStudent + 2, colIdx, 2, 1)
      .setFontColor('black')
      .setBackground(COLOR_AVAILABLE);
  } else {
    sheetStudent
      .getRange(
        rowIdxStartFromStudent,
        colIdx,
        LABELS_INPUT_ROW_STUDENT.length,
        1
      )
      .setFontColor(COLOR_UNAVAILABLE)
      .setBackground(COLOR_UNAVAILABLE);
  }
}
