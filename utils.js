function calcSectionNum(rowIdx, type) {
  // define a set of LABELS_INPUT_ROW_** as section
  const labelsLen =
    type === 'rss'
      ? LABELS_INPUT_ROW_RSS.length
      : LABELS_INPUT_ROW_STUDENT.length;
  return (
    Math.floor((rowIdx - (LABELS_TT_ROW.length + OFFSET_ROW) - 1) / labelsLen) +
    1
  );
}

function calcQuestionNum(rowIdx, type) {
  // define a set of LABELS_INPUT_ROW_** as section
  const labelsLen =
    type === 'rss'
      ? LABELS_INPUT_ROW_RSS.length
      : LABELS_INPUT_ROW_STUDENT.length;
  return ((rowIdx - (LABELS_TT_ROW.length + OFFSET_ROW) - 1) % labelsLen) + 1;
}

function calcSectionRowStartFrom(sectionNum, type) {
  const labelsLen =
    type === 'rss'
      ? LABELS_INPUT_ROW_RSS.length
      : LABELS_INPUT_ROW_STUDENT.length;
  return (sectionNum - 1) * labelsLen + LABELS_TT_ROW.length + OFFSET_ROW + 1;
}

function calcDataId(date, colIdx) {
  const dateStr = dateToStr(date);
  const period = calcPeriod(colIdx);
  const session = calcSession(colIdx);
  return `${dateStr}_${period}_${session}`;
}

function calcPeriod(colIdx) {
  return Math.floor((colIdx - OFFSET_COL - 2) / MAX_SESSION_NUM) + 1;
}

function calcSession(colIdx) {
  return ((colIdx - OFFSET_COL - 2) % MAX_SESSION_NUM) + 1;
}

function getEmail(studentId) {
  return 'e' + studentId.toString() + '@soka-u.jp';
}

function dateToStr(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}${month}${day}`;
}

function validateStudentId(value) {
  const str = value.toString();
  return !str || str.match(/^\d{2}(m|\d)\d{4}$/);
}

function fetchDataByDate(date) {
  const dateStr = dateToStr(date);
  let matchedData = {};
  INDUSTRIES.forEach((industry) => {
    matchedData[industry] = [];
    const sheetData = ssData.getSheetByName(industry);
    const lastRowIdx = sheetData.getLastRow() - 1;
    if (lastRowIdx === 0) return;

    const ids = sheetData.getRange(2, 1, lastRowIdx, 1).getValues();
    const matchedIds = ids.filter((id) => id[0].slice(0, 8) === dateStr);
    matchedIds.forEach((id) => {
      const rowIdxData = getRowIdxData(sheetData, id[0]);

      const studentId = sheetData.getRange(rowIdxData, 2).getValue().toString();
      const rssId = sheetData.getRange(rowIdxData, 4).getValue().toString();
      const studentName = sheetData.getRange(rowIdxData, 3).getValue();
      const rssName = sheetData.getRange(rowIdxData, 5).getValue();
      const date = sheetData.getRange(rowIdxData, 6).getValue();
      const period = sheetData.getRange(rowIdxData, 7).getValue();
      const session = sheetData.getRange(rowIdxData, 8).getValue();
      const content = sheetData.getRange(rowIdxData, 11).getValue();
      const rssEmail = getEmail(rssId);
      const studentEmail = getEmail(studentId);

      matchedData[industry].push({
        studentId,
        rssId,
        studentName,
        rssName,
        studentEmail,
        rssEmail,
        date,
        period,
        session,
        content,
        industry,
      });
    });
  });
  return matchedData;
}
