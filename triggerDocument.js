function createMemberList() {
  const today = new Date();
  const data = fetchDataByDate(today);

  let isBlank = true;
  for (const prop in data) {
    if (data[prop].length === 0) continue;
    isBlank = false;
    break;
  }
  if (isBlank) return;

  const doc = createDocument(dateToStr(today));
  const body = doc.getBody();
  let concated = [];
  Object.values(data).forEach((d) => (concated = concated.concat(d)));

  body
    .insertParagraph(0, '--- 詳細 ---（※GDを含まない）')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const overview = {};
  const arrDevidedByPeriod = devideArrayByProp(concated, 'period');
  for (const periodStr in arrDevidedByPeriod) {
    if(!TIMETABLE[parseInt(periodStr) - 1]) continue;

    addList(body, TIMETABLE[parseInt(periodStr) - 1], 0).setHeading(
      DocumentApp.ParagraphHeading.HEADING3
    );

    let rssNum = 0;
    let studentNum = 0;

    const arrDevidedByIndustry = devideArrayByProp(
      arrDevidedByPeriod[periodStr],
      'industry'
    );
    for (const industry in arrDevidedByIndustry) {
      addList(body, industry, 1);

      const sorted = arrDevidedByIndustry[industry].sort(
        (a, b) => a.session - b.session
      );
      sorted.forEach((obj) => {
        let rssInfo = null;
        let studentInfo = null;
        if (validateStudentId(obj.rssId) && obj.rssName) {
          rssInfo = { name: obj.rssName, email: obj.rssEmail };
          rssNum++;
        }
        if (validateStudentId(obj.studentId) && obj.studentName) {
          studentInfo = { name: obj.studentName, email: obj.studentEmail };
          studentNum++;
        }
        if (rssInfo || studentInfo) {
          addList(body, `session: ${obj.session}`, 2);
          if (rssInfo) {
            addList(body, 'RSS', 3);
            addList(body, `名前: ${rssInfo.name}`, 4);
            addList(body, `メール: ${rssInfo.email}`, 4);
          }
          if (studentInfo) {
            addList(body, '学生', 3);
            addList(body, `名前: ${studentInfo.name}`, 4);
            addList(body, `メール: ${studentInfo.email}`, 4);
          }
        }
      });
    }
    overview[TIMETABLE[parseInt(periodStr) - 1]] = {
      rssNum,
      studentNum,
      sum: rssNum + studentNum,
    };
  }
  body
    .insertParagraph(0, '--- 概要 ---（※GDを含まない）')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const periods = Object.keys(overview);
  periods.reverse().forEach((period) => {
    addList(body, `合計: ${overview[period]['sum']}人`, 1, 1);
    addList(body, `学生: ${overview[period]['studentNum']}人`, 1, 1);
    addList(body, `RSS: ${overview[period]['rssNum']}人`, 1, 1);
    addList(body, period, 0, 1);
  });
  setGlyph(body);
}

function createDocument(name) {
  const doc = DocumentApp.create(name);
  const file = DriveApp.getFileById(doc.getId());
  const id = getProperty('DOCUMENT_FOLDER_ID');
  DriveApp.getFolderById(id).addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return doc;
}

function addList(body, item, nestLevel, insertIdx) {
  if (insertIdx)
    return body.insertListItem(insertIdx, item).setNestingLevel(nestLevel);
  return body.appendListItem(item).setNestingLevel(nestLevel);
}

const GLYPH = [
  DocumentApp.GlyphType.BULLET,
  DocumentApp.GlyphType.HOLLOW_BULLET,
  DocumentApp.GlyphType.SQUARE_BULLET,
];
function setGlyph(body) {
  body.getListItems().forEach((item) => {
    level = item.getNestingLevel();
    item.setGlyphType(GLYPH[level % 3]);
  });
}

function devideArrayByProp(arr, prop) {
  const devided = {};
  arr.forEach((obj) => {
    if (!devided[obj[prop]]) devided[obj[prop]] = [];
    devided[obj[prop]].push(obj);
  });
  return devided;
}
