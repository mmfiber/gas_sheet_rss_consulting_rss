function setProperty() {
  const prop = PropertiesService.getScriptProperties();
  // prop.setProperty("", "");
  console.log(prop.getKeys());
}

function getProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}
// consts
var ssRss = SpreadsheetApp.openById(getProperty('SPREADSHEET_ID_RSS'));
var ssStudent = SpreadsheetApp.openById(getProperty('SPREADSHEET_ID_STUDENT'));
var ssData = SpreadsheetApp.openById(getProperty('SPREADSHEET_ID_DATA'));

const INDUSTRIES = [
  'メーカー',
  'IT・コンサル',
  '建設',
  '小売・流通・印刷',
  '金融',
  '人材・教育・サービス業・エンタメ',
];
const TIMETABLE = [
  '1コマ(9:00-10:30)',
  '2コマ(10:45-12:15)',
  '3コマ(13:05-14:35)',
  '4コマ(14:50-16:20)',
];
const MAX_SESSION_NUM = 3;
const LABELS_TT_ROW = [['時間割'], ['セッション']];
const LABELS_INPUT_ROW_RSS = [['名前'], ['学籍番号'], ['詳しい業界']];
const LABELS_INPUT_ROW_STUDENT = [
  ['名前'],
  ['学籍番号'],
  ['相談内容'],
  ['担当RSSの詳しい業界'],
];
const CONSULTING_CONTENT = ['面接対策', '相談', 'ES添削'];
const LABEL_COLOR_BASE = '#d9ead3';
const LABEL_COLOR_SESSION = '#93c47d';
const COLOR_AVAILABLE = '#ffffff';
const COLOR_UNAVAILABLE = '#666666';
const OFFSET_ROW = 2;
const OFFSET_COL = 1;
const CHECK_RANGE = {
  from: {
    rowIdx: OFFSET_ROW + LABELS_TT_ROW.length + 1,
    colIdx: OFFSET_COL + 2,
  },
};
const DATA_COLMUNS = [
  'id',
  'student_id',
  'student_name',
  'rss_id',
  'rss_name',
  'date',
  'period',
  'session_num',
  'created_at',
  'updated_at',
  'content',
];
const ZOOM_URL =
  'https://us02web.zoom.us/j/9522499468?pwd=WmYzdjkwT0U1N1J6d01GWXZQbENBZz09';
const CONSULTING_FESTA_DETAIL =
  'https://docs.google.com/document/d/1IKKQP_77raPiTC2fyCp2DJUHKWii_-fdQFoCyoV7DR0/edit';

const CONSULTING_CONTENT_RULE = SpreadsheetApp.newDataValidation()
  .requireValueInList(CONSULTING_CONTENT)
  .build();
