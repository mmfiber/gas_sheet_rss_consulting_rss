function sendRemidMail() {
  const infoRss = sendEmail('rss');
  const infoStudent = sendEmail('student');
  send(
    'e19m5207@soka-u.jp',
    'remid notification',
    JSON.stringify({ infoRss, infoStudent })
  );
}

function sendEmail(type) {
  const now = new Date();
  const tomorrow = dateToStr(new Date(now.setDate(now.getDate() + 1)));

  let sentInfo = {};
  INDUSTRIES.forEach((name) => {
    sentInfo[name] = [];
    const sheetData = ssData.getSheetByName(name);
    const lastRowIdx = sheetData.getLastRow() - 1;
    if (lastRowIdx === 0) return;

    const ids = sheetData.getRange(2, 1, lastRowIdx, 1).getValues();
    const matchedIds = ids.filter((id) => id[0].slice(0, 8) === tomorrow);
    matchedIds.forEach((id) => {
      const rowIdxData = getRowIdxData(sheetData, id[0]);

      const studentId = sheetData.getRange(rowIdxData, 2).getValue().toString();
      if (
        type === 'student' &&
        (studentId.length !== 7 || !validateStudentId(studentId))
      )
        return;

      const rssId = sheetData.getRange(rowIdxData, 4).getValue().toString();
      if (type === 'rss' && (rssId.length !== 7 || !validateStudentId(rssId)))
        return;

      const studentName = sheetData.getRange(rowIdxData, 3).getValue();
      const rssName = sheetData.getRange(rowIdxData, 5).getValue();
      const date = sheetData.getRange(rowIdxData, 6).getValue();
      const period = sheetData.getRange(rowIdxData, 7).getValue();
      const session = sheetData.getRange(rowIdxData, 8).getValue();
      const content = sheetData.getRange(rowIdxData, 11).getValue();

      const body =
        type === 'rss'
          ? getBodyMsgRss(
              rssName,
              studentName,
              date,
              period,
              session,
              content,
              name
            )
          : getBodyMsgStudent(
              rssName,
              studentName,
              date,
              period,
              session,
              content,
              name
            );
      const email = getEmail(type === 'rss' ? rssId : studentId);
      send(email, '就活相談会リマインド', body);

      const info =
        type === 'rss'
          ? `${email}, ${rssId}, ${rssName}`
          : `${email}, ${studentId}, ${studentName}`;
      sentInfo[name].push(info);
    });
  });
  return sentInfo;
}

function send(to, subject, body) {
  GmailApp.sendEmail(to, subject, body, {
    name: 'RSS 就活相談会',
    noReply: true,
  });
}

function getBodyMsgRss(
  rssName,
  studentName,
  date,
  period,
  session,
  content,
  insdustry
) {
  return `
    ${rssName} さん

    こんにちは！ 
    日頃から就活生の使命の進路決着のために尽力いただきありがとうございます！ 
      
    就活相談会の詳細を連絡させて頂きます。 
    ------------------------------------------------------------------- 
    ・日時：${date.getMonth() + 1}月${date.getDate()}日 ${TIMETABLE[period - 1]}
    ・業界：${insdustry}
    ・セッション番号: ${session}

    ・担当学生：${studentName}
    ・相談内容：${content}
    
    ※5 分前集合を心がけてください 
    ・開催形態：ZOOM 
    ・URL：${ZOOM_URL}
    ------------------------------------------------------------------- 

    注意点 
    ・予約されていなくても必ず予定を開けといてください。直前まで予約が可能です。
    ・遅刻、欠席は厳禁となっています。万が一、当日遅刻・やむを得ない事情で欠席する場合は、ホストまでメールにて連絡してください。
    ・業務内容を下記URLから確認してください。
    ${CONSULTING_FESTA_DETAIL}

    最後まで後輩のために、一緒に頑張りましょう！ 当日お会いできる事を楽しみにしています。
  `.replace('\t', '');
}

function getBodyMsgStudent(
  rssName,
  studentName,
  date,
  period,
  session,
  content,
  insdustry
) {
  return `
    ${studentName} さん

    こんにちは！ 
    就活相談会へ参加いただきありがとうございます！ 
      
    就活相談会の詳細を連絡させて頂きます。 
    ------------------------------------------------------------------- 
    ・日時：${date.getMonth() + 1}月${date.getDate()}日 ${TIMETABLE[period - 1]}
    ・業界：${insdustry}
    ・セッション番号: ${session}
    ・担当RSS：${rssName}
    ・服装：自由
    
    ※5 分前集合を心がけてください 
    ・開催形態：ZOOM 
    ・URL：${ZOOM_URL}
    ------------------------------------------------------------------- 
    アドバイス 
    ・RSS に聞きたい就職活動への疑問や不安は、事前にまとめておいて下さい。 
    注意点 
    ・遅刻、欠席の際はrss17hr@gmail.comまでメールにて連絡お願いします。

    進路実現に向かって、一緒に頑張りましょう！ 当日お会いできる事を楽しみにしています。
  `.replace('\t', '');
}
