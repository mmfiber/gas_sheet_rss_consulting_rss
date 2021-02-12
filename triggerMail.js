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
  const tomorrow = new Date(now.setDate(now.getDate() + 1));
  const data = fetchDataByDate(tomorrow);

  let sentInfo = {};
  for (const prop in data) {
    sentInfo[prop] = [];
    data[prop].forEach((d) => {
      if (
        type === 'student' &&
        (d.studentId.length !== 7 || !validateStudentId(d.studentId))
      )
        return;

      if (
        type === 'rss' &&
        (d.rssId.length !== 7 || !validateStudentId(d.rssId))
      )
        return;

      const body = type === 'rss' ? getBodyMsgRss(d) : getBodyMsgStudent(d);
      if (!body) return;

      const email = type === 'rss' ? d.rssEmail : d.studentEmail;
      send(email, '就活相談会リマインド', body);

      const info =
        type === 'rss'
          ? `${email}, ${d.rssName}`
          : `${email}, ${d.studentName}`;
      sentInfo[prop].push(info);
    });
  }
  return sentInfo;
}

function send(to, subject, body) {
  GmailApp.sendEmail(to, subject, body, {
    name: 'RSS 就活相談会',
    noReply: true,
  });
}

function getBodyMsgRss(info) {
  try {
    return `
      ${info.rssName} さん

      こんにちは！
      日頃から就活生の使命の進路決着のために尽力いただきありがとうございます！ 

      就活相談会の詳細を連絡させて頂きます。 
      ------------------------------------------------------------------- 
      ・日時：${info.date.getMonth() + 1}月${info.date.getDate()}日 ${
      TIMETABLE[info.period - 1]
    }
      ・業界：${info.industry}
      ・セッション番号: ${info.session}

      ・担当学生：${info.studentName}
      ・相談内容：${info.content}

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
    `
      .replace(/^\n/, '')
      .replace(/^ {6}/gm, '');
  } catch (e) {
    logger('getBodyMsgRss', e);
    return null;
  }
}

function getBodyMsgStudent(info) {
  try {
    return `
      ${info.studentName} さん

      こんにちは！
      就活相談会へ参加いただきありがとうございます！ 

      就活相談会の詳細を連絡させて頂きます。 
      ------------------------------------------------------------------- 
      ・日時：${info.date.getMonth() + 1}月${info.date.getDate()}日 ${
      TIMETABLE[info.period - 1]
    }
      ・業界：${info.industry}
      ・セッション番号: ${info.session}
      ・担当RSS：${info.rssName}
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
    `
      .replace(/^\n/, '')
      .replace(/^ {6}/gm, '');
  } catch (e) {
    logger('getBodyMsgRss', e);
    return null;
  }
}

function testMail() {
  const studentId = '11111111';
  const rssId = '1111111';
  const studentName = 'STUDENT';
  const rssName = 'RSS';
  const date = new Date();
  const period = 1;
  const session = 1;
  const content = 'test';
  const rssEmail = getEmail(rssId);
  const studentEmail = getEmail(studentId);
  const industry = 'test';

  const obj = {
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
  };

  // send(
  //   'e19m5207@soka-u.jp',
  //   'test',
  //   getBodyMsgStudent(obj)
  // );
  console.log(getBodyMsgRss(obj));
  console.log(getBodyMsgStudent(obj));
}
