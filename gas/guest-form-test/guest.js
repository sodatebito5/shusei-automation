// ===============================
// 設定（ゲスト申請専用）
// ===============================

// 共用スプレッドシート
const SHEET_ID = '1vMbqCubLhJvZbli7h8yo_TwdEo5pPvkueLJl0Yqmz90';

// ゲスト用シート
const GUEST_SHEET_NAME = 'ゲスト出欠状況（自動）';

// LINE Messaging API
const CHANNEL_ACCESS_TOKEN =
  '0bpbYqJNtUI6f8Xbo0ZpgNm43ptT6rDteTF+JmwHpM0N9RNBc/Hu/oSlSWkObbiD7eA1JgBQYifNnkhkIac5xAHjaakI5DfoM5udktGipdmdXsJmm+Lws6FiLu5w8qKR8FqY/Q0vXH8AkMLS+YNlFQdB04t89/1O/w1cDnyilFU=';
const OWNER_USER_ID = 'U9e236db4178e6dd6a11ec761b0612a73';

// 名刺画像を保存するフォルダID
const CARD_FOLDER_ID = '1krsSv1oAKIyhWm3MY3iekGoUX-_NJ5Lf';


// ===============================
// ★ エンドポイント：doPost（guest専用）
// ===============================
function doPost(e) {
  // テキスト系は JSON でまとめて渡す想定
  const payloadJson = (e && e.parameter && e.parameter.payload) || '{}';
  const data = JSON.parse(payloadJson);

  const files = (e && e.files) || {};  // card_front, card_back が入ってくる

  return handleGuestPost(data, files);
}


// ===============================
// ■ ゲスト申請：登録処理
// ===============================
function handleGuestPost(data, files) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(GUEST_SHEET_NAME);

  if (!sheet) {
    return _out({ success:false, error:'ゲスト用シートが見つかりません: ' + GUEST_SHEET_NAME });
  }

  const requiredFields = ['eventKey', 'name', 'kana', 'company', 'title', 'business'];
  for (const key of requiredFields) {
    const v = String(data[key] || '').trim();
    if (!v) {
      return _out({ success:false, error:`${key} is required` });
    }
  }

  const now = new Date();
  const nextGuestId = getNextGuestId(sheet);

  // ① 先にシートへ行を作る
  const rowBeforeH = [
    nextGuestId,
    now,
    String(data.eventKey || ''),
    String(data.name || ''),
    String(data.kana || ''),
    String(data.company || ''),
    String(data.title || '')
  ];

  const rowAfterH = [
    String(data.displayName || ''),
    String(data.business || '')
  ];

  const targetRow = findFirstEmptyRowFrom3(sheet, 1);
  const maxRows   = sheet.getMaxRows();

  if (targetRow > maxRows) {
    sheet.getRange(maxRows + 1, 1, 1, 7).setValues([rowBeforeH]);
    sheet.getRange(maxRows + 1, 9, 1, 2).setValues([rowAfterH]);
  } else {
    sheet.getRange(targetRow, 1, 1, 7).setValues([rowBeforeH]);
    sheet.getRange(targetRow, 9, 1, 2).setValues([rowAfterH]);
  }

  // 実際に書き込まれた行番号
  const writeRow = (targetRow > maxRows) ? maxRows + 1 : targetRow;

  // ② 名刺画像をドライブに保存
  const folder = DriveApp.getFolderById(CARD_FOLDER_ID);
  let frontUrl = '';
  let backUrl  = '';

  if (files.card_front) {
    const blob = files.card_front;
    blob.setName(`Guest_${nextGuestId}_front_${blob.getName()}`);
    const file = folder.createFile(blob);
    frontUrl = file.getUrl();
  }

  if (files.card_back) {
    const blob = files.card_back;
    blob.setName(`Guest_${nextGuestId}_back_${blob.getName()}`);
    const file = folder.createFile(blob);
    backUrl = file.getUrl();
  }

  // ③ シートにURLを書き込み（例：K列=表, L列=裏）
  if (frontUrl || backUrl) {
    const range  = sheet.getRange(writeRow, 11, 1, 2); // K列=11, L列=12
    range.setValues([[frontUrl, backUrl]]);
  }

  // ④ LINE 通知（管理者向け）
  try {
    const msgLines = [
      '【ゲスト申請を登録しました】',
      `Guest_ID : ${nextGuestId}`,
      `例会     : ${data.eventKey || ''}`,
      `氏名     : ${data.name || ''}`,
      `会社名   : ${data.company || ''}`,
      `役職     : ${data.title || ''}`,
      `営業内容 : ${data.business || ''}`
    ];

    if (frontUrl) msgLines.push(`名刺(表): ${frontUrl}`);
    if (backUrl)  msgLines.push(`名刺(裏): ${backUrl}`);

    const msg = msgLines.join('\n');
    pushLineMessage(OWNER_USER_ID, msg);
  } catch (err) {
    Logger.log('push error: ' + err);
  }

  return _out({
    success: true,
    mode: 'guest',
    message: 'ゲスト申請を登録しました。',
    guestId: nextGuestId
  });
}


// ===============================
// ● ヘルパー関数
// ===============================
function _out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getNextGuestId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  const idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const nums = idValues
    .map(r => Number(r[0]))
    .filter(n => !isNaN(n) && n > 0);

  if (nums.length === 0) return 1;
  return Math.max.apply(null, nums) + 1;
}

function findFirstEmptyRowFrom3(sheet, col) {
  const maxRows = sheet.getMaxRows();
  for (let r = 3; r <= maxRows; r++) {
    const v = sheet.getRange(r, col).getValue();
    if (v === '' || v === null) {
      return r;
    }
  }
  return maxRows + 1;
}

function pushLineMessage(toUserId, messageText) {
  if (!CHANNEL_ACCESS_TOKEN) {
    Logger.log('no CHANNEL_ACCESS_TOKEN');
    return;
  }
  if (!toUserId) {
    Logger.log('no toUserId');
    return;
  }

  const url = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: toUserId,
    messages: [{ type: 'text', text: messageText }]
  };

  const params = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, params);
  Logger.log('LINE push status: ' + res.getResponseCode());
  Logger.log('LINE push body  : ' + res.getContentText());
}
