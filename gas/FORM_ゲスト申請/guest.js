// ===============================
// 設定（ゲスト申請専用）
// ===============================

// 共用スプレッドシート
const SHEET_ID = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';

// ゲスト用シート
const GUEST_SHEET_NAME = 'ゲスト出欠状況（自動）';

// LINE Messaging API
const CHANNEL_ACCESS_TOKEN = 'h0EwnRvQt+stn4OpyTv12UdZCpYa+KOm736YQuULhuygATdHdXaGmXqwLben8m9TxPnT5UZ59Uzd3gchFemLEmbFXHuaF5TRo44nZV+Qvs36njrFWUxfqhf7zoQTxOCHfpOUofjisza9VwhN+ZzNoAdB04t89/1O/w1cDnyilFU=';
const OWNER_USER_ID = 'U9e236db4178e6dd6a11ec761b0612a73';


// ===============================
// ★ エンドポイント：doPost（guest専用）
// ===============================
function doPost(e) {
  const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  return handleGuestPost(data);
}


// ===============================
// ■ ゲスト申請：登録処理（現状と同じロジック）
// ===============================
function handleGuestPost(data) {
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
  const maxRows = sheet.getMaxRows();

  if (targetRow > maxRows) {
    sheet.getRange(maxRows + 1, 1, 1, 7).setValues([rowBeforeH]);
    sheet.getRange(maxRows + 1, 9, 1, 2).setValues([rowAfterH]);
  } else {
    sheet.getRange(targetRow, 1, 1, 7).setValues([rowBeforeH]);
    sheet.getRange(targetRow, 9, 1, 2).setValues([rowAfterH]);
  }

  try {
    const msg =
      '【ゲスト申請を登録しました】\n' +
      `Guest_ID : ${nextGuestId}\n` +
      `例会     : ${data.eventKey || ''}\n` +
      `氏名     : ${data.name || ''}\n` +
      `会社名   : ${data.company || ''}\n` +
      `役職     : ${data.title || ''}\n` +
      `営業内容 : ${data.business || ''}`;

    pushLineMessage(OWNER_USER_ID, msg);
  } catch (err) {
    Logger.log('push error: ' + err);
  }

  return _out({
    success: true,
    mode: 'guest',
    message: 'ゲスト申請を登録しました。'
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
