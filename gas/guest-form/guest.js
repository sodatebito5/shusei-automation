// ===============================
// 設定（ゲスト申請専用）
// ===============================

// 共用スプレッドシート
const SHEET_ID = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';

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
  try {
    const body = e.postData.contents;
    const data = JSON.parse(body);
    return handleGuestPost(data);
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return _out({ success: false, error: 'リクエスト解析エラー: ' + err.message });
  }
}



// ===============================
// ■ ゲスト申請：登録処理
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

  // ★名刺表面は必須チェック
  if (!data.cardFront) {
    return _out({ success:false, error:'名刺画像（表面）は必須です' });
  }

  const now = new Date();
  const nextGuestId = getNextGuestId(sheet);
  const guestName = String(data.name || '');

  // ★名刺画像をDriveに保存
  let cardFrontUrl = '';
  let cardBackUrl = '';

  try {
    if (data.cardFront) {
      cardFrontUrl = saveBase64ToDrive(data.cardFront, `Guest_${nextGuestId}_${guestName}_名刺_表`);
    }
    if (data.cardBack) {
      cardBackUrl = saveBase64ToDrive(data.cardBack, `Guest_${nextGuestId}_${guestName}_名刺_裏`);
    }
  } catch (err) {
    Logger.log('Drive save error: ' + err);
    return _out({ success:false, error:'画像保存に失敗しました: ' + err.message });
  }

  const rowBeforeH = [
    nextGuestId,
    now,
    String(data.eventKey || ''),
    String(data.name || ''),
    String(data.kana || ''),
    String(data.company || ''),
    String(data.title || '')
  ];

  const cardUrls = [];
if (cardFrontUrl) cardUrls.push(cardFrontUrl);
if (cardBackUrl) cardUrls.push(cardBackUrl);

// I〜J列に書き込む
const rowIJ = [
  String(data.displayName || ''),  // I列
  String(data.business || '')      // J列
];

// M〜N列に書き込む
const rowMN = [
  cardFrontUrl,                     // M列：名刺表面URL
  cardBackUrl                       // N列：名刺裏面URL
];

const targetRow = findFirstEmptyRowFrom3(sheet, 1);
const maxRows = sheet.getMaxRows();

if (targetRow > maxRows) {
  sheet.getRange(maxRows + 1, 1, 1, 7).setValues([rowBeforeH]);
  sheet.getRange(maxRows + 1, 9, 1, 2).setValues([rowIJ]);   // I〜J
  sheet.getRange(maxRows + 1, 13, 1, 2).setValues([rowMN]);  // M〜N
} else {
  sheet.getRange(targetRow, 1, 1, 7).setValues([rowBeforeH]);
  sheet.getRange(targetRow, 9, 1, 2).setValues([rowIJ]);     // I〜J
  sheet.getRange(targetRow, 13, 1, 2).setValues([rowMN]);    // M〜N
}

  try {
    const msg =
      '【ゲスト申請を登録しました】\n' +
      `Guest_ID : ${nextGuestId}\n` +
      `例会     : ${data.eventKey || ''}\n` +
      `氏名     : ${data.name || ''}\n` +
      `会社名   : ${data.company || ''}\n` +
      `役職     : ${data.title || ''}\n` +
      `営業内容 : ${data.business || ''}\n` +
      `名刺表   : ${cardFrontUrl ? 'あり' : 'なし'}\n` +
      `名刺裏   : ${cardBackUrl ? 'あり' : 'なし'}`;

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

/**
 * Base64画像をDriveに保存してURLを返す
 */
function saveBase64ToDrive(base64Data, fileName) {
  // data:image/jpeg;base64,xxxxx の形式から分離
  const matches = base64Data.match(/^data:(.+);base64,(.+)$/);
  if (!matches) {
    throw new Error('Invalid base64 format');
  }

  const mimeType = matches[1];  // image/jpeg など
  const base64 = matches[2];
  
  // 拡張子を決定
  let ext = 'jpg';
  if (mimeType.includes('png')) ext = 'png';
  else if (mimeType.includes('gif')) ext = 'gif';
  else if (mimeType.includes('webp')) ext = 'webp';

  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    mimeType,
    `${fileName}.${ext}`
  );

  const folder = DriveApp.getFolderById(CARD_FOLDER_ID);
  const file = folder.createFile(blob);
  
  // 閲覧可能なリンクを返す
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
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
