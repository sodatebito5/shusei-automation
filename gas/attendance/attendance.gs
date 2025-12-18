// ===============================
// 設定
// ===============================

// 1つのスプレッドシートを共用
const SHEET_ID = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';

// 出欠用
const ATTEND_SHEET_NAME    = '出欠状況（自動）';      // 出欠履歴を書き込むシート
const ATTEND_EVENT_KEY_DEF = '2025年12月例会';        // 出欠用デフォルトeventKey
const ROSTER_SHEET_NAME    = '福岡飯塚_参加者名簿';  // 名簿シート

// ゲスト用
const GUEST_SHEET_NAME     = 'ゲスト出欠状況（自動）'; // ゲスト申請シート名

// LINE Messaging API 用
const CHANNEL_ACCESS_TOKEN = 'h0EwnRvQt+stn4OpyTv12UdZCpYa+KOm736YQuULhuygATdHdXaGmXqwLben8m9TxPnT5UZ59Uzd3gchFemLEmbFXHuaF5TRo44nZV+Qvs36njrFWUxfqhf7zoQTxOCHfpOUofjisza9VwhN+ZzNoAdB04t89/1O/w1cDnyilFU=';
const OWNER_USER_ID = 'U9e236db4178e6dd6a11ec761b0612a73';



// ===============================
// ★ エンドポイント：doPost
// ===============================
function doPost(e) {
  const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  const mode = String(data.mode || 'attend');

  if (mode === 'guest') {
    return handleGuestPost(data);
  } else {
    return handleAttendPost(data);
  }
}



// ===============================
// ★ エンドポイント：doGet
// ===============================
function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || '';

  if (mode === 'seatParticipants') {
    return handleSeatGetParticipants_();
  }
  
  if (mode === 'getTitle') {
    return handleGetTitle_();
  }
  
  if (mode === 'getAttendanceStatus') {
    return getAttendanceStatus(e.parameter.userId);
  }

  return _out({ success: false, error: 'invalid mode' });
}



// ===============================
// ■ 出欠確認：登録処理
// ===============================
function handleAttendPost(data) {
  const userId = String(data.userId || '').trim();
  const name   = String(data.displayName || '').trim();
  const status = String(data.status || '').trim();
  const booth  = String(data.booth || '').trim();

  const boothMark = (booth === '出店したいです。') ? '○' : '×';

  if (!userId) {
    return _out({ success:false, error:'userId is required' });
  }

  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ATTEND_SHEET_NAME);
  if (!sh) {
    return _out({ success:false, error:'出欠シートが見つかりません: ' + ATTEND_SHEET_NAME });
  }

  const lastRow = sh.getLastRow();
  let targetRow = -1;
  
  if (lastRow >= 3) {
    for (let row = 3; row <= lastRow; row++) {
      const cellUserId = String(sh.getRange(row, 3).getValue() || '').trim();
      if (cellUserId === userId) {
        targetRow = row;
        break;
      }
    }
  }

  const eventKey = String(data.eventKey || ATTEND_EVENT_KEY_DEF);
  const rowData = [new Date(), eventKey, userId, name, status, boothMark];

  let action;
  if (targetRow > 0) {
    sh.getRange(targetRow, 1, 1, 6).setValues([rowData]);
    action = 'updated';
  } else {
    sh.appendRow(rowData);
    action = 'inserted';
  }

  try {
    syncAttendanceByLineId();
  } catch (err) {
    Logger.log('sync error: ' + err);
  }

  try {
    const msg = 
      `【出欠登録を受付けました】\n` +
      `例会: ${eventKey}\n` +
      `出欠: ${status}\n` +
      `ブース出店: ${boothMark}`;
    
    pushLineMessage(userId, msg);
  } catch (err) {
    Logger.log('LINE push error: ' + err);
  }

  return _out({ success:true, mode:'attend', action, targetRow });
}



// ===============================
// ■ 出欠確認：状態取得
// ===============================
function getAttendanceStatus(userId) {
  if (!userId) {
    return _out({ success: false, error: 'userId is required' });
  }
  
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ATTEND_SHEET_NAME);
  if (!sh) {
    return _out({ success: false, error: '出欠シートが見つかりません' });
  }
  
  const data = sh.getDataRange().getValues();
  
  // 3行目から検索
  for (let i = 2; i < data.length; i++) {
    const rowUserId = String(data[i][2] || '').trim(); // C列
    if (rowUserId === userId) {
      const status = String(data[i][4] || '').trim(); // E列
      return _out({ success: true, status: status });
    }
  }
  
  // 未登録
  return _out({ success: true, status: '未回答' });
}



// ===============================
// ■ 出欠確認：名簿へ同期
// ===============================
function syncAttendanceByLineId() {
  const ss     = SpreadsheetApp.openById(SHEET_ID);
  const master = ss.getSheetByName(ATTEND_SHEET_NAME);
  const roster = ss.getSheetByName(ROSTER_SHEET_NAME);

  if (!master || !roster) {
    throw new Error('シート名を確認してください');
  }

  const mVals = master.getDataRange().getValues();
  if (mVals.length < 2) return;

  const mHeaderInfo = findHeaderRowAndCols(mVals, ['userId','status']);
  const mHeaderRow  = mHeaderInfo.row;
  const mIdx        = mHeaderInfo.indexMap;

  const latest = new Map();
  for (let r = mVals.length - 1; r > mHeaderRow; r--) {
    const row = mVals[r];
    const uid = String(row[mIdx.userId] || '').trim();
    if (!uid) continue;

    if (!latest.has(uid)) {
      const raw = String(row[mIdx.status] || '').trim();
      latest.set(uid, mapToMark(raw));
    }
  }

  const rVals = roster.getDataRange().getValues();
  if (rVals.length === 0) return;

  let rHeaderRow = -1, lineIdCol = -1, statusCol = -1;
  for (let r = 0; r < Math.min(10, rVals.length); r++) {
    const cols = rVals[r].map(v => String(v || '').trim());
    const li = cols.indexOf('LINE_ID');
    if (li >= 0) {
      rHeaderRow = r;
      lineIdCol  = li;
      statusCol  = cols.indexOf('出欠');
      break;
    }
  }

  if (rHeaderRow < 0 || lineIdCol < 0) {
    throw new Error(`「${ROSTER_SHEET_NAME}」に LINE_ID の見出しが見つかりません`);
  }

  if (statusCol < 0) {
    roster.insertColumnAfter(lineIdCol + 1);
    statusCol = lineIdCol + 1;
    roster.getRange(rHeaderRow + 1, statusCol + 1).setValue('出欠');
  }

  const lastRow = roster.getLastRow();
  if (lastRow <= rHeaderRow + 1) return;

  const idRange    = roster.getRange(rHeaderRow + 2, lineIdCol + 1, lastRow - rHeaderRow - 1, 1);
  const ids        = idRange.getValues();
  const writeRange = roster.getRange(rHeaderRow + 2, statusCol + 1, ids.length, 1);

  const out = new Array(ids.length).fill(0).map(_ => ['']);

  for (let i = 0; i < ids.length; i++) {
    const uid = String(ids[i][0] || '').trim();
    if (uid && latest.has(uid)) {
      out[i][0] = latest.get(uid);
    }
  }

  writeRange.setValues(out);
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
// ■ 配席：参加者一覧取得（seatParticipants）
// ===============================
function handleSeatGetParticipants_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('自動配席用');  // 参加者一覧のシート名

  if (!sh) {
    return _out({
      success: false,
      error: 'sheet not found: 自動配席用'
    });
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return _out({
      success: true,
      participants: []
    });
  }

  const [header, ...rows] = values;

  // 見出し名 → 列インデックス
  const idx = {};
  header.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) idx[key] = i;
  });
  const col = (name) => (typeof idx[name] === 'number' ? idx[name] : -1);

  const colId    = col('ID');
  const colName  = col('氏名');
  const colType  = col('区分');
  const colAff   = col('所属');
  const colRole  = col('役割');
  const colTeam  = col('チーム');
  const colBiz   = col('営業内容');
  const colTable = col('卓');
  const colSeat  = col('席');

  const participants = rows
    .filter(r => String(colName >= 0 ? (r[colName] || '') : '').trim() !== '')
    .map(r => ({
      id:          colId    >= 0 ? String(r[colId]    || '') : '',
      name:        colName  >= 0 ? String(r[colName]  || '') : '',
      category:    colType  >= 0 ? String(r[colType]  || '') : '',
      affiliation: colAff   >= 0 ? String(r[colAff]   || '') : '',
      role:        colRole  >= 0 ? String(r[colRole]  || '') : '',
      team:        colTeam  >= 0 ? String(r[colTeam]  || '') : '',
      business:    colBiz   >= 0 ? String(r[colBiz]   || '') : '',
      table:       colTable >= 0 ? String(r[colTable] || '') : '',
      seat:        colSeat  >= 0 ? String(r[colSeat]  || '') : ''
    }));

  return _out({
    success: true,
    participants: participants
  });
}



// ===============================
// ■ 配席：タイトル取得（getTitle）
// ===============================
function handleGetTitle_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('他会場/ゲスト_参加者名簿（自動）'); // B2セルを読むシート

  if (!sh) {
    return _out({
      success: false,
      error: 'sheet not found: 他会場/ゲスト_参加者名簿（自動）'
    });
  }

  const title = sh.getRange('B2').getValue();

  return _out({
    success: true,
    title: String(title || '')
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

function findHeaderRowAndCols(values2d, requiredKeys) {
  const maxScan = Math.min(10, values2d.length);

  for (let r = 0; r < maxScan; r++) {
    const row   = values2d[r].map(v => String(v || '').trim());
    const lower = row.map(c => c.toLowerCase());
    let ok      = true;
    const indexMap = {};

    for (const key of requiredKeys) {
      const idx = lower.indexOf(key.toLowerCase());
      if (idx < 0) { ok = false; break; }
      indexMap[key] = idx;
    }

    if (ok) {
      return { row: r, indexMap };
    }
  }

  throw new Error(`ヘッダー行が見つかりません（必須: ${requiredKeys.join(', ')}）`);
}

function mapToMark(s) {
  if (s === '出席') return '○';
  if (s === '欠席') return '×';
  return '';
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

// ===============================
// ■ 未回答者リスト作成（LINE未登録も含める）
// ===============================
function updateUnregisteredList() {
  const ss     = SpreadsheetApp.openById(SHEET_ID);
  const master = ss.getSheetByName('会員名簿マスター');
  const attend = ss.getSheetByName(ATTEND_SHEET_NAME);
  
  if (!master || !attend) {
    throw new Error('シートが見つかりません（会員名簿マスター or 出欠状況（自動））');
  }

  // ===== 1. 出欠シート側：回答済み userId セットを作成 =====
  const aLastRow = attend.getLastRow();
  const answeredUserIds = new Set();

  if (aLastRow >= 4) {
    const aRange = attend.getRange(4, 3, aLastRow - 3, 1); // C列 = userId
    const aVals  = aRange.getValues();

    for (let i = 0; i < aVals.length; i++) {
      const uid = String(aVals[i][0] || '').trim();
      if (uid) answeredUserIds.add(uid);
    }
  }

  // ===== 2. 名簿マスター → 未回答者抽出 =====
  // C列：氏名
  // P列：LINE_userId
  const mLastRow = master.getLastRow();
  const unregisteredRows = [];

  if (mLastRow >= 3) {
    const nameRange = master.getRange(3, 3,  mLastRow - 2, 1);  // C列
    const idRange   = master.getRange(3, 16, mLastRow - 2, 1); // P列
    const nameVals  = nameRange.getValues();
    const idVals    = idRange.getValues();

    for (let i = 0; i < nameVals.length; i++) {
      const name = String(nameVals[i][0] || '').trim();
      const uid  = String(idVals[i][0]   || '').trim();

      if (!name) continue;

      // ▼ 判定ロジック
      if (uid && answeredUserIds.has(uid)) {
        // 回答済み → 何もしない
        continue;
      }

      if (uid) {
        // LINEあり ＆ 未回答
        unregisteredRows.push([name, uid, '未回答（LINEあり）']);
      } else {
        // LINE未登録
        unregisteredRows.push([name, '', 'LINE未登録']);
      }
    }
  }

  // ===== 3. 未回答者リスト シートを毎回作り直す =====
  const sheetName = '未回答者リスト';
  let outSheet = ss.getSheetByName(sheetName);
  if (outSheet) {
    ss.deleteSheet(outSheet);
  }
  outSheet = ss.insertSheet(sheetName);

  // ヘッダー
  outSheet.appendRow(['氏名', 'LINE_userId', '区分']);

  if (unregisteredRows.length > 0) {
    outSheet
      .getRange(2, 1, unregisteredRows.length, 3)
      .setValues(unregisteredRows);
  }

  SpreadsheetApp.getUi().alert(
    `未回答者リストを更新しました\n` +
    `回答済み：${answeredUserIds.size}名\n` +
    `未回答＋LINE未登録：${unregisteredRows.length}名`
  );
}




// ===============================
// ★ onOpen：メニュー追加
// ===============================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('管理者専用')
    .addItem('式次第を更新', 'updateMeetingDate')
    .addItem('未回答者は誰だ', 'updateUnregisteredList')
    .addToUi();
}