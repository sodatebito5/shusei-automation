// ===============================
// seat-maker：自動配席アプリ用 API（統合版）
// ===============================

// 参加者一覧を置いているスプレッドシート
const SEAT_SHEET_ID   = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';
const SEAT_SHEET_NAME = '自動配席用';

// 受付名簿（自動）
const MAIN_BOOK_NAME  = '受付名簿（自動）';

// 受付名簿（他会場・ゲスト）
const OTHER_BOOK_NAME = '受付名簿（他会場・ゲスト）';

// 他会場の反映先範囲
const OTHER_RANGE_START = 3;
const OTHER_RANGE_END   = 20;

// ゲストの反映先範囲
const GUEST_RANGE_START = 25;
const GUEST_RANGE_END   = 42;



// ===============================
// タイトル（例会名）取得 API
// mode=getTitle
// ===============================
function handleGetTitle_() {
  const ss = SpreadsheetApp.openById(SEAT_SHEET_ID);
  const sh = ss.getSheetByName('他会場/ゲスト_参加者名簿（自動）');

  if (!sh) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: 'sheet not found: 他会場/ゲスト_参加者名簿（自動）'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const title = sh.getRange('B2').getValue();

  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      title: String(title || '')
    }))
    .setMimeType(ContentService.MimeType.JSON);
}



// ===============================
// 参加者一覧取得（自動配席用）
// mode=seatParticipants
// ===============================
function handleSeatGetParticipants_() {
  const ss = SpreadsheetApp.openById(SEAT_SHEET_ID);
  const sh = ss.getSheetByName(SEAT_SHEET_NAME);

  if (!sh) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: 'sheet not found: ' + SEAT_SHEET_NAME
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        participants: []
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const [header, ...rows] = values;

  const idx = {};
  header.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) idx[key] = i;
  });

  const col = (name) => (typeof idx[name] === 'number' ? idx[name] : -1);

  const colId        = col('ID');
  const colName      = col('氏名');
  const colType      = col('区分');
  const colAff       = col('所属');
  const colRole      = col('役割');
  const colTeam      = col('チーム');
  const colBiz       = col('営業内容');
  const colTable     = col('卓');
  const colSeat      = col('席');

  const participants = rows
    .filter(r => String(colName >= 0 ? (r[colName] || '') : '').trim() !== '')
    .map(r => ({
      id:          colId   >= 0 ? String(r[colId]   || '') : '',
      name:        colName >= 0 ? String(r[colName] || '') : '',
      category:    colType >= 0 ? String(r[colType] || '') : '',
      affiliation: colAff  >= 0 ? String(r[colAff]  || '') : '',
      role:        colRole >= 0 ? String(r[colRole] || '') : '',
      team:        colTeam >= 0 ? String(r[colTeam] || '') : '',
      business:    colBiz  >= 0 ? String(r[colBiz]  || '') : '',
      table:       colTable >= 0 ? String(r[colTable] || '') : '',
      seat:        colSeat  >= 0 ? String(r[colSeat]  || '') : ''
    }));

  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      participants
    }))
    .setMimeType(ContentService.MimeType.JSON);
}



// ===============================
// 座席情報の書き込み（内部処理）
// ===============================
function syncSeatsInternal_(assignments) {
  const ss = SpreadsheetApp.openById(SEAT_SHEET_ID);

  const mainSheet  = ss.getSheetByName(MAIN_BOOK_NAME);
  const otherSheet = ss.getSheetByName(OTHER_BOOK_NAME);

  if (!mainSheet || !otherSheet) {
    return {
      success: false,
      error: '受付名簿シートが不足しています'
    };
  }

  const mainValues  = mainSheet.getDataRange().getValues();
  const otherValues = otherSheet.getDataRange().getValues();

  const mainHeader = mainValues[0];
  const idxMainName = mainHeader.indexOf('氏名');
  const idxMainSeat = mainHeader.indexOf('座席');

  const otherHeader = otherValues[0];
  const idxOtherName = otherHeader.indexOf('氏名');
  const idxOtherSeat = otherHeader.indexOf('座席');

  if (idxMainName < 0 || idxMainSeat < 0 || idxOtherName < 0 || idxOtherSeat < 0) {
    return {
      success: false,
      error: '氏名/座席の見出しが不足しています'
    };
  }

  let updatedMain = 0;
  let updatedOther = 0;
  const unmatched = [];

  for (const a of assignments) {
    const name = String(a.name || '').trim();
    const table = String(a.table || '').trim();
    const category = String(a.category || '').trim();

    if (!name || !table) continue;

    // 会員
    if (category === '会員') {
      let found = false;

      for (let r = 1; r < mainValues.length; r++) {
        const rowName = String(mainValues[r][idxMainName] || '').trim();
        if (rowName === name) {
          mainSheet.getRange(r + 1, idxMainSeat + 1).setValue(table);
          updatedMain++;
          found = true;
          break;
        }
      }

      if (!found) {
        unmatched.push({
          name, table, category, reason: 'not found in 会員受付名簿'
        });
      }

    } else {
      // 他会場・ゲスト
      let found = false;

      for (let r = 1; r < otherValues.length; r++) {
        const rowName = String(otherValues[r][idxOtherName] || '').trim();
        if (rowName === name) {
          otherSheet.getRange(r + 1, idxOtherSeat + 1).setValue(table);
          updatedOther++;
          found = true;
          break;
        }
      }

      if (!found) {
        unmatched.push({
          name, table, category, reason: 'not found in 他会場/ゲスト受付名簿'
        });
      }
    }
  }

  return {
    success: true,
    updatedMain,
    updatedOther,
    unmatched
  };
}



// ===============================
// doPost：座席反映 API
// mode=syncSeats
// ===============================
function handleSyncSeats_(jsonText) {
  let payload = null;

  try {
    payload = JSON.parse(jsonText);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: 'invalid JSON: ' + err
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (!payload.assignments || !Array.isArray(payload.assignments)) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: 'assignments is required'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const result = syncSeatsInternal_(payload.assignments);

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}



// ===============================
// doPost：モード切り替え
// ===============================
function doPost(e) {
  const mode = (e && e.parameter && e.parameter.mode) || '';

  if (mode === 'syncSeats') {
    const jsonText = (e.postData && e.postData.contents) || '{}';
    return handleSyncSeats_(jsonText);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      success: false,
      error: 'invalid mode (doPost)'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}



// ===============================
// doGet：座席一覧 or タイトル取得
// ===============================
function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || '';

  if (mode === 'seatParticipants') {
    return handleSeatGetParticipants_();
  }

  if (mode === 'getTitle') {
    return handleGetTitle_();
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      success: false,
      error: 'invalid mode'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}



// ===============================
// ★ テスト：内部処理
// ===============================
function test_syncSeatsInternal_() {
  const payload = {
    assignments: [
      { name: '佐藤 太郎', table: 'A',  category: '会員' },
      { name: '鈴木 花子', table: 'B',  category: '他会場' },
      { name: '山木 和夫', table: 'PA', category: '会員' },
      { name: '鈴木 幸則', table: 'MC', category: 'ゲスト' }
    ]
  };

  const result = syncSeatsInternal_(payload.assignments);
  Logger.log(JSON.stringify(result, null, 2));
}



// ===============================
// ★ テスト：doPost 経由
// ===============================
function test_syncSeatsByDoPost_() {
  const e = {
    parameter: { mode: 'syncSeats' },
    postData: {
      type: 'application/json',
      contents: JSON.stringify({
        assignments: [
          { name: '佐藤 太郎', table: 'A',  category: '会員' },
          { name: '鈴木 花子', table: 'B',  category: '他会場' },
          { name: '山木 和夫', table: 'PA', category: '会員' },
          { name: '鈴木 幸則', table: 'MC', category: 'ゲスト' }
        ]
      })
    }
  };

  const res = doPost(e);
  Logger.log(res.getContent());
}
