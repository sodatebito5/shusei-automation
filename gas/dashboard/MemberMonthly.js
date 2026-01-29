/** =========================================================
 *  会員別月次サマリー
 *  - シート自動生成
 *  - 出欠・売上・ゲストを集計してシートに書き込む
 * ======================================================= */

const MEMBER_MONTHLY_SHEET_NAME = '会員別月次サマリー';

/**
 * 会員別月次サマリーシートを作成（なければ新規、あればスキップ）
 */
function createMemberMonthlySheet() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  let sh = ss.getSheetByName(MEMBER_MONTHLY_SHEET_NAME);
  
  if (sh) {
    Logger.log('シート既存: ' + MEMBER_MONTHLY_SHEET_NAME);
    return { created: false, message: 'シートは既に存在します' };
  }
  
  // 新規作成
  sh = ss.insertSheet(MEMBER_MONTHLY_SHEET_NAME);
  
  // ヘッダー設定
  const headers = [
    'eventKey',
    'userId', 
    '氏名',
    '出欠',
    '商談件数',
    '成約件数',
    '売上金額',
    'ゲスト招待',
    '更新日時'
  ];
  
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  const headerRange = sh.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#e5e7eb');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅調整
  sh.setColumnWidth(1, 140); // eventKey
  sh.setColumnWidth(2, 180); // userId
  sh.setColumnWidth(3, 100); // 氏名
  sh.setColumnWidth(4, 60);  // 出欠
  sh.setColumnWidth(5, 80);  // 商談件数
  sh.setColumnWidth(6, 80);  // 成約件数
  sh.setColumnWidth(7, 100); // 売上金額
  sh.setColumnWidth(8, 80);  // ゲスト招待
  sh.setColumnWidth(9, 140); // 更新日時
  
  // 1行目を固定
  sh.setFrozenRows(1);
  
  Logger.log('シート作成完了: ' + MEMBER_MONTHLY_SHEET_NAME);
  return { created: true, message: 'シートを作成しました' };
}

/**
 * 指定例会の会員別月次データを集計してシートに書き込む
 * @param {string} eventName - 例会名（例: 2025年12月例会）省略時は現在例会
 */
function updateMemberMonthly(eventName) {
  // 例会名が指定されていなければ現在例会を使用
  if (!eventName) {
    const eventInfo = getCurrentEventInfo_();
    eventName = eventInfo.name;
  }
  
  if (!eventName) {
    Logger.log('例会名が取得できません');
    return { success: false, message: '例会名が取得できません' };
  }
  
  // シート存在確認（なければ作成）
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  let sh = ss.getSheetByName(MEMBER_MONTHLY_SHEET_NAME);
  if (!sh) {
    createMemberMonthlySheet();
    sh = ss.getSheetByName(MEMBER_MONTHLY_SHEET_NAME);
  }
  
  // 集計実行
  const data = aggregateMemberMonthly_(eventName);
  
  if (!data.length) {
    Logger.log('集計データがありません: ' + eventName);
    return { success: false, message: '集計データがありません' };
  }
  
  // 既存データ削除（同じeventKeyの行を消す）
  deleteMemberMonthlyRows_(sh, eventName);
  
  // 新規データ追記
  const lastRow = sh.getLastRow();
  sh.getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`${eventName}: ${data.length}件 書き込み完了`);
  return { success: true, message: `${data.length}件 書き込み完了` };
}

/**
 * 同じeventKeyの既存行を削除
 */
function deleteMemberMonthlyRows_(sh, eventName) {
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return; // ヘッダーのみ
  
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues(); // A列のみ
  
  // 下から削除（行番号ズレ防止）
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === eventName) {
      sh.deleteRow(i + 2); // +2 はヘッダー分
    }
  }
}

/**
 * 会員別月次データを集計（内部関数）
 */
function aggregateMemberMonthly_(eventName) {
  const now = new Date();
  
  // 1. 全会員リスト取得
  const members = getAllMembersWithUserId_();
  if (!members.length) return [];
  
  // 2. 出欠マップ作成 { name: '○' | '×' | '未回答' }
  const attendMap = getAttendMapByEvent_(eventName);
  
  // 3. 売上マップ作成 { name: { meetings, deals, amount } }
  const salesMap = getSalesMapByEvent_(eventName);
  
  // 4. ゲスト招待マップ作成 { name: count }
  const guestMap = getGuestCountMapByEvent_(eventName);
  
  // 5. 全会員分のデータ作成
  const result = [];
  
  members.forEach(m => {
    const name = m.name;
    const userId = m.userId || '';
    
    const attend = attendMap[name] || '未回答';
    const sales = salesMap[name] || { meetings: 0, deals: 0, amount: 0 };
    const guestCount = guestMap[name] || 0;
    
    result.push([
      eventName,                    // A: eventKey
      userId,                       // B: userId
      name,                         // C: 氏名
      attend,                       // D: 出欠
      sales.meetings,               // E: 商談件数
      sales.deals,                  // F: 成約件数
      sales.amount,                 // G: 売上金額
      guestCount,                   // H: ゲスト招待
      now                           // I: 更新日時
    ]);
  });
  
  return result;
}

/**
 * 会員名簿から全会員（userId付き）を取得
 */
function getAllMembersWithUserId_() {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sh) return [];
  
  const values = sh.getDataRange().getValues();
  if (values.length < 3) return [];
  
  const header = values[1]; // 2行目がヘッダー
  const rows = values.slice(2); // 3行目以降がデータ
  
  const idxUserId = findColumnIndex_(header, ['LINE_userId', 'userId', 'ユーザーID']);
  const idxName = findColumnIndex_(header, ['displayName', '氏名', '名前']);
  const nameCol = idxName !== -1 ? idxName : 2; // デフォルトC列
  
  const list = [];
  const seen = {};
  
  rows.forEach(row => {
    const name = String(row[nameCol] || '').trim();
    if (!name || seen[name]) return;
    
    const userId = idxUserId !== -1 ? String(row[idxUserId] || '').trim() : '';
    list.push({ name, userId });
    seen[name] = true;
  });
  
  return list;
}

/**
 * 出欠マップ取得 { 氏名: '○' | '×' | '未回答' }
 */
function getAttendMapByEvent_(eventName) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(ATTEND_SHEET_NAME);
  if (!sh) return {};
  
  const values = sh.getDataRange().getValues();
  if (values.length < 4) return {};
  
  const HEADER_ROW_INDEX = 2;
  const header = values[HEADER_ROW_INDEX];
  const rows = values.slice(HEADER_ROW_INDEX + 1);
  
  const idxEventKey = findColumnIndex_(header, ['eventKey', 'イベントキー']);
  const idxStatus = findColumnIndex_(header, ['status', '出欠', '出欠状況', 'ステータス']);
  const idxName = findColumnIndex_(header, ['displayName', '氏名', '名前']);
  
  if (idxEventKey === -1 || idxStatus === -1 || idxName === -1) return {};
  
  const map = {};
  
  rows.forEach(row => {
    const key = String(row[idxEventKey] || '').trim();
    if (key !== eventName) return;
    
    const name = String(row[idxName] || '').trim();
    const st = String(row[idxStatus] || '').trim();
    
    if (!name) return;
    
    if (st === '○' || st === '出席') {
      map[name] = '○';
    } else if (st === '×' || st === '欠席') {
      map[name] = '×';
    }
  });
  
  return map;
}

/**
 * 売上マップ取得 { 氏名: { meetings, deals, amount } }
 */
function getSalesMapByEvent_(eventName) {
  const ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  const sh = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sh) return {};
  
  const lastRow = sh.getLastRow();
  if (lastRow < 7) return {};
  
  // eventName（2025年1月例会）→ 前月の売上eventKey（12月_売上報告）
  const monthMatch = eventName.match(/(\d{4})年(\d{1,2})月/);
  if (!monthMatch) return {};
  
  const year = Number(monthMatch[1]);
  const month = Number(monthMatch[2]);
  const prevDate = new Date(year, month - 2, 1);
  const prevMonth = prevDate.getMonth() + 1;
  const salesEventKey = `${prevMonth}月_売上報告`;
  
  const values = sh.getRange(7, 1, lastRow - 6, sh.getLastColumn()).getValues();
  
  const map = {};
  
  values.forEach(row => {
    const rowEventKey = String(row[14] || '').trim(); // O列
    if (rowEventKey !== salesEventKey) return;
    
    const member = String(row[13] || '').trim(); // N列：氏名
    if (!member) return;
    
    const meetings = Number(row[15]) || 0; // P列：商談件数
    const deals = Number(row[4]) || 0;     // E列：成約件数
    const amount = Number(row[5]) || 0;    // F列：売上金額
    
    // 同一会員の複数報告は合算
    if (!map[member]) {
      map[member] = { meetings: 0, deals: 0, amount: 0 };
    }
    map[member].meetings += meetings;
    map[member].deals += deals;
    map[member].amount += amount;
  });
  
  return map;
}

/**
 * ゲスト招待数マップ取得 { 紹介者名: count }
 */
function getGuestCountMapByEvent_(eventName) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(GUEST_SHEET_NAME);
  if (!sh) return {};
  
  const values = sh.getDataRange().getValues();
  if (values.length < 3) return {};
  
  const HEADER_ROW_INDEX = 1; // 2行目がヘッダー
  const header = values[HEADER_ROW_INDEX];
  const rows = values.slice(HEADER_ROW_INDEX + 1);
  
  const idxEvent = findColumnIndex_(header, ['eventKey', 'イベントキー']);
  const idxIntro = findColumnIndex_(header, ['紹介者', '紹介者名']);
  
  const cEvent = idxEvent !== -1 ? idxEvent : 2;
  const cIntro = idxIntro !== -1 ? idxIntro : 7;
  
  const map = {};
  
  rows.forEach(row => {
    const key = String(row[cEvent] || '').trim();
    if (key !== eventName) return;
    
    const introName = String(row[cIntro] || '').trim();
    if (!introName) return;
    
    map[introName] = (map[introName] || 0) + 1;
  });
  
  return map;
}

/** =========================================================
 *  テスト用
 * ======================================================= */

/**
 * テスト：現在例会で実行
 */
function testUpdateMemberMonthly() {
  const result = updateMemberMonthly();
  Logger.log(result);
}

/**
 * テスト：シート作成のみ
 */
function testCreateMemberMonthlySheet() {
  const result = createMemberMonthlySheet();
  Logger.log(result);
}