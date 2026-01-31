/** =========================================================
 *  設定
 * ======================================================= */

// 管理ダッシュボード（設定シートが入っているスプシ）
const CONFIG_SHEET_ID = '1R4GR1GZg6mJP9zPX5MTE0IsYEAdVLNIM314o7vBqrg8';
const CONFIG_SHEET_NAME = '設定';
const CURRENT_EVENT_KEY_CELL = 'A2';  // 202512_01 が入っているセル

// 出欠＆ゲスト用スプシ
const ATTEND_GUEST_SHEET_ID = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';
const ATTEND_SHEET_NAME = '出欠状況（自動）';
const GUEST_SHEET_NAME  = 'ゲスト出欠状況（自動）';

// 会員マスタ
const MEMBER_SHEET_NAME = '会員名簿マスター';

// 売上報告用スプシ
const SALES_SHEET_ID = '1QYLHr7wMj0jQW5ApgWf-l1Bk9_9M5EV54FDHQp2Rkic';
const SALES_SHEET_NAME = 'sales';

// 紹介記録シート名（追加）
const REFERRAL_SHEET_NAME = '紹介記録';

// LINE Messaging API 用
const LINE_CHANNEL_ACCESS_TOKEN = 'h0EwnRvQt+stn4OpyTv12UdZCpYa+KOm736YQuULhuygATdHdXaGmXqwLben8m9TxPnT5UZ59Uzd3gchFemLEmbFXHuaF5TRo44nZV+Qvs36njrFWUxfqhf7zoQTxOCHfpOUofjisza9VwhN+ZzNoAdB04t89/1O/w1cDnyilFU=';
const LINE_MESSAGE_LIMIT = 200;  // LINE無料枠（月200通）

// 役割分担シート名
const ROLE_ASSIGNMENT_SHEET_NAME = '役割分担';

// 参加者名簿アーカイブ用フォルダ（Google Drive）
const ROSTER_ARCHIVE_FOLDER_ID = '1tp3NCGTpLkwpS9zb3c87b8xZ_Bdlule8';

/** =========================================================
 *  ダッシュボード高速化：事前集計シート方式
 * ======================================================= */

const DASHBOARD_CACHE_SHEET_NAME = 'ダッシュボード集計';

/**
 * 集計シートを作成（なければ新規、あればスキップ）
 * ★最初に1回だけ手動実行
 */
function createDashboardCacheSheet() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  let sh = ss.getSheetByName(DASHBOARD_CACHE_SHEET_NAME);
  
  if (sh) {
    Logger.log('シート既存: ' + DASHBOARD_CACHE_SHEET_NAME);
    return;
  }
  
  sh = ss.insertSheet(DASHBOARD_CACHE_SHEET_NAME);
  sh.getRange('A1').setValue('json');
  sh.getRange('B1').setValue('updatedAt');
  sh.setColumnWidth(1, 800);
  
  Logger.log('シート作成完了: ' + DASHBOARD_CACHE_SHEET_NAME);

}

/** =========================================================
 *  配席アーカイブ
 * ======================================================= */

const SEATING_ARCHIVE_SHEET_NAME = '配席アーカイブ';

/**
 * 配席アーカイブシートを作成（なければ新規、あればスキップ）
 * ★最初に1回だけ手動実行
 */
function createSeatingArchiveSheet() {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  let sh = ss.getSheetByName(SEATING_ARCHIVE_SHEET_NAME);

  if (sh) {
    Logger.log('シート既存: ' + SEATING_ARCHIVE_SHEET_NAME);
    return { success: true, message: 'シート既存' };
  }

  sh = ss.insertSheet(SEATING_ARCHIVE_SHEET_NAME);

  const headers = ['例会キー', '例会日', '参加者ID', '氏名', '区分', '所属', '役割', 'チーム', 'テーブル', '席番', '確定日時', '確定者'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sh.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 180);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 80);
  sh.setColumnWidth(6, 100);
  sh.setColumnWidth(7, 120);
  sh.setColumnWidth(8, 60);
  sh.setColumnWidth(9, 80);
  sh.setColumnWidth(10, 60);
  sh.setColumnWidth(11, 150);
  sh.setColumnWidth(12, 180);

  sh.setFrozenRows(1);

  Logger.log('シート作成完了: ' + SEATING_ARCHIVE_SHEET_NAME);
  return { success: true, message: 'シート作成完了' };
}

/**
 * ダッシュボードデータを集計してシートに保存
 * ★トリガーで5分おきに実行
 */
function updateDashboardCache() {
  const data = getDashboardData();
  const json = JSON.stringify(data);
  const now = new Date();
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  let sh = ss.getSheetByName(DASHBOARD_CACHE_SHEET_NAME);
  
  if (!sh) {
    createDashboardCacheSheet();
    sh = ss.getSheetByName(DASHBOARD_CACHE_SHEET_NAME);
  }
  
  sh.getRange('A2').setValue(json);
  sh.getRange('B2').setValue(now);
  
  Logger.log('集計完了: ' + now);
}

/**
 * 集計シートからデータを読み込む（高速版）
 * ★フロントから呼ぶのはこっち
 */
function getDashboardDataFast() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(DASHBOARD_CACHE_SHEET_NAME);
  
  if (!sh) {
    // 集計シートがない場合はフォールバック
    Logger.log('集計シートなし、従来版で取得');
    return getDashboardDataSafe();
  }
  
  const json = sh.getRange('A2').getValue();
  
  if (!json) {
    // データがない場合はフォールバック
    Logger.log('集計データなし、従来版で取得');
    return getDashboardDataSafe();
  }
  
  try {
    return JSON.parse(json);
  } catch (e) {
    Logger.log('JSONパースエラー: ' + e);
    return getDashboardDataSafe();
  }
}


/** =========================================================
 *  イベントキー自動切り替え
 * ======================================================= */

/**
 * 設定シートにF列・G列（次月イベントキー・次月開催日）を追加
 * ★初回セットアップ時に1回だけ実行
 */
function setupEventKeyColumns() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!sh) {
    Logger.log('設定シートが見つかりません');
    return { success: false, error: '設定シートが見つかりません' };
  }

  // 見出し追加
  sh.getRange('F1').setValue('次月イベントキー');
  sh.getRange('G1').setValue('次月開催日');

  // 数式追加
  // F2: 次月イベントキーを計算（G2の日付から）
  sh.getRange('F2').setFormula('=IF(G2="","",TEXT(YEAR(G2),"0000")&TEXT(MONTH(G2),"00")&"_01")');

  // G2: D列から次の開催日を取得
  sh.getRange('G2').setFormula('=IFERROR(INDEX(D:D, MATCH(D2, D:D, 0) + 1), "")');

  Logger.log('F列・G列のセットアップ完了');
  return { success: true, message: 'F列・G列のセットアップ完了' };
}

/**
 * 設定シートからイベント情報を取得
 * @returns {Object} { currentEventKey, currentEventDate, nextEventKey, nextEventDate }
 */
function getEventSettings() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!sh) {
    return { currentEventKey: '', currentEventDate: null, nextEventKey: '', nextEventDate: null };
  }

  const currentEventKey = String(sh.getRange('A2').getValue() || '').trim();
  const currentEventDateRaw = sh.getRange('D2').getValue();
  const nextEventKey = String(sh.getRange('F2').getValue() || '').trim();
  const nextEventDateRaw = sh.getRange('G2').getValue();

  return {
    currentEventKey: currentEventKey,
    currentEventDate: currentEventDateRaw ? new Date(currentEventDateRaw) : null,
    nextEventKey: nextEventKey,
    nextEventDate: nextEventDateRaw ? new Date(nextEventDateRaw) : null
  };
}

/**
 * 指定時刻を過ぎたかチェック
 * @param {Date} eventDate - 基準日
 * @param {number} hour - 時刻（0-23）
 * @param {number} dayOffset - 日付オフセット（-1=前日, 0=当日）
 * @returns {boolean}
 */
function isPastSwitchTime(eventDate, hour, dayOffset) {
  if (!eventDate) return false;

  const now = new Date();
  const switchTime = new Date(eventDate);
  switchTime.setDate(switchTime.getDate() + (dayOffset || 0));
  switchTime.setHours(hour, 0, 0, 0);

  return now >= switchTime;
}

/**
 * ダッシュボード用イベントキー
 * 切り替え: 例会開催日 11:00
 */
function getDashboardEventKey() {
  const settings = getEventSettings();
  if (!settings.currentEventDate) {
    return settings.currentEventKey;
  }

  if (isPastSwitchTime(settings.currentEventDate, 11, 0)) {
    return settings.nextEventKey || settings.currentEventKey;
  }
  return settings.currentEventKey;
}

/**
 * ダッシュボード用開催日
 * 切り替え: 例会開催日 11:00
 */
function getDashboardEventDate() {
  const settings = getEventSettings();
  if (!settings.currentEventDate) {
    return settings.currentEventDate;
  }

  if (isPastSwitchTime(settings.currentEventDate, 11, 0)) {
    return settings.nextEventDate || settings.currentEventDate;
  }
  return settings.currentEventDate;
}

/**
 * 出欠アプリ用イベントキー
 * 切り替え: 例会開催日 12:00
 */
function getAttendanceEventKey() {
  const settings = getEventSettings();
  if (!settings.currentEventDate) {
    return settings.currentEventKey;
  }

  if (isPastSwitchTime(settings.currentEventDate, 12, 0)) {
    return settings.nextEventKey || settings.currentEventKey;
  }
  return settings.currentEventKey;
}

/**
 * 例会当日ページ用イベントキー
 * 設定シートのI2セル（例会当日イベントキー）を参照
 */
function getEventDayEventKey() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sh) {
    return '';
  }
  const eventDayKey = String(sh.getRange('I2').getValue() || '').trim();
  return eventDayKey;
}

/**
 * 例会当日ページ用開催日
 * I2のイベントキーに対応する開催日を返す
 */
function getEventDayEventDate() {
  const settings = getEventSettings();
  const eventDayKey = getEventDayEventKey();

  // I2のキーがcurrentEventKeyと同じなら currentEventDate
  if (eventDayKey === settings.currentEventKey) {
    return settings.currentEventDate;
  }
  // nextEventKeyと同じなら nextEventDate
  if (eventDayKey === settings.nextEventKey) {
    return settings.nextEventDate;
  }
  // どちらでもない場合はnull
  return null;
}

/**
 * 例会当日ページ用イベント情報を取得
 * 例会開催日が終了するまで現在月の情報を表示
 */
function getEventDayEventInfo_() {
  const settings = getEventSettings();
  const keyRaw = getEventDayEventKey();
  const eventDate = getEventDayEventDate();

  // デバッグログ
  Logger.log('=== getEventDayEventInfo_ DEBUG ===');
  Logger.log('settings.currentEventKey: ' + settings.currentEventKey);
  Logger.log('settings.currentEventDate: ' + settings.currentEventDate);
  Logger.log('settings.nextEventKey: ' + settings.nextEventKey);
  Logger.log('settings.nextEventDate: ' + settings.nextEventDate);
  Logger.log('getEventDayEventKey result: ' + keyRaw);
  Logger.log('now: ' + new Date());

  if (!keyRaw) {
    return { key: '', name: '', eventDate: '' };
  }

  const name = eventKeyToJapanese(keyRaw);
  return { key: keyRaw, name, eventDate };
}

/**
 * ゲスト申請用イベントキー
 * 切り替え: 例会開催日 12:00（出欠と同じ）
 */
function getGuestEventKey() {
  return getAttendanceEventKey();
}

/**
 * 売上アプリ用イベントキー
 * 切り替え: 例会開催前日 12:00（締め）
 */
function getSalesEventKey() {
  const settings = getEventSettings();
  if (!settings.currentEventDate) {
    return settings.currentEventKey;
  }

  // 前日12:00で切り替え
  if (isPastSwitchTime(settings.currentEventDate, 12, -1)) {
    return settings.nextEventKey || settings.currentEventKey;
  }
  return settings.currentEventKey;
}

/**
 * イベントキーを日本語形式に変換
 * 例: 202601_01 → 2026年1月例会
 */
function eventKeyToJapanese(eventKey) {
  if (!eventKey) return '';
  const match = String(eventKey).match(/^(\d{4})(\d{2})_\d+$/);
  if (!match) return eventKey;
  const year = match[1];
  const month = parseInt(match[2], 10);
  return `${year}年${month}月例会`;
}

/**
 * 日本語形式をイベントキーに変換
 * 例: 2026年1月例会 → 202601_01
 */
function japaneseToEventKey(japanese) {
  if (!japanese) return '';
  const match = String(japanese).match(/^(\d{4})年(\d{1,2})月例会$/);
  if (!match) return japanese;
  const year = match[1];
  const month = match[2].padStart(2, '0');
  return `${year}${month}_01`;
}


/** =========================================================
 *  共通ヘルパー
 * ======================================================= */

/**
 * Webアプリのエントリポイント
 */
function doGet(e) {
  // パラメータがない場合のガード
  if (!e || !e.parameter) {
    e = { parameter: {} };
  }

  const action = e.parameter.action || '';
  
  // ★LIFF用エンドポイント（イベント情報を返す）
  if (action === 'getCurrentEvent') {
    const eventInfo = getCurrentEventInfo_();
    return ContentService
      .createTextOutput(JSON.stringify({
        eventKey: eventInfo.key,
        eventName: eventInfo.name
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★デバッグ用：設定情報を返す
  if (action === 'debugEventSettings') {
    const settings = getEventSettings();
    const now = new Date();
    const dashboardKey = getDashboardEventKey();
    const eventDayKey = getEventDayEventKey();
    return ContentService
      .createTextOutput(JSON.stringify({
        now: now.toISOString(),
        settings: {
          currentEventKey: settings.currentEventKey,
          currentEventDate: settings.currentEventDate ? settings.currentEventDate.toISOString() : null,
          nextEventKey: settings.nextEventKey,
          nextEventDate: settings.nextEventDate ? settings.nextEventDate.toISOString() : null
        },
        dashboardKey,
        eventDayKey_I2: eventDayKey
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★配席アーカイブシート作成（1回だけ実行）
  if (action === 'createSeatingArchiveSheet') {
    const result = createSeatingArchiveSheet();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★配席アプリ用API: 参加者一覧取得
  if (action === 'getSeatingParticipants') {
    const eventKey = e.parameter.eventKey || '';
    const result = getSeatingParticipants(eventKey);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★配席アプリ用API: 過去配席取得
  if (action === 'getSeatingArchive') {
    const eventKey = e.parameter.eventKey || '';
    const result = getSeatingArchive(eventKey);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★配席アプリ用API: アーカイブ一覧
  if (action === 'listSeatingArchives') {
    const result = listSeatingArchives();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★イベントキー自動切替: 初期セットアップ（1回だけ実行）
  if (action === 'setupEventKeyColumns') {
    const result = setupEventKeyColumns();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★出欠アプリ用: 現在のイベントキー取得
  if (action === 'getAttendanceEventKey') {
    const eventKey = getAttendanceEventKey();
    const eventName = eventKeyToJapanese(eventKey);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        eventKey: eventKey,
        eventName: eventName
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★ゲスト申請用: 現在のイベントキー取得
  if (action === 'getGuestEventKey') {
    const eventKey = getGuestEventKey();
    const eventName = eventKeyToJapanese(eventKey);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        eventKey: eventKey,
        eventName: eventName
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★個別売上: パスワード認証
  if (action === 'verifySalesPassword') {
    const password = e.parameter.password || '';
    const result = verifySalesPassword(password);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★個別売上: 月一覧取得
  if (action === 'getSalesMonthList') {
    const result = getSalesMonthList();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★個別売上: 月別データ取得
  if (action === 'getIndividualSales') {
    const month = e.parameter.month || '';
    const password = e.parameter.password || '';
    const result = getIndividualSales(month, password);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★役割分担: 取得
  if (action === 'getRoleAssignments') {
    const eventKey = e.parameter.eventKey || '';
    const result = getRoleAssignments(eventKey);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★役割分担: デフォルト役割一覧取得
  if (action === 'getDefaultRoles') {
    const result = getDefaultRoles();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★役割分担: 役割候補者一覧取得（役割列が空でない会員）
  if (action === 'getRoleCandidates') {
    const result = getRoleCandidates();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★他会場: フォーム回答→他会場名簿マスター同期
  if (action === 'syncFormToOtherVenueMaster') {
    const result = syncFormResponsesToOtherVenueMaster();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★ダッシュボード用API: ダッシュボードデータ取得（高速版）
  if (action === 'getDashboardDataFast') {
    try {
      const result = getDashboardDataFast();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★ダッシュボード用API: イベント当日データ取得
  if (action === 'getEventDayData') {
    try {
      const result = getEventDayData();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★ダッシュボード用API: イベントキー一覧取得
  if (action === 'getEventKeyList') {
    try {
      const result = getEventKeyList();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★ダッシュボード用API: 参加者履歴取得
  if (action === 'getParticipantHistory') {
    try {
      const eventKey = e.parameter.eventKey || '';
      const result = getParticipantHistory(eventKey);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★ダッシュボード用API: 受付名簿データ取得
  if (action === 'getReceptionRosterData') {
    try {
      const result = getReceptionRosterData();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★チェックイン用API: チェックイン実行（LIFF用）
  if (action === 'checkin') {
    try {
      const lineUserId = e.parameter.userId || '';
      const result = doCheckin(lineUserId);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★チェックイン用API: チェックインキャンセル（LIFF用）
  if (action === 'cancelCheckin') {
    try {
      const lineUserId = e.parameter.userId || '';
      const result = cancelCheckin(lineUserId);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★チェックイン用API: 参加者情報取得（LIFF用）
  if (action === 'getParticipantInfo') {
    try {
      const lineUserId = e.parameter.userId || '';
      const result = getParticipantInfo(lineUserId);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★チェックイン用API: チェックイン状況一覧取得（ダッシュボード用）
  if (action === 'getCheckinStatus') {
    try {
      const result = getCheckinStatus();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★チェックインデータ初期化（例会当日の朝に実行）
  if (action === 'initCheckinData') {
    try {
      const result = initCheckinData();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★ダッシュボード用API: 式次第データ取得
  if (action === 'getShikidaiData') {
    try {
      const result = getShikidaiData();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★ダッシュボード用API: PDF名簿生成
  if (action === 'exportRosterPdf') {
    try {
      const result = exportRosterPdf();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★ダッシュボード用API: 他会場スケジュール更新
  if (action === 'updateOtherVenueSchedule') {
    try {
      const result = updateOtherVenueSchedule();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★会員編集機能: パスワード検証
  if (action === 'verifyMemberEditPassword') {
    const password = e.parameter.password || '';
    const result = verifyMemberEditPassword(password);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ★会員編集機能: 会員一覧取得
  if (action === 'getMembersForEdit') {
    try {
      const filter = {
        team: e.parameter.team || '',
        badge: e.parameter.badge || '',
        renewalMonth: e.parameter.renewalMonth || '',
        includeRetired: e.parameter.includeRetired === 'true',
        searchText: e.parameter.searchText || ''
      };
      const result = getMembersForEdit(filter);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★出欠リマインド機能: 未回答者リスト取得
  if (action === 'getUnrespondedMembersForReminder') {
    try {
      const result = getUnrespondedMembersForReminder();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★出欠リマインド機能: ドライラン（テスト実行）
  if (action === 'sendAttendanceReminderDryRun') {
    try {
      const result = sendAttendanceReminderDryRun();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★出欠リマインド機能: トリガー状態取得
  if (action === 'getReminderTriggerStatus') {
    try {
      const result = getReminderTriggerStatus();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★連絡網グループ取得
  if (action === 'getContactGroups') {
    try {
      const result = getContactGroups();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★連絡網専用ページ（全会員向け）
  const page = e.parameter.page || '';
  if (page === 'contact') {
    return HtmlService.createTemplateFromFile('contact')
      .evaluate()
      .setTitle('連絡網グループ - 守成クラブ福岡飯塚')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const userId = e.parameter.userId;
  
  if (!userId) {
    // 管理画面
    const template = HtmlService.createTemplateFromFile('index');
    // 式次第データを埋め込み（エラー時は空オブジェクト）
    let shikidaiData = {};
    try {
      shikidaiData = getShikidaiData() || {};
    } catch (e) {
      console.error('getShikidaiData failed in doGet:', e.message);
      shikidaiData = { error: e.message };
    }
    template.shikidaiData = JSON.stringify(shikidaiData);
    return template
      .evaluate()
      .setTitle('守成クラブ福岡飯塚 ダッシュボード')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    // 会員専用ページ
    const template = HtmlService.createTemplateFromFile('mypage');
    template.userId = userId;
    return template
      .evaluate()
      .setTitle('マイページ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// 会員名簿マスターから「氏名 → チーム名」のマップを作る
function getMemberTeamMap() {
  // 出欠＆ゲスト用スプシ（会員名簿マスターが入っている方）
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME); // '会員名簿マスター'

  const values = sh.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const name = values[i][0]; // A列：氏名
    const team = values[i][8]; // I列：チーム（0始まり index 8）

    if (name && team) {
      map[name] = team;
    }
  }
  return map; // 例: { "木下 智子": "Aチーム", ... }
}

/**
 * チーム別売上ランキング（例会月の「前月」の売上報告を集計）
 */
function getTeamSalesRanking(eventKey) {
  const ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  const sh = ss.getSheetByName(SALES_SHEET_NAME);

  const lastRow = sh.getLastRow();
  if (lastRow < 7) return [];

  // ★ eventKey("202501_01") → 前月の売上報告eventKey("2024年12月_売上報告")
  const ym = parseYearMonthFromEventKey_(eventKey);
  let salesEventKey = '';
  if (ym) {
    const prevDate = new Date(ym.year, ym.month - 2, 1);
    const prevYear = prevDate.getFullYear();
    const prevMonth = prevDate.getMonth() + 1;
    salesEventKey = `${prevYear}年${prevMonth}月_売上報告`;
  }

  // 7行目以降を取得
  const values = sh.getRange(7, 1, lastRow - 6, sh.getLastColumn()).getValues();

  const memberTeamMap = getMemberTeamMap();
  const teamSalesMap = {};

  values.forEach(row => {
    const rowEventKey = String(row[14] || '').trim(); // O列
    if (rowEventKey !== salesEventKey) return;

    const amount = Number(row[5]) || 0;  // F列：金額
    const member = row[13];              // N列：氏名

    if (!amount) return;

    const team = memberTeamMap[member] || '未所属';

    if (!teamSalesMap[team]) teamSalesMap[team] = 0;
    teamSalesMap[team] += amount;
  });

  const ranking = Object.keys(teamSalesMap).map(team => ({
    team,
    amount: teamSalesMap[team]
  })).sort((a, b) => b.amount - a.amount);

  return ranking.slice(0, 7);
}



/**
 * ダッシュボード用イベント情報を取得
 * ★ 例会開催日 11:00 で自動切り替え
 *  - 現在イベントキー（202512_01）
 *  - 現在イベント名（2025年12月例会）
 *  - 開催日
 */
function getCurrentEventInfo_() {
  // ★ 自動切り替え対応: getDashboardEventKey() を使用
  const keyRaw = getDashboardEventKey();
  const eventDate = getDashboardEventDate();

  if (!keyRaw) {
    return { key: '', name: '', eventDate: '' };
  }

  // 例：202512_01 → 2025年12月例会
  const name = eventKeyToJapanese(keyRaw);

  return { key: keyRaw, name, eventDate };
}

// eventKey(202512_01 など) から 年月を取り出す
function parseYearMonthFromEventKey_(key) {
  if (!key || key.length < 6) return null;
  const year  = Number(key.substring(0, 4));
  const month = Number(key.substring(4, 6));
  if (!year || !month) return null;
  return { year, month };
}

// 年月に月数を加算（翌月・翌々月計算用）
function addMonth_(year, month, add) {
  const d = new Date(year, month - 1 + add, 1);
  return {
    year: d.getFullYear(),
    month: d.getMonth() + 1,
  };
}

// 入会月セル(M列)から 年月を取り出す
function parseYearMonthCell_(value) {
  if (!value) return null;

  // 日付型で入っている場合
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return {
      year: value.getFullYear(),
      month: value.getMonth() + 1,
    };
  }

  const str = String(value).trim();
  if (!str) return null;

  // 例: 2025/03, 2025-3, 2025年3月 → 2025 と 3 を抜き出す
  const m = str.match(/(\d{4})\D+(\d{1,2})/);
  if (!m) return null;

  const year  = Number(m[1]);
  const month = Number(m[2]);
  if (!year || !month) return null;

  return { year, month };
}

/**
 * 会員数（会員名簿マスターの行数）を取得
 */
function getMemberCount_() {
  return getAllMembers_().length;
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sh) return 0;
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return 0; // ヘッダーのみ
  return lastRow - 1;
}

/**
 * ヘッダー行から任意の候補名の列インデックスを探す
 */
function findColumnIndex_(headerRow, candidates) {
  for (let i = 0; i < headerRow.length; i++) {
    const v = String(headerRow[i] || '').trim();
    if (!v) continue;
    if (candidates.includes(v)) return i;
  }
  return -1;
}

/**
 * 会員名簿マスターから全会員（氏名＋バッジ）の一覧を取得
 * 氏名は主に C列（ヘッダー: displayName / 氏名 / 名前）を優先
 * バッジは B列（ゴ/正/準）
 */
function getAllMembers_() {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (values.length < 3) return [];

  const header = values[1];
  const idxName = findColumnIndex_(header, ['displayName', '氏名', '名前']);
  const nameCol = idxName !== -1 ? idxName : 2; // デフォルト C列
  const badgeCol = 1; // B列

  const list = [];
  const seen = {};

  values.slice(2).forEach(row => {
    const name = String(row[nameCol] || '').trim();
    if (!name || seen[name]) return;
    const badge = String(row[badgeCol] || '').trim();
    list.push({ name, badge });
    seen[name] = true;
  });

  return list;
}

/**
 * 会員名簿マスターから「userId → {name,badge}」マップを作る
 */
function getMemberMapByUserId_() {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sh) return {};

  const values = sh.getDataRange().getValues();
  if (values.length < 3) return {};

  const header = values[1];
  const idxUserId = findColumnIndex_(header, ['LINE_userId', 'userId', 'ユーザーID']);
  const idxName   = findColumnIndex_(header, ['displayName', '氏名', '名前']);
  const nameCol   = idxName !== -1 ? idxName : 2; // C列
  const badgeCol  = 1; // B列

  const map = {};

  if (idxUserId === -1) {
    return map;
  }

  values.slice(2).forEach(row => {
    const uid = String(row[idxUserId] || '').trim();
    if (!uid) return;
    const name  = String(row[nameCol]  || '').trim();
    const badge = String(row[badgeCol] || '').trim();
    map[uid] = { name, badge };
  });

  return map;
}

/**
 * 他会場会員集計（詳細表示用）
 */
function aggregateOtherVenues_(eventName) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName('他会場名簿マスター');
  if (!sh) {
    return { detail: [] };
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 3) {
    return { detail: [] };
  }

  const HEADER_ROW_INDEX = 2; // 3行目（0-based）
  const header = values[HEADER_ROW_INDEX];
  const rows   = values.slice(HEADER_ROW_INDEX + 1);

  // 列インデックス
  const idxEvent    = findColumnIndex_(header, ['イベント', 'eventKey', '例会']);
  const idxName     = findColumnIndex_(header, ['氏名', '名前']);
  const idxFurigana = findColumnIndex_(header, ['フリガナ']);
  const idxVenue    = findColumnIndex_(header, ['所属', '会場']);
  const idxRole     = findColumnIndex_(header, ['役割']);
  const idxCompany  = findColumnIndex_(header, ['会社名']);
  const idxPosition = findColumnIndex_(header, ['役職']);
  const idxIntro    = findColumnIndex_(header, ['紹介者']);
  const idxBusiness = findColumnIndex_(header, ['営業内容']);

  // デフォルト値（列が見つからない場合）
  const cEvent    = idxEvent    !== -1 ? idxEvent    : 1;  // B列
  const cName     = idxName     !== -1 ? idxName     : 3;  // D列
  const cFurigana = idxFurigana !== -1 ? idxFurigana : 4;  // E列
  const cVenue    = idxVenue    !== -1 ? idxVenue    : 5;  // F列
  const cRole     = idxRole     !== -1 ? idxRole     : 6;  // G列
  const cCompany  = idxCompany  !== -1 ? idxCompany  : 7;  // H列
  const cPosition = idxPosition !== -1 ? idxPosition : 8;  // I列
  const cIntro    = idxIntro    !== -1 ? idxIntro    : 9;  // J列
  const cBusiness = idxBusiness !== -1 ? idxBusiness : 10; // K列

  const detail = [];

  rows.forEach(row => {
    const event = String(row[cEvent] || '').trim();
    
    // eventName（例：2025年12月例会）が含まれているかチェック
    if (!event.includes(eventName.replace('例会', ''))) return;

    const name     = String(row[cName]     || '').trim();
    const furigana = String(row[cFurigana] || '').trim();
    const venue    = String(row[cVenue]    || '').trim();
    const role     = String(row[cRole]     || '').trim();
    const company  = String(row[cCompany]  || '').trim();
    const position = String(row[cPosition] || '').trim();
    const introName= String(row[cIntro]    || '').trim();
    const business = String(row[cBusiness] || '').trim();

    if (!name) return;

    detail.push({
      name,
      furigana,
      venue,
      role,
      company,
      position,
      introName,
      business
    });
  });

  return { detail };
}

/** =========================================================
 *  バッジ集計（会員名簿マスターのB列）
 * ======================================================= */

/**
 * 会員名簿マスターのB列から「ゴ」「正」「準」の数を集計
 */
function aggregateBadges_() {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sh) {
    return { gold: 0, red: 0, green: 0 };
  }

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) {
    return { gold: 0, red: 0, green: 0 };
  }

  const values = sh.getRange(2, 2, lastRow - 1, 1).getValues(); // B2:B[最終行]

  let gold = 0;   // ゴ
  let red = 0;    // 正
  let green = 0;  // 準

  values.forEach(row => {
    const val = String(row[0] || '').trim();
    if (val === 'ゴ') gold++;
    else if (val === '正') red++;
    else if (val === '準') green++;
  });

  return { gold, red, green };
}

/** =========================================================
 *  チーム集計（会員名簿マスターの I 列）
 * ======================================================= */

/**
 * 会員名簿マスターの I 列（3行目以降）を見て、
 * チームごとにメンバーをグルーピングする。
 * I 列が空欄の人は「未配属」としてまとめる。
 */
function aggregateTeams_() {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sh) {
    return { teams: [], unassigned: [] };
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 3) {
    return { teams: [], unassigned: [] };
  }

  // 3行目以降がデータ
  const rows = values.slice(2);

  const idxBadge = 1; // B列（0-based）
  const idxName  = 2; // C列
  const idxTeam  = 8; // I列

  const teamMap = {};
  const unassigned = [];

  rows.forEach(row => {
    const name  = String(row[idxName]  || '').trim();
    const team  = String(row[idxTeam]  || '').trim();
    const badge = String(row[idxBadge] || '').trim(); // ゴ / 正 / 準

    if (!name) return;

    const member = { name, badge };

    if (team) {
      if (!teamMap[team]) {
        teamMap[team] = [];
      }
      teamMap[team].push(member);
    } else {
      unassigned.push(member);
    }
  });

  const teams = Object.keys(teamMap)
    .sort()
    .map(teamName => ({
      teamName,
      members: teamMap[teamName],
    }));

  return {
    teams,
    unassigned,
  };
}

// 入会月ではなく「更新月(T列: 1〜12)」から
// イベントキーの翌月・翌々月に更新となる会員を取得
function getRenewalMembers_() {
  const eventInfo = getCurrentEventInfo_();       // { key: '202512_01', name: '2025年12月例会' }
  const ym = parseYearMonthFromEventKey_(eventInfo.key);
  if (!ym) {
    return {
      next: [],
      next2: [],
      nextMonthLabel: '来月更新',
      next2MonthLabel: '再来月更新',
    };
  }

  // 例: 202512_01 → ym.month = 12
  const next  = addMonth_(ym.year, ym.month, 1);  // 翌月
  const next2 = addMonth_(ym.year, ym.month, 2);  // 再来月

  const nextMonth  = next.month;   // 1月
  const next2Month = next2.month;  // 2月

  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sh) {
    return {
      next: [],
      next2: [],
      nextMonthLabel: nextMonth  + '月更新',
      next2MonthLabel: next2Month + '月更新',
    };
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 3) {
    return {
      next: [],
      next2: [],
      nextMonthLabel: nextMonth  + '月更新',
      next2MonthLabel: next2Month + '月更新',
    };
  }

  // 3行目以降がデータ
  const rows = values.slice(2);

  // 列インデックス（0始まり）
  const idxBadge          = 1;   // B列：バッジ
  const idxName           = 2;   // C列：氏名
  const idxReferralCount  = 18;  // Q列：紹介在籍人数
  const idxUpdateMonth    = 21;  // V列：更新月(1〜12)

  const nextArr  = [];  // 翌月更新
  const next2Arr = [];  // 再来月更新

  rows.forEach(row => {
    const name = String(row[idxName] || '').trim();
    if (!name) return;

    const badge = String(row[idxBadge] || '').trim();

    const mRaw = row[idxUpdateMonth];
    if (mRaw === '' || mRaw == null) return;

    // 更新月（1〜12）を数値に
    const n = Number(String(mRaw).trim());
    if (!Number.isFinite(n)) return;
    const updateMonth = n;

    // 紹介在籍人数（Q列）
    const referralCountRaw = row[idxReferralCount];
    const referralCount = Number(referralCountRaw) || 0;
    const hasZeroReferral = (referralCount === 0);

    const member = {
      name,
      badge,
      hasZeroReferral,    // ← フロントで ⚠️ に使うフラグ
      referralCount,      // ついでに数も渡しておく
    };

    if (updateMonth === nextMonth) {
      nextArr.push(member);
    } else if (updateMonth === next2Month) {
      next2Arr.push(member);
    }
  });

  return {
    next: nextArr,
    next2: next2Arr,
    nextMonthLabel: nextMonth  + '月更新',
    next2MonthLabel: next2Month + '月更新',
  };
}



/** =========================================================
 *  ダッシュボードメイン
 * ======================================================= */

/**
 * ダッシュボード用データをまとめて返す
 */
function getDashboardData() {
  const eventInfo   = getCurrentEventInfo_();
  const eventName   = eventInfo.name;  // 2025年12月例会
  const eventKey    = eventInfo.key;   // ★追加
  const eventDate   = eventInfo.eventDate;  // ★追加
  const memberCount = getMemberCount_();
  

  const attendance       = aggregateAttendance_(eventName, memberCount);
  const attendanceDetail = aggregateAttendanceDetail_(eventName);
  const sales            = aggregateSales_(eventName);
  const guests           = aggregateGuests_(eventName);
  const badges           = aggregateBadges_();
  const teams            = aggregateTeams_();
  const renewals         = getRenewalMembers_();
  const otherVenues      = aggregateOtherVenues_(eventName);
  const prep             = aggregatePrep_(eventName);
  const monthlySummary   = getMonthlySummary_();

  return {
    eventName,
    eventKey,      // ★追加
    eventDate,     // ★追加
    attendance,
    attendanceDetail,
    sales,
    guests,
    otherVenues,
    badges,
    teams,
    renewals,
    prep,
    monthlySummary,
  };
}

/**
 * 例会当日用データ取得（役割分担）
 * ★ 例会開催日が終了するまで現在月の情報を表示
 */
function getEventDayData() {
  const eventInfo = getEventDayEventInfo_();  // ★ 例会当日ページ用イベント情報
  const eventName = eventInfo.name;  // 2025年12月例会
  const eventKey  = eventInfo.key;   // 202512_01

  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sh) {
    return { eventName, roles: {}, tableMasters: [] };
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 3) {
    return { eventName, roles: {}, tableMasters: [] };
  }

  const header = values[1]; // 2行目
  const rows   = values.slice(2); // 3行目以降

  const idxName  = findColumnIndex_(header, ['displayName', '氏名', '名前']);
  const nameCol  = idxName !== -1 ? idxName : 2; // C列
  const badgeCol = 1; // B列

  // 役割は X/Y/Z列（23, 24, 25）
  const roleCol1 = 23; // X列（0始まり）
  const roleCol2 = 24; // Y列
  const roleCol3 = 25; // Z列

  // 役割ごとに分類
  const rolesMap = {};

  rows.forEach(row => {
    const name  = String(row[nameCol]  || '').trim();
    const badge = String(row[badgeCol] || '').trim();

    if (!name) return;

    const role1 = String(row[roleCol1] || '').trim();
    const role2 = String(row[roleCol2] || '').trim();
    const role3 = String(row[roleCol3] || '').trim();

    [role1, role2, role3].forEach(r => {
      if (!r) return;
      if (!rolesMap[r]) rolesMap[r] = [];
      rolesMap[r].push({ name, badge });
    });
  });

  // 配席アーカイブからテーブルマスターを取得
  const seatingArchive = getSeatingArchive(eventKey);
  const seatingTables = (seatingArchive.success && seatingArchive.tables) ? seatingArchive.tables : {};

  // テーブルマスター一覧（配席アーカイブに実際にあるテーブルを全て取得）
  // ★テーブル記号が空欄のものは除外
  const tableLabels = Object.keys(seatingTables)
    .filter(t => t && t.trim() !== '')  // 空欄を除外
    .sort((a, b) => {
      // PA, MC を最後にソート、それ以外はアルファベット順
      const specialTables = ['PA', 'MC'];
      const aSpecial = specialTables.includes(a);
      const bSpecial = specialTables.includes(b);
      if (aSpecial && !bSpecial) return 1;
      if (!aSpecial && bSpecial) return -1;
      return a.localeCompare(b);
    });
  const tableMasters = tableLabels.map(t => ({
    table: t,
    tm: seatingTables[t]?.master || ''
  }));

  return {
    eventName,
    eventKey,
    roles: rolesMap,
    tableMasters,
  };
}

/**
 * 例会事前準備用：参加者名簿データ
 *  - local : 福岡飯塚会場_参加者名簿 用
 *  - others: 他会場・ゲスト_参加者名簿 用（今回はまだ空）
 *
 * 出欠欄のルール：
 *   出欠状況（自動）で
 *     ○ / 出席 → そのまま出力
 *     × / 欠席 → 空欄にする
 *     それ以外・未回答 → 空欄
 */
/**
 * 例会事前準備用：参加者名簿データ
 *  - local : 福岡飯塚会場_参加者名簿 用
 *  - others: 他会場・ゲスト_参加者名簿 用（今回はまだ空）
 *
 * データソース：
 *   会員名簿マスター: L列=受賞, N列=紹介数, D列=フリガナ, K列=営業内容
 *   出欠状況（自動）: E列=出欠, F列=ブース
 */
function aggregatePrep_(eventName) {
  // 1) 出欠状況（自動）から「userId → info」と「displayName → info」マップを作る
  const ssAttend = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const shAttend = ssAttend.getSheetByName(ATTEND_SHEET_NAME);

  const attendInfoByUserId = {};
  const attendInfoByDisplayName = {};  // ★displayNameでマッチング

  if (shAttend) {
    const valuesA = shAttend.getDataRange().getValues();
    if (valuesA.length >= 4) {
      const HEADER_ROW_INDEX = 2;
      const headerA = valuesA[HEADER_ROW_INDEX];
      const rowsA   = valuesA.slice(HEADER_ROW_INDEX + 1);

      const idxEventKey    = findColumnIndex_(headerA, ['eventKey', 'イベントキー']);
      const idxUserId      = findColumnIndex_(headerA, ['userId', 'LINE_userId', 'ユーザーID']);
      const idxDisplayName = findColumnIndex_(headerA, ['displayName', '氏名', '名前']);
      const idxStatus      = 4;
      const idxBooth       = 5;

      rowsA.forEach(row => {
        if (idxEventKey === -1) return;

        const key         = String(row[idxEventKey] || '').trim();
        if (key !== eventName) return;

        const userId      = idxUserId !== -1 ? String(row[idxUserId] || '').trim() : '';
        const displayName = idxDisplayName !== -1 ? String(row[idxDisplayName] || '').trim() : '';
        const st          = String(row[idxStatus] || '').trim();
        const booth       = String(row[idxBooth]  || '').trim();

        const info = { status: st, booth: booth };

        if (userId) {
          attendInfoByUserId[userId] = info;
        }
        if (displayName) {
          attendInfoByDisplayName[displayName] = info;
        }
      });
    }
  }

  // 2) 会員名簿マスターから名簿用の行を作る
  const ssMember = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const shMember = ssMember.getSheetByName(MEMBER_SHEET_NAME);
  if (!shMember) {
    return { local: [], others: [] };
  }

  const values = shMember.getDataRange().getValues();
  if (values.length < 3) {
    return { local: [], others: [] };
  }

  const rows = values.slice(2);

  // 固定列位置（0始まり）
  const idxBadge       = 1;   // B列
  const idxName        = 2;   // C列：氏名
  const idxFurigana    = 3;   // D列
  const idxBusiness    = 10;  // K列
  const idxAward       = 11;  // L列
  const idxReferral    = 13;  // N列
  const idxDisplayName = 14;  // O列：displayName ★追加
  const idxUserId      = 15;  // P列：LINE_userId ★修正

  const header = values[1];
  const idxRole       = findColumnIndex_(header, ['役割']);
  const idxCompany    = findColumnIndex_(header, ['会社名', '屋号', '店舗名']);
  const idxPosition   = findColumnIndex_(header, ['役職', '肩書']);
  const idxIntroducer = findColumnIndex_(header, ['紹介者', '紹介者名']);

  const localList = [];

  rows.forEach(row => {
    const name = String(row[idxName] || '').trim();
    if (!name) return;

    // ★修正：userId → displayName の順でフォールバック
    const userId      = String(row[idxUserId]      || '').trim();
    const displayName = String(row[idxDisplayName] || '').trim();
    
    let attendInfo = null;
    
    if (userId && attendInfoByUserId[userId]) {
      // LINE_userIdでマッチ
      attendInfo = attendInfoByUserId[userId];
    } else if (displayName && attendInfoByDisplayName[displayName]) {
      // displayNameでマッチ
      attendInfo = attendInfoByDisplayName[displayName];
    } else {
      attendInfo = {};
    }
    
    const rawStatus = attendInfo.status || '';
    const booth     = attendInfo.booth  || '';

    // 出欠欄のルール
    let displayStatus = '';
    const statusTrimmed = rawStatus.trim();

    if (statusTrimmed.includes('○') || statusTrimmed.includes('o') || 
        statusTrimmed.includes('〇') || statusTrimmed === '出席') {
      displayStatus = '○';
    } else if (statusTrimmed.includes('×') || statusTrimmed === '欠席') {
      displayStatus = '';
    } else {
      displayStatus = '';
    }

    localList.push({
      award:        String(row[idxAward]     || '').trim(),
      referralCnt:  row[idxReferral],
      badge:        String(row[idxBadge]     || '').trim(),
      attendStatus: displayStatus,
      booth:        booth,
      name:         name,
      furigana:     String(row[idxFurigana]  || '').trim(),
      role:         idxRole      !== -1 ? String(row[idxRole]      || '').trim() : '',
      company:      idxCompany   !== -1 ? String(row[idxCompany]   || '').trim() : '',
      position:     idxPosition  !== -1 ? String(row[idxPosition]  || '').trim() : '',
      introducer:   idxIntroducer!== -1 ? String(row[idxIntroducer]|| '').trim() : '',
      business:     String(row[idxBusiness]  || '').trim(),
    });
  });

  // ========================================
  // ★ others（他会場 + ゲスト）の取得
  // ========================================
  
  // 1) 他会場名簿マスターから取得
  const otherVenueList = [];
  const shOther = ssAttend.getSheetByName('他会場名簿マスター');
  if (shOther) {
    const valuesO = shOther.getDataRange().getValues();
    if (valuesO.length >= 3) {
      const HEADER_ROW_O = 1; // 2行目がヘッダー（0-based index 1）
      const rowsO = valuesO.slice(HEADER_ROW_O + 1); // 3行目以降がデータ

      // 列インデックス（0始まり）
      const cEvent    = 1;  // B列: EVENT_KEY
      const cBadge    = 2;  // C列: バッジ
      const cName     = 3;  // D列: 氏名
      const cFurigana = 4;  // E列: フリガナ
      const cVenue    = 5;  // F列: 所属
      const cRole     = 6;  // G列: 役割
      const cCompany  = 7;  // H列: 会社名
      const cPosition = 8;  // I列: 役職
      const cIntro    = 9;  // J列: 紹介者
      const cBusiness = 10; // K列: 営業内容
      const cBooth    = 12; // M列: ブース

      rowsO.forEach(row => {
        const event = String(row[cEvent] || '').trim();
        if (event !== eventName) return;

        const name = String(row[cName] || '').trim();
        if (!name) return;

        otherVenueList.push({
          badge:      String(row[cBadge] || '').trim(),
          booth:      String(row[cBooth] || '').trim(),
          name:       name,
          furigana:   String(row[cFurigana] || '').trim(),
          venue:      String(row[cVenue] || '').trim(),
          role:       String(row[cRole] || '').trim(),
          company:    String(row[cCompany] || '').trim(),
          position:   String(row[cPosition] || '').trim(),
          introducer: String(row[cIntro] || '').trim(),
          business:   String(row[cBusiness] || '').trim(),
        });
      });
    }
  }

  // 2) ゲスト出欠状況から取得
  const guestList = [];
  const shGuest = ssAttend.getSheetByName(GUEST_SHEET_NAME);
  if (shGuest) {
    const valuesG = shGuest.getDataRange().getValues();
    if (valuesG.length >= 3) {
      const HEADER_ROW_G = 1; // 2行目がヘッダー
      const rowsG = valuesG.slice(HEADER_ROW_G + 1);

      const gcEvent    = 2;  // C列: eventKey
      const gcName     = 3;  // D列: 氏名
      const gcFurigana = 4;  // E列: フリガナ
      const gcCompany  = 5;  // F列: 会社名
      const gcPosition = 6;  // G列: 役職
      const gcIntro    = 7;  // H列: 紹介者
      const gcBusiness = 9;  // J列: 営業内容

      rowsG.forEach(row => {
        const event = String(row[gcEvent] || '').trim();
        if (event !== eventName) return;

        const name = String(row[gcName] || '').trim();
        if (!name) return;

        guestList.push({
          introducer: String(row[gcIntro] || '').trim(),
          name:       name,
          furigana:   String(row[gcFurigana] || '').trim(),
          company:    String(row[gcCompany] || '').trim(),
          position:   String(row[gcPosition] || '').trim(),
          business:   String(row[gcBusiness] || '').trim(),
        });
      });
    }
  }

  return {
    local: localList,
    others: {
      otherVenue: otherVenueList,
      guest: guestList,
    },
  };
}





/** =========================================================
 *  出欠集計
 * ======================================================= */

/**
 * 出欠状況の集計
 * 想定：
 *   出欠状況（自動）
 *   3行目：timestamp / eventKey / userId / displayName / status ...
 *   4行目以降：データ
 */
function aggregateAttendance_(eventKey, memberCount) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(ATTEND_SHEET_NAME);
  if (!sh) {
    return { memberCount, attend: 0, absent: 0, noAnswer: memberCount };
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 4) {
    return { memberCount, attend: 0, absent: 0, noAnswer: memberCount };
  }

  const HEADER_ROW_INDEX = 2; // 3行目（0-based）
  const header = values[HEADER_ROW_INDEX];

  const idxEventKey = findColumnIndex_(header, ['eventKey', 'イベントキー']);
  const idxStatus   = findColumnIndex_(header, ['status', '出欠', '出欠状況', 'ステータス']);

  if (idxEventKey === -1 || idxStatus === -1) {
    return { memberCount, attend: 0, absent: 0, noAnswer: memberCount };
  }

  const rows = values.slice(HEADER_ROW_INDEX + 1);

  let attend = 0;
  let absent = 0;

  rows.forEach(row => {
    const key = String(row[idxEventKey] || '').trim();
    const st  = String(row[idxStatus]   || '').trim();

    if (key !== eventKey) return;

    if (st === '○' || st === '出席') {
      attend++;
    } else if (st === '×' || st === '欠席') {
      absent++;
    }
  });

  const noAnswer = Math.max(memberCount - attend - absent, 0);

  return { memberCount, attend, absent, noAnswer };
}

/**
 * 出欠詳細（出席 / 欠席 / 未回答）の一覧
 * 氏名・バッジは会員名簿マスター基準（C列氏名 / B列バッジ）
 */
function aggregateAttendanceDetail_(eventKey) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(ATTEND_SHEET_NAME);
  if (!sh) {
    const allMembers = getAllMembers_();
    return {
      attendList: [],
      absentList: [],
      noAnswerList: allMembers,
    };
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 4) {
    const allMembers = getAllMembers_();
    return {
      attendList: [],
      absentList: [],
      noAnswerList: allMembers,
    };
  }

  const HEADER_ROW_INDEX = 2; // 3行目
  const header = values[HEADER_ROW_INDEX];
  const rows   = values.slice(HEADER_ROW_INDEX + 1);

  const idxEventKey = findColumnIndex_(header, ['eventKey', 'イベントキー']);
  const idxStatus   = findColumnIndex_(header, ['status', '出欠', '出欠状況', 'ステータス']);
  const idxUserId   = findColumnIndex_(header, ['userId', 'LINE_userId', 'ユーザーID']);
  const idxNameCol  = findColumnIndex_(header, ['displayName', '氏名', '名前']); // fallback用
  const idxTimestamp = 0; // A列：timestamp

  const memberMapByUserId = getMemberMapByUserId_();
  const allMembers        = getAllMembers_();

  const attendList = [];
  const absentList = [];
  const repliedNameSet = {};

  rows.forEach(row => {
    if (idxEventKey === -1 || idxStatus === -1) return;

    const key = String(row[idxEventKey] || '').trim();
    if (key !== eventKey) return;

    const status = String(row[idxStatus] || '').trim();

    const userId = idxUserId !== -1 ? String(row[idxUserId] || '').trim() : '';
    const memberInfo = userId ? memberMapByUserId[userId] : null;

    let name = memberInfo ? memberInfo.name : '';
    let badge = memberInfo ? memberInfo.badge : '';

    // userId から引けなかった場合は、出欠シート側の氏名をfallbackに使う
    if (!name && idxNameCol !== -1) {
      name = String(row[idxNameCol] || '').trim();
    }

    if (!name) return;

    // timestamp
    const tsRaw = row[idxTimestamp];
    let timestamp = 0;
    let timeStr = '';

    if (tsRaw) {
      const d = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
      if (!isNaN(d)) {
        timestamp = d.getTime();
        timeStr = Utilities.formatDate(
          d,
          Session.getScriptTimeZone() || 'Asia/Tokyo',
          'MM/dd HH:mm'
        );
      }
    }

    const obj = { name, badge, time: timeStr, timestamp, status };

    if (status === '○' || status === '出席') {
      attendList.push(obj);
      repliedNameSet[name] = true;
    } else if (status === '×' || status === '欠席') {
      absentList.push(obj);
      repliedNameSet[name] = true;
    } else {
      // その他ステータスは今回は対象外
      return;
    }
  });

  // 回答時間の早い順
  attendList.sort((a, b) => (a.timestamp || 0) - (b.timestamp || 0));
  absentList.sort((a, b) => (a.timestamp || 0) - (b.timestamp || 0));

  // 未回答 = 会員名簿 - (出席 + 欠席)
  const noAnswerList = allMembers
    .filter(m => !repliedNameSet[m.name])
    .map(m => ({ name: m.name, badge: m.badge }));

  return {
    attendList,
    absentList,
    noAnswerList,
  };
}


/** =========================================================
 *  売上集計（sales!E3 = 商談件数 / F3 = 売上金額）
 *  + 会員別ランキングTOP3
 *  + チーム別売上ランキング（1〜7位）
 * ======================================================= */
/**
 * 売上集計（例会月の「前月」の売上報告を集計）
 */
function aggregateSales_(eventName) {
  const ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  const sh = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sh) {
    return { total: 0, dealCount: 0, meetingsCount: 0, reporterCount: 0, salesMonth: null, zeroList: [], ranking: [], teamRanking: [] };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 7) {
    return { total: 0, dealCount: 0, meetingsCount: 0, reporterCount: 0, salesMonth: null, zeroList: [], ranking: [], teamRanking: [] };
  }

  // eventName("2025年1月例会") → 前月の売上報告eventKey("2024年12月_売上報告")
  const monthMatch = eventName.match(/(\d{4})年(\d{1,2})月/);
  let salesEventKey = '';
  let salesMonth = null;
  let salesYear = null;
  if (monthMatch) {
    const year = Number(monthMatch[1]);
    const month = Number(monthMatch[2]);
    const prevDate = new Date(year, month - 2, 1);
    const prevYear = prevDate.getFullYear();
    const prevMonth = prevDate.getMonth() + 1;
    salesEventKey = `${prevYear}年${prevMonth}月_売上報告`;
    salesMonth = prevMonth;
    salesYear = prevYear;
  }

  const values = sh.getRange(7, 1, lastRow - 6, sh.getLastColumn()).getValues();

  let total = 0;
  let dealCount = 0;
  let meetingsCount = 0;
  let reporterCount = 0;
  const memberSales = {};
  const zeroList = [];  // ★追加

  values.forEach(row => {
    const rowEventKey = String(row[14] || '').trim(); // O列
    if (rowEventKey !== salesEventKey) return;

    reporterCount++;

    const amount   = Number(row[5]) || 0;   // F列：売上金額
    const deals    = Number(row[4]) || 0;   // E列：成約件数
    const meetings = Number(row[15]) || 0;  // P列：商談件数
    const member   = String(row[13] || '').trim(); // N列：氏名

    total += amount;
    dealCount += deals;
    meetingsCount += meetings;

    // ★商談件数と売上金額が両方0の人をリストに追加
if (meetings === 0 && amount === 0) {
  zeroList.push({
    name: member,
    meetings: meetings,
    deals: deals,
    amount: amount
  });
}

    if (member && amount > 0) {
      if (!memberSales[member]) memberSales[member] = 0;
      memberSales[member] += amount;
    }
  });

  // 会員別ランキング（TOP3）
  const ranking = Object.entries(memberSales)
    .map(([name, amount]) => ({ name, amount }))
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 3)
    .map(item => ({ name: item.name }));

  // チーム別売上ランキング
  const eventInfo = getCurrentEventInfo_();
  const eventKey = eventInfo.key || '';
  const teamRanking = eventKey ? getTeamSalesRanking(eventKey) : [];

  return {
    total,
    dealCount,
    meetingsCount,
    reporterCount,
    salesMonth,
    zeroList,  // ★追加
    ranking,
    teamRanking
  };
}


/** =========================================================
 *  ゲスト集計
 * ======================================================= */

/**
 * ゲスト申請状況の集計
 * 想定：
 *   ゲスト出欠状況（自動）
 *   2行目：Guest_ID / timestamp / eventKey / 氏名 / フリガナ / 会社名 / 役職 / 紹介者 / displayName / 営業内容 / 承認
 *   3行目以降：データ
 *   K列「承認」はチェックボックス（TRUE / FALSE）
 */
function aggregateGuests_(eventName) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(GUEST_SHEET_NAME);
  if (!sh) {
    return { total: 0, approved: 0, pending: 0, list: [], detail: [] };
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 3) {
    return { total: 0, approved: 0, pending: 0, list: [], detail: [] };
  }

  const HEADER_ROW_INDEX = 1; // 2行目（0-based）
  const header = values[HEADER_ROW_INDEX];
  const rows   = values.slice(HEADER_ROW_INDEX + 1);

  // 各列の位置をヘッダー名から取得（なければ既定位置）
  const idxEvent    = findColumnIndex_(header, ['eventKey', 'イベントキー']);
  const idxGuest    = findColumnIndex_(header, ['氏名', 'ゲスト名']);
  const idxFurigana = findColumnIndex_(header, ['フリガナ']);
  const idxCompany  = findColumnIndex_(header, ['会社名']);
  const idxPosition = findColumnIndex_(header, ['役職']);
  const idxIntro    = findColumnIndex_(header, ['紹介者', '紹介者名']);
  const idxBusiness = findColumnIndex_(header, ['営業内容']);
  const idxApprove  = findColumnIndex_(header, ['承認']);

  const cEvent    = idxEvent    !== -1 ? idxEvent    : 2;  // C列
  const cGuest    = idxGuest    !== -1 ? idxGuest    : 3;  // D列
  const cFurigana = idxFurigana !== -1 ? idxFurigana : 4;  // E列
  const cCompany  = idxCompany  !== -1 ? idxCompany  : 5;  // F列
  const cPosition = idxPosition !== -1 ? idxPosition : 6;  // G列
  const cIntro    = idxIntro    !== -1 ? idxIntro    : 7;  // H列
  const cBusiness = idxBusiness !== -1 ? idxBusiness : 9;  // J列
  const cApprove  = idxApprove  !== -1 ? idxApprove  : 10; // K列

  let total    = 0;
  let approved = 0;

  const list   = []; // 概要タブ用（今まで通り）
  const detail = []; // 詳細タブ用

  rows.forEach(row => {
    const key = String(row[cEvent] || '').trim();
    if (key !== eventName) return;  // eventKey が現在の例会名と一致するものだけ

    const guestName = String(row[cGuest]    || '').trim();
    const furigana  = String(row[cFurigana] || '').trim();
    const company   = String(row[cCompany]  || '').trim();
    const position  = String(row[cPosition] || '').trim();
    const introName = String(row[cIntro]    || '').trim();
    const business  = String(row[cBusiness] || '').trim();
    const isApproved = row[cApprove] === true; // チェックボックス

    total++;
    if (isApproved) approved++;

    // 概要タブ右側のリスト
    list.push({
      guestName,
      introName,
      approved: isApproved,
    });

    // ゲスト詳細タブ用
    detail.push({
      name:      guestName,
      furigana:  furigana,
      company:   company,
      position:  position,
      introName: introName,
      business:  business,
    });
  });

  const pending = total - approved;

  return { total, approved, pending, list, detail };
}

/** =========================================================
 *  月次サマリー取得（グラフ用）
 * ======================================================= */
function getMonthlySummary_() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName('月次サマリー');
  if (!sh) return { monthly: [], totals: null };

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { monthly: [], totals: null };

  // 1行目はヘッダー、2行目〜13行目がデータ（12ヶ月分）
  const monthly = [];
  for (let i = 1; i <= 12 && i < values.length; i++) {
    const row = values[i];
    const eventKey = String(row[0] || '').trim();
    if (!eventKey || eventKey === '年計') break;

    // 月ラベル抽出（例：2025年4月例会 → 4月）
    const monthMatch = eventKey.match(/(\d{1,2})月/);
    const monthLabel = monthMatch ? monthMatch[1] + '月' : eventKey;

    monthly.push({
      eventKey: eventKey,
      label: monthLabel,
      date: row[1],
      memberCount: Number(row[2]) || 0,
      attendCount: Number(row[3]) || 0,
      attendRate: parsePercent_(row[4]),
      reporterCount: Number(row[5]) || 0,
      submitRate: parsePercent_(row[6]),
      dealCount: Number(row[7]) || 0,
      salesAmount: Number(row[8]) || 0,
    });
  }

  // 総累計（17行目想定）
  let totals = null;
  if (values.length >= 17) {
    const row17 = values[16]; // 0-indexed
    totals = {
      dealCount: Number(row17[7]) || 0,
      salesAmount: Number(row17[8]) || 0,
    };
  }

  return { monthly, totals };
}

/**
 * 月次サマリーの売上データを更新
 * @param {number} salesMonth - 対象月（1-12）。省略時は全月更新
 */
function updateMonthlySalesSummary(salesMonth) {
  const configSs = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const summarySheet = configSs.getSheetByName('月次サマリー');
  if (!summarySheet) {
    return { success: false, message: '月次サマリーシートが見つかりません' };
  }

  // 売上報告シートからデータ取得
  const salesSs = SpreadsheetApp.openById(SALES_SHEET_ID);
  const salesSheet = salesSs.getSheetByName(SALES_SHEET_NAME);
  if (!salesSheet) {
    return { success: false, message: '売上報告シートが見つかりません' };
  }

  const lastRow = salesSheet.getLastRow();
  if (lastRow < 7) {
    return { success: false, message: '売上データがありません' };
  }

  // 売上データを月別に集計
  const salesData = salesSheet.getRange(7, 1, lastRow - 6, salesSheet.getLastColumn()).getValues();
  const monthlySales = {}; // { 月: { total, dealCount, reporterCount } }

  salesData.forEach(row => {
    const eventKey = String(row[14] || '').trim(); // O列
    const monthMatch = eventKey.match(/^(\d{4})年(\d{1,2})月_売上報告$/);
    if (!monthMatch) return;

    const year = Number(monthMatch[1]);
    const month = Number(monthMatch[2]);
    const key = `${year}-${month}`; // 年月をキーにする
    if (salesMonth && month !== salesMonth) return; // 特定月のみ処理

    if (!monthlySales[key]) {
      monthlySales[key] = { year, month, total: 0, dealCount: 0, reporterCount: 0 };
    }

    monthlySales[key].total += Number(row[5]) || 0;      // F列：売上金額
    monthlySales[key].dealCount += Number(row[4]) || 0;  // E列：成約件数
    monthlySales[key].reporterCount++;
  });

  // 月次サマリーシートを更新
  const summaryValues = summarySheet.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < summaryValues.length; i++) {
    const eventKey = String(summaryValues[i][0] || '').trim();
    if (!eventKey || eventKey === '年計') continue;

    // eventKey（例：2025年12月例会）から年月を抽出
    const monthMatch = eventKey.match(/(\d{4})年(\d{1,2})月例会$/);
    if (!monthMatch) continue;

    const year = Number(monthMatch[1]);
    const month = Number(monthMatch[2]);
    if (salesMonth && month !== salesMonth) continue; // 特定月のみ処理

    const key = `${year}-${month}`;
    const sales = monthlySales[key];
    if (sales) {
      // F列(6): reporterCount, H列(8): dealCount, I列(9): salesAmount
      summarySheet.getRange(i + 1, 6).setValue(sales.reporterCount); // F列
      summarySheet.getRange(i + 1, 8).setValue(sales.dealCount);     // H列
      summarySheet.getRange(i + 1, 9).setValue(sales.total);         // I列
      updatedCount++;
      Logger.log(`${year}年${month}月: 売上 ${sales.total.toLocaleString()}円 / 成約 ${sales.dealCount}件 / 報告者 ${sales.reporterCount}人`);
    }
  }

  // 年計行も更新（17行目想定）
  if (!salesMonth && summaryValues.length >= 17) {
    let yearTotal = 0, yearDealCount = 0;
    Object.values(monthlySales).forEach(s => {
      yearTotal += s.total;
      yearDealCount += s.dealCount;
    });
    summarySheet.getRange(17, 8).setValue(yearDealCount); // H列
    summarySheet.getRange(17, 9).setValue(yearTotal);     // I列
    Logger.log(`年計: 売上 ${yearTotal.toLocaleString()}円 / 成約 ${yearDealCount}件`);
  }

  return {
    success: true,
    message: `${updatedCount}ヶ月分の売上データを更新しました`,
    monthlySales
  };
}

/**
 * 12月売上データを更新（テスト用）
 */
function updateDecemberSales() {
  return updateMonthlySalesSummary(12);
}

/**
 * 全月の売上データを更新
 */
function updateAllMonthlySales() {
  return updateMonthlySalesSummary();
}

// パーセント文字列を数値に変換（40.2% → 0.402）
function parsePercent_(val) {
  if (typeof val === 'number') return val;
  const str = String(val || '').replace('%', '').trim();
  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  return num / 100;
}


/** =========================================================
 *  会員マイページ用（新規追加）
 * ======================================================= */

/**
 * 会員個人のデータ取得
 */
function getMyData(userId) {
  const eventInfo = getCurrentEventInfo_();
  const eventName = eventInfo.name;
  
  // 会員名取得
  const userName = getUserName_(userId);
  
  // その人の出欠状況
  const myAttendance = getMyAttendance_(userId, eventName);
  
  // その人の紹介実績
  const myReferrals = getMyReferrals_(userId);
  
  // その人のランキング順位
  const myRanking = getMyRanking_(userId);
  
  return {
    userName,
    eventName,
    myAttendance,
    myReferrals,
    myRanking,
  };
}

/**
 * userIdから会員名を取得
 */
function getUserName_(userId) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sh) return '名前不明';
  
  const values = sh.getDataRange().getValues();
  if (values.length < 3) return '名前不明'; // 2行目ヘッダー、3行目以降データ
  
  // ヘッダー行は2行目（index=1）
  const header = values[1];
  const idxUserId = header.indexOf('LINE_userId') !== -1 ? header.indexOf('LINE_userId') : -1;
  const idxName = header.indexOf('displayName') !== -1 ? header.indexOf('displayName') : 
                  header.indexOf('氏名') !== -1 ? header.indexOf('氏名') : 2; // C列
  
  if (idxUserId === -1) return '名前不明';
  
  const rows = values.slice(2); // 3行目以降がデータ
  for (let row of rows) {
    if (String(row[idxUserId]).trim() === userId) {
      return String(row[idxName]).trim() || '名前不明';
    }
  }
  
  return '名前不明';
}

/**
 * その人の出欠状況
 */
function getMyAttendance_(userId, eventName) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sh = ss.getSheetByName(ATTEND_SHEET_NAME);
  if (!sh) {
    return { attended: false, status: '未回答' };
  }
  
  const values = sh.getDataRange().getValues();
  if (values.length < 4) {
    return { attended: false, status: '未回答' };
  }
  
  const HEADER_ROW_INDEX = 2; // 3行目
  const header = values[HEADER_ROW_INDEX];
  const rows = values.slice(HEADER_ROW_INDEX + 1);
  
  const idxEventKey = findColumnIndex_(header, ['eventKey', 'イベントキー']);
  const idxUserId = findColumnIndex_(header, ['userId', 'ユーザーID']);
  const idxStatus = findColumnIndex_(header, ['status', '出欠', '出欠状況']);
  
  if (idxEventKey === -1 || idxUserId === -1 || idxStatus === -1) {
    return { attended: false, status: '未回答' };
  }
  
  for (let row of rows) {
    const key = String(row[idxEventKey] || '').trim();
    const uid = String(row[idxUserId] || '').trim();
    const st = String(row[idxStatus] || '').trim();
    
    if (key === eventName && uid === userId) {
      const attended = (st === '○' || st === '出席');
      return { attended, status: st };
    }
  }
  
  return { attended: false, status: '未回答' };
}

/**
 * その人の紹介実績
 */
function getMyReferrals_(userId) {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(REFERRAL_SHEET_NAME);
  if (!sh) {
    return { count: 0, totalAmount: 0, closedCount: 0, hasChain: false };
  }
  
  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return { count: 0, totalAmount: 0, closedCount: 0, hasChain: false };
  }
  
  const header = values[0];
  const rows = values.slice(1);
  
  const idxFromId = header.indexOf('紹介者ID') !== -1 ? header.indexOf('紹介者ID') : 1; // B列
  const idxStatus = header.indexOf('ステータス') !== -1 ? header.indexOf('ステータス') : 7; // H列
  const idxAmount = header.indexOf('成約金額') !== -1 ? header.indexOf('成約金額') : 8; // I列
  
  let count = 0;
  let totalAmount = 0;
  let closedCount = 0;
  
  rows.forEach(row => {
    const fromId = String(row[idxFromId] || '').trim();
    if (fromId !== userId) return;
    
    count++;
    
    const status = String(row[idxStatus] || '').trim();
    const amount = Number(row[idxAmount]) || 0;
    
    if (status === '成約') {
      closedCount++;
      totalAmount += amount;
    }
  });
  
  // チェーン判定（簡易版：2件以上紹介してたらチェーンありとみなす）
  const hasChain = count >= 2;
  
  return { count, totalAmount, closedCount, hasChain };
}

/**
 * その人のランキング順位
 */
function getMyRanking_(userId) {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(REFERRAL_SHEET_NAME);
  if (!sh) {
    return { rank: null, total: 0 };
  }
  
  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return { rank: null, total: 0 };
  }
  
  const header = values[0];
  const rows = values.slice(1);
  
  const idxFromId = header.indexOf('紹介者ID') !== -1 ? header.indexOf('紹介者ID') : 1;
  
  // 会員ごとの紹介数を集計
  const countMap = {};
  rows.forEach(row => {
    const fromId = String(row[idxFromId] || '').trim();
    if (!fromId) return;
    countMap[fromId] = (countMap[fromId] || 0) + 1;
  });
  
  // ランキング作成
  const ranking = Object.entries(countMap)
    .map(([uid, cnt]) => ({ userId: uid, count: cnt }))
    .sort((a, b) => b.count - a.count);
  
  // 自分の順位を探す
  const myRankIndex = ranking.findIndex(r => r.userId === userId);
  const rank = myRankIndex !== -1 ? myRankIndex + 1 : null;
  const total = ranking.length;
  
  return { rank, total };
}
function testGetDashboardData() {
  const data = getDashboardData();
  
  Logger.log('eventName: ' + data.eventName);
  Logger.log('出席数: ' + (data.attendance ? data.attendance.attend : 'なし'));
  Logger.log('会員数: ' + (data.attendance ? data.attendance.memberCount : 'なし'));
  Logger.log('ゲスト数: ' + (data.guests ? data.guests.total : 'なし'));
  Logger.log('月次サマリー: ' + (data.monthlySummary && data.monthlySummary.monthly ? data.monthlySummary.monthly.length + 'ヶ月分' : 'なし'));
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function testGetDashboardData() {
  const data = getDashboardData();
  
  Logger.log('eventName: ' + data.eventName);
  Logger.log('出席数: ' + (data.attendance ? data.attendance.attend : 'なし'));
  Logger.log('会員数: ' + (data.attendance ? data.attendance.memberCount : 'なし'));
  Logger.log('ゲスト数: ' + (data.guests ? data.guests.total : 'なし'));
  Logger.log('月次サマリー: ' + (data.monthlySummary && data.monthlySummary.monthly ? data.monthlySummary.monthly.length + 'ヶ月分' : 'なし'));
}

/**
 * ダッシュボード用データをJSON安全な形で返す
 */
function getDashboardDataSafe() {
  const data = getDashboardData();
  // Date型を文字列に変換するためにJSON経由で返す
  return JSON.parse(JSON.stringify(data));
}

/**
 * 受付名簿用データを集計（出席者のみ）
 * @param {string} eventName - 例会名
 * @return {Object} { local: [...], others: { otherVenue: [...], guest: [...] } }
 */
function aggregateReceptionRoster_(eventName) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);

  // ========================================
  // 0-1. 配席アーカイブから座席情報を取得
  // ========================================
  const seatMap = {};  // 氏名 → テーブル記号
  const eventInfo = getCurrentEventInfo_();
  const targetEventKey = eventInfo.key;  // 202601_01 形式

  const shArchive = ss.getSheetByName(SEATING_ARCHIVE_SHEET_NAME);
  if (shArchive) {
    const archiveValues = shArchive.getDataRange().getValues();
    if (archiveValues.length >= 2) {
      // ヘッダー: 例会キー / 例会日 / 参加者ID / 氏名 / 区分 / 所属 / 役割 / チーム / テーブル / 席番 / 確定日時 / 確定者
      const archiveRows = archiveValues.slice(1);  // ヘッダー行をスキップ

      archiveRows.forEach(row => {
        const rowEventKey = String(row[0] || '').trim();  // A列: 例会キー
        if (rowEventKey !== targetEventKey) return;

        const name = String(row[3] || '').trim();   // D列: 氏名
        const table = String(row[8] || '').trim();  // I列: テーブル

        if (name && table) {
          seatMap[name] = table;
        }
      });
    }
  }

  // ========================================
  // 0-2. 出欠状況から出席者セットを作成
  // ========================================
  const attendeeByUserId = {};
  const attendeeByDisplayName = {};
  
  const shAttend = ss.getSheetByName(ATTEND_SHEET_NAME);
  if (shAttend) {
    const valuesA = shAttend.getDataRange().getValues();
    if (valuesA.length >= 4) {
      const HEADER_ROW_INDEX = 2;
      const headerA = valuesA[HEADER_ROW_INDEX];
      const rowsA = valuesA.slice(HEADER_ROW_INDEX + 1);
      
      const idxEventKey    = findColumnIndex_(headerA, ['eventKey', 'イベントキー']);
      const idxUserId      = findColumnIndex_(headerA, ['userId', 'LINE_userId', 'ユーザーID']);
      const idxDisplayName = findColumnIndex_(headerA, ['displayName', '氏名', '名前']);
      const idxStatus      = findColumnIndex_(headerA, ['status', '出欠', '出欠状況', 'ステータス']);
      
      rowsA.forEach(row => {
        if (idxEventKey === -1 || idxStatus === -1) return;
        
        const key = String(row[idxEventKey] || '').trim();
        if (key !== eventName) return;
        
        const st = String(row[idxStatus] || '').trim();
        // 出席者のみ
        if (st !== '○' && st !== '出席') return;
        
        const userId = idxUserId !== -1 ? String(row[idxUserId] || '').trim() : '';
        const displayName = idxDisplayName !== -1 ? String(row[idxDisplayName] || '').trim() : '';
        
        if (userId) attendeeByUserId[userId] = true;
        if (displayName) attendeeByDisplayName[displayName] = true;
      });
    }
  }
  
  // ========================================
  // 1. 福岡飯塚会員（出席者のみ）
  // ========================================
  const localList = [];
  const shMember = ss.getSheetByName(MEMBER_SHEET_NAME);
  
  if (shMember) {
    const values = shMember.getDataRange().getValues();
    if (values.length >= 3) {
      const rows = values.slice(2); // 3行目以降
      
      // 列インデックス（0始まり）
      const idxName        = 2;  // C列: 氏名
      const idxFurigana    = 3;  // D列: フリガナ
      const idxBusiness    = 10; // K列: 営業内容
      const idxDisplayName = 14; // O列: displayName
      const idxUserId      = 15; // P列: LINE_userId
      
      rows.forEach(row => {
        const name = String(row[idxName] || '').trim();
        if (!name) return;
        
        const userId = String(row[idxUserId] || '').trim();
        const displayName = String(row[idxDisplayName] || '').trim();
        
        // 出席者チェック
        const isAttendee = 
          (userId && attendeeByUserId[userId]) ||
          (displayName && attendeeByDisplayName[displayName]);
        
        if (!isAttendee) return;
        
        localList.push({
          seat: seatMap[name] || '',  // ★配席アーカイブから座席を取得
          name: name,
          furigana: String(row[idxFurigana] || '').trim(),
          business: String(row[idxBusiness] || '').trim(),
        });
      });
    }
  }
  
  // フリガナ順ソート
  localList.sort((a, b) => (a.furigana || '').localeCompare(b.furigana || '', 'ja'));
  
  // ========================================
  // 2. 他会場参加者（そのまま）
  // ========================================
  const otherVenueList = [];
  const shOther = ss.getSheetByName('他会場名簿マスター');
  
  if (shOther) {
    const valuesO = shOther.getDataRange().getValues();
    if (valuesO.length >= 3) {
      const rowsO = valuesO.slice(2);
      
      const cEvent    = 1;
      const cName     = 3;
      const cFurigana = 4;
      const cVenue    = 5;
      
      rowsO.forEach(row => {
        const event = String(row[cEvent] || '').trim();
        if (event !== eventName) return;
        
        const name = String(row[cName] || '').trim();
        if (!name) return;
        
        otherVenueList.push({
          seat: seatMap[name] || '',  // ★配席アーカイブから座席を取得
          name: name,
          furigana: String(row[cFurigana] || '').trim(),
          venue: String(row[cVenue] || '').trim(),
        });
      });
    }
  }
  
  otherVenueList.sort((a, b) => (a.furigana || '').localeCompare(b.furigana || '', 'ja'));
  
  // ========================================
  // 3. ゲスト参加者（そのまま）
  // ========================================
  const guestList = [];
  const shGuest = ss.getSheetByName(GUEST_SHEET_NAME);
  
  if (shGuest) {
    const valuesG = shGuest.getDataRange().getValues();
    if (valuesG.length >= 3) {
      const rowsG = valuesG.slice(2);
      
      const gcEvent    = 2;
      const gcName     = 3;
      const gcFurigana = 4;
      const gcCompany  = 5;
      
      rowsG.forEach(row => {
        const event = String(row[gcEvent] || '').trim();
        if (event !== eventName) return;
        
        const name = String(row[gcName] || '').trim();
        if (!name) return;
        
        guestList.push({
          seat: seatMap[name] || '',  // ★配席アーカイブから座席を取得
          name: name,
          furigana: String(row[gcFurigana] || '').trim(),
          company: String(row[gcCompany] || '').trim(),
        });
      });
    }
  }
  
  guestList.sort((a, b) => (a.furigana || '').localeCompare(b.furigana || '', 'ja'));
  
  return {
    local: localList,
    others: {
      otherVenue: otherVenueList,
      guest: guestList,
    }
  };
}

/**
 * 受付名簿データを取得（フロントから直接呼び出し用）
 * キャッシュに含めず、必要時にリアルタイム取得
 */
function getReceptionRosterData() {
  const eventInfo = getCurrentEventInfo_();
  const eventName = eventInfo.name;
  const reception = aggregateReceptionRoster_(eventName);
  
  return {
    eventName,
    reception
  };
}

// ===============================
// 式次第関連
// ===============================

/**
 * タイムテーブルマスターシート初期化
 * ※ 初回のみ実行
 */
function initTimeTableMaster() {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  
  // シートがなければ作成
  let sheet = ss.getSheetByName('タイムテーブルマスター');
  if (!sheet) {
    sheet = ss.insertSheet('タイムテーブルマスター');
  } else {
    sheet.clear();
  }
  
  // ========== タイムテーブル部分（A〜D列） ==========
  const timeTableHeader = [['時間', '項目', '担当者', '備考']];
  const timeTableData = [
    ['17:45', 'ゲスト説明', '田邊会員', '世話人'],
    ['18:10', 'ゲストさま入場', '', ''],
    ['18:15', 'オープニングDVD', '', ''],
    ['18:17', '開会宣言', '杉村会員', '副代表'],
    ['18:20', '代表挨拶', '髙山会員', '代表'],
    ['18:25', 'バッジ授与', '', ''],
    ['18:35', '他会場紹介', '', ''],
    ['18:40', 'ゲスト紹介', '', ''],
    ['18:45', '伝達事項報告', '', ''],
    ['18:55', '車座商談会', '', '各テーブルマスター'],
    ['19:20', 'ブース紹介', '', ''],
    ['19:25', '大名刺交換会', '', ''],
    ['19:35', '商談懇親会　乾杯', '', ''],
    ['20:45', '入会申し込み締め切り　入会希望者紹介', '', ''],
    ['20:50', '二部会案内', '', ''],
    ['20:55', '出発進行', '神谷会員', ''],
    ['21:00', '閉会', '', '']
  ];
  
  sheet.getRange(1, 1, 1, 4).setValues(timeTableHeader);
  sheet.getRange(2, 1, timeTableData.length, 4).setValues(timeTableData);
  
  // ヘッダー書式
  sheet.getRange(1, 1, 1, 4)
    .setBackground('#4a5568')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  // 列幅
  sheet.setColumnWidth(1, 80);   // 時間
  sheet.setColumnWidth(2, 300);  // 項目
  sheet.setColumnWidth(3, 100);  // 担当者
  sheet.setColumnWidth(4, 150);  // 備考
  
  // ========== お知らせ部分（F〜G列） ==========
  const noticeHeader = [['種別', '内容']];
  const noticeData = [
    ['来月案内', '◎ 来月の例会は○/○（○）　　原則　第二水曜日開催です！！'],
    ['キャンセル', '※ 参加予定の方で、前々日10時以降にキャンセルされた方は、キャンセル料として4000円頂きます。'],
    ['設営案内', '※ 毎月、例会当日 16：30～　会場設営を世話人さん中心で行ってます。\n会員さんも早く来れる方は是非、一緒にご準備しましょう！'],
    ['ゲスト連絡先', 'ゲストさまご紹介報告は公式ライン入力フォームよりお願いします。\n※ご不明の方の申し出先 TEL090-6423-7100（事務局 桑村）'],
    ['他会場連絡先', '他会場参加報告申し出先 TEL090-2503-4192（事務局 泊）'],
    ['車座補足', '（自己紹介・会社紹介・PRタイム一人3分×1周）'],
    ['懇親会補足', '「新会員さん紹介」　　（ブース出店　20:30まで）']
  ];
  
  sheet.getRange(1, 6, 1, 2).setValues(noticeHeader);
  sheet.getRange(2, 6, noticeData.length, 2).setValues(noticeData);
  
  // ヘッダー書式
  sheet.getRange(1, 6, 1, 2)
    .setBackground('#2b6cb0')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  // 列幅
  sheet.setColumnWidth(5, 30);   // 区切り
  sheet.setColumnWidth(6, 100);  // 種別
  sheet.setColumnWidth(7, 500);  // 内容
  
  Logger.log('タイムテーブルマスター初期化完了');
  return { success: true };
}


/**
 * 式次第データ取得（ダッシュボード用）
 */
function getShikidaiData() {
  try {
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);

    // 1. タイムテーブルマスターからデータ取得
    const ttSheet = ss.getSheetByName('タイムテーブルマスター');
  if (!ttSheet) {
    return { error: 'タイムテーブルマスターシートが見つかりません。initTimeTableMaster()を実行してください。' };
  }
  
  // タイムテーブル（A〜D列）- 最終行まで動的に取得
  const ttLastRow = ttSheet.getLastRow();
  const ttData = ttLastRow > 1 ? ttSheet.getRange(2, 1, ttLastRow - 1, 4).getValues() : [];
  const timeTable = ttData
    .filter(row => row[0] || row[1])
    .map(row => ({
      time: row[0] instanceof Date 
        ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'HH:mm')
        : (row[0] || ''),
      item: row[1] || '',
      person: row[2] || '',
      note: row[3] || ''
    }));
  
  // お知らせ（F〜G列）- 同じシートなのでttLastRowを使用
  const noticeData = ttLastRow > 1 ? ttSheet.getRange(2, 6, ttLastRow - 1, 2).getValues() : [];
  const notices = {};
  let mcPersons = ['', ''];
  noticeData.forEach(row => {
    if (row[0]) {
      // 司会進行は特別に処理
      if (row[0] === '司会進行') {
        const mcStr = row[1] || '';
        const parts = mcStr.split('・');
        mcPersons = [parts[0] || '', parts[1] || ''];
      } else {
        notices[row[0]] = row[1] || '';
      }
    }
  });
  
  // 2. 式次第シートから基本情報を取得
  const shikidaiSheet = ss.getSheetByName('式次第');
  let meetingTitle = '';
  let nextMeetingInfo = '';

  if (shikidaiSheet) {
    // タイトル（A2）
    meetingTitle = shikidaiSheet.getRange('A2').getValue() || '';

    // 来月案内（C4）
    nextMeetingInfo = shikidaiSheet.getRange('C4').getValue() || '';
  }

  // 他会場日程はキャッシュから取得
  const cacheResult = getOtherVenueCache();
  let otherVenueSchedule = cacheResult.success ? cacheResult.data : [];
  
  // 3. 設定シートから例会情報を取得
  let eventKey = '';
  let meetingCount = null;
  let currentEventDate = null;
  let nextEventDate = null;
  let nextEventDateStr = '';

  try {
    const configSs = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const settingSheet = configSs.getSheetByName('設定');
    if (settingSheet) {
      // イベントキー（A2）
      eventKey = settingSheet.getRange('A2').getValue() || '';

      // 例会回数を計算
      meetingCount = getMeetingCount(eventKey);

      // 年間日程を取得（D2:D13）
      const schedules = settingSheet.getRange('D2:D13').getValues()
        .flat()
        .filter(d => d instanceof Date);

      // 今日より後の最初の日程 = 今月の例会
      const today = new Date();
      today.setHours(0, 0, 0, 0);

      let currentIndex = schedules.findIndex(d => {
        const scheduleDate = new Date(d);
        scheduleDate.setHours(0, 0, 0, 0);
        return scheduleDate >= today;
      });
      if (currentIndex === -1) currentIndex = 0;

      currentEventDate = schedules[currentIndex] || null;
      nextEventDate = schedules[currentIndex + 1] || null;

      // 来月日程を「○/○（○）」形式にフォーマット
      if (nextEventDate) {
        nextEventDateStr = formatDateWithDay_(nextEventDate);
      }
    }
  } catch (e) {
    console.log('Failed to get settings:', e.message);
  }

  // 定義シートから開催回を取得（フォールバック用）
  const defSheet = ss.getSheetByName('定義');
  let meetingNo = '';
  if (defSheet) {
    meetingNo = defSheet.getRange('B2').getValue() || '';
  }

  // 4. 例会情報（getCurrentEventInfo_は既存ロジック用に残す）
  const eventInfo = getCurrentEventInfo_();

  // 5. 会員リスト（担当者選択用）
  const memberSheet = ss.getSheetByName(MEMBER_SHEET_NAME);
  let memberList = [];
  if (memberSheet) {
    const memberData = memberSheet.getRange('C2:C').getValues();
    memberList = memberData
      .filter(row => row[0])
      .map(row => String(row[0]).trim());
  }

  const result = {
    meetingTitle: meetingTitle,
    meetingNo: meetingNo,
    meetingCount: meetingCount,
    nextMeetingInfo: nextMeetingInfo,
    nextEventDate: nextEventDate,
    nextEventDateStr: nextEventDateStr,
    currentEventDate: currentEventDate,
    eventName: eventInfo.name,
    eventKey: eventKey || eventInfo.eventKey,
    eventDate: eventInfo.eventDate,
    timeTable: timeTable,
    notices: notices,
    mcPersons: mcPersons,
    otherVenueSchedule: otherVenueSchedule,
    memberList: memberList
  };
    console.log('getShikidaiData: returning result', JSON.stringify(result));
    return result;
  } catch (e) {
    console.error('getShikidaiData ERROR:', e.message, e.stack);
    return { error: e.message };
  }
}


/**
 * タイムテーブル保存
 */
function saveTimeTableData(data) {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const sheet = ss.getSheetByName('タイムテーブルマスター');
  
  if (!sheet) {
    return { success: false, error: 'シートが見つかりません' };
  }
  
  try {
    // タイムテーブル保存（A〜D列）
    if (data.timeTable && data.timeTable.length > 0) {
      // 既存データクリア（ヘッダー以外）
      const lastRow = sheet.getRange('A:A').getValues().filter(r => r[0]).length;
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
      }
      
      // 新データ書き込み
      const ttValues = data.timeTable.map(row => [
        row.time || '',
        row.item || '',
        row.person || '',
        row.note || ''
      ]);
      sheet.getRange(2, 1, ttValues.length, 4).setValues(ttValues);
    }
    
    // お知らせ保存（F〜G列）
    // 既存データクリア（ヘッダー以外）
    const noticeLastRow = sheet.getRange('F:F').getValues().filter(r => r[0]).length;
    if (noticeLastRow > 1) {
      sheet.getRange(2, 6, noticeLastRow - 1, 2).clearContent();
    }

    // お知らせデータを準備
    const noticeValues = [];

    // 通常のお知らせ
    if (data.notices) {
      Object.entries(data.notices).forEach(([key, value]) => {
        noticeValues.push([key, value]);
      });
    }

    // 司会進行を追加（「・」区切りで保存）
    if (data.mcPersons && (data.mcPersons[0] || data.mcPersons[1])) {
      const mcStr = data.mcPersons.filter(p => p).join('・');
      noticeValues.push(['司会進行', mcStr]);
    }

    // 書き込み
    if (noticeValues.length > 0) {
      sheet.getRange(2, 6, noticeValues.length, 2).setValues(noticeValues);
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function testGetShikidaiData() {
  const data = getShikidaiData();
  Logger.log('timeTable count: ' + (data.timeTable ? data.timeTable.length : 0));
  Logger.log('notices keys: ' + (data.notices ? Object.keys(data.notices).join(', ') : 'none'));
  Logger.log('memberList count: ' + (data.memberList ? data.memberList.length : 0));
  if (data.timeTable && data.timeTable.length > 0) {
    Logger.log('First row: ' + JSON.stringify(data.timeTable[0]));
  }
  if (data.error) {
    Logger.log('Error: ' + data.error);
  }
}

// ========================================
// 式次第自動取得機能
// ========================================

/**
 * イベントキーから例会回数を計算
 * 基準: 2025年11月 = 第91回
 */
function getMeetingCount(eventKey) {
  if (!eventKey || eventKey.length < 6) return null;

  const year = parseInt(eventKey.substring(0, 4));
  const month = parseInt(eventKey.substring(4, 6));

  // 基準: 2025年11月 = 第91回
  const baseYear = 2025;
  const baseMonth = 11;
  const baseCount = 91;

  const monthDiff = (year - baseYear) * 12 + (month - baseMonth);
  return baseCount + monthDiff;
}

/**
 * 日付を「○/○（○）」形式にフォーマット（タイムゾーン対応）
 */
function formatDateWithDay_(date) {
  if (!date) return '';
  const d = new Date(date);
  // Utilities.formatDateでタイムゾーンを正しく処理
  const formatted = Utilities.formatDate(d, 'Asia/Tokyo', 'M/d');
  const dayOfWeek = Utilities.formatDate(d, 'Asia/Tokyo', 'u'); // 1=月, 7=日
  const dayNames = ['', '月', '火', '水', '木', '金', '土', '日'];
  const dayName = dayNames[parseInt(dayOfWeek)];
  return `${formatted}（${dayName}）`;
}

/**
 * 他会場日程をスクレイピングして取得
 */
function scrapeOtherVenueSchedule() {
  // 九州・沖縄エリアの対象会場（福岡飯塚含む - 基準日特定用）
  const targetVenues = [
    '福岡飯塚', '福岡イースト', '福岡中央', '福岡筑後', '福岡',
    'ヒルノ福岡セントラル', '久留米セントラル', '久留米',
    '北九州八幡', '北九州門司', '北九州', '小倉', '宗像福津', '博多',
    'ながさき出島', '長崎いさはや',
    '熊本県北玉名', 'ヒルノ熊本',
    '中津諭吉の里', '延岡', '宮崎', '都城',
    '鹿児島南', '鳥栖',
    '沖縄北部やんばる', 'ヒルノ沖縄', '沖縄中部', '那覇', '沖縄', '琉球'
  ];

  try {
    const url = 'https://shusei-honbu.jp/shusei';
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const html = response.getContentText();

    // 全スケジュールを抽出（venue, dateStr, dateObj）
    const allSchedules = [];
    const today = new Date();
    const currentYear = today.getFullYear();
    const sortedVenues = [...targetVenues].sort((a, b) => b.length - a.length);

    // HTMLを行単位で処理
    const lines = html.split(/(?:<br\s*\/?>|\n|\r)+/);

    for (const line of lines) {
      const dateMatch = line.match(/(\d{1,2})月(\d{1,2})日\s*[（\(]([日月火水木金土])[）\)]/);
      if (!dateMatch) continue;

      const month = parseInt(dateMatch[1]);
      const day = parseInt(dateMatch[2]);
      const weekday = dateMatch[3];
      const dateStr = `${month}/${day}（${weekday}）`;

      // 年をまたぐ場合の処理（今月より6ヶ月以上前なら来年）
      let year = currentYear;
      if (month < today.getMonth() + 1 - 6) {
        year = currentYear + 1;
      }
      const dateObj = new Date(year, month - 1, day);

      for (const venue of sortedVenues) {
        if (line.includes(venue)) {
          allSchedules.push({ venue, dateStr, dateObj });
        }
      }
    }

    // テーブル行からも抽出
    const rowPattern = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    let rowMatch;

    while ((rowMatch = rowPattern.exec(html)) !== null) {
      const rowHtml = rowMatch[1];
      const dateMatch = rowHtml.match(/(\d{1,2})月(\d{1,2})日\s*[（\(]([日月火水木金土])[）\)]/);
      if (!dateMatch) continue;

      const month = parseInt(dateMatch[1]);
      const day = parseInt(dateMatch[2]);
      const weekday = dateMatch[3];
      const dateStr = `${month}/${day}（${weekday}）`;

      let year = currentYear;
      if (month < today.getMonth() + 1 - 6) {
        year = currentYear + 1;
      }
      const dateObj = new Date(year, month - 1, day);

      for (const venue of sortedVenues) {
        if (rowHtml.includes(venue)) {
          allSchedules.push({ venue, dateStr, dateObj });
        }
      }
    }

    // 1. 福岡飯塚の日程を取得して、今月・来月の例会日を特定
    const iizukaSchedules = allSchedules
      .filter(s => s.venue === '福岡飯塚')
      .sort((a, b) => a.dateObj - b.dateObj);

    // 今日以降の最初の福岡飯塚例会（今月例会）
    const currentSchedule = iizukaSchedules.find(s => s.dateObj >= today);
    if (!currentSchedule) {
      return { success: true, data: [], updatedAt: new Date().toISOString() };
    }

    // その次の福岡飯塚例会（来月例会）
    const currentIdx = iizukaSchedules.indexOf(currentSchedule);
    const nextSchedule = iizukaSchedules[currentIdx + 1];

    const baseDate = currentSchedule.dateObj;
    // 来月例会がなければ35日後を終了日とする
    const endDate = nextSchedule ? nextSchedule.dateObj : new Date(baseDate.getTime() + 35 * 24 * 60 * 60 * 1000);

    // 2. 各会場の基準日より後の最初の1回だけを抽出
    const venueFirstSchedule = {};

    for (const schedule of allSchedules) {
      // 福岡飯塚は除外
      if (schedule.venue === '福岡飯塚') continue;

      // 今月例会後 〜 来月例会前 の範囲
      if (schedule.dateObj > baseDate && schedule.dateObj < endDate) {
        // 各会場につき最初の1回だけ
        if (!venueFirstSchedule[schedule.venue] ||
            schedule.dateObj < venueFirstSchedule[schedule.venue].dateObj) {
          venueFirstSchedule[schedule.venue] = schedule;
        }
      }
    }

    const futureSchedules = Object.values(venueFirstSchedule);
    futureSchedules.sort((a, b) => a.dateObj - b.dateObj);

    // 3. 同じ日の会場をグループ化
    const groupedByDate = {};
    for (const schedule of futureSchedules) {
      const key = schedule.dateStr;
      if (!groupedByDate[key]) {
        groupedByDate[key] = [];
      }
      groupedByDate[key].push(schedule.venue);
    }

    // 4. 結果を作成（日付順）
    const scheduleData = [];
    const sortedDates = Object.keys(groupedByDate).sort((a, b) => {
      const [aM, aD] = a.match(/(\d+)\/(\d+)/).slice(1).map(Number);
      const [bM, bD] = b.match(/(\d+)\/(\d+)/).slice(1).map(Number);
      return aM !== bM ? aM - bM : aD - bD;
    });

    for (const date of sortedDates) {
      scheduleData.push({
        date: date,
        venues: groupedByDate[date].join(' / ')
      });
    }

    return {
      success: true,
      data: scheduleData,
      updatedAt: new Date().toISOString()
    };

  } catch (e) {
    console.error('Scraping error:', e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 他会場日程キャッシュを保存
 */
function saveOtherVenueCache(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    let cacheSheet = ss.getSheetByName('他会場日程キャッシュ');

    // シートがなければ作成
    if (!cacheSheet) {
      cacheSheet = ss.insertSheet('他会場日程キャッシュ');
      cacheSheet.getRange('A1').setValue('キャッシュデータ');
      cacheSheet.getRange('B1').setValue('更新日時');
    }

    // JSONで保存
    cacheSheet.getRange('A2').setValue(JSON.stringify(data));
    cacheSheet.getRange('B2').setValue(new Date());

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * 他会場日程キャッシュを取得
 */
function getOtherVenueCache() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const cacheSheet = ss.getSheetByName('他会場日程キャッシュ');

    if (!cacheSheet) {
      return { success: false, data: [], error: 'キャッシュシートが存在しません' };
    }

    const jsonStr = cacheSheet.getRange('A2').getValue();
    const updatedAt = cacheSheet.getRange('B2').getValue();

    if (!jsonStr) {
      return { success: false, data: [], error: 'キャッシュが空です' };
    }

    const data = JSON.parse(jsonStr);
    return {
      success: true,
      data: data,
      updatedAt: updatedAt
    };
  } catch (e) {
    return { success: false, data: [], error: e.message };
  }
}

/**
 * 他会場日程を更新（スクレイピング→キャッシュ保存）
 */
function updateOtherVenueSchedule() {
  const result = scrapeOtherVenueSchedule();

  if (!result.success) {
    return { success: false, error: 'スクレイピングに失敗しました: ' + result.error };
  }

  const saveResult = saveOtherVenueCache(result.data);

  if (!saveResult.success) {
    return { success: false, error: 'キャッシュ保存に失敗しました: ' + saveResult.error };
  }

  return {
    success: true,
    count: result.data.length,
    message: `${result.data.length}件の他会場日程を更新しました`,
    data: result.data
  };
}

/**
 * 来月の例会日を取得（他会場日程から福岡飯塚の次回日程を探す）
 */
function getNextMeetingDate() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const settingSheet = ss.getSheetByName('設定');
    if (settingSheet) {
      const nextDate = settingSheet.getRange('D2').getValue();
      if (nextDate instanceof Date) {
        return nextDate;
      }
    }
    return null;
  } catch (e) {
    console.error('getNextMeetingDate error:', e.message);
    return null;
  }
}


/** =========================================================
 *  配席アプリ用 API
 *  - getSeatingParticipants: 参加者一覧取得
 *  - getSeatingArchive: 過去配席取得
 *  - syncSeats: 座席確定・保存
 *  - listSeatingArchives: アーカイブ一覧
 * ======================================================= */

// 受付名簿シート名
const RECEPTION_MAIN_SHEET = '受付名簿（自動）';
const RECEPTION_OTHER_SHEET = '受付名簿（他会場・ゲスト）';

/**
 * 配席用参加者データを取得
 * @param {string} eventKey - 例会キー（省略時は現在の例会）
 * @returns {Object} 参加者データ
 */
function getSeatingParticipants(eventKey) {
  try {
    const eventInfo = getCurrentEventInfo_();
    const targetEventKey = eventKey || eventInfo.key;
    const targetEventName = eventKey ? eventKey : eventInfo.name;

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);

    // ========================================
    // 0. 出欠状況から出席者セットを作成
    // ========================================
    const attendeeMap = {}; // userId -> { displayName, booth }

    const shAttend = ss.getSheetByName(ATTEND_SHEET_NAME);
    if (shAttend) {
      const valuesA = shAttend.getDataRange().getValues();
      if (valuesA.length >= 4) {
        const HEADER_ROW_INDEX = 2;
        const headerA = valuesA[HEADER_ROW_INDEX];
        const rowsA = valuesA.slice(HEADER_ROW_INDEX + 1);

        const idxEventKey    = findColumnIndex_(headerA, ['eventKey', 'イベントキー']);
        const idxUserId      = findColumnIndex_(headerA, ['userId', 'LINE_userId', 'ユーザーID']);
        const idxDisplayName = findColumnIndex_(headerA, ['displayName', '氏名', '名前']);
        const idxStatus      = findColumnIndex_(headerA, ['status', '出欠', '出欠状況', 'ステータス']);
        const idxBooth       = findColumnIndex_(headerA, ['booth', 'ブース']);

        rowsA.forEach(row => {
          if (idxEventKey === -1 || idxStatus === -1) return;

          const key = String(row[idxEventKey] || '').trim();
          if (key !== targetEventName) return;

          const st = String(row[idxStatus] || '').trim();
          if (st !== '○' && st !== '出席') return;

          const userId = idxUserId !== -1 ? String(row[idxUserId] || '').trim() : '';
          const displayName = idxDisplayName !== -1 ? String(row[idxDisplayName] || '').trim() : '';
          const booth = idxBooth !== -1 ? String(row[idxBooth] || '').trim() : '';

          if (userId) {
            attendeeMap[userId] = { displayName, booth };
          }
        });
      }
    }

    const participants = [];
    let memberCount = 0;
    let guestCount = 0;
    let otherVenueCount = 0;

    // ========================================
    // 1. 福岡飯塚会員（出席者のみ）
    // ========================================
    const shMember = ss.getSheetByName(MEMBER_SHEET_NAME);
    if (shMember) {
      const values = shMember.getDataRange().getValues();
      if (values.length >= 3) {
        const rows = values.slice(2);

        // 列インデックス（0始まり）
        const idxId          = 0;   // A列: 会員ID
        const idxBadge       = 1;   // B列: バッジ
        const idxName        = 2;   // C列: 氏名
        const idxFurigana    = 3;   // D列: フリガナ
        const idxAffiliation = 4;   // E列: 所属
        const idxCompany     = 5;   // F列: 会社名
        const idxPosition    = 6;   // G列: 役職
        const idxRole        = 7;   // H列: 役割
        const idxTeam        = 8;   // I列: チーム
        const idxReferrer    = 9;   // J列: 紹介者
        const idxBusiness    = 10;  // K列: 営業内容
        const idxUserId      = 15;  // P列: LINE_userId

        rows.forEach(row => {
          const name = String(row[idxName] || '').trim();
          if (!name) return;

          const oderId = String(row[idxId] || '').trim();
          const userId = String(row[idxUserId] || '').trim();

          // 出席者チェック
          if (!userId || !attendeeMap[userId]) return;

          participants.push({
            id: oderId,
            oderId: oderId,
            name: name,
            furigana: String(row[idxFurigana] || '').trim(),
            category: '会員',
            affiliation: String(row[idxAffiliation] || '').trim() || '福岡飯塚',
            role: String(row[idxRole] || '').trim(),
            team: String(row[idxTeam] || '').trim(),
            business: String(row[idxBusiness] || '').trim(),
            booth: attendeeMap[userId].booth === '○' ? true : false,
            userId: userId,
            badge: String(row[idxBadge] || '').trim(),
            lastTable: '',
            lastSeat: ''
          });
          memberCount++;
        });
      }
    }

    // ========================================
    // 2. 他会場参加者
    // ========================================
    const shOther = ss.getSheetByName('他会場名簿マスター');
    if (shOther) {
      const valuesO = shOther.getDataRange().getValues();
      if (valuesO.length >= 3) {
        const rowsO = valuesO.slice(2);

        const cId       = 0;  // A列: ID
        const cEvent    = 1;  // B列: 例会名
        const cDate     = 2;  // C列: 登録日
        const cName     = 3;  // D列: 氏名
        const cFurigana = 4;  // E列: フリガナ
        const cVenue    = 5;  // F列: 所属

        rowsO.forEach((row, idx) => {
          const event = String(row[cEvent] || '').trim();
          if (event !== targetEventName) return;

          const name = String(row[cName] || '').trim();
          if (!name) return;

          // ★修正: 会員IDとの衝突を防ぐためプレフィックスを追加
          const rawId = String(row[cId] || '').trim();
          const id = rawId ? `other_${rawId}` : `other_${idx}`;

          participants.push({
            id: id,
            oderId: rawId || `${idx}`,
            name: name,
            furigana: String(row[cFurigana] || '').trim(),
            category: '他会場',
            affiliation: String(row[cVenue] || '').trim(),
            role: '',
            team: '',
            business: '',
            booth: false,
            userId: '',
            badge: '',
            lastTable: '',
            lastSeat: ''
          });
          otherVenueCount++;
        });
      }
    }

    // ========================================
    // 3. ゲスト参加者
    // ========================================
    const shGuest = ss.getSheetByName(GUEST_SHEET_NAME);
    if (shGuest) {
      const valuesG = shGuest.getDataRange().getValues();
      if (valuesG.length >= 3) {
        const rowsG = valuesG.slice(2);

        // シート構造: Guest_ID / timestamp / 例会名 / 氏名 / フリガナ / 会社名 / 役職 / 紹介者 / displayName / 営業内容 / 承認 / 区分
        const gcId       = 0;  // A列: Guest_ID
        const gcDate     = 1;  // B列: timestamp
        const gcEvent    = 2;  // C列: 例会名（例: 2026年1月例会）
        const gcName     = 3;  // D列: 氏名
        const gcFurigana = 4;  // E列: フリガナ
        const gcCompany  = 5;  // F列: 会社名
        const gcPosition = 6;  // G列: 役職
        const gcReferrer = 7;  // H列: 紹介者
        const gcDisplayName = 8;  // I列: displayName
        const gcBusiness = 9;  // J列: 営業内容
        const gcApproved = 10; // K列: 承認
        const gcCategory = 11; // L列: 区分（ゲスト/他会場）

        rowsG.forEach((row, idx) => {
          const eventName = String(row[gcEvent] || '').trim();
          if (eventName !== targetEventName) return;

          const name = String(row[gcName] || '').trim();
          if (!name) return;

          // L列の区分を取得（ゲスト or 他会場）
          const category = String(row[gcCategory] || '').trim();
          const rawId = String(row[gcId] || '').trim();
          // ★修正: 会員IDとの衝突を防ぐためプレフィックスを追加
          const id = rawId ? `guest_${rawId}` : `guest_${idx}`;

          if (category === 'ゲスト' || category === '') {
            // ゲストとして追加
            participants.push({
              id: id,
              oderId: rawId || `${idx}`,
              name: name,
              furigana: String(row[gcFurigana] || '').trim(),
              category: 'ゲスト',
              affiliation: String(row[gcCompany] || '').trim(),
              role: '',
              team: '',
              business: String(row[gcBusiness] || '').trim(),
              booth: false,
              userId: '',
              badge: '',
              lastTable: '',
              lastSeat: ''
            });
            guestCount++;
          } else if (category === '他会場') {
            // 他会場として追加
            participants.push({
              id: id,
              oderId: rawId || `${idx}`,
              name: name,
              furigana: String(row[gcFurigana] || '').trim(),
              category: '他会場',
              affiliation: String(row[gcCompany] || '').trim(),
              role: '',
              team: '',
              business: String(row[gcBusiness] || '').trim(),
              booth: false,
              userId: '',
              badge: '',
              lastTable: '',
              lastSeat: ''
            });
            otherVenueCount++;
          }
        });
      }
    }

    // フリガナ順ソート
    participants.sort((a, b) => (a.furigana || '').localeCompare(b.furigana || '', 'ja'));

    return {
      success: true,
      eventKey: targetEventKey,
      eventName: targetEventName,
      eventDate: '',
      eventTitle: targetEventName,
      participants: participants,
      summary: {
        total: participants.length,
        members: memberCount,
        guests: guestCount,
        otherVenue: otherVenueCount
      }
    };

  } catch (e) {
    console.error('getSeatingParticipants error:', e.message);
    return {
      success: false,
      error: e.message,
      participants: [],
      summary: { total: 0, members: 0, guests: 0, otherVenue: 0 }
    };
  }
}

/**
 * 過去の配席データを取得
 * @param {string} eventKey - 例会キー
 * @returns {Object} 配席アーカイブ
 */
function getSeatingArchive(eventKey) {
  try {
    if (!eventKey) {
      return { success: false, error: 'eventKey is required' };
    }

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sh = ss.getSheetByName(SEATING_ARCHIVE_SHEET_NAME);

    if (!sh) {
      return { success: false, error: 'アーカイブシートがありません' };
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) {
      return { success: true, eventKey: eventKey, assignments: [], tables: {} };
    }

    const [header, ...rows] = values;

    // 列インデックス取得
    const idx = {};
    header.forEach((h, i) => { idx[String(h || '').trim()] = i; });

    const assignments = [];
    let confirmedAt = '';
    let confirmedBy = '';
    let eventDate = '';
    let layout = [];
    const tables = {};

    rows.forEach(row => {
      const rowEventKey = String(row[idx['例会キー']] || '').trim();
      if (rowEventKey !== eventKey) return;

      if (!eventDate) eventDate = String(row[idx['例会日']] || '').trim();
      if (!confirmedAt) confirmedAt = String(row[idx['確定日時']] || '').trim();
      if (!confirmedBy) confirmedBy = String(row[idx['確定者']] || '').trim();

      // レイアウト情報（最初の行のみに保存されている）
      if (layout.length === 0 && idx['レイアウト'] !== undefined) {
        const layoutStr = String(row[idx['レイアウト']] || '').trim();
        if (layoutStr) {
          try {
            layout = JSON.parse(layoutStr);
          } catch (e) {
            // パースエラーは無視
          }
        }
      }

      const table = String(row[idx['テーブル']] || '').trim();
      const seat = parseInt(row[idx['席番']] || 0, 10);
      const name = String(row[idx['氏名']] || '').trim();
      const category = String(row[idx['区分']] || '').trim();

      assignments.push({
        id: String(row[idx['参加者ID']] || '').trim(),
        name: name,
        category: category,
        affiliation: String(row[idx['所属']] || '').trim(),
        role: String(row[idx['役割']] || '').trim(),
        team: String(row[idx['チーム']] || '').trim(),
        table: table,
        seat: seat
      });

      // テーブル別に整理
      if (!tables[table]) {
        tables[table] = { master: '', members: [] };
      }
      if (seat === 0) {
        tables[table].master = name;
      } else {
        tables[table].members.push(name);
      }
    });

    return {
      success: true,
      eventKey: eventKey,
      eventDate: eventDate,
      confirmedAt: confirmedAt,
      confirmedBy: confirmedBy,
      assignments: assignments,
      tables: tables,
      layout: layout
    };

  } catch (e) {
    console.error('getSeatingArchive error:', e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 座席情報を保存（受付名簿反映 + アーカイブ保存）
 * @param {Object} payload - 座席データ
 * @returns {Object} 結果
 */
function syncSeats(payload) {
  try {
    if (!payload || !payload.assignments || !Array.isArray(payload.assignments)) {
      return { success: false, error: 'assignments is required' };
    }

    const eventKey = payload.eventKey || '';
    const userId = payload.userId || '';
    const assignments = payload.assignments;
    const layout = payload.layout || [];  // テーブル配列パターン

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);

    // ========================================
    // 1. 受付名簿に反映
    // ========================================
    const mainSheet = ss.getSheetByName(RECEPTION_MAIN_SHEET);
    const otherSheet = ss.getSheetByName(RECEPTION_OTHER_SHEET);

    let updatedMain = 0;
    let updatedOther = 0;
    const unmatched = [];

    if (mainSheet && otherSheet) {
      const mainValues = mainSheet.getDataRange().getValues();
      const otherValues = otherSheet.getDataRange().getValues();

      // ★修正: ヘッダーは2行目（index 1）にある
      const mainHeader = mainValues[1] || [];
      const idxMainName = mainHeader.indexOf('氏名');
      const idxMainSeat = mainHeader.indexOf('座席');

      const otherHeader = otherValues[1] || [];
      const idxOtherName = otherHeader.indexOf('氏名');
      const idxOtherSeat = otherHeader.indexOf('座席');

      if (idxMainName >= 0 && idxMainSeat >= 0 && idxOtherName >= 0 && idxOtherSeat >= 0) {
        for (const a of assignments) {
          const name = String(a.name || '').trim();
          const table = String(a.table || '').trim();
          const category = String(a.category || '').trim();

          if (!name || !table) continue;

          if (category === '会員') {
            let found = false;
            // ★修正: データは3行目（index 2）から始まる
            for (let r = 2; r < mainValues.length; r++) {
              const rowName = String(mainValues[r][idxMainName] || '').trim();
              if (rowName === name) {
                mainSheet.getRange(r + 1, idxMainSeat + 1).setValue(table);
                updatedMain++;
                found = true;
                break;
              }
            }
            if (!found) {
              unmatched.push({ name, table, category, reason: 'not found in 会員受付名簿' });
            }
          } else {
            let found = false;
            // ★修正: データは3行目（index 2）から始まる
            for (let r = 2; r < otherValues.length; r++) {
              const rowName = String(otherValues[r][idxOtherName] || '').trim();
              if (rowName === name) {
                otherSheet.getRange(r + 1, idxOtherSeat + 1).setValue(table);
                updatedOther++;
                found = true;
                break;
              }
            }
            if (!found) {
              unmatched.push({ name, table, category, reason: 'not found in 他会場/ゲスト受付名簿' });
            }
          }
        }
      }
    }

    // ========================================
    // 2. 配席アーカイブに保存
    // ========================================
    let archivedCount = 0;
    const archiveSheet = ss.getSheetByName(SEATING_ARCHIVE_SHEET_NAME);

    if (archiveSheet && eventKey) {
      const now = new Date();
      const eventDate = '';  // TODO: 設定から取得

      // 同じeventKeyの既存データを削除
      const existingValues = archiveSheet.getDataRange().getValues();
      const rowsToDelete = [];
      for (let r = existingValues.length - 1; r >= 1; r--) {
        if (String(existingValues[r][0] || '').trim() === eventKey) {
          rowsToDelete.push(r + 1);
        }
      }
      rowsToDelete.forEach(rowNum => archiveSheet.deleteRow(rowNum));

      // 新しいデータを追加（最初の行にレイアウト情報を含める）
      const layoutStr = JSON.stringify(layout);
      const newRows = assignments.map((a, idx) => [
        eventKey,                                      // 例会キー
        eventDate,                                     // 例会日
        String(a.id || '').trim(),                     // 参加者ID
        String(a.name || '').trim(),                   // 氏名
        String(a.category || '').trim(),              // 区分
        String(a.affiliation || '').trim(),            // 所属
        String(a.role || '').trim(),                   // 役割
        String(a.team || '').trim(),                   // チーム
        String(a.table || '').trim(),                  // テーブル
        parseInt(a.seat || 0, 10),                     // 席番
        now,                                           // 確定日時
        userId,                                        // 確定者
        idx === 0 ? layoutStr : ''                     // レイアウト（最初の行のみ）
      ]);

      if (newRows.length > 0) {
        const lastRow = archiveSheet.getLastRow();
        archiveSheet.getRange(lastRow + 1, 1, newRows.length, 13).setValues(newRows);
        archivedCount = newRows.length;
      }
    }

    return {
      success: true,
      eventKey: eventKey,
      updatedMain: updatedMain,
      updatedOther: updatedOther,
      archivedCount: archivedCount,
      unmatched: unmatched
    };

  } catch (e) {
    console.error('syncSeats error:', e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 配席アーカイブの一覧を取得
 * @returns {Object} アーカイブ一覧
 */
function listSeatingArchives() {
  try {
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sh = ss.getSheetByName(SEATING_ARCHIVE_SHEET_NAME);

    if (!sh) {
      return { success: false, error: 'アーカイブシートがありません', archives: [] };
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) {
      return { success: true, archives: [] };
    }

    const [header, ...rows] = values;

    // 例会キーごとに集計
    const archiveMap = {};

    rows.forEach(row => {
      const eventKey = String(row[0] || '').trim();
      if (!eventKey) return;

      if (!archiveMap[eventKey]) {
        archiveMap[eventKey] = {
          eventKey: eventKey,
          eventDate: String(row[1] || '').trim(),
          confirmedAt: String(row[10] || '').trim(),
          count: 0
        };
      }
      archiveMap[eventKey].count++;
    });

    // 配列に変換してソート（新しい順）
    const archives = Object.values(archiveMap).sort((a, b) => {
      return b.eventKey.localeCompare(a.eventKey);
    });

    return {
      success: true,
      archives: archives
    };

  } catch (e) {
    console.error('listSeatingArchives error:', e.message);
    return { success: false, error: e.message, archives: [] };
  }
}

/**
 * POSTリクエストのエントリポイント
 */
function doPost(e) {
  // パラメータがない場合のガード
  if (!e || !e.parameter) {
    e = { parameter: {} };
  }

  const action = e.parameter.action || '';

  // ★配席アプリ用API: 座席確定・保存
  if (action === 'syncSeats') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const result = syncSeats(payload);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★役割分担: 保存
  if (action === 'saveRoleAssignments') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const result = saveRoleAssignments(payload.eventKey, payload.assignments);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★式次第: 保存
  if (action === 'saveTimeTableData') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const result = saveTimeTableData(payload);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★チェックイン: フィールド更新（ダッシュボード用）
  if (action === 'updateCheckinField') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const result = updateCheckinField(
        payload.rowIndex,
        payload.field,
        payload.value
      );
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★チェックイン: 手動チェックイン（ダッシュボード用）
  if (action === 'manualCheckin') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const result = manualCheckin(payload.rowIndex);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★会員編集機能: 会員情報更新
  if (action === 'updateMember') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const result = updateMember(payload.rowIndex, payload.data, payload.password);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★会員編集機能: 新規会員追加
  if (action === 'addMember') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const result = addMember(payload.data, payload.password);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★会員編集機能: 退会処理
  if (action === 'retireMember') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const result = retireMember(payload.rowIndex, payload.password);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★出欠リマインド機能: リマインド送信（本番）
  if (action === 'sendAttendanceReminder') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const result = sendAttendanceReminder(payload.userIds || []);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★出欠リマインド機能: トリガー設定
  if (action === 'setupReminderTriggers') {
    try {
      const jsonText = (e.postData && e.postData.contents) || '{}';
      const payload = JSON.parse(jsonText);
      const year = payload.year || new Date().getFullYear();
      const result = setupReminderTriggers(year);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'invalid JSON: ' + err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ★出欠リマインド機能: トリガー削除
  if (action === 'clearReminderTriggers') {
    try {
      const result = clearReminderTriggers();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: err.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 未対応のaction
  return ContentService
    .createTextOutput(JSON.stringify({
      success: false,
      error: 'unknown action: ' + action
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** =========================================================
 *  参加者名簿PDF出力
 * ======================================================= */

/**
 * Drive権限テスト用（GASエディタから実行して認証を通す）
 */
function testDriveAccess() {
  try {
    const folder = DriveApp.getFolderById(ROSTER_ARCHIVE_FOLDER_ID);
    Logger.log('フォルダ名: ' + folder.getName());
    Logger.log('フォルダURL: ' + folder.getUrl());
    return { success: true, folderName: folder.getName() };
  } catch (e) {
    Logger.log('エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 参加者名簿をスプレッドシートで生成してPDFをBase64で返す（ダウンロード用）
 * @returns {Object} { success, pdfBase64, fileName }
 */
function exportRosterPdf() {
  try {
    // 1. イベント情報取得
    const eventInfo = getCurrentEventInfo_();
    const eventKey = eventInfo.key;
    const eventName = eventInfo.name;

    if (!eventKey) {
      return { success: false, error: 'イベントキーが取得できません' };
    }

    // 2. 名簿データ取得
    const prepData = aggregatePrep_(eventName);
    const localList = prepData.local || [];
    const others = prepData.others || {};
    const otherVenueList = others.otherVenue || [];
    const guestList = others.guest || [];

    if (localList.length === 0 && otherVenueList.length === 0 && guestList.length === 0) {
      return { success: false, error: '名簿データがありません' };
    }

    // 3. 既存スプレッドシートに名簿シートを作成
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sheetName = '名簿_' + eventKey;

    // 既存シートがあれば削除して再作成
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet(sheetName);

    // 4. 名簿データをシートに描画
    const meetingNo = getMeetingCount(eventKey);
    renderRosterSheet_(sheet, localList, otherVenueList, guestList, eventName, meetingNo);

    // 5. PDF出力
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    const pdfBlob = exportSheetToPdf_(ss.getId(), sheet.getSheetId(), {
      orientation: 'landscape',
      paperSize: 'A4',
      fitToPage: true
    });

    // 6. Base64エンコードして返す
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    const fileName = '参加者名簿_' + eventKey + '.pdf';

    return {
      success: true,
      pdfBase64: pdfBase64,
      fileName: fileName,
      sheetUrl: ss.getUrl() + '#gid=' + sheet.getSheetId()
    };

  } catch (e) {
    console.error('exportRosterPdf error:', e);
    return { success: false, error: e.message };
  }
}

/**
 * フォルダ内にサブフォルダを取得/作成
 */
function getOrCreateFolder_(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

/**
 * 名簿シートを描画（完成形に整形）
 * html2canvas版と同じレイアウトに合わせる
 */
function renderRosterSheet_(sheet, localList, otherVenueList, guestList, eventName, meetingNo) {
  // ヘッダー（縦書き風に短縮）
  const headers = ['受賞', '紹介', 'バッジ', '出欠', 'ブース', '氏名', 'フリガナ', '役割', '会社名', '役職', '紹介者', '営業内容'];
  // 列幅（html2canvas版に合わせる: 22,22,22,22,25,70,100,56,180,155,70,auto）
  const colWidths = [25, 25, 25, 25, 28, 75, 105, 60, 190, 160, 75, 250];

  // 列幅設定
  colWidths.forEach((w, i) => {
    sheet.setColumnWidth(i + 1, w);
  });

  let currentRow = 1;

  // === タイトル行 ===
  const roundText = meetingNo ? `第${meetingNo}回 ` : '';
  const title = `守成クラブ福岡飯塚　${roundText}仕事バンバンプラザ 例会　参加者名簿`;
  sheet.getRange(currentRow, 1).setValue(title);
  sheet.getRange(currentRow, 1, 1, headers.length).merge();
  sheet.getRange(currentRow, 1)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 28);
  currentRow++;

  // === サブタイトル行（左: 稼働会員数・参加者数、右: 開催日）===
  const attendCount = localList.filter(r => r.attendStatus === '○').length;
  const memberCount = localList.length;

  // 左側の情報
  sheet.getRange(currentRow, 1).setValue(`稼働会員数：${memberCount}　会員参加者：${attendCount}`);
  sheet.getRange(currentRow, 1, 1, 6).merge();
  sheet.getRange(currentRow, 1).setFontSize(9).setHorizontalAlignment('left');

  // 右側の情報（開催日）
  sheet.getRange(currentRow, 7).setValue(`開催日：${eventName}`);
  sheet.getRange(currentRow, 7, 1, 6).merge();
  sheet.getRange(currentRow, 7).setFontSize(9).setHorizontalAlignment('right');

  sheet.setRowHeight(currentRow, 18);
  currentRow++;

  // === ヘッダー行 ===
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(currentRow, 1, 1, headers.length);
  headerRange
    .setFontWeight('bold')
    .setFontSize(8)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(null)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(currentRow, 22);

  const headerStartRow = currentRow;
  currentRow++;
  const dataStartRow = currentRow;

  // === 福岡飯塚会員データ ===
  localList.forEach(row => {
    const boothVal = (row.booth === '○' || row.booth === true) ? '○' : '';
    const values = [
      row.award || '',
      row.referralCnt != null ? row.referralCnt : '',
      row.badge || '',
      row.attendStatus || '',
      boothVal,
      row.name || '',
      row.furigana || '',
      row.role || '',
      row.company || '',
      row.position || '',
      row.introducer || '',
      row.business || ''
    ];
    sheet.getRange(currentRow, 1, 1, values.length).setValues([values]);
    sheet.setRowHeight(currentRow, 18);
    currentRow++;
  });

  const localEndRow = currentRow - 1;

  // === 他会場セクション（あれば）===
  if (otherVenueList.length > 0) {
    currentRow++; // 空行
    sheet.setRowHeight(currentRow - 1, 10);

    sheet.getRange(currentRow, 1).setValue('【他会場参加者】　' + otherVenueList.length + '名');
    sheet.getRange(currentRow, 1, 1, headers.length).merge();
    sheet.getRange(currentRow, 1).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('left');
    sheet.setRowHeight(currentRow, 20);
    currentRow++;

    // 他会場ヘッダー
    const otherHeaders = ['バッジ', '氏名', 'フリガナ', '所属', '役割', '会社名', '役職', '紹介者', '営業内容', '', '', ''];
    sheet.getRange(currentRow, 1, 1, otherHeaders.length).setValues([otherHeaders]);
    sheet.getRange(currentRow, 1, 1, 9)
      .setFontWeight('bold')
      .setFontSize(8)
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(currentRow, 20);
    currentRow++;

    const otherDataStart = currentRow;
    otherVenueList.forEach(row => {
      const values = [
        row.badge || '',
        row.name || '',
        row.furigana || '',
        row.venue || '',
        row.role || '',
        row.company || '',
        row.position || '',
        row.introducer || '',
        row.business || '',
        '', '', ''
      ];
      sheet.getRange(currentRow, 1, 1, values.length).setValues([values]);
      sheet.setRowHeight(currentRow, 18);
      currentRow++;
    });

    // 他会場データに罫線
    if (currentRow > otherDataStart) {
      sheet.getRange(otherDataStart, 1, currentRow - otherDataStart, 9)
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  // === ゲストセクション（あれば）===
  if (guestList.length > 0) {
    currentRow++; // 空行
    sheet.setRowHeight(currentRow - 1, 10);

    sheet.getRange(currentRow, 1).setValue('【ゲスト参加者】　' + guestList.length + '名');
    sheet.getRange(currentRow, 1, 1, headers.length).merge();
    sheet.getRange(currentRow, 1).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('left');
    sheet.setRowHeight(currentRow, 20);
    currentRow++;

    // ゲストヘッダー
    const guestHeaders = ['紹介者', '氏名', 'フリガナ', '会社名', '役職', '営業内容', '', '', '', '', '', ''];
    sheet.getRange(currentRow, 1, 1, guestHeaders.length).setValues([guestHeaders]);
    sheet.getRange(currentRow, 1, 1, 6)
      .setFontWeight('bold')
      .setFontSize(8)
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(currentRow, 20);
    currentRow++;

    const guestDataStart = currentRow;
    guestList.forEach(row => {
      const values = [
        row.introducer || '',
        row.name || '',
        row.furigana || '',
        row.company || '',
        row.position || '',
        row.business || '',
        '', '', '', '', '', ''
      ];
      sheet.getRange(currentRow, 1, 1, values.length).setValues([values]);
      sheet.setRowHeight(currentRow, 18);
      currentRow++;
    });

    // ゲストデータに罫線
    if (currentRow > guestDataStart) {
      sheet.getRange(guestDataStart, 1, currentRow - guestDataStart, 6)
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  // === 全体スタイル適用 ===
  const lastRow = currentRow - 1;
  const lastCol = headers.length;

  // 全体にフォント設定
  const allRange = sheet.getRange(1, 1, lastRow, lastCol);
  allRange.setFontFamily('Noto Sans JP');
  allRange.setVerticalAlignment('middle');

  // データ部分のフォントサイズ
  if (localEndRow >= dataStartRow) {
    sheet.getRange(dataStartRow, 1, localEndRow - dataStartRow + 1, lastCol).setFontSize(9);
  }

  // 福岡飯塚会員データに罫線
  if (localEndRow >= dataStartRow) {
    sheet.getRange(dataStartRow, 1, localEndRow - dataStartRow + 1, lastCol)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      .setBackground(null);
  }

  // 中央揃え列（受賞, 紹介, バッジ, 出欠, ブース）
  const centerCols = [1, 2, 3, 4, 5];
  centerCols.forEach(col => {
    if (localEndRow >= dataStartRow) {
      sheet.getRange(dataStartRow, col, localEndRow - dataStartRow + 1, 1).setHorizontalAlignment('center');
    }
  });
}

/**
 * スプレッドシートの特定シートをPDFにエクスポート
 */
function exportSheetToPdf_(spreadsheetId, sheetId, options) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
    'format=pdf' +
    '&gid=' + sheetId +
    '&portrait=' + (options.orientation === 'portrait' ? 'true' : 'false') +
    '&size=' + (options.paperSize || 'A4') +
    '&fitw=true' +  // 幅に合わせる
    '&fith=false' +
    '&gridlines=false' +
    '&printtitle=false' +
    '&sheetnames=false' +
    '&pagenum=false' +
    '&fzr=false' +  // 固定行なし
    '&fzc=false' +  // 固定列なし
    '&top_margin=0.3' +
    '&bottom_margin=0.3' +
    '&left_margin=0.3' +
    '&right_margin=0.3';

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export failed: ' + response.getContentText());
  }

  return response.getBlob().setName('roster.pdf');
}

/** =========================================================
 *  個別売上機能（パスワード保護）
 * ======================================================= */

/**
 * 個別売上パスワードを認証（公開関数）
 * 設定シートのJ2セルにパスワードを保存
 */
function verifySalesPassword(password) {
  if (!password) {
    return { success: false, message: 'パスワードを入力してください' };
  }

  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sh) {
    return { success: false, message: '設定シートが見つかりません' };
  }

  // J2セルからパスワードを取得
  const correctPassword = String(sh.getRange('J2').getValue() || '').trim();
  if (!correctPassword) {
    return { success: false, message: 'パスワードが設定されていません' };
  }

  if (password === correctPassword) {
    return { success: true };
  } else {
    return { success: false, message: 'パスワードが違います' };
  }
}

/**
 * 売上報告の月一覧を取得（公開関数）
 */
function getSalesMonthList() {
  const ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  const sh = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sh) {
    return { success: false, months: [], message: '売上シートが見つかりません' };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 7) {
    return { success: true, months: [] };
  }

  const values = sh.getRange(7, 15, lastRow - 6, 1).getValues(); // O列のみ
  const monthList = [];
  const seen = new Set();

  values.forEach(row => {
    const eventKey = String(row[0] || '').trim();
    const match = eventKey.match(/^(\d{4})年(\d{1,2})月_売上報告$/);
    if (match) {
      const year = Number(match[1]);
      const month = Number(match[2]);
      const key = `${year}-${month}`;
      if (!seen.has(key)) {
        seen.add(key);
        monthList.push({ year, month, label: `${year}年${month}月`, key });
      }
    }
  });

  // 年月を昇順でソート
  monthList.sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.month - b.month;
  });

  return { success: true, months: monthList };
}

/**
 * 指定月の個別売上データを取得（公開関数）
 * @param {string} yearMonth - "2025-11" 形式の年月キー
 * @param {string} password - パスワード
 */
function getIndividualSales(yearMonth, password) {
  // パスワード認証
  const authResult = verifySalesPassword(password);
  if (!authResult.success) {
    return authResult;
  }

  if (!yearMonth) {
    return { success: false, message: '年月を指定してください' };
  }

  // "2025-11" → "2025年11月_売上報告"
  const parts = yearMonth.split('-');
  if (parts.length !== 2) {
    return { success: false, message: '年月フォーマットが不正です' };
  }
  const year = parts[0];
  const month = parts[1];
  const salesEventKey = `${year}年${month}月_売上報告`;

  const ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  const sh = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sh) {
    return { success: false, message: '売上シートが見つかりません' };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 7) {
    return { success: true, data: [], summary: { total: 0, dealCount: 0, meetingsCount: 0, reporterCount: 0 } };
  }

  const values = sh.getRange(7, 1, lastRow - 6, sh.getLastColumn()).getValues();
  const data = [];
  let total = 0, dealCount = 0, meetingsCount = 0;

  values.forEach(row => {
    const rowEventKey = String(row[14] || '').trim(); // O列
    if (rowEventKey !== salesEventKey) return;

    const name = String(row[13] || '').trim();     // N列：氏名
    const meetings = Number(row[15]) || 0;          // P列：商談件数
    const deals = Number(row[4]) || 0;              // E列：成約件数
    const amount = Number(row[5]) || 0;             // F列：売上金額

    data.push({ name, meetings, deals, amount });
    total += amount;
    dealCount += deals;
    meetingsCount += meetings;
  });

  // 売上金額で降順ソート
  data.sort((a, b) => b.amount - a.amount);

  return {
    success: true,
    month: Number(month),
    data,
    summary: {
      total,
      dealCount,
      meetingsCount,
      reporterCount: data.length
    }
  };
}

/** =========================================================
 *  例会参加履歴API
 * ======================================================= */

/**
 * 例会キーの一覧を取得（ドロップダウン用）
 * @returns {Object} 例会キー一覧
 */
function getEventKeyList() {
  try {
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sh = ss.getSheetByName(SEATING_ARCHIVE_SHEET_NAME);

    if (!sh) {
      return { success: false, eventKeys: [], message: 'アーカイブシートがありません' };
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return { success: true, eventKeys: [] };
    }

    // A列（例会キー）を取得
    const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();

    // ユニークな例会キーを取得
    const eventKeySet = new Set();
    values.forEach(row => {
      const key = String(row[0] || '').trim();
      if (key) eventKeySet.add(key);
    });

    // 配列に変換してソート（新しい順）
    const eventKeys = Array.from(eventKeySet).sort().reverse();

    return { success: true, eventKeys };
  } catch (e) {
    Logger.log('getEventKeyList error: ' + e.message);
    return { success: false, eventKeys: [], message: e.message };
  }
}

/**
 * 指定例会の参加者履歴を取得
 * @param {string} eventKey - 例会キー（例: "202601_01"）
 * @returns {Object} 参加者データ
 */
function getParticipantHistory(eventKey) {
  try {
    if (!eventKey) {
      return { success: false, message: 'eventKeyを指定してください' };
    }

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sh = ss.getSheetByName(SEATING_ARCHIVE_SHEET_NAME);

    if (!sh) {
      return { success: false, message: 'アーカイブシートがありません' };
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return {
        success: true,
        eventKey,
        total: 0,
        members: 0,
        guests: 0,
        others: 0,
        participants: []
      };
    }

    // 全データ取得（ヘッダー除く）
    const data = sh.getRange(2, 1, lastRow - 1, 9).getValues();

    // 例会キーでフィルタ
    const participants = [];
    data.forEach(row => {
      const rowEventKey = String(row[0] || '').trim();
      if (rowEventKey !== eventKey) return;

      participants.push({
        name: String(row[3] || '').trim(),      // D列: 氏名
        category: String(row[4] || '').trim(),  // E列: 区分
        affiliation: String(row[5] || '').trim(), // F列: 所属
        role: String(row[6] || '').trim(),      // G列: 役割
        team: String(row[7] || '').trim(),      // H列: チーム
        table: String(row[8] || '').trim()      // I列: テーブル
      });
    });

    // テーブル順 → 氏名順でソート
    participants.sort((a, b) => {
      if (a.table !== b.table) {
        return a.table.localeCompare(b.table);
      }
      return a.name.localeCompare(b.name, 'ja');
    });

    // 集計
    const members = participants.filter(p => p.category === '会員').length;
    const guests = participants.filter(p => p.category === 'ゲスト').length;
    const others = participants.filter(p => p.category === '他会場').length;

    return {
      success: true,
      eventKey,
      total: participants.length,
      members,
      guests,
      others,
      participants
    };
  } catch (e) {
    Logger.log('getParticipantHistory error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/** =========================================================
 *  役割分担管理
 * ======================================================= */

// デフォルトの役割リスト（group: 同じ役割名をグループ化、count: 枠数）
const DEFAULT_ROLES = [
  { id: '統括', name: '統括', order: 1, group: '統括' },
  { id: 'PA', name: 'PA', order: 2, group: 'PA' },
  { id: 'MC1', name: 'MC', order: 3, group: 'MC' },
  { id: 'MC2', name: 'MC', order: 4, group: 'MC' },
  { id: 'バッジ授与', name: 'バッジ授与', order: 5, group: 'バッジ授与' },
  { id: '受付1', name: '受付', order: 6, group: '受付' },
  { id: '受付2', name: '受付', order: 7, group: '受付' },
  { id: '受付3', name: '受付', order: 8, group: '受付' },
  { id: '受付4', name: '受付', order: 9, group: '受付' },
  { id: 'ゲスト誘導1', name: 'ゲスト誘導', order: 10, group: 'ゲスト誘導' },
  { id: 'ゲスト誘導2', name: 'ゲスト誘導', order: 11, group: 'ゲスト誘導' },
  { id: 'ゲスト誘導3', name: 'ゲスト誘導', order: 12, group: 'ゲスト誘導' },
  { id: 'ゲスト誘導4', name: 'ゲスト誘導', order: 13, group: 'ゲスト誘導' },
  { id: '会員サポート1', name: '会員サポート', order: 14, group: '会員サポート' },
  { id: '会員サポート2', name: '会員サポート', order: 15, group: '会員サポート' },
  { id: '会員サポート3', name: '会員サポート', order: 16, group: '会員サポート' },
  { id: '会員サポート4', name: '会員サポート', order: 17, group: '会員サポート' },
  { id: '名刺回収1', name: '名刺回収', order: 18, group: '名刺回収' },
  { id: '名刺回収2', name: '名刺回収', order: 19, group: '名刺回収' },
  { id: '名刺回収3', name: '名刺回収', order: 20, group: '名刺回収' },
  { id: '名刺回収4', name: '名刺回収', order: 21, group: '名刺回収' },
  { id: '撮影', name: '撮影', order: 22, group: '撮影' }
];

/**
 * デフォルト役割一覧を取得
 */
function getDefaultRoles() {
  return {
    success: true,
    roles: DEFAULT_ROLES
  };
}

/**
 * 役割分担シートを作成（なければ新規、あればスキップ）
 */
function createRoleAssignmentSheet() {
  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  let sh = ss.getSheetByName(ROLE_ASSIGNMENT_SHEET_NAME);

  if (sh) {
    Logger.log('シート既存: ' + ROLE_ASSIGNMENT_SHEET_NAME);
    return { success: true, message: 'シート既存' };
  }

  sh = ss.insertSheet(ROLE_ASSIGNMENT_SHEET_NAME);

  const headers = ['例会キー', '役割ID', '役割名', '担当者名', '更新日時'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sh.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#6aa84f');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sh.setColumnWidth(1, 100);  // 例会キー
  sh.setColumnWidth(2, 100);  // 役割ID
  sh.setColumnWidth(3, 100);  // 役割名
  sh.setColumnWidth(4, 120);  // 担当者名
  sh.setColumnWidth(5, 150);  // 更新日時

  sh.setFrozenRows(1);

  Logger.log('シート作成完了: ' + ROLE_ASSIGNMENT_SHEET_NAME);
  return { success: true, message: 'シート作成完了' };
}

/**
 * 指定例会の役割分担を取得
 * @param {string} eventKey - 例会キー（例: "202601_01"）
 * @returns {Object} 役割分担データ
 */
function getRoleAssignments(eventKey) {
  try {
    if (!eventKey) {
      // eventKey未指定の場合は現在のイベントキーを使用
      eventKey = getDashboardEventKey();
    }
    Logger.log('getRoleAssignments: eventKey=' + eventKey);

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    let sh = ss.getSheetByName(ROLE_ASSIGNMENT_SHEET_NAME);

    // シートがなければ作成
    if (!sh) {
      createRoleAssignmentSheet();
      sh = ss.getSheetByName(ROLE_ASSIGNMENT_SHEET_NAME);
    }

    const lastRow = sh.getLastRow();
    Logger.log('getRoleAssignments: lastRow=' + lastRow);

    // デフォルト役割をベースに作成
    const assignments = DEFAULT_ROLES.map(role => ({
      roleId: role.id,
      roleName: role.name,
      assignee: '',
      order: role.order
    }));

    if (lastRow >= 2) {
      // 既存データを取得
      const data = sh.getRange(2, 1, lastRow - 1, 4).getValues();
      Logger.log('getRoleAssignments: 既存データ行数=' + data.length);

      // 例会キーでフィルタして担当者を設定
      let matchCount = 0;
      data.forEach(row => {
        const rowEventKey = String(row[0] || '').trim();
        if (rowEventKey !== eventKey) return;

        const roleId = String(row[1] || '').trim();
        const assignee = String(row[3] || '').trim();

        // 対応する役割を探して担当者を設定
        const assignment = assignments.find(a => a.roleId === roleId);
        if (assignment) {
          assignment.assignee = assignee;
          matchCount++;
          Logger.log('getRoleAssignments: マッチ ' + roleId + '=' + assignee);
        }
      });
      Logger.log('getRoleAssignments: マッチ数=' + matchCount);
    }

    // order順でソート
    assignments.sort((a, b) => a.order - b.order);

    const assignedCount = assignments.filter(a => a.assignee).length;
    Logger.log('getRoleAssignments: 担当設定数=' + assignedCount);

    return {
      success: true,
      eventKey,
      assignments
    };
  } catch (e) {
    Logger.log('getRoleAssignments error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * 役割分担を保存
 * @param {Object} payload - { eventKey, assignments: [{roleId, roleName, assignee}] }
 * @returns {Object} 結果
 */
function saveRoleAssignments(eventKey, assignments) {
  Logger.log('=== saveRoleAssignments 開始 ===');
  try {
    Logger.log('saveRoleAssignments: eventKey=' + eventKey);
    Logger.log('saveRoleAssignments: assignments type=' + typeof assignments);
    Logger.log('saveRoleAssignments: assignments length=' + (assignments ? assignments.length : 'null'));

    // パラメータチェック
    if (!eventKey) {
      Logger.log('saveRoleAssignments: eventKeyが空');
      return { success: false, error: 'eventKeyが空です' };
    }

    if (!assignments || !Array.isArray(assignments)) {
      Logger.log('saveRoleAssignments: assignmentsが無効');
      return { success: false, error: 'assignmentsが無効です' };
    }

    Logger.log('saveRoleAssignments: assignments数=' + assignments.length);

    // 担当者が設定されているものをログ
    const withAssignee = assignments.filter(a => a.assignee && a.assignee.trim() !== '');
    Logger.log('saveRoleAssignments: 担当設定数=' + withAssignee.length);
    withAssignee.forEach(a => {
      Logger.log('saveRoleAssignments: 保存対象 ' + a.roleId + '=' + a.assignee);
    });

    if (false && !eventKey) { // 既にチェック済みなのでスキップ
      return { success: false, error: 'eventKeyを指定してください' };
    }

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    let sh = ss.getSheetByName(ROLE_ASSIGNMENT_SHEET_NAME);

    // シートがなければ作成
    if (!sh) {
      createRoleAssignmentSheet();
      sh = ss.getSheetByName(ROLE_ASSIGNMENT_SHEET_NAME);
    }

    const lastRow = sh.getLastRow();
    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    // 既存の該当eventKeyの行を削除（更新用）
    if (lastRow >= 2) {
      const data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      // 下から削除していく（行番号ずれ防止）
      for (let i = data.length - 1; i >= 0; i--) {
        const rowEventKey = String(data[i][0] || '').trim();
        if (rowEventKey === eventKey) {
          sh.deleteRow(i + 2); // ヘッダー分 +1, 0-index分 +1
        }
      }
    }

    // 新しいデータを追加
    const newRows = [];
    assignments.forEach(a => {
      if (a.assignee && a.assignee.trim() !== '') {
        newRows.push([
          eventKey,
          a.roleId,
          a.roleName,
          a.assignee.trim(),
          timestamp
        ]);
      }
    });

    Logger.log('saveRoleAssignments: newRows数=' + newRows.length);

    if (newRows.length > 0) {
      const insertRow = sh.getLastRow() + 1;
      sh.getRange(insertRow, 1, newRows.length, 5).setValues(newRows);
    }

    return {
      success: true,
      message: `${newRows.length}件の役割分担を保存しました`,
      savedCount: newRows.length
    };
  } catch (e) {
    Logger.log('saveRoleAssignments error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 役割候補者一覧を取得（会員名簿マスターの「役割」列が空でない人）
 * @returns {Object} { success, candidates: [{name, role}] }
 */
function getRoleCandidates() {
  try {
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sh = ss.getSheetByName(MEMBER_SHEET_NAME);

    if (!sh) {
      return { success: false, message: '会員名簿マスターが見つかりません', candidates: [] };
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 3) {
      return { success: true, candidates: [] };
    }

    // ヘッダー行（2行目）から列インデックスを取得
    const header = values[1];
    const idxName = 2;  // C列：氏名
    const idxRole = findColumnIndex_(header, ['役割']);

    if (idxRole < 0) {
      return { success: false, message: '役割列が見つかりません', candidates: [] };
    }

    const candidates = [];
    const rows = values.slice(2);  // 3行目以降

    rows.forEach(row => {
      const name = String(row[idxName] || '').trim();
      const role = String(row[idxRole] || '').trim();

      // 役割列が空でない人だけ追加
      if (name && role) {
        candidates.push({ name, role });
      }
    });

    // 名前順でソート
    candidates.sort((a, b) => a.name.localeCompare(b.name, 'ja'));

    return { success: true, candidates };
  } catch (e) {
    Logger.log('getRoleCandidates error: ' + e.message);
    return { success: false, message: e.message, candidates: [] };
  }
}

/** =========================================================
 *  Googleフォーム回答 → 他会場名簿マスター 転記
 * ======================================================= */

// フォーム回答シート名（最新のシート）
const FORM_RESPONSE_SHEET_NAME = 'フォームの回答 10';
const OTHER_VENUE_MASTER_SHEET_NAME = '他会場名簿マスター';

/**
 * 日付文字列からEVENT_KEYに変換
 * 例: 「3月11日（水）」→「2026年3月例会」
 * @param {string} dateStr - 日付文字列
 * @returns {string} EVENT_KEY形式
 */
function convertDateToEventKey_(dateStr) {
  if (!dateStr) return '';

  // 「3月11日（水）」形式をパース
  const match = dateStr.match(/(\d+)月(\d+)日/);
  if (!match) return dateStr;

  const month = parseInt(match[1], 10);
  const day = parseInt(match[2], 10);

  // 現在の年を基準に判定（12月→来年1月なら翌年）
  const now = new Date();
  let year = now.getFullYear();

  // 現在が12月で、対象が1-3月なら翌年
  if (now.getMonth() === 11 && month <= 3) {
    year++;
  }
  // 現在が1-2月で、対象が11-12月なら前年
  if (now.getMonth() <= 1 && month >= 11) {
    year--;
  }

  return `${year}年${month}月例会`;
}

/**
 * バッジを短縮形に変換
 * 例: 「正会員」→「正」、「ゴールド」→「ゴ」
 * @param {string} badge - バッジ名
 * @returns {string} 短縮形
 */
function convertBadge_(badge) {
  if (!badge) return '';
  const normalized = String(badge).trim();

  const mapping = {
    '正会員': '正',
    '正': '正',
    'ゴールド': 'ゴ',
    'ゴ': 'ゴ',
    'プラチナ': 'プ',
    'プ': 'プ',
    'ダイヤモンド': 'ダ',
    'ダ': 'ダ'
  };

  return mapping[normalized] || normalized;
}

/**
 * ブース出店を変換
 * 例: 「有り」→「○」、「あり」→「○」
 * @param {string} booth - ブース出店
 * @returns {string} 「○」または空文字
 */
function convertBooth_(booth) {
  if (!booth) return '';
  const normalized = String(booth).trim().toLowerCase();

  if (normalized === '有り' || normalized === 'あり' || normalized === '有' || normalized === 'yes') {
    return '○';
  }
  return '';
}

/**
 * フォーム回答シートから他会場名簿マスターに転記
 * 重複チェック: 同じEVENT_KEY + 氏名の組み合わせがあればスキップ
 * @returns {Object} { success, added, skipped, errors }
 */
function syncFormResponsesToOtherVenueMaster() {
  try {
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);

    // フォーム回答シート取得
    const shForm = ss.getSheetByName(FORM_RESPONSE_SHEET_NAME);
    if (!shForm) {
      return { success: false, message: `${FORM_RESPONSE_SHEET_NAME}が見つかりません` };
    }

    // 他会場名簿マスター取得
    const shMaster = ss.getSheetByName(OTHER_VENUE_MASTER_SHEET_NAME);
    if (!shMaster) {
      return { success: false, message: `${OTHER_VENUE_MASTER_SHEET_NAME}が見つかりません` };
    }

    // フォーム回答データ取得（ヘッダー除く）
    const formValues = shForm.getDataRange().getValues();
    if (formValues.length < 2) {
      return { success: true, message: 'フォーム回答なし', added: 0, skipped: 0 };
    }
    const formRows = formValues.slice(1);  // 1行目はヘッダー

    // 他会場名簿マスターの既存データ取得（重複チェック用）
    const masterValues = shMaster.getDataRange().getValues();
    const existingKeys = new Set();

    // ヘッダーは2行目（インデックス1）、データは3行目以降
    if (masterValues.length > 2) {
      for (let i = 2; i < masterValues.length; i++) {
        const eventKey = String(masterValues[i][1] || '').trim();  // B列: EVENT_KEY
        const name = String(masterValues[i][3] || '').trim();      // D列: 氏名
        if (eventKey && name) {
          existingKeys.add(`${eventKey}|${name}`);
        }
      }
    }

    // 追加するデータを準備
    const newRows = [];
    let skipped = 0;
    const errors = [];

    // フォーム回答列インデックス（0-based）
    const F_TIMESTAMP = 0;      // A列: タイムスタンプ
    const F_EVENT_DATE = 1;     // B列: 参加する例会
    const F_VENUE = 2;          // C列: 所属会場
    const F_BADGE = 3;          // D列: バッジ
    const F_ROLE = 4;           // E列: 役割
    const F_NAME = 5;           // F列: 氏名
    const F_FURIGANA = 6;       // G列: フリガナ
    const F_COMPANY = 7;        // H列: 会社名
    const F_POSITION = 8;       // I列: 役職
    const F_BUSINESS = 9;       // J列: 営業内容
    const F_REFERRER = 10;      // K列: 紹介者
    const F_BOOTH = 11;         // L列: ブース出店
    const F_GUEST_WITH = 12;    // M列: ゲスト同伴
    const F_GUEST_INFO = 13;    // N列: ゲスト情報

    formRows.forEach((row, idx) => {
      try {
        const name = String(row[F_NAME] || '').trim();
        if (!name) return;  // 氏名なしはスキップ

        const eventDateStr = String(row[F_EVENT_DATE] || '').trim();
        const eventKey = convertDateToEventKey_(eventDateStr);

        // 重複チェック
        const key = `${eventKey}|${name}`;
        if (existingKeys.has(key)) {
          skipped++;
          return;
        }

        // 他会場名簿マスター形式で行を作成
        // [会員ID, EVENT_KEY, バッジ, 氏名, フリガナ, 所属, 役割, 会社名, 役職, 紹介者, 営業内容, 賞, ブース, 区分]
        const newRow = [
          '',                                          // A: 会員ID（空）
          eventKey,                                    // B: EVENT_KEY
          convertBadge_(row[F_BADGE]),                 // C: バッジ
          name,                                        // D: 氏名
          String(row[F_FURIGANA] || '').trim(),        // E: フリガナ
          String(row[F_VENUE] || '').trim(),           // F: 所属
          String(row[F_ROLE] || '').trim(),            // G: 役割
          String(row[F_COMPANY] || '').trim(),         // H: 会社名
          String(row[F_POSITION] || '').trim(),        // I: 役職
          String(row[F_REFERRER] || '').trim(),        // J: 紹介者
          String(row[F_BUSINESS] || '').trim(),        // K: 営業内容
          '',                                          // L: 賞（空）
          convertBooth_(row[F_BOOTH]),                 // M: ブース
          '会員'                                       // N: 区分
        ];

        newRows.push(newRow);
        existingKeys.add(key);  // 同一バッチ内での重複防止

        // ゲスト同伴の処理（ゲスト情報があれば別行追加）
        const guestInfo = String(row[F_GUEST_INFO] || '').trim();
        if (guestInfo) {
          // ゲスト情報をパース（「氏名, 会社名」形式を想定）
          const guestParts = guestInfo.split(/[,、]/);
          const guestName = guestParts[0] ? guestParts[0].trim() : '';
          const guestCompany = guestParts[1] ? guestParts[1].trim() : '';

          if (guestName) {
            const guestKey = `${eventKey}|${guestName}`;
            if (!existingKeys.has(guestKey)) {
              const guestRow = [
                '',                                    // A: 会員ID（空）
                eventKey,                              // B: EVENT_KEY
                '',                                    // C: バッジ（空）
                guestName,                             // D: 氏名
                '',                                    // E: フリガナ（空）
                String(row[F_VENUE] || '').trim(),     // F: 所属（紹介者と同じ会場）
                '',                                    // G: 役割（空）
                guestCompany,                          // H: 会社名
                '',                                    // I: 役職（空）
                name,                                  // J: 紹介者（フォーム回答者）
                '',                                    // K: 営業内容（空）
                '',                                    // L: 賞（空）
                '',                                    // M: ブース（空）
                'ゲスト'                               // N: 区分
              ];
              newRows.push(guestRow);
              existingKeys.add(guestKey);
            }
          }
        }
      } catch (e) {
        errors.push(`行${idx + 2}: ${e.message}`);
      }
    });

    // データ追加
    if (newRows.length > 0) {
      const lastRow = shMaster.getLastRow();
      shMaster.getRange(lastRow + 1, 1, newRows.length, 14).setValues(newRows);
    }

    Logger.log(`syncFormResponsesToOtherVenueMaster: added=${newRows.length}, skipped=${skipped}`);

    return {
      success: true,
      added: newRows.length,
      skipped: skipped,
      errors: errors
    };
  } catch (e) {
    Logger.log('syncFormResponsesToOtherVenueMaster error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * フォーム回答同期（手動実行用 or ダッシュボードから呼び出し）
 */
function runFormSync() {
  const result = syncFormResponsesToOtherVenueMaster();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/** =========================================================
 *  チェックイン機能（専用シート版）
 *  - initCheckinData: 受付名簿から参加者をチェックインシートにコピー
 *  - doCheckin: 参加者のチェックイン実行（LIFF用）
 *  - getCheckinStatus: 全員のチェックイン状況取得（ダッシュボード用）
 *  - updateCheckinField: チェックイン関連フィールド更新（ダッシュボード用）
 *  - manualCheckin: 手動チェックイン（ダッシュボード用）
 * ======================================================= */

// チェックインシート名
const CHECKIN_SHEET_NAME = 'チェックイン';

// チェックインシートの列インデックス（0始まり）
const CHECKIN_COL = {
  EVENT_KEY: 0,      // A列: 例会キー
  NAME: 1,           // B列: 氏名
  CATEGORY: 2,       // C列: 区分
  AFFILIATION: 3,    // D列: 所属
  SEAT: 4,           // E列: 座席
  CHECKIN: 5,        // F列: チェックイン
  CHECKIN_TIME: 6,   // G列: チェックイン時刻
  BADGE_LENT: 7,     // H列: バッジ貸出
  BADGE_RETURNED: 8, // I列: バッジ返却
  RECEIPT_NO: 9,     // J列: 領収書番号
  CANCELLED: 10      // K列: キャンセル（欠席）
};

/**
 * 配席アーカイブから参加者データをチェックインシートに初期化
 * ★例会当日の朝にダッシュボードから実行
 * ※配席アーカイブの該当eventKeyのデータをコピー
 */
function initCheckinData() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const checkinSheet = ss.getSheetByName(CHECKIN_SHEET_NAME);

    if (!checkinSheet) {
      return { success: false, message: 'チェックインシートが見つかりません' };
    }

    // 例会当日イベントキーを取得
    const eventDayKey = getEventDayEventKey();
    if (!eventDayKey) {
      return { success: false, message: '例会当日イベントキーが設定されていません（設定シートI2）' };
    }

    // 既存データをクリア（ヘッダー行は残す）
    const lastRow = checkinSheet.getLastRow();
    if (lastRow > 1) {
      checkinSheet.getRange(2, 1, lastRow - 1, 11).clearContent();
    }

    // ========================================
    // 配席アーカイブから該当eventKeyのデータを取得
    // ========================================
    const archiveSheet = ss.getSheetByName(SEATING_ARCHIVE_SHEET_NAME);
    if (!archiveSheet) {
      return { success: false, message: '配席アーカイブシートが見つかりません' };
    }

    const archiveValues = archiveSheet.getDataRange().getValues();
    if (archiveValues.length < 2) {
      return { success: false, message: '配席アーカイブにデータがありません' };
    }

    // 配席アーカイブの列インデックス
    // A:例会キー, B:例会日, C:参加者ID, D:氏名, E:区分, F:所属, G:役割, H:チーム, I:テーブル, J:席番, K:確定日時
    const ARCHIVE_COL = {
      EVENT_KEY: 0,
      NAME: 3,
      CATEGORY: 4,
      AFFILIATION: 5,
      TABLE: 8
    };

    const newRows = [];

    for (let r = 1; r < archiveValues.length; r++) {
      const rowEventKey = String(archiveValues[r][ARCHIVE_COL.EVENT_KEY] || '').trim();
      if (rowEventKey !== eventDayKey) continue;

      const name = String(archiveValues[r][ARCHIVE_COL.NAME] || '').trim();
      if (!name) continue;

      const category = String(archiveValues[r][ARCHIVE_COL.CATEGORY] || '').trim() || '会員';
      const affiliation = String(archiveValues[r][ARCHIVE_COL.AFFILIATION] || '').trim();
      const table = String(archiveValues[r][ARCHIVE_COL.TABLE] || '').trim();

      newRows.push([
        eventDayKey,    // A: 例会キー
        name,           // B: 氏名
        category,       // C: 区分
        affiliation,    // D: 所属
        table,          // E: 座席（テーブル）
        '',             // F: チェックイン
        '',             // G: チェックイン時刻
        '',             // H: バッジ貸出
        '',             // I: バッジ返却
        '',             // J: 領収書番号
        ''              // K: キャンセル
      ]);
    }

    if (newRows.length === 0) {
      return { success: false, message: '配席アーカイブに該当する例会キー(' + eventDayKey + ')のデータがありません' };
    }

    // データを書き込み
    checkinSheet.getRange(2, 1, newRows.length, 11).setValues(newRows);

    Logger.log('チェックインデータ初期化完了: ' + newRows.length + '件 (eventKey: ' + eventDayKey + ')');
    return { success: true, count: newRows.length, eventKey: eventDayKey };

  } catch (e) {
    Logger.log('initCheckinData error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * チェックイン実行（参加者向けLIFF用）
 * @param {string} lineUserId - LINE UserID
 * @returns {Object} チェックイン結果
 */
function doCheckin(lineUserId) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    if (!lineUserId) {
      return { success: false, message: 'LINE UserIDが必要です' };
    }

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const memberSs = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const eventInfo = getCurrentEventInfo_();
    const eventKey = eventInfo.key;

    // 1. 会員名簿からLINE UserIDで会員を検索
    const memberSheet = memberSs.getSheetByName(MEMBER_SHEET_NAME);
    if (!memberSheet) {
      return { success: false, message: '会員名簿が見つかりません' };
    }

    const memberValues = memberSheet.getDataRange().getValues();
    let memberInfo = null;
    const idxUserId = 15;  // P列: LINE_userId
    const idxName = 2;     // C列: 氏名
    const idxCompany = 5;  // F列: 会社名

    for (let i = 2; i < memberValues.length; i++) {
      const userId = String(memberValues[i][idxUserId] || '').trim();
      if (userId === lineUserId) {
        memberInfo = {
          name: String(memberValues[i][idxName] || '').trim(),
          company: String(memberValues[i][idxCompany] || '').trim(),
          category: '会員'
        };
        break;
      }
    }

    if (!memberInfo) {
      return {
        success: false,
        message: '会員情報が見つかりません。受付係にお声がけください。',
        isUnknown: true
      };
    }

    // 2. チェックインシートで該当者を検索・更新
    const checkinSheet = ss.getSheetByName(CHECKIN_SHEET_NAME);
    if (!checkinSheet) {
      return { success: false, message: 'チェックインシートが見つかりません。initCheckinData()を実行してください。' };
    }

    const checkinValues = checkinSheet.getDataRange().getValues();
    let checkinSuccess = false;
    let seatInfo = '';

    for (let r = 1; r < checkinValues.length; r++) {
      const rowEventKey = String(checkinValues[r][CHECKIN_COL.EVENT_KEY] || '').trim();
      const rowName = String(checkinValues[r][CHECKIN_COL.NAME] || '').trim();

      if (rowEventKey === eventKey && rowName === memberInfo.name) {
        // 既にチェックイン済みかチェック
        const currentCheckin = String(checkinValues[r][CHECKIN_COL.CHECKIN] || '').trim();
        seatInfo = String(checkinValues[r][CHECKIN_COL.SEAT] || '').trim();

        if (currentCheckin === '○') {
          return {
            success: true,
            alreadyCheckedIn: true,
            message: '既にチェックイン済みです',
            participant: {
              name: memberInfo.name,
              company: memberInfo.company,
              category: memberInfo.category,
              seat: seatInfo
            }
          };
        }

        // チェックイン記録
        const now = new Date();
        checkinSheet.getRange(r + 1, CHECKIN_COL.CHECKIN + 1).setValue('○');
        checkinSheet.getRange(r + 1, CHECKIN_COL.CHECKIN_TIME + 1).setValue(now);
        checkinSuccess = true;
        break;
      }
    }

    if (checkinSuccess) {
      return {
        success: true,
        message: 'チェックイン完了',
        participant: {
          name: memberInfo.name,
          company: memberInfo.company,
          category: memberInfo.category,
          seat: seatInfo
        }
      };
    } else {
      return {
        success: false,
        message: '参加者名簿に名前が見つかりません。initCheckinData()を実行してください。',
        memberFound: true
      };
    }

  } catch (e) {
    Logger.log('doCheckin error: ' + e.message);
    return { success: false, message: 'エラーが発生しました: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * チェックインキャンセル（LIFF用）
 * @param {string} lineUserId - LINE UserID
 * @returns {Object} キャンセル結果
 */
function cancelCheckin(lineUserId) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    if (!lineUserId) {
      return { success: false, message: 'LINE UserIDが必要です' };
    }

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const memberSs = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const eventInfo = getCurrentEventInfo_();
    const eventKey = eventInfo.key;

    // 1. 会員名簿からLINE UserIDで会員を検索
    const memberSheet = memberSs.getSheetByName(MEMBER_SHEET_NAME);
    if (!memberSheet) {
      return { success: false, message: '会員名簿が見つかりません' };
    }

    const memberValues = memberSheet.getDataRange().getValues();
    let memberName = null;
    const idxUserId = 15;  // P列: LINE_userId
    const idxName = 2;     // C列: 氏名

    for (let i = 2; i < memberValues.length; i++) {
      const userId = String(memberValues[i][idxUserId] || '').trim();
      if (userId === lineUserId) {
        memberName = String(memberValues[i][idxName] || '').trim();
        break;
      }
    }

    if (!memberName) {
      return { success: false, message: '会員情報が見つかりません' };
    }

    // 2. チェックインシートで該当者を検索・更新
    const checkinSheet = ss.getSheetByName(CHECKIN_SHEET_NAME);
    if (!checkinSheet) {
      return { success: false, message: 'チェックインシートが見つかりません' };
    }

    const checkinValues = checkinSheet.getDataRange().getValues();

    for (let r = 1; r < checkinValues.length; r++) {
      const rowEventKey = String(checkinValues[r][CHECKIN_COL.EVENT_KEY] || '').trim();
      const rowName = String(checkinValues[r][CHECKIN_COL.NAME] || '').trim();

      if (rowEventKey === eventKey && rowName === memberName) {
        // チェックインをキャンセル（クリア）
        checkinSheet.getRange(r + 1, CHECKIN_COL.CHECKIN + 1).clearContent();
        checkinSheet.getRange(r + 1, CHECKIN_COL.CHECKIN_TIME + 1).clearContent();

        return {
          success: true,
          message: 'チェックインをキャンセルしました',
          name: memberName
        };
      }
    }

    return { success: false, message: '該当するチェックイン記録が見つかりません' };

  } catch (e) {
    Logger.log('cancelCheckin error: ' + e.message);
    return { success: false, message: 'エラーが発生しました: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 参加者情報取得（LIFF用）
 * @param {string} lineUserId - LINE UserID
 * @returns {Object} 参加者情報
 */
function getParticipantInfo(lineUserId) {
  if (!lineUserId) {
    return { success: false, message: 'LINE UserIDが必要です' };
  }

  const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const memberSs = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
  const eventInfo = getCurrentEventInfo_();
  const eventKey = eventInfo.key;

  // 会員名簿から検索
  const memberSheet = memberSs.getSheetByName(MEMBER_SHEET_NAME);
  if (!memberSheet) {
    return { success: false, message: '会員名簿が見つかりません' };
  }

  const memberValues = memberSheet.getDataRange().getValues();
  const idxUserId = 15;
  const idxName = 2;
  const idxCompany = 5;

  for (let i = 2; i < memberValues.length; i++) {
    const userId = String(memberValues[i][idxUserId] || '').trim();
    if (userId === lineUserId) {
      const name = String(memberValues[i][idxName] || '').trim();

      // チェックインシートから状態と座席を取得
      const checkinSheet = ss.getSheetByName(CHECKIN_SHEET_NAME);
      let checkedIn = false;
      let seat = '';

      if (checkinSheet) {
        const checkinValues = checkinSheet.getDataRange().getValues();
        for (let r = 1; r < checkinValues.length; r++) {
          const rowEventKey = String(checkinValues[r][CHECKIN_COL.EVENT_KEY] || '').trim();
          const rowName = String(checkinValues[r][CHECKIN_COL.NAME] || '').trim();

          if (rowEventKey === eventKey && rowName === name) {
            checkedIn = String(checkinValues[r][CHECKIN_COL.CHECKIN] || '').trim() === '○';
            seat = String(checkinValues[r][CHECKIN_COL.SEAT] || '').trim();
            break;
          }
        }
      }

      return {
        success: true,
        participant: {
          name: name,
          company: String(memberValues[i][idxCompany] || '').trim(),
          category: '会員',
          seat: seat,
          checkedIn: checkedIn
        }
      };
    }
  }

  return { success: false, message: '会員情報が見つかりません' };
}

/**
 * チェックイン状況を一括取得（ダッシュボード用）
 * @returns {Object} チェックイン状況
 */
function getCheckinStatus() {
  try {
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const eventInfo = getCurrentEventInfo_();
    const eventKey = eventInfo.key;
    const eventName = eventInfo.name;

    const result = {
      eventName: eventName,
      eventKey: eventKey,
      members: [],      // 福岡飯塚会員
      otherVenue: [],   // 他会場
      guests: [],       // ゲスト
      summary: {
        total: 0,
        checkedIn: 0,
        badgeLent: 0,
        badgeNotReturned: 0
      }
    };

    const checkinSheet = ss.getSheetByName(CHECKIN_SHEET_NAME);
    if (!checkinSheet) {
      return { success: false, message: 'チェックインシートが見つかりません。initCheckinData()を実行してください。', needsInit: true };
    }

    const values = checkinSheet.getDataRange().getValues();

    for (let r = 1; r < values.length; r++) {
      const rowEventKey = String(values[r][CHECKIN_COL.EVENT_KEY] || '').trim();
      if (rowEventKey !== eventKey) continue;

      const name = String(values[r][CHECKIN_COL.NAME] || '').trim();
      if (!name) continue;

      const category = String(values[r][CHECKIN_COL.CATEGORY] || '').trim();
      const checkedIn = String(values[r][CHECKIN_COL.CHECKIN] || '').trim() === '○';
      const badgeLent = String(values[r][CHECKIN_COL.BADGE_LENT] || '').trim() === '○';
      const badgeReturned = String(values[r][CHECKIN_COL.BADGE_RETURNED] || '').trim() === '○';
      const cancelled = String(values[r][CHECKIN_COL.CANCELLED] || '').trim() === '○';

      const person = {
        rowIndex: r + 1,  // シートの実際の行番号
        name: name,
        category: category,
        affiliation: String(values[r][CHECKIN_COL.AFFILIATION] || '').trim(),
        seat: String(values[r][CHECKIN_COL.SEAT] || '').trim(),
        checkedIn: checkedIn,
        checkinTime: values[r][CHECKIN_COL.CHECKIN_TIME] ? formatTime_(values[r][CHECKIN_COL.CHECKIN_TIME]) : '',
        badgeLent: category === 'ゲスト' ? false : badgeLent,
        badgeReturned: category === 'ゲスト' ? false : badgeReturned,
        receiptNo: String(values[r][CHECKIN_COL.RECEIPT_NO] || '').trim(),
        cancelled: cancelled
      };

      // 区分に応じて振り分け
      if (category === '会員') {
        result.members.push(person);
        if (badgeLent) result.summary.badgeLent++;
        if (badgeLent && !badgeReturned) result.summary.badgeNotReturned++;
      } else if (category === 'ゲスト') {
        result.guests.push(person);
      } else {
        result.otherVenue.push(person);
        if (badgeLent) result.summary.badgeLent++;
        if (badgeLent && !badgeReturned) result.summary.badgeNotReturned++;
      }

      result.summary.total++;
      if (checkedIn) result.summary.checkedIn++;
    }

    return { success: true, data: result };

  } catch (e) {
    Logger.log('getCheckinStatus error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * 時刻フォーマット
 */
function formatTime_(date) {
  if (!date) return '';
  if (date instanceof Date) {
    return Utilities.formatDate(date, 'Asia/Tokyo', 'HH:mm');
  }
  return String(date);
}

/**
 * チェックイン関連フィールドを更新（ダッシュボード用）
 * @param {number} rowIndex - シートの行番号
 * @param {string} field - 更新するフィールド
 * @param {any} value - 設定する値
 * @returns {Object} 結果
 */
function updateCheckinField(rowIndex, field, value) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sheet = ss.getSheetByName(CHECKIN_SHEET_NAME);

    if (!sheet) {
      return { success: false, message: 'チェックインシートが見つかりません' };
    }

    // フィールド名から列インデックスを取得
    let colIndex = -1;
    switch (field) {
      case 'checkin':
        colIndex = CHECKIN_COL.CHECKIN;
        break;
      case 'badgeLent':
        colIndex = CHECKIN_COL.BADGE_LENT;
        break;
      case 'badgeReturned':
        colIndex = CHECKIN_COL.BADGE_RETURNED;
        break;
      case 'receiptNo':
        colIndex = CHECKIN_COL.RECEIPT_NO;
        break;
      default:
        return { success: false, message: '不明なフィールド: ' + field };
    }

    // 値を設定
    let setValue = value;
    if (field === 'checkin' || field === 'badgeLent' || field === 'badgeReturned') {
      setValue = value ? '○' : '';
    }

    sheet.getRange(rowIndex, colIndex + 1).setValue(setValue);

    // チェックインの場合は時刻も記録
    if (field === 'checkin' && value) {
      sheet.getRange(rowIndex, CHECKIN_COL.CHECKIN_TIME + 1).setValue(new Date());
    }

    return { success: true };

  } catch (e) {
    Logger.log('updateCheckinField error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 手動チェックイン（ダッシュボード用）
 * @param {number} rowIndex - シートの行番号
 * @returns {Object} 結果
 */
function manualCheckin(rowIndex) {
  return updateCheckinField(rowIndex, 'checkin', true);
}

/**
 * キャンセル（欠席）マーク設定/解除（ダッシュボード用）
 * @param {number} rowIndex - シートの行番号
 * @param {boolean} cancelled - キャンセル状態
 * @returns {Object} 結果
 */
function markCheckinCancelled(rowIndex, cancelled) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sheet = ss.getSheetByName(CHECKIN_SHEET_NAME);

    if (!sheet) {
      return { success: false, message: 'チェックインシートが見つかりません' };
    }

    // K列（CANCELLED）に値を設定
    sheet.getRange(rowIndex, CHECKIN_COL.CANCELLED + 1).setValue(cancelled ? '○' : '');

    return { success: true };

  } catch (e) {
    Logger.log('markCheckinCancelled error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * チェックイン済みからのキャンセル（ダッシュボード用）
 * チェックインフラグと時刻をクリア
 * @param {number} rowIndex - シートの行番号
 * @returns {Object} 結果
 */
function cancelCheckinByRow(rowIndex) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sheet = ss.getSheetByName(CHECKIN_SHEET_NAME);

    if (!sheet) {
      return { success: false, message: 'チェックインシートが見つかりません' };
    }

    // チェックインとチェックイン時刻をクリア
    sheet.getRange(rowIndex, CHECKIN_COL.CHECKIN + 1).clearContent();
    sheet.getRange(rowIndex, CHECKIN_COL.CHECKIN_TIME + 1).clearContent();

    return { success: true };

  } catch (e) {
    Logger.log('cancelCheckinByRow error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/** =========================================================
 *  会員名簿編集機能
 * ======================================================= */

// 会員名簿マスターの列インデックス（0始まり）
const MEMBER_COL = {
  MEMBER_ID: 0,      // A: 会員ID
  BADGE: 1,          // B: バッジ
  NAME: 2,           // C: 氏名
  FURIGANA: 3,       // D: フリガナ
  AFFILIATION: 4,    // E: 所属
  COMPANY: 5,        // F: 会社名
  POSITION: 6,       // G: 役職
  ROLE: 7,           // H: 役割
  TEAM: 8,           // I: チーム
  INTRODUCER: 9,     // J: 紹介者
  BUSINESS: 10,      // K: 営業内容
  AWARD: 11,         // L: 賞
  RENEWAL_MONTH_OLD: 12, // M: 更新月（使用しない）
  REFERRAL_COUNT: 13,    // N: 紹介数
  DISPLAY_NAME: 14,      // O: displayName
  LINE_USER_ID: 15,      // P: LINE_userId
  CATEGORY: 16,          // Q: 区分
  AFFILIATION_OLD: 17,   // R: 所属（使用しない）
  REFERRAL_ACTIVE: 18,   // S: 自会場紹介在籍人数
  PURPLE_BADGE: 19,      // T: 紫バッジ
  JOIN_DATE: 20,         // U: 入会日
  RENEWAL_MONTH: 21,     // V: 更新月
  CONTINUED_MONTHS: 22,  // W: 継続月
  ROLE_X: 23,            // X: 役割分担
  ROLE_Y: 24,            // Y: 役割分担
  ROLE_Z: 25,            // Z: 役割分担
  RETIRED: 26,           // AA: 退会
  GROUP: 27,             // AB: グループ（連絡網）
  LEADER_FLAG: 28,       // AC: リーダーフラグ（〇）
  CONTACT_ORDER: 29      // AD: 連絡網順序（例: A1-2）
};

/**
 * 連絡網グループ情報を取得
 * 会員名簿マスターのAB列（グループ）、AC列（リーダーフラグ）、AD列（連絡網順序）を使用
 * AD列の形式: A1-2 = Aグループ、1行目、2番目
 * 退会者（AA列）は除外
 * @returns {Object} { success, groups: [{name, leader, rows: [[名前,...], [名前,...]]}, ...] }
 */
function getContactGroups() {
  try {
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sh = ss.getSheetByName(MEMBER_SHEET_NAME);

    if (!sh) {
      return { success: false, error: '会員名簿マスターが見つかりません', groups: [] };
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 3) {
      return { success: true, groups: [] };
    }

    // グループごとにメンバーを集計
    const groupMap = {};

    // 3行目以降がデータ（1,2行目はヘッダー）
    for (let i = 2; i < values.length; i++) {
      const row = values[i];
      const name = String(row[MEMBER_COL.NAME] || '').trim();
      const retired = String(row[MEMBER_COL.RETIRED] || '').trim();
      const group = String(row[MEMBER_COL.GROUP] || '').trim();
      const leaderFlag = String(row[MEMBER_COL.LEADER_FLAG] || '').trim();
      const contactOrder = String(row[MEMBER_COL.CONTACT_ORDER] || '').trim();

      // 退会者は除外
      if (retired) continue;
      // 名前がない行は除外
      if (!name) continue;
      // グループが未設定の場合は除外
      if (!group) continue;

      // グループ初期化
      if (!groupMap[group]) {
        groupMap[group] = {
          name: group,
          leader: null,
          membersWithOrder: []  // {name, row, pos} の配列
        };
      }

      // リーダーフラグが「〇」ならリーダー
      if (leaderFlag === '〇' || leaderFlag === '○') {
        groupMap[group].leader = name;
      } else {
        // 連絡網順序をパース（例: A1-2 → row=1, pos=2）
        let rowNum = 1;
        let posNum = 999;
        if (contactOrder) {
          const match = contactOrder.match(/^[A-J](\d+)-(\d+)$/);
          if (match) {
            rowNum = parseInt(match[1], 10);
            posNum = parseInt(match[2], 10);
          }
        }
        groupMap[group].membersWithOrder.push({ name, rowNum, posNum });
      }
    }

    // A〜Jの順にソートして配列化、メンバーを行ごとに整理
    const sortedGroups = Object.keys(groupMap)
      .sort()
      .map(groupName => {
        const g = groupMap[groupName];

        // メンバーを行番号・位置番号でソート
        g.membersWithOrder.sort((a, b) => {
          if (a.rowNum !== b.rowNum) return a.rowNum - b.rowNum;
          return a.posNum - b.posNum;
        });

        // 行ごとにグループ化
        const rowsMap = {};
        g.membersWithOrder.forEach(m => {
          if (!rowsMap[m.rowNum]) {
            rowsMap[m.rowNum] = [];
          }
          rowsMap[m.rowNum].push(m.name);
        });

        // 行番号順に配列化
        const rows = Object.keys(rowsMap)
          .sort((a, b) => parseInt(a) - parseInt(b))
          .map(rowNum => rowsMap[rowNum]);

        return {
          name: g.name,
          leader: g.leader,
          rows: rows
        };
      });

    return {
      success: true,
      groups: sortedGroups
    };

  } catch (e) {
    Logger.log('getContactGroups error: ' + e.message);
    return { success: false, error: e.message, groups: [] };
  }
}

/**
 * 会員編集用パスワードを検証
 * 設定シートのM2セルにパスワードを保存
 */
function verifyMemberEditPassword(password) {
  if (!password) {
    return { success: false, message: 'パスワードを入力してください' };
  }

  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sh) {
    return { success: false, message: '設定シートが見つかりません' };
  }

  // M2セルからパスワードを取得
  const correctPassword = String(sh.getRange('M2').getValue() || '').trim();
  if (!correctPassword) {
    return { success: false, message: '会員編集パスワードが設定されていません（設定シートM2）' };
  }

  if (password === correctPassword) {
    return { success: true };
  } else {
    return { success: false, message: 'パスワードが違います' };
  }
}

/**
 * 会員一覧を取得（フィルタ対応）
 * @param {Object} filter - { team, badge, renewalMonth, includeRetired, searchText }
 * @returns {Object} { success, members: [{rowIndex, memberId, name, ...}] }
 */
function getMembersForEdit(filter) {
  try {
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);

    // フィルタ条件
    const filterTeam = filter && filter.team ? filter.team : '';
    const filterBadge = filter && filter.badge ? filter.badge : '';
    const filterRenewalMonth = filter && filter.renewalMonth ? String(filter.renewalMonth) : '';
    const showRetiredOnly = filter && filter.showRetiredOnly === true;
    const searchText = filter && filter.searchText ? filter.searchText.toLowerCase() : '';

    const members = [];
    const teamSet = new Set();
    const badgeSet = new Set();

    // 退会者のみ表示モードの場合は退会者シートから取得
    if (showRetiredOnly) {
      const retiredSheet = ss.getSheetByName('退会者');
      if (retiredSheet) {
        const retiredValues = retiredSheet.getDataRange().getValues();
        // 退会者シートの列構成（0始まり）
        const RETIRED_COL = {
          MEMBER_ID: 0, BADGE: 1, NAME: 2, FURIGANA: 3, AFFILIATION: 4,
          COMPANY: 5, POSITION: 6, ROLE: 7, TEAM: 8, INTRODUCER: 9,
          BUSINESS: 10, AWARD: 11, JOIN_MONTH: 12, RETIRED_MONTH: 13,
          REFERRAL_COUNT: 14, DISPLAY_NAME: 15, LINE_USER_ID: 16
        };

        // 3行目以降がデータ（1,2行目はヘッダー）
        for (let i = 2; i < retiredValues.length; i++) {
          const row = retiredValues[i];
          const memberId = String(row[RETIRED_COL.MEMBER_ID] || '').trim();
          const name = String(row[RETIRED_COL.NAME] || '').trim();

          if (!name) continue;

          const badge = String(row[RETIRED_COL.BADGE] || '').trim();
          const furigana = String(row[RETIRED_COL.FURIGANA] || '').trim();
          const team = String(row[RETIRED_COL.TEAM] || '').trim();
          const company = String(row[RETIRED_COL.COMPANY] || '').trim();
          const affiliation = String(row[RETIRED_COL.AFFILIATION] || '').trim();
          const retiredMonth = String(row[RETIRED_COL.RETIRED_MONTH] || '').trim();

          if (team) teamSet.add(team);
          if (badge) badgeSet.add(badge);

          if (filterTeam && team !== filterTeam) continue;
          if (filterBadge && badge !== filterBadge) continue;
          if (searchText) {
            const searchTarget = (name + furigana + company).toLowerCase();
            if (searchTarget.indexOf(searchText) === -1) continue;
          }

          members.push({
            rowIndex: i + 1,
            memberId: memberId,
            badge: badge,
            name: name,
            furigana: furigana,
            affiliation: affiliation,
            company: company,
            position: String(row[RETIRED_COL.POSITION] || '').trim(),
            role: String(row[RETIRED_COL.ROLE] || '').trim(),
            team: team,
            introducer: String(row[RETIRED_COL.INTRODUCER] || '').trim(),
            business: String(row[RETIRED_COL.BUSINESS] || '').trim(),
            award: String(row[RETIRED_COL.AWARD] || '').trim(),
            referralCount: String(row[RETIRED_COL.REFERRAL_COUNT] || '').trim(),
            category: '',
            referralActive: '',
            purpleBadge: '',
            joinDate: String(row[RETIRED_COL.JOIN_MONTH] || '').trim(),
            renewalMonth: '',
            continuedMonths: '',
            retired: '退会',
            retiredMonth: retiredMonth,
            isFromRetiredSheet: true
          });
        }
      }

      return {
        success: true,
        members: members,
        teams: Array.from(teamSet).sort(),
        badges: Array.from(badgeSet)
      };
    }

    // 通常モード：会員名簿マスターから取得（退会者を除外）
    const sh = ss.getSheetByName(MEMBER_SHEET_NAME);
    if (!sh) {
      return { success: false, message: '会員名簿マスターが見つかりません', members: [] };
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 3) {
      return { success: true, members: [], teams: [], badges: [] };
    }

    // 3行目以降がデータ（1,2行目はヘッダー）
    for (let i = 2; i < values.length; i++) {
      const row = values[i];
      const memberId = String(row[MEMBER_COL.MEMBER_ID] || '').trim();
      const name = String(row[MEMBER_COL.NAME] || '').trim();

      if (!name) continue;

      const badge = String(row[MEMBER_COL.BADGE] || '').trim();
      const furigana = String(row[MEMBER_COL.FURIGANA] || '').trim();
      const team = String(row[MEMBER_COL.TEAM] || '').trim();
      const renewalMonth = String(row[MEMBER_COL.RENEWAL_MONTH] || '').trim();
      const retired = String(row[MEMBER_COL.RETIRED] || '').trim();
      const company = String(row[MEMBER_COL.COMPANY] || '').trim();
      const affiliation = String(row[MEMBER_COL.AFFILIATION] || '').trim();

      if (team) teamSet.add(team);
      if (badge) badgeSet.add(badge);

      // 退会者は除外
      if (retired) continue;

      if (filterTeam && team !== filterTeam) continue;
      if (filterBadge && badge !== filterBadge) continue;
      if (filterRenewalMonth && renewalMonth !== filterRenewalMonth) continue;
      if (searchText) {
        const searchTarget = (name + furigana + company).toLowerCase();
        if (searchTarget.indexOf(searchText) === -1) continue;
      }

      members.push({
        rowIndex: i + 1,
        memberId: memberId,
        badge: badge,
        name: name,
        furigana: furigana,
        affiliation: affiliation,
        company: company,
        position: String(row[MEMBER_COL.POSITION] || '').trim(),
        role: String(row[MEMBER_COL.ROLE] || '').trim(),
        team: team,
        introducer: String(row[MEMBER_COL.INTRODUCER] || '').trim(),
        business: String(row[MEMBER_COL.BUSINESS] || '').trim(),
        award: String(row[MEMBER_COL.AWARD] || '').trim(),
        referralCount: String(row[MEMBER_COL.REFERRAL_COUNT] || '').trim(),
        category: String(row[MEMBER_COL.CATEGORY] || '').trim(),
        referralActive: String(row[MEMBER_COL.REFERRAL_ACTIVE] || '').trim(),
        purpleBadge: String(row[MEMBER_COL.PURPLE_BADGE] || '').trim(),
        joinDate: formatDate_(row[MEMBER_COL.JOIN_DATE]),
        renewalMonth: renewalMonth,
        continuedMonths: String(row[MEMBER_COL.CONTINUED_MONTHS] || '').trim(),
        retired: retired,
        isFromRetiredSheet: false
      });
    }

    return {
      success: true,
      members: members,
      teams: Array.from(teamSet).sort(),
      badges: Array.from(badgeSet)
    };

  } catch (e) {
    Logger.log('getMembersForEdit error: ' + e.message);
    return { success: false, message: e.message, members: [] };
  }
}

/**
 * 会員情報を更新
 * @param {number} rowIndex - シートの行番号（1始まり）
 * @param {Object} data - 更新データ
 * @param {string} password - パスワード
 * @returns {Object} 結果
 */
function updateMember(rowIndex, data, password) {
  // パスワード検証
  const authResult = verifyMemberEditPassword(password);
  if (!authResult.success) {
    return authResult;
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sh = ss.getSheetByName(MEMBER_SHEET_NAME);

    if (!sh) {
      return { success: false, message: '会員名簿マスターが見つかりません' };
    }

    // 行番号の妥当性チェック
    if (rowIndex < 3) {
      return { success: false, message: '無効な行番号です' };
    }

    // 変更前の値を取得
    const oldRow = sh.getRange(rowIndex, 1, 1, 27).getValues()[0];
    const memberId = String(oldRow[MEMBER_COL.MEMBER_ID] || '');
    const memberName = String(oldRow[MEMBER_COL.NAME] || '');

    // 更新対象の列と値のマッピング（フィールド名も追加）
    const updates = [];

    if (data.badge !== undefined) updates.push({ col: MEMBER_COL.BADGE + 1, val: data.badge, field: 'badge', colIdx: MEMBER_COL.BADGE });
    if (data.name !== undefined) updates.push({ col: MEMBER_COL.NAME + 1, val: data.name, field: 'name', colIdx: MEMBER_COL.NAME });
    if (data.furigana !== undefined) updates.push({ col: MEMBER_COL.FURIGANA + 1, val: data.furigana, field: 'furigana', colIdx: MEMBER_COL.FURIGANA });
    if (data.affiliation !== undefined) updates.push({ col: MEMBER_COL.AFFILIATION + 1, val: data.affiliation, field: 'affiliation', colIdx: MEMBER_COL.AFFILIATION });
    if (data.company !== undefined) updates.push({ col: MEMBER_COL.COMPANY + 1, val: data.company, field: 'company', colIdx: MEMBER_COL.COMPANY });
    if (data.position !== undefined) updates.push({ col: MEMBER_COL.POSITION + 1, val: data.position, field: 'position', colIdx: MEMBER_COL.POSITION });
    if (data.role !== undefined) updates.push({ col: MEMBER_COL.ROLE + 1, val: data.role, field: 'role', colIdx: MEMBER_COL.ROLE });
    if (data.team !== undefined) updates.push({ col: MEMBER_COL.TEAM + 1, val: data.team, field: 'team', colIdx: MEMBER_COL.TEAM });
    if (data.introducer !== undefined) updates.push({ col: MEMBER_COL.INTRODUCER + 1, val: data.introducer, field: 'introducer', colIdx: MEMBER_COL.INTRODUCER });
    if (data.business !== undefined) updates.push({ col: MEMBER_COL.BUSINESS + 1, val: data.business, field: 'business', colIdx: MEMBER_COL.BUSINESS });
    if (data.award !== undefined) updates.push({ col: MEMBER_COL.AWARD + 1, val: data.award, field: 'award', colIdx: MEMBER_COL.AWARD });
    if (data.referralCount !== undefined) updates.push({ col: MEMBER_COL.REFERRAL_COUNT + 1, val: data.referralCount, field: 'referralCount', colIdx: MEMBER_COL.REFERRAL_COUNT });
    if (data.displayName !== undefined) updates.push({ col: MEMBER_COL.DISPLAY_NAME + 1, val: data.displayName, field: 'displayName', colIdx: MEMBER_COL.DISPLAY_NAME });
    if (data.lineUserId !== undefined) updates.push({ col: MEMBER_COL.LINE_USER_ID + 1, val: data.lineUserId, field: 'lineUserId', colIdx: MEMBER_COL.LINE_USER_ID });
    if (data.category !== undefined) updates.push({ col: MEMBER_COL.CATEGORY + 1, val: data.category, field: 'category', colIdx: MEMBER_COL.CATEGORY });
    if (data.referralActive !== undefined) updates.push({ col: MEMBER_COL.REFERRAL_ACTIVE + 1, val: data.referralActive, field: 'referralActive', colIdx: MEMBER_COL.REFERRAL_ACTIVE });
    if (data.purpleBadge !== undefined) updates.push({ col: MEMBER_COL.PURPLE_BADGE + 1, val: data.purpleBadge, field: 'purpleBadge', colIdx: MEMBER_COL.PURPLE_BADGE });
    if (data.joinDate !== undefined) updates.push({ col: MEMBER_COL.JOIN_DATE + 1, val: data.joinDate, field: 'joinDate', colIdx: MEMBER_COL.JOIN_DATE });
    if (data.renewalMonth !== undefined) updates.push({ col: MEMBER_COL.RENEWAL_MONTH + 1, val: data.renewalMonth, field: 'renewalMonth', colIdx: MEMBER_COL.RENEWAL_MONTH });
    if (data.continuedMonths !== undefined) updates.push({ col: MEMBER_COL.CONTINUED_MONTHS + 1, val: data.continuedMonths, field: 'continuedMonths', colIdx: MEMBER_COL.CONTINUED_MONTHS });
    if (data.retired !== undefined) updates.push({ col: MEMBER_COL.RETIRED + 1, val: data.retired ? '退会' : '', field: 'retired', colIdx: MEMBER_COL.RETIRED });

    // 変更があった項目を記録
    const changes = [];
    updates.forEach(u => {
      const oldValue = String(oldRow[u.colIdx] || '');
      const newValue = String(u.val || '');
      if (oldValue !== newValue) {
        changes.push({ field: u.field, oldValue: oldValue, newValue: newValue });
      }
    });

    // 各列を更新
    updates.forEach(u => {
      sh.getRange(rowIndex, u.col).setValue(u.val);
    });

    // 変更履歴を記録（変更があった場合のみ）
    // ★デバッグ: 一時的にコメントアウト
    // if (changes.length > 0) {
    //   logMemberChange_('編集', memberId, memberName, changes);
    // }

    return { success: true, message: '会員情報を更新しました' };

  } catch (e) {
    Logger.log('updateMember error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 新規会員を追加
 * @param {Object} data - 会員データ
 * @param {string} password - パスワード
 * @returns {Object} 結果
 */
function addMember(data, password) {
  // パスワード検証
  const authResult = verifyMemberEditPassword(password);
  if (!authResult.success) {
    return authResult;
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sh = ss.getSheetByName(MEMBER_SHEET_NAME);

    if (!sh) {
      return { success: false, message: '会員名簿マスターが見つかりません' };
    }

    // 必須項目チェック
    if (!data.name || !data.name.trim()) {
      return { success: false, message: '氏名は必須です' };
    }

    // 会員IDを自動採番（A列の最大値 + 1）
    const values = sh.getRange('A:A').getValues();
    let maxId = 0;
    for (let i = 2; i < values.length; i++) {
      const id = parseInt(values[i][0], 10);
      if (!isNaN(id) && id > maxId) {
        maxId = id;
      }
    }
    const newMemberId = maxId + 1;

    // 最終行に追加
    const lastRow = sh.getLastRow();
    const newRow = lastRow + 1;

    // 新しい行を構築
    const newRowData = new Array(27).fill('');
    newRowData[MEMBER_COL.MEMBER_ID] = newMemberId;
    newRowData[MEMBER_COL.BADGE] = data.badge || '';
    newRowData[MEMBER_COL.NAME] = data.name || '';
    newRowData[MEMBER_COL.FURIGANA] = data.furigana || '';
    newRowData[MEMBER_COL.AFFILIATION] = data.affiliation || '';
    newRowData[MEMBER_COL.COMPANY] = data.company || '';
    newRowData[MEMBER_COL.POSITION] = data.position || '';
    newRowData[MEMBER_COL.ROLE] = data.role || '';
    newRowData[MEMBER_COL.TEAM] = data.team || '';
    newRowData[MEMBER_COL.INTRODUCER] = data.introducer || '';
    newRowData[MEMBER_COL.BUSINESS] = data.business || '';
    newRowData[MEMBER_COL.AWARD] = data.award || '';
    newRowData[MEMBER_COL.REFERRAL_COUNT] = data.referralCount || '';
    newRowData[MEMBER_COL.DISPLAY_NAME] = data.displayName || '';
    newRowData[MEMBER_COL.LINE_USER_ID] = data.lineUserId || '';
    newRowData[MEMBER_COL.CATEGORY] = data.category || '';
    newRowData[MEMBER_COL.REFERRAL_ACTIVE] = data.referralActive || '';
    newRowData[MEMBER_COL.PURPLE_BADGE] = data.purpleBadge || '';
    newRowData[MEMBER_COL.JOIN_DATE] = data.joinDate || '';
    newRowData[MEMBER_COL.RENEWAL_MONTH] = data.renewalMonth || '';
    newRowData[MEMBER_COL.CONTINUED_MONTHS] = data.continuedMonths || '';
    newRowData[MEMBER_COL.RETIRED] = '';

    sh.getRange(newRow, 1, 1, newRowData.length).setValues([newRowData]);

    // 変更履歴を記録
    logMemberChange_('新規追加', String(newMemberId), data.name || '', null);

    return {
      success: true,
      message: '会員を追加しました',
      memberId: newMemberId,
      rowIndex: newRow
    };

  } catch (e) {
    Logger.log('addMember error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 会員を退会処理（論理削除）
 * @param {number} rowIndex - シートの行番号
 * @param {string} password - パスワード
 * @returns {Object} 結果
 */
function retireMember(rowIndex, password) {
  // パスワード検証
  const authResult = verifyMemberEditPassword(password);
  if (!authResult.success) {
    return authResult;
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const sh = ss.getSheetByName(MEMBER_SHEET_NAME);

    if (!sh) {
      return { success: false, message: '会員名簿マスターが見つかりません' };
    }

    // 会員情報を取得（履歴記録用）
    const row = sh.getRange(rowIndex, 1, 1, 27).getValues()[0];
    const memberId = String(row[MEMBER_COL.MEMBER_ID] || '');
    const memberName = String(row[MEMBER_COL.NAME] || '');

    // AA列に「退会」をセット
    sh.getRange(rowIndex, MEMBER_COL.RETIRED + 1).setValue('退会');

    // 変更履歴を記録
    logMemberChange_('退会', memberId, memberName, [{ field: 'retired', oldValue: '', newValue: '退会' }]);

    return { success: true, message: '退会処理が完了しました' };

  } catch (e) {
    Logger.log('retireMember error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 日付をフォーマット（内部ヘルパー）
 * Date型、YYYYMMDD形式の数値/文字列に対応
 * HTML input[type="date"]で使用するためYYYY-MM-DD形式（ハイフン区切り）で出力
 */
function formatDate_(date) {
  if (!date) return '';
  if (date instanceof Date) {
    return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  // YYYYMMDD形式（数値または文字列）をYYYY-MM-DD形式に変換
  const str = String(date);
  if (/^\d{8}$/.test(str)) {
    return str.slice(0, 4) + '-' + str.slice(4, 6) + '-' + str.slice(6, 8);
  }
  return str;
}

// 会員変更履歴シート名
const MEMBER_HISTORY_SHEET_NAME = '会員変更履歴';

// フィールド名の日本語マッピング
const MEMBER_FIELD_LABELS = {
  badge: 'バッジ',
  name: '氏名',
  furigana: 'フリガナ',
  affiliation: '所属',
  company: '会社名',
  position: '役職',
  role: '役割',
  team: 'チーム',
  introducer: '紹介者',
  business: '営業内容',
  award: '賞',
  referralCount: '紹介数',
  displayName: '表示名',
  lineUserId: 'LINE ID',
  category: '区分',
  referralActive: '自会場紹介在籍人数',
  purpleBadge: '紫バッジ',
  joinDate: '入会日',
  renewalMonth: '更新月',
  continuedMonths: '継続月',
  retired: '退会'
};

/**
 * 会員変更履歴を記録
 * @param {string} operationType - 操作種別（編集/新規追加/退会）
 * @param {string} memberId - 会員ID
 * @param {string} memberName - 会員名
 * @param {Array} changes - 変更内容の配列 [{field, oldValue, newValue}, ...]
 */
function logMemberChange_(operationType, memberId, memberName, changes) {
  try {
    const ss = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    let historySheet = ss.getSheetByName(MEMBER_HISTORY_SHEET_NAME);

    if (!historySheet) {
      // シートがなければ作成
      historySheet = ss.insertSheet(MEMBER_HISTORY_SHEET_NAME);
      historySheet.getRange('A1:G1').setValues([['変更日時', '操作種別', '会員ID', '会員名', '変更項目', '変更前', '変更後']]);
    }

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const rows = [];

    if (changes && changes.length > 0) {
      // 各変更項目ごとに1行記録
      changes.forEach(change => {
        const fieldLabel = MEMBER_FIELD_LABELS[change.field] || change.field;
        rows.push([timestamp, operationType, memberId, memberName, fieldLabel, change.oldValue || '', change.newValue || '']);
      });
    } else {
      // 変更項目がない場合（新規追加など）
      rows.push([timestamp, operationType, memberId, memberName, '', '', '']);
    }

    // 最終行の次に追加
    const lastRow = historySheet.getLastRow();
    historySheet.getRange(lastRow + 1, 1, rows.length, 7).setValues(rows);

  } catch (e) {
    Logger.log('logMemberChange_ error: ' + e.message);
    // 履歴記録のエラーは無視（メイン処理を止めない）
  }
}

/** =========================================================
 *  出欠リマインド送信機能
 * ======================================================= */

/**
 * 未回答者リスト取得（リマインド送信用）
 * 会員名簿マスターと出欠状況を突合して未回答者を抽出
 * @returns {Object} { success, eventInfo, members: [{name, userId, team, hasLineId}, ...], noLineIdMembers: [...] }
 */
function getUnrespondedMembersForReminder() {
  try {
    // 1. 設定シートから次回例会情報を取得
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!configSheet) {
      return { success: false, message: '設定シートが見つかりません' };
    }

    // F2: 次回イベントキー、G2: 次回開催日、J2: 回答期限
    const nextEventKey = String(configSheet.getRange('F2').getValue() || '').trim();
    const nextEventDateRaw = configSheet.getRange('G2').getValue();
    const deadlineRaw = configSheet.getRange('J2').getValue();

    if (!nextEventKey) {
      return { success: false, message: '次回イベントキーが設定されていません（設定シートF2）' };
    }

    const nextEventDate = nextEventDateRaw ? new Date(nextEventDateRaw) : null;

    // 回答期限: J2が有効な日付なら使用、そうでなければ開催月の前月末日
    let deadline = null;
    if (deadlineRaw && deadlineRaw instanceof Date && !isNaN(deadlineRaw.getTime())) {
      deadline = deadlineRaw;
    } else if (nextEventDate) {
      // new Date(year, month, 0) で前月末日を取得（2月例会→1/31）
      deadline = new Date(nextEventDate.getFullYear(), nextEventDate.getMonth(), 0);
    }

    // イベントキーを日本語形式に変換（2026年2月例会）
    const eventKeyJp = eventKeyToJapanese_(nextEventKey);

    // 日付をフォーマット
    const formatDate = (date) => {
      if (!date) return '';
      const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
      return Utilities.formatDate(date, 'Asia/Tokyo', 'M月d日') + '（' + weekdays[date.getDay()] + '）';
    };

    const eventInfo = {
      eventKey: nextEventKey,
      eventKeyJp: eventKeyJp,
      eventDate: formatDate(nextEventDate),
      deadline: formatDate(deadline)
    };

    // 2. 出欠状況シートから回答済みuserIdセットを作成
    const attendSs = SpreadsheetApp.openById(ATTEND_GUEST_SHEET_ID);
    const attendSheet = attendSs.getSheetByName(ATTEND_SHEET_NAME);

    if (!attendSheet) {
      return { success: false, message: '出欠状況シートが見つかりません' };
    }

    const attendData = attendSheet.getDataRange().getValues();
    const answeredUserIds = new Set();

    // ヘッダー行を確認（B列: eventKey, C列: userId）
    for (let i = 3; i < attendData.length; i++) {
      const rowEventKey = String(attendData[i][1] || '').trim(); // B列
      const rowUserId = String(attendData[i][2] || '').trim();   // C列

      // 対象イベントの回答のみ抽出
      if (rowEventKey === eventKeyJp && rowUserId) {
        answeredUserIds.add(rowUserId);
      }
    }

    // 3. 会員名簿マスターから未回答者を抽出
    const memberSheet = attendSs.getSheetByName(MEMBER_SHEET_NAME);
    if (!memberSheet) {
      return { success: false, message: '会員名簿マスターが見つかりません' };
    }

    const memberData = memberSheet.getDataRange().getValues();
    const membersWithLineId = [];
    const membersNoLineId = [];

    // 列インデックス（0始まり）
    const COL = {
      BADGE: 1,       // B列: バッジ
      NAME: 2,        // C列: 氏名
      TEAM: 8,        // I列: チーム
      LINE_USER_ID: 15, // P列: LINE_userId
      RETIRED: 26     // AA列: 退会
    };

    for (let i = 2; i < memberData.length; i++) {
      const row = memberData[i];
      const name = String(row[COL.NAME] || '').trim();
      const userId = String(row[COL.LINE_USER_ID] || '').trim();
      const team = String(row[COL.TEAM] || '').trim();
      const badge = String(row[COL.BADGE] || '').trim();
      const retired = String(row[COL.RETIRED] || '').trim();

      // 退会者はスキップ
      if (retired) continue;
      if (!name) continue;

      // 回答済みはスキップ
      if (userId && answeredUserIds.has(userId)) continue;

      const member = { name, team, badge };

      if (userId) {
        member.userId = userId;
        member.hasLineId = true;
        membersWithLineId.push(member);
      } else {
        member.hasLineId = false;
        membersNoLineId.push(member);
      }
    }

    return {
      success: true,
      eventInfo: eventInfo,
      members: membersWithLineId,
      noLineIdMembers: membersNoLineId,
      answeredCount: answeredUserIds.size
    };

  } catch (e) {
    Logger.log('getUnrespondedMembersForReminder error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * イベントキーを日本語形式に変換
 * @param {string} eventKey - イベントキー（例: 202602_01）
 * @returns {string} 日本語形式（例: 2026年2月例会）
 */
function eventKeyToJapanese_(eventKey) {
  if (!eventKey) return '';

  // 既に日本語形式なら変換しない
  if (eventKey.includes('年')) return eventKey;

  // 202602_01 → 2026年2月例会
  const match = eventKey.match(/^(\d{4})(\d{2})_\d{2}$/);
  if (match) {
    const year = match[1];
    const month = parseInt(match[2], 10);
    return year + '年' + month + '月例会';
  }

  return eventKey;
}

/**
 * リマインドメッセージを生成
 * 設定シートに保存されたテンプレートを使用（なければデフォルト）
 * プレースホルダー: {{eventDate}}, {{deadline}}
 * @param {string} eventDate - 例会日程（例: 2月13日（木））
 * @param {string} deadline - 回答期限（例: 2月10日（月））
 * @returns {string} メッセージテキスト
 */
function buildReminderMessage_(eventDate, deadline) {
  // テンプレートを取得
  const templateResult = getReminderMessageTemplate();
  let template = templateResult.success ? templateResult.template : getDefaultReminderTemplate_();

  // プレースホルダーを置換
  const message = template
    .replace(/\{\{eventDate\}\}/g, eventDate || '未設定')
    .replace(/\{\{deadline\}\}/g, deadline || '未設定');

  return message;
}

/**
 * LINEメッセージ送信（単一ユーザー）
 * @param {string} toUserId - 送信先のLINE userId
 * @param {string} messageText - メッセージテキスト
 * @returns {Object} { success, statusCode, response }
 */
function pushLineMessage_(toUserId, messageText) {
  if (!LINE_CHANNEL_ACCESS_TOKEN) {
    return { success: false, error: 'LINE_CHANNEL_ACCESS_TOKEN が設定されていません' };
  }
  if (!toUserId) {
    return { success: false, error: 'toUserId が指定されていません' };
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
      'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(url, params);
    const statusCode = res.getResponseCode();
    const responseText = res.getContentText();

    Logger.log('LINE push to ' + toUserId + ': status=' + statusCode);

    return {
      success: statusCode === 200,
      statusCode: statusCode,
      response: responseText
    };
  } catch (e) {
    Logger.log('LINE push error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * リマインド送信（ドライラン/テスト実行）
 * 実際には送信せず、対象者リストと件数を返す
 * @returns {Object} { success, eventInfo, targetCount, targets, message }
 */
function sendAttendanceReminderDryRun() {
  const result = getUnrespondedMembersForReminder();

  if (!result.success) {
    return result;
  }

  const targets = result.members.map(m => ({
    name: m.name,
    team: m.team,
    badge: m.badge
  }));

  const messagePreview = buildReminderMessage_(
    result.eventInfo.eventDate,
    result.eventInfo.deadline
  );

  return {
    success: true,
    eventInfo: result.eventInfo,
    targetCount: result.members.length,
    targets: targets,
    noLineIdCount: result.noLineIdMembers.length,
    noLineIdMembers: result.noLineIdMembers,
    messagePreview: messagePreview,
    message: `テスト実行完了: ${result.members.length}名が送信対象です（LINE未登録: ${result.noLineIdMembers.length}名）`
  };
}

/**
 * リマインド送信（本番）
 * @param {Array<string>} userIds - 送信対象のuserIdリスト（省略時は全未回答者）
 * @param {string} customMessage - カスタムメッセージ（省略時はデフォルトメッセージ）
 * @returns {Object} { success, sent, failed, errors, message }
 */
function sendAttendanceReminder(userIds, customMessage) {
  try {
    const result = getUnrespondedMembersForReminder();

    if (!result.success) {
      return result;
    }

    // 送信対象を決定
    let targets = result.members;
    if (userIds && userIds.length > 0) {
      // 指定されたuserIdのみ送信
      const targetSet = new Set(userIds);
      targets = result.members.filter(m => targetSet.has(m.userId));
    }

    if (targets.length === 0) {
      return { success: false, message: '送信対象者がいません' };
    }

    // メッセージ: カスタムメッセージがあればそれを使用、なければデフォルト
    const message = customMessage || buildReminderMessage_(
      result.eventInfo.eventDate,
      result.eventInfo.deadline
    );

    // 送信実行
    let sentCount = 0;
    let failedCount = 0;
    const errors = [];

    for (let i = 0; i < targets.length; i++) {
      const target = targets[i];

      // API制限対策: 100ms間隔
      if (i > 0) {
        Utilities.sleep(100);
      }

      const sendResult = pushLineMessage_(target.userId, message);

      if (sendResult.success) {
        sentCount++;
        Logger.log('送信成功: ' + target.name);
      } else {
        failedCount++;
        errors.push({ name: target.name, error: sendResult.error || sendResult.response });
        Logger.log('送信失敗: ' + target.name + ' - ' + (sendResult.error || sendResult.response));
      }
    }

    return {
      success: true,
      sent: sentCount,
      failed: failedCount,
      errors: errors,
      message: `送信完了: 成功 ${sentCount}件、失敗 ${failedCount}件`
    };

  } catch (e) {
    Logger.log('sendAttendanceReminder error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/** =========================================================
 *  出欠リマインド自動送信トリガー管理
 * ======================================================= */

/**
 * 指定年の月末日リストを取得
 * @param {number} year - 年（例: 2026）
 * @returns {Date[]} 月末日の配列（12個）
 */
function getLastDaysOfYear_(year) {
  const lastDays = [];
  for (let month = 1; month <= 12; month++) {
    // new Date(year, month, 0) で月末日を取得
    const lastDay = new Date(year, month, 0);
    lastDays.push(lastDay);
  }
  return lastDays;
}

/**
 * 自動送信用リマインド関数（トリガーから呼ばれる）
 * 未回答者全員に自動でリマインドを送信
 */
function autoSendAttendanceReminder() {
  try {
    Logger.log('autoSendAttendanceReminder: 自動送信開始');

    const result = getUnrespondedMembersForReminder();

    if (!result.success) {
      Logger.log('autoSendAttendanceReminder: データ取得失敗 - ' + result.message);
      return { success: false, message: result.message };
    }

    if (!result.members || result.members.length === 0) {
      Logger.log('autoSendAttendanceReminder: 未回答者なし');
      return { success: true, sent: 0, message: '未回答者なし' };
    }

    // デフォルトメッセージを生成
    const message = buildReminderMessage_(
      result.eventInfo.eventDate,
      result.eventInfo.deadline
    );

    // 送信実行
    let sentCount = 0;
    let failedCount = 0;
    const errors = [];

    for (let i = 0; i < result.members.length; i++) {
      const target = result.members[i];

      // API制限対策: 100ms間隔
      if (i > 0) {
        Utilities.sleep(100);
      }

      const sendResult = pushLineMessage_(target.userId, message);

      if (sendResult.success) {
        sentCount++;
        Logger.log('自動送信成功: ' + target.name);
      } else {
        failedCount++;
        errors.push({ name: target.name, error: sendResult.error || sendResult.response });
        Logger.log('自動送信失敗: ' + target.name + ' - ' + (sendResult.error || sendResult.response));
      }
    }

    const resultMessage = `自動送信完了: 成功 ${sentCount}件、失敗 ${failedCount}件`;
    Logger.log('autoSendAttendanceReminder: ' + resultMessage);

    return {
      success: true,
      sent: sentCount,
      failed: failedCount,
      errors: errors,
      message: resultMessage
    };

  } catch (e) {
    Logger.log('autoSendAttendanceReminder error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * 1年分のリマインドトリガーを一括作成
 * 毎月末日の12:00にautoSendAttendanceReminderを実行
 * @param {number} year - 対象年
 * @returns {Object} { success, created, skipped, message }
 */
function setupReminderTriggers(year) {
  try {
    // まず既存のリマインドトリガーを削除
    clearReminderTriggers();

    const lastDays = getLastDaysOfYear_(year);
    const now = new Date();
    let createdCount = 0;
    let skippedCount = 0;
    const createdDates = [];

    for (const lastDay of lastDays) {
      // 過去の日付はスキップ
      if (lastDay < now) {
        skippedCount++;
        continue;
      }

      // 12:00 JSTに設定
      const triggerDate = new Date(lastDay.getFullYear(), lastDay.getMonth(), lastDay.getDate(), 12, 0, 0);

      ScriptApp.newTrigger('autoSendAttendanceReminder')
        .timeBased()
        .at(triggerDate)
        .create();

      createdCount++;
      createdDates.push(Utilities.formatDate(triggerDate, 'Asia/Tokyo', 'M/d'));
    }

    const message = `${year}年のトリガーを${createdCount}件作成しました（スキップ: ${skippedCount}件）`;
    Logger.log('setupReminderTriggers: ' + message);

    return {
      success: true,
      created: createdCount,
      skipped: skippedCount,
      dates: createdDates,
      message: message
    };

  } catch (e) {
    Logger.log('setupReminderTriggers error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * リマインドトリガーを全削除
 * @returns {Object} { success, deleted, message }
 */
function clearReminderTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;

    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'autoSendAttendanceReminder') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }

    const message = `${deletedCount}件のトリガーを削除しました`;
    Logger.log('clearReminderTriggers: ' + message);

    return {
      success: true,
      deleted: deletedCount,
      message: message
    };

  } catch (e) {
    Logger.log('clearReminderTriggers error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * リマインドトリガーの状態を取得
 * @returns {Object} { success, count, nextTrigger, triggers }
 */
function getReminderTriggerStatus() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const reminderTriggers = [];
    let nextTrigger = null;
    const now = new Date();

    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'autoSendAttendanceReminder') {
        const triggerSource = trigger.getTriggerSource();

        // 時間ベースのトリガーの場合
        if (triggerSource === ScriptApp.TriggerSource.CLOCK) {
          // トリガーの詳細情報を取得（制限あり）
          const triggerId = trigger.getUniqueId();

          reminderTriggers.push({
            id: triggerId,
            type: 'time-based'
          });
        }
      }
    }

    // 次回送信日を計算（現在日以降の最初の月末日）
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();

    // 今月末日
    let nextLastDay = new Date(currentYear, currentMonth + 1, 0, 12, 0, 0);

    // 今月末日が過去なら来月末日
    if (nextLastDay <= now) {
      nextLastDay = new Date(currentYear, currentMonth + 2, 0, 12, 0, 0);
    }

    // トリガーが存在する場合のみ次回送信日を返す
    if (reminderTriggers.length > 0) {
      const weekdays = ['日', '月', '火', '水', '木', '金', '土'];

      // 次回送信日（月末日）に対応する例会情報を計算
      // 2/28送信 → 3月例会、回答期限は2/28
      const targetEventMonth = nextLastDay.getMonth() + 2; // 翌月（0始まりなので+2）
      const targetEventYear = targetEventMonth > 12
        ? nextLastDay.getFullYear() + 1
        : nextLastDay.getFullYear();
      const adjustedMonth = targetEventMonth > 12 ? targetEventMonth - 12 : targetEventMonth;

      // 回答期限は送信日（月末日）
      const deadlineFormatted = Utilities.formatDate(nextLastDay, 'Asia/Tokyo', 'M月d日') + '（' + weekdays[nextLastDay.getDay()] + '）';

      // 例会日程は設定シートのD列から取得
      // D列には各月の例会日程が入っている（D2から順に日付リスト）
      let eventDateFormatted = `${adjustedMonth}月例会`;
      try {
        const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
        const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
        if (configSheet) {
          // D列から日付を取得（D2:D13）
          const schedules = configSheet.getRange('D2:D13').getValues();
          for (let i = 0; i < schedules.length; i++) {
            const dateVal = schedules[i][0];
            if (dateVal && dateVal instanceof Date) {
              const month = dateVal.getMonth() + 1; // 1-12
              if (month === adjustedMonth) {
                eventDateFormatted = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'M月d日') + '（' + weekdays[dateVal.getDay()] + '）';
                break;
              }
            }
          }
        }
      } catch (e) {
        Logger.log('getReminderTriggerStatus: 設定シート読み込みエラー - ' + e.message);
      }

      nextTrigger = {
        date: Utilities.formatDate(nextLastDay, 'Asia/Tokyo', 'M月d日'),
        weekday: weekdays[nextLastDay.getDay()],
        time: '12:00',
        formatted: Utilities.formatDate(nextLastDay, 'Asia/Tokyo', 'M月d日') + '（' + weekdays[nextLastDay.getDay()] + '）12:00',
        // 配信メッセージ用の情報
        eventDate: eventDateFormatted,
        deadline: deadlineFormatted
      };
    }

    // 未回答者数も取得
    let unrespondedCount = 0;
    let noLineIdCount = 0;
    let noLineIdNames = [];
    try {
      const reminderData = getUnrespondedMembersForReminder();
      if (reminderData.success) {
        unrespondedCount = reminderData.members ? reminderData.members.length : 0;
        noLineIdCount = reminderData.noLineIdMembers ? reminderData.noLineIdMembers.length : 0;
        noLineIdNames = reminderData.noLineIdMembers ? reminderData.noLineIdMembers.map(m => m.name) : [];
      }
    } catch (e) {
      Logger.log('getReminderTriggerStatus: 未回答者取得エラー - ' + e.message);
    }

    return {
      success: true,
      count: reminderTriggers.length,
      nextTrigger: nextTrigger,
      triggers: reminderTriggers,
      unrespondedCount: unrespondedCount,
      noLineIdCount: noLineIdCount,
      noLineIdNames: noLineIdNames
    };

  } catch (e) {
    Logger.log('getReminderTriggerStatus error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/** =========================================================
 *  リマインドメッセージテンプレート管理
 *  「出欠リマインド管理」シートを使用
 * ======================================================= */

const REMINDER_SHEET_NAME = '出欠リマインド管理';

/**
 * 出欠リマインド管理シートを取得（なければ作成）
 */
function getReminderSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  let sheet = ss.getSheetByName(REMINDER_SHEET_NAME);

  if (!sheet) {
    // シートを作成
    sheet = ss.insertSheet(REMINDER_SHEET_NAME);
    // ヘッダー設定
    sheet.getRange('A1').setValue('メッセージテンプレート');
    sheet.getRange('A1').setFontWeight('bold');
    sheet.setColumnWidth(1, 600);
    Logger.log('getReminderSheet_: シートを作成しました');
  }

  return sheet;
}

/**
 * デフォルトのリマインドメッセージテンプレートを取得
 * プレースホルダー: {{eventDate}}, {{deadline}}
 */
function getDefaultReminderTemplate_() {
  return `【出欠回答のお願い】

いつもお世話になっております。
守成クラブ福岡飯塚 事務局です。

※本メッセージと行き違いで
　すでにご回答済みの場合は
　何卒ご容赦ください。

次回例会の出欠につきまして、
まだご回答をいただいておりません。

━━━━━━━━━━━━━━━━━━
■ 例会日程
　{{eventDate}}

■ 回答期限
　{{deadline}}

■ ご注意
　期限を過ぎますと
　自動的に「出席」扱いとなります。

　その後のキャンセルには
　キャンセル料 4,000円が発生いたします。
━━━━━━━━━━━━━━━━━━

画面下部のリッチメニューより
「例会出欠」ボタンを押して
ご回答をお願いいたします。

よろしくお願いいたします。`;
}

/**
 * リマインドメッセージテンプレートを取得
 * @returns {Object} { success, template, isDefault }
 */
function getReminderMessageTemplate() {
  try {
    const sheet = getReminderSheet_();
    const savedTemplate = String(sheet.getRange('A2').getValue() || '').trim();

    if (savedTemplate) {
      return {
        success: true,
        template: savedTemplate,
        isDefault: false
      };
    } else {
      return {
        success: true,
        template: getDefaultReminderTemplate_(),
        isDefault: true
      };
    }
  } catch (e) {
    Logger.log('getReminderMessageTemplate error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * リマインドメッセージテンプレートを保存
 * @param {string} template - メッセージテンプレート
 * @returns {Object} { success, message }
 */
function saveReminderMessageTemplate(template) {
  try {
    const sheet = getReminderSheet_();
    sheet.getRange('A2').setValue(template);

    Logger.log('saveReminderMessageTemplate: テンプレートを保存しました');
    return { success: true, message: 'メッセージを保存しました' };
  } catch (e) {
    Logger.log('saveReminderMessageTemplate error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * リマインドメッセージテンプレートをデフォルトにリセット
 * @returns {Object} { success, template, message }
 */
function resetReminderMessageTemplate() {
  try {
    const sheet = getReminderSheet_();
    // セルをクリア（デフォルトを使用する状態に）
    sheet.getRange('A2').clearContent();

    const defaultTemplate = getDefaultReminderTemplate_();
    Logger.log('resetReminderMessageTemplate: デフォルトにリセットしました');

    return {
      success: true,
      template: defaultTemplate,
      message: 'デフォルトメッセージにリセットしました'
    };
  } catch (e) {
    Logger.log('resetReminderMessageTemplate error: ' + e.message);
    return { success: false, message: e.message };
  }
}
