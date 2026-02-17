// ===============================
// 例会アンケート機能
// ===============================

// アンケート設定
const SURVEY_SHEET_NAME = 'アンケート回答';
const SURVEY_URL = 'https://shusei-survey.pages.dev/';

// ===============================
// ★ アンケート回答保存API
// ===============================

/**
 * アンケート回答を保存
 * doPostから呼び出される
 */
function handleSurveyPost(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SURVEY_SHEET_NAME);

    // シートがなければ作成
    if (!sheet) {
      sheet = createSurveySheet_(ss);
    }

    // データを抽出
    const meetingKey = String(data.meetingKey || '').trim();

    // Q1: 例会満足度
    const meetingSatisfaction = parseInt(data.meetingSatisfaction, 10) || 0;
    const meetingGoodPoints = String(data.meetingGoodPoints || '').trim();
    const meetingImprovements = String(data.meetingImprovements || '').trim();

    // Q2: システム（配列はカンマ区切りに変換）
    const attendanceSystem = Array.isArray(data.attendanceSystem) ? data.attendanceSystem.join(', ') : '';
    const attendanceComment = String(data.attendanceComment || '').trim();
    const salesSystem = Array.isArray(data.salesSystem) ? data.salesSystem.join(', ') : '';
    const salesComment = String(data.salesComment || '').trim();
    const otherSystemComment = String(data.otherSystemComment || '').trim();

    // Q3: 福岡飯塚満足度
    const clubSatisfaction = parseInt(data.clubSatisfaction, 10) || 0;
    const clubGoodPoints = String(data.clubGoodPoints || '').trim();
    const clubImprovements = String(data.clubImprovements || '').trim();

    // Q4: その他
    const otherComments = String(data.otherComments || '').trim();

    // バリデーション
    if (!meetingSatisfaction || !clubSatisfaction) {
      return _out({ success: false, error: '必須項目が入力されていません' });
    }

    // 行を追加
    const now = new Date();
    const rowData = [
      now,                    // A: タイムスタンプ
      meetingKey,             // B: 例会キー
      meetingSatisfaction,    // C: 例会満足度
      meetingGoodPoints,      // D: 例会_良かった点
      meetingImprovements,    // E: 例会_改善点
      attendanceSystem,       // F: 出欠システム評価
      attendanceComment,      // G: 出欠コメント
      salesSystem,            // H: 売上システム評価
      salesComment,           // I: 売上コメント
      otherSystemComment,     // J: その他システムコメント
      clubSatisfaction,       // K: 福岡飯塚満足度
      clubGoodPoints,         // L: 福岡飯塚_良かった点
      clubImprovements,       // M: 福岡飯塚_改善点
      otherComments           // N: その他コメント
    ];

    sheet.appendRow(rowData);

    return _out({
      success: true,
      message: '回答を保存しました'
    });

  } catch (error) {
    Logger.log('Survey save error: ' + error);
    return _out({
      success: false,
      error: 'サーバーエラー: ' + error.message
    });
  }
}

/**
 * アンケート回答シートを作成
 */
function createSurveySheet_(ss) {
  const sheet = ss.insertSheet(SURVEY_SHEET_NAME);

  // ヘッダー行を設定
  const headers = [
    'タイムスタンプ',        // A
    '例会キー',              // B
    '例会満足度',            // C
    '例会_良かった点',       // D
    '例会_改善点',           // E
    '出欠システム',          // F
    '出欠コメント',          // G
    '売上システム',          // H
    '売上コメント',          // I
    'その他システム',        // J
    '福岡飯塚満足度',        // K
    '福岡飯塚_良かった点',   // L
    '福岡飯塚_改善点',       // M
    'その他コメント'         // N
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a5c38');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅を調整
  sheet.setColumnWidth(1, 150);  // タイムスタンプ
  sheet.setColumnWidth(2, 120);  // 例会キー
  sheet.setColumnWidth(3, 80);   // 例会満足度
  sheet.setColumnWidth(4, 200);  // 例会_良かった点
  sheet.setColumnWidth(5, 200);  // 例会_改善点
  sheet.setColumnWidth(6, 180);  // 出欠システム
  sheet.setColumnWidth(7, 200);  // 出欠コメント
  sheet.setColumnWidth(8, 180);  // 売上システム
  sheet.setColumnWidth(9, 200);  // 売上コメント
  sheet.setColumnWidth(10, 200); // その他システム
  sheet.setColumnWidth(11, 80);  // 福岡飯塚満足度
  sheet.setColumnWidth(12, 200); // 福岡飯塚_良かった点
  sheet.setColumnWidth(13, 200); // 福岡飯塚_改善点
  sheet.setColumnWidth(14, 250); // その他コメント

  // 1行目を固定
  sheet.setFrozenRows(1);

  Logger.log('アンケート回答シートを作成しました');
  return sheet;
}


// ===============================
// ★ LINE送信機能
// ===============================

/**
 * 当月例会参加者（自会場のみ）にアンケートリンクを送信
 * @param {string} eventKey - 例会キー（例: 2026年2月例会）
 * @param {boolean} dryRun - テスト実行（実際には送信しない）
 * @returns {Object} 送信結果
 */
function sendSurveyToParticipants(eventKey, dryRun) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const rosterSheet = ss.getSheetByName(ROSTER_SHEET_NAME);

  if (!rosterSheet) {
    throw new Error('参加者名簿シートが見つかりません: ' + ROSTER_SHEET_NAME);
  }

  // 参加者名簿からLINE_IDと出欠を取得
  const data = rosterSheet.getDataRange().getValues();

  // ヘッダー行を探す
  let headerRow = -1;
  let lineIdCol = -1;
  let attendCol = -1;
  let nameCol = -1;

  for (let r = 0; r < Math.min(10, data.length); r++) {
    const cols = data[r].map(v => String(v || '').trim());
    const li = cols.indexOf('LINE_ID');
    if (li >= 0) {
      headerRow = r;
      lineIdCol = li;
      attendCol = cols.indexOf('出欠');
      nameCol = cols.indexOf('氏名');
      break;
    }
  }

  if (headerRow < 0 || lineIdCol < 0) {
    throw new Error('LINE_IDの見出しが見つかりません');
  }

  // 出席者のLINE_IDを抽出
  const participants = [];
  for (let r = headerRow + 1; r < data.length; r++) {
    const row = data[r];
    const lineId = String(row[lineIdCol] || '').trim();
    const attend = String(attendCol >= 0 ? (row[attendCol] || '') : '').trim();
    const name = String(nameCol >= 0 ? (row[nameCol] || '') : '').trim();

    // 出席者（○）かつLINE_IDがある人
    if (lineId && attend === '○') {
      participants.push({
        lineId: lineId,
        name: name
      });
    }
  }

  if (participants.length === 0) {
    return {
      success: true,
      message: '送信対象者がいません',
      count: 0,
      dryRun: dryRun
    };
  }

  // アンケートURLを生成（例会キーをパラメータに含める）
  const encodedKey = encodeURIComponent(eventKey);
  const surveyUrl = SURVEY_URL + '?key=' + encodedKey;

  // LINEメッセージ作成
  const message =
    '【例会アンケートのお願い】\n\n' +
    '本日は例会にご参加いただき、ありがとうございました。\n\n' +
    '今後のより良い運営のため、簡単なアンケートにご協力をお願いいたします。\n' +
    '（所要時間：約1分）\n\n' +
    '▼ アンケートはこちら\n' +
    surveyUrl + '\n\n' +
    '※ 回答は匿名で収集されます\n\n' +
    '守成クラブ福岡飯塚';

  // 送信処理
  const results = {
    success: [],
    failed: []
  };

  for (const p of participants) {
    if (dryRun) {
      // テスト実行：送信せずに記録のみ
      results.success.push({ name: p.name, lineId: p.lineId });
    } else {
      // 本番送信
      try {
        pushLineMessage(p.lineId, message);
        results.success.push({ name: p.name, lineId: p.lineId });
        // API制限回避のため少し待機
        Utilities.sleep(100);
      } catch (error) {
        Logger.log('LINE送信エラー: ' + p.name + ' - ' + error);
        results.failed.push({ name: p.name, lineId: p.lineId, error: error.message });
      }
    }
  }

  return {
    success: true,
    message: dryRun ? 'テスト実行完了' : 'アンケート送信完了',
    count: results.success.length,
    failed: results.failed.length,
    dryRun: dryRun,
    details: results
  };
}


// ===============================
// ★ 自動送信トリガー（開催日ベース）
// ===============================

/**
 * 例会開催日21時にアンケートを自動送信
 * トリガーから呼び出される
 */
function autoSendSurvey() {
  try {
    // 設定シートからイベント情報を取得
    const settings = getEventSettings_();
    const eventKey = settings.currentEventKey;

    // イベントキーを日本語形式に変換
    const japaneseEventKey = eventKeyToJapanese_(eventKey);

    // アンケート送信実行
    const result = sendSurveyToParticipants(japaneseEventKey, false);

    Logger.log('autoSendSurvey 完了: ' + JSON.stringify(result));

    // 管理者に結果を通知
    const notifyMessage =
      '【アンケート自動送信完了】\n\n' +
      '例会: ' + japaneseEventKey + '\n' +
      '送信成功: ' + result.count + '名\n' +
      '送信失敗: ' + result.failed + '名';

    pushLineMessage(OWNER_USER_ID, notifyMessage);

    // 次回例会のトリガーを設定
    setupNextSurveyTrigger_();

  } catch (error) {
    Logger.log('autoSendSurvey エラー: ' + error);

    // エラー通知
    pushLineMessage(OWNER_USER_ID,
      '【アンケート自動送信エラー】\n' + error.message);
  }
}


/**
 * アンケート送信トリガーを設定（開催日ベース）
 * 設定シートP2（現在開催日）の21時にトリガーを設定
 * @returns {Object} { success, message, triggerDate }
 */
function setupSurveyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoSendSurvey') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  }
  Logger.log('setupSurveyTrigger: 既存トリガー削除数: ' + deletedCount);

  // 設定シートから開催日を取得（D2セル）
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!sh) {
    Logger.log('setupSurveyTrigger: 設定シートが見つかりません');
    return { success: false, message: '設定シートが見つかりません' };
  }

  const eventDate = sh.getRange('D2').getValue();
  if (!eventDate) {
    Logger.log('setupSurveyTrigger: 現在開催日（D2）が空です');
    return { success: false, message: '現在開催日（D2）が空です' };
  }

  // 開催日の21時にトリガーを設定
  const d = new Date(eventDate);
  const triggerTime = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 21, 0, 0);

  // 過去の日付ならスキップ
  if (triggerTime <= new Date()) {
    Logger.log('setupSurveyTrigger: 過去の日付のためスキップ: ' + triggerTime);
    return { success: false, message: '開催日が過去です: ' + triggerTime };
  }

  // トリガー作成
  ScriptApp.newTrigger('autoSendSurvey')
    .timeBased()
    .at(triggerTime)
    .create();

  Logger.log('setupSurveyTrigger: トリガー設定完了');
  Logger.log('  開催日: ' + d);
  Logger.log('  送信時刻: ' + triggerTime);

  return {
    success: true,
    message: 'アンケート送信トリガーを設定しました',
    triggerDate: triggerTime.toISOString()
  };
}


/**
 * 次回例会のアンケートトリガーを設定（内部用）
 * autoSendSurvey完了後に呼び出される
 */
function setupNextSurveyTrigger_() {
  // 設定シートから次月開催日を取得（G2セル）
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!sh) {
    Logger.log('setupNextSurveyTrigger_: 設定シートが見つかりません');
    return;
  }

  const nextEventDate = sh.getRange('G2').getValue();
  if (!nextEventDate) {
    Logger.log('setupNextSurveyTrigger_: 次月開催日（G2）が空です');
    return;
  }

  // 次月開催日の21時にトリガーを設定
  const d = new Date(nextEventDate);
  const triggerTime = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 21, 0, 0);

  // 過去の日付ならスキップ
  if (triggerTime <= new Date()) {
    Logger.log('setupNextSurveyTrigger_: 過去の日付のためスキップ: ' + triggerTime);
    return;
  }

  // トリガー作成
  ScriptApp.newTrigger('autoSendSurvey')
    .timeBased()
    .at(triggerTime)
    .create();

  Logger.log('setupNextSurveyTrigger_: 次回トリガー設定完了: ' + triggerTime);
}


/**
 * トリガーを削除
 */
function removeSurveyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoSendSurvey') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }

  Logger.log('アンケート自動送信トリガーを削除しました: ' + removed + '件');
  return { success: true, removed: removed };
}


/**
 * 現在設定されているトリガーを確認
 */
function getSurveyTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const surveyTriggers = [];

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoSendSurvey') {
      surveyTriggers.push({
        id: trigger.getUniqueId(),
        nextRun: trigger.getTriggerSource()
      });
    }
  }

  return {
    count: surveyTriggers.length,
    triggers: surveyTriggers
  };
}


// ===============================
// ★ 手動送信（テスト用）
// ===============================

/**
 * 手動でアンケート送信（ドライラン）
 */
function testSendSurveyDryRun() {
  const eventKey = getAttendanceEventKey_();
  const result = sendSurveyToParticipants(eventKey, true);
  Logger.log('テスト実行結果: ' + JSON.stringify(result, null, 2));
  return result;
}

/**
 * 手動でアンケート送信（本番）
 */
function manualSendSurvey() {
  const eventKey = getAttendanceEventKey_();
  const result = sendSurveyToParticipants(eventKey, false);
  Logger.log('送信結果: ' + JSON.stringify(result, null, 2));
  return result;
}


// ===============================
// ★ アンケート結果取得API
// ===============================

/**
 * 指定した例会のアンケート結果を取得
 * @param {string} eventKey - 例会キー
 * @returns {Object} アンケート結果サマリー
 */
function getSurveyResults(eventKey) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SURVEY_SHEET_NAME);

  if (!sheet) {
    return {
      success: false,
      error: 'アンケート回答シートが見つかりません'
    };
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return {
      success: true,
      eventKey: eventKey,
      count: 0,
      averages: {},
      comments: []
    };
  }

  // ヘッダー行をスキップしてフィルタ
  const responses = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowEventKey = String(row[1] || '').trim();

    if (!eventKey || rowEventKey === eventKey) {
      responses.push({
        timestamp: row[0],
        eventKey: rowEventKey,
        systemRating: Number(row[2]) || 0,
        systemComment: String(row[3] || '').trim(),
        meetingSatisfaction: Number(row[4]) || 0,
        clubSatisfaction: Number(row[5]) || 0,
        improvementRequest: String(row[6] || '').trim()
      });
    }
  }

  if (responses.length === 0) {
    return {
      success: true,
      eventKey: eventKey,
      count: 0,
      averages: {},
      comments: []
    };
  }

  // 平均を計算
  const sum = {
    systemRating: 0,
    meetingSatisfaction: 0,
    clubSatisfaction: 0
  };

  for (const r of responses) {
    sum.systemRating += r.systemRating;
    sum.meetingSatisfaction += r.meetingSatisfaction;
    sum.clubSatisfaction += r.clubSatisfaction;
  }

  const count = responses.length;
  const averages = {
    systemRating: Math.round((sum.systemRating / count) * 10) / 10,
    meetingSatisfaction: Math.round((sum.meetingSatisfaction / count) * 10) / 10,
    clubSatisfaction: Math.round((sum.clubSatisfaction / count) * 10) / 10
  };

  // コメントを抽出（空でないもの）
  const comments = [];
  for (const r of responses) {
    if (r.systemComment) {
      comments.push({ type: 'system', text: r.systemComment });
    }
    if (r.improvementRequest) {
      comments.push({ type: 'improvement', text: r.improvementRequest });
    }
  }

  return {
    success: true,
    eventKey: eventKey,
    count: count,
    averages: averages,
    comments: comments
  };
}
