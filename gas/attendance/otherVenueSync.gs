// ===============================
// 他会場フォーム回答の自動同期
// ===============================

// シート名定数
const FORM_RESPONSE_SHEET = 'フォームの回答 11';
const OTHER_VENUE_MASTER_SHEET = '他会場名簿マスター';

/**
 * フォーム送信時に自動実行されるトリガー関数
 * 「他会場参加申込フォーム」の回答を「他会場名簿マスター」に転記
 * @param {Object} e - フォーム送信イベントオブジェクト
 */
function onOtherVenueFormSubmit(e) {
  try {
    // フォーム回答の値を取得
    const response = e.namedValues;
    if (!response) {
      Logger.log('フォーム回答が取得できませんでした');
      return;
    }

    // 必須フィールドのチェック
    const name = getFirstValue_(response['氏名']);
    if (!name) {
      Logger.log('氏名がありません - スキップ');
      return;
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const shMaster = ss.getSheetByName(OTHER_VENUE_MASTER_SHEET);

    if (!shMaster) {
      Logger.log('他会場名簿マスターが見つかりません');
      return;
    }

    // EVENT_KEY生成（日付から変換）
    const eventDateStr = getFirstValue_(response['参加する例会']);
    const eventKey = convertDateToEventKey_(eventDateStr);

    // 重複チェック（同じEVENT_KEY + 氏名の組み合わせ）
    if (isDuplicate_(shMaster, eventKey, name)) {
      Logger.log(`重複のためスキップ: ${eventKey} - ${name}`);
      return;
    }

    // データ行を作成
    const newRow = [
      '',                                                    // A: 会員ID（空）
      eventKey,                                              // B: EVENT_KEY
      convertBadge_(getFirstValue_(response['バッジ'])),     // C: バッジ
      name,                                                  // D: 氏名
      getFirstValue_(response['フリガナ']),                   // E: フリガナ
      extractVenueName_(getFirstValue_(response['所属会場'])), // F: 所属（会場名のみ）
      convertRole_(getFirstValue_(response['役割'])),         // G: 役割（「なし」→空欄）
      getFirstValue_(response['会社名']),                     // H: 会社名
      getFirstValue_(response['役職']),                       // I: 役職
      getFirstValue_(response['紹介者']),                     // J: 紹介者
      getFirstValue_(response['営業内容']),                   // K: 営業内容
      '',                                                    // L: 賞（空）
      convertBooth_(getFirstValue_(response['ブース出店'])),  // M: ブース
      '会員'                                                 // N: 区分
    ];

    // 他会場名簿マスターに追加
    const lastRow = shMaster.getLastRow();
    shMaster.getRange(lastRow + 1, 1, 1, 14).setValues([newRow]);
    Logger.log(`他会場名簿マスターに追加: ${name} (${eventKey})`);

    // ゲスト同伴の処理
    const guestWith = getFirstValue_(response['ゲスト同伴']);
    const guestInfo = getFirstValue_(response['ゲスト情報']);

    if ((guestWith === 'あり' || guestWith === '有り') && guestInfo) {
      addGuestRow_(shMaster, eventKey, guestInfo, name, getFirstValue_(response['所属会場']));
    }

  } catch (error) {
    Logger.log('onOtherVenueFormSubmit エラー: ' + error.message);
    Logger.log(error.stack);
  }
}

/**
 * ゲスト行を追加
 */
function addGuestRow_(shMaster, eventKey, guestInfo, referrerName, venue) {
  // ゲスト情報をパース（「氏名, 会社名」形式を想定）
  const guestParts = guestInfo.split(/[,、]/);
  const guestName = guestParts[0] ? guestParts[0].trim() : '';
  const guestCompany = guestParts[1] ? guestParts[1].trim() : '';

  if (!guestName) return;

  // 重複チェック
  if (isDuplicate_(shMaster, eventKey, guestName)) {
    Logger.log(`ゲスト重複のためスキップ: ${eventKey} - ${guestName}`);
    return;
  }

  const guestRow = [
    '',              // A: 会員ID（空）
    eventKey,        // B: EVENT_KEY
    '',              // C: バッジ（空）
    guestName,       // D: 氏名
    '',              // E: フリガナ（空）
    venue || '',     // F: 所属（紹介者と同じ会場）
    '',              // G: 役割（空）
    guestCompany,    // H: 会社名
    '',              // I: 役職（空）
    referrerName,    // J: 紹介者（フォーム回答者）
    '',              // K: 営業内容（空）
    '',              // L: 賞（空）
    '',              // M: ブース（空）
    'ゲスト'         // N: 区分
  ];

  const lastRow = shMaster.getLastRow();
  shMaster.getRange(lastRow + 1, 1, 1, 14).setValues([guestRow]);
  Logger.log(`ゲスト追加: ${guestName} (紹介者: ${referrerName})`);
}

/**
 * 重複チェック
 */
function isDuplicate_(shMaster, eventKey, name) {
  const masterValues = shMaster.getDataRange().getValues();

  // ヘッダーは2行目（インデックス1）、データは3行目以降
  if (masterValues.length > 2) {
    for (let i = 2; i < masterValues.length; i++) {
      const existEventKey = String(masterValues[i][1] || '').trim();  // B列: EVENT_KEY
      const existName = String(masterValues[i][3] || '').trim();       // D列: 氏名
      if (existEventKey === eventKey && existName === name) {
        return true;
      }
    }
  }
  return false;
}

/**
 * namedValuesから最初の値を取得
 */
function getFirstValue_(arr) {
  if (!arr || arr.length === 0) return '';
  return String(arr[0] || '').trim();
}

/**
 * 日付文字列をEVENT_KEYに変換
 * 例: "2月4日（水）" → "2026年2月例会"
 */
function convertDateToEventKey_(dateStr) {
  if (!dateStr) return '';

  // 「○月○日」形式からマッチ
  const match = String(dateStr).match(/(\d{1,2})月(\d{1,2})日/);
  if (!match) return dateStr;

  const month = parseInt(match[1], 10);

  // 年を推定（現在の年、または翌年）
  const now = new Date();
  let year = now.getFullYear();
  const currentMonth = now.getMonth() + 1;

  // 現在の月より小さい月の場合は来年と判断
  if (month < currentMonth - 1) {
    year++;
  }

  return `${year}年${month}月例会`;
}

/**
 * 役割変換（「なし」→ 空欄）
 */
function convertRole_(role) {
  if (!role || role === 'なし') return '';
  return role;
}

/**
 * 所属会場から会場名のみ抽出
 * 例: "九州・沖縄 - 佐賀" → "佐賀"
 */
function extractVenueName_(venue) {
  if (!venue) return '';
  const parts = String(venue).split('-');
  if (parts.length >= 2) {
    return parts.slice(1).join('-').trim();
  }
  return venue.trim();
}

/**
 * バッジ変換
 */
function convertBadge_(badge) {
  if (!badge) return '';
  const normalized = String(badge).trim();
  if (normalized === '正会員') return '正';
  if (normalized === '準会員') return '準';
  if (normalized === 'ゴールド') return 'ゴ';
  if (normalized === 'ダイヤ') return 'ダ';
  return normalized;
}

/**
 * ブース変換
 */
function convertBooth_(booth) {
  if (!booth) return '';
  const normalized = String(booth).trim();
  if (normalized === 'あり' || normalized === '有り' || normalized === '有' || normalized === 'yes') {
    return '○';
  }
  return '';
}

/**
 * 他会場名簿マスターF列の所属会場を会場名のみに一括変換
 * 例: "九州・沖縄 - 佐賀" → "佐賀"
 */
function fixExistingVenueNames() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(OTHER_VENUE_MASTER_SHEET);
  if (!sh) {
    Logger.log('他会場名簿マスターが見つかりません');
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 3) {
    Logger.log('データなし');
    return;
  }

  const range = sh.getRange(3, 6, lastRow - 2, 1); // F列、3行目から
  const values = range.getValues();
  let count = 0;

  for (let i = 0; i < values.length; i++) {
    const original = String(values[i][0] || '').trim();
    const converted = extractVenueName_(original);
    if (converted !== original) {
      values[i][0] = converted;
      count++;
    }
  }

  if (count > 0) {
    range.setValues(values);
    Logger.log(count + '件のF列を変換しました');
  } else {
    Logger.log('変換対象なし');
  }
}

// ===============================
// トリガー設定用関数
// ===============================

/**
 * フォーム送信トリガーを設定
 * この関数を一度だけ手動実行してください
 */
function setupOtherVenueFormTrigger() {
  // 既存のトリガーを削除（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onOtherVenueFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('既存トリガーを削除しました');
    }
  });

  // スプレッドシートのフォーム送信トリガーを設定
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ScriptApp.newTrigger('onOtherVenueFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log('フォーム送信トリガーを設定しました');
  Logger.log('対象スプレッドシート: ' + ss.getName());
}

/**
 * トリガー一覧を表示（確認用）
 */
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log('=== 現在のトリガー一覧 ===');
  triggers.forEach((trigger, i) => {
    Logger.log(`${i + 1}. ${trigger.getHandlerFunction()} - ${trigger.getEventType()}`);
  });
  if (triggers.length === 0) {
    Logger.log('トリガーはありません');
  }
}
