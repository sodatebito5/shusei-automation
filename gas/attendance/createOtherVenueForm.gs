// ★ 既存フォームのID（一度作成したら固定）
const OTHER_VENUE_FORM_ID = '1pqg-V0AyPBzNdmIYUVCS55TtQnFWzY4efhkZISSm7OA';

/**
 * 既存フォームを完全に再構築（URLはそのまま）
 */
function rebuildOtherVenueForm() {
  const form = FormApp.openById(OTHER_VENUE_FORM_ID);
  const ssId = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';

  // 既存の項目をすべて削除（末尾から逆順に削除）
  const items = form.getItems();
  for (let i = items.length - 1; i >= 0; i--) {
    try {
      form.deleteItem(i);
    } catch (e) {
      Logger.log('項目削除スキップ: index=' + i + ', error=' + e);
    }
  }

  form.setTitle('他会場参加申込フォーム');
  form.setDescription('守成クラブ福岡飯塚への他会場からの参加申込フォームです。');
  form.setConfirmationMessage('受付完了しました。ご参加ありがとうございます。当日お会いできることを楽しみにしております。');

  // 会場マスターからデータ取得
  const ss = SpreadsheetApp.openById(ssId);
  const venueMaster = ss.getSheetByName('会場マスター');
  const venueData = venueMaster.getDataRange().getValues();

  const venuesByRegion = {};
  for (let i = 1; i < venueData.length; i++) {
    const region = venueData[i][0];
    const venue = venueData[i][2];
    if (!region || !venue) continue;
    if (!venuesByRegion[region]) venuesByRegion[region] = [];
    if (!venuesByRegion[region].includes(venue)) venuesByRegion[region].push(venue);
  }

  const regions = Object.keys(venuesByRegion).sort();
  const allVenues = [];
  regions.forEach(region => {
    venuesByRegion[region].sort().forEach(venue => {
      allVenues.push(region + ' - ' + venue);
    });
  });

  // 設定シートから例会日程を取得
  const configSheetId = '1R4GR1GZg6mJP9zPX5MTE0IsYEAdVLNIM314o7vBqrg8';
  const configSs = SpreadsheetApp.openById(configSheetId);
  const configSheet = configSs.getSheetByName('設定');

  const eventOptions = [];
  if (configSheet) {
    const lastRow = configSheet.getLastRow();
    const dateRange = configSheet.getRange(2, 4, lastRow - 1, 1);
    const dates = dateRange.getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];

    for (let i = 0; i < dates.length && eventOptions.length < 3; i++) {
      const eventDate = dates[i][0];
      if (eventDate) {
        const d = new Date(eventDate);
        if (d >= today) {
          const month = d.getMonth() + 1;
          const day = d.getDate();
          const dayOfWeek = dayNames[d.getDay()];
          eventOptions.push(month + '月' + day + '日（' + dayOfWeek + '）');
        }
      }
    }
  }

  if (eventOptions.length === 0) {
    const now = new Date();
    for (let i = 1; i <= 3; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      eventOptions.push(d.getFullYear() + '年' + (d.getMonth() + 1) + '月例会');
    }
  }

  // フォーム項目を追加
  form.addListItem().setTitle('参加する例会').setChoiceValues(eventOptions).setRequired(true);
  form.addListItem().setTitle('所属会場').setChoiceValues(allVenues).setRequired(true);
  form.addListItem().setTitle('バッジ').setChoiceValues(['正会員', '準会員', 'ゴールド', 'ダイヤ']).setRequired(true);
  form.addListItem().setTitle('役割').setChoiceValues(['代表', '副代表', '世話人', '事務局', '会計', '全連常務理事', 'なし']).setRequired(false);
  form.addTextItem().setTitle('氏名').setRequired(true);
  form.addTextItem().setTitle('フリガナ').setRequired(true);
  form.addTextItem().setTitle('会社名').setRequired(true);
  form.addTextItem().setTitle('役職').setRequired(true);
  form.addParagraphTextItem().setTitle('営業内容').setRequired(true);
  form.addTextItem().setTitle('紹介者').setRequired(true);
  form.addMultipleChoiceItem().setTitle('ブース出店').setChoiceValues(['あり', 'なし']).setRequired(true);
  form.addMultipleChoiceItem().setTitle('ゲスト同伴').setChoiceValues(['あり', 'なし']).setRequired(false);
  form.addParagraphTextItem().setTitle('ゲスト情報').setHelpText('ゲスト同伴ありの場合：お名前（フリガナ）、会社名、役職').setRequired(false);
  form.addTextItem().setTitle('メールアドレス').setHelpText('確認メールを送信します').setRequired(true);

  // 回答先スプレッドシートを設定（既に同じ宛先なら再設定しない）
  try {
    const currentDest = form.getDestinationId();
    if (currentDest !== ssId) {
      form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
      Logger.log('回答先を再設定しました');
    } else {
      Logger.log('回答先は設定済み: ' + currentDest);
    }
  } catch (e) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
    Logger.log('回答先を新規設定しました');
  }

  Logger.log('フォーム再構築完了');
  Logger.log('編集URL: ' + form.getEditUrl());
  Logger.log('回答URL: ' + form.getPublishedUrl());
}

/**
 * フォームの回答先スプレッドシートを再リンク
 * 新しい「フォームの回答」シートが作成される
 */
function relinkFormToSpreadsheet() {
  const form = FormApp.openById(OTHER_VENUE_FORM_ID);
  const ssId = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';

  // 既存のリンクを解除してから再設定
  try {
    form.removeDestination();
    Logger.log('既存のリンクを解除しました');
  } catch (e) {
    Logger.log('リンク解除スキップ: ' + e);
  }

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);

  // 新しく作成されたシート名を特定
  const ss = SpreadsheetApp.openById(ssId);
  const sheets = ss.getSheets();
  const responseSheets = sheets.filter(s => s.getName().startsWith('フォームの回答'));
  responseSheets.forEach(s => {
    Logger.log('回答シート: ' + s.getName());
  });

  Logger.log('フォーム回答先を再リンクしました');
  Logger.log('スプレッドシートで新しい「フォームの回答」シートを確認してください');
}

/**
 * 既存の他会場参加申込フォームを更新（例会選択肢のみ更新）
 */
function updateOtherVenueForm() {
  const form = FormApp.openById(OTHER_VENUE_FORM_ID);

  const configSheetId = '1R4GR1GZg6mJP9zPX5MTE0IsYEAdVLNIM314o7vBqrg8';
  const configSs = SpreadsheetApp.openById(configSheetId);
  const configSheet = configSs.getSheetByName('設定');

  const eventOptions = [];
  if (configSheet) {
    const lastRow = configSheet.getLastRow();
    const dateRange = configSheet.getRange(2, 4, lastRow - 1, 1);
    const dates = dateRange.getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];

    for (let i = 0; i < dates.length && eventOptions.length < 3; i++) {
      const eventDate = dates[i][0];
      if (eventDate) {
        const d = new Date(eventDate);
        if (d >= today) {
          const month = d.getMonth() + 1;
          const day = d.getDate();
          const dayOfWeek = dayNames[d.getDay()];
          eventOptions.push(month + '月' + day + '日（' + dayOfWeek + '）');
        }
      }
    }
  }

  if (eventOptions.length === 0) {
    const now = new Date();
    for (let i = 1; i <= 3; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      eventOptions.push(d.getFullYear() + '年' + (d.getMonth() + 1) + '月例会');
    }
  }

  const items = form.getItems();
  for (const item of items) {
    if (item.getTitle() === '参加する例会') {
      item.asListItem().setChoiceValues(eventOptions);
      Logger.log('例会選択肢を更新: ' + eventOptions.join(', '));
      break;
    }
  }

  Logger.log('フォーム更新完了');
}

/**
 * 既存フォームの選択肢のみ更新（バッジ・役割）
 * フォーム構造・URLは一切変わらない
 */
function updateFormChoices() {
  const form = FormApp.openById(OTHER_VENUE_FORM_ID);
  const items = form.getItems();

  for (const item of items) {
    const title = item.getTitle();
    if (title === 'バッジ') {
      item.asListItem().setChoiceValues(['正会員', '準会員', 'ゴールド', 'ダイヤ']);
      Logger.log('バッジ選択肢を更新');
    } else if (title === '役割') {
      item.asListItem().setChoiceValues(['代表', '副代表', '世話人', '事務局', '会計', '全連常務理事', 'なし']);
      Logger.log('役割選択肢を更新');
    }
  }

  form.setConfirmationMessage('受付完了しました。ご参加ありがとうございます。当日お会いできることを楽しみにしております。');
  Logger.log('確認メッセージを更新');
  Logger.log('フォーム選択肢更新完了（URL変更なし）');
}

/**
 * 出欠確認シートのシート一覧を確認
 */
function listSheetsInAttendance() {
  const ssId = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';
  const ss = SpreadsheetApp.openById(ssId);
  const sheets = ss.getSheets();

  Logger.log('=== シート一覧 ===');
  sheets.forEach((sheet, i) => {
    Logger.log((i + 1) + '. ' + sheet.getName());
  });
}

/**
 * 会場マスターのデータを確認
 */
function checkVenueMaster() {
  const ssId = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';
  const ss = SpreadsheetApp.openById(ssId);

  // 会場マスターという名前のシートを探す
  const possibleNames = ['会場マスター', '会場マスタ', '会場一覧', '他会場マスター', '他会場名簿マスター'];

  for (const name of possibleNames) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      Logger.log('シート発見: ' + name);
      const data = sheet.getDataRange().getValues();
      Logger.log('ヘッダー: ' + JSON.stringify(data[0]));
      Logger.log('データ行数: ' + (data.length - 1));
      if (data.length > 1) {
        Logger.log('サンプル行1: ' + JSON.stringify(data[1]));
        if (data.length > 2) {
          Logger.log('サンプル行2: ' + JSON.stringify(data[2]));
        }
      }
      return;
    }
  }

  Logger.log('会場マスターが見つかりません');
}

/**
 * 他会場参加申込フォームを自動作成（シンプル版 - 1列）
 */
function createOtherVenueFormV4() {
  const ssId = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';

  // 会場マスターからデータ取得
  const ss = SpreadsheetApp.openById(ssId);
  const venueMaster = ss.getSheetByName('会場マスター');
  const venueData = venueMaster.getDataRange().getValues();

  // 地域ごとに会場をグループ化
  const venuesByRegion = {};
  for (let i = 1; i < venueData.length; i++) {
    const region = venueData[i][0];
    const venue = venueData[i][2];
    if (!region || !venue) continue;

    if (!venuesByRegion[region]) {
      venuesByRegion[region] = [];
    }
    if (!venuesByRegion[region].includes(venue)) {
      venuesByRegion[region].push(venue);
    }
  }

  const regions = Object.keys(venuesByRegion).sort();
  Logger.log('地域数: ' + regions.length);
  Logger.log('地域: ' + regions.join(', '));

  // 全会場リストを作成（地域 - 会場名 形式）
  const allVenues = [];
  regions.forEach(region => {
    const venues = venuesByRegion[region].sort();
    venues.forEach(venue => {
      allVenues.push(region + ' - ' + venue);
    });
  });
  Logger.log('会場数: ' + allVenues.length);

  // フォーム作成
  const form = FormApp.create('他会場参加申込フォーム');
  form.setDescription('守成クラブ福岡飯塚への他会場からの参加申込フォームです。');
  form.setConfirmationMessage('受付完了しました。\nご参加ありがとうございます。当日お会いできることを楽しみにしております。');

  // === 基本情報 ===

  // 1. 参加例会（設定シートのD列から取得）
  const configSheetId = '1R4GR1GZg6mJP9zPX5MTE0IsYEAdVLNIM314o7vBqrg8';
  const configSs = SpreadsheetApp.openById(configSheetId);
  const configSheet = configSs.getSheetByName('設定');

  const eventOptions = [];
  if (configSheet) {
    // D列（開催日）を読み取り、今日以降の日付を例会として追加（直近3回分のみ）
    const lastRow = configSheet.getLastRow();
    const dateRange = configSheet.getRange(2, 4, lastRow - 1, 1); // D列、2行目から
    const dates = dateRange.getValues();

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];

    for (let i = 0; i < dates.length && eventOptions.length < 3; i++) {
      const eventDate = dates[i][0];
      if (eventDate) {
        const d = new Date(eventDate);
        if (d >= today) {
          // 日付を「○月○日（曜日）」形式に変換
          const month = d.getMonth() + 1;
          const day = d.getDate();
          const dayOfWeek = dayNames[d.getDay()];

          eventOptions.push(month + '月' + day + '日（' + dayOfWeek + '）');
        }
      }
    }
  }

  // フォールバック：設定シートから取得できない場合は6ヶ月分を自動生成
  if (eventOptions.length === 0) {
    const now = new Date();
    for (let i = 1; i <= 6; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      eventOptions.push(d.getFullYear() + '年' + (d.getMonth() + 1) + '月例会');
    }
  }

  Logger.log('例会選択肢: ' + eventOptions.join(', '));

  form.addListItem()
    .setTitle('参加する例会')
    .setChoiceValues(eventOptions)
    .setRequired(true);

  // 2. 所属会場（全会場を1つのドロップダウンに）
  form.addListItem()
    .setTitle('所属会場')
    .setHelpText('地域 - 会場名 の形式で選択してください')
    .setChoiceValues(allVenues)
    .setRequired(true);

  // === 個人情報 ===

  // 3. バッジ
  form.addListItem()
    .setTitle('バッジ')
    .setChoiceValues(['正会員', '準会員', 'ゴールド', 'ダイヤ'])
    .setRequired(true);

  // 4. 役割
  form.addListItem()
    .setTitle('役割')
    .setChoiceValues(['代表', '副代表', '世話人', '事務局', '会計', '全連常務理事', 'なし'])
    .setRequired(false);

  // 5. 氏名
  form.addTextItem()
    .setTitle('氏名')
    .setRequired(true);

  // 6. フリガナ
  form.addTextItem()
    .setTitle('フリガナ')
    .setRequired(true);

  // 7. 会社名
  form.addTextItem()
    .setTitle('会社名')
    .setRequired(true);

  // 8. 役職
  form.addTextItem()
    .setTitle('役職')
    .setRequired(true);

  // 9. 営業内容
  form.addParagraphTextItem()
    .setTitle('営業内容')
    .setRequired(true);

  // 10. 紹介者
  form.addTextItem()
    .setTitle('紹介者')
    .setRequired(true);

  // 11. ブース出店
  form.addMultipleChoiceItem()
    .setTitle('ブース出店')
    .setChoiceValues(['あり', 'なし'])
    .setRequired(true);

  // 11. ゲスト同伴
  form.addMultipleChoiceItem()
    .setTitle('ゲスト同伴')
    .setChoiceValues(['あり', 'なし'])
    .setRequired(false);

  // 12. ゲスト情報
  form.addParagraphTextItem()
    .setTitle('ゲスト情報')
    .setHelpText('ゲスト同伴ありの場合：お名前（フリガナ）、会社名、役職')
    .setRequired(false);

  // 13. メールアドレス
  form.addTextItem()
    .setTitle('メールアドレス')
    .setHelpText('確認メールを送信します')
    .setRequired(true);

  // 回答先スプレッドシートを設定
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);

  // 作成完了ログ
  Logger.log('フォーム作成完了（セクション分岐版）');
  Logger.log('編集URL: ' + form.getEditUrl());
  Logger.log('回答URL: ' + form.getPublishedUrl());

  return {
    editUrl: form.getEditUrl(),
    publishedUrl: form.getPublishedUrl()
  };
}
