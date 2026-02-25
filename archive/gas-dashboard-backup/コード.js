/** =========================================================
 *  設定（のぶさん環境用）
 * ======================================================= */

// 管理ダッシュボード（設定シートが入っているスプシ）
const CONFIG_SHEET_ID = '195LiOaYhTq_BQw-KspFEtOHk8oHkuoVKcta8a7DYQuo';
const CONFIG_SHEET_NAME = '設定';
const CURRENT_EVENT_KEY_CELL = 'A2';  // 202512_01 が入っているセル

// 出欠＆ゲスト用スプシ
const ATTEND_GUEST_SHEET_ID = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';
const ATTEND_SHEET_NAME = '出欠状況（自動）';
const GUEST_SHEET_NAME  = 'ゲスト出欠状況（自動）';

// 会員マスタ（※シート名はあとで実名に合わせて変えてOK）
const MEMBER_SHEET_NAME = '会員名簿マスター';

// 売上報告用スプシ
const SALES_SHEET_ID = '1QYLHr7wMj0jQW5ApgWf-l1Bk9_9M5EV54FDHQp2Rkic';
const SALES_SHEET_NAME = 'sales';

// 紹介記録シート名（追加）
const REFERRAL_SHEET_NAME = '紹介記録';


/** =========================================================
 *  共通ヘルパー
 * ======================================================= */

/**
 * Webアプリのエントリポイント
 */
function doGet(e) {
  const userId = e.parameter.userId;
  
  if (!userId) {
    // 管理画面
    return HtmlService
      .createTemplateFromFile('index')
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

// チーム別売上ランキング（1〜7位まで）
function getTeamSalesRanking(eventKey) {
  // 売上報告用スプシ（上の const SALES_SHEET_ID / SALES_SHEET_NAME を利用）
  const ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  const sh = ss.getSheetByName(SALES_SHEET_NAME); // 'sales'

  const lastRow = sh.getLastRow();
  if (lastRow < 7) return [];

  // 7行目以降を取得
  const values = sh.getRange(7, 1, lastRow - 6, sh.getLastColumn()).getValues();

  const memberTeamMap = getMemberTeamMap(); // 氏名→チーム
  const teamSalesMap = {};

  values.forEach(row => {
    const amount = Number(row[5]) || 0;  // F列：金額
    const member = row[13];             // N列：氏名
    const rowKey = row[14];             // O列：eventKey

    if (!rowKey || rowKey !== eventKey) return;
    if (!amount) return;

    const team = memberTeamMap[member] || '未所属';

    if (!teamSalesMap[team]) teamSalesMap[team] = 0;
    teamSalesMap[team] += amount;
  });

  // { team: 'Aチーム', amount: 123456 } にして降順ソート
  const ranking = Object.keys(teamSalesMap).map(team => ({
    team,
    amount: teamSalesMap[team]
  })).sort((a, b) => b.amount - a.amount);

  // 上位7チームだけ
  return ranking.slice(0, 7);
}



/**
 * 設定シートから
 *  - 現在イベントキー（202512_01）
 *  - 現在イベント名（2025年12月例会）
 * を取得
 */
function getCurrentEventInfo_() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sh = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sh) {
    return { key: '', name: '' };
  }

  const keyRaw = String(sh.getRange(CURRENT_EVENT_KEY_CELL).getValue() || '').trim();
  if (!keyRaw) {
    return { key: '', name: '' };
  }

  // 例：202512_01 → 2025年12月例会
  const y = keyRaw.substring(0, 4);         // 2025
  let m = keyRaw.substring(4, 6);           // 12
  if (m.startsWith('0')) m = m.substring(1);
  const name = `${y}年${m}月例会`;

  return { key: keyRaw, name };
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
  const idxReferralCount  = 16;  // Q列：紹介在籍人数
  const idxUpdateMonth    = 19;  // T列：更新月(1〜12)

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
  const memberCount = getMemberCount_();

  const attendance       = aggregateAttendance_(eventName, memberCount);
  const attendanceDetail = aggregateAttendanceDetail_(eventName); // ★追加
  const sales            = aggregateSales_(eventName);
  const guests           = aggregateGuests_(eventName);
  const badges           = aggregateBadges_();
  const teams            = aggregateTeams_();      // すでにあるチーム集計
  const renewals         = getRenewalMembers_();   // 更新予定

  return {
    eventName,
    attendance,
    attendanceDetail,   // ★フロントに渡す
    sales,
    guests,
    badges,
    teams,
    renewals,
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
function aggregateAttendance_(eventName, memberCount) {
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

    if (key !== eventName) return;

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
function aggregateAttendanceDetail_(eventName) {
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
    if (key !== eventName) return;

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
function aggregateSales_(eventName) {
  const ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  const sh = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sh) {
    return { total: 0, dealCount: 0, ranking: [], teamRanking: [] };
  }

  // サマリ（商談件数 / 売上合計）
  const dealCount = Number(sh.getRange('E3').getValue()) || 0;
  const total     = Number(sh.getRange('F3').getValue()) || 0;

  // ── 会員別ランキング（N列:氏名 / F列:金額） ──
  const lastRow = sh.getLastRow();
  let ranking = [];
  if (lastRow >= 7) {
    const nameCol   = 14; // N列（A=1 → N=14）
    const amountCol = 6;  // F列

    const data = [];
    for (let i = 7; i <= lastRow; i++) {
      const name   = String(sh.getRange(i, nameCol).getValue() || '').trim();
      const amount = Number(sh.getRange(i, amountCol).getValue()) || 0;
      if (!name || amount === 0) continue;
      data.push({ name, amount });
    }

    data.sort((a, b) => b.amount - a.amount);

    // TOP3だけ名前を返す
    ranking = data.slice(0, 3).map(item => ({
      name: item.name
    }));
  }

  // ── チーム別売上ランキング（eventKeyでフィルタ） ──
  const eventInfo = getCurrentEventInfo_();   // { key: '202512_01', name: '2025年12月例会' }
  const eventKey  = eventInfo.key || '';
  const teamRanking = eventKey ? getTeamSalesRanking(eventKey) : [];

  return {
    total,        // 売上合計
    dealCount,    // 商談件数
    ranking,      // 会員別TOP3
    teamRanking   // チーム別1〜7位（{team, amount}）
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
    return { total: 0, approved: 0, pending: 0, list: [] };
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 3) {
    return { total: 0, approved: 0, pending: 0, list: [] };
  }

  const HEADER_ROW_INDEX = 1; // 2行目（0-based）
  const header = values[HEADER_ROW_INDEX];
  const rows   = values.slice(HEADER_ROW_INDEX + 1);

  const idxEvent   = findColumnIndex_(header, ['eventKey', 'イベントキー']);
  const idxGuest   = findColumnIndex_(header, ['氏名', 'ゲスト名']);
  const idxIntro   = findColumnIndex_(header, ['紹介者', '紹介者名']);
  const idxApprove = findColumnIndex_(header, ['承認']);

  const cEvent   = idxEvent   !== -1 ? idxEvent   : 2;  // C列
  const cGuest   = idxGuest   !== -1 ? idxGuest   : 3;  // D列
  const cIntro   = idxIntro   !== -1 ? idxIntro   : 7;  // H列
  const cApprove = idxApprove !== -1 ? idxApprove : 10; // K列

  let total    = 0;
  let approved = 0;

  const list = [];

  rows.forEach(row => {
    const key = String(row[cEvent] || '').trim();
    if (key !== eventName) return;

    const guestName = String(row[cGuest] || '').trim();
    const introName = String(row[cIntro] || '').trim();
    const isApproved = row[cApprove] === true; // チェックボックス

    total++;
    if (isApproved) approved++;

    list.push({
      guestName,
      introName,
      approved: isApproved,
    });
  });

  const pending = total - approved;

  return { total, approved, pending, list };
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
