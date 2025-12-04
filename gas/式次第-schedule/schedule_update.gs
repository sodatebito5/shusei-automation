// ===============================
// スプレッドシート開いたときにメニュー追加
// ===============================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('日程更新')
    .addItem('日程を更新', 'updateMeetingDate')
    .addToUi();
}

// ===============================
// 日程取得・更新処理
// ===============================
function updateMeetingDate() {
  // スプレッドシートIDとシート名
  const SHEET_ID = '1IPyjDi3uD-pSxtkF9JK7Uc5isi4lNw6nQKpv9hWUvic';
  const SHEET_NAME = '式次第';
  
  // 基準となる会場
  const BASE_VENUE = '福岡飯塚';
  
  // 基準日と回数（2025.11.12を第91回とする）
  const BASE_DATE_REF = new Date(2025, 10, 12); // 2025年11月12日
  const BASE_COUNT = 91;
  
  // 対象会場リスト（27会場）
  const TARGET_VENUES = [
    '沖縄北部やんばる',
    'ヒルノ福岡セントラル',
    '久留米セントラル',
    '中津諭吉の里',
    'ながさき出島',
    '長崎いさはや',
    '熊本県北玉名',
    '北九州八幡',
    '北九州門司',
    'ヒルノ沖縄',
    'ヒルノ熊本',
    '沖縄中部',
    '福岡イースト',
    '福岡中央',
    '福岡筑後',
    '鹿児島南',
    '北九州',
    '博多',
    '宮崎',
    '那覇',
    '鳥栖',
    '沖縄',
    '延岡',
    '琉球',
    '宗像福津',
    '小倉',
    '都城'
  ];
  
  // ページ取得
  const url = 'https://shusei-honbu.jp/shusei';
  const response = UrlFetchApp.fetch(url);
  const html = response.getContentText();
  
  // <tbody>...</tbody> の部分だけを抽出
  const tbodyMatch = html.match(/<tbody>([\s\S]*?)<\/tbody>/i);
  
  if (!tbodyMatch) {
    Logger.log('tbodyが見つかりませんでした');
    return;
  }
  
  const tbody = tbodyMatch[1];
  const tableRowPattern = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  const rows = tbody.match(tableRowPattern);
  
  if (!rows) {
    Logger.log('行が見つかりませんでした');
    return;
  }
  
  // 全会場の日程を抽出
  const schedules = [];
  
  for (let row of rows) {
    // <td>タグを個別に抽出
    const tdPattern = /<td[^>]*>([\s\S]*?)<\/td>/gi;
    const cells = [];
    let match;
    
    while ((match = tdPattern.exec(row)) !== null) {
      const cellText = match[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
      cells.push(cellText);
    }
    
    if (cells.length < 3) continue;
    
    const dateText = cells[0];
    const venueText = cells[2];
    
    // 日付パターン
    const datePattern = /(\d+)月(\d+)日\(([月火水木金土日])\)/;
    const dateMatch = dateText.match(datePattern);
    
    if (!dateMatch) continue;
    
    const month = parseInt(dateMatch[1]);
    const day = parseInt(dateMatch[2]);
    const weekday = dateMatch[3];
    
    // 会場名を検索
    let venueName = '';
    
    if (venueText.indexOf(BASE_VENUE) > -1) {
      venueName = BASE_VENUE;
    } else {
      for (let venue of TARGET_VENUES) {
        if (venueText.indexOf(venue) > -1) {
          venueName = venue;
          break;
        }
      }
    }
    
    if (venueName) {
      const now = new Date();
      let year = now.getFullYear();
      
      if (now.getMonth() >= 10 && month <= 3) {
        year++;
      }
      
      const dateObj = new Date(year, month - 1, day);
      
      schedules.push({
        venue: venueName,
        month: month,
        day: day,
        weekday: weekday,
        dateObj: dateObj,
        dateStr: `${month}/${day}（${weekday}）`
      });
    }
  }
  
  // 福岡飯塚の日程を基準日として取得
  const baseSchedule = schedules.find(s => s.venue === BASE_VENUE);
  
  if (!baseSchedule) {
    Logger.log('福岡飯塚の日程が見つかりませんでした');
    return;
  }
  
  const baseDate = baseSchedule.dateObj;
  
  // 各会場の基準日より後の最初の1回だけを抽出
  const venueFirstSchedule = {};
  
  for (let schedule of schedules) {
    if (TARGET_VENUES.includes(schedule.venue) && schedule.dateObj > baseDate) {
      if (!venueFirstSchedule[schedule.venue] || 
          schedule.dateObj < venueFirstSchedule[schedule.venue].dateObj) {
        venueFirstSchedule[schedule.venue] = schedule;
      }
    }
  }
  
  const futureSchedules = Object.values(venueFirstSchedule);
  futureSchedules.sort((a, b) => a.dateObj - b.dateObj);
  
  // 同じ日の会場をグループ化
  const groupedByDate = {};
  
  for (let schedule of futureSchedules) {
    const key = schedule.dateStr;
    if (!groupedByDate[key]) {
      groupedByDate[key] = [];
    }
    groupedByDate[key].push(schedule.venue);
  }
  
  // 配置用データ作成
  const output = [];
  for (let dateStr in groupedByDate) {
    const venues = groupedByDate[dateStr].join(' / ');
    output.push({
      date: dateStr,
      venues: venues
    });
  }
  
  // 回数を計算
  const monthDiff = (baseDate.getFullYear() - BASE_DATE_REF.getFullYear()) * 12 
                    + (baseDate.getMonth() - BASE_DATE_REF.getMonth());
  const meetingCount = BASE_COUNT + monthDiff;
  
  // スプレッドシートに書き込み
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  // A2に例会回数を含むタイトルを書き込み
  const titleText = `守成クラブ福岡飯塚　第${meetingCount}回 仕事バンバンプラザ 例会　式次第`;
  const cellA2 = sheet.getRange('A2');
  cellA2.setValue(titleText);
  
  const countText = `第${meetingCount}回`;
  const startIndex = titleText.indexOf(countText);
  const endIndex = startIndex + countText.length;
  
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(titleText)
    .setTextStyle(startIndex, endIndex, 
      SpreadsheetApp.newTextStyle()
        .setFontSize(18)
        .build())
    .build();
  
  cellA2.setRichTextValue(richText);
  
  // C4に福岡飯塚の日程を書き込み
  const iizukaText = `◎ 来月の例会は${baseSchedule.dateStr}　　原則　第二水曜日開催です！！`;
  const cellC4 = sheet.getRange('C4');
  cellC4.setValue(iizukaText);
  cellC4.setFontColor('#FF0000');
  cellC4.setFontLine('underline');
  cellC4.setFontSize(14);
  
  // 既存データをクリア
  const dateRanges = ['C8:C12', 'G8:G12', 'K8:K12'];
  const venueRanges = ['D8:D12', 'H8:H12', 'L8:L12'];
  
  for (let range of dateRanges) {
    sheet.getRange(range).clearContent();
  }
  for (let range of venueRanges) {
    sheet.getRange(range).clearContent();
  }
  
  // データを配置
  let index = 0;
  const columns = [
    { dateCol: 'C', venueCol: 'D' },
    { dateCol: 'G', venueCol: 'H' },
    { dateCol: 'K', venueCol: 'L' }
  ];
  
  for (let col of columns) {
    for (let row = 8; row <= 12; row++) {
      if (index >= output.length) break;
      
      sheet.getRange(`${col.dateCol}${row}`).setValue(output[index].date);
      sheet.getRange(`${col.venueCol}${row}`).setValue(output[index].venues);
      
      index++;
    }
    if (index >= output.length) break;
  }
  
  Logger.log('書き込み完了');
}