/** Code.gs（event_key対応版） **/
const SHEET_ID   = '1QYLHr7wMj0jQW5ApgWf-l1Bk9_9M5EV54FDHQp2Rkic';
const SHEET_NAME = 'sales';
const API_KEY    = 'shusei_2025_secret_1016';

function doGet(e) {
  return _json({ ok: true, ping: 'pong' });
}

function doOptions(e) {
  return _json({ ok: true });
}

function doPost(e) {
  try {
    if (!e.postData) return _json({ ok:false, error:'no post data' });

    const body = JSON.parse(e.postData.contents || '{}');

    // 認証
    if (body.api_key !== API_KEY) {
      return _json({ ok:false, error:'unauthorized' });
    }

    // 受け取り
    const {
      line_user_id = '',
      display_name = '',
      report_date  = '',
      deals_count      = '',
      sales_amount     = '',
      join_next_seat   = '',
      desired_industry = '',
      next_guest       = '',
      meetings_count   = '',
      event_key  = ''
    } = body;

    if (!line_user_id) {
      return _json({ ok:false, error:'missing field: line_user_id' });
    }

    // 数値整形
    const nDeals    = deals_count    === '' ? '' : Number(deals_count);
    const nSales    = sales_amount   === '' ? '' : Number(sales_amount);
    const nMeetings = meetings_count === '' ? '' : Number(meetings_count);

    // シート準備
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return _json({ ok:false, error:'sheet not found: ' + SHEET_NAME });
    }

    const lastRow = sheet.getLastRow();
    const DATA_START_ROW = 7;

    // 既存行を検索（B列: line_user_id, O列: event_key で一致）
    let targetRow = -1;

    if (lastRow >= DATA_START_ROW && event_key) {
      const dataRows = lastRow - DATA_START_ROW + 1;
      const data = sheet.getRange(DATA_START_ROW, 1, dataRows, 16).getValues();

      for (let i = data.length - 1; i >= 0; i--) {
        const rowUserId   = String(data[i][1] || '').trim();  // B列
        const rowEventKey = String(data[i][14] || '').trim(); // O列

        if (rowUserId === line_user_id && rowEventKey === event_key) {
          targetRow = i + DATA_START_ROW;
          break;
        }
      }
    }

    const now = new Date();

    if (targetRow !== -1) {
      // 既存行を上書き
      sheet.getRange(targetRow, 1).setValue(now);              // A: timestamp
      sheet.getRange(targetRow, 4).setValue(report_date);      // D: 回答日
      sheet.getRange(targetRow, 5).setValue(nDeals);           // E: 成約件数
      sheet.getRange(targetRow, 6).setValue(nSales);           // F: 金額
      sheet.getRange(targetRow, 7).setValue(join_next_seat);   // G: 同席希望
      sheet.getRange(targetRow, 8).setValue(desired_industry); // H: 入会希望業種
      sheet.getRange(targetRow, 9).setValue(next_guest);       // I: 次回ゲスト
      sheet.getRange(targetRow, 15).setValue(event_key);       // O: event_key
      sheet.getRange(targetRow, 16).setValue(nMeetings);       // P: 商談件数

      return _json({ ok:true, mode:'updated', row: targetRow });
    }

    // 新規行を追加
    const newRow = lastRow + 1;
    sheet.getRange(newRow, 1).setValue(now);              // A: timestamp
    sheet.getRange(newRow, 2).setValue(line_user_id);     // B: LINEユーザーID
    sheet.getRange(newRow, 3).setValue(display_name);     // C: LINE名
    sheet.getRange(newRow, 4).setValue(report_date);      // D: 回答日
    sheet.getRange(newRow, 5).setValue(nDeals);           // E: 成約件数
    sheet.getRange(newRow, 6).setValue(nSales);           // F: 金額
    sheet.getRange(newRow, 7).setValue(join_next_seat);   // G: 同席希望
    sheet.getRange(newRow, 8).setValue(desired_industry); // H: 入会希望業種
    sheet.getRange(newRow, 9).setValue(next_guest);       // I: 次回ゲスト
    sheet.getRange(newRow, 14).setValue(display_name);    // N: 氏名
    sheet.getRange(newRow, 15).setValue(event_key);       // O: event_key
    sheet.getRange(newRow, 16).setValue(nMeetings);       // P: 商談件数

    return _json({ ok:true, mode:'inserted', row: newRow });

  } catch (err) {
    return _json({ ok:false, error:String(err) });
  }
}

/** 共通：JSON出力 */
function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
