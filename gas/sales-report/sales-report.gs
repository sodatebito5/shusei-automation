/** Code.gs（上書き対応：line_user_id だけでUPSERT） **/
const SHEET_ID   = '1QYLHr7wMj0jQW5ApgWf-l1Bk9_9M5EV54FDHQp2Rkic';
const SHEET_NAME = 'Sales';
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

    // 受け取り（report_date は任意に）
    const {
      line_user_id = '',
      display_name = '',
      report_date  = '',

      deals_count      = '',
      sales_amount     = '',
      join_next_seat   = '',
      desired_industry = '',
      next_guest       = '',

      product = '',
      amount  = '',
      payment = '',
      notes   = ''
    } = body;

    if (!line_user_id) {
      return _json({ ok:false, error:'missing field: line_user_id' });
    }

    // 数値整形
    const nDeals  = deals_count  === '' ? '' : Number(deals_count);
    const nSales  = sales_amount === '' ? '' : Number(sales_amount);
    const nAmount = amount       === '' ? '' : Number(amount);

    // シート準備
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    const headers = [
      'timestamp','line_user_id','display_name','report_date',
      'deals_count','sales_amount','join_next_seat','desired_industry','next_guest',
      'product','amount','payment','notes'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    } else {
      const firstRow = sheet.getRange(1,1,1,headers.length).getValues()[0];
      const needFix  = headers.some((h, i) => firstRow[i] !== h);
      if (needFix) {
        sheet.insertRowBefore(1);
        sheet.getRange(1,1,1,headers.length).setValues([headers]);
      }
    }

    // 1行分のデータをヘッダー順で整形
    const rowData = [
      new Date(),
      line_user_id,
      display_name,
      report_date,      // 任意。渡さなければ空欄で上書き
      nDeals,
      nSales,
      join_next_seat,
      desired_industry,
      next_guest,
      product,
      nAmount,
      payment,
      notes
    ];

    // 既存検索：line_user_id が一致する最新行を上から（下から）探す
    const lastRow = sheet.getLastRow();
    const lastCol = headers.length;

    if (lastRow > 1) {
      const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues(); // データ部
      const IDX = { line_user_id: headers.indexOf('line_user_id') };

      let targetRow = -1; // シート上の絶対行番号
      for (let i = values.length - 1; i >= 0; i--) {
        if (values[i][IDX.line_user_id] === line_user_id) {
          targetRow = i + 2; // データは2行目開始なので +2
          break;
        }
      }

      if (targetRow !== -1) {
        // 上書き（timestamp は毎回更新）
        sheet.getRange(targetRow, 1, 1, lastCol).setValues([rowData]);
        return _json({ ok:true, mode:'updated', row: targetRow });
      }
    }

    // 見つからなければ追記（= 初回送信）
    sheet.appendRow(rowData);
    return _json({ ok:true, mode:'inserted', row: sheet.getLastRow() });

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
