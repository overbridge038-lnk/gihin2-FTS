/**
 * 誤品照合 FTS版 → スプレッドシート受け取り用 GAS（Webアプリ）
 *
 * 【設定のしかた】
 * 1. 記録を貯めたいスプレッドシートを開く → 拡張機能 → Apps Script
 * 2. このコードを貼り付けて保存
 * 3. （任意）プロジェクトの設定 → スクリプト プロパティ に「KEY」を作り、合言葉を入れる
 *    → アプリの⚙設定「合言葉」に同じ文字を入れる（空欄ならチェックしない）
 * 4. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *    実行ユーザー：自分 ／ アクセスできるユーザー：全員 → デプロイ
 * 5. 出てきた「ウェブアプリのURL」（…/exec）を、アプリの⚙設定「GAS連携」に入れて「接続テスト」
 *
 * 同じ「キー（端末ID-通し番号）」は1行にまとまるので、2回送られても増えません。
 */
const SHEET_NAME = '照合記録';
const HEADER = ['キー','端末ID','通し番号','日付','タイムスタンプ','客先品番','社内品番','判定','客先QR全文','備考','処置記録(JSON)','受信日時'];

function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e){
  try{
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const need = PropertiesService.getScriptProperties().getProperty('KEY');
    if(need && body.key !== need) return out_({ok:false, error:'合言葉が違います'});
    if(body.action === 'ping') return out_({ok:true});
    if(body.action !== 'append' || !Array.isArray(body.records)) return out_({ok:false, error:'送られた内容が正しくありません'});

    const lock = LockService.getScriptLock(); lock.waitLock(30000);
    try{
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
      if(sh.getLastRow() === 0){ sh.appendRow(HEADER); sh.setFrozenRows(1); }
      const last = sh.getLastRow();
      const have = new Set(last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : []);
      const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
      const rows = [];
      body.records.forEach(r => {
        const key = String(r.key || '');
        if(!key || have.has(key)) return;            // すでにある記録は足さない（二重送信対策）
        have.add(key);
        rows.push([key, r.dev || '', r.no || '', r.day || '', r.t || '', r.hin || '', r.inner || '', r.r || '', r.raw || '', r.note || '',
                   r.treat ? JSON.stringify(r.treat) : '', now]);
      });
      if(rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADER.length).setValues(rows);
      // received＝受け取った件数（すでにあった分も含む）。アプリはこの数が送った数と合ったときだけ「送信済み」にする
      return out_({ok:true, received:body.records.length, added:rows.length});
    }finally{ lock.releaseLock(); }
  }catch(err){
    return out_({ok:false, error:String(err)});
  }
}

// ブラウザで URL を開いたときの確認用
function doGet(){ return out_({ok:true, message:'誤品照合 受け取り用GAS（POSTで使います）'}); }
