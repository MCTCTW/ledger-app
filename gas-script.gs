/**
 * 金洲工作室 · 記帳 ledger-app 後端（Google Apps Script）
 *
 * ══ 安裝步驟（一次就好）══
 * 1. 建一個新的 Google 試算表（例如命名「記帳 ledger」）
 * 2. 上方選單「擴充功能」→「Apps Script」
 * 3. 把這整個檔案貼上（取代原本的 myFunction）→ 儲存
 * 4. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    - 執行身分：我
 *    - 誰可以存取：任何人
 * 5. 複製 /exec 結尾的網址 → 貼到記帳 app 設定頁「記帳後端 (b)」
 *
 * ══ 分頁（自動建立）══
 *  _state       app 資料原檔（JSON，勿手動編輯）
 *  換匯台帳     每批換匯/代收/退款回池：台幣、外幣、匯率、已用、剩餘
 *  記帳明細     每筆帳逐商品攤台幣成本（會計主要看這張）
 *  批次成本明細 批次結算頁按「存這批結算到會計表」後寫入
 *  修改軌跡     append-only，誰在什麼時候改了什麼，永不覆蓋
 *
 * ⚠️ 更新過 script 之後要「部署 → 管理部署作業 → 編輯 → 版本選新版本」，網址才會維持不變。
 */

var TABS = { STATE:'_state', FX:'換匯台帳', LEDGER:'記帳明細', SETTLE:'批次成本明細', AUDIT:'修改軌跡' };
var CHUNK = 40000; // _state 每格 JSON 字數上限（Sheets 單格上限 50,000）

function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if(!sh) sh = ss.insertSheet(name);
  return sh;
}

function doGet(e){
  var action = (e && e.parameter && e.parameter.action) || '';
  try{
    if(action==='state') return json_({success:true, state:readState_()});
    if(action==='audit') return json_({success:true, audit:readAuditTail_(50)});
    return json_({success:true, ping:'ledger-gas', tabs:Object.keys(TABS).length});
  }catch(err){ return json_({success:false, error:String(err)}); }
}

function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if(body.state)  writeState_(body.state);
    if(body.tables){
      if(body.tables.fx)     rewriteTable_(TABS.FX, body.tables.fx);
      if(body.tables.ledger) rewriteTable_(TABS.LEDGER, body.tables.ledger);
    }
    if(body.settle) upsertSettle_(body.settle);
    if(body.audit && body.audit.length) appendAudit_(body.audit);
    return json_({success:true});
  }catch(err){
    return json_({success:false, error:String(err)});
  }finally{
    lock.releaseLock();
  }
}

// ---------- _state：app 資料原檔（JSON 分塊存） ----------
function writeState_(state){
  var sh = sheet_(TABS.STATE);
  var s = JSON.stringify(state);
  var rows = [];
  for(var i=0;i<s.length;i+=CHUNK) rows.push([rows.length, s.substr(i, CHUNK)]);
  if(!rows.length) rows=[[0,'{}']];
  sh.clearContents();
  sh.getRange(1,1,1,2).setValues([['chunk','json（勿手動編輯，app 資料原檔）']]);
  sh.getRange(2,1,rows.length,2).setValues(rows);
}
function readState_(){
  var sh = sheet_(TABS.STATE);
  var last = sh.getLastRow();
  if(last<2) return null;
  var vals = sh.getRange(2,1,last-1,2).getValues();
  vals.sort(function(a,b){ return a[0]-b[0]; });
  var s = vals.map(function(r){ return r[1]; }).join('');
  try{ return JSON.parse(s); }catch(e){ return null; }
}

// ---------- 換匯台帳 / 記帳明細：整表重寫（app 端算好、這裡只存） ----------
function rewriteTable_(name, rows){
  var sh = sheet_(name);
  sh.clearContents();
  if(!rows || !rows.length){ sh.getRange(1,1).setValue('（目前沒有資料）'); return; }
  var headers = Object.keys(rows[0]);
  var data = rows.map(function(r){ return headers.map(function(h){ var v=r[h]; return v==null?'':v; }); });
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.getRange(2,1,data.length,headers.length).setValues(data);
  sh.setFrozenRows(1);
}

// ---------- 批次成本明細：按批次 upsert（先刪同批舊列再寫） ----------
function upsertSettle_(settle){
  var sh = sheet_(TABS.SETTLE);
  var rows = settle.rows || [];
  if(!rows.length) return;
  var headers = Object.keys(rows[0]);
  var last = sh.getLastRow();
  if(last===0){
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    last = 1;
  }
  // 刪掉同批次的舊列（批次ID 欄），由下往上刪
  var existHeaders = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0];
  var idCol = existHeaders.indexOf('批次ID') + 1;
  if(idCol>0 && last>1){
    var ids = sh.getRange(2,idCol,last-1,1).getValues();
    for(var i=ids.length-1;i>=0;i--){
      if(String(ids[i][0])===String(settle.batchId)) sh.deleteRow(i+2);
    }
  }
  var data = rows.map(function(r){ return existHeaders.map(function(h){ var v=r[h]; return v==null?'':v; }); });
  sh.getRange(sh.getLastRow()+1,1,data.length,existHeaders.length).setValues(data);
}

// ---------- 修改軌跡：append-only 永不覆蓋 ----------
function appendAudit_(audit){
  var sh = sheet_(TABS.AUDIT);
  if(sh.getLastRow()===0){
    sh.getRange(1,1,1,5).setValues([['時間','操作者','動作','項目','內容']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var rows = audit.map(function(a){ return [a.t||'', a.by||'', a.action||'', a.what||'', a.detail||'']; });
  sh.getRange(sh.getLastRow()+1,1,rows.length,5).setValues(rows);
}
function readAuditTail_(n){
  var sh = sheet_(TABS.AUDIT);
  var last = sh.getLastRow();
  if(last<2) return [];
  var start = Math.max(2, last-n+1);
  return sh.getRange(start,1,last-start+1,5).getValues().map(function(r){
    return {t:r[0], by:r[1], action:r[2], what:r[3], detail:r[4]};
  });
}
