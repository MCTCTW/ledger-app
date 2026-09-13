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
 *  _state 備份   每次覆寫前自動留一版（最多 5 版，救命用）
 *  換匯台帳     每批換匯/代收/退款回池：台幣、外幣、匯率、已用、剩餘
 *  記帳明細     每筆帳逐商品攤台幣成本（會計主要看這張）
 *  批次成本明細 批次結算頁按「存這批結算到會計表」後寫入
 *  修改軌跡     append-only，誰在什麼時候改了什麼，永不覆蓋
 *
 * ⚠️ 更新過 script 之後要「部署 → 管理部署作業 → 編輯 → 版本選新版本」，網址才會維持不變。
 *
 * ══ v2（2026-08-21）══
 * 寫入改「合併」：收到 mergeState 的資料會跟表上現有資料取聯集（同一筆取較新的、刪除用 dels 註記），
 * 不再整包覆蓋 → 兩個人同時開著記帳也不會把對方剛存的洗掉。修改軌跡新增「紀錄ID」欄用來去重。
 */

var TABS = { STATE:'_state', BAK:'_state 備份', FX:'換匯台帳', LEDGER:'記帳明細', SETTLE:'批次成本明細', AUDIT:'修改軌跡' };
var CHUNK = 40000; // _state 每格 JSON 字數上限（Sheets 單格上限 50,000）
var VER = 3;       // v3：換匯台帳／記帳明細覆寫前也留備份＋列數驟減寫進修改軌跡
                   // v2：強制合併寫入（不再整包覆蓋）＋拒收舊版分頁＋備份 5 版＋修改軌跡去重

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
    if(action==='state') return json_({success:true, state:readState_(), ver:VER});
    if(action==='audit') return json_({success:true, audit:readAuditTail_(Number(e.parameter.n)||50)});
    return json_({success:true, ping:'ledger-gas', ver:VER, tabs:Object.keys(TABS).length});
  }catch(err){ return json_({success:false, error:String(err)}); }
}

function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    // 🚨 後門上鎖：沒帶 mergeState＝這個分頁還在跑舊版（整包覆蓋）程式 → 直接拒收，不讓它寫。
    //    正常 v2 前端每次都會帶 mergeState:true；批次結算的 {settle} 不受影響。
    if((body.state || body.tables) && !body.mergeState){
      return json_({success:false, error:'STALE_CLIENT', need:'reload', ver:VER});
    }
    if(body.state)  writeState_(mergeStates_(readState_(), body.state));
    if(body.tables){
      if(body.tables.fx)     rewriteTable_(TABS.FX, body.tables.fx);
      if(body.tables.ledger) rewriteTable_(TABS.LEDGER, body.tables.ledger);
    }
    if(body.settle) upsertSettle_(body.settle);
    if(body.audit && body.audit.length) appendAudit_(body.audit);
    return json_({success:true, ver:VER});
  }catch(err){
    return json_({success:false, error:String(err)});
  }finally{
    lock.releaseLock();
  }
}

// ---------- _state：app 資料原檔（JSON 分塊存） ----------
function putChunks_(sh, s, title){
  var rows = [];
  for(var i=0;i<s.length;i+=CHUNK) rows.push([rows.length, s.substr(i, CHUNK)]);
  if(!rows.length) rows=[[0,'{}']];
  sh.clearContents();
  sh.getRange(1,1,1,2).setValues([['chunk', title]]);
  sh.getRange(2,1,rows.length,2).setValues(rows);
}
function rawState_(sh){
  var last = sh.getLastRow();
  if(last<2) return '';
  var vals = sh.getRange(2,1,last-1,2).getValues();
  vals.sort(function(a,b){ return a[0]-b[0]; });
  return vals.map(function(r){ return r[1]; }).join('');
}
function writeState_(state){
  var sh = sheet_(TABS.STATE);
  var prev = rawState_(sh);
  var s = JSON.stringify(state);
  if(prev && prev !== s){
    backupState_(prev);            // 覆寫前先留一份（最多 5 版，出事才有得救）
    // 筆數警告：帳目、換匯分開比。只比總數的話，「洗掉 20 筆舊的＋自己新增 25 筆」會變成總數增加，警告根本不會跳。
    var o = safeParse_(prev), lost = [];
    if(o){
      if((state.entries||[]).length < (o.entries||[]).length) lost.push('帳目 '+(o.entries||[]).length+' → '+(state.entries||[]).length);
      if((state.fxs||[]).length     < (o.fxs||[]).length)     lost.push('換匯 '+(o.fxs||[]).length+' → '+(state.fxs||[]).length);
    }
    if(lost.length) appendAudit_([{ aid:'sys'+Date.now(), t:new Date().toLocaleString('zh-TW',{hour12:false}),
      by:'（系統）', action:'⚠️ 筆數減少', what:'_state',
      detail:lost.join('、')+'；覆寫前的內容已備份在「_state 備份」分頁' }]);
  }
  putChunks_(sh, s, 'json（勿手動編輯，app 資料原檔）');
}
function safeParse_(str){ try{ return JSON.parse(str); }catch(e){ return null; } }
function countRecs_(str){ var o=safeParse_(str); return o ? (o.entries||[]).length + (o.fxs||[]).length : 0; }
// 備份：一列一版（時間／筆數／json 分塊），最新的排在最上面，只留 5 版
function backupState_(raw){
  var sh = sheet_(TABS.BAK);
  if(sh.getLastRow()===0){
    sh.getRange(1,1,1,3).setValues([['備份時間','筆數','json（自動備份，覆寫前留存，最多 5 版）']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var cells = [];
  for(var i=0;i<raw.length;i+=CHUNK) cells.push(raw.substr(i, CHUNK));
  var row = [new Date().toLocaleString('zh-TW',{hour12:false}), countRecs_(raw)].concat(cells);
  if(sh.getMaxColumns() < row.length) sh.insertColumnsAfter(sh.getMaxColumns(), row.length - sh.getMaxColumns());
  sh.insertRowBefore(2);
  sh.getRange(2,1,1,row.length).setValues([row]);
  var last = sh.getLastRow();
  if(last > 6) sh.deleteRows(7, last-6);
}
function readState_(){
  var s = rawState_(sheet_(TABS.STATE));
  if(!s) return null;              // 真的還沒有資料
  try{ return JSON.parse(s); }
  // 🚨 讀壞掉 ≠ 沒資料。回 null 會讓合併變成「雲端本來就空的」→ 整包覆蓋，等於前功盡棄。
  catch(e){ throw new Error('雲端 _state 讀取失敗（內容損毀），為避免覆蓋已停止寫入'); }
}

// ---------- 合併：舊資料與這次送上來的取聯集，同一筆(id)取較新的，刪除註記(dels)擋掉已刪的 ----------
function recT_(r){ return Number((r && (r.updated || r.created)) || 0); }
function mergeDels_(a, b){
  var m = {}, cut = Date.now() - 90*86400000;
  (a||[]).concat(b||[]).forEach(function(d){
    if(d && d.id && (!m[d.id] || Number(d.t) > Number(m[d.id].t))) m[d.id] = d;
  });
  return Object.keys(m).map(function(k){ return m[k]; }).filter(function(d){ return Number(d.t) >= cut; });
}
function mergeLists_(a, b, delMap){
  var m = {}, order = [];
  (a||[]).concat(b||[]).forEach(function(r){
    if(!r || !r.id) return;
    if(!m[r.id]){ order.push(r.id); m[r.id] = r; }
    else if(recT_(r) > recT_(m[r.id])) m[r.id] = r;
  });
  return order.map(function(id){ return m[id]; }).filter(function(r){
    var dt = delMap[r.id];
    return !(dt && dt >= recT_(r));
  });
}
function mergeStates_(old, inc){
  inc = inc || {};
  if(!old) return { fxs: inc.fxs||[], entries: inc.entries||[], dels: inc.dels||[] };
  var dels = mergeDels_(old.dels, inc.dels), dm = {};
  dels.forEach(function(d){ dm[d.id] = Number(d.t); });
  return {
    fxs:     mergeLists_(old.fxs,     inc.fxs,     dm),
    entries: mergeLists_(old.entries, inc.entries, dm),
    dels:    dels
  };
}

// ---------- 換匯台帳 / 記帳明細：整表重寫（app 端算好、這裡只存） ----------
// 🚨 這兩張表是 clearContents 後整表重寫，原本沒有任何備份：只要有人用「批次沒讀到」的裝置存一次檔，
// 自動運費列就會整批消失、已結帳月份的數字也會跟著變，而且沒有舊值可以比對。
// → 覆寫前先留一版備份，列數明顯變少時寫進修改軌跡。
function rewriteTable_(name, rows){
  var sh = sheet_(name);
  var before = sh.getLastRow();
  bakTable_(name, sh);
  var after = (rows && rows.length) ? rows.length + 1 : 0;   // +1 表頭
  if(before > 1 && after < before * 0.8){
    appendAudit_([{ aid:'sys'+Date.now(), t:new Date().toLocaleString('zh-TW',{hour12:false}),
      by:'（系統）', action:'⚠️ 會計表列數驟減', what:name,
      detail:'列數 '+before+' → '+after+'；覆寫前的內容已備份在「'+name+' 備份」分頁。'
             +'常見原因：存檔的那台裝置沒讀到 weight-app 批次，自動帶入的運費列整批不見。' }]);
  }
  sh.clearContents();
  if(!rows || !rows.length){ sh.getRange(1,1).setValue('（目前沒有資料）'); return; }
  var headers = Object.keys(rows[0]);
  var data = rows.map(function(r){ return headers.map(function(h){ var v=r[h]; return v==null?'':v; }); });
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.getRange(2,1,data.length,headers.length).setValues(data);
  sh.setFrozenRows(1);
}

// 覆寫前把整張表另存一版（只留最近一版；這兩張表是衍生報表，_state 才是真身）
function bakTable_(name, sh){
  try{
    var last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if(last < 2 || lastCol < 1) return;                 // 空表不用備份
    var bs = sheet_(name + ' 備份');
    bs.clearContents();
    var vals = sh.getRange(1, 1, last, lastCol).getValues();
    bs.getRange(1, 1, 1, 1)
      .setValue('（' + name + ' 覆寫前備份 ' + new Date().toLocaleString('zh-TW',{hour12:false}) + '，只保留最近一版）')
      .setFontWeight('bold');
    bs.getRange(2, 1, vals.length, lastCol).setValues(vals);
  }catch(e){ /* 備份失敗不能擋住正常寫入 */ }
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
  var last = sh.getLastRow();
  if(last===0){
    sh.getRange(1,1,1,6).setValues([['時間','操作者','動作','項目','內容','紀錄ID']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    last = 1;
  } else if(!sh.getRange(1,6).getValue()){
    sh.getRange(1,6).setValue('紀錄ID').setFontWeight('bold');
  }
  // 去重：同一筆軌跡若因回應遺失被重送，不再重複寫一列
  var seen = {};
  if(last > 1){
    var back = Math.min(last-1, 800);
    sh.getRange(last-back+1, 6, back, 1).getValues().forEach(function(r){ if(r[0]) seen[String(r[0])] = 1; });
  }
  var rows = [];
  audit.forEach(function(a){
    var aid = String(a.aid||'');
    if(aid && seen[aid]) return;
    if(aid) seen[aid] = 1;
    rows.push([a.t||'', a.by||'', a.action||'', a.what||'', a.detail||'', aid]);
  });
  if(!rows.length) return;
  sh.getRange(sh.getLastRow()+1,1,rows.length,6).setValues(rows);
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
