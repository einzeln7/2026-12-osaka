/**
 * 聖誕大阪小旅行 — Google 試算表 API（Apps Script）
 * 只會讀寫「行程-網頁」「Expenses-明細」「Expense-設定」三個分頁。
 * 第一次：上方函式選 setup → 執行 → 授權。之後部署成「網頁應用程式」。
 */
const TOKEN = 'f7nfS_5IDzfnQGKMWI2M4oWE';   // 要跟網頁裡的 token 一樣，不要外流
const SH_IT = '行程-網頁', SH_EX = 'Expenses-明細', SH_SET = 'Expense-設定';
const IT_COLS = ['ID','Day','排序','類型','時間','停留','標題','地點','地圖連結','說明','備註','參考連結','已訂位','緯度','經度','對象'];
const IT_KEYS = ['id','day','order','type','time','stay','title','place','mapUrl','desc','note','ref','booked','lat','lng','who'];
const EX_COLS = ['日期','項目','幣值','總花費','墊付人','攤分人','類型','備註','ID'];
const DEFAULT_RATE = 0.2051;

function setup() {
  const ss = SpreadsheetApp.getActive();
  sheet_(SH_IT, IT_COLS);
  const ex = sheet_(SH_EX, EX_COLS);
  let st = ss.getSheetByName(SH_SET);
  if (!st) { st = ss.insertSheet(SH_SET); st.getRange('A1:B1').setValues([['日幣匯率（1 JPY = ? NTD）', DEFAULT_RATE]]); }
  // 補 ID
  const n = ex.getLastRow();
  if (n > 1) {
    const ids = ex.getRange(2, 9, n - 1, 1).getValues();
    ids.forEach((r, i) => { if (!r[0]) ids[i][0] = newId_('ex'); });
    ex.getRange(2, 9, n - 1, 1).setValues(ids);
  }
}

function sheet_(name, cols) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(cols);
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@'); // 全部當文字，避免日期、ID 被自動轉換
  sh.setFrozenRows(1);
  return sh;
}
const newId_ = p => p + Utilities.getUuid().replace(/-/g, '').slice(0, 10);

function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(p);
}
function handle_(p) {
  try {
    if (p.token !== TOKEN) throw new Error('驗證碼不對');
    const write = p.action && p.action !== 'all';
    const lock = LockService.getScriptLock();
    if (write) lock.waitLock(20000);
    try {
      return json_(Object.assign({ ok: true }, route_(p)));
    } finally { if (write) lock.releaseLock(); }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function route_(p) {
  switch (p.action) {
    case 'all': return all_();
    case 'initItems': {
      const sh = sheet_(SH_IT, IT_COLS);
      if (sh.getLastRow() <= 1) (p.items || []).forEach((it, i) => sh.appendRow(itemRow_(Object.assign({ order: (i + 1) * 10 }, it))));
      const ex = sheet_(SH_EX, EX_COLS);
      if (ex.getLastRow() <= 1) (p.expenses || []).forEach(e => ex.appendRow([e.date, e.item, e.cur, String(e.amount), e.payer, (e.split || []).join(', '), e.kind, e.note || '', e.id || newId_('ex')]));
      return all_();
    }
    case 'saveItem': upsert_(sheet_(SH_IT, IT_COLS), 1, p.item.id, itemRow_(p.item)); return all_();
    case 'deleteItem': remove_(sheet_(SH_IT, IT_COLS), 1, p.id); return all_();
    case 'reorder': {
      const sh = sheet_(SH_IT, IT_COLS), data = sh.getDataRange().getValues();
      (p.list || []).forEach(r => {
        for (let i = 1; i < data.length; i++) if (String(data[i][0]) === r.id) { sh.getRange(i + 1, 2, 1, 2).setValues([[String(r.day), String(r.order)]]); }
      });
      return all_();
    }
    case 'saveExpense': {
      const e = p.expense;
      if (!e.id) e.id = newId_('ex');
      upsert_(sheet_(SH_EX, EX_COLS), 9, e.id, [e.date, e.item, e.cur, String(e.amount), e.payer, (e.split || []).join(', '), e.kind, e.note || '', e.id]);
      return all_();
    }
    case 'deleteExpense': remove_(sheet_(SH_EX, EX_COLS), 9, p.id); return all_();
    case 'setMeta': {
      const st = SpreadsheetApp.getActive().getSheetByName(SH_SET);
      st.getRange('A2:B3').setValues([['網站名字', String((p.meta || {}).title || '')], ['副標', String((p.meta || {}).sub || '')]]);
      return all_();
    }
    case 'setRate': SpreadsheetApp.getActive().getSheetByName(SH_SET).getRange('B1').setValue(Number(p.rate)); return all_();
    default: throw new Error('不認得的動作：' + p.action);
  }
}

function itemRow_(it) {
  return IT_KEYS.map(k => {
    const v = it[k];
    if (k === 'booked') return v ? '已訂位' : '';
    if (k === 'who') return Array.isArray(v) ? v.join(', ') : (v || '');
    return v == null ? '' : String(v);
  });
}
function upsert_(sh, idCol, id, row) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (String(data[i][idCol - 1]) === String(id)) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  sh.appendRow(row);
}
function remove_(sh, idCol, id) {
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][idCol - 1]) === String(id)) sh.deleteRow(i + 1);
}
function all_() {
  const ss = SpreadsheetApp.getActive();
  const it = sheet_(SH_IT, IT_COLS).getDataRange().getDisplayValues().slice(1).filter(r => r[0]);
  const ex = sheet_(SH_EX, EX_COLS).getDataRange().getDisplayValues().slice(1).filter(r => r[1] || r[3]);
  const st = ss.getSheetByName(SH_SET);
  const rate = st ? Number(st.getRange('B1').getValue()) || DEFAULT_RATE : DEFAULT_RATE;
  return {
    items: it.map(r => { const o = {}; IT_KEYS.forEach((k, i) => o[k] = r[i]); return o; }),
    expenses: ex.map(r => ({ date: r[0], item: r[1], cur: r[2] === 'JPY' || r[2] === 'KRW' ? 'JPY' : 'NTD', amount: Number(String(r[3]).replace(/,/g, '')) || 0,
      payer: r[4], split: String(r[5]).split(/[,，、]\s*/).filter(String), kind: r[6] === '還款' ? '還款' : '支出', note: r[7], id: r[8] })),
    rate: rate,
    meta: st ? { title: String(st.getRange('B2').getValue() || ''), sub: String(st.getRange('B3').getValue() || '') } : {},
  };
}
