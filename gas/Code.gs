/**
 * 金禾鍋貼 · 線上訂單 → Google 試算表
 * 容器綁定腳本：在試算表按「擴充功能 → Apps Script」貼上本檔，先執行一次 setup()。
 * 部署：部署 → 新增部署作業 → 網頁應用程式（執行身分：我；存取權：所有人），把網址貼到網站 index.html 的 ORDER_ENDPOINT。
 */

var SHEET = { ORDERS: 'Orders', ITEMS: 'OrderItems', MENU: 'Menu', CONFIG: 'Config' };
var STATUSES = ['新訂單', '製作中', '可取餐', '已完成', '已取消'];
var ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '姓名', '電話', '地址', '品項', '總額', '備註', '狀態更新時間'];
var ITEM_HEADERS = ['訂單編號', '建立時間', '商品ID', '商品名稱', '單價', '數量', '小計'];
var MENU_HEADERS = ['商品ID', '商品名稱', '單價', '供應中'];
var MENU_SEED = [
  ['p1', '金禾招牌鍋貼 10 入', 80, true], ['p2', '香辣韭菜鍋貼 10 入', 90, true], ['p3', '鮮蝦鍋貼 10 入', 110, true],
  ['p4', '高麗菜水餃 10 入', 80, true], ['p5', '韭菜水餃 10 入', 85, true], ['p6', '暖心酸辣湯', 45, true],
  ['p7', '古早味紅茶', 30, true], ['p8', '爽脆小黃瓜', 40, true], ['p9', '胡麻時蔬', 50, true], ['gift', '招財鍋貼禮盒 50 入', 399, true]
];
var CONFIG_SEED = [
  ['NOTIFY_EMAIL', '', '新訂單與每日彙整寄送的 Email（留空則不寄）'],
  ['ORDERING_OPEN', 'TRUE', 'TRUE 接單／FALSE 暫停接單'],
  ['DAILY_SUMMARY_HOUR', '22', '每日彙整寄送時間（0–23 點，改後請重新執行 installTriggers）']
];
var MAX_LINES = 30, MAX_QTY = 99;

/* ───────── 一次性設定 ───────── */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orders = ensureSheet_(ss, SHEET.ORDERS, ORDER_HEADERS);
  ensureSheet_(ss, SHEET.ITEMS, ITEM_HEADERS);
  var menu = ensureSheet_(ss, SHEET.MENU, MENU_HEADERS);
  var cfg = ensureSheet_(ss, SHEET.CONFIG, ['設定', '值', '說明']);
  if (menu.getLastRow() < 2) menu.getRange(2, 1, MENU_SEED.length, 4).setValues(MENU_SEED);
  if (cfg.getLastRow() < 2) cfg.getRange(2, 1, CONFIG_SEED.length, 3).setValues(CONFIG_SEED);
  menu.getRange('D2:D500').insertCheckboxes();
  // 狀態下拉選單 + 顏色
  orders.getRange('C2:C5000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  orders.getRange('F2:F5000').setNumberFormat('@');
  var colors = { '新訂單': '#fde2c8', '製作中': '#fff2b3', '可取餐': '#cfe8d5', '已完成': '#e3e3e3', '已取消': '#f4c7c3' };
  orders.setConditionalFormatRules(STATUSES.map(function (s) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s]).setRanges([orders.getRange('C2:C5000')]).build();
  }));
  orders.setFrozenRows(1); orders.setColumnWidth(8, 360); orders.setColumnWidth(7, 220);
  installTriggers();
  SpreadsheetApp.getUi().alert('設定完成！請到 Config 分頁填入 NOTIFY_EMAIL，再依說明部署為網頁應用程式。');
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#39271f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  var hour = parseInt(getConfig_('DAILY_SUMMARY_HOUR', '22'), 10);
  ScriptApp.newTrigger('sendDailySummary').timeBased().everyDays(1).atHour(isNaN(hour) ? 22 : hour).create();
  ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

/* ───────── 網頁應用程式 ───────── */
function doGet() {
  return json_({ ok: true, service: 'jinhe-orders', open: isOpen_() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var cache = CacheService.getScriptCache();
    var rid = String(p.requestId || '').slice(0, 64);
    if (rid) { var hit = cache.get('rid_' + rid); if (hit) return textOut_(hit); }   // 重複送出 → 回傳同一筆結果
    if (p.website) return json_({ ok: false, error: '無法處理此訂單' });               // honeypot
    if (!isOpen_()) return json_({ ok: false, error: '目前暫停接單，請稍後再試' });

    var order = buildOrder_(p, readMenu_());
    if (order.error) return json_({ ok: false, error: order.error });
    if (cache.get('ph_' + order.phone)) return json_({ ok: false, error: '送出太頻繁，請 30 秒後再試' });

    lock.waitLock(20000);
    var id = nextOrderId_();
    var now = new Date();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheetByName(SHEET.ORDERS).appendRow([id, now, STATUSES[0], order.method, safe_(order.name), order.phone, safe_(order.address),
      safe_(order.lines.map(function (l) { return l.name + ' × ' + l.qty; }).join('、')), order.total, safe_(order.note), now]);
    var items = ss.getSheetByName(SHEET.ITEMS);
    order.lines.forEach(function (l) { items.appendRow([id, now, l.id, l.name, l.price, l.qty, l.price * l.qty]); });
    SpreadsheetApp.flush();
    lock.releaseLock();

    var result = JSON.stringify({ ok: true, orderId: id, total: order.total });
    if (rid) cache.put('rid_' + rid, result, 600);
    cache.put('ph_' + order.phone, '1', 30);
    notifyNewOrder_(id, order);
    return textOut_(result);
  } catch (err) {
    try { lock.releaseLock(); } catch (x) {}
    console.error(err && err.stack || err);
    return json_({ ok: false, error: '系統忙碌，請稍後再試或來電訂購' });
  }
}

/* 純函式：驗證並以「試算表價格」重新計算金額（可在 Node 測試） */
function buildOrder_(p, menu) {
  var byId = {}; menu.forEach(function (m) { byId[m.id] = m; });
  var raw = Array.isArray(p.items) ? p.items : [];
  if (!raw.length || raw.length > MAX_LINES) return { error: '購物車內容不正確' };
  var seen = {}, lines = [];
  for (var i = 0; i < raw.length; i++) {
    var id = String(raw[i].id), qty = Number(raw[i].qty), m = byId[id];
    if (!m || seen[id] || !(qty >= 1 && qty <= MAX_QTY) || Math.floor(qty) !== qty) return { error: '購物車內容不正確' };
    if (!m.available) return { error: '「' + m.name + '」目前暫停供應，請移除後再送出' };
    seen[id] = 1; lines.push({ id: id, name: m.name, price: m.price, qty: qty });
  }
  var onlyGift = lines.every(function (l) { return l.id === 'gift'; });
  var hasGift = lines.some(function (l) { return l.id === 'gift'; });
  var allowed = onlyGift ? ['外帶', '宅配'] : hasGift ? ['外帶'] : ['內用', '外帶'];
  var method = String(p.method || '');
  if (allowed.indexOf(method) < 0) return { error: '取餐方式不適用於這些餐點' };
  var name = String(p.name || '').trim(), phone = String(p.phone || '').trim(), address = String(p.address || '').trim();
  if (!name || name.length > 50) return { error: '請填寫聯絡姓名' };
  if (!/^[0-9+() -]{8,20}$/.test(phone)) return { error: '聯絡電話格式不正確' };
  if (method === '宅配' && (!address || address.length > 200)) return { error: '請填寫宅配地址' };
  var total = lines.reduce(function (a, l) { return a + l.price * l.qty; }, 0);
  return { lines: lines, method: method, name: name, phone: phone, address: method === '宅配' ? address : '', note: String(p.note || '').slice(0, 200), total: total };
}

/* ───────── 輔助 ───────── */
function readMenu_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.MENU);
  var rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 4).getValues();
  return rows.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { id: String(r[0]), name: String(r[1]), price: Number(r[2]), available: r[3] === true || String(r[3]).toUpperCase() === 'TRUE' };
  });
}
function getConfig_(key, dflt) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CONFIG); if (!sh) return dflt;
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (rows[i][0] === key) return String(rows[i][1] === '' ? dflt : rows[i][1]);
  return dflt;
}
function isOpen_() { return getConfig_('ORDERING_OPEN', 'TRUE').toUpperCase() !== 'FALSE'; }
function nextOrderId_() {
  var props = PropertiesService.getScriptProperties();
  var day = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd'), key = 'seq_' + day;
  var n = parseInt(props.getProperty(key) || '0', 10) + 1;
  props.setProperty(key, String(n));
  return 'JH-' + day + '-' + ('000' + n).slice(-3);
}
function safe_(v) { v = String(v == null ? '' : v); return /^[=+\-@\t\r]/.test(v) ? "'" + v : v; }   // 防試算表公式注入
function json_(o) { return textOut_(JSON.stringify(o)); }
function textOut_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON); }

function notifyNewOrder_(id, o) {
  var to = getConfig_('NOTIFY_EMAIL', ''); if (!to) return;
  var body = ['訂單 ' + id, '取餐方式：' + o.method, '姓名：' + o.name, '電話：' + o.phone, o.address ? '地址：' + o.address : '', '',
    o.lines.map(function (l) { return l.name + ' × ' + l.qty + '　NT$ ' + l.price * l.qty; }).join('\n'), '', '總額：NT$ ' + o.total,
    o.note ? '備註：' + o.note : '', '', '試算表：' + SpreadsheetApp.getActiveSpreadsheet().getUrl()].filter(function (x) { return x !== ''; }).join('\n');
  try { MailApp.sendEmail(to, '【金禾鍋貼】新訂單 ' + id + '（NT$ ' + o.total + '）', body); } catch (e) { console.error(e); }
}

/* 狀態欄被修改 → 自動記錄更新時間 */
function onStatusEdit(e) {
  var r = e && e.range; if (!r) return;
  var sh = r.getSheet(); if (sh.getName() !== SHEET.ORDERS || r.getColumn() !== 3 || r.getRow() < 2) return;
  sh.getRange(r.getRow(), 11).setValue(new Date());
}

/* 每日彙整：今日訂單數、營業額（不含已取消）、熱銷品項 */
function sendDailySummary() {
  var to = getConfig_('NOTIFY_EMAIL', ''); if (!to) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd');
  var orders = ss.getSheetByName(SHEET.ORDERS).getDataRange().getValues().slice(1).filter(function (r) { return String(r[0]).indexOf('JH-' + today) === 0; });
  var valid = orders.filter(function (r) { return r[2] !== '已取消'; });
  var ids = {}; valid.forEach(function (r) { ids[r[0]] = 1; });
  var tally = {};
  ss.getSheetByName(SHEET.ITEMS).getDataRange().getValues().slice(1).forEach(function (r) { if (ids[r[0]]) tally[r[3]] = (tally[r[3]] || 0) + Number(r[5]); });
  var top = Object.keys(tally).sort(function (a, b) { return tally[b] - tally[a]; }).slice(0, 5).map(function (k) { return k + ' × ' + tally[k]; });
  var revenue = valid.reduce(function (a, r) { return a + Number(r[8]); }, 0);
  var open = valid.filter(function (r) { return r[2] === '新訂單' || r[2] === '製作中'; }).length;
  MailApp.sendEmail(to, '【金禾鍋貼】' + today + ' 每日彙整', ['今日訂單：' + valid.length + ' 筆（取消 ' + (orders.length - valid.length) + ' 筆）', '營業額：NT$ ' + revenue,
    '尚未完成：' + open + ' 筆', '', '熱銷品項：', top.join('\n') || '（無）'].join('\n'));
}
