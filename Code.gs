/**
 * 中秋餐會座位預約 —— Google Apps Script 後端
 *
 * 資料存在同一個 Google 試算表裡（= 你說的「Excel 表格紀錄」），
 * 用 LockService 確保兩個人同時搶同一桌時不會超賣。
 *
 * 部署方式見 README.md。
 */

const SHEET_NAME = '報名紀錄';
const HEADERS = ['登記時間', '預約碼', '桌號', '姓名', '工號', '餐別', '備註'];

// 必須和 index.html 的 TABLE_CONFIG 一致，否則容量會對不上。
const TABLE_CONFIG = [
  { prefix: '甲', size: 4, count: 7 },
  { prefix: '乙', size: 6, count: 9 },
  { prefix: '丙', size: 8, count: 1 },
];

function tableMap_() {
  const m = {};
  TABLE_CONFIG.forEach(function (c) {
    for (let i = 1; i <= c.count; i++) m[c.prefix + i] = c.size;
  });
  return m;
}

/* ------------------------- 入口 ------------------------- */

function doGet() {
  return json_({ ok: true, bookings: readAll_() });
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '請求格式錯誤' });
  }

  try {
    switch (body.action) {
      case 'state':  return json_({ ok: true, bookings: readAll_() });
      case 'book':   return json_(book_(body));
      case 'cancel': return json_(cancel_(body));
      default:       return json_({ ok: false, error: '不支援的動作：' + body.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------- 資料存取 ------------------------- */

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function readAll_() {
  const sh = sheet_();
  if (sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
  return rows
    .filter(function (r) { return r[3]; })   // 有姓名才算一筆
    .map(function (r) {
      return {
        ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        code: String(r[1]),
        tableId: r[2] ? String(r[2]) : null, // 空白 = 交由福委安排
        name: String(r[3]),
        empId: String(r[4]),
        meal: String(r[5]),
        note: String(r[6] || ''),
      };
    });
}

/* ------------------------- 預約 ------------------------- */

function book_(body) {
  const tableId = body.tableId ? String(body.tableId) : null;
  const people = body.people || [];

  if (!people.length) throw new Error('沒有收到任何人員資料');
  if (people.length > 12) throw new Error('一次最多登記 12 位');

  people.forEach(function (p) {
    if (!p.name || !String(p.name).trim())   throw new Error('有人沒填姓名');
    if (!p.empId || !String(p.empId).trim()) throw new Error('有人沒填工號');
  });

  const sizes = tableMap_();
  if (tableId && !sizes[tableId]) throw new Error('沒有「' + tableId + '」這張桌子');

  // 關鍵：整段讀取→驗證→寫入都在鎖裡，避免兩人同時送出造成超賣
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中，請三秒後再送出一次');

  try {
    const existing = readAll_();

    // 同一個工號只能登記一次
    const byEmp = {};
    existing.forEach(function (r) { byEmp[r.empId] = r; });
    for (let i = 0; i < people.length; i++) {
      const hit = byEmp[String(people[i].empId).trim()];
      if (hit) {
        throw new Error('工號 ' + hit.empId + '（' + hit.name + '）已登記過'
          + (hit.tableId ? ' ' + hit.tableId + ' 桌' : '（交由福委安排）')
          + '，請先用預約碼取消原登記');
      }
    }

    // 容量檢查
    if (tableId) {
      let used = 0;
      existing.forEach(function (r) { if (r.tableId === tableId) used++; });
      const left = sizes[tableId] - used;
      if (people.length > left) {
        throw new Error(tableId + ' 桌只剩 ' + left + ' 個位子（你要登記 '
          + people.length + ' 位），請重新選擇');
      }
    }

    const code = newCode_(existing.map(function (r) { return r.code; }));
    const now = new Date();
    const rows = people.map(function (p) {
      return [now, code, tableId || '', String(p.name).trim(),
              String(p.empId).trim(), p.meal || 'meat', p.note || ''];
    });

    const sh = sheet_();
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();

    return { ok: true, code: code, bookings: readAll_() };
  } finally {
    lock.releaseLock();
  }
}

function cancel_(body) {
  const code = String(body.code || '').trim().toUpperCase();
  if (!code) throw new Error('請提供預約碼');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中，請稍後再試');

  try {
    const sh = sheet_();
    if (sh.getLastRow() < 2) throw new Error('查無此預約碼');

    const values = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
    const targets = [];
    values.forEach(function (r, i) {
      if (String(r[1]).trim().toUpperCase() === code) targets.push(i + 2); // 試算表列號
    });
    if (!targets.length) throw new Error('查無此預約碼');

    // 由下往上刪，列號才不會位移
    targets.reverse().forEach(function (rowNum) { sh.deleteRow(rowNum); });
    SpreadsheetApp.flush();

    return { ok: true, removed: targets.length, bookings: readAll_() };
  } finally {
    lock.releaseLock();
  }
}

function newCode_(existing) {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 去掉易混淆的 I O 0 1
  let code;
  do {
    code = '';
    for (let i = 0; i < 6; i++) {
      code += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
    }
  } while (existing.indexOf(code) !== -1);
  return code;
}

/* ------------------------- 福委工具 ------------------------- */

/**
 * 在試算表選單加一個「餐會工具」，可一鍵產生目前的座位總表。
 * 重新整理試算表後就會出現。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌕 餐會工具')
    .addItem('產生座位總表', 'buildSummary')
    .addToUi();
}

function buildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '座位總表';
  let sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);

  const all = readAll_();
  const sizes = tableMap_();
  const out = [['桌號', '容量', '已登記', '空位', '葷食', '素食', '其他', '名單']];

  Object.keys(sizes).forEach(function (id) {
    const ppl = all.filter(function (r) { return r.tableId === id; });
    const cnt = function (m) { return ppl.filter(function (r) { return r.meal === m; }).length; };
    out.push([id, sizes[id], ppl.length, sizes[id] - ppl.length,
      cnt('meat'), cnt('veg'), cnt('other'),
      ppl.map(function (r) { return r.name; }).join('、')]);
  });

  const pool = all.filter(function (r) { return !r.tableId; });
  const cntP = function (m) { return pool.filter(function (r) { return r.meal === m; }).length; };
  out.push(['待福委安排', '', pool.length, '', cntP('meat'), cntP('veg'), cntP('other'),
    pool.map(function (r) { return r.name; }).join('、')]);

  const totalSeats = Object.keys(sizes).reduce(function (s, k) { return s + sizes[k]; }, 0);
  out.push([]);
  out.push(['總計', totalSeats, all.length, totalSeats - all.filter(function (r) { return r.tableId; }).length,
    all.filter(function (r) { return r.meal === 'meat'; }).length,
    all.filter(function (r) { return r.meal === 'veg'; }).length,
    all.filter(function (r) { return r.meal === 'other'; }).length, '']);

  sh.getRange(1, 1, out.length, 8).setValues(
    out.map(function (r) { while (r.length < 8) r.push(''); return r; })
  );
  sh.getRange(1, 1, 1, 8).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 7);
}
