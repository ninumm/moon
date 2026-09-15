/**
 * 中秋餐會座位與餐點登記 —— Google Apps Script 後端
 *
 * 資料存在同一個 Google 試算表裡（= 你說的「Excel 表格紀錄」），
 * 用 LockService 確保兩組人同時搶同一桌時不會超賣。
 *
 * 部署方式見 README.md。
 */

const SHEET_NAME = '報名紀錄';

/* ============ 設定區：必須和 config.js 完全一致 ============ */

const TABLE_CONFIG = [
  { prefix: '甲', size: 4, count: 8 },
  { prefix: '乙', size: 6, count: 9 },
  { prefix: '丙', size: 8, count: 1 },
];

// per: 'pair' = 一組兩人共用一份；'person' = 每人各選
// skipIf: 符合條件時這一項不用選（依據的欄位要排在它前面）
const MENU = [
  { key: 'main', label: '主餐套餐', per: 'pair', options: ['牛豚', '全豚', '素'] },
  { key: 'seafood', label: '海鮮', per: 'pair', options: ['干貝/大蝦', '干貝/花枝'],
    skipIf: { main: '素' } },
  { key: 'drink', label: '飲品', per: 'person',
    options: ['紅茶', '可樂', '可爾必思', '青梅可爾必思', '海鹽奶霜紅茶', '烏龍茶', '美式咖啡'] },
  { key: 'dessert', label: '甜點', per: 'person', options: ['手工塔', '紅豆湯', '冰棒'] },
];

/* ========================================================== */

const BASE_HEADERS = ['登記時間', '預約碼', '桌號', '組別', '姓名', '工號'];
const HEADERS = BASE_HEADERS.concat(MENU.map(function (it) { return it.label; }));

function tableMap_() {
  const m = {};
  TABLE_CONFIG.forEach(function (c) {
    for (let i = 1; i <= c.count; i++) m[c.prefix + i] = c.size;
  });
  return m;
}

function isSkipped_(item, vals) {
  if (!item.skipIf) return false;
  return Object.keys(item.skipIf).every(function (k) { return vals[k] === item.skipIf[k]; });
}

function clean_(v) { return v == null ? '' : String(v).trim(); }

/* ------------------------- 入口 ------------------------- */

function doGet() {
  try {
    return json_({ ok: true, bookings: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
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
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  } else {
    // 欄位和程式不一致時直接報錯，避免資料寫進錯的欄位
    const current = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0].map(clean_);
    if (current.join('|') !== HEADERS.join('|')) {
      throw new Error('「' + SHEET_NAME + '」工作表的欄位和程式不一致，'
        + '請把舊的工作表改名後再試（系統會自動建立新的）');
    }
  }
  return sh;
}

function readAll_() {
  const sh = sheet_();
  if (sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
  const base = BASE_HEADERS.length;
  return rows
    .filter(function (r) { return clean_(r[4]); })   // 有姓名才算一筆
    .map(function (r) {
      const o = {
        ts: r[0] instanceof Date ? r[0].toISOString() : clean_(r[0]),
        code: clean_(r[1]),
        tableId: clean_(r[2]) || null,   // 空白 = 待福委排座位
        pairId: clean_(r[3]),            // 空白 = 找不到搭檔、待福委配對
        name: clean_(r[4]),
        empId: clean_(r[5]),
      };
      MENU.forEach(function (it, i) { o[it.key] = clean_(r[base + i]); });
      return o;
    });
}

/* ------------------------- 登記 ------------------------- */

function book_(body) {
  const tableId = clean_(body.tableId) || null;
  const groups = body.groups || [];

  if (!groups.length) throw new Error('沒有收到任何登記資料');
  if (groups.length > 6) throw new Error('一次最多登記 6 組');

  const sizes = tableMap_();
  if (tableId && !sizes[tableId]) throw new Error('沒有「' + tableId + '」這張桌子');

  // 先在鎖外把資料格式驗證完，縮短持鎖時間
  const people = [];
  groups.forEach(function (g, gi) {
    const ms = g.members || [];
    const where = '第 ' + (gi + 1) + ' 組';

    if (ms.length === 1) {
      if (tableId) throw new Error('找不到搭檔的登記請選「交由福委安排」');
    } else if (ms.length !== 2) {
      throw new Error(where + '必須剛好兩人');
    }

    ms.forEach(function (m) {
      m.name = clean_(m.name);
      m.empId = clean_(m.empId);
      if (!m.name) throw new Error(where + '有人沒填姓名');
      if (!m.empId) throw new Error(where + '「' + m.name + '」沒填工號');

      // MENU 依序處理，skipIf 依據的欄位已經先清理過
      MENU.forEach(function (it) {
        const v = clean_(m[it.key]);
        m[it.key] = v;
        if (isSkipped_(it, m)) { m[it.key] = ''; return; }
        if (it.options.indexOf(v) === -1) {
          throw new Error('「' + m.name + '」的「' + it.label + '」'
            + (v ? '選項無效：' + v + '（請重新整理頁面再試）' : '還沒選'));
        }
      });
      people.push(m);
    });

    if (ms.length === 2) {
      MENU.forEach(function (it) {
        if (it.per === 'pair' && ms[0][it.key] !== ms[1][it.key]) {
          throw new Error(where + '的「' + it.label + '」兩人必須相同');
        }
      });
    }
  });

  const seen = {};
  people.forEach(function (p) {
    if (seen[p.empId]) throw new Error('工號 ' + p.empId + ' 重複填寫了');
    seen[p.empId] = true;
  });

  // 關鍵：讀取→檢查→寫入都在鎖裡，避免兩組人同時送出造成超賣
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中，請三秒後再送出一次');

  try {
    const existing = readAll_();

    const byEmp = {};
    existing.forEach(function (r) { byEmp[r.empId] = r; });
    people.forEach(function (p) {
      const hit = byEmp[p.empId];
      if (hit) {
        throw new Error('工號 ' + hit.empId + '（' + hit.name + '）已登記過'
          + (hit.tableId ? ' ' + hit.tableId + ' 桌' : '（交由福委安排）')
          + '，請先用預約碼取消原登記');
      }
    });

    if (tableId) {
      let used = 0;
      existing.forEach(function (r) { if (r.tableId === tableId) used++; });
      const left = sizes[tableId] - used;
      if (people.length > left) {
        throw new Error(tableId + ' 桌只剩 ' + left + ' 個位子（這次要登記 '
          + people.length + ' 位），請重新選擇');
      }
    }

    const code = newCode_(existing.map(function (r) { return r.code; }));
    const now = new Date();
    const rows = [];
    groups.forEach(function (g, gi) {
      const pairId = g.members.length === 2 ? code + '-' + (gi + 1) : '';
      g.members.forEach(function (m) {
        rows.push([now, code, tableId || '', pairId, m.name, m.empId]
          .concat(MENU.map(function (it) { return m[it.key]; })));
      });
    });

    const sh = sheet_();
    // 工號可能是 00123 這種格式，先設成純文字避免前導 0 被吃掉
    const range = sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length);
    range.offset(0, 5, rows.length, 1).setNumberFormat('@');
    range.setValues(rows);
    SpreadsheetApp.flush();

    return { ok: true, code: code, bookings: readAll_() };
  } finally {
    lock.releaseLock();
  }
}

function cancel_(body) {
  const code = clean_(body.code).toUpperCase();
  if (!code) throw new Error('請提供預約碼');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中，請稍後再試');

  try {
    const sh = sheet_();
    if (sh.getLastRow() < 2) throw new Error('查無此預約碼');

    const codes = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
    const targets = [];
    codes.forEach(function (r, i) {
      if (clean_(r[0]).toUpperCase() === code) targets.push(i + 2); // 試算表列號
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

/** 在試算表選單加「🌕 餐會工具」，重新整理試算表後出現。 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌕 餐會工具')
    .addItem('產生座位與點餐總表', 'buildSummary')
    .addToUi();
}

function buildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '總表';
  let sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);

  const all = readAll_();
  const sizes = tableMap_();
  const COLS = 4;
  const out = [];
  const pad = function (r) { while (r.length < COLS) r.push(''); return r; };
  const boldRows = [];

  // 把同組的人合併
  const pairsOf = function (rows) {
    const map = {}, order = [];
    rows.forEach(function (r) {
      const k = r.pairId || 'solo:' + r.empId;
      if (!map[k]) { map[k] = []; order.push(k); }
      map[k].push(r);
    });
    return order.map(function (k) { return map[k]; });
  };
  const pairText = function (ms) {
    const tags = ms.length === 2
      ? MENU.filter(function (it) { return it.per === 'pair'; })
          .map(function (it) { return ms[0][it.key]; }).filter(String)
      : [];
    return ms.map(function (m) { return m.name; }).join('＆')
      + (tags.length ? '（' + tags.join('/') + '）' : '');
  };

  // 1. 各桌名單
  boldRows.push(out.length + 1);
  out.push(['桌號', '人數', '空位', '組別與套餐']);
  Object.keys(sizes).forEach(function (id) {
    const ppl = all.filter(function (r) { return r.tableId === id; });
    out.push([id, ppl.length + '/' + sizes[id], sizes[id] - ppl.length,
      pairsOf(ppl).map(pairText).join('\n')]);
  });
  const unseated = all.filter(function (r) { return !r.tableId && r.pairId; });
  const solos = all.filter(function (r) { return !r.pairId; });
  out.push(['待排座位', unseated.length, '', pairsOf(unseated).map(pairText).join('\n')]);
  out.push(['待配對', solos.length, '', solos.map(function (r) {
    return r.name + '（偏好 ' + MENU.filter(function (it) { return it.per === 'pair'; })
      .map(function (it) { return r[it.key]; }).filter(String).join('/') + '）';
  }).join('\n')]);

  // 2. 點餐統計
  out.push([]);
  boldRows.push(out.length + 1);
  out.push(['點餐統計', '選項', '數量', '備註']);
  const pairHeads = pairsOf(all.filter(function (r) { return r.pairId; }))
    .map(function (p) { return p[0]; });
  MENU.forEach(function (it) {
    const perPair = it.per === 'pair';
    const src = perPair ? pairHeads : all;
    it.options.forEach(function (o, i) {
      const n = src.filter(function (r) { return r[it.key] === o; }).length;
      out.push([i === 0 ? it.label : '', o, n + (perPair ? ' 組' : ' 份'),
        i === 0 && perPair && solos.length ? '另有 ' + solos.length + ' 人待配對未計入' : '']);
    });
  });

  sh.getRange(1, 1, out.length, COLS).setValues(out.map(pad));
  boldRows.forEach(function (r) { sh.getRange(r, 1, 1, COLS).setFontWeight('bold'); });
  sh.getRange(1, 4, out.length, 1).setWrap(true);
  sh.setColumnWidth(4, 420);
  sh.autoResizeColumns(1, 3);
}
