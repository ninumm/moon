/**
 * 中秋餐會座位與餐點登記 —— Google Apps Script 後端
 *
 * 資料存在同一個 Google 試算表裡（= 你說的「Excel 表格紀錄」）。
 * 同仁只能新增登記；修改或刪除一律由福委直接改試算表。
 *
 * 併發處理：
 *   1. 讀取→檢查→寫入全部在 Script Lock 裡，同一時間只有一筆登記在寫，不會超賣
 *   2. 鎖內只做必要的試算表呼叫（整張表只讀一次），縮短排隊時間
 *   3. 每次送出帶 requestId；前端逾時重送時回傳第一次的結果，不會重複登記
 *   4. 搶不到鎖時回傳 retry:true，前端自動退避重試
 *   5. 讀取座位狀態走短暫快取，減少佔用 Apps Script 同時執行的名額
 *
 * 部署方式見 README.md。
 */

const SHEET_NAME = '報名紀錄';

/* ============ 設定區：必須和 config.js 完全一致 ============ */

// 一桌一行：id = 桌號，size = 人數；veg: true = 素桌，一人一位登記，餐點自動帶入 VEG_FIXED
const TABLE_CONFIG = [
  { id: 'A1', size: 4 },
  { id: 'A2', size: 4 },
  { id: 'A3', size: 4 },
  { id: 'A4', size: 4 },
  { id: 'A5', size: 4 },
  { id: 'A6', size: 4 },
  { id: 'B1', size: 4 },
  { id: 'B2', size: 4 },
  { id: 'B3', size: 6, veg: true },   // 素桌
  { id: 'B4', size: 8 },
  { id: 'C1', size: 6 },
  { id: 'C2', size: 6 },
  { id: 'C3', size: 6 },
  { id: 'C4', size: 6 },
  { id: 'C5', size: 6 },
  { id: 'C6', size: 6 },
  { id: 'C7', size: 6 },
  { id: 'C8', size: 6 },
];   // 合計 94 席

const VEG_FIXED = { main: '素' };

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

/* ============ 併發參數 ============ */

const LOCK_WAIT_MS = 10000;          // 等鎖最多 10 秒，等不到就請前端稍後重試
const STATE_CACHE_SECONDS = 15;      // 座位狀態快取秒數（福委手動改試算表後，最多這麼久會反映到表單頁）
const REQUEST_CACHE_SECONDS = 21600; // 記住已處理過的 requestId 6 小時

/* ========================================================== */

// 「預約碼」欄是同一次送出的登記編號，只在內部使用（串起同一次送出、產生組別編號），不顯示給同仁
const BASE_HEADERS = ['登記時間', '預約碼', '桌號', '組別', '姓名', '工號'];
const HEADERS = BASE_HEADERS.concat(MENU.map(function (it) { return it.label; }));
const COL_CODE = 1;   // 預約碼在第幾欄（從 0 算）
const COL_EMP = 5;    // 工號
const STATE_KEY = 'state';
const STATE_VER_KEY = 'state-ver';

// 桌號 → 人數
function tableMap_() {
  const m = {};
  TABLE_CONFIG.forEach(function (c) { m[String(c.id)] = c.size; });
  return m;
}

function isVegTable_(id) {
  return TABLE_CONFIG.some(function (c) { return c.veg && String(c.id) === id; });
}

function isSkipped_(item, vals) {
  if (!item.skipIf) return false;
  return Object.keys(item.skipIf).every(function (k) { return vals[k] === item.skipIf[k]; });
}

function clean_(v) { return v == null ? '' : String(v).trim(); }

// retry = true 代表「沒有寫入任何資料，前端可以安全地再送一次」
function fail_(message, retry) {
  const e = new Error(message);
  e.retry = !!retry;
  return e;
}

/* ------------------------- 入口 ------------------------- */

function doGet() {
  try {
    return json_({ ok: true, bookings: getState_(false) });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err), retry: !!err.retry });
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
      case 'state':  return json_({ ok: true, bookings: getState_(!!body.fresh) });
      case 'book':   return json_(book_(body));
      default:       return json_({ ok: false, error: '不支援的動作：' + body.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err), retry: !!err.retry });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------- 資料存取 ------------------------- */

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 一次讀出整張表（只呼叫一次試算表），並確認欄位沒被改掉
function readValues_(sh) {
  const values = sh.getDataRange().getValues();
  const header = (values[0] || []).slice(0, HEADERS.length).map(clean_);
  if (header.join('|') !== HEADERS.join('|')) {
    throw fail_('「' + SHEET_NAME + '」工作表的欄位和程式不一致，'
      + '請把舊的工作表改名後再試（系統會自動建立新的）');
  }
  return values;
}

function toBooking_(r) {
  const o = {
    ts: r[0] instanceof Date ? r[0].toISOString() : clean_(r[0]),
    code: clean_(r[COL_CODE]),
    tableId: clean_(r[2]) || null,   // 空白 = 待福委排座位
    pairId: clean_(r[3]),            // 空白 = 找不到搭檔、待福委配對
    name: clean_(r[4]),
    empId: clean_(r[COL_EMP]),
  };
  MENU.forEach(function (it, i) { o[it.key] = clean_(r[BASE_HEADERS.length + i]); });
  return o;
}

// values 第 0 列是標題
function toBookings_(values) {
  return values.slice(1).map(toBooking_).filter(function (b) { return b.name; });
}

function readAll_() {
  return toBookings_(readValues_(sheet_()));
}

/* ------------------------- 座位狀態快取 ------------------------- */

function getState_(fresh) {
  const cache = CacheService.getScriptCache();
  if (!fresh) {
    const hit = cache.get(STATE_KEY);
    if (hit) return JSON.parse(hit);
  }
  // 讀取期間若剛好有人寫入（版本號變了），就不要用這份可能較舊的資料覆蓋快取
  const verBefore = cache.get(STATE_VER_KEY);
  const list = readAll_();
  if (cache.get(STATE_VER_KEY) === verBefore) putCache_(STATE_KEY, JSON.stringify(list), STATE_CACHE_SECONDS);
  return list;
}

// 寫入後呼叫（在鎖內），讓所有人馬上看到最新座位
function publishState_(list) {
  putCache_(STATE_KEY, JSON.stringify(list), STATE_CACHE_SECONDS);
  putCache_(STATE_VER_KEY, String(Date.now()) + Math.random(), REQUEST_CACHE_SECONDS);
}

function putCache_(key, value, seconds) {
  try {
    CacheService.getScriptCache().put(key, value, seconds);
  } catch (e) {
    // 超過快取大小上限（100KB）時就不快取，直接讀試算表即可
  }
}

/* ------------------------- 登記 ------------------------- */

function book_(body) {
  const tableId = clean_(body.tableId) || null;
  const groups = body.groups || [];
  const requestId = clean_(body.requestId).slice(0, 80);

  if (!groups.length) throw fail_('沒有收到任何登記資料');
  if (groups.length > 8) throw fail_('一次最多登記 8 組');

  const sizes = tableMap_();
  if (tableId && !sizes[tableId]) throw fail_('沒有「' + tableId + '」這張桌子');
  const vegTable = isVegTable_(tableId);

  // 先在鎖外把資料格式驗證完，縮短持鎖時間
  const people = [];
  groups.forEach(function (g, gi) {
    const ms = g.members || [];
    const where = vegTable ? '第 ' + (gi + 1) + ' 位' : '第 ' + (gi + 1) + ' 組';

    if (vegTable) {
      if (ms.length !== 1) throw fail_('素桌請一人一位登記');
    } else if (ms.length === 1) {
      if (tableId) throw fail_('找不到搭檔的登記請選「找不到搭檔」');
    } else if (ms.length !== 2) {
      throw fail_(where + '必須剛好兩人');
    }

    ms.forEach(function (m) {
      m.name = clean_(m.name);
      m.empId = clean_(m.empId);
      if (!m.name) throw fail_(where + '有人沒填姓名');
      if (!m.empId) throw fail_(where + '「' + m.name + '」沒填工號');
      // 素桌的餐點以後端為準，前端傳什麼都覆蓋掉
      if (vegTable) Object.keys(VEG_FIXED).forEach(function (k) { m[k] = VEG_FIXED[k]; });

      // MENU 依序處理，skipIf 依據的欄位已經先清理過
      MENU.forEach(function (it) {
        const v = clean_(m[it.key]);
        m[it.key] = v;
        if (isSkipped_(it, m)) { m[it.key] = ''; return; }
        if (it.options.indexOf(v) === -1) {
          throw fail_('「' + m.name + '」的「' + it.label + '」'
            + (v ? '選項無效：' + v + '（請重新整理頁面再試）' : '還沒選'));
        }
      });
      people.push(m);
    });

    if (ms.length === 2) {
      MENU.forEach(function (it) {
        if (it.per === 'pair' && ms[0][it.key] !== ms[1][it.key]) {
          throw fail_(where + '的「' + it.label + '」兩人必須相同');
        }
      });
    }
  });

  const seen = {};
  people.forEach(function (p) {
    if (seen[p.empId]) throw fail_('工號 ' + p.empId + ' 重複填寫了');
    seen[p.empId] = true;
  });

  // 關鍵：讀取→檢查→寫入都在鎖裡，避免兩組人同時送出造成超賣
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) throw fail_('目前登記的人很多，正在排隊中，請稍候', true);

  try {
    const cache = CacheService.getScriptCache();
    const sh = sheet_();
    const values = readValues_(sh);
    const existing = toBookings_(values);

    // 同一個 requestId 已經成功過（前端逾時後重送）→ 回傳原本的結果，不重複寫入
    if (requestId) {
      const prevCode = cache.get('req:' + requestId);
      if (prevCode && existing.some(function (r) { return r.code === prevCode; })) {
        return { ok: true, code: prevCode, bookings: existing, replay: true };
      }
    }

    const byEmp = {};
    existing.forEach(function (r) { byEmp[r.empId] = r; });
    people.forEach(function (p) {
      const hit = byEmp[p.empId];
      if (hit) {
        throw fail_('工號 ' + hit.empId + '（' + hit.name + '）已登記過'
          + (hit.tableId ? ' ' + hit.tableId + ' 桌' : '（交由福委安排）')
          + '，如需修改請洽福委會');
      }
    });

    if (tableId) {
      let used = 0;
      existing.forEach(function (r) { if (r.tableId === tableId) used++; });
      const left = sizes[tableId] - used;
      if (people.length > left) {
        throw fail_(tableId + ' 桌只剩 ' + left + ' 個位子（這次要登記 '
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

    const range = sh.getRange(values.length + 1, 1, rows.length, HEADERS.length);
    // 工號可能是 00123 這種格式，先設成純文字避免前導 0 被吃掉
    range.offset(0, COL_EMP, rows.length, 1).setNumberFormat('@');
    range.setValues(rows);
    SpreadsheetApp.flush();

    // 不再重讀試算表，直接組出最新狀態
    const list = existing.concat(rows.map(toBooking_));
    if (requestId) putCache_('req:' + requestId, code, REQUEST_CACHE_SECONDS);
    publishState_(list);

    return { ok: true, code: code, bookings: list };
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
    const veg = isVegTable_(id);
    out.push([veg ? id + '（素桌）' : id, ppl.length + '/' + sizes[id], sizes[id] - ppl.length,
      veg
        ? ppl.map(function (r) { return r.name + '（' + r.drink + '/' + r.dessert + '）'; }).join('\n')
        : pairsOf(ppl).map(pairText).join('\n')]);
  });
  const unseated = all.filter(function (r) { return !r.tableId && r.pairId; });
  const solos = all.filter(function (r) { return !r.pairId && !r.tableId; });
  const vegPeople = all.filter(function (r) { return isVegTable_(r.tableId); });
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
      // 共用項目：兩人一組算「組」，素桌一人一位算「份」
      const veg = perPair ? vegPeople.filter(function (r) { return r[it.key] === o; }).length : 0;
      const notes = [];
      if (veg) notes.push('另加素桌 ' + veg + ' 份');
      if (i === 0 && perPair && solos.length) notes.push('另有 ' + solos.length + ' 人待配對未計入');
      out.push([i === 0 ? it.label : '', o, n + (perPair ? ' 組' : ' 份'), notes.join('；')]);
    });
  });

  sh.getRange(1, 1, out.length, COLS).setValues(out.map(pad));
  boldRows.forEach(function (r) { sh.getRange(r, 1, 1, COLS).setFontWeight('bold'); });
  sh.getRange(1, 4, out.length, 1).setWrap(true);
  sh.setColumnWidth(4, 420);
  sh.autoResizeColumns(1, 3);
}
