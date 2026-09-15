/* 表單頁與福委頁共用的資料層與小工具（需先載入 config.js） */

const TABLES = TABLE_CONFIG.map(c => ({ id: String(c.id), size: c.size, veg: !!c.veg }));
const TOTAL_SEATS = TABLES.reduce((s, t) => s + t.size, 0);
const VEG_TABLES = TABLES.filter(t => t.veg);
const tableById = id => TABLES.find(t => t.id === id);
const isVegTable = id => !!(id && tableById(id)?.veg);
const menuItem = key => MENU.find(m => m.key === key);
const isSkipped = (item, vals) =>
  !!item.skipIf && Object.keys(item.skipIf).every(k => vals[k] === item.skipIf[k]);

// bookings: 每人一筆 {code, tableId|null, pairId|'', name, empId, main, seafood, drink, dessert, ts}
//   有 pairId              = 兩人一組
//   沒 pairId、有 tableId  = 素桌（一人一位）
//   沒 pairId、沒 tableId  = 找不到搭檔、待福委配對
//   code 是同一次送出的登記編號，只在內部使用，不顯示給同仁
let bookings = [];
const isAwaitingPair = b => !b.pairId && !b.tableId;
// 桌號填了、但不在 TABLE_CONFIG 裡（舊桌號或福委手動打錯字）
const hasUnknownTable = b => !!b.tableId && !tableById(b.tableId);

/* ---------------- 儲存層：後端 or 單機 ---------------- */

const LOCAL_KEY = 'moon-seating-v2';

const sleep = ms => new Promise(r => setTimeout(r, ms));

// 每次「送出」產生一組，重送時沿用，後端靠它判斷是不是同一筆
function newRequestId() {
  try { if (crypto.randomUUID) return crypto.randomUUID(); } catch (e) {}
  return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 12);
}

// 呼叫後端。遇到「可安全重試」的錯誤（排隊中、逾時、網路斷線、Apps Script 暫時過載）
// 會自動退避重試；座位已滿、工號重複這類錯誤則直接回報。
//   opts.retries：最多再試幾次
//   opts.onRetry(第幾次重試, 錯誤)：重試前通知畫面
async function api(action, payload, opts = {}) {
  if (!API_URL) return localApi(action, payload);
  const retries = opts.retries ?? 2;
  const body = JSON.stringify({ action, ...payload });

  for (let attempt = 0; ; attempt++) {
    try {
      return await callBackend(body);
    } catch (e) {
      if (!e.retry || attempt >= retries) throw e;
      if (opts.onRetry) opts.onRetry(attempt + 1, e);
      // 指數退避 + 隨機抖動，避免大家同時重送又撞在一起
      await sleep(Math.min(8000, 1000 * 2 ** attempt) * (0.5 + Math.random()));
    }
  }
}

async function callBackend(body) {
  const fail = (message, retry) => Object.assign(new Error(message), { retry });
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 30000);
  try {
    let res;
    try {
      res = await fetch(API_URL, {
        method: 'POST',
        // text/plain 可避開 CORS preflight，Apps Script 才收得到
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body,
        signal: ctrl.signal,
      });
    } catch (e) {
      throw fail(ctrl.signal.aborted ? '伺服器回應逾時' : '網路連線失敗', true);
    }
    if (!res.ok) throw fail('伺服器回應 ' + res.status, res.status === 429 || res.status >= 500);

    let data;
    try {
      data = JSON.parse(await res.text());
    } catch (e) {
      // Apps Script 同時執行太多時會回傳 HTML 錯誤頁
      throw fail('伺服器暫時忙碌', true);
    }
    if (!data.ok) throw fail(data.error || '未知錯誤', !!data.retry);
    return data;
  } finally {
    clearTimeout(timer);
  }
}

let memoryRows = null;   // localStorage 被停用時（無痕模式等）的退路

function localApi(action, payload) {
  const read = () => {
    if (memoryRows) return memoryRows;
    try { return JSON.parse(localStorage.getItem(LOCAL_KEY) || '[]'); }
    catch (e) { memoryRows = []; return memoryRows; }
  };
  const write = rows => {
    if (memoryRows) { memoryRows = rows; return; }
    try { localStorage.setItem(LOCAL_KEY, JSON.stringify(rows)); }
    catch (e) { memoryRows = rows; }
  };
  let rows = read();

  if (action === 'state') return Promise.resolve({ ok: true, bookings: rows });

  if (action === 'book') {
    const { tableId, groups } = payload;
    const err = checkBooking(rows, tableId, groups);
    if (err) return Promise.reject(new Error(err));
    const code = newCode(rows.map(r => r.code));
    const ts = new Date().toISOString();
    groups.forEach((g, i) => {
      const pairId = g.members.length === 2 ? `${code}-${i + 1}` : '';
      g.members.forEach(m => rows.push({ code, tableId, pairId, ...m, ts }));
    });
    write(rows);
    return Promise.resolve({ ok: true, code, bookings: rows });
  }
  return Promise.reject(new Error('不支援的動作'));
}

// 和 Code.gs 的 book_() 規則相同（單機模式用）
function checkBooking(rows, tableId, groups) {
  const veg = isVegTable(tableId);
  for (const g of groups) {
    if (veg) {
      if (g.members.length !== 1) return '素桌請一人一位登記';
      Object.assign(g.members[0], VEG_FIXED);
    } else if (g.members.length === 1) {
      if (tableId) return '找不到搭檔的登記請選「找不到搭檔」';
    } else if (g.members.length !== 2) {
      return '每組必須是兩人';
    }
  }
  const people = groups.flatMap(g => g.members);
  const byEmp = new Map(rows.map(r => [r.empId, r]));
  for (const p of people) {
    const hit = byEmp.get(p.empId);
    if (hit) return `工號 ${hit.empId}（${hit.name}）已經登記過，如需修改請洽福委會`;
  }
  if (tableId) {
    const t = tableById(tableId);
    if (!t) return `沒有「${tableId}」這張桌子`;
    const left = t.size - rows.filter(r => r.tableId === tableId).length;
    if (people.length > left) {
      return `${tableId} 桌只剩 ${left} 個位子（這次要登記 ${people.length} 位），請重新選擇`;
    }
  }
  return '';
}

function newCode(existing) {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 去掉易混淆的 I O 0 1
  let code;
  do {
    code = Array.from({ length: 6 }, () =>
      alphabet[Math.floor(Math.random() * alphabet.length)]).join('');
  } while (existing.includes(code));
  return code;
}

/* ---------------- 讀取加速 ----------------
   Apps Script 每次讀取約 1.5～2 秒，閒置後第一次更久。
   做法：先顯示上次讀到的資料（快取），背景再抓最新的。
   送出登記時後端一定會重新檢查座位，所以快取稍舊也不會超賣。 */

const CACHE_KEY = 'moon-state-cache-v1';
let lastSyncAt = 0;
let syncing = null;

function setBookings(list) {
  bookings = list || [];
  if (!API_URL) return;   // 單機模式本來就是即時的，不用快取
  try { localStorage.setItem(CACHE_KEY, JSON.stringify(bookings)); } catch (e) {}
}

// 有快取就先載入，回傳是否成功
function loadCachedBookings() {
  if (!API_URL) return false;
  try {
    const cached = JSON.parse(localStorage.getItem(CACHE_KEY) || 'null');
    if (Array.isArray(cached)) { bookings = cached; return true; }
  } catch (e) {}
  return false;
}

// 定時更新：只在頁面正在被看的時候才打後端，減輕 Apps Script 同時執行的負擔
function autoRefresh(fn, intervalMs, canRun) {
  const ok = () => document.visibilityState === 'visible' && (!canRun || canRun());
  setInterval(() => { if (ok()) fn(); }, intervalMs);
  document.addEventListener('visibilitychange', () => {
    if (ok() && Date.now() - lastSyncAt > 15000) fn();
  });
}

// 取得最新登記資料，並把連線狀態顯示在 #netBanner。
// 同時間只會有一個請求在跑，重複呼叫會共用同一個結果。
//   fresh = true：略過後端快取，直接讀試算表（福委頁用）
function syncBookings(fresh) {
  if (!syncing) syncing = doSync(fresh).finally(() => { syncing = null; });
  return syncing;
}

async function doSync(fresh) {
  const banner = document.querySelector('#netBanner');
  try {
    const r = await api('state', fresh ? { fresh: true } : {}, { retries: 1 });
    setBookings(r.bookings);
    lastSyncAt = Date.now();
    if (banner) {
      banner.className = API_URL ? 'banner' : 'banner info';
      banner.innerHTML = API_URL ? '' :
        '📋 目前是<b>單機示範模式</b>：資料只存在這台電腦的瀏覽器，換一台裝置看不到。正式使用請先設定 <span class="code">config.js</span> 的 <span class="code">API_URL</span>（見 README）。';
    }
    return true;
  } catch (e) {
    // 後端掛掉時頁面仍照常畫出來，只是資料可能不是最新
    if (banner) {
      banner.className = 'banner err';
      banner.textContent = '⚠️ 無法取得最新登記狀態：' + e.message + '（畫面可能不是最新的，請稍後重新整理）';
    }
    return false;
  }
}

/* ---------------- 小工具 ---------------- */

const $ = sel => document.querySelector(sel);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;' }[c]));

const seatsUsed = id => bookings.filter(b => b.tableId === id).length;
const pairSlots = id => Math.floor((tableById(id).size - seatsUsed(id)) / 2);

// 把同一組的人合在一起
function pairsOf(rows) {
  const map = new Map();
  rows.forEach(r => {
    const k = r.pairId || 'solo:' + r.empId;
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(r);
  });
  return [...map.values()];
}

// 「王小明＆陳美玲」，withTags 時加上兩人共用的套餐標籤
function pairLabel(members, withTags) {
  const tags = withTags && members.length === 2
    ? MENU.filter(it => it.per === 'pair').map(it => members[0][it.key]).filter(Boolean)
    : [];
  return members.map(m => esc(m.name)).join('＆')
    + tags.map(t => `<span class="tag">${esc(t)}</span>`).join('');
}
