/* 表單頁與福委頁共用的資料層與小工具（需先載入 config.js） */

const TABLES = [];
TABLE_CONFIG.forEach(c => {
  for (let i = 1; i <= c.count; i++) TABLES.push({ id: c.prefix + i, size: c.size });
});
const TOTAL_SEATS = TABLES.reduce((s, t) => s + t.size, 0);
const tableById = id => TABLES.find(t => t.id === id);
const menuItem = key => MENU.find(m => m.key === key);
const isSkipped = (item, vals) =>
  !!item.skipIf && Object.keys(item.skipIf).every(k => vals[k] === item.skipIf[k]);

// bookings: 每人一筆 {code, tableId|null, pairId|'', name, empId, main, seafood, drink, dessert, ts}
//   pairId 空白 = 找不到搭檔、待福委配對
//   tableId 空白 = 待福委排座位
let bookings = [];

/* ---------------- 儲存層：後端 or 單機 ---------------- */

const LOCAL_KEY = 'moon-seating-v2';

async function api(action, payload) {
  if (!API_URL) return localApi(action, payload);
  const res = await fetch(API_URL, {
    method: 'POST',
    // text/plain 可避開 CORS preflight，Apps Script 才收得到
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...payload }),
  });
  if (!res.ok) throw new Error('伺服器回應 ' + res.status);
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || '未知錯誤');
  return data;
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

  if (action === 'cancel') {
    const code = String(payload.code || '').toUpperCase();
    const hit = rows.filter(r => r.code === code).length;
    if (!hit) return Promise.reject(new Error('查無此預約碼'));
    rows = rows.filter(r => r.code !== code);
    write(rows);
    return Promise.resolve({ ok: true, removed: hit, bookings: rows });
  }
  return Promise.reject(new Error('不支援的動作'));
}

// 和 Code.gs 的 book_() 規則相同（單機模式用）
function checkBooking(rows, tableId, groups) {
  for (const g of groups) {
    if (g.members.length === 1 && tableId) return '找不到搭檔的登記請選「交由福委安排」';
    if (g.members.length < 1 || g.members.length > 2) return '每組必須是兩人';
  }
  const people = groups.flatMap(g => g.members);
  const byEmp = new Map(rows.map(r => [r.empId, r]));
  for (const p of people) {
    const hit = byEmp.get(p.empId);
    if (hit) return `工號 ${hit.empId}（${hit.name}）已經登記過，請先用預約碼取消原登記`;
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
function syncBookings() {
  if (!syncing) syncing = doSync().finally(() => { syncing = null; });
  return syncing;
}

async function doSync() {
  const banner = document.querySelector('#netBanner');
  try {
    const r = await api('state', {});
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
