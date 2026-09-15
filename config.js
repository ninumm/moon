/* =======================================================================
   設定檔 —— 表單頁（index.html）和福委頁（admin.html）共用
   ⚠️ Code.gs 開頭也有一份 TABLE_CONFIG 與 MENU，修改時要兩邊一起改
   ======================================================================= */

// 貼上 Google Apps Script 部署後拿到的網址（/exec 結尾）。
// 留空 = 單機示範模式（資料只存在這台電腦的瀏覽器，僅供試玩）。
const API_URL = 'https://script.google.com/macros/s/AKfycbxJPzxyNUuLEfW2DhlQB_2GglEUFmZmchVFl7ezMz295Es2CChkUTay8JgLWt0pQjOf/exec';

// 報名人數（福委頁用來對照登記進度）
const EXPECTED_TOTAL = 93;

// 桌次設定：一桌一行，id = 桌號，size = 人數（對應餐廳的「A1 4P」）
//   veg: true = 素桌。一人一位登記（不用兩人一組），餐點自動帶入 VEG_FIXED，只選飲品和甜點
//   頁面上的桌子依這裡的順序排列
// ⚠️ 已經有人登記的桌號不要改名或刪除，否則那些人會從頁面上消失（資料還在試算表裡）
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

// 素桌自動帶入的餐點（這些項目不會出現在素桌的表單上）
const VEG_FIXED = { main: '素' };

// 餐點選項
//   per:    'pair' = 一組兩人共用一份；'person' = 每人各選一份
//   step:   顯示在表單第幾步（第 1 步固定是填成員）
//   skipIf: 符合條件時這一項不用選（依據的欄位要排在它前面）
const MENU = [
  { key: 'main', label: '主餐套餐', per: 'pair', step: 2,
    options: ['牛豚', '全豚', '素'] },
  { key: 'seafood', label: '海鮮', per: 'pair', step: 3,
    options: ['干貝/大蝦', '干貝/花枝'],
    skipIf: { main: '素' }, skipNote: '素食套餐不含海鮮，不用選' },
  { key: 'drink', label: '飲品', per: 'person', step: 3,
    options: ['紅茶', '可樂', '可爾必思', '青梅可爾必思', '海鹽奶霜紅茶', '烏龍茶', '美式咖啡'] },
  { key: 'dessert', label: '甜點', per: 'person', step: 3,
    options: ['手工塔', '紅豆湯', '冰棒'] },
];
const STEP_TITLES = { 2: '選擇主餐套餐', 3: '選擇細項（海鮮・飲品・甜點）' };

// 不指定桌次時，一次最多登記幾組
const MAX_UNSEATED_GROUPS = 4;
