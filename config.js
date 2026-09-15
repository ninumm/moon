/* =======================================================================
   設定檔 —— 表單頁（index.html）和福委頁（admin.html）共用
   ⚠️ Code.gs 開頭也有一份 TABLE_CONFIG 與 MENU，修改時要兩邊一起改
   ======================================================================= */

// 貼上 Google Apps Script 部署後拿到的網址（/exec 結尾）。
// 留空 = 單機示範模式（資料只存在這台電腦的瀏覽器，僅供試玩）。
const API_URL = 'https://script.google.com/macros/s/AKfycbxJPzxyNUuLEfW2DhlQB_2GglEUFmZmchVFl7ezMz295Es2CChkUTay8JgLWt0pQjOf/exec';

// 報名人數（福委頁用來對照登記進度）
const EXPECTED_TOTAL = 93;

// 桌次設定
const TABLE_CONFIG = [
  { prefix: '甲', size: 4, count: 8 },   // 8 張 4 人桌 = 32
  { prefix: '乙', size: 6, count: 9 },   // 9 張 6 人桌 = 54
  { prefix: '丙', size: 8, count: 1 },   // 1 張 8 人桌 =  8
];

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
