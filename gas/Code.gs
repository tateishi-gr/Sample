/**
 * 見積書 明細検索 Mock - サーバー側 (Google Apps Script)
 *
 * データはスプレッドシートの 2 シートで管理する。
 *   - 見積書: 1 見積 = 1 行（ヘッダー情報）
 *   - 明細  : 1 明細 = 1 行（見積ID で見積書に紐づく）
 *
 * スクリプトプロパティ（任意）
 *   SPREADSHEET_ID  : データ保存先。未設定ならコンテナバインドのスプレッドシートを使う
 *   DRIVE_FOLDER_ID : 設定すると登録時に元ファイルを Drive に保存する
 */

const SHEET_QUOTES = '見積書';
const SHEET_ITEMS = '明細';

const QUOTE_HEADERS = ['見積ID', '登録日時', '見積日', '取引先', '件名', '合計金額', '取込元', 'ファイル名', 'ファイルURL', 'メモ'];
const ITEM_HEADERS = ['明細ID', '見積ID', '行No', '品名', '仕様・型番', '数量', '単位', '単価', '金額', '備考'];

const SOURCE_TYPES = ['excel', 'pdf', 'handwritten', 'manual'];
const MAX_SEARCH_RESULTS = 500;

// ---------------------------------------------------------------------------
// Web app entry point
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('見積明細検索')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------

/** Run once from the editor: creates both sheets and fills sample data if empty. */
function setup() {
  const ss = getSpreadsheet_();
  const quotes = ensureSheet_(ss, SHEET_QUOTES, QUOTE_HEADERS);
  ensureSheet_(ss, SHEET_ITEMS, ITEM_HEADERS);
  if (quotes.getLastRow() <= 1) {
    SAMPLE_QUOTES.forEach(q => registerQuote(q));
  }
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#eef2f7');
  }
  return sheet;
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートが見つかりません。SPREADSHEET_ID を設定してください。');
  return ss;
}

// ---------------------------------------------------------------------------
// API (called from the client via google.script.run)
// ---------------------------------------------------------------------------

/** Initial data for the UI: vendor list and counts. */
function getInitData() {
  const quotes = readQuotes_();
  const vendors = Array.from(new Set(quotes.map(q => q.vendor).filter(Boolean))).sort();
  return {
    vendors: vendors,
    quoteCount: quotes.length,
    itemCount: Math.max(sheet_(SHEET_ITEMS).getLastRow() - 1, 0),
  };
}

/**
 * Registers one quote and its line items.
 * @param {{header: Object, items: Object[], file?: {name: string, mimeType: string, base64: string}}} payload
 * @return {{quoteId: string, itemCount: number}}
 */
function registerQuote(payload) {
  const header = payload.header || {};
  const items = (payload.items || []).filter(it => String(it.name || '').trim() !== '');
  if (!String(header.vendor || '').trim()) throw new Error('取引先は必須です。');
  if (items.length === 0) throw new Error('明細が 1 行もありません。');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const quoteId = 'Q' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' +
      Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    const fileUrl = payload.file ? saveFile_(payload.file, quoteId) : '';

    const rows = items.map((it, i) => {
      const qty = toNumber_(it.qty);
      const price = toNumber_(it.unitPrice);
      const amount = it.amount !== '' && it.amount != null ? toNumber_(it.amount) : qty * price;
      return [quoteId + '-' + (i + 1), quoteId, i + 1, String(it.name).trim(), it.spec || '', qty, it.unit || '', price, amount, it.note || ''];
    });
    const total = rows.reduce((sum, r) => sum + (Number(r[8]) || 0), 0);

    sheet_(SHEET_QUOTES).appendRow([
      quoteId,
      new Date(),
      header.quoteDate ? new Date(header.quoteDate) : '',
      String(header.vendor).trim(),
      header.subject || '',
      total,
      SOURCE_TYPES.indexOf(header.sourceType) >= 0 ? header.sourceType : 'manual',
      header.fileName || (payload.file && payload.file.name) || '',
      fileUrl,
      header.memo || '',
    ]);
    const itemSheet = sheet_(SHEET_ITEMS);
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, rows.length, ITEM_HEADERS.length).setValues(rows);

    return { quoteId: quoteId, itemCount: rows.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Searches line items joined with their quote header.
 * @param {{keyword?: string, vendor?: string, dateFrom?: string, dateTo?: string,
 *          priceMin?: number, priceMax?: number, sourceType?: string}} q
 */
function searchItems(q) {
  q = q || {};
  const quotes = {};
  readQuotes_().forEach(h => { quotes[h.quoteId] = h; });

  const terms = normalize_(q.keyword || '').split(/\s+/).filter(Boolean);
  const priceMin = q.priceMin === '' || q.priceMin == null ? null : Number(q.priceMin);
  const priceMax = q.priceMax === '' || q.priceMax == null ? null : Number(q.priceMax);

  const results = [];
  let total = 0;
  readItems_().forEach(it => {
    const h = quotes[it.quoteId];
    if (!h) return;
    if (q.vendor && h.vendor !== q.vendor) return;
    if (q.sourceType && h.sourceType !== q.sourceType) return;
    if (q.dateFrom && (!h.quoteDate || h.quoteDate < q.dateFrom)) return;
    if (q.dateTo && (!h.quoteDate || h.quoteDate > q.dateTo)) return;
    if (priceMin !== null && it.unitPrice < priceMin) return;
    if (priceMax !== null && it.unitPrice > priceMax) return;
    if (terms.length) {
      const haystack = normalize_([it.name, it.spec, it.note, h.vendor, h.subject].join(' '));
      if (!terms.every(t => haystack.indexOf(t) >= 0)) return;
    }
    total++;
    if (results.length < MAX_SEARCH_RESULTS) {
      results.push(Object.assign({}, it, {
        vendor: h.vendor, subject: h.subject, quoteDate: h.quoteDate, sourceType: h.sourceType,
      }));
    }
  });
  results.sort((a, b) => (b.quoteDate || '').localeCompare(a.quoteDate || ''));
  return { total: total, items: results };
}

/** All quote headers, newest first. */
function listQuotes() {
  return readQuotes_().sort((a, b) => b.registeredAt.localeCompare(a.registeredAt));
}

/** One quote with its line items. */
function getQuote(quoteId) {
  const header = readQuotes_().filter(h => h.quoteId === quoteId)[0];
  if (!header) throw new Error('見積が見つかりません: ' + quoteId);
  const items = readItems_().filter(it => it.quoteId === quoteId).sort((a, b) => a.lineNo - b.lineNo);
  return { header: header, items: items };
}

/** Deletes a quote and its line items. */
function deleteQuote(quoteId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    deleteRowsWhere_(sheet_(SHEET_ITEMS), 1, quoteId);
    deleteRowsWhere_(sheet_(SHEET_QUOTES), 0, quoteId);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function sheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error('シート「' + name + '」がありません。先に setup() を実行してください。');
  return sheet;
}

function readRows_(name, width) {
  const sheet = sheet_(name);
  const last = sheet.getLastRow();
  if (last <= 1) return [];
  return sheet.getRange(2, 1, last - 1, width).getValues();
}

function readQuotes_() {
  return readRows_(SHEET_QUOTES, QUOTE_HEADERS.length).filter(r => r[0]).map(r => ({
    quoteId: String(r[0]),
    registeredAt: formatDate_(r[1], 'yyyy-MM-dd HH:mm'),
    quoteDate: formatDate_(r[2], 'yyyy-MM-dd'),
    vendor: String(r[3]),
    subject: String(r[4]),
    total: Number(r[5]) || 0,
    sourceType: String(r[6]),
    fileName: String(r[7]),
    fileUrl: String(r[8]),
    memo: String(r[9]),
  }));
}

function readItems_() {
  return readRows_(SHEET_ITEMS, ITEM_HEADERS.length).filter(r => r[0]).map(r => ({
    itemId: String(r[0]),
    quoteId: String(r[1]),
    lineNo: Number(r[2]) || 0,
    name: String(r[3]),
    spec: String(r[4]),
    qty: Number(r[5]) || 0,
    unit: String(r[6]),
    unitPrice: Number(r[7]) || 0,
    amount: Number(r[8]) || 0,
    note: String(r[9]),
  }));
}

function deleteRowsWhere_(sheet, colIndex, value) {
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][colIndex]) === value) sheet.deleteRow(i + 1);
  }
}

function saveFile_(file, quoteId) {
  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  if (!folderId) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(file.base64), file.mimeType, quoteId + '_' + file.name);
  return DriveApp.getFolderById(folderId).createFile(blob).getUrl();
}

// google.script.run cannot return Date objects, so dates go to the client as strings.
function formatDate_(v, pattern) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', pattern);
  return v ? String(v) : '';
}

function toNumber_(v) {
  const n = Number(String(v == null ? '' : v).normalize('NFKC').replace(/[,¥￥円\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

// Full-width/half-width and case insensitive matching ("ＬＡＮ" == "lan").
function normalize_(s) {
  return String(s).normalize('NFKC').toLowerCase();
}

// ---------------------------------------------------------------------------
// Sample data (used by setup())
// ---------------------------------------------------------------------------

const SAMPLE_QUOTES = [
  {
    header: { quoteDate: '2026-07-03', vendor: '東和電機株式会社', subject: '本社3F 照明LED化工事', sourceType: 'excel', fileName: '見積_東和電機_20260703.xlsx' },
    items: [
      { name: 'LEDベースライト', spec: 'LDL40 昼白色 4000lm', qty: 48, unit: '台', unitPrice: 12800 },
      { name: '既設照明器具撤去', spec: '40W×2灯', qty: 48, unit: '台', unitPrice: 1500 },
      { name: '配線工事', spec: 'VVF1.6-3C', qty: 1, unit: '式', unitPrice: 85000 },
      { name: '産廃処分費', spec: '', qty: 1, unit: '式', unitPrice: 32000 },
    ],
  },
  {
    header: { quoteDate: '2026-07-18', vendor: 'ネットワークス山田', subject: '会議室 LAN 増設', sourceType: 'pdf', fileName: 'Q-2026-0718.pdf' },
    items: [
      { name: 'LANケーブル', spec: 'Cat6A 305m巻', qty: 2, unit: '巻', unitPrice: 38000 },
      { name: 'スイッチングハブ', spec: '24ポート PoE+ GS1900-24HP', qty: 1, unit: '台', unitPrice: 64500 },
      { name: '情報コンセント', spec: 'Cat6A 2口', qty: 12, unit: '個', unitPrice: 2400 },
      { name: '敷設・結線作業', spec: '', qty: 2, unit: '人日', unitPrice: 45000 },
    ],
  },
  {
    header: { quoteDate: '2026-08-05', vendor: '東和電機株式会社', subject: '倉庫 照明・コンセント増設', sourceType: 'pdf', fileName: '見積書_倉庫.pdf' },
    items: [
      { name: 'LED高天井照明', spec: '150W 水銀灯400W相当', qty: 16, unit: '台', unitPrice: 29800 },
      { name: 'コンセント増設', spec: '2口 接地付', qty: 10, unit: '箇所', unitPrice: 8500 },
      { name: '高所作業車', spec: '10m 1日', qty: 2, unit: '日', unitPrice: 38000 },
    ],
  },
  {
    header: { quoteDate: '2026-08-22', vendor: '丸山建材', subject: '事務所 内装補修', sourceType: 'handwritten', fileName: 'memo_0822.jpg', memo: '現地打合せ時の手書きメモより' },
    items: [
      { name: '石膏ボード', spec: '12.5mm 910×1820', qty: 20, unit: '枚', unitPrice: 980 },
      { name: 'クロス張替', spec: '量産品', qty: 45, unit: 'm2', unitPrice: 1200 },
      { name: 'タイルカーペット', spec: '500角 グレー', qty: 60, unit: '枚', unitPrice: 750 },
    ],
  },
  {
    header: { quoteDate: '2026-09-10', vendor: 'ネットワークス山田', subject: '無線LAN 更新', sourceType: 'excel', fileName: '無線LAN更新_見積.xlsx' },
    items: [
      { name: '無線アクセスポイント', spec: 'Wi-Fi 6E WAX630E', qty: 6, unit: '台', unitPrice: 52000 },
      { name: 'LANケーブル', spec: 'Cat6 20m', qty: 6, unit: '本', unitPrice: 2200 },
      { name: '設定作業', spec: 'SSID/VLAN設計含む', qty: 1, unit: '式', unitPrice: 60000 },
    ],
  },
];
