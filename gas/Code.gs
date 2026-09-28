/**
 * 見積書 明細検索 Mock - サーバー側 (Google Apps Script)
 *
 * データはスプレッドシートの 2 シートで管理する。
 *   - 見積書: 1 見積 = 1 行（ヘッダー情報）
 *   - 明細  : 1 明細 = 1 行（見積ID で見積書に紐づく）
 *
 * 列は見出し名で読み書きするので、シート上で列を並べ替えたり独自の列を足したりしても動く。
 *
 * スクリプトプロパティ（任意）
 *   SPREADSHEET_ID  : データ保存先。未設定ならコンテナバインドのスプレッドシートを使う
 *   DRIVE_FOLDER_ID : 設定すると登録時に元ファイルを Drive に保存する
 */

const SHEET_QUOTES = '見積書';
const SHEET_ITEMS = '明細';

// [key, 見出し]
const QUOTE_COLUMNS = [
  ['quoteId', '見積ID'],
  ['registeredAt', '登録日時'],
  ['updatedAt', '更新日時'],
  ['status', 'ステータス'],
  ['quoteNo', '見積番号'],
  ['quoteDate', '見積日'],
  ['validUntil', '有効期限'],
  ['vendor', '取引先'],
  ['vendorContact', '取引先担当者'],
  ['subject', '件名'],
  ['deliveryDate', '納期'],
  ['deliveryPlace', '納入場所'],
  ['paymentTerms', '支払条件'],
  ['staff', '自社担当者'],
  ['subtotal', '小計（税抜）'],
  ['tax', '消費税'],
  ['total', '合計（税込）'],
  ['sourceType', '取込元'],
  ['fileName', 'ファイル名'],
  ['fileUrl', 'ファイルURL'],
  ['memo', 'メモ'],
];
const ITEM_COLUMNS = [
  ['itemId', '明細ID'],
  ['quoteId', '見積ID'],
  ['lineNo', '行No'],
  ['name', '品名'],
  ['spec', '仕様・型番'],
  ['qty', '数量'],
  ['unit', '単位'],
  ['unitPrice', '単価'],
  ['amount', '金額'],
  ['note', '備考'],
];
// Old header names from earlier versions, renamed by setup().
const COLUMN_ALIASES = { '合計金額': '合計（税込）' };

// Header fields the user can edit from the UI.
const EDITABLE_KEYS = ['status', 'quoteNo', 'quoteDate', 'validUntil', 'vendor', 'vendorContact', 'subject',
  'deliveryDate', 'deliveryPlace', 'paymentTerms', 'staff', 'memo'];
const DATE_KEYS = ['quoteDate', 'validUntil'];
const DATETIME_KEYS = ['registeredAt', 'updatedAt'];
const NUMBER_KEYS = ['subtotal', 'tax', 'total', 'lineNo', 'qty', 'unitPrice', 'amount'];

const STATUSES = ['検討中', '採用', '不採用', '保留'];
const SOURCE_TYPES = ['excel', 'pdf', 'handwritten', 'manual'];
const TAX_RATE = 0.1;
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

/**
 * Run from the editor. Creates both sheets (or adds missing columns to existing ones)
 * and fills sample data if there is no quote yet.
 */
function setup() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, SHEET_QUOTES, QUOTE_COLUMNS);
  ensureSheet_(ss, SHEET_ITEMS, ITEM_COLUMNS);
  if (ss.getSheetByName(SHEET_QUOTES).getLastRow() <= 1) {
    SAMPLE_QUOTES.forEach(q => registerQuote(q));
  }
}

function ensureSheet_(ss, name, columns) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const labels = columns.map(c => c[1]);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(labels);
    sheet.setFrozenRows(1);
  } else {
    const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const current = range.getValues()[0].map(h => COLUMN_ALIASES[h] || String(h));
    range.setValues([current]);
    const missing = labels.filter(l => current.indexOf(l) < 0);
    if (missing.length) sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
  }
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#eef2f7');
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

/** Initial data for the UI: choices for selects and counts. */
function getInitData() {
  const quotes = readTable_(openTable_(SHEET_QUOTES, QUOTE_COLUMNS));
  const uniq = key => Array.from(new Set(quotes.map(q => q[key]).filter(Boolean))).sort();
  return {
    vendors: uniq('vendor'),
    staff: uniq('staff'),
    statuses: STATUSES,
    quoteCount: quotes.length,
    itemCount: Math.max(sheet_(SHEET_ITEMS).getLastRow() - 1, 0),
  };
}

/**
 * Registers one quote (header + line items).
 * @param {{header: Object, items: Object[], file?: {name: string, mimeType: string, base64: string}}} payload
 * @return {{quoteId: string, itemCount: number}}
 */
function registerQuote(payload) {
  const built = buildQuote_(payload);
  return withLock_(() => {
    const quoteId = 'Q' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + '-' +
      Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    const h = payload.header || {};
    const now = new Date();
    const header = Object.assign(built.header, {
      quoteId: quoteId,
      registeredAt: now,
      updatedAt: now,
      sourceType: SOURCE_TYPES.indexOf(h.sourceType) >= 0 ? h.sourceType : 'manual',
      fileName: h.fileName || (payload.file && payload.file.name) || '',
      fileUrl: payload.file ? saveFile_(payload.file, quoteId) : '',
    });
    const quotes = openTable_(SHEET_QUOTES, QUOTE_COLUMNS);
    quotes.sheet.appendRow(toRow_(quotes, header));
    appendItems_(quoteId, built.items);
    return { quoteId: quoteId, itemCount: built.items.length };
  });
}

/** Replaces a quote's header fields and line items. File and source type are kept. */
function updateQuote(quoteId, payload) {
  const built = buildQuote_(payload);
  return withLock_(() => {
    const quotes = openTable_(SHEET_QUOTES, QUOTE_COLUMNS);
    const row = findRow_(quotes, quoteId);
    const header = Object.assign(built.header, { updatedAt: new Date() });
    quotes.sheet.getRange(row._row, 1, 1, quotes.width).setValues([toRow_(quotes, header, row._raw)]);
    deleteRowsWhere_(openTable_(SHEET_ITEMS, ITEM_COLUMNS), 'quoteId', quoteId);
    appendItems_(quoteId, built.items);
    return { quoteId: quoteId, itemCount: built.items.length };
  });
}

/** Changes only the status of a quote. */
function updateQuoteStatus(quoteId, status) {
  if (STATUSES.indexOf(status) < 0) throw new Error('不正なステータスです: ' + status);
  return withLock_(() => {
    const quotes = openTable_(SHEET_QUOTES, QUOTE_COLUMNS);
    const row = findRow_(quotes, quoteId);
    quotes.sheet.getRange(row._row, quotes.index.status + 1).setValue(status);
    quotes.sheet.getRange(row._row, quotes.index.updatedAt + 1).setValue(new Date());
    return true;
  });
}

/**
 * Searches line items joined with their quote header.
 * @param {{keyword?: string, vendor?: string, status?: string, dateFrom?: string, dateTo?: string,
 *          priceMin?: number, priceMax?: number, sourceType?: string}} q
 */
function searchItems(q) {
  q = q || {};
  const quotes = {};
  readTable_(openTable_(SHEET_QUOTES, QUOTE_COLUMNS)).forEach(h => { quotes[h.quoteId] = h; });

  const terms = normalize_(q.keyword || '').split(/\s+/).filter(Boolean);
  const priceMin = q.priceMin === '' || q.priceMin == null ? null : Number(q.priceMin);
  const priceMax = q.priceMax === '' || q.priceMax == null ? null : Number(q.priceMax);

  const results = [];
  let total = 0;
  readTable_(openTable_(SHEET_ITEMS, ITEM_COLUMNS)).forEach(it => {
    const h = quotes[it.quoteId];
    if (!h) return;
    if (q.vendor && h.vendor !== q.vendor) return;
    if (q.status && h.status !== q.status) return;
    if (q.sourceType && h.sourceType !== q.sourceType) return;
    if (q.dateFrom && (!h.quoteDate || h.quoteDate < q.dateFrom)) return;
    if (q.dateTo && (!h.quoteDate || h.quoteDate > q.dateTo)) return;
    if (priceMin !== null && it.unitPrice < priceMin) return;
    if (priceMax !== null && it.unitPrice > priceMax) return;
    if (terms.length) {
      const haystack = normalize_([it.name, it.spec, it.note, h.vendor, h.subject, h.quoteNo].join(' '));
      if (!terms.every(t => haystack.indexOf(t) >= 0)) return;
    }
    total++;
    if (results.length < MAX_SEARCH_RESULTS) {
      results.push(Object.assign(strip_(it), {
        vendor: h.vendor, subject: h.subject, quoteNo: h.quoteNo, quoteDate: h.quoteDate,
        status: h.status, sourceType: h.sourceType,
      }));
    }
  });
  results.sort((a, b) => (b.quoteDate || '').localeCompare(a.quoteDate || ''));
  return { total: total, items: results };
}

/**
 * All quote headers, newest quote date first.
 * Filtering for the 見積管理 screen is done on the client.
 */
function listQuotes() {
  return readTable_(openTable_(SHEET_QUOTES, QUOTE_COLUMNS)).map(strip_).sort((a, b) =>
    (b.quoteDate || '').localeCompare(a.quoteDate || '') || b.registeredAt.localeCompare(a.registeredAt));
}

/** One quote with its line items. */
function getQuote(quoteId) {
  const header = findRow_(openTable_(SHEET_QUOTES, QUOTE_COLUMNS), quoteId);
  const items = readTable_(openTable_(SHEET_ITEMS, ITEM_COLUMNS))
    .filter(it => it.quoteId === quoteId)
    .sort((a, b) => a.lineNo - b.lineNo)
    .map(strip_);
  return { header: strip_(header), items: items };
}

/** Deletes a quote and its line items. */
function deleteQuote(quoteId) {
  return withLock_(() => {
    deleteRowsWhere_(openTable_(SHEET_ITEMS, ITEM_COLUMNS), 'quoteId', quoteId);
    deleteRowsWhere_(openTable_(SHEET_QUOTES, QUOTE_COLUMNS), 'quoteId', quoteId);
    return true;
  });
}

// ---------------------------------------------------------------------------
// Quote building
// ---------------------------------------------------------------------------

/** Validates the payload and computes amounts. Tax defaults to 10% of the subtotal. */
function buildQuote_(payload) {
  const h = payload.header || {};
  const items = (payload.items || []).filter(it => String(it.name || '').trim() !== '').map((it, i) => {
    const qty = toNumber_(it.qty);
    const unitPrice = toNumber_(it.unitPrice);
    const amount = it.amount !== '' && it.amount != null ? toNumber_(it.amount) : qty * unitPrice;
    return { lineNo: i + 1, name: String(it.name).trim(), spec: it.spec || '', qty: qty, unit: it.unit || '',
      unitPrice: unitPrice, amount: amount, note: it.note || '' };
  });
  if (!String(h.vendor || '').trim()) throw new Error('取引先は必須です。');
  if (items.length === 0) throw new Error('明細が 1 行もありません。');

  const header = {};
  EDITABLE_KEYS.forEach(k => { header[k] = h[k] == null ? '' : String(h[k]).trim(); });
  if (STATUSES.indexOf(header.status) < 0) header.status = STATUSES[0];
  header.subtotal = items.reduce((sum, it) => sum + it.amount, 0);
  header.tax = h.tax !== '' && h.tax != null ? toNumber_(h.tax) : Math.floor(header.subtotal * TAX_RATE);
  header.total = header.subtotal + header.tax;
  return { header: header, items: items };
}

function appendItems_(quoteId, items) {
  const t = openTable_(SHEET_ITEMS, ITEM_COLUMNS);
  const rows = items.map(it => toRow_(t, Object.assign({ itemId: quoteId + '-' + it.lineNo, quoteId: quoteId }, it)));
  t.sheet.getRange(t.sheet.getLastRow() + 1, 1, rows.length, t.width).setValues(rows);
}

// ---------------------------------------------------------------------------
// Sheet access by header name
// ---------------------------------------------------------------------------

function sheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error('シート「' + name + '」がありません。先に setup() を実行してください。');
  return sheet;
}

/** @return {{sheet: Sheet, width: number, index: Object<string, number>}} index is 0-based column per key */
function openTable_(name, columns) {
  const sheet = sheet_(name);
  const width = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(String);
  const index = {};
  columns.forEach(([key, label]) => {
    const i = headers.indexOf(label);
    if (i < 0) throw new Error('シート「' + name + '」に列「' + label + '」がありません。setup() を実行してください。');
    index[key] = i;
  });
  return { sheet: sheet, width: width, index: index };
}

/** Rows as objects. _row is the 1-based sheet row, _raw the original values. */
function readTable_(t) {
  const last = t.sheet.getLastRow();
  if (last <= 1) return [];
  return t.sheet.getRange(2, 1, last - 1, t.width).getValues()
    .map((raw, i) => {
      const obj = { _row: i + 2, _raw: raw };
      Object.keys(t.index).forEach(k => { obj[k] = fromCell_(k, raw[t.index[k]]); });
      return obj;
    })
    .filter(o => o[Object.keys(t.index)[0]]);
}

/** Builds a sheet row from an object, keeping values of columns the object doesn't cover. */
function toRow_(t, obj, base) {
  const row = base ? base.slice() : new Array(t.width).fill('');
  Object.keys(t.index).forEach(k => { if (k in obj) row[t.index[k]] = toCell_(k, obj[k]); });
  return row;
}

// _raw may contain Date objects, which google.script.run cannot return.
function strip_(obj) {
  const copy = Object.assign({}, obj);
  delete copy._row;
  delete copy._raw;
  return copy;
}

function findRow_(t, quoteId) {
  const row = readTable_(t).filter(r => r.quoteId === quoteId)[0];
  if (!row) throw new Error('見積が見つかりません: ' + quoteId);
  return row;
}

function deleteRowsWhere_(t, key, value) {
  const col = t.index[key];
  const values = t.sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][col]) === value) t.sheet.deleteRow(i + 1);
  }
}

// google.script.run cannot return Date objects, so dates go to the client as strings.
function fromCell_(key, v) {
  if (NUMBER_KEYS.indexOf(key) >= 0) return Number(v) || 0;
  if (key === 'status') return String(v || '') || STATUSES[0]; // rows from before statuses existed
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Tokyo', DATETIME_KEYS.indexOf(key) >= 0 ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  return v == null ? '' : String(v);
}

function toCell_(key, v) {
  if (DATE_KEYS.indexOf(key) >= 0) {
    const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : (v || '');
  }
  return v == null ? '' : v;
}

// ---------------------------------------------------------------------------
// Misc helpers
// ---------------------------------------------------------------------------

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function saveFile_(file, quoteId) {
  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  if (!folderId) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(file.base64), file.mimeType, quoteId + '_' + file.name);
  return DriveApp.getFolderById(folderId).createFile(blob).getUrl();
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
    header: { status: '採用', quoteNo: 'TW-2026-0412', quoteDate: '2026-07-03', validUntil: '2026-08-02', vendor: '東和電機株式会社', vendorContact: '佐藤',
      subject: '本社3F 照明LED化工事', deliveryDate: '受注後3週間', deliveryPlace: '本社3F', paymentTerms: '月末締め翌月末払い', staff: '田中',
      sourceType: 'excel', fileName: '見積_東和電機_20260703.xlsx' },
    items: [
      { name: 'LEDベースライト', spec: 'LDL40 昼白色 4000lm', qty: 48, unit: '台', unitPrice: 12800 },
      { name: '既設照明器具撤去', spec: '40W×2灯', qty: 48, unit: '台', unitPrice: 1500 },
      { name: '配線工事', spec: 'VVF1.6-3C', qty: 1, unit: '式', unitPrice: 85000 },
      { name: '産廃処分費', spec: '', qty: 1, unit: '式', unitPrice: 32000 },
    ],
  },
  {
    header: { status: '不採用', quoteNo: 'Q-2026-0718', quoteDate: '2026-07-18', validUntil: '2026-08-17', vendor: 'ネットワークス山田', vendorContact: '山田',
      subject: '会議室 LAN 増設', deliveryDate: '2026-08-末', deliveryPlace: '本社2F 会議室', paymentTerms: '検収後30日', staff: '鈴木',
      sourceType: 'pdf', fileName: 'Q-2026-0718.pdf', memo: '他社相見積の結果、不採用' },
    items: [
      { name: 'LANケーブル', spec: 'Cat6A 305m巻', qty: 2, unit: '巻', unitPrice: 38000 },
      { name: 'スイッチングハブ', spec: '24ポート PoE+ GS1900-24HP', qty: 1, unit: '台', unitPrice: 64500 },
      { name: '情報コンセント', spec: 'Cat6A 2口', qty: 12, unit: '個', unitPrice: 2400 },
      { name: '敷設・結線作業', spec: '', qty: 2, unit: '人日', unitPrice: 45000 },
    ],
  },
  {
    header: { status: '検討中', quoteNo: 'TW-2026-0533', quoteDate: '2026-08-05', validUntil: '2026-09-04', vendor: '東和電機株式会社', vendorContact: '佐藤',
      subject: '倉庫 照明・コンセント増設', deliveryDate: '受注後1ヶ月', deliveryPlace: '第2倉庫', paymentTerms: '月末締め翌月末払い', staff: '田中',
      sourceType: 'pdf', fileName: '見積書_倉庫.pdf' },
    items: [
      { name: 'LED高天井照明', spec: '150W 水銀灯400W相当', qty: 16, unit: '台', unitPrice: 29800 },
      { name: 'コンセント増設', spec: '2口 接地付', qty: 10, unit: '箇所', unitPrice: 8500 },
      { name: '高所作業車', spec: '10m 1日', qty: 2, unit: '日', unitPrice: 38000 },
    ],
  },
  {
    header: { status: '保留', quoteDate: '2026-08-22', validUntil: '2026-10-03', vendor: '丸山建材', vendorContact: '丸山',
      subject: '事務所 内装補修', deliveryPlace: '本社1F 事務所', paymentTerms: '現金', staff: '鈴木',
      sourceType: 'handwritten', fileName: 'memo_0822.jpg', memo: '現地打合せ時の手書きメモより' },
    items: [
      { name: '石膏ボード', spec: '12.5mm 910×1820', qty: 20, unit: '枚', unitPrice: 980 },
      { name: 'クロス張替', spec: '量産品', qty: 45, unit: 'm2', unitPrice: 1200 },
      { name: 'タイルカーペット', spec: '500角 グレー', qty: 60, unit: '枚', unitPrice: 750 },
    ],
  },
  {
    header: { status: '検討中', quoteNo: 'Q-2026-0910', quoteDate: '2026-09-10', validUntil: '2026-10-10', vendor: 'ネットワークス山田', vendorContact: '山田',
      subject: '無線LAN 更新', deliveryDate: '2026-11-中旬', deliveryPlace: '本社 全フロア', paymentTerms: '検収後30日', staff: '田中',
      sourceType: 'excel', fileName: '無線LAN更新_見積.xlsx' },
    items: [
      { name: '無線アクセスポイント', spec: 'Wi-Fi 6E WAX630E', qty: 6, unit: '台', unitPrice: 52000 },
      { name: 'LANケーブル', spec: 'Cat6 20m', qty: 6, unit: '本', unitPrice: 2200 },
      { name: '設定作業', spec: 'SSID/VLAN設計含む', qty: 1, unit: '式', unitPrice: 60000 },
    ],
  },
];
