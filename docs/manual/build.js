/**
 * Builds the employee manual PDF.
 *
 *   1. Opens gas/Index.html in local preview mode (built-in sample data) and saves
 *      annotated screenshots to docs/manual/img/.
 *   2. Prints docs/manual/manual.html to docs/見積明細検索_取扱説明書.pdf and fills the
 *      table of contents with the page each chapter starts on.
 *
 * Requirements: the `playwright` package with its full Chromium build (`npx playwright install chromium`)
 * and the Noto Sans CJK JP font (Debian/Ubuntu: fonts-noto-cjk).
 *
 * Usage: node docs/manual/build.js [--lib-dir DIR]
 *   --lib-dir DIR  Load xlsx.full.min.js / pdf.min.js / pdf.worker.min.js from DIR
 *                  instead of cdnjs (for machines without internet access).
 */
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const { pathToFileURL } = require('url');
const { chromium } = require('playwright');

const ROOT = path.resolve(__dirname, '..', '..');
const APP = path.join(ROOT, 'gas', 'Index.html');
const MANUAL = path.join(__dirname, 'manual.html');
const IMG_DIR = path.join(__dirname, 'img');
const OUT = path.join(ROOT, 'docs', '見積明細検索_取扱説明書.pdf');
const SAMPLE_XLSX = path.join(__dirname, 'sample-quote.xlsx');

// Screens are captured on a fixed date so expiry tags (期限切れ / 期限間近) come out the same every build.
const FIXED_NOW = new Date('2026-09-28T10:00:00+09:00');
const VIEWPORT = { width: 1180, height: 800 };
const ANNO_COLOR = '#e8590c';
const APP_CSS = `
  #mode, .toast { display: none !important; }
  body { font-family: "Noto Sans CJK JP", "Noto Sans JP", "Hiragino Sans", "Yu Gothic UI", Meiryo, sans-serif !important; }
`;
// Chapter titles as written in manual.html; used to find the page each chapter starts on.
const CHAPTERS = {
  ch1: 'はじめに',
  ch2: '画面の開き方と基本操作',
  ch3: '見積書を登録する',
  ch4: '明細を検索する',
  ch5: '見積を管理する',
  ch6: '見積を確認・修正・削除する',
  ch7: '困ったときは',
};
// Text printed by the @page footer in manual.html (not part of the page body).
const FOOTER_TEXT = '見積明細検索　取扱説明書 0123456789/';
const TINY_PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=', 'base64');

function argValue(name) {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : null;
}

async function routeLibs(context, libDir) {
  if (!libDir) return;
  await context.route('https://cdnjs.cloudflare.com/**', route =>
    route.fulfill({ path: path.join(libDir, route.request().url().split('/').pop()), contentType: 'application/javascript' }));
}

/**
 * Draws numbered boxes: marks = [[selector, number, padding?], ...].
 * A box surrounds every element the selector matches (e.g. all tab buttons).
 */
async function annotate(page, marks) {
  await page.evaluate(({ marks, color }) => {
    for (const [sel, n, pad = 4] of marks) {
      const rects = [...document.querySelectorAll(sel)].map(el => el.getBoundingClientRect());
      if (!rects.length) throw new Error('annotation target not found: ' + sel);
      const left = Math.min(...rects.map(r => r.left)), top = Math.min(...rects.map(r => r.top));
      const width = Math.max(...rects.map(r => r.right)) - left, height = Math.max(...rects.map(r => r.bottom)) - top;
      const x = left + scrollX - pad;
      const y = top + scrollY - pad;
      const box = document.createElement('div');
      box.className = 'manual-anno';
      box.style.cssText = `position:absolute;left:${x}px;top:${y}px;width:${width + pad * 2}px;height:${height + pad * 2}px;` +
        `border:3px solid ${color};border-radius:8px;z-index:9998;pointer-events:none;box-sizing:border-box`;
      const badge = document.createElement('div');
      badge.className = 'manual-anno';
      badge.textContent = n;
      badge.style.cssText = `position:absolute;left:${x - 13}px;top:${y - 13}px;width:26px;height:26px;border-radius:50%;` +
        `background:${color};color:#fff;font:700 15px/26px "Noto Sans CJK JP",sans-serif;text-align:center;` +
        `z-index:9999;box-shadow:0 0 0 2px #fff;pointer-events:none`;
      document.body.append(box, badge);
    }
  }, { marks, color: ANNO_COLOR });
}

/** Screenshots the union of the given elements (plus annotations), then removes the annotations. */
async function capture(page, name, selectors, marks = [], margin = 16) {
  await annotate(page, marks);
  const clip = await page.evaluate(({ selectors, margin }) => {
    const rects = selectors.map(s => {
      const el = document.querySelector(s);
      if (!el) throw new Error('clip target not found: ' + s);
      return el.getBoundingClientRect();
    });
    document.querySelectorAll('.manual-anno').forEach(e => rects.push(e.getBoundingClientRect()));
    const doc = document.documentElement;
    const x0 = Math.max(0, Math.min(...rects.map(r => r.left)) + scrollX - margin);
    const y0 = Math.max(0, Math.min(...rects.map(r => r.top)) + scrollY - margin);
    const x1 = Math.min(doc.scrollWidth, Math.max(...rects.map(r => r.right)) + scrollX + margin);
    const y1 = Math.min(doc.scrollHeight, Math.max(...rects.map(r => r.bottom)) + scrollY + margin);
    return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
  }, { selectors, margin });
  await page.screenshot({ path: path.join(IMG_DIR, name + '.png'), clip, fullPage: true });
  await page.evaluate(() => document.querySelectorAll('.manual-anno').forEach(e => e.remove()));
  console.log('  img/' + name + '.png');
}

async function captureScreens(browser, libDir) {
  const context = await browser.newContext({
    viewport: VIEWPORT, deviceScaleFactor: 2, colorScheme: 'light', locale: 'ja-JP', timezoneId: 'Asia/Tokyo',
  });
  await routeLibs(context, libDir);
  await context.clock.setFixedTime(FIXED_NOW);
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e));
  await page.goto(pathToFileURL(APP).href);
  await page.addStyleTag({ content: APP_CSS });
  await page.waitForFunction(() => /見積 \d+ 件/.test(document.querySelector('#stats').textContent));
  await page.evaluate(() => document.fonts.ready);

  // Cover and overview: the 見積管理 list.
  await page.click('nav button[data-tab=list]');
  await page.waitForSelector('#list-body select.st');
  await capture(page, 'cover', ['header', '#tab-list'], [], 0);

  // 2章: header, captured narrower so its small text stays readable in the manual
  await page.setViewportSize({ width: 640, height: VIEWPORT.height });
  await capture(page, 'fig-header', ['header'], [['nav button', 1, 2], ['.stats', 2, 6]], 18);
  await page.setViewportSize(VIEWPORT);

  // 5章: 見積管理
  await capture(page, 'fig-list', ['header', '#tab-list'], [
    ['#l-keyword', 1], ['#list-form .filters', 2], ['#list-summary .chips', 3],
    ['#list-body tr:first-child select.st', 4], ['#list-body .exp.expired', 5, 3], ['#list-body tr:first-child a.link', 6, 3],
  ], 0);

  // 4章: 明細検索
  await page.click('nav button[data-tab=search]');
  await page.fill('#q-keyword', 'LANケーブル');
  await page.waitForFunction(() => /該当 2 件/.test(document.querySelector('#search-summary').textContent));
  await capture(page, 'fig-search', ['header', '#tab-search'], [
    ['#q-keyword', 1], ['#search-form .filters', 2], ['#search-summary', 3], ['#search-body tr:first-child a.link', 4, 3],
  ], 0);

  // 3章: 見積登録
  await page.click('nav button[data-tab=register]');
  await capture(page, 'fig-drop', ['#reg-step1'], [['#drop', 1], ['#manual-btn', 2]]);

  await page.setInputFiles('#file', SAMPLE_XLSX);
  await page.waitForFunction(() => /抽出しました/.test(document.querySelector('#extract-notice').textContent));
  await page.fill('#h-staff', '田中');
  await capture(page, 'fig-review', ['#reg-step2'], [
    ['#extract-notice', 1], ['#header-form', 2], ['#reg-step2 table.items', 3], ['#add-row', 4], ['.amounts', 5], ['#reg-submit', 6],
  ]);

  await page.click('#reg-cancel');
  await page.setInputFiles('#file', { name: 'memo.png', mimeType: 'image/png', buffer: TINY_PNG });
  await page.waitForFunction(() => /OCR はモック/.test(document.querySelector('#extract-notice').textContent));
  await capture(page, 'fig-lowconf', ['#reg-step2 fieldset:has(#items-body)'], [['#items-body tr.lowconf:nth-child(2)', 1, 2]], 8);
  await page.click('#reg-cancel');

  // 6章: 詳細と削除
  await page.click('nav button[data-tab=list]');
  await page.click('#list-body a.link:has-text("倉庫 照明・コンセント増設")');
  await page.waitForSelector('#m-edit');
  await capture(page, 'fig-detail', ['#modal .modal'], [['#modal dl.hd select.st', 1, 3], ['#m-edit', 2, 3], ['#m-del', 3, 3]], 8);
  await page.click('#m-del'); // first press only arms the button
  // Close-up of the armed button; the rest of the dialog is hidden so the crop shows nothing else.
  await page.evaluate(() => document.querySelectorAll('#modal .modal > :not(.reg-foot)').forEach(e => { e.style.visibility = 'hidden'; }));
  await capture(page, 'fig-delete', ['#m-del'], [['#m-del', 1, 3]], 4);
  await page.click('#m-close');

  await context.close();
  if (errors.length) throw errors[0];
}

// Chromium records some kanji in the PDF text layer as look-alike radical characters when the font
// maps both to one glyph (見 -> U+2F92 ⾒), which breaks searching and copying the PDF text.
// Kangxi radicals (U+2F00-2FDF) turn back into kanji via NFKC; CJK Radicals Supplement characters
// (U+2E80-2EFF) have no NFKC form, so these pairs list the kanji sharing each glyph in Noto Sans CJK JP.
// ⼾ is listed too because NFKC gives 戶 while the Japanese glyph is 戸.
const RADICAL_PAIRS = '⼾戸⺂乛⺃乚⺅亻⺇𠘨⺉刂⺋㔾⺍𭕄⺎兀⺏尣⺐尢⺒巳⺓幺⺔彑⺖忄⺘扌⺙攵⺛旡⺞歺⺠民⺡氵⺣灬⺦丬⺨犭⺩𤣩⺫罒⺭礻⺮𥫗⺯糹⺰纟⺱罓⺲罒⺹耂⺺肀⺽𦥑⺾艹⻂衤⻃覀⻄西⻅见⻇𧢲⻈讠⻉贝⻊𧾷⻋车⻍辶⻐钅⻑長⻒镸⻓长⻕𨸏⻖阝⻘青⻙韦⻚页⻛风⻜飞⻞𩙿⻟飠⻠饣⻡𩠐⻢马⻥鱼⻦鸟⻧卤⻨麦⻩黄⻪黾⻫斉⻬齐⻭歯⻮齿⻯竜⻰龙⻲亀';
const RADICALS = new Map();
{
  const cps = Array.from(RADICAL_PAIRS);
  for (let i = 0; i < cps.length; i += 2) RADICALS.set(cps[i], cps[i + 1]);
}
// Other characters the font draws with the glyph of a character we actually use.
const SHARED_GLYPHS = new Map([['‧', '・']]); // U+2027 HYPHENATION POINT -> U+30FB KATAKANA MIDDLE DOT
const isRadical = ch => ch.codePointAt(0) >= 0x2E80 && ch.codePointAt(0) <= 0x2FDF;
const fixChar = ch => SHARED_GLYPHS.get(ch) || (isRadical(ch) ? RADICALS.get(ch) || ch.normalize('NFKC') : ch);
const fixText = str => Array.from(str, fixChar).join('');

const utf16HexToString = hex => String.fromCharCode(...hex.match(/.{4}/g).map(h => parseInt(h, 16)));
const stringToUtf16Hex = str => Array.from({ length: str.length }, (_, i) => str.charCodeAt(i).toString(16).toUpperCase().padStart(4, '0')).join('');

/** Rewrites a ToUnicode CMap through fixText; returns null when nothing needs changing. */
function fixToUnicode(cmap) {
  const head = cmap.slice(0, cmap.indexOf('endcodespacerange') + 'endcodespacerange'.length);
  const tail = cmap.slice(cmap.indexOf('endcmap'));
  const body = cmap.slice(head.length, cmap.length - tail.length);
  const map = new Map(); // source code (hex) -> text
  for (const [, block] of body.matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const [, src, dst] of block.matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g)) map.set(src, utf16HexToString(dst));
  }
  for (const [, block] of body.matchAll(/beginbfrange([\s\S]*?)endbfrange/g)) {
    for (const [, lo, hi, dst] of block.matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]+>|\[[^\]]*\])/g)) {
      const from = parseInt(lo, 16), to = parseInt(hi, 16);
      const list = dst.startsWith('[') ? [...dst.matchAll(/<([0-9A-Fa-f]+)>/g)].map(m => m[1]) : null;
      for (let c = from; c <= to; c++) {
        let text;
        if (list) text = utf16HexToString(list[c - from]);
        else { // consecutive range: the last UTF-16 unit counts up
          const units = utf16HexToString(dst.slice(1, -1));
          text = units.slice(0, -1) + String.fromCharCode(units.charCodeAt(units.length - 1) + c - from);
        }
        map.set(c.toString(16).toUpperCase().padStart(lo.length, '0'), text);
      }
    }
  }
  if (![...map.values()].some(t => fixText(t) !== t)) return null;
  const entries = [...map.entries()].sort((a, b) => parseInt(a[0], 16) - parseInt(b[0], 16));
  let out = '';
  for (let i = 0; i < entries.length; i += 100) {
    const chunk = entries.slice(i, i + 100);
    out += `\n${chunk.length} beginbfchar\n` + chunk.map(([src, t]) => `<${src}> <${stringToUtf16Hex(fixText(t))}>`).join('\n') + '\nendbfchar';
  }
  return head + out + '\n' + tail;
}

/**
 * Applies fixToUnicode to every ToUnicode stream of a Chromium PDF (classic xref table, objects
 * numbered 1..N) and returns the file rewritten with a new xref table.
 */
function fixPdfText(pdf) {
  const s = pdf.toString('latin1');
  const startxref = s.lastIndexOf('startxref');
  const xrefPos = parseInt(s.slice(startxref + 9).trim(), 10);
  if (!s.startsWith('xref', xrefPos)) throw new Error('unsupported PDF: expected a classic xref table');
  const trailerPos = s.indexOf('trailer', xrefPos);
  const lines = s.slice(xrefPos + 4, trailerPos).trim().split(/\r?\n/);
  const objs = [];
  for (let i = 0; i < lines.length;) {
    const [first, count] = lines[i++].trim().split(/\s+/).map(Number);
    for (let k = 0; k < count; k++) {
      const [off, , type] = lines[i++].trim().split(/\s+/);
      if (type === 'n') objs.push({ num: first + k, off: Number(off) });
    }
  }
  objs.sort((a, b) => a.off - b.off);
  objs.forEach((o, i) => { o.text = s.slice(o.off, i + 1 < objs.length ? objs[i + 1].off : xrefPos); });
  if (objs.some((o, i, all) => !all.find(x => x.num === i + 1))) throw new Error('unsupported PDF: object numbers are not 1..N');

  const toUnicode = new Set([...s.matchAll(/\/ToUnicode (\d+) 0 R/g)].map(m => Number(m[1])));
  for (const o of objs) {
    if (!toUnicode.has(o.num)) continue;
    const m = o.text.match(/^(\d+ 0 obj\s*<<)([\s\S]*?)(>>\s*stream\r?\n)/);
    const length = m && Number((m[2].match(/\/Length (\d+)(?! \d+ R)/) || [])[1]);
    if (!length) throw new Error('unsupported ToUnicode stream in object ' + o.num);
    const start = m[0].length;
    const raw = Buffer.from(o.text.slice(start, start + length), 'latin1');
    const flate = /\/FlateDecode/.test(m[2]);
    const cmap = fixToUnicode((flate ? zlib.inflateSync(raw) : raw).toString('latin1'));
    if (cmap === null) continue;
    const data = flate ? zlib.deflateSync(Buffer.from(cmap, 'latin1')) : Buffer.from(cmap, 'latin1');
    o.text = m[1] + m[2].replace(/\/Length \d+/, '/Length ' + data.length) + m[3] + data.toString('latin1') + '\nendstream\nendobj\n';
  }

  let out = s.slice(0, objs[0].off);
  const offsets = [];
  for (const o of objs) { offsets[o.num] = out.length; out += o.text; }
  const xref = out.length;
  out += `xref\n0 ${offsets.length}\n0000000000 65535 f \n` +
    offsets.slice(1).map(off => String(off).padStart(10, '0') + ' 00000 n \n').join('') +
    s.slice(trailerPos, startxref) + `startxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

/**
 * Returns { ch1: page, ... } by finding each chapter heading ("3見積書を登録する") in the page text,
 * searching after the table of contents and in chapter order (every chapter starts a new page).
 * PDF bookmarks would be simpler, but Chromium duplicates heading text in them with this font.
 */
async function chapterPages(browser, libDir, pdf, expectedText) {
  const context = await browser.newContext();
  await routeLibs(context, libDir);
  const page = await context.newPage();
  await page.goto(pathToFileURL(APP).href); // loads pdf.js the same way the app does
  const texts = await page.evaluate(async b64 => {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    const doc = await pdfjsLib.getDocument({ data: Uint8Array.from(atob(b64), c => c.charCodeAt(0)) }).promise;
    const texts = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const items = (await (await doc.getPage(p)).getTextContent()).items;
      texts.push(items.map(i => i.str).join('').replace(/\s+/g, ''));
    }
    return texts;
  }, pdf.toString('base64'));
  await context.close();
  // Every character in the PDF text must appear in the manual itself (or its footer), so search and copy work.
  const allowed = new Set(Array.from(expectedText + FOOTER_TEXT));
  const unexpected = [...new Set(Array.from(texts.join('')))].filter(ch => !allowed.has(ch));
  if (unexpected.length) {
    throw new Error('PDF text has characters that are not in the manual (add them to SHARED_GLYPHS): ' +
      unexpected.map(ch => `${ch} U+${ch.codePointAt(0).toString(16).toUpperCase()}`).join(', '));
  }

  const found = {};
  let after = texts.findIndex(t => t.includes('目次')) + 1; // pages are 1-based
  for (const [id, title] of Object.entries(CHAPTERS)) {
    const i = texts.findIndex((t, idx) => idx + 1 > after && t.includes(id.replace('ch', '') + title));
    if (i < 0) throw new Error(`chapter heading not found in PDF: ${id} ${title}`);
    found[id] = after = i + 1;
  }
  return { found, total: texts.length };
}

async function printManual(browser, libDir) {
  const page = await browser.newPage();
  await page.goto(pathToFileURL(MANUAL).href, { waitUntil: 'load' });
  await page.evaluate(() => document.fonts.ready);
  const print = async () => fixPdfText(await page.pdf({ preferCSSPageSize: true, printBackground: true }));

  // Pass 1 finds where chapters land; pass 2 prints with those numbers in the table of contents.
  const expectedText = await page.evaluate(() => document.body.innerText);
  const first = await chapterPages(browser, libDir, await print(), expectedText);
  await page.evaluate(found => {
    for (const [id, p] of Object.entries(found)) document.querySelector(`[data-page-of="${id}"]`).textContent = p;
  }, first.found);
  const pdf = await print();
  const second = await chapterPages(browser, libDir, pdf, expectedText);
  if (JSON.stringify(first.found) !== JSON.stringify(second.found)) throw new Error('page numbers moved after filling the table of contents');
  fs.writeFileSync(OUT, pdf);
  console.log(`  ${path.relative(ROOT, OUT)} (${second.total} pages)`, second.found);
  await page.close();
}

(async () => {
  const libDir = argValue('--lib-dir');
  // Date inputs should render in Japanese format (2026/09/15), as employees will see them. That takes
  // the full Chromium build (the default headless shell only has en-US locale data) started with a
  // Japanese UI language, which on Linux comes from LANGUAGE together with --lang.
  const browser = await chromium.launch({
    args: ['--lang=ja-JP'],
    env: { ...process.env, LANGUAGE: 'ja' },
    ...(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : { channel: 'chromium' }),
  });
  try {
    fs.mkdirSync(IMG_DIR, { recursive: true });
    console.log('Capturing screens');
    await captureScreens(browser, libDir);
    console.log('Printing PDF');
    await printManual(browser, libDir);
  } finally {
    await browser.close();
  }
})().catch(e => { console.error(e); process.exit(1); });
