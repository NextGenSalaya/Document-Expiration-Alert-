/**
 * แจ้งเตือนเอกสารหมดอายุผ่าน LINE: Work Permit / Passport / Visa
 * ใช้กับ Google Sheet ที่นำเข้าจากไฟล์ "รายชื่อ AEC และวันที่หมดอายุของเอกสาร"
 *
 * วิธีใช้: เมนู "แจ้งเตือน LINE" ด้านบนของ Google Sheet
 * LINE token เก็บใน Script Properties (ไม่อยู่ในชีต คนที่เปิดชีตจะไม่เห็น)
 */

// ===================== ตั้งค่า (แก้ได้) =====================
const CONFIG = {
  DAYS_AHEAD: 90,             // แจ้งเตือนล่วงหน้ากี่วัน
  INCLUDE_EXPIRED: true,      // รวมเอกสารที่หมดอายุแล้วด้วย
  INCLUDE_DOC_NUMBERS: false, // true = ส่งเลขหนังสือเดินทาง / เลขคนต่างด้าว / เลขใบอนุญาตทำงานไปใน LINE ด้วย
  SKIP_IF_EMPTY: true,        // ไม่ส่งถ้าไม่มีรายการ
  MESSAGE_STYLE: 'card',      // 'card' = การ์ดสรุป (Flex Message) · 'text' = ข้อความตัวหนังสือ
  SEND_HOUR: 8,               // เวลาส่งอัตโนมัติ (ชั่วโมง 0-23)
  SHEET_NAME: '',             // ว่างไว้ = หาชีตที่มีคอลัมน์ "PeopleSoft ID" ให้เอง
  WEB_URL: 'https://nextgensalaya.github.io/Document-Expiration-Alert-/',
};
// ============================================================

const DOCS = [
  { col: 'wpExp',   label: 'Work Permit' },
  { col: 'ppExp',   label: 'Passport' },
  { col: 'visaExp', label: 'Visa' },
];
const MATCHERS = [ // expiry columns first so "Passport หมดอายุ" never lands in a number column
  ['wpExp',      h => h.includes('workpermit') && h.includes('หมดอายุ')],
  ['ppExp',      h => h.includes('passport') && h.includes('หมดอายุ')],
  ['visaExp',    h => h.includes('visa') && h.includes('หมดอายุ')],
  ['psid',       h => h.includes('peoplesoft')],
  ['prefix',     h => h.includes('คำนำหน้า')],
  ['nameTh',     h => h.startsWith('ชื่อ') && h.includes('ไทย')],
  ['dept',       h => h.includes('หน่วยงาน')],
  ['passportNo', h => h.includes('เลขที่หนังสือเดินทาง') || h.includes('หนังสือเดินทางเลขที่')],
  ['alienId',    h => h.includes('คนต่างด้าว')],
  ['wpNo',       h => h.includes('ใบอนุญาตทำงานเลขที่') || h.includes('เลขที่ใบอนุญาตทำงาน')],
];
const TEXT_FIELDS = ['psid', 'prefix', 'nameTh', 'dept', 'passportNo', 'alienId', 'wpNo'];
const TH_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const MONTH_NAMES = {
  'มกราคม':1,'กุมภาพันธ์':2,'มีนาคม':3,'เมษายน':4,'พฤษภาคม':5,'มิถุนายน':6,'กรกฎาคม':7,'สิงหาคม':8,'กันยายน':9,'ตุลาคม':10,'พฤศจิกายน':11,'ธันวาคม':12,
  jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12,
};
TH_MONTHS.forEach((m, i) => MONTH_NAMES[m.replace(/\./g, '')] = i + 1);
const MAX_CHARS = 4800; // LINE text message limit is 5000

// ---------------- menu ----------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('แจ้งเตือน LINE')
    .addItem('ดูตัวอย่างข้อความ (ยังไม่ส่ง)', 'previewAlerts')
    .addItem('ส่งแจ้งเตือนตอนนี้', 'sendAlertsNow')
    .addSeparator()
    .addItem('ตั้งให้ส่งอัตโนมัติทุกวัน', 'scheduleDaily')
    .addItem('ตั้งให้ส่งอัตโนมัติทุกวันจันทร์', 'scheduleWeekly')
    .addItem('ปิดการส่งอัตโนมัติ', 'unschedule')
    .addSeparator()
    .addItem('ตั้งค่า LINE token', 'setToken')
    .addItem('ส่งเข้ากลุ่ม LINE: ขอรหัสลงทะเบียนกลุ่ม', 'showRegisterCode')
    .addItem('ส่งหาทุกคนที่แอดบอท (ยกเลิกการส่งเข้ากลุ่ม)', 'clearTarget')
    .addItem('ดูสถานะ', 'showStatus')
    .addToUi();
}

function previewAlerts() {
  const { messages, cards, total } = buildFromSheet_();
  const text = messages.length ? messages.join('\n\n────────── (ข้อความถัดไป) ──────────\n\n') : 'ไม่มีเอกสารที่จะหมดอายุในช่วงนี้';
  const style = CONFIG.MESSAGE_STYLE === 'card'
    ? `ส่งเป็นการ์ด ${cards.reduce((s, m) => s + m.contents.contents.length, 0)} ใบ (${cards.length} ข้อความ LINE) · ด้านล่างคือเนื้อหาเดียวกันแบบตัวหนังสือ`
    : `${messages.length} ข้อความ LINE`;
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px/1.5 sans-serif"><p><b>' + total + ' รายการ · ' + esc_(style) + '</b></p>' +
    '<pre style="white-space:pre-wrap;font:13px/1.5 sans-serif;background:#f4f5f7;padding:12px;border-radius:8px">' + esc_(text) + '</pre></div>')
    .setWidth(640).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'ตัวอย่างข้อความ LINE');
}

function sendAlertsNow() {
  const ui = SpreadsheetApp.getUi();
  const r = sendAlerts_();
  if (!r.sent) { ui.alert('ไม่มีรายการที่ต้องแจ้ง จึงไม่ได้ส่ง'); return; }
  ui.alert(`[เวอร์ชันการ์ด] ส่งแล้ว ${r.total} รายการ (${r.messages}) ไปที่ ${r.target}` +
    (r.fallback ? `\n\nLINE ไม่รับการ์ด จึงส่งเป็นตัวหนังสือแทน\nเหตุผลจาก LINE:\n${r.fallback.slice(0, 800)}` : ''));
}

/** ตัวที่ trigger อัตโนมัติเรียกใช้ */
function sendAlertsScheduled() { sendAlerts_(); }

function sendAlerts_() {
  const { messages, cards, total } = buildFromSheet_();
  const target = PropertiesService.getScriptProperties().getProperty('TARGET_ID');
  if (!messages.length && CONFIG.SKIP_IF_EMPTY) return { sent: false, total: 0 };
  const textMsgs = (messages.length ? messages : ['✅ ไม่มีเอกสารที่จะหมดอายุภายใน ' + CONFIG.DAYS_AHEAD + ' วัน'])
    .map(t => ({ type: 'text', text: t }));
  let msgs = textMsgs, style = 'ข้อความตัวหนังสือ', fallback = '';
  if (CONFIG.MESSAGE_STYLE === 'card' && cards.length) {
    // Ask LINE to check the cards first (sends nothing); fall back to plain text if they are rejected.
    try {
      for (let i = 0; i < cards.length; i += 5) {
        const batch = cards.slice(i, i + 5);
        linePost_(target ? 'validate/push' : 'validate/broadcast', { messages: batch });
      }
      msgs = cards; style = 'ข้อความแบบการ์ด';
    } catch (err) {
      fallback = err.message;
      console.warn('LINE rejected the card layout, sending text instead: ' + err.message);
      PropertiesService.getScriptProperties().setProperty('LAST_CARD_ERROR', String(err.message).slice(0, 1500));
    }
  }
  for (let i = 0; i < msgs.length; i += 5) { // LINE accepts up to 5 messages per request
    const batch = msgs.slice(i, i + 5);
    if (target) linePost_('push', { to: target, messages: batch });
    else linePost_('broadcast', { messages: batch });
  }
  PropertiesService.getScriptProperties().setProperty('LAST_SENT', new Date().toISOString());
  return { sent: true, total, messages: msgs.length + ' ' + style, fallback, target: target ? 'กลุ่ม/ผู้รับที่ลงทะเบียนไว้' : 'ทุกคนที่แอดบอทเป็นเพื่อน' };
}

// ---------------- reading the sheet ----------------
function buildFromSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = tz_();
  const sheets = CONFIG.SHEET_NAME ? [ss.getSheetByName(CONFIG.SHEET_NAME)] : ss.getSheets();
  for (const sh of sheets) {
    if (!sh) continue;
    const range = sh.getDataRange();
    const raw = range.getValues(), disp = range.getDisplayValues();
    if (findHeader_(disp) < 0) continue;
    const today = ymd_(new Date(), tz);
    const items = collectAlerts_(raw, disp, today, d => ymd_(d, tz));
    return { messages: buildMessages_(items, today), cards: buildCards_(items, today), total: items.length };
  }
  throw new Error('ไม่พบคอลัมน์ "PeopleSoft ID" ในชีตใดเลย ตรวจสอบว่านำเข้าไฟล์ถูกต้อง');
}

function collectAlerts_(raw, disp, today, dateToYmd) {
  const r = findHeader_(disp);
  const map = mapColumns_(disp, r);
  const missing = DOCS.filter(d => !(d.col in map)).map(d => d.label);
  if (!('psid' in map) || missing.length === DOCS.length) throw new Error('ไม่พบคอลัมน์วันหมดอายุ');
  const items = [];
  for (let i = r + 1; i < raw.length; i++) {
    const p = {};
    TEXT_FIELDS.forEach(k => p[k] = map[k] === undefined ? '' : String(disp[i][map[k]] || '').trim());
    if (!p.psid && !p.nameTh) continue;
    DOCS.forEach(d => {
      const c = map[d.col]; if (c === undefined) return;
      const dt = parseDate_(raw[i][c], dateToYmd) || parseDate_(disp[i][c], dateToYmd);
      if (!dt) return;
      const n = Math.round((Date.UTC(dt[0], dt[1] - 1, dt[2]) - Date.UTC(today[0], today[1] - 1, today[2])) / 864e5);
      if (n > CONFIG.DAYS_AHEAD) return;
      if (n < 0 && !CONFIG.INCLUDE_EXPIRED) return;
      items.push({ p, doc: d, date: dt, n });
    });
  }
  return items.sort((a, b) => a.n - b.n);
}

function buildMessages_(items, today) {
  if (!items.length) return [];
  const lines = [];
  lines.push(`⚠️ เอกสารใกล้หมดอายุ (ภายใน ${CONFIG.DAYS_AHEAD} วัน)`);
  lines.push(`ข้อมูล ณ ${thDate_(today)} · รวม ${items.length} รายการ`);
  DOCS.forEach(d => {
    const mine = items.filter(i => i.doc === d);
    if (!mine.length) return;
    const exp = mine.filter(i => i.n < 0).length, urgent = mine.filter(i => i.n >= 0 && i.n <= 30).length;
    lines.push('');
    lines.push(`■ ${d.label} หมดอายุ: ${mine.length} คน` + (exp || urgent ? ` (หมดแล้ว ${exp} · ≤30 วัน ${urgent})` : ''));
    mine.forEach(it => {
      const icon = it.n < 0 ? '🔴' : it.n <= 30 ? '🟠' : it.n <= 60 ? '🟡' : '🔵';
      const when = it.n < 0 ? `เลยมา ${-it.n} วัน` : it.n === 0 ? 'หมดวันนี้' : `อีก ${it.n} วัน`;
      const p = it.p;
      let line = `${icon} ${p.psid} ${[p.prefix, p.nameTh].filter(Boolean).join(' ')}` + (p.dept ? ` | ${p.dept}` : '') + ` | ${thDate_(it.date)} (${when})`;
      if (CONFIG.INCLUDE_DOC_NUMBERS) {
        const nums = [p.passportNo && 'PP ' + p.passportNo, p.alienId && 'ต่างด้าว ' + p.alienId, p.wpNo && 'WP ' + p.wpNo].filter(Boolean);
        if (nums.length) line += '\n    ' + nums.join(' · ');
      }
      lines.push(line);
    });
  });
  lines.push('');
  lines.push('🔴 หมดอายุแล้ว 🟠 ≤30 วัน 🟡 31–60 วัน 🔵 61 วันขึ้นไป');
  if (CONFIG.WEB_URL) lines.push('ดูรายละเอียด: ' + CONFIG.WEB_URL);

  // split into LINE-sized messages on line boundaries
  const out = []; let cur = '';
  lines.forEach(l => {
    if (cur && (cur.length + l.length + 1) > MAX_CHARS) { out.push(cur.replace(/\n+$/, '')); cur = '(ต่อ)\n'; }
    cur += l + '\n';
  });
  if (cur.trim()) out.push(cur.replace(/\n+$/, ''));
  return out;
}

// ---------------- LINE cards (Flex Message) ----------------
const CARD = { navy: '#1C4F8F', yellow: '#F2D21B', ink: '#17202C', muted: '#5A6574', faint: '#8A94A3',
  crit: '#D93025', high: '#E8710A', warn: '#B8860B', watch: '#3F6DB3' };
const ROWS_PER_CARD = 15;

function buildCards_(items, today) {
  if (!items.length) return [];
  const color = n => n < 0 ? CARD.crit : n <= 30 ? CARD.high : n <= 60 ? CARD.warn : CARD.watch;
  const when = n => n < 0 ? `เลย ${-n} วัน` : n === 0 ? 'หมดวันนี้' : `อีก ${n} วัน`;
  const txt = (text, o) => Object.assign({ type: 'text', text: String(text || '-') }, o || {});

  // card 1: summary per document type
  const counts = DOCS.map(d => {
    const m = items.filter(i => i.doc === d);
    return { d, all: m.length, exp: m.filter(i => i.n < 0).length, urg: m.filter(i => i.n >= 0 && i.n <= 30).length };
  });
  const summaryRows = [];
  counts.forEach((c, idx) => {
    if (idx) summaryRows.push({ type: 'separator', margin: 'md' });
    summaryRows.push({ type: 'box', layout: 'horizontal', margin: 'md', contents: [
      { type: 'box', layout: 'vertical', flex: 5, contents: [
        txt(c.d.label, { weight: 'bold', size: 'md', color: CARD.ink }),
        txt(c.all ? `หมดแล้ว ${c.exp} · ภายใน 30 วัน ${c.urg}` : 'ไม่มีรายการ',
          { size: 'xs', color: c.exp ? CARD.crit : c.urg ? CARD.high : CARD.faint }),
      ] },
      txt(c.all, { flex: 2, align: 'end', gravity: 'center', size: 'xxl', weight: 'bold', color: c.all ? CARD.ink : '#C0C6CF' }),
      txt('คน', { flex: 1, align: 'end', gravity: 'center', size: 'xs', color: CARD.faint }),
    ] });
  });
  summaryRows.push({ type: 'separator', margin: 'lg' });
  summaryRows.push(txt('🔴 หมดแล้ว  🟠 ≤30 วัน  🟡 31–60  🔵 61 วันขึ้นไป', { size: 'xxs', color: CARD.faint, margin: 'md', wrap: true }));
  const summary = { type: 'bubble', size: 'mega',
    header: { type: 'box', layout: 'vertical', backgroundColor: CARD.navy, paddingAll: '16px', contents: [
      txt('⚠️ เอกสารใกล้หมดอายุ', { color: '#FFFFFF', weight: 'bold', size: 'lg' }),
      txt(`ภายใน ${CONFIG.DAYS_AHEAD} วัน · ข้อมูล ณ ${thDate_(today)} · รวม ${items.length} รายการ`, { color: '#DCE6F5', size: 'xs', wrap: true }),
    ] },
    body: { type: 'box', layout: 'vertical', contents: summaryRows },
  };
  if (CONFIG.WEB_URL) summary.footer = { type: 'box', layout: 'vertical', contents: [
    { type: 'button', style: 'primary', color: CARD.navy, height: 'sm', action: { type: 'uri', label: 'เปิดหน้ารายละเอียด', uri: CONFIG.WEB_URL } },
  ] };

  // following cards: names per document type, ROWS_PER_CARD per card
  const bubbles = [summary];
  counts.filter(c => c.all).forEach(c => {
    const mine = items.filter(i => i.doc === c.d);
    const pages = Math.ceil(mine.length / ROWS_PER_CARD);
    for (let pg = 0; pg < pages; pg++) {
      const rows = [];
      mine.slice(pg * ROWS_PER_CARD, (pg + 1) * ROWS_PER_CARD).forEach((it, k) => {
        const p = it.p;
        let sub = [p.psid, p.dept, thDate_(it.date)].filter(Boolean).join(' · ');
        if (CONFIG.INCLUDE_DOC_NUMBERS) {
          const nums = [p.passportNo && 'PP ' + p.passportNo, p.alienId && 'ต่างด้าว ' + p.alienId, p.wpNo && 'WP ' + p.wpNo].filter(Boolean);
          if (nums.length) sub += '\n' + nums.join(' · ');
        }
        if (k) rows.push({ type: 'separator', margin: 'md' });
        rows.push({ type: 'box', layout: 'vertical', margin: 'md', contents: [
          { type: 'box', layout: 'horizontal', contents: [
            txt([p.prefix, p.nameTh].filter(Boolean).join(' '), { size: 'sm', weight: 'bold', color: CARD.ink, flex: 5, wrap: true }),
            txt(when(it.n), { size: 'sm', weight: 'bold', color: color(it.n), flex: 3, align: 'end' }),
          ] },
          txt(sub, { size: 'xxs', color: CARD.muted, wrap: true }),
        ] });
      });
      bubbles.push({ type: 'bubble', size: 'mega',
        header: { type: 'box', layout: 'vertical', backgroundColor: CARD.yellow, paddingAll: '14px', contents: [
          txt(`${c.d.label} หมดอายุ`, { weight: 'bold', size: 'md', color: CARD.ink }),
          txt(`${c.all} คน` + (pages > 1 ? ` · หน้า ${pg + 1}/${pages}` : ''), { size: 'xs', color: CARD.ink }),
        ] },
        body: { type: 'box', layout: 'vertical', contents: rows },
      });
    }
  });

  // pack bubbles into carousels: at most 12 bubbles and ~45 KB per LINE message
  const alt = '⚠️ เอกสารใกล้หมดอายุ: ' + counts.map(c => `${c.d.label} ${c.all}`).join(' · ');
  const out = []; let cur = [];
  const size = arr => Utilities.newBlob(JSON.stringify({ type: 'carousel', contents: arr })).getBytes().length;
  bubbles.forEach(b => {
    if (cur.length && (cur.length >= 12 || size(cur.concat([b])) > 45000)) { out.push(cur); cur = []; }
    cur.push(b);
  });
  if (cur.length) out.push(cur);
  return out.map((arr, i) => ({ type: 'flex', altText: i ? alt + ' (ต่อ)' : alt, contents: { type: 'carousel', contents: arr } }));
}

// ---------------- helpers ----------------
function norm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[\s​ ]+/g, ''); }
function findHeader_(rows) {
  for (let r = 0; r < Math.min(rows.length, 40); r++) if (rows[r].some(c => norm_(c).includes('peoplesoft'))) return r;
  return -1;
}
function mapColumns_(rows, r) {
  const map = {}, width = rows[r].length, used = () => Object.keys(map).map(k => map[k]);
  const pass = getText => MATCHERS.forEach(([key, fn]) => {
    if (key in map) return;
    for (let c = 0; c < width; c++) if (!used().includes(c) && fn(getText(c))) { map[key] = c; break; }
  });
  pass(c => norm_(rows[r][c]));
  if (r > 0) pass(c => norm_(rows[r - 1][c]) + norm_(rows[r][c]));
  if (r + 1 < rows.length) pass(c => norm_(rows[r][c]) + norm_(rows[r + 1][c]));
  return map;
}
function mk_(y, m, d) {
  if (y > 2400) y -= 543; // พ.ศ. → ค.ศ.
  if (y < 100) y += 2000;
  const t = new Date(Date.UTC(y, m - 1, d));
  return (t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d) ? [y, m, d] : null;
}
function parseDate_(v, dateToYmd) {
  if (v === null || v === undefined || v === '') return null;
  // +6h guards against dates stored as 23:59:xx of the day before
  if (Object.prototype.toString.call(v) === '[object Date]') { if (isNaN(v)) return null; const a = dateToYmd(new Date(v.getTime() + 6 * 3600e3)); return mk_(a[0], a[1], a[2]); }
  if (typeof v === 'number') {
    if (v < 60 || v > 2958465) return null;
    const u = new Date(Math.round((v - 25569) * 864e5));
    return mk_(u.getUTCFullYear(), u.getUTCMonth() + 1, u.getUTCDate());
  }
  const s = String(v).trim();
  if (!s || /^[-–—\s]+$/.test(s)) return null;
  let m;
  if ((m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/))) return mk_(+m[1], +m[2], +m[3]);
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/))) return mk_(+m[3], +m[2], +m[1]);
  if ((m = s.match(/^(\d{1,2})[\s\-\/]*([^\d\s\-\/]+)[\s\-\/,]*(\d{2,4})$/))) {
    const key = m[2].toLowerCase().replace(/\./g, '');
    const mo = MONTH_NAMES[key] || MONTH_NAMES[key.slice(0, 3)];
    if (mo) return mk_(+m[3], mo, +m[1]);
  }
  return null;
}
/** Time zone as a plain string; falls back to Bangkok when the sheet or script returns nothing usable. */
function tz_() {
  let tz = '';
  try { tz = String(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || '').trim(); } catch (e) {}
  if (!tz) { try { tz = String(Session.getScriptTimeZone() || '').trim(); } catch (e) {} }
  return tz || 'Asia/Bangkok';
}
function ymd_(d, tz) { return Utilities.formatDate(d, String(tz || 'Asia/Bangkok'), 'yyyy-M-d').split('-').map(Number); }
function thDate_(a) { return `${a[2]} ${TH_MONTHS[a[1] - 1]} ${a[0] + 543}`; }
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

// ---------------- LINE ----------------
function linePost_(path, body) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) throw new Error('ยังไม่ได้ตั้งค่า LINE token: เมนู แจ้งเตือน LINE > ตั้งค่า LINE token');
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + path, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token }, payload: JSON.stringify(body),
  });
  const code = res.getResponseCode();
  if (code !== 200) {
    const hint = code === 401 ? ' (token ไม่ถูกต้อง ตั้งค่าใหม่ที่เมนู)' : code === 429 ? ' (ส่งเกินโควตาข้อความของเดือนนี้)' : code === 400 ? ' (รหัสกลุ่มหรือข้อความไม่ถูกต้อง)' : '';
    throw new Error('LINE ตอบกลับ ' + code + hint + ': ' + res.getContentText());
  }
}

/** LINE webhook: ใช้ลงทะเบียนกลุ่ม LINE ที่จะรับแจ้งเตือน (พิมพ์ "ลงทะเบียน 123456" ในกลุ่ม) */
function doPost(e) {
  const props = PropertiesService.getScriptProperties();
  const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  (body.events || []).forEach(ev => {
    const src = ev.source || {};
    const text = (ev.type === 'message' && ev.message && ev.message.type === 'text') ? ev.message.text.trim() : '';
    const m = text.match(/^ลงทะเบียน\s*(\d{6})$/);
    if (!m || !ev.replyToken) return;
    const code = props.getProperty('REGISTER_CODE');
    const id = src.groupId || src.roomId || src.userId;
    let reply;
    if (code && m[1] === code && id) {
      props.setProperty('TARGET_ID', id);
      props.deleteProperty('REGISTER_CODE');
      reply = '✅ ลงทะเบียนแล้ว แจ้งเตือนเอกสารหมดอายุจะส่งเข้า' + (src.groupId ? 'กลุ่มนี้' : 'แชตนี้');
    } else {
      reply = '❌ รหัสไม่ถูกต้องหรือหมดอายุ ขอรหัสใหม่ที่ Google Sheet เมนู แจ้งเตือน LINE';
    }
    try { linePost_('reply', { replyToken: ev.replyToken, messages: [{ type: 'text', text: reply }] }); } catch (err) { console.error(err); }
  });
  return ContentService.createTextOutput('OK');
}

// ---------------- settings ----------------
function setToken() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('ตั้งค่า LINE token', 'วาง Channel access token (long-lived) จาก LINE Developers > Messaging API', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const t = r.getResponseText().trim();
  if (t.length < 50) { ui.alert('token สั้นเกินไป ตรวจสอบว่าคัดลอก Channel access token มาครบ'); return; }
  PropertiesService.getScriptProperties().setProperty('LINE_TOKEN', t);
  ui.alert('บันทึก token แล้ว');
}
function showRegisterCode() {
  const code = String(Math.floor(100000 + Math.random() * 900000));
  PropertiesService.getScriptProperties().setProperty('REGISTER_CODE', code);
  SpreadsheetApp.getUi().alert('ลงทะเบียนกลุ่ม LINE',
    `1. เชิญบอทเข้ากลุ่ม LINE ที่ต้องการ\n2. พิมพ์ในกลุ่มว่า:\n\nลงทะเบียน ${code}\n\nบอทจะตอบกลับเมื่อสำเร็จ (ต้องตั้ง Webhook ตามคู่มือก่อน)`, SpreadsheetApp.getUi().ButtonSet.OK);
}
function clearTarget() {
  PropertiesService.getScriptProperties().deleteProperty('TARGET_ID');
  SpreadsheetApp.getUi().alert('จากนี้จะส่งหาทุกคนที่แอดบอทเป็นเพื่อน');
}
function scheduleDaily() { schedule_(b => b.everyDays(1), 'ทุกวัน'); }
function scheduleWeekly() { schedule_(b => b.onWeekDay(ScriptApp.WeekDay.MONDAY), 'ทุกวันจันทร์'); }
function schedule_(fn, label) {
  removeTriggers_();
  const tz = tz_();
  fn(ScriptApp.newTrigger('sendAlertsScheduled').timeBased()).atHour(CONFIG.SEND_HOUR).inTimezone(tz).create();
  SpreadsheetApp.getUi().alert(`ตั้งแล้ว: ส่งอัตโนมัติ${label} ช่วง ${CONFIG.SEND_HOUR}:00–${CONFIG.SEND_HOUR + 1}:00 น. (${tz})`);
}
function unschedule() { removeTriggers_(); SpreadsheetApp.getUi().alert('ปิดการส่งอัตโนมัติแล้ว'); }
function removeTriggers_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendAlertsScheduled').forEach(t => ScriptApp.deleteTrigger(t));
}
function showStatus() {
  const p = PropertiesService.getScriptProperties();
  const trig = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendAlertsScheduled').length;
  const last = p.getProperty('LAST_SENT');
  SpreadsheetApp.getUi().alert('สถานะ',
    `LINE token: ${p.getProperty('LINE_TOKEN') ? 'ตั้งแล้ว' : 'ยังไม่ตั้ง'}\n` +
    `ผู้รับ: ${p.getProperty('TARGET_ID') ? 'กลุ่ม/แชตที่ลงทะเบียนไว้' : 'ทุกคนที่แอดบอทเป็นเพื่อน'}\n` +
    `ส่งอัตโนมัติ: ${trig ? 'เปิด' : 'ปิด'}\n` +
    `ส่งล่าสุด: ${last ? new Date(last).toLocaleString('th-TH') : '-'}\n` +
    `แจ้งล่วงหน้า: ${CONFIG.DAYS_AHEAD} วัน · ส่งเลขเอกสาร: ${CONFIG.INCLUDE_DOC_NUMBERS ? 'ใช่' : 'ไม่'}\n` +
    `รูปแบบ: ${CONFIG.MESSAGE_STYLE === 'card' ? 'การ์ดสรุป' : 'ข้อความ'}`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}
