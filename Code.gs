/**
 * Lalitha Event Decors – Estimate Maker: Google Sheets storage (v2)
 * All app data lives in this spreadsheet:
 *   Estimates  – one row per estimate/invoice (full data in the last column)
 *   Items      – every line item of every estimate (for your own analysis)
 *   Payments   – every payment received
 *   Clients    – client list (name, phone, address, GSTIN)
 *   Catalog    – your item catalog. You can edit this tab directly in Sheets.
 *   Packages   – saved packages
 *   Photos     – catalog photos and seal/signature image
 *   Settings   – business details, layout settings, number counters
 *
 * Setup / update:
 * 1. Extensions → Apps Script → paste this file (replace everything). Set SECRET.
 * 2. Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy
 *    (first time: Deploy → New deployment → Web app, Execute as Me, Access Anyone).
 */
const SECRET = 'change-this-password';
// Google Calendar: confirmed bookings are added to this calendar automatically.
const CALENDAR_SYNC = true;   // set to false to switch off
const CALENDAR_ID = '';       // '' = your main Google Calendar, or paste another calendar's ID

const CHUNK = 45000; // Google Sheets cell limit is 50,000 characters
const ss_ = () => SpreadsheetApp.getActiveSpreadsheet();
function sh_(name, head) {
  let s = ss_().getSheetByName(name);
  if (!s) { s = ss_().insertSheet(name); s.appendRow(head); s.setFrozenRows(1); s.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#f3eee0'); }
  return s;
}
const EST_H = ['ID', 'Number', 'Type', 'Date', 'Client', 'Phone', 'Event', 'Event date', 'Venue', 'Status', 'Grand total', 'Paid', 'Balance', 'Saved at', 'Deleted', 'Data (do not edit)'];
const est_ = () => sh_('Estimates', EST_H);
const items_ = () => sh_('Items', ['Estimate ID', 'Number', 'Event date', 'Status', 'Section', 'Item', 'Size', 'Qty', 'Days', 'Rate', 'Amount', 'Optional']);
const pays_ = () => sh_('Payments', ['Estimate ID', 'Number', 'Client', 'Date', 'Amount', 'Mode', 'Note']);
const clients_ = () => sh_('Clients', ['Client', 'Phone', 'Address', 'GSTIN', 'Contact person', 'Last estimate', 'Updated']);
const CAT_H = ['Item', 'Size', 'Rate', 'Description', 'Peak rate', 'Telugu name', 'Stock owned', 'My cost', 'Rate type'];
const cat_ = () => sh_('Catalog', CAT_H);
const pkg_ = () => sh_('Packages', ['Package', 'Items', 'Value', 'Data (do not edit)']);
const ph_ = () => sh_('Photos', ['Name', 'Updated', 'Image data (do not edit) →']);
const set_ = () => sh_('Settings', ['Key', 'Updated', 'Value (do not edit) →']);
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const props_ = () => PropertiesService.getScriptProperties();
const chunks_ = s => { s = String(s); const a = []; for (let i = 0; i < s.length; i += CHUNK) a.push(s.slice(i, i + CHUNK)); return a.length ? a : ['']; };

/* ---------- key/value rows split over several cells ---------- */
function kvRead_(sheet) {
  const v = sheet.getDataRange().getValues(), o = {};
  for (let i = 1; i < v.length; i++) if (v[i][0] !== '') o[v[i][0]] = { at: +v[i][1] || 0, raw: v[i].slice(2).join('') };
  return o;
}
function kvWrite_(sheet, key, str) {
  const v = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(r => r[0]);
  let i = v.indexOf(key), row;
  if (i > 0) { row = i + 1; const lc = sheet.getLastColumn(); if (lc > 2) sheet.getRange(row, 3, 1, lc - 2).clearContent(); }
  else { row = sheet.getLastRow() + 1; }
  const c = chunks_(str), now = Date.now();
  sheet.getRange(row, 1, 1, 2 + c.length).setValues([[key, now].concat(c)]);
  return now;
}
function kvDelete_(sheet, key) {
  const v = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(r => r[0]);
  const i = v.indexOf(key); if (i > 0) sheet.deleteRow(i + 1);
}

/* ---------- catalog as a normal table ---------- */
function catRead_() {
  const v = cat_().getDataRange().getValues().slice(1).filter(r => r[0] !== '');
  return v.map(r => [String(r[0]), String(r[1]), +r[2] || 0, String(r[3]), r[4] === '' ? '' : +r[4], String(r[5]), r[6] === '' ? '' : +r[6], r[7] === '' ? '' : +r[7], String(r[8])]);
}
function catWrite_(rows) {
  const s = cat_(); const n = s.getLastRow(); if (n > 1) s.getRange(2, 1, n - 1, CAT_H.length).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, CAT_H.length).setValues(rows.map(r => { r = (r || []).slice(0, 9); while (r.length < 9) r.push(''); return r; }));
  props_().setProperty('catalogAt', String(Date.now()));
}
/* Edits made directly in the Catalog or Packages tab are picked up by the app */
function onEdit(e) {
  const n = e && e.range && e.range.getSheet().getName();
  if (n === 'Catalog') props_().setProperty('catalogAt', String(Date.now()));
}

/* ---------- packages ---------- */
function pkgRead_() { return pkg_().getDataRange().getValues().slice(1).filter(r => r[0] !== '' && r[3]).map(r => JSON.parse(r[3])); }
function pkgWrite_(list) {
  const s = pkg_(); const n = s.getLastRow(); if (n > 1) s.getRange(2, 1, n - 1, 4).clearContent();
  const rows = (list || []).map(p => { const it = (p.rows || []).filter(r => !r.sec); const val = it.filter(r => !r.opt).reduce((a, r) => a + (+r.amount || (+r.qty || 0) * (+r.rate || 0) * (+r.days || 1)), 0); return [p.name, it.map(r => r.name).join(', '), val, JSON.stringify(p)]; });
  if (rows.length) s.getRange(2, 1, rows.length, 4).setValues(rows);
  props_().setProperty('packagesAt', String(Date.now()));
}

/* ---------- derived tables for an estimate ---------- */
function deleteRowsWhere_(sheet, id) {
  const v = sheet.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === id) sheet.deleteRow(i + 1);
}
function writeDerived_(id, d) {
  const it = items_(), py = pays_();
  deleteRowsWhere_(it, id); deleteRowsWhere_(py, id);
  if (!d) return;
  let sec = ''; const rows = [];
  (d.rows || []).forEach(r => { if (r.sec) { sec = r.name; return; } rows.push([id, d.no || '', d.edate || '', d.status || '', sec, r.name || '', r.uom || '', +r.qty || 0, +r.days || 1, +r.rate || 0, +r.amount || 0, r.opt ? 'yes' : '']); });
  if (rows.length) it.getRange(it.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  const p = (d.pays || []).filter(x => +x.amt).map(x => [id, d.no || '', d.client || '', x.date || '', +x.amt || 0, x.mode || '', x.note || '']);
  if (p.length) py.getRange(py.getLastRow() + 1, 1, p.length, p[0].length).setValues(p);
  if (d.client) {
    const c = clients_(), v = c.getDataRange().getValues(); let i = -1;
    for (let k = 1; k < v.length; k++) if (String(v[k][0]).trim().toLowerCase() === d.client.trim().toLowerCase()) { i = k; break; }
    const row = [d.client.trim(), d.phone || '', d.addr || '', d.cgst || '', d.attn || '', d.no || '', new Date()];
    if (i > 0) c.getRange(i + 1, 1, 1, row.length).setValues([row]); else c.appendRow(row);
  }
}

/* ---------- read everything ---------- */
function doGet(e) {
  ensureV2_();
  if (e.parameter.view) return clientView_(e.parameter.view, e.parameter.t);
  if (e.parameter.card !== undefined) return out_(cardData_(String(e.parameter.card || '')));
  if (e.parameter.key !== SECRET) return out_({ ok: false, error: 'Wrong password' });
  const q = {};
  est_().getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    if (r[14]) { q[r[0]] = { deleted: true, saved: +r[13] }; return; }
    try { q[r[0]] = JSON.parse(r[15]); } catch (x) { }
  });
  const st = {};
  const kv = kvRead_(set_());
  for (const k in kv) { try { st[k] = { value: JSON.parse(kv[k].raw), at: kv[k].at }; } catch (x) { } }
  const ph = kvRead_(ph_());
  for (const k in ph) st[k === '__stamp' ? 'stamp' : 'ph:' + k] = { value: ph[k].raw, at: ph[k].at };
  const P = props_();
  st.catalog = { value: catRead_(), at: +P.getProperty('catalogAt') || 1 };
  st.packages = { value: pkgRead_(), at: +P.getProperty('packagesAt') || 1 };
  return out_({ ok: true, estimates: q, settings: st, sheet: ss_().getUrl(), name: ss_().getName() });
}

/* ---------- write ---------- */
function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.key !== SECRET) return out_({ ok: false, error: 'Wrong password' });
  ensureV2_();
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    if (b.action === 'save' || b.action === 'delete') {
      const s = est_(), ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => r[0]);
      const d = b.data || {}, i = ids.indexOf(b.id), paid = (+d.adv || 0) + (d.pays || []).reduce((a, p) => a + (+p.amt || 0), 0);
      const row = b.action === 'delete'
        ? [b.id, '', '', '', '', '', '', '', '', '', '', '', '', Date.now(), 'yes', '']
        : [b.id, d.no || '', d.doctype || '', d.date || '', d.client || '', d.phone || '', d.ename || d.etype || '', d.edate || '', d.venue || '', d.status || 'Draft', +d.total || 0, paid, (+d.total || 0) - paid, d.saved || Date.now(), '', JSON.stringify(d)];
      if (i > 0) s.getRange(i + 1, 1, 1, row.length).setValues([row]); else s.appendRow(row);
      writeDerived_(b.id, b.action === 'delete' ? null : d);
      try { calSync_(b.id, b.action === 'delete' ? null : d); } catch (x) { }
    } else if (b.action === 'calsyncall') {
      return out_({ ok: true, count: calSyncAll_() });
    } else if (b.action === 'setting') {
      const n = b.name;
      if (n === 'catalog') catWrite_(b.value || []);
      else if (n === 'packages') pkgWrite_(b.value || []);
      else if (n === 'stamp' || n.indexOf('ph:') === 0) {
        const key = n === 'stamp' ? '__stamp' : n.slice(3);
        if (b.value) kvWrite_(ph_(), key, b.value); else kvDelete_(ph_(), key);
      } else if (n === 'seq' || n === 'iseq') {
        const kv = kvRead_(set_()), cur = kv[n] ? +JSON.parse(kv[n].raw) : 0;
        if (+b.value > cur) kvWrite_(set_(), n, JSON.stringify(+b.value));
      } else kvWrite_(set_(), n, JSON.stringify(b.value));
    }
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

/* ---------- old v1 layout: rename old tabs, the app re-uploads its data ---------- */
function ensureV2_() {
  const st = ss_().getSheetByName('Settings');
  if (st && st.getRange(1, 2).getValue() === 'Value (do not edit)') st.setName('Settings_old_' + Date.now());
  const e = ss_().getSheetByName('Estimates');
  if (e && e.getRange(1, 2).getValue() === 'Estimate No') e.setName('Estimates_old_' + Date.now());
}

/* ---------- Client accept page ---------- */
function findEst_(id) {
  const s = est_(), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === id && !v[i][14] && v[i][15]) return { sheet: s, row: i + 1, data: JSON.parse(v[i][15]) };
  return null;
}
const inr_ = n => '₹' + Math.round(+n || 0).toLocaleString('en-IN');
const h_ = x => String(x == null ? '' : x).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
function clientView_(id, t) {
  const f = findEst_(id);
  if (!f || !f.data.tok || f.data.tok !== t) return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:20px">This link is not valid.</p>');
  const d = f.data, done = ['Confirmed', 'Completed'].includes(d.status);
  let n = 0;
  const rows = (d.rows || []).map(r => r.sec ? `<tr><td colspan="3" class="s">${h_(r.name)}</td></tr>` :
    `<tr class="${r.opt ? 'o' : ''}"><td>${++n}. ${h_(r.name)}${r.opt ? ' <i>(optional)</i>' : ''}<br><small>${h_(r.uom || '')}</small></td><td>${r.qty}</td><td>${inr_(r.amount)}</td></tr>`).join('');
  const btn = done ? '<div class="ok">✓ Accepted. Thank you!</div>'
    : `<button id="b">✓ Accept this ${h_(String(d.doctype || 'quotation').toLowerCase())}</button><p><small>Tapping accept confirms the booking as per the terms shared with you.</small></p>`;
  const html = `<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><style>
body{font-family:system-ui,sans-serif;margin:0;background:#f4f1ec;color:#222}.w{max-width:560px;margin:0 auto;padding:16px}.c{background:#fff;border-radius:12px;padding:16px;margin-bottom:12px}
h1{font-size:20px;margin:0 0 4px;color:#d9161c}table{width:100%;border-collapse:collapse}td{padding:6px 4px;border-bottom:1px solid #eee;vertical-align:top}td:nth-child(2),td:nth-child(3){text-align:right;white-space:nowrap}
.s{font-weight:700;color:#d9161c;text-transform:uppercase;font-size:12px}.o{color:#999}.t{display:flex;justify-content:space-between;font-size:20px;font-weight:700;margin-top:10px}
button{width:100%;padding:14px;font-size:17px;border:0;border-radius:10px;background:#1d7a3a;color:#fff;font-weight:700}.ok{background:#e2f5e7;color:#1d7a3a;padding:14px;border-radius:10px;text-align:center;font-weight:700}
small{color:#888}</style></head><body><div class="w"><div class="c"><h1>Lalitha Event Decors</h1><div>${h_(d.doctype)} <b>${h_(d.no)}</b></div>
<div>For: <b>${h_(d.client)}</b></div><div>${h_(d.ename || d.etype || '')} ${d.edate ? '· ' + h_(d.edate.split('-').reverse().join('-')) : ''} ${d.venue ? '· ' + h_(d.venue) : ''}</div></div>
<div class="c"><table>${rows}</table><div class="t"><span>Total</span><span>${inr_(d.total)}</span></div><small>Optional items are not included in the total.</small></div>
<div class="c" id="act">${btn}</div></div>
<script>var b=document.getElementById('b');if(b)b.onclick=function(){b.disabled=true;b.textContent='Please wait…';google.script.run.withSuccessHandler(function(){document.getElementById('act').innerHTML='<div class=ok>✓ Accepted. Thank you! We will contact you shortly.</div>'}).withFailureHandler(function(){b.disabled=false;b.textContent='Try again'}).accept(${JSON.stringify(id)},${JSON.stringify(t)})};</script></body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('Lalitha Event Decors – ' + d.no).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function accept(id, t) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = findEst_(id); if (!f || f.data.tok !== t) throw new Error('Invalid link');
    f.data.status = 'Confirmed'; f.data.follow = ''; f.data.acceptedAt = new Date().toISOString(); f.data.saved = Date.now();
    f.sheet.getRange(f.row, 10).setValue('Confirmed'); f.sheet.getRange(f.row, 14).setValue(f.data.saved); f.sheet.getRange(f.row, 16).setValue(JSON.stringify(f.data));
    writeDerived_(id, f.data);
    try { calSync_(id, f.data); } catch (x) { }
    return true;
  } finally { lock.releaseLock(); }
}

/* ---------- Google Calendar ---------- */
function cal_() { return CALENDAR_ID ? CalendarApp.getCalendarById(CALENDAR_ID) : CalendarApp.getDefaultCalendar(); }
function calGuests_() {
  try { const kv = kvRead_(set_()); const c = kv.config ? JSON.parse(kv.config.raw) : {}; return String(c.calGuests || '').split(/[\s,;]+/).map(x => x.trim().toLowerCase()).filter(x => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x)); } catch (x) { return []; }
}
function calSync_(id, d) {
  if (!CALENDAR_SYNC) return;
  const guests = calGuests_();
  const P = props_(), key = 'cal_' + id, evId = P.getProperty(key), cal = cal_();
  let ev = null; if (evId) { try { ev = cal.getEventById(evId); } catch (x) { ev = null; } }
  const want = d && d.edate && ['Confirmed', 'Completed'].includes(d.status);
  if (!want) { if (ev) ev.deleteEvent(); P.deleteProperty(key); return; }
  const [y, m, dd] = d.edate.split('-').map(Number);
  const title = '🎉 ' + (d.client || 'Event') + ' – ' + (d.ename || d.etype || 'Decoration') + ' (' + (d.no || '') + ')';
  const paid = (+d.adv || 0) + (d.pays || []).reduce((a, p) => a + (+p.amt || 0), 0);
  let n = 0;
  const items = (d.rows || []).filter(r => !r.opt).map(r => r.sec ? '\n' + String(r.name).toUpperCase() : (++n) + '. ' + r.name + (r.uom ? ' (' + r.uom + ')' : '') + ' × ' + r.qty).join('\n');
  const desc = [(d.doctype || 'Quotation') + ' ' + (d.no || ''), 'Client: ' + (d.client || '') + (d.phone ? ' · ' + d.phone : ''), d.attn ? 'Contact: ' + d.attn : '', d.guests ? 'Guests: ' + d.guests : '',
    d.setup ? 'Setup: ' + d.setup : '', d.team ? 'Team: ' + d.team : '', d.vehicle ? 'Vehicle: ' + d.vehicle : '',
    'Total: ' + inr_(d.total) + ' · Paid: ' + inr_(paid) + ' · Balance: ' + inr_((+d.total || 0) - paid), '', 'Items:' + '\n' + items, d.notes ? '\nNotes: ' + d.notes : ''].filter(x => x !== '').join('\n');
  let start, end, allDay = !d.etime;
  if (!allDay) { const [h, mi] = d.etime.split(':').map(Number); start = new Date(y, m - 1, dd, h, mi); end = new Date(start.getTime() + 4 * 3600 * 1000); }
  else start = new Date(y, m - 1, dd);
  if (ev && ev.isAllDayEvent() !== allDay) { ev.deleteEvent(); ev = null; }
  if (!ev) {
    const opt = { location: d.venue || '', description: desc, guests: guests.join(','), sendInvites: guests.length > 0 };
    ev = allDay ? cal.createAllDayEvent(title, start, opt) : cal.createEvent(title, start, end, opt);
    ev.addPopupReminder(24 * 60); ev.addPopupReminder(3 * 60);
    P.setProperty(key, ev.getId());
  } else {
    ev.setTitle(title); ev.setLocation(d.venue || ''); ev.setDescription(desc);
    if (allDay) ev.setAllDayDate(start); else ev.setTime(start, end);
  }
  try { const have = ev.getGuestList().map(g => g.getEmail().toLowerCase()); guests.forEach(g => { if (have.indexOf(g) < 0) ev.addGuest(g); }); have.forEach(g => { if (guests.indexOf(g) < 0) ev.removeGuest(g); }); } catch (x) { }
  try { ev.setColor(d.status === 'Completed' ? CalendarApp.EventColor.GRAY : CalendarApp.EventColor.RED); } catch (x) { }
}
function calSyncAll_() {
  let c = 0;
  est_().getDataRange().getValues().slice(1).forEach(r => { if (!r[0] || r[14] || !r[15]) return; try { const d = JSON.parse(r[15]); calSync_(r[0], d); if (d.edate && ['Confirmed', 'Completed'].includes(d.status)) c++; } catch (x) { } });
  return c;
}
/* Run this once from the Apps Script editor to give calendar permission */
function authorizeCalendar() { return cal_().getName(); }

/* ---------- Public visiting card data (no password; only card-safe fields) ---------- */
function cardData_(tok) {
  const kv = kvRead_(set_()), J = k => { try { return kv[k] ? JSON.parse(kv[k].raw) : {}; } catch (x) { return {}; } };
  const c = J('config'), b = J('business');
  const pick = (a, d) => (a !== undefined && a !== '' ? a : d || '');
  const card = {
    style: pick(c.cardStyle, 'glass'), name: pick(c.bizName, 'Lalitha Event Decors'), tagline: pick(b.btag, ''),
    owner: pick(c.cardOwner), role: pick(c.cardRole), phone: pick(c.cardPhone, b.bphone), wa: pick(c.cardWa, c.cardPhone || b.bphone),
    email: pick(c.cardEmail, b.bmail), addr: pick(c.cardAddr, b.baddr), maps: pick(c.cardMaps), web: pick(c.cardWeb, b.bweb),
    insta: pick(c.cardInsta), fb: pick(c.cardFb), yt: pick(c.cardYt), review: pick(c.reviewLink),
    services: String(c.cardServices || '').split('\n').map(x => x.trim()).filter(Boolean),
    about: pick(c.cardAbout, c.about), greet: pick(c.cardGreet, 'Dear {name}, thank you for your enquiry. We would love to decorate your {event}.'),
    upi: c.cardUpi ? pick(b.bupi) : '', since: pick(c.cardSince, '1985'), halls: Array.isArray(c.halls) ? c.halls.filter(h => h && h.name) : [], acc: pick(c.acc, '#d9161c'), gold: pick(c.gold, '#9a7a1c'), photos: []
  };
  const names = String(c.cardGallery || '').split(',').map(x => x.trim()).filter(Boolean).slice(0, 9);
  if (names.length) { const ph = kvRead_(ph_()); names.forEach(n => { if (ph[n]) card.photos.push({ name: n, src: ph[n].raw }); }); }
  let enq = null;
  if (tok && tok.length >= 4) {
    const v = est_().getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (v[i][14] || !v[i][15] || String(v[i][15]).indexOf(tok) < 0) continue;
      try {
        const d = JSON.parse(v[i][15]); if (d.ctok !== tok) continue;
        enq = { no: d.no || '', client: d.client || '', event: d.ename || d.etype || '', etype: d.etype || '', edate: d.edate || '', etime: d.etime || '', venue: d.venue || '', guests: d.guests || '', notes: c.cardNotes ? (d.notes || '') : '', halls: Array.isArray(d.hallq) ? d.hallq : [] };
      } catch (x) { }
      break;
    }
  }
  return { ok: true, card: card, enq: enq };
}
