// Runs apps-script/Code.gs against an in-memory Google Sheets mock.
// Simulates a week of submissions and prints the weekly tab + email.
// Usage: node tools/sheet_mock_test.js [out-dir]
const fs = require('fs'), path = require('path');
const out = process.argv[2] || path.join(__dirname);

function makeSheet(name, id) {
  const cells = []; // cells[r][c] = {v, f}
  const get = (r, c) => ((cells[r] || [])[c]) || { v: '', f: '' };
  const set = (r, c, o) => { (cells[r] = cells[r] || [])[c] = Object.assign({ v: '', f: '' }, o); };
  const noop = new Proxy({}, { get: (t, k) => () => noop });
  const sh = {
    name, id,
    getName: () => name, getSheetId: () => id,
    getLastRow: () => { let n = 0; cells.forEach((row, r) => { if (row && row.some(x => x && (x.v !== '' || x.f !== ''))) n = r; }); return n; },
    getRange(r, c, nr = 1, nc = 1) {
      const rng = {
        setValues(vals) { vals.forEach((row, i) => row.forEach((v, j) => { const md = typeof v === 'string' && /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun) (\d+)\/(\d+)$/.exec(v); set(r + i, c + j, { v: md ? new Date(2026, +md[2] - 1, +md[3]) : v }); })); return rng; }, // HEADER_AS_DATE: real Sheets turns 'Mon 9/21' into a date
        setValue(v) { set(r, c, { v }); return rng; },
        setFormulas(fs) { fs.forEach((row, i) => row.forEach((f, j) => set(r + i, c + j, String(f).startsWith('=') ? { f } : { v: f }))); return rng; },
        setFormula(f) { set(r, c, { f }); return rng; },
        getValue() { return get(r, c).v; },
        getValues() { const a = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) row.push(get(r + i, c + j).v); a.push(row); } return a; },
      };
      const px = new Proxy(rng, { get: (t, k) => (k in t ? (...a) => { const res = t[k](...a); return res === rng ? px : res; } : () => px) });
      return px;
    },
    insertRowBefore(r) { cells.splice(r, 0, []); },
    insertRowsBefore(r, n) { cells.splice(r, 0, ...Array.from({ length: n }, () => [])); },
    appendRow(vals) { const r = sh.getLastRow() + 1; vals.forEach((v, j) => set(r, j + 1, { v })); },
    setFrozenRows() {}, setColumnWidth() {}, setConditionalFormatRules(r) { sh.rules = r; },
    dump() { const n = sh.getLastRow(); const lines = []; for (let r = 1; r <= n; r++) { const row = []; for (let c = 1; c <= 13; c++) { const x = get(r, c); row.push(x.f ? x.f : x.v); } lines.push(r + '\t' + row.map(v => String(v).slice(0, 22)).join(' | ')); } return lines.join('\n'); },
  };
  return sh;
}
const sheets = []; let nextId = 100;
const ss = {
  getSheetByName: n => sheets.find(s => s.name === n) || null,
  insertSheet: (n, idx) => { const s = makeSheet(n, nextId++); idx === 0 ? sheets.unshift(s) : sheets.push(s); return s; },
  getUrl: () => 'https://docs.google.com/spreadsheets/d/TEST',
  getSheets: () => sheets,
};
const ruleBuilder = () => { const b = new Proxy({}, { get: (t, k) => k === 'build' ? () => ({}) : () => b }); return b; };
global.SpreadsheetApp = { getActiveSpreadsheet: () => ss, newConditionalFormatRule: ruleBuilder };
global.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
const mails = []; global.MailApp = { sendEmail: o => mails.push(o) };
global.Utilities = { formatDate: (d, tz, f) => f === 'EEE M/d' ? d.toDateString().slice(0, 3) + ' ' + (d.getMonth() + 1) + '/' + d.getDate() : d.toDateString() };
global.ContentService = { createTextOutput: t => ({ setMimeType: () => t }), MimeType: { JSON: 'json' } };

eval(fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'Code.gs'), 'utf8') + '\n;global.doPost = doPost; global.SHARED_KEY = SHARED_KEY; global.sendWeeklyReportNow = sendWeeklyReportNow;');

const ITEMS = {
  'Pechanga Fried Chicken': [['Chicken', 'Legs', 'ea.', 0.517], ['Chicken', 'Biscuits', 'ea.', 0.3], ['Panned Sides', 'Corn', 'Lbs.', 1.34]],
  'Pronto': [['Pizza & Mains', 'Pepperoni Pizza', 'slice', 10.75], ['Pizza & Mains', 'Lasagna', 'pan', 2.87]],
  'Little Wok': [['Entrees', 'Orange Chicken', 'LBS', 4], ['Sides', 'Pork Eggroll', 'Ea.', 0.99]],
  'Agave': [['Proteins', 'Carne Asada', 'Lbs.', 8.6]],
};
function payload(date, id, qtyFn, extraItem) {
  const stations = Object.entries(ITEMS).map(([label, items], si) => {
    const list = items.slice(); if (extraItem && label === 'Agave') list.push(extraItem);
    const its = list.map(([category, item, uom, c], ii) => { const qty = qtyFn(si, ii); return { category, item, uom, qty, costPerUnit: c, cost: Math.round(qty * c * 100) / 100 }; });
    const w = its.filter(i => i.qty > 0);
    return { station: label, stationLabel: label, closeTime: '21:30', notes: '', totals: { itemsWasted: w.length, units: w.reduce((a, i) => a + i.qty, 0), cost: Math.round(w.reduce((a, i) => a + i.cost, 0) * 100) / 100 }, items: its };
  });
  const sum = k => Math.round(stations.reduce((a, s) => a + s.totals[k], 0) * 100) / 100;
  return { id, key: SHARED_KEY, chef: 'Test Chef', badge: '123', date, submittedAt: new Date().toISOString(), totals: { itemsWasted: sum('itemsWasted'), units: sum('units'), cost: sum('cost') }, stations };
}
const post = p => JSON.parse(doPost({ postData: { contents: JSON.stringify(p) } }));
const res = [];
res.push(post(payload('2026-09-21', 'a', (s, i) => s + i)));                  // Mon
const pb = payload('2026-09-23', 'b', (s, i) => (s + 1) * 3); pb.stations[1].notes = 'Oven 2 down, pizzas remade'; res.push(post(pb));           // Wed
res.push(post(payload('2026-09-23', 'c', (s, i) => 1)));                     // Wed re-send (correction) overwrites
res.push(post(payload('2026-09-23', 'c', (s, i) => 99)));                    // duplicate id ignored
res.push(post(payload('2026-09-27', 'd', (s, i) => 12, ['Proteins', 'Birria', 'Lbs.', 5]))); // Sun + new menu item
res.push(post(payload('2026-09-27', 'd2', (s, i) => 11)));                   // Sun re-send → UPDATED weekly
res.push(post(payload('2026-09-28', 'e', (s, i) => 2)));                     // next Monday → new tab, no email
res.push(post(payload('2026-10-06', 'f', (s, i) => 1)));                     // week 2 had no Sunday → late weekly for 9/28 week
console.log('responses', JSON.stringify(res));
console.log('tabs', sheets.map(s => s.name).join(' , '));
const wk = ss.getSheetByName('Week 9-21 to 9-27');
console.log(wk.dump());
mails.forEach((x, i) => console.log('email', i + 1, '|', x.subject));
fs.writeFileSync(path.join(out, 'mock_email.html'), mails[0].htmlBody);
fs.writeFileSync(path.join(out, 'mock_email_late.html'), mails[mails.length - 1].htmlBody);
// Manual "send last week now" must pick the most recent FINISHED week, not the newest tab
global.sendWeeklyReportNow = sendWeeklyReportNow;
const before = mails.length; sendWeeklyReportNow();
console.log('sendWeeklyReportNow ->', mails.length > before ? mails[mails.length - 1].subject : '(nothing sent)');
