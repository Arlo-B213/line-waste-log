/**
 * Food Court Line Waste Log — email + logging service
 * ---------------------------------------------------
 * Receives ONE nightly report covering every outlet, emails it to
 * management, and logs it to this Google Sheet.
 *
 * Paste into Extensions ▸ Apps Script of a Google Sheet, fill in the
 * settings below, then Deploy ▸ New deployment ▸ Web app
 * (Execute as: Me · Who has access: Anyone). Copy the /exec URL into the app.
 */

// ====== SETTINGS ======
const RECIPIENTS  = 'abedolla@pechanga.com'; // management inbox(es), comma-separated
const CC          = '';                                           // optional
const SHARED_KEY  = 'pechanga-fc-waste';                          // must match the app
const AMBER_AT = 5, RED_AT = 10;                                  // template status key
const TIMEZONE = 'America/Los_Angeles';
const DAILY_EMAIL  = false;  // true = also email every night; false = weekly report only
// ======================

const LOG_SHEET = 'Waste Log';
const SUB_SHEET = 'Submissions';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.key !== SHARED_KEY) return json_({ ok: false, error: 'Invalid key' });

    if (p.test) {
      MailApp.sendEmail({ to: RECIPIENTS, subject: '✅ Line Waste Log — test email',
        htmlBody: '<p>The Line Waste Log app is connected. The weekly waste report will arrive here after each Sunday count.</p>' });
      return json_({ ok: true, test: true });
    }
    if (!p.stations || !p.stations.length) return json_({ ok: false, error: 'No outlets in report' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Day / Swing logs are kept apart from the closing count: own tab, own email section, never in the week tab totals.
    if ((p.shift || 'closing') !== 'closing') return json_(logShift_(ss, p));
    const subs = sheet_(ss, SUB_SHEET, ['Submission ID','Submitted At','Business Date','Station','Closing Time','Chef','Badge #','Items Wasted','Total Units','Waste Cost ($)','Notes']);

    // Ignore duplicates (the app retries when the tablet was offline)
    const ids = subs.getLastRow() > 1 ? subs.getRange(2, 1, subs.getLastRow() - 1, 1).getValues().flat() : [];
    if (ids.indexOf(p.id) !== -1) return json_({ ok: true, duplicate: true });

    const submitted = new Date(p.submittedAt);
    const subRows = p.stations.map(s => [p.id, submitted, p.date, s.stationLabel, s.closeTime, p.chef, "'" + p.badge,
      s.totals.itemsWasted, s.totals.units, s.totals.cost, s.notes || '']);
    subs.getRange(subs.getLastRow() + 1, 1, subRows.length, subRows[0].length).setValues(subRows);

    const log = sheet_(ss, LOG_SHEET, ['Business Date','Station','Category','Item','UOM','Waste Qty','Cost/Unit ($)','Waste Cost ($)','Chef','Badge #','Submission ID']);
    const rows = [];
    p.stations.forEach(s => s.items.filter(i => i.qty > 0).forEach(i =>
      rows.push([p.date, s.stationLabel, i.category, i.item, i.uom, i.qty, i.costPerUnit, i.cost, p.chef, "'" + p.badge, p.id])));
    if (rows.length) log.getRange(log.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    const week = updateWeek_(ss, p);
    if (DAILY_EMAIL) sendReport_(p, ss.getUrl(), week);
    // Weekly report: goes out when Sunday's count comes in (week complete).
    if (week.idx === 6) sendWeeklyReport_(ss, week.sheet, false);
    // Safety net: if Sunday was never submitted, send last week's report
    // with the first count of the new week.
    const prev = ss.getSheetByName(weekName_(new Date(week.mon.getFullYear(), week.mon.getMonth(), week.mon.getDate() - 7)));
    if (prev && !weeklySent_(prev)) sendWeeklyReport_(ss, prev, true);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

const VERSION = 'shifts-1';  // bump when the app depends on new backend behavior; checked before app releases
function doGet() { return json_({ ok: true, service: 'Line Waste Log', version: VERSION }); }

/* =====================================================================
   SHIFT LOG — optional Day / Swing waste, one row per discarded item.
   ===================================================================== */
const SHIFT_SHEET = 'Shift Log';
const SHIFT_LABEL = { day: 'Day', swing: 'Swing' };

function logShift_(ss, p) {
  const sh = sheet_(ss, SHIFT_SHEET, ['Submission ID','Submitted At','Business Date','Shift','Station','Category','Item','UOM','Waste Qty','Cost/Unit ($)','Waste Cost ($)','Time','Chef','Badge #','Notes']);
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat() : [];
  if (ids.indexOf(p.id) !== -1) return { ok: true, duplicate: true };
  const label = SHIFT_LABEL[p.shift] || String(p.shift);
  const submitted = new Date(p.submittedAt);
  const rows = [];
  p.stations.forEach(s => s.items.filter(i => i.qty > 0).forEach(i =>
    rows.push([p.id, submitted, p.date, label, s.stationLabel, i.category, i.item, i.uom, i.qty, i.costPerUnit, i.cost,
      hm12_(s.closeTime), p.chef, "'" + p.badge, s.notes || ''])));
  if (!rows.length) return { ok: false, error: 'No waste in shift log' };
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { ok: true, shift: label };
}

function weekShifts_(ss, sh) {
  const log = ss.getSheetByName(SHIFT_SHEET); if (!log || log.getLastRow() < 2) return [];
  const wk = weekBounds_(sh); if (!wk) return [];
  return log.getRange(2, 1, log.getLastRow() - 1, 15).getValues()
    .map(r => ({ d: dateOnly_(r[2]), shift: String(r[3]), station: String(r[4]), item: String(r[6]), uom: String(r[7]), qty: Number(r[8]) || 0, cost: Number(r[10]) || 0 }))
    .filter(x => x.d && x.d >= wk.mon && x.d <= wk.sun && x.qty > 0);
}

/* =====================================================================
   WEEKLY LOG (Mon–Sun) — one tab per week, laid out like the Excel
   template: items down the side, Mon…Sun across, filled in nightly.
   Re-sending a night overwrites that day's column.
   Columns: A Category · B Item · C UOM · D Cost/Unit · E–K Mon–Sun ·
            L Week Qty · M Week Cost ($)
   ===================================================================== */
const DAY_COL = 5, NCOLS = 13;
const WK_NAVY = '#0f1b33', WK_GOLD = '#d4af37';

function mondayOf_(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const dt = new Date(y, m - 1, d);
  const idx = (dt.getDay() + 6) % 7;          // Mon=0 … Sun=6
  const mon = new Date(y, m - 1, d - idx);
  return { mon, idx };
}
function md_(dt) { return (dt.getMonth() + 1) + '/' + dt.getDate(); }
function weekName_(mon) {
  const sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 6);
  return 'Week ' + (mon.getMonth() + 1) + '-' + mon.getDate() + ' to ' + (sun.getMonth() + 1) + '-' + sun.getDate();
}
function dayLabels_(mon) {
  const n = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  return n.map((x, i) => x + ' ' + md_(new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + i)));
}

function updateWeek_(ss, p) {
  const { mon, idx } = mondayOf_(p.date);
  const name = weekName_(mon), days = dayLabels_(mon);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = buildWeekSheet_(ss, name, mon, days, p.stations);

  const col = DAY_COL + idx;
  p.stations.forEach(s => {
    let b = findBlock_(sh, s.stationLabel);
    if (!b) { appendBlock_(sh, s, days); b = findBlock_(sh, s.stationLabel); }
    // add any items that aren't on this week's tab yet (menu changes)
    s.items.forEach(i => {
      if (b.items[i.item] === undefined) {
        sh.insertRowBefore(b.costRow);
        const r = b.costRow;
        sh.getRange(r, 1, 1, 4).setValues([[i.category, i.item, i.uom, i.costPerUnit]]);
        sh.getRange(r, 12, 1, 2).setFormulas([['=SUM(E' + r + ':K' + r + ')', '=L' + r + '*D' + r]]);
        b = findBlock_(sh, s.stationLabel);
      }
    });
    // write tonight's quantities into the day column (0 shows as –)
    const vals = [];
    for (let r = b.first; r <= b.last; r++) vals.push(['']);
    s.items.forEach(i => { vals[b.items[i.item] - b.first] = [i.qty]; });
    sh.getRange(b.first, col, vals.length, 1).setValues(vals);
    sh.getRange(b.closeRow, col).setValue(hm12_(s.closeTime));
    sh.getRange(b.chefRow, col).setValue(p.chef + ' (' + p.badge + ')');
    writeBlockFormulas_(sh, b);
  });
  writeSummary_(sh);
  applyRules_(sh);
  sh.getRange(2, 1).setValue('Last updated ' + Utilities.formatDate(new Date(), TIMEZONE, 'EEE M/d h:mm a') +
    ' · ' + p.chef + '  ·  0 = counted, no waste (shows as –)  ·  🟠 5–9  🔴 10+');

  // week-to-date $ per outlet per day, computed from the tab's values
  const stationDaily = {};
  const all = sh.getRange(1, 1, sh.getLastRow(), NCOLS).getValues();
  let cur = null;
  all.forEach(row => {
    const a = String(row[0]);
    if (a.indexOf('■ ') === 0) { cur = a.slice(2); stationDaily[cur] = [0, 0, 0, 0, 0, 0, 0]; return; }
    if (!cur || row[1] === '' || row[1] === 'Item' || typeof row[3] !== 'number') return;
    for (let d = 0; d < 7; d++) { const q = row[DAY_COL - 1 + d]; if (typeof q === 'number') stationDaily[cur][d] += q * row[3]; }
  });
  return { name, gid: sh.getSheetId(), days, idx, stationDaily, sheet: sh, mon };
}

function buildWeekSheet_(ss, name, mon, days, stations) {
  const sh = ss.insertSheet(name, 0);                          // newest week first
  const sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 6);
  sh.getRange(1, 1).setValue('LINE WASTE LOG  ·  Week of Mon ' + md_(mon) + ' – Sun ' + md_(sun) + '/' + sun.getFullYear());
  sh.getRange(1, 1, 1, NCOLS).setBackground(WK_NAVY).setFontColor(WK_GOLD).setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1, 1, NCOLS).setFontColor('#777777').setFontStyle('italic');
  // summary block: rows 4–(5+stations)
  sh.getRange(4, 1, 1, NCOLS).setValues([['WEEK SUMMARY', 'Outlet', '', ''].concat(days, ['', 'Week Cost ($)'])])
    .setBackground(WK_NAVY).setFontColor('#ffffff').setFontWeight('bold');
  let r = 6 + stations.length + 2;
  sh.getRange(5 + stations.length, 1, 1, NCOLS).setFontWeight('bold').setBorder(true, false, false, false, false, false);
  stations.forEach(s => { r = writeBlockSkeleton_(sh, r, s, days) + 2; });
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 170); sh.setColumnWidth(3, 55); sh.setColumnWidth(4, 75);
  for (let c = DAY_COL; c < DAY_COL + 7; c++) sh.setColumnWidth(c, 78);
  sh.setColumnWidth(12, 80); sh.setColumnWidth(13, 100);
  return sh;
}

// highlight 5–9 amber / 10+ red — item rows only (not the $ rows)
function applyRules_(sh) {
  const v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  const ranges = [];
  v.forEach(r => { const a = String(r[0]); if (a.indexOf('■ ') === 0) { const b = findBlock_(sh, a.slice(2)); ranges.push(sh.getRange(b.first, DAY_COL, b.last - b.first + 1, 7)); } });
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(RED_AT).setBackground('#f8d0d0').setFontColor('#b02a2a').setRanges(ranges).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(AMBER_AT, RED_AT - 0.001).setBackground('#fde5c0').setFontColor('#8a5a00').setRanges(ranges).build()
  ]);
}

function appendBlock_(sh, s, days) {
  writeBlockSkeleton_(sh, sh.getLastRow() + 3, s, days);
}

// Block: header row "■ Station", column header row, item rows, Daily Cost row, Closing Time row, Counted By row
function writeBlockSkeleton_(sh, r, s, days) {
  sh.getRange(r, 1, 1, NCOLS).setBackground(WK_NAVY).setFontColor(WK_GOLD).setFontWeight('bold');
  sh.getRange(r, 1).setValue('■ ' + s.stationLabel);
  sh.getRange(r + 1, 1, 1, NCOLS).setValues([['Category', 'Item', 'UOM', 'Cost/Unit'].concat(days, ['Week Qty', 'Week Cost ($)'])])
    .setBackground('#e9e4d4').setFontWeight('bold');
  const items = s.items.map(i => [i.category, i.item, i.uom, i.costPerUnit]);
  const first = r + 2, last = first + items.length - 1;
  sh.getRange(first, 1, items.length, 4).setValues(items);
  sh.getRange(first, 4, items.length, 1).setNumberFormat('$0.00');
  sh.getRange(first, DAY_COL, items.length, 8).setNumberFormat('0.##;-0.##;"–"');
  sh.getRange(first, 13, items.length, 1).setNumberFormat('$#,##0.00');
  sh.getRange(first, DAY_COL, items.length, 7).setHorizontalAlignment('center');
  const costRow = last + 1;
  sh.getRange(costRow, 1, 3, 2).setValues([['', 'Daily Waste Cost ($)'], ['', 'Closing Time'], ['', 'Counted By (Badge)']]);
  sh.getRange(costRow, 1, 1, NCOLS).setFontWeight('bold').setBackground('#f5f1e4');
  sh.getRange(costRow, DAY_COL, 1, 9).setNumberFormat('$#,##0.00');
  sh.getRange(costRow + 1, 1, 2, NCOLS).setFontColor('#555555').setFontSize(9);
  sh.getRange(costRow + 1, DAY_COL, 2, 7).setHorizontalAlignment('center').setWrap(true);
  writeBlockFormulas_(sh, findBlock_(sh, s.stationLabel));
  return costRow + 2;
}

function findBlock_(sh, label) {
  const n = sh.getLastRow(); if (n < 1) return null;
  const v = sh.getRange(1, 1, n, 2).getValues();
  const start = v.findIndex(r => r[0] === '■ ' + label);
  if (start < 0) return null;
  const first = start + 3;                                   // 1-based: header row + col header + 1
  let r = first, items = {};
  while (r <= n && v[r - 1][1] !== 'Daily Waste Cost ($)') { if (v[r - 1][1] !== '') items[v[r - 1][1]] = r; r++; }
  return { head: start + 1, first, last: r - 1, costRow: r, closeRow: r + 1, chefRow: r + 2, items };
}

function writeBlockFormulas_(sh, b) {
  const n = b.last - b.first + 1;
  const rowF = [];
  for (let r = b.first; r <= b.last; r++) rowF.push(['=SUM(E' + r + ':K' + r + ')', '=L' + r + '*D' + r]);
  sh.getRange(b.first, 12, n, 2).setFormulas(rowF);
  const cols = ['E', 'F', 'G', 'H', 'I', 'J', 'K'];
  const f = cols.map(c => '=SUMPRODUCT(' + c + b.first + ':' + c + b.last + ',$D' + b.first + ':$D' + b.last + ')');
  sh.getRange(b.costRow, DAY_COL, 1, 7).setFormulas([f]);
  sh.getRange(b.costRow, 12, 1, 2).setFormulas([['', '=SUM(M' + b.first + ':M' + b.last + ')']]);
}

function writeSummary_(sh) {
  const n = sh.getLastRow();
  const v = sh.getRange(1, 1, n, 2).getValues();
  const blocks = [];
  v.forEach((row, i) => { if (String(row[0]).indexOf('■ ') === 0) blocks.push(String(row[0]).slice(2)); });
  const cols = ['E', 'F', 'G', 'H', 'I', 'J', 'K'];
  const rows = blocks.map(label => {
    const b = findBlock_(sh, label);
    return cols.map(c => '=' + c + b.costRow).concat(['', '=M' + b.costRow]);
  });
  const labels = blocks.map(l => [l]);
  const totalRow = 5 + rows.length;
  rows.push(cols.map(c => '=SUM(' + c + '5:' + c + (totalRow - 1) + ')').concat(['', '=SUM(M5:M' + (totalRow - 1) + ')']));
  labels.push(['ALL OUTLETS']);
  // make sure summary never overruns the first outlet block
  const firstBlockHead = findBlock_(sh, blocks[0]).head;
  if (5 + rows.length > firstBlockHead - 2) { sh.insertRowsBefore(firstBlockHead, 5 + rows.length - firstBlockHead + 2); return writeSummary_(sh); }
  sh.getRange(5, 2, labels.length, 1).setValues(labels);          // plain text labels
  sh.getRange(5, DAY_COL, rows.length, 9).setFormulas(rows);     // E–M formulas
  sh.getRange(5, DAY_COL, rows.length, 9).setNumberFormat('$#,##0.00');
  sh.getRange(totalRow, 1, 1, NCOLS).setFontWeight('bold');
}

function sendReport_(p, sheetUrl, week) {
  const [y, m, d] = p.date.split('-').map(Number);
  const dateStr = Utilities.formatDate(new Date(y, m - 1, d), TIMEZONE, 'EEE M/d/yyyy');
  const money = n => '$' + Number(n).toFixed(2);
  const color = q => q >= RED_AT ? '#d93b3b' : q >= AMBER_AT ? '#e08a00' : '#2e9e62';
  const dot = q => `<span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:${color(q)};margin-right:6px"></span>`;
  const td = 'padding:6px 8px;border-bottom:1px solid #f0f0f0';

  const all = [];
  p.stations.forEach(s => s.items.filter(i => i.qty > 0).forEach(i => all.push(Object.assign({ station: s.stationLabel }, i))));
  const top = all.slice().sort((a, b) => b.cost - a.cost)[0];
  const flagged = all.filter(i => i.qty >= RED_AT);
  const topStation = p.stations.slice().sort((a, b) => b.totals.cost - a.totals.cost)[0];

  // Outlet summary
  let summary = p.stations.map(s => {
    const pct = p.totals.cost ? Math.round(100 * s.totals.cost / p.totals.cost) : 0;
    const st = s.items.filter(i => i.qty > 0).sort((a, b) => b.cost - a.cost)[0];
    return `<tr>
      <td style="${td};font-weight:600">${esc_(s.stationLabel)}</td>
      <td style="${td};color:#777">${hm12_(s.closeTime)}</td>
      <td style="${td};text-align:right">${s.totals.units}</td>
      <td style="${td};text-align:right;font-weight:600">${money(s.totals.cost)}</td>
      <td style="${td};text-align:right;color:#777">${pct}%</td>
      <td style="${td};color:#555">${st ? esc_(st.item) : '—'}</td></tr>`;
  }).join('');
  summary += `<tr style="font-weight:700"><td style="padding:8px">All outlets</td><td></td>
      <td style="padding:8px;text-align:right">${p.totals.units}</td><td style="padding:8px;text-align:right">${money(p.totals.cost)}</td><td></td><td></td></tr>`;

  // Detail per outlet
  const details = p.stations.map(s => {
    const wasted = s.items.filter(i => i.qty > 0);
    let rows = '', lastCat = '';
    wasted.forEach(i => {
      if (i.category !== lastCat) { rows += `<tr><td colspan="4" style="padding:10px 8px 3px;font-weight:700;color:#8a6d12;font-size:12px">${esc_(i.category)}</td></tr>`; lastCat = i.category; }
      rows += `<tr><td style="${td}">${dot(i.qty)}${esc_(i.item)}</td><td style="${td};text-align:right">${i.qty} ${esc_(i.uom)}</td>
        <td style="${td};text-align:right;color:#777">${money(i.costPerUnit)}</td><td style="${td};text-align:right;font-weight:600">${money(i.cost)}</td></tr>`;
    });
    if (!rows) rows = `<tr><td colspan="4" style="padding:10px 8px;color:#2e9e62">No waste recorded.</td></tr>`;
    return `<div style="margin-top:22px">
      <div style="background:#0f1b33;color:#fff;padding:8px 12px;border-radius:6px;font-weight:700">${esc_(s.stationLabel)}
        <span style="float:right;font-weight:400;color:#d4af37">${money(s.totals.cost)} · closed ${hm12_(s.closeTime)}</span></div>
      <table style="width:100%;border-collapse:collapse;font-size:14px">${rows}</table>
      ${s.notes ? `<div style="margin-top:8px;padding:10px;background:#fafafa;border-left:3px solid #d4af37;font-size:13px"><b>Notes:</b> ${esc_(s.notes).replace(/\n/g, '<br>')}</div>` : ''}
    </div>`;
  }).join('');

  const kpi = (label, val) => `<td style="padding:10px 14px;background:#f5f1e4;border-radius:8px"><div style="font-size:12px;color:#777">${label}</div><div style="font-size:20px;font-weight:700;color:#0f1b33">${val}</div></td>`;
  const th = 'padding:6px 8px;font-size:12px;color:#777;font-weight:600;border-bottom:1px solid #ddd';

  // Week-to-date (Mon–Sun) table
  let weekHtml = '';
  if (week) {
    const labels = Object.keys(week.stationDaily);
    const dayTot = [0, 0, 0, 0, 0, 0, 0];
    labels.forEach(l => week.stationDaily[l].forEach((v, i) => dayTot[i] += v));
    const wtd = dayTot.reduce((x, y) => x + y, 0);
    const hd = i => i === week.idx ? 'background:#f5f1e4;font-weight:700' : '';
    const cell = (v, i) => `<td style="padding:5px 6px;text-align:right;border-bottom:1px solid #f0f0f0;${hd(i)}">${i > week.idx ? '' : (v ? money(v) : '–')}</td>`;
    weekHtml = `<div style="margin-top:22px">
      <div style="font-weight:700;color:#0f1b33;margin-bottom:6px">This week so far (Mon–Sun): ${money(wtd)}</div>
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <tr style="color:#777"><th style="text-align:left;padding:5px 6px;border-bottom:1px solid #ddd">Outlet</th>${week.days.map((d, i) => `<th style="padding:5px 6px;text-align:right;border-bottom:1px solid #ddd;${hd(i)}">${d.split(' ')[0]}</th>`).join('')}<th style="padding:5px 6px;text-align:right;border-bottom:1px solid #ddd">Week</th></tr>
        ${labels.map(l => `<tr><td style="padding:5px 6px;border-bottom:1px solid #f0f0f0">${esc_(l)}</td>${week.stationDaily[l].map(cell).join('')}<td style="padding:5px 6px;text-align:right;border-bottom:1px solid #f0f0f0;font-weight:600">${money(week.stationDaily[l].reduce((x, y) => x + y, 0))}</td></tr>`).join('')}
        <tr style="font-weight:700"><td style="padding:5px 6px">All outlets</td>${dayTot.map(cell).join('')}<td style="padding:5px 6px;text-align:right">${money(wtd)}</td></tr>
      </table>
      <div style="font-size:12px;margin-top:4px"><a href="${sheetUrl}#gid=${week.gid}" style="color:#8a6d12">Open the ${esc_(week.name)} log</a></div>
    </div>`;
  }

  const html = `
  <div style="font-family:Segoe UI,Arial,sans-serif;max-width:680px;color:#222">
    <div style="background:#0f1b33;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0;border-bottom:3px solid #d4af37">
      <div style="font-size:12px;letter-spacing:1px;color:#d4af37">FOOD COURT · NIGHTLY LINE WASTE · ALL OUTLETS</div>
      <div style="font-size:22px;font-weight:700">${dateStr}</div>
      <div style="font-size:13px;color:#c8d2e4">Counted by ${esc_(p.chef)} (Badge ${esc_(p.badge)})</div>
    </div>
    <div style="border:1px solid #e5e5e5;border-top:0;padding:16px 20px;border-radius:0 0 10px 10px">
      <table cellspacing="8" style="margin:-8px"><tr>
        ${kpi('Total waste cost', money(p.totals.cost))}${kpi('Items wasted', p.totals.itemsWasted)}${kpi('Total units', p.totals.units)}
      </tr></table>
      ${topStation && topStation.totals.cost ? `<p style="margin:14px 0 4px">Highest-cost outlet: <b>${esc_(topStation.stationLabel)}</b> (${money(topStation.totals.cost)})</p>` : ''}
      ${top ? `<p style="margin:4px 0">Highest $ item: <b>${esc_(top.item)}</b> at ${esc_(top.station)} — ${top.qty} ${esc_(top.uom)} (${money(top.cost)})</p>` : ''}
      ${flagged.length ? `<p style="margin:4px 0;color:#d93b3b">🔴 10+ units: ${flagged.map(i => esc_(i.item) + ' – ' + esc_(i.station) + ' (' + i.qty + ' ' + esc_(i.uom) + ')').join(', ')}. Review prep / par levels.</p>` : ''}
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin-top:14px">
        <tr style="text-align:left"><th style="${th}">Outlet</th><th style="${th}">Closed</th><th style="${th};text-align:right">Units</th><th style="${th};text-align:right">Cost</th><th style="${th};text-align:right">% of total</th><th style="${th}">Top $ item</th></tr>
        ${summary}
      </table>
      ${weekHtml}
      ${details}
      <p style="font-size:12px;color:#999;margin-top:18px">🟢 &lt;5 units · 🟠 5–9 · 🔴 10+ &nbsp;|&nbsp; <a href="${sheetUrl}" style="color:#8a6d12">Open the waste log sheet</a></p>
    </div>
  </div>`;

  const subject = `Line Waste · All Outlets · ${dateStr} · ${money(p.totals.cost)} (${p.chef})`;
  const opts = { to: RECIPIENTS, subject, htmlBody: html, name: 'Food Court Waste Log' };
  if (CC) opts.cc = CC;
  MailApp.sendEmail(opts);
}

/* =====================================================================
   WEEKLY REPORT EMAIL (Mon–Sun) — built from the week tab
   ===================================================================== */
// A week counts as sent only once its full Mon–Sun report has gone out (marker ends with "(complete)").
function weeklySent_(sh) { const v = String(sh.getRange(3, 1).getValue()); return v.indexOf('Weekly report emailed') === 0 && v.indexOf('(complete)') > 0; }
function dayLabel_(x) { return x instanceof Date ? Utilities.formatDate(x, TIMEZONE, 'EEE M/d') : String(x); }

function readWeek_(sh) {
  const n = sh.getLastRow();
  const v = sh.getRange(1, 1, n, NCOLS).getValues();
  const days = v[3].slice(DAY_COL - 1, DAY_COL + 6).map(dayLabel_);        // row 4 headers (Sheets may store these as dates)
  const out = [];
  v.forEach(row => {
    const a = String(row[0]);
    if (a.indexOf('■ ') !== 0) return;
    const b = findBlock_(sh, a.slice(2));
    const items = [];
    for (let r = b.first; r <= b.last; r++) {
      const x = v[r - 1];
      const q = x.slice(DAY_COL - 1, DAY_COL + 6).map(z => typeof z === 'number' ? z : null);
      items.push({ cat: x[0], item: x[1], uom: x[2], cost: Number(x[3]) || 0, days: q });
    }
    out.push({ label: a.slice(2), items,
      close: v[b.closeRow - 1].slice(DAY_COL - 1, DAY_COL + 6).map(String),
      chef: v[b.chefRow - 1].slice(DAY_COL - 1, DAY_COL + 6).map(String) });
  });
  return { days, stations: out };
}

function weekBounds_(sh) {
  const m = String(sh.getRange(1, 1).getValue()).match(/Sun (\d+)\/(\d+)\/(\d{4})/); if (!m) return null;
  const sun = new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2]));
  return { mon: new Date(sun.getFullYear(), sun.getMonth(), sun.getDate() - 6), sun };
}
// Sheets may store a business date as a Date or keep the ISO string.
function dateOnly_(x) {
  if (x instanceof Date) return new Date(x.getFullYear(), x.getMonth(), x.getDate());
  const s = String(x);
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? new Date(Number(s.slice(0, 4)), Number(s.slice(5, 7)) - 1, Number(s.slice(8, 10))) : null;
}

function weekNotes_(ss, sh) {
  const subs = ss.getSheetByName(SUB_SHEET); if (!subs || subs.getLastRow() < 2) return [];
  const wk = weekBounds_(sh); if (!wk) return [];
  const { mon, sun } = wk;
  const rows = subs.getRange(2, 1, subs.getLastRow() - 1, 11).getValues();
  const toDate = dateOnly_;
  const seen = {};
  return rows.filter(r => { const d = toDate(r[2]); return d && d >= mon && d <= sun && String(r[10]).trim(); })
    .map(r => { const d = toDate(r[2]); return { day: ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][d.getDay()] + ' ' + md_(d), station: r[3], chef: r[5], note: String(r[10]).trim() }; })
    .filter(x => { const k = x.day + x.station + x.note; if (seen[k]) return false; seen[k] = 1; return true; });
}

function sendWeeklyReport_(ss, sh, late) {
  const W = readWeek_(sh);
  const money = n => '$' + Number(n).toFixed(2);
  const fmtQ = q => (Math.round(q * 100) / 100).toString();
  const color = q => q >= RED_AT ? '#d93b3b' : q >= AMBER_AT ? '#e08a00' : '';

  // day coverage: a day counts as logged if any outlet has a closing time for it
  const logged = [0, 1, 2, 3, 4, 5, 6].map(i => W.stations.some(s => s.close[i] && s.close[i] !== ''));
  const missing = W.days.filter((d, i) => !logged[i]);

  const stDay = W.stations.map(s => [0, 1, 2, 3, 4, 5, 6].map(i => s.items.reduce((a, it) => a + (it.days[i] || 0) * it.cost, 0)));
  const dayTot = [0, 1, 2, 3, 4, 5, 6].map(i => stDay.reduce((a, r) => a + r[i], 0));
  const weekCost = dayTot.reduce((a, b) => a + b, 0);
  const all = [];
  W.stations.forEach(s => s.items.forEach(it => {
    const q = it.days.reduce((a, b) => a + (b || 0), 0);
    if (q > 0) all.push({ station: s.label, item: it.item, uom: it.uom, qty: q, cost: q * it.cost, peak: Math.max.apply(null, it.days.map(x => x || 0)) });
  }));
  const units = all.reduce((a, b) => a + b.qty, 0);
  const top = all.slice().sort((a, b) => b.cost - a.cost).slice(0, 10);
  const flagged = all.filter(i => i.peak >= RED_AT);
  const worstDay = dayTot.indexOf(Math.max.apply(null, dayTot));
  const stWeek = W.stations.map((s, i) => ({ label: s.label, cost: stDay[i].reduce((a, b) => a + b, 0) }));
  const topSt = stWeek.slice().sort((a, b) => b.cost - a.cost)[0];

  // ---- Phone-first layout: no table wider than 3 columns; full Mon–Sun grid lives in the Sheet ----
  const NAVY = '#0f1b33', GOLD = '#c9a54a', INK = '#1f2937', MUTE = '#6b7280', LINE = '#eceef2';
  const cellL = `padding:9px 0;border-bottom:1px solid ${LINE};vertical-align:top`;
  const cellR = `padding:9px 0 9px 10px;border-bottom:1px solid ${LINE};text-align:right;white-space:nowrap;vertical-align:top`;
  const h = t => `<div style="margin:26px 0 8px;font-size:12px;font-weight:700;letter-spacing:.8px;text-transform:uppercase;color:${MUTE}">${t}</div>`;
  const kpi = (label, val) => `<td width="50%" style="padding:6px"><div style="background:#f6f3ea;border-radius:10px;padding:12px 14px"><div style="font-size:12px;color:${MUTE}">${label}</div><div style="font-size:22px;font-weight:700;color:${NAVY};margin-top:2px">${val}</div></div></td>`;
  const bar = pct => `<div style="height:5px;background:${LINE};border-radius:5px;margin-top:6px"><div style="height:5px;width:${Math.max(2, Math.round(pct))}%;background:${GOLD};border-radius:5px"></div></div>`;

  const outletRows = W.stations.map((s, i) => ({ s, i, cost: stWeek[i].cost })).sort((a, b) => b.cost - a.cost).map(({ s, cost }) => {
    const pct = weekCost ? 100 * cost / weekCost : 0;
    return `<tr><td style="${cellL}"><div style="font-weight:600;color:${INK}">${esc_(s.label)}</div>${bar(pct)}</td><td style="${cellR}"><div style="font-weight:700;color:${INK}">${money(cost)}</div><div style="font-size:12px;color:${MUTE}">${Math.round(pct)}%</div></td></tr>`;
  }).join('');

  const dayRows = W.days.map((d, i) => `<tr><td style="${cellL};color:${INK}">${esc_(d)}${i === worstDay && weekCost ? ` <span style="font-size:11px;font-weight:700;color:#b02a2a;background:#fbe3e1;border-radius:4px;padding:1px 6px;margin-left:4px">Highest</span>` : ''}</td><td style="${cellR};${logged[i] ? `font-weight:600;color:${INK}` : `color:#b02a2a`}">${logged[i] ? money(dayTot[i]) : 'No count'}</td></tr>`).join('');

  const topRows = top.map((t, i) => `<tr><td style="${cellL}"><span style="color:${MUTE};display:inline-block;width:20px">${i + 1}</span><span style="font-weight:600;color:${INK}">${esc_(t.item)}</span><div style="font-size:12px;color:${MUTE};margin-left:20px">${esc_(t.station)} · ${fmtQ(t.qty)} ${esc_(t.uom)}</div></td><td style="${cellR};font-weight:700;color:${INK}">${money(t.cost)}</td></tr>`).join('');

  const dayAbbr = W.days.map(d => d.split(' ')[0]);
  const details = W.stations.map((s, si) => {
    const rows = s.items.filter(it => it.days.some(q => q > 0)).map(it => {
      const q = it.days.reduce((a, b) => a + (b || 0), 0);
      const perDay = it.days.map((x, d) => x > 0 ? `<span style="white-space:nowrap;${x >= AMBER_AT ? 'color:' + color(x) + ';font-weight:700' : ''}">${dayAbbr[d]} ${fmtQ(x)}</span>` : '').filter(Boolean).join(' · ');
      return `<tr><td style="${cellL}"><div style="font-weight:600;color:${INK}">${esc_(it.item)}</div><div style="font-size:12px;color:${MUTE};margin-top:2px;line-height:1.5">${perDay}</div></td><td style="${cellR}"><div style="font-weight:700;color:${INK}">${money(q * it.cost)}</div><div style="font-size:12px;color:${MUTE}">${fmtQ(q)} ${esc_(it.uom)}</div></td></tr>`;
    }).join('') || `<tr><td colspan="2" style="${cellL};color:#2e9e62">No waste recorded this week.</td></tr>`;
    return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px;border-collapse:collapse">
        <tr><td style="background:${NAVY};color:#fff;padding:9px 12px;border-radius:8px 0 0 8px;font-weight:700">${esc_(s.label)}</td><td style="background:${NAVY};color:${GOLD};padding:9px 12px;border-radius:0 8px 8px 0;text-align:right;font-weight:700;white-space:nowrap">${money(stWeek[si].cost)}</td></tr>
        ${rows}
      </table>`;
  }).join('');

  const notes = weekNotes_(ss, sh);
  const notesHtml = notes.length ? h('Chef notes') + `<div style="background:#f9fafb;border-left:3px solid ${GOLD};border-radius:0 8px 8px 0;padding:10px 12px;font-size:13px;color:${INK}">${notes.map(n => `<div style="margin:3px 0"><b>${esc_(n.day)} · ${esc_(n.station)}</b><br>${esc_(n.note)} <span style="color:${MUTE}">— ${esc_(n.chef)}</span></div>`).join('')}</div>` : '';

  // Day & Swing shift waste: reported on its own, never added into the closing totals above.
  const sw = weekShifts_(ss, sh);
  let shiftsHtml = '';
  if (sw.length) {
    const sumBy = (key) => { const o = {}; sw.forEach(x => { const k = x[key]; o[k] = o[k] || { cost: 0, qty: 0 }; o[k].cost += x.cost; o[k].qty += x.qty; }); return o; };
    const byShift = sumBy('shift'), bySt = sumBy('station');
    const items = {}; sw.forEach(x => { const k = x.station + '|' + x.item; items[k] = items[k] || { station: x.station, item: x.item, uom: x.uom, qty: 0, cost: 0 }; items[k].qty += x.qty; items[k].cost += x.cost; });
    const swTop = Object.keys(items).map(k => items[k]).sort((a, b) => b.cost - a.cost).slice(0, 5);
    const swTotal = sw.reduce((a, x) => a + x.cost, 0);
    const row = (l, cost, sub) => `<tr><td style="${cellL};color:${INK}">${l}</td><td style="${cellR}"><div style="font-weight:700;color:${INK}">${money(cost)}</div>${sub ? `<div style="font-size:12px;color:${MUTE}">${sub}</div>` : ''}</td></tr>`;
    shiftsHtml = h('Day &amp; Swing shifts') +
      `<div style="font-size:12px;color:${MUTE};margin-top:-2px">Waste logged during the shift. Not included in the totals above.</div>` +
      `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:6px;border-collapse:collapse">` +
      ['Day', 'Swing'].filter(s => byShift[s]).map(s => row(`<b>${s} shift</b>`, byShift[s].cost, fmtQ(byShift[s].qty) + ' units')).join('') +
      `<tr><td style="padding:10px 0;font-weight:700">Shift total</td><td style="padding:10px 0 10px 10px;text-align:right;font-weight:700">${money(swTotal)}</td></tr></table>` +
      `<div style="font-size:12px;font-weight:600;color:${MUTE};margin:12px 0 2px">By outlet</div>` +
      `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">` +
      Object.keys(bySt).sort((a, b) => bySt[b].cost - bySt[a].cost).map(s => row(esc_(s), bySt[s].cost, '')).join('') + `</table>` +
      `<div style="font-size:12px;font-weight:600;color:${MUTE};margin:12px 0 2px">Top items</div>` +
      `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">` +
      swTop.map(t => row(`<span style="font-weight:600">${esc_(t.item)}</span><div style="font-size:12px;color:${MUTE}">${esc_(t.station)} · ${fmtQ(t.qty)} ${esc_(t.uom)}</div>`, t.cost, '')).join('') + `</table>`;
  }

  const alerts = [];
  if (late && !logged[6]) alerts.push(`<div style="background:#fdf1dc;color:#8a5a00;border-radius:8px;padding:10px 12px;margin-bottom:8px;font-size:13px">Sunday's count wasn't submitted, so this report went out with the first count of the next week.</div>`);
  if (missing.length) alerts.push(`<div style="background:#fbe3e1;color:#9b2c24;border-radius:8px;padding:10px 12px;margin-bottom:8px;font-size:13px"><b>No count submitted:</b> ${missing.map(esc_).join(', ')}</div>`);
  if (flagged.length) alerts.push(`<div style="background:#fbe3e1;color:#9b2c24;border-radius:8px;padding:10px 12px;margin-bottom:8px;font-size:13px"><b>10+ units in a single day:</b> ${flagged.map(i => esc_(i.item) + ' (' + esc_(i.station) + ')').join(', ')}. Review prep / par levels.</div>`);

  const range = W.days[0].replace(/^Mon /, '') + ' – ' + W.days[6].replace(/^Sun /, '');
  const html = `
  <div style="margin:0;padding:0;background:#f3f4f6">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;background:#f3f4f6"><tr><td align="center" style="padding:12px 8px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;border-collapse:collapse;background:#ffffff;border-radius:12px;font-family:-apple-system,'Segoe UI',Roboto,Arial,sans-serif;font-size:14px;line-height:1.45;color:${INK}">
    <tr><td style="background:${NAVY};border-radius:12px 12px 0 0;padding:18px 18px 16px;border-bottom:3px solid ${GOLD}">
      <div style="font-size:11px;letter-spacing:1.2px;font-weight:700;color:${GOLD}">FOOD COURT · WEEKLY LINE WASTE</div>
      <div style="font-size:22px;font-weight:700;color:#ffffff;margin-top:4px">Mon ${esc_(range)}</div>
      <div style="font-size:13px;color:#c8d2e4;margin-top:2px">PFC · Pronto · Little Wok · Agave</div>
    </td></tr>
    <tr><td style="padding:14px 12px 4px">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
        <tr>${kpi('Week waste cost', money(weekCost))}${kpi('Days logged', logged.filter(Boolean).length + ' / 7')}</tr>
        <tr>${kpi('Items wasted', all.length)}${kpi('Total units', fmtQ(units))}</tr>
      </table>
    </td></tr>
    <tr><td style="padding:4px 18px 22px">
      ${alerts.join('')}
      ${h('By outlet')}
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">${outletRows}
        <tr><td style="padding:10px 0;font-weight:700">All outlets</td><td style="padding:10px 0 10px 10px;text-align:right;font-weight:700">${money(weekCost)}</td></tr></table>
      ${h('By day')}
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">${dayRows}</table>
      ${top.length ? h('Top ' + top.length + ' items by cost') + `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">${topRows}</table>` : ''}
      ${h('Outlet detail')}
      <div style="font-size:12px;color:${MUTE};margin-top:-2px">Days with waste are listed under each item · <span style="color:#e08a00;font-weight:700">5–9</span> · <span style="color:#d93b3b;font-weight:700">10+</span></div>
      ${details}
      ${shiftsHtml}
      ${notesHtml}
      <div style="margin-top:24px;text-align:center"><a href="${ss.getUrl()}#gid=${sh.getSheetId()}" style="display:inline-block;background:${NAVY};color:#ffffff;text-decoration:none;font-weight:600;padding:11px 18px;border-radius:8px">Open the full Mon–Sun log</a></div>
    </td></tr>
  </table>
  </td></tr></table>
  </div>`;

  const already = weeklySent_(sh);
  const mSun = String(sh.getRange(1, 1).getValue()).match(/Sun (\d+)\/(\d+)\/(\d{4})/);
  const weekOver = mSun ? new Date() >= new Date(Number(mSun[3]), Number(mSun[1]) - 1, Number(mSun[2]) + 1) : true;
  const inProgress = !logged[6] && !late && !weekOver;
  const subject = `${already ? 'UPDATED · ' : ''}${inProgress ? 'Week so far · ' : ''}Weekly Line Waste · Mon ${range} · ${money(weekCost)}`;
  const opts = { to: RECIPIENTS, subject, htmlBody: html, name: 'Food Court Waste Log' };
  if (CC) opts.cc = CC;
  MailApp.sendEmail(opts);
  // Mark the week sent only when the week is over (so a mid-week preview doesn't block Sunday's report)
  const complete = logged[6] || weekOver;
  sh.getRange(3, 1).setValue('Weekly report emailed ' + Utilities.formatDate(new Date(), TIMEZONE, 'EEE M/d h:mm a') + (complete ? ' (complete)' : ' (preview)'));
}

/** Run from the editor to email LAST week's (most recent finished Mon–Sun) report right now. */
function sendWeeklyReportNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const done = ss.getSheets().filter(s => /^Week \d+-\d+ to /.test(s.getName())).map(s => {
    const m = String(s.getRange(1, 1).getValue()).match(/Sun (\d+)\/(\d+)\/(\d{4})/);
    return { s, sun: m ? new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2])) : null };
  }).filter(x => x.sun && x.sun < today).sort((a, b) => b.sun - a.sun);
  if (done.length) sendWeeklyReport_(ss, done[0].s, false);
}

/** Run from the editor to email THIS week's numbers so far (a preview; Sunday's report still goes out). */
function sendThisWeekSoFar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheets().find(s => /^Week \d+-\d+ to /.test(s.getName()));
  if (sh) sendWeeklyReport_(ss, sh, false);
}

function sheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f1b33').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function hm12_(t) { if (!t) return '—'; const [h, m] = t.split(':').map(Number); return ((h % 12) || 12) + ':' + String(m).padStart(2, '0') + (h < 12 ? ' AM' : ' PM'); }
function esc_(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Run once from the editor to authorize Gmail + Sheets and send yourself a test. */
function authorizeAndTest() {
  MailApp.sendEmail(RECIPIENTS, '✅ Line Waste Log — setup test', 'Apps Script can send email. Now deploy as a Web app.');
}
