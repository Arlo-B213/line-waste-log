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
const RECIPIENTS  = 'arlo.bedolla@gmail.com'; // management inbox(es), comma-separated
const CC          = '';                                           // optional
const SHARED_KEY  = 'pechanga-fc-waste';                          // must match the app
const AMBER_AT = 5, RED_AT = 10;                                  // template status key
const TIMEZONE = 'America/Los_Angeles';
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
        htmlBody: '<p>The Line Waste Log app is connected. Nightly all-outlet waste reports will arrive here.</p>' });
      return json_({ ok: true, test: true });
    }
    if (!p.stations || !p.stations.length) return json_({ ok: false, error: 'No outlets in report' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
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

    sendReport_(p, ss.getUrl());
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return json_({ ok: true, service: 'Line Waste Log' }); }

function sendReport_(p, sheetUrl) {
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
      ${details}
      <p style="font-size:12px;color:#999;margin-top:18px">🟢 &lt;5 units · 🟠 5–9 · 🔴 10+ &nbsp;|&nbsp; <a href="${sheetUrl}" style="color:#8a6d12">Open the waste log sheet</a></p>
    </div>
  </div>`;

  const subject = `Line Waste · All Outlets · ${dateStr} · ${money(p.totals.cost)} (${p.chef})`;
  const opts = { to: RECIPIENTS, subject, htmlBody: html, name: 'Food Court Waste Log' };
  if (CC) opts.cc = CC;
  MailApp.sendEmail(opts);
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
