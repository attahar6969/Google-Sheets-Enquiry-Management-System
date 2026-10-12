/**
 * SALES DASHBOARD  (new file: Dashboard.gs)
 *  - showDashboard()      : popup dashboard (menu: Sales Tracker > Dashboard)
 *  - getDashboardData(f)  : read-only data for the popup (runs as the user)
 *  - refreshDashboard_()  : corrected "Dashboard" sheet  (REPLACES the old refreshDashboard_ in Code.gs)
 */

function showDashboard() { showDialog_('Dashboard', 'Sales Dashboard', 1250, 720); }

function openDashboardSheet() {
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheetByName(SHEETS.DASH));
}

function refreshSystem() {
  callOwner_('refresh', {});
  return 'ok';
}

/* ------------------------------------------------------------------ */
/*  POPUP DATA                                                         */
/* ------------------------------------------------------------------ */
function getDashboardData(f) {
  f = f || {};
  const ss = ss_();
  const tz = ss.getSpreadsheetTimeZone();
  const key = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const sh = ss.getSheetByName(SHEETS.ENQ);
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, 20).getValues() : [];

  // Date a deal was marked Won = latest "Won" row in Follow-Up History (fallback: enquiry date)
  const wonAt = {};
  const hi = ss.getSheetByName(SHEETS.HIST);
  const hl = hi.getLastRow();
  if (hl > 1) {
    hi.getRange(2, 1, hl - 1, 8).getValues().forEach(r => {
      if (String(r[7]).trim().toLowerCase() === 'won' && r[0] instanceof Date) {
        const id = String(r[1]).trim();
        if (!wonAt[id] || r[0] > wonAt[id]) wonAt[id] = r[0];
      }
    });
  }

  const all = [];
  rows.forEach(r => {
    if (!r[0]) return;
    const d = r[1] instanceof Date ? r[1] : null;
    const st = String(r[11] || '').trim();
    const lc = st.toLowerCase();
    const amt = Number(r[8]) || 0, bill = Number(r[18]) || 0;
    const won = lc === 'won';
    const wd = won ? (wonAt[String(r[0]).trim()] || d) : null;
    all.push({
      id: String(r[0]), d: d ? key(d) : '', cust: String(r[2]),
      rep: String(r[10] || '').trim() || '(Blank)',
      cat: String(r[6] || '').trim() || '(Blank)',
      src: String(r[9] || '').trim() || '(Blank)',
      status: st || '(Blank)', lc: lc,
      amt: amt, sales: won ? (bill > 0 ? bill : amt) : 0,
      fup: r[13] instanceof Date ? key(r[13]) : '',
      wd: wd ? key(wd) : ''
    });
  });

  // filter options
  const dd = getDropdowns_();
  const union = (a, b) => a.concat(b).filter((x, i, arr) => x && arr.indexOf(x) === i);
  const opts = {
    reps: union(dd.reps, all.map(x => x.rep)),
    statuses: union(dd.statuses, all.map(x => x.status)),
    categories: union(dd.categories, all.map(x => x.cat)),
    sources: union(dd.sources, all.map(x => x.src))
  };

  // apply filters (on Date of Enquiry)
  const data = all.filter(x => {
    if (f.from && (!x.d || x.d < f.from)) return false;
    if (f.to && (!x.d || x.d > f.to)) return false;
    if (f.month && x.d.slice(0, 7) !== f.month) return false;
    if (f.rep && x.rep !== f.rep) return false;
    if (f.status && x.status !== f.status) return false;
    if (f.category && x.cat !== f.category) return false;
    if (f.source && x.src !== f.source) return false;
    return true;
  });

  const isFin = x => x.lc === 'won' || x.lc === 'lost' || x.lc === 'closed';
  const k = { total: data.length, won: 0, lost: 0, closed: 0, open: 0, quote: 0, sales: 0, pipeline: 0 };
  const statusMap = {}, sourceMap = {}, repMap = {}, monMap = {}, rm = {}, cal = {}, days = {};
  const months = {};

  const rep = n => repMap[n] || (repMap[n] = { name: n, enq: 0, open: 0, won: 0, lost: 0, closed: 0, quote: 0, sales: 0 });
  const mon = m => monMap[m] || (monMap[m] = { enq: 0, won: 0, sales: 0 });
  const rmc = (r, m) => { rm[r] = rm[r] || {}; return rm[r][m] || (rm[r][m] = { enq: 0, won: 0, sales: 0, quote: 0 }); };
  const calc = dk => cal[dk] || (cal[dk] = { enq: 0, fup: 0, won: 0, amt: 0 });
  const addDay = (dk, o) => { (days[dk] = days[dk] || []).push(o); };

  data.forEach(x => {
    statusMap[x.status] = (statusMap[x.status] || 0) + 1;
    sourceMap[x.src] = (sourceMap[x.src] || 0) + 1;
    k.quote += x.amt;
    const R = rep(x.rep);
    R.enq++; R.quote += x.amt;
    if (x.lc === 'won') { k.won++; k.sales += x.sales; R.won++; R.sales += x.sales; }
    else if (x.lc === 'lost') { k.lost++; R.lost++; }
    else if (x.lc === 'closed') { k.closed++; R.closed++; }
    else { k.open++; k.pipeline += x.amt; R.open++; }

    const em = x.d.slice(0, 7);
    if (em) {
      months[em] = 1;
      mon(em).enq++;
      const c = rmc(x.rep, em); c.enq++; c.quote += x.amt;
    }
    if (x.lc === 'won') {
      const wm = (x.wd || x.d).slice(0, 7);
      if (wm) {
        months[wm] = 1;
        const m = mon(wm); m.won++; m.sales += x.sales;
        const c = rmc(x.rep, wm); c.won++; c.sales += x.sales;
      }
    }
    // calendar
    const base = { id: x.id, cust: x.cust, rep: x.rep, status: x.status };
    if (x.d) { calc(x.d).enq++; addDay(x.d, Object.assign({ t: 'Enquiry', amt: x.amt }, base)); }
    if (!isFin(x) && x.fup) { calc(x.fup).fup++; addDay(x.fup, Object.assign({ t: 'Follow-up due', amt: x.amt }, base)); }
    if (x.lc === 'won') {
      const wd = x.wd || x.d;
      if (wd) { const c = calc(wd); c.won++; c.amt += x.sales; addDay(wd, Object.assign({ t: 'Won', amt: x.sales }, base)); }
    }
  });

  const monthList = Object.keys(months).sort().slice(-12);
  const monthly = monthList.map(m => ({ m: m, enq: (monMap[m] || {}).enq || 0, won: (monMap[m] || {}).won || 0, sales: (monMap[m] || {}).sales || 0 }));
  const repList = Object.keys(repMap).map(n => repMap[n]).sort((a, b) => b.sales - a.sales || b.enq - a.enq);
  const toArr = o => Object.keys(o).map(n => [n, o[n]]).sort((a, b) => b[1] - a[1]);

  const recent = data.slice().sort((a, b) => b.id.localeCompare(a.id)).slice(0, 10)
    .map(x => ({ id: x.id, date: x.d, cust: x.cust, rep: x.rep, status: x.status, quote: x.amt, fup: x.fup }));

  k.conv = k.total ? k.won / k.total : 0;
  const fin = k.won + k.lost + k.closed;
  k.winRate = fin ? k.won / fin : 0;

  return {
    opts: opts, kpi: k, statusCounts: toArr(statusMap), sourceCounts: toArr(sourceMap),
    reps: repList, months: monthList, repMonth: rm, monthly: monthly,
    cal: cal, days: days, recent: recent
  };
}

/* ------------------------------------------------------------------ */
/*  "Dashboard" SHEET  (corrected version)                             */
/*  - wider columns (no cut-off labels)                                */
/*  - Closed card + Win Rate card added                                */
/*  - "Other / Blank" rows so every table adds up to the Total         */
/*  - monthly table starts at the first month that has data            */
/* ------------------------------------------------------------------ */
function refreshDashboard_() {
  const ss = ss_();
  const d = ss.getSheetByName(SHEETS.DASH);
  d.getCharts().forEach(c => d.removeChart(c));
  d.clear();

  const T = "'" + SHEETS.ENQ + "'!";
  const R = c => T + '$' + c + '$2:$' + c + '$' + MAXROW;
  const uniq = a => a.filter((x, i) => a.indexOf(x) === i).slice(0, 17);
  const dd = getDropdowns_();

  d.setColumnWidths(1, 13, 150);
  d.getRange('A1').setValue('SALES ENQUIRY DASHBOARD').setFontSize(20).setFontWeight('bold').setFontColor('#1a4e8a');
  d.getRange('A2').setFormula('="Updated: "&TEXT(NOW(),"dd-mmm-yyyy hh:mm")').setFontColor('#5f6368');

  const labels = ['Total Enquiries', 'Open', 'Won', 'Lost', 'Closed', 'Conversion % (Won ÷ All)',
    'Win Rate % (Won ÷ Won+Lost+Closed)', 'Total Quotation Value', 'Won Revenue', 'Pipeline Value'];
  const f = [
    '=COUNTA(' + R('A') + ')',
    '=A4-C4-D4-E4',
    '=COUNTIF(' + R('L') + ',"Won")',
    '=COUNTIF(' + R('L') + ',"Lost")',
    '=COUNTIF(' + R('L') + ',"Closed")',
    '=IFERROR(C4/A4,0)',
    '=IFERROR(C4/(C4+D4+E4),0)',
    '=SUM(' + R('I') + ')',
    '=SUMIF(' + R('L') + ',"Won",' + R('S') + ')',
    '=SUMIFS(' + R('I') + ',' + R('L') + ',"<>Won",' + R('L') + ',"<>Lost",' + R('L') + ',"<>Closed",' + R('A') + ',"<>")'
  ];
  d.getRange(3, 1, 1, 10).setValues([labels]).setBackground('#e8f0fe').setFontWeight('bold')
    .setFontColor('#1a4e8a').setHorizontalAlignment('center').setWrap(true).setVerticalAlignment('middle');
  d.setRowHeight(3, 44);
  d.getRange(4, 1, 1, 10).setFormulas([f]).setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground('#f8f9fa');
  d.getRange('F4:G4').setNumberFormat('0.0%');
  d.getRange('H4:J4').setNumberFormat(INR_FORMAT);
  d.setRowHeight(4, 42);

  const head = (r, c, arr) => d.getRange(r, c, 1, arr.length).setValues([arr]).setBackground('#1a4e8a')
    .setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  const other = (r, c, sumCol, n) => {
    d.getRange(r, c).setValue('Other / Blank').setFontStyle('italic');
    d.getRange(r, c + 1).setFormula('=MAX(0,$A$4-SUM(' + sumCol + '8:' + sumCol + (7 + n) + '))').setFontStyle('italic');
  };

  // Status
  const st = uniq(dd.statuses);
  head(7, 1, ['Status', 'Enquiries']);
  st.forEach((s, i) => {
    d.getRange(8 + i, 1).setValue(s);
    d.getRange(8 + i, 2).setFormula('=COUNTIF(' + R('L') + ',A' + (8 + i) + ')');
  });
  other(8 + st.length, 1, 'B', st.length);

  // Source
  const so = uniq(dd.sources);
  head(7, 4, ['Source', 'Enquiries']);
  so.forEach((s, i) => {
    d.getRange(8 + i, 4).setValue(s);
    d.getRange(8 + i, 5).setFormula('=COUNTIF(' + R('J') + ',D' + (8 + i) + ')');
  });
  other(8 + so.length, 4, 'E', so.length);

  // Representative
  const rp = uniq(dd.reps);
  head(7, 7, ['Representative', 'Enquiries', 'Won', 'Won Revenue']);
  rp.forEach((s, i) => {
    const r = 8 + i;
    d.getRange(r, 7).setValue(s);
    d.getRange(r, 8).setFormula('=COUNTIF(' + R('K') + ',G' + r + ')');
    d.getRange(r, 9).setFormula('=COUNTIFS(' + R('K') + ',G' + r + ',' + R('L') + ',"Won")');
    d.getRange(r, 10).setFormula('=SUMIFS(' + R('S') + ',' + R('K') + ',G' + r + ',' + R('L') + ',"Won")')
      .setNumberFormat(INR_FORMAT);
  });
  other(8 + rp.length, 7, 'H', rp.length);

  // Monthly trend: from first month with data (max 12 months)
  const tr = ss.getSheetByName(SHEETS.ENQ);
  const tl = tr.getLastRow();
  const now = new Date();
  let start = new Date(now.getFullYear(), now.getMonth() - 11, 1);
  if (tl > 1) {
    let minD = null;
    tr.getRange(2, COL.DATE, tl - 1, 1).getValues().forEach(r => {
      if (r[0] instanceof Date && (!minD || r[0] < minD)) minD = r[0];
    });
    if (minD) {
      const ms = new Date(minD.getFullYear(), minD.getMonth(), 1);
      if (ms > start) start = ms;
    }
  }
  const nMonths = (now.getFullYear() - start.getFullYear()) * 12 + now.getMonth() - start.getMonth() + 1;
  head(7, 12, ['Month', 'Enquiries']);
  for (let i = 0; i < nMonths; i++) {
    const r = 8 + i;
    d.getRange(r, 12).setValue(new Date(start.getFullYear(), start.getMonth() + i, 1)).setNumberFormat('mmm yy');
    d.getRange(r, 13).setFormula('=COUNTIFS(' + R('B') + ',">="&L' + r + ',' + R('B') + ',"<"&EDATE(L' + r + ',1))');
  }

  // Charts
  const opts = (b, title, w) => b.setNumHeaders(1).setOption('title', title).setOption('width', w || 470)
    .setOption('height', 300).build();
  d.insertChart(opts(d.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(d.getRange(7, 1, st.length + 2, 2)).setPosition(30, 1, 0, 0), 'Status Breakdown'));
  d.insertChart(opts(d.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(d.getRange(7, 4, so.length + 2, 2)).setPosition(30, 5, 20, 0), 'Enquiries by Source'));
  if (rp.length) d.insertChart(opts(d.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(d.getRange(7, 7, rp.length + 1, 3)).setPosition(47, 1, 0, 0), 'Representative Performance'));
  d.insertChart(opts(d.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(d.getRange(7, 12, nMonths + 1, 2)).setPosition(47, 5, 20, 0), 'Monthly Enquiries'));
}
