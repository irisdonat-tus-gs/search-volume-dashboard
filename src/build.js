#!/usr/bin/env node
'use strict';

const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const MARKETS = ['DE','ES','FR','UK','IT'];
const BRANDS  = ['GS','TUS','BOTH'];

// Month name → zero-padded number (multilingual)
const MONTH_NUM = {
  jan:'01', january:'01', janvier:'01', enero:'01', januar:'01',
  feb:'02', february:'02', fev:'02', fevrier:'02', febrero:'02', februar:'02',
  mar:'03', march:'03', mars:'03', marzo:'03', marz:'03',
  apr:'04', april:'04', avril:'04', abril:'04',
  may:'05', mai:'05', mayo:'05',
  jun:'06', june:'06', juin:'06', junio:'06', juni:'06',
  jul:'07', july:'07', juillet:'07', julio:'07', juli:'07',
  aug:'08', august:'08', aout:'08', agosto:'08',
  sep:'09', september:'09', septembre:'09', septiembre:'09',
  oct:'10', october:'10', octobre:'10', octubre:'10', okt:'10',
  nov:'11', november:'11', novembre:'11', noviembre:'11',
  dec:'12', december:'12', decembre:'12', diciembre:'12', dez:'12',
};

// Excel serial date (1900 system) → "YYYY-MM"
function serialToMonthKey(n) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

// Normalize any header variant to "YYYY-MM": real Date cells, Excel serial numbers,
// "Searches: Jan 2025" prefixed labels, and bare "Jan 2025" / "January 2025".
function toMonthKey(header) {
  if (header === null || header === undefined || header === '') return null;
  // Real Date cell
  if (header instanceof Date) return serialToMonthKey((header.getTime() / 86400000) + 25569);
  // Numeric Excel serial date (date-formatted headers come through as ~45000–46000)
  if (typeof header === 'number') return (header > 20000 && header < 60000) ? serialToMonthKey(header) : null;

  // Strip a leading "Searches:" (or similar) prefix some Semrush exports add
  let s = header.toString().trim().replace(/^[a-zA-ZÀ-ɏ]+\s*:\s*/, '');
  // Numeric string serial
  if (/^\d+(\.\d+)?$/.test(s)) {
    const n = parseFloat(s);
    return (n > 20000 && n < 60000) ? serialToMonthKey(n) : null;
  }
  // "Jan 2025", "Jan. 2025", "January 2025"
  const m = s.match(/^([a-zA-ZÀ-ɏ]+)\.?\s+(\d{4})$/);
  if (!m) return null;
  const word = m[1].toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); // strip diacritics
  const num  = MONTH_NUM[word];
  if (!num) return null;
  return `${m[2]}-${num}`;
}

function parseSheet(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  // raw:true keeps numbers as numbers (percentages as 0.xx decimals)
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (rows.length < 3) return [];

  const hdrs = rows[1]; // row 0 = TOTAL, row 1 = headers
  let kwCol = -1;
  const monthCols = {};
  let ytd25Col = -1, ytd26Col = -1, ytdYoyCol = -1;

  hdrs.forEach((h, i) => {
    if (h === null || h === undefined) return;
    const s   = h.toString().trim();
    const low = s.toLowerCase();

    if (kwCol === -1 && /^keywords?$/i.test(s)) { kwCol = i; return; }

    const mk = toMonthKey(h);
    if (mk) { monthCols[mk] = i; return; }

    if (low.includes('ytd') || (low.includes('jan') && low.includes('may'))) {
      if ((low.includes('25') || low.includes('2025')) && ytd25Col === -1) { ytd25Col = i; return; }
      if ((low.includes('26') || low.includes('2026')) && ytd26Col === -1) { ytd26Col = i; return; }
    }
    if ((low.includes('yoy') || low.includes('%')) && ytdYoyCol === -1 && ytd26Col > -1 && i > ytd26Col) {
      ytdYoyCol = i;
    }
  });

  // Positional fallback: assume last 3 non-month cols after all months are ytd25, ytd26, ytdYoy
  if (ytd25Col === -1 || ytd26Col === -1) {
    const lastMon = Object.values(monthCols).length ? Math.max(...Object.values(monthCols)) : -1;
    const trailing = hdrs.map((h, i) => ({ h, i }))
      .filter(x => x.i > lastMon && !Object.values(monthCols).includes(x.i) && x.h !== null);
    const n = trailing.length;
    if (n >= 2 && ytd25Col === -1) ytd25Col = trailing[n - 3]?.i ?? trailing[0].i;
    if (n >= 2 && ytd26Col === -1) ytd26Col = trailing[n - 2]?.i ?? trailing[1].i;
    if (n >= 1 && ytdYoyCol === -1) ytdYoyCol = trailing[n - 1].i;
  }

  if (kwCol === -1) { console.warn(`  ⚠  No keyword column in ${sheetName}`); return []; }

  const result = [];
  for (let r = 2; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row[kwCol] == null) continue;
    const kw = row[kwCol].toString().trim();
    if (!kw || /^total$/i.test(kw)) continue;

    const obj = { keyword: kw };
    for (const [mk, ci] of Object.entries(monthCols)) {
      const v = row[ci];
      obj[mk] = (v !== null && v !== undefined && v !== '') ? Number(v) : null;
    }
    const n = v => (v !== null && v !== undefined && v !== '') ? Number(v) : 0;
    obj._ytd25   = ytd25Col  >= 0 ? n(row[ytd25Col])  : 0;
    obj._ytd26   = ytd26Col  >= 0 ? n(row[ytd26Col])  : 0;
    obj._ytd_yoy = ytdYoyCol >= 0 ? n(row[ytdYoyCol]) : 0;
    result.push(obj);
  }
  return result;
}

function build() {
  // Find Excel file in data/
  const dataDir  = path.join(__dirname, '..', 'data');
  const xlsxFile = fs.readdirSync(dataDir).find(f => /\.(xlsx|xls)$/i.test(f));
  if (!xlsxFile) throw new Error('No .xlsx file found in data/');
  const xlsxPath = path.join(dataDir, xlsxFile);
  console.log(`📂  Reading ${xlsxFile}`);

  const wb = XLSX.readFile(xlsxPath);

  // Parse every market-brand sheet
  const allRows = [];
  for (const name of wb.SheetNames) {
    const m = name.match(/^(DE|ES|FR|UK|IT)[_\-\s]*(GS|TUS)$/i);
    if (!m) continue;
    const market = m[1].toUpperCase();
    const brand  = m[2].toUpperCase();
    const rows   = parseSheet(wb, name);
    rows.forEach(r => { r.market = market; r.brand = brand; });
    allRows.push(...rows);
    console.log(`  ✓  ${name}: ${rows.length} rows`);
  }
  console.log(`\n     Raw total: ${allRows.length} rows`);

  // Collect all month keys across all sheets, sort chronologically
  const monthSet = new Set();
  for (const row of allRows)
    for (const k of Object.keys(row))
      if (/^\d{4}-\d{2}$/.test(k)) monthSet.add(k);
  let sortedMonths = [...monthSet].sort();

  // Dynamic scope: keep only months whose month-of-year also appears in the LATEST year
  // present. Auto-extends as new months are added (add Jun 2026 → Jun 2025 is pulled in for
  // comparison) while dropping lonely months with no counterpart in the latest year (e.g.
  // Jun–Dec 2025 from sheets that carry a full rolling history).
  if (sortedMonths.length) {
    const years = [...new Set(sortedMonths.map(k => k.slice(0, 4)))].sort();
    const latestYear = years[years.length - 1];
    const latestMM = new Set(sortedMonths.filter(k => k.slice(0, 4) === latestYear).map(k => k.slice(5)));
    const dropped = sortedMonths.filter(k => !latestMM.has(k.slice(5)));
    sortedMonths = sortedMonths.filter(k => latestMM.has(k.slice(5)));
    if (dropped.length) console.log(`     Scoped out ${dropped.length} month(s) with no ${latestYear} match: ${dropped.join(', ')}`);
  }

  const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const monthLabels = sortedMonths.map(k => {
    const [y, m] = k.split('-');
    return `${MN[parseInt(m, 10) - 1]} ${y}`;
  });

  // Deduplicate by (keyword_lower, market) — average volumes across GS+TUS
  const grouped = new Map();
  for (const row of allRows) {
    const key = row.keyword.toLowerCase() + '|' + row.market;
    const arr = grouped.get(key) || [];
    arr.push(row);
    grouped.set(key, arr);
  }

  const compact = [];
  for (const [, group] of grouped) {
    const r0   = group[0];
    const mIdx = MARKETS.indexOf(r0.market);
    if (mIdx === -1) continue;

    const brands = group.map(r => r.brand);
    const bIdx   = (brands.includes('GS') && brands.includes('TUS')) ? 2
                 : brands.includes('GS') ? 0 : 1;

    const avgField = field => {
      const vals = group.map(r => r[field]).filter(v => v != null);
      return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
    };

    const vols = sortedMonths.map(mk => {
      const vals = group.map(r => r[mk]).filter(v => v != null);
      return vals.length ? Math.round(vals.reduce((a, b) => a + b, 0) / vals.length) : null;
    });

    compact.push([r0.keyword, mIdx, bIdx, ...vols,
      avgField('_ytd25'), avgField('_ytd26'), avgField('_ytd_yoy')]);
  }

  const dupes = allRows.length - grouped.size;
  console.log(`     Deduplicated: ${compact.length} unique keywords (merged ${dupes} duplicates)`);
  console.log(`     Months: ${monthLabels[0]} → ${monthLabels[monthLabels.length - 1]}`);

  // Inject into template
  const tplPath = path.join(__dirname, '..', 'dashboard_template.html');
  let html = fs.readFileSync(tplPath, 'utf8');
  html = html
    .replace('__DATA__',         JSON.stringify(compact))
    .replace('__MONTH_KEYS__',   JSON.stringify(sortedMonths))
    .replace('__MONTH_LABELS__', JSON.stringify(monthLabels));

  const outDir  = path.join(__dirname, '..', 'public');
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, 'index.html');
  fs.writeFileSync(outPath, html, 'utf8');

  const kb = Math.round(html.length / 1024);
  console.log(`\n✅  Written public/index.html (${kb} KB)`);
}

build();
