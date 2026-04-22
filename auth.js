let _token = null;
let _msalApp = null;

function excelDateToJS(val) {
  if (!val && val !== 0) return null;
  if (typeof val === 'number' && val > 40000) {
    // Convert Excel serial to JS date using UTC components to avoid timezone shift
    var d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  if (typeof val === 'string' && /^\d+\.\d+\.\d+$/.test(val.trim())) {
    const p = val.trim().split('.');
    return new Date(parseInt(p[2]), parseInt(p[0])-1, parseInt(p[1]));
  }
  if (typeof val === 'string' && (val.includes('/') || val.includes('-'))) {
    const d = new Date(val);
    return isNaN(d) ? null : d;
  }
  return null;
}

function pctToNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  const str = String(val).trim();
  if (str.endsWith('%')) return parseFloat(str) || 0;
  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  if (num > 0 && num <= 2) return Math.round(num * 100);
  return Math.round(num);
}

async function getMSAL() {
  if (_msalApp) return _msalApp;
  _msalApp = new msal.PublicClientApplication({
    auth: {
      clientId:    '5a8e24ef-86dc-443d-93fc-e74b11365575',
      authority:   'https://login.microsoftonline.com/33ee9db4-64f2-4c9f-a3c3-ae3cd287d483',
      redirectUri: 'https://JadeABA.github.io/jade-aba-ceo-dashboard',
    },
    cache: {
      cacheLocation:          'sessionStorage',
      storeAuthStateInCookie: true,
    }
  });
  return _msalApp;
}

async function signInPopup() {
  const app = await getMSAL();
  await app.loginRedirect({ scopes: ['Files.Read', 'User.Read'] });
}

async function handleLogin() {
  const app = await getMSAL();
  try {
    const result = await app.handleRedirectPromise();
    if (result && result.accessToken) {
      _token = result.accessToken;
      sessionStorage.setItem('jade_token', _token);
      return true;
    }
  } catch(e) { console.error('Redirect error:', e); }
  const accounts = app.getAllAccounts();
  if (accounts.length > 0) {
    try {
      const silent = await app.acquireTokenSilent({
        scopes:  ['Files.Read', 'User.Read'],
        account: accounts[0]
      });
      _token = silent.accessToken;
      sessionStorage.setItem('jade_token', _token);
      return true;
    } catch(e) { console.log('Silent failed, need login'); }
  }
  return false;
}

async function readSheet(fileId, sheetName) {
  if (!_token) return [];
  const url = 'https://graph.microsoft.com/v1.0/me/drive/items/' + fileId +
    '/workbook/worksheets(\'' + encodeURIComponent(sheetName) + '\')/usedRange';
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + _token } });
  if (!res.ok) { console.error('Could not read ' + sheetName, res.status); return []; }
  const data = await res.json();
  const rows = data.values || [];
  if (rows.length < 2) return [];
  const headers = rows[0].map(function(h) { return String(h).trim(); });
  return rows.slice(1)
    .filter(function(row) {
      return row.some(function(cell) { return cell !== '' && cell !== null; });
    })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) {
        obj[h] = row[i] !== undefined ? row[i] : '';
      });
      obj.__date = excelDateToJS(row[0]);
      // Check only columns 2-7 (skip the two date columns at start)
      // This avoids formula columns like WoW% and date formulas triggering false positives
      obj.__hasData = row.slice(2, 8).filter(function(cell) {
        return cell !== '' && cell !== null && cell !== 0 &&
               cell !== '0' && cell !== false && cell !== '0.0%' &&
               !(typeof cell === 'number' && cell === 0);
      }).length >= 1;
      return obj;
    });
}

async function readCell(fileId, sheetName, cellAddress) {
  if (!_token) return null;
  const url = 'https://graph.microsoft.com/v1.0/me/drive/items/' + fileId +
    '/workbook/worksheets(\'' + encodeURIComponent(sheetName) + '\')/range(address=\'' + cellAddress + '\')';
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + _token } });
  if (!res.ok) return null;
  const data = await res.json();
  return data.values && data.values[0] && data.values[0][0] !== undefined
    ? data.values[0][0] : null;
}

async function readRange(fileId, sheetName, rangeAddress) {
  if (!_token) return [];
  const url = 'https://graph.microsoft.com/v1.0/me/drive/items/' + fileId +
    '/workbook/worksheets(\'' + encodeURIComponent(sheetName) + '\')/range(address=\'' + rangeAddress + '\')';
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + _token } });
  if (!res.ok) return [];
  const data = await res.json();
  return data.values || [];
}

async function loadAllData() {
  const f = CONFIG.files;
  const results = await Promise.all([
    readSheet(f.ar.fileId,         f.ar.sheetName),
    readSheet(f.bcba.fileId,       f.bcba.sheetName),
    readSheet(f.intake.fileId,     f.intake.sheetName),
    readSheet(f.recruiting.fileId, f.recruiting.sheetName),
    readSheet(f.caseCoord.fileId,  f.caseCoord.sheetName),
  ]);

  const uc = f.utilization.cells;
  const cellResults = await Promise.all([
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.totalAuthorized),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.totalAvailable),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.totalScheduled),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.overallUtilPct),
  ]);

  const totalAvail = n(cellResults[1]);
  const totalSched = n(cellResults[2]);
  const openHrs    = Math.round((totalAvail - totalSched) * 10) / 10;

  const clientRows    = await readRange(f.utilization.fileId, f.utilization.sheetName, 'A12:D41');
  const therapistRows = await readRange(f.utilization.fileId, f.utilization.sheetName, 'F12:I31');

  let clientsLow = 0, clientsTotal = 0;
  clientRows.forEach(function(row) {
    const name = row[0];
    const util = n(row[3]);
    if (name && String(name).trim() !== '') {
      clientsTotal++;
      const utilPct = util <= 1 ? util * 100 : util;
      if (utilPct < 80) clientsLow++;
    }
  });

  let therapistsLow = 0, therapistsTotal = 0;
  therapistRows.forEach(function(row) {
    const name = row[0];
    const util = n(row[3]);
    if (name && String(name).trim() !== '') {
      therapistsTotal++;
      const utilPct = util <= 1 ? util * 100 : util;
      if (utilPct < 70) therapistsLow++;
    }
  });

  return {
    arRows:      results[0],
    bcbaRows:    results[1],
    intakeRows:  results[2],
    recruitRows: results[3],
    caseRows:    results[4],
    util: {
      totalAuth:      cellResults[0],
      totalAvail:     cellResults[1],
      totalSched:     cellResults[2],
      utilPct:        cellResults[3],
      clientsLow:     clientsLow,
      clientsTotal:   clientsTotal,
      therapistsLow:  therapistsLow,
      therapistsTotal:therapistsTotal,
      openHrs:        openHrs,
    }
  };
}

function nthLatest(rows, n) {
  if (!rows || !rows.length) return {};
  var withData = rows.filter(function(r) {
    return r.__date && r.__hasData;
  });
  if (!withData.length) {
    withData = rows.filter(function(r) { return r.__date; });
  }
  if (!withData.length) return rows[rows.length - 1] || {};
  var sorted = withData.slice().sort(function(a, b) { return b.__date - a.__date; });
  return sorted[n] || sorted[sorted.length - 1] || {};
}

function rowFromWeeksAgo(rows, weeksAgo) {
  if (!rows || !rows.length) return {};
  var latest = nthLatest(rows, 0);
  if (!latest.__date) return {};
  var target   = new Date(latest.__date.getTime() - weeksAgo * 7 * 24 * 60 * 60 * 1000);
  var best     = null, bestDiff = Infinity;
  rows.forEach(function(r) {
    if (!r.__date || !r.__hasData) return;
    var diff = Math.abs(r.__date.getTime() - target.getTime());
    if (diff < bestDiff && diff < 10 * 24 * 60 * 60 * 1000) { bestDiff = diff; best = r; }
  });
  return best || {};
}

function rowFromMonthsAgo(rows, monthsAgo) {
  if (!rows || !rows.length) return {};
  var latest = nthLatest(rows, 0);
  if (!latest.__date) return {};
  var target   = new Date(latest.__date);
  target.setMonth(target.getMonth() - monthsAgo);
  var best     = null, bestDiff = Infinity;
  rows.forEach(function(r) {
    if (!r.__date || !r.__hasData) return;
    var diff = Math.abs(r.__date.getTime() - target.getTime());
    if (diff < bestDiff && diff < 21 * 24 * 60 * 60 * 1000) { bestDiff = diff; best = r; }
  });
  return best || {};
}

function filterRange(rows, weekCol, range) {
  if (!rows || !rows.length) return [];
  const now    = new Date();
  var   cutoff = new Date(0);
  if      (range === 'wtd') { cutoff = new Date(now); cutoff.setDate(now.getDate() - now.getDay()); }
  else if (range === 'mtd') { cutoff = new Date(now.getFullYear(), now.getMonth(), 1); }
  else if (range === 'qtd') { cutoff = new Date(now.getFullYear(), Math.floor(now.getMonth()/3)*3, 1); }
  else if (range === 'ytd') { cutoff = new Date(now.getFullYear(), 0, 1); }
  const filtered = rows.filter(function(row) {
    return row.__date && row.__date >= cutoff;
  });
  if (filtered.length === 0) {
    const latest = nthLatest(rows, 0);
    if (!latest || !latest.__date) return rows.slice(-5);
    const latestTime = latest.__date.getTime();
    return rows.filter(function(r) {
      return r.__date && Math.abs(r.__date.getTime() - latestTime) < 8 * 24 * 60 * 60 * 1000;
    });
  }
  return filtered;
}

function n(val) {
  if (val === null || val === undefined || val === '') return 0;
  return parseFloat(String(val).replace(/[$,%]/g, '').trim()) || 0;
}
