let _token = null;
let _msalApp = null;

function excelDateToJS(val) {
  if (!val && val !== 0) return null;
  if (typeof val === 'number' && val > 40000) {
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

async function loadUtilizationData() {
  if (!_token) return { weeks: [], current: null, previous: null };
  const fileId = CONFIG.files.utilization.fileId;
  const sheetName = CONFIG.files.utilization.sheetName;
  const url = 'https://graph.microsoft.com/v1.0/me/drive/items/' + fileId +
    '/workbook/worksheets(\'' + encodeURIComponent(sheetName) + '\')/usedRange';
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + _token } });
  if (!res.ok) { console.error('Could not read utilization sheet', res.status); return { weeks: [], current: null, previous: null }; }
  const data = await res.json();
  const rows = data.values || [];
  var weeks = [];
  var currentWeekLabel = null;
  var clientCount = 0;
  var clientsUnder80 = 0;
  rows.forEach(function(r) {
    var cell0 = String(r[0] || '').trim();
    var cell1 = String(r[1] || '').trim();
    var isWeekHeader = cell0.includes('WEEK:');
    var isTotal = cell0.includes('WEEKLY TOTAL') || cell1.includes('WEEKLY TOTAL');
    if (isWeekHeader) {
      currentWeekLabel = cell0;
      clientCount = 0;
      clientsUnder80 = 0;
    } else if (isTotal && currentWeekLabel) {
      var authorized  = n(r[3]);
      var clientAvail = n(r[4]);
      var scheduled   = n(r[5]);
      var rendered    = n(r[6]);
      var utilPct     = clientAvail > 0 ? Math.round((rendered / clientAvail) * 100) : 0;
      var openHrs     = Math.round((clientAvail - rendered) * 10) / 10;
      weeks.push({
        label: currentWeekLabel, authorized: authorized, clientAvail: clientAvail,
        scheduled: scheduled, rendered: rendered, utilPct: utilPct,
        openHrs: openHrs, clientCount: clientCount, clientsUnder80: clientsUnder80,
      });
      currentWeekLabel = null;
    } else if (currentWeekLabel && cell1 && !cell0.includes('Date') && cell1 !== 'Case' && cell1 !== '') {
      var rend = n(r[6]);
      var avail = n(r[4]);
      clientCount++;
      if (avail > 0 && rend / avail < 0.80) clientsUnder80++;
    }
  });
  var dataWeeks = weeks.filter(function(w) { return w.clientAvail > 0; });
  var current  = dataWeeks.length > 0 ? dataWeeks[dataWeeks.length - 1] : null;
  var previous = dataWeeks.length > 1 ? dataWeeks[dataWeeks.length - 2] : null;
  var monthAgo = dataWeeks.length > 4 ? dataWeeks[dataWeeks.length - 5] : null;
  return { weeks: dataWeeks, current: current, previous: previous, monthAgo: monthAgo };
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
  const utilData = await loadUtilizationData();
  return {
    arRows:      results[0],
    bcbaRows:    results[1],
    intakeRows:  results[2],
    recruitRows: results[3],
    caseRows:    results[4],
    util:        utilData,
  };
}

function nthLatest(rows, n) {
  if (!rows || !rows.length) return {};
  var withData = rows.filter(function(r) { return r.__date && r.__hasData; });
  if (!withData.length) withData = rows.filter(function(r) { return r.__date; });
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
