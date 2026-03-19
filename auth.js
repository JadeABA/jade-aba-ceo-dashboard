let _token = null;

function signInPopup() {
  // Redirect method — no popup needed, goes to Microsoft and comes back
  const authUrl = new URL('https://login.microsoftonline.com/' + CONFIG.tenantId + '/oauth2/v2.0/authorize');
  authUrl.searchParams.set('client_id',     CONFIG.clientId);
  authUrl.searchParams.set('response_type', 'token');
  authUrl.searchParams.set('redirect_uri',  CONFIG.redirectUri);
  authUrl.searchParams.set('scope',         'Files.Read User.Read');
  authUrl.searchParams.set('response_mode', 'fragment');
  authUrl.searchParams.set('nonce',         Date.now().toString());
  authUrl.searchParams.set('prompt',        'select_account');
  window.location.href = authUrl.toString();
  return new Promise(function() {});
}

function handleLogin() {
  // Check if Microsoft just sent us back with a token in the URL
  const hash = window.location.hash.substring(1);
  const params = new URLSearchParams(hash);
  const token = params.get('access_token');
  if (token) {
    _token = token;
    sessionStorage.setItem('jade_token', token);
    window.history.replaceState(null, '', window.location.pathname);
    return true;
  }
  // Check if we already have a saved token
  const saved = sessionStorage.getItem('jade_token');
  if (saved) {
    _token = saved;
    return true;
  }
  return false;
}

async function readSheet(fileId, sheetName) {
  if (!_token) return [];
  const url = 'https://graph.microsoft.com/v1.0/me/drive/items/' + fileId + '/workbook/worksheets(\'' + encodeURIComponent(sheetName) + '\')/usedRange';
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + _token } });
  if (!res.ok) { console.error('Could not read ' + sheetName, res.status); return []; }
  const data = await res.json();
  const rows = data.values || [];
  if (rows.length < 2) return [];
  const headers = rows[0].map(function(h) { return String(h).trim(); });
  return rows.slice(1)
    .filter(function(row) { return row.some(function(cell) { return cell !== '' && cell !== null; }); })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    });
}

async function readCell(fileId, sheetName, cellAddress) {
  if (!_token) return null;
  const url = 'https://graph.microsoft.com/v1.0/me/drive/items/' + fileId + '/workbook/worksheets(\'' + encodeURIComponent(sheetName) + '\')/range(address=\'' + cellAddress + '\')';
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + _token } });
  if (!res.ok) return null;
  const data = await res.json();
  return data.values && data.values[0] && data.values[0][0] !== undefined ? data.values[0][0] : null;
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
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.clientsUnder80),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.therapistsUnder70),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.openHours),
  ]);
  return {
    arRows:      results[0],
    bcbaRows:    results[1],
    intakeRows:  results[2],
    recruitRows: results[3],
    caseRows:    results[4],
    util: {
      totalAuth:     cellResults[0],
      totalAvail:    cellResults[1],
      totalSched:    cellResults[2],
      utilPct:       cellResults[3],
      clientsLow:    cellResults[4],
      therapistsLow: cellResults[5],
      openHrs:       cellResults[6],
    }
  };
}

function latestRow(rows, weekCol) {
  if (!rows || !rows.length) return {};
  return rows.reduce(function(best, row) {
    const d = new Date(row[weekCol]);
    return !isNaN(d) && d > new Date(best[weekCol] || 0) ? row : best;
  }, {});
}

function filterRange(rows, weekCol, range) {
  if (!rows || !rows.length) return [];
  const now = new Date();
  var cutoff = new Date(0);
  if      (range === 'wtd') { cutoff = new Date(now); cutoff.setDate(now.getDate() - now.getDay()); }
  else if (range === 'mtd') { cutoff = new Date(now.getFullYear(), now.getMonth(), 1); }
  else if (range === 'qtd') { cutoff = new Date(now.getFullYear(), Math.floor(now.getMonth()/3)*3, 1); }
  else if (range === 'ytd') { cutoff = new Date(now.getFullYear(), 0, 1); }
  return rows.filter(function(row) {
    const d = new Date(row[weekCol]);
    return !isNaN(d) && d >= cutoff;
  });
}

function n(val) {
  if (val === null || val === undefined || val === '') return 0;
  return parseFloat(String(val).replace(/[$,%]/g, '').trim()) || 0;
}
