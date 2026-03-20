let _token = null;
let _msalApp = null;

// Converts ANY date format Excel uses into a real JavaScript date
function excelDateToJS(val) {
  if (!val && val !== 0) return null;
  // Excel serial number (e.g. 46081)
  if (typeof val === 'number' && val > 40000) {
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }
  // Dot format like "3.6.2026" (month.day.year) or "1.1.2024"
  if (typeof val === 'string' && /^\d+\.\d+\.\d+$/.test(val.trim())) {
    const p = val.trim().split('.');
    return new Date(parseInt(p[2]), parseInt(p[0])-1, parseInt(p[1]));
  }
  // Slash format like "3/6/2026"
  if (typeof val === 'string' && val.includes('/')) {
    const d = new Date(val);
    return isNaN(d) ? null : d;
  }
  // Try generic parse
  const d = new Date(val);
  return isNaN(d) ? null : d;
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
      // Parse date from first column for sorting
      obj.__date = excelDateToJS(row[0]);
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
  // Calculate open hours directly — avail minus scheduled
  const totalAvail = cellResults[1] || 0;
  const totalSched = cellResults[2] || 0;
  const openHrs = Math.round((totalAvail - totalSched) * 10) / 10;
  // Get client/therapist alert counts from the utilization sheet directly
  const utilSheet = await readSheet(f.utilization.fileId, f.utilizat
