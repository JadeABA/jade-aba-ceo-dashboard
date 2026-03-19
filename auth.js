let _token = null;
let _msalApp = null;

const MSAL_CONFIG = {
  auth: {
    clientId: '5a8e24ef-86dc-443d-93fc-e74b11365575',
    authority: 'https://login.microsoftonline.com/33ee9db4-64f2-4c9f-a3c3-ae3cd287d483',
    redirectUri: 'https://JadeABA.github.io/jade-aba-ceo-dashboard',
    navigateToLoginRequestUrl: true,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: true,
  }
};

const SCOPES = ['Files.Read', 'User.Read'];

async function getMSAL() {
  if (_msalApp) return _msalApp;
  _msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
  await _msalApp.initialize();
  return _msalApp;
}

async function signInPopup() {
  const app = await getMSAL();
  try {
    const result = await app.loginRedirect({ scopes: SCOPES });
    return result ? true : false;
  } catch(e) {
    console.error('Login error:', e);
    return false;
  }
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
  } catch(e) {
    console.error('Redirect handle error:', e);
  }
  const accounts = app.getAllAccounts();
  if (accounts.length > 0) {
    try {
      const silentResult = await app.acquireTokenSilent({
        scopes: SCOPES,
        account: accounts[0]
      });
      _token = silentResult.accessToken;
      sessionStorage.setItem('jade_token', _token);
      return true;
    } catch(e) {
      console.log('Silent failed, need login');
    }
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
```

Also update `index.html` — find the line that loads MSAL at the top and make sure it says exactly this. Go to `index.html` → pencil → find this line near the very top:
```
<script src="https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js"></script>
```

Change it to:
```
<script src="https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.28.1/lib/msal-browser.min.js"></script>
