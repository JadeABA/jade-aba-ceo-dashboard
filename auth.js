// Microsoft login using MSAL — modern browser-based auth
let _token = null;

async function signIn() {
  const authUrl = new URL(`https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/authorize`);
  authUrl.searchParams.set('client_id',     CONFIG.clientId);
  authUrl.searchParams.set('response_type', 'code');
  authUrl.searchParams.set('redirect_uri',  CONFIG.redirectUri);
  authUrl.searchParams.set('scope',         'openid profile Files.Read User.Read offline_access');
  authUrl.searchParams.set('response_mode', 'fragment');
  authUrl.searchParams.set('code_challenge_method', 'plain');
  authUrl.searchParams.set('code_challenge', 'jade_aba_dashboard_2026');
  window.location.href = authUrl.toString();
}

function handleLogin() {
  // Check for token in session
  const saved = sessionStorage.getItem('jade_token');
  if (saved) { _token = saved; return true; }

  // Check URL fragment for auth code
  const hash = window.location.hash.substring(1);
  const params = new URLSearchParams(hash);

  if (params.get('error')) {
    console.error('Auth error:', params.get('error_description'));
    return false;
  }
  return false;
}

async function signInPopup() {
  return new Promise((resolve) => {
    const authUrl = new URL(`https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/authorize`);
    authUrl.searchParams.set('client_id',     CONFIG.clientId);
    authUrl.searchParams.set('response_type', 'token');
    authUrl.searchParams.set('redirect_uri',  CONFIG.redirectUri);
    authUrl.searchParams.set('scope',         'Files.Read User.Read');
    authUrl.searchParams.set('response_mode', 'fragment');
    authUrl.searchParams.set('nonce',         Date.now().toString());
    authUrl.searchParams.set('prompt',        'select_account');

    const popup = window.open(authUrl.toString(), 'auth', 'width=500,height=600,scrollbars=yes');

    const timer = setInterval(() => {
      try {
        if (popup.closed) { clearInterval(timer); resolve(false); return; }
        const hash = popup.location.hash;
        if (hash && hash.includes('access_token')) {
          const p = new URLSearchParams(hash.substring(1));
          _token = p.get('access_token');
          sessionStorage.setItem('jade_token', _token);
          popup.close();
          clearInterval(timer);
          resolve(true);
        }
      } catch(e) { /* cross-origin, keep waiting */ }
    }, 500);
  });
}

async function readSheet(fileId, sheetName) {
  if (!_token) return [];
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`;
  const res  = await fetch(url, { headers: { Authorization: `Bearer ${_token}` } });
  if (!res.ok) { console.error('Could not read', sheetName, res.status); return []; }
  const data    = await res.json();
  const rows    = data.values || [];
  if (rows.length < 2) return [];
  const headers = rows[0].map(h => String(h).trim());
  return rows.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] ?? '');
      return obj;
    });
}

async function readCell(fileId, sheetName, cellAddress) {
  if (!_token) return null;
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/range(address='${cellAddress}')`;
  const res  = await fetch(url, { headers: { Authorization: `Bearer ${_token}` } });
  if (!res.ok) return null;
  const data = await res.json();
  return data.values?.[0]?.[0] ?? null;
}

async function loadAllData() {
  const f = CONFIG.files;
  const [arRows, bcbaRows, intakeRows, recruitRows, caseRows] = await Promise.all([
    readSheet(f.ar.fileId,         f.ar.sheetName),
    readSheet(f.bcba.fileId,       f.bcba.sheetName),
    readSheet(f.intake.fileId,     f.intake.sheetName),
    readSheet(f.recruiting.fileId, f.recruiting.sheetName),
    readSheet(f.caseCoord.fileId,  f.caseCoord.sheetName),
  ]);
  const uc = f.utilization.cells;
  const [totalAuth, totalAvail, totalSched, utilPct, clientsLow, therapistsLow, openHrs] = await Promise.all([
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.totalAuthorized),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.totalAvailable),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.totalScheduled),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.overallUtilPct),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.clientsUnder80),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.therapistsUnder70),
    readCell(f.utilization.fileId, f.utilization.sheetName, uc.openHours),
  ]);
  return {
    arRows, bcbaRows, intakeRows, recruitRows, caseRows,
    util: { totalAuth, totalAvail, totalSched, utilPct, clientsLow, therapistsLow, openHrs }
  };
}

function latestRow(rows, weekCol) {
  if (!rows?.length) return {};
  return rows.reduce((best, row) => {
    const d = new Date(row[weekCol]);
    return !isNaN(d) && d > new Date(best[weekCol] || 0) ? row : best;
  }, {});
}

function filterRange(rows, weekCol, range) {
  if (!rows?.length) return [];
  const now    = new Date();
  let   cutoff = new Date(0);
  if      (range === 'wtd') { cutoff = new Date(now); cutoff.setDate(now.getDate() - now.getDay()); }
  else if (range === 'mtd') { cutoff = new Date(now.getFullYear(), now.getMonth(), 1); }
  else if (range === 'qtd') { cutoff = new Date(now.getFullYear(), Math.floor(now.getMonth()/3)*3, 1); }
  else if (range === 'ytd') { cutoff = new Date(now.getFullYear(), 0, 1); }
  return rows.filter(row => { const d = new Date(row[weekCol]); return !isNaN(d) && d >= cutoff; });
}

function n(val) {
  if (val === null || val === undefined || val === '') return 0;
  return parseFloat(String(val).replace(/[$,%]/g, '').trim()) || 0;
}
```

Also go to `index.html` → pencil ✏️ → find this line:
```
window.addEventListener('load', () => {
  if (handleLogin()) { startApp(); }
});
