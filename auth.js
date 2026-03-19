// Handles Microsoft login and reads your OneDrive Excel files

let _token = null;

// Called when you click "Sign in with Microsoft"
async function signIn() {
  const url = new URL(`https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/authorize`);
  url.searchParams.set('client_id',     CONFIG.clientId);
  url.searchParams.set('response_type', 'token');
  url.searchParams.set('redirect_uri',  CONFIG.redirectUri);
  url.searchParams.set('scope',         'Files.Read User.Read');
  url.searchParams.set('response_mode', 'fragment');
  window.location.href = url.toString();
}

// Runs automatically when the page loads — checks if you just logged in
function handleLogin() {
  const params = new URLSearchParams(window.location.hash.substring(1));
  const token  = params.get('access_token');
  if (token) {
    _token = token;
    sessionStorage.setItem('jade_token', token);
    window.history.replaceState(null, '', window.location.pathname);
    return true;
  }
  const saved = sessionStorage.getItem('jade_token');
  if (saved) { _token = saved; return true; }
  return false;
}

// Reads one tab from one Excel file on OneDrive
async function readSheet(fileId, sheetName) {
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

// Reads a specific cell from the utilization dashboard tab
async function readCell(fileId, sheetName, cellAddress) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/range(address='${cellAddress}')`;
  const res  = await fetch(url, { headers: { Authorization: `Bearer ${_token}` } });
  if (!res.ok) return null;
  const data = await res.json();
  return data.values?.[0]?.[0] ?? null;
}

// Loads ALL your Excel data at once
async function loadAllData() {
  const f = CONFIG.files;
  const [arRows, bcbaRows, intakeRows, recruitRows, caseRows] = await Promise.all([
    readSheet(f.ar.fileId,         f.ar.sheetName),
    readSheet(f.bcba.fileId,       f.bcba.sheetName),
    readSheet(f.intake.fileId,     f.intake.sheetName),
    readSheet(f.recruiting.fileId, f.recruiting.sheetName),
    readSheet(f.caseCoord.fileId,  f.caseCoord.sheetName),
  ]);
  // Read the utilization summary cells individually
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

// Gets the most recent row from any dataset
function latestRow(rows, weekCol) {
  if (!rows?.length) return {};
  return rows.reduce((best, row) => {
    const d = new Date(row[weekCol]);
    return !isNaN(d) && d > new Date(best[weekCol] || 0) ? row : best;
  }, {});
}

// Filters rows to only those within the selected date range
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

// Converts any value to a clean number
function n(val) {
  if (val === null || val === undefined || val === '') return 0;
  const clean = String(val).replace(/[$,%]/g, '').trim();
  return parseFloat(clean) || 0;
}
