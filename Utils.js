/* ---------- FAST sheet cache ----------------------------------- */
const _CACHE = {};
function sheet_(name) {
  if (!_CACHE[name]) {
    Logger.log(`sheet_(): caching sheet "${name}"`);
    const s = SS.getSheetByName(name);
    if (!s) {
        Logger.log(`sheet_(): Sheet "${name}" not found!`);
        SpreadsheetApp.getUi().alert(`Error: Critical sheet "${name}" not found. Please check configuration.`);
        throw new Error(`Sheet not found: ${name}`);
    }
    _CACHE[name] = s;
  }
  return _CACHE[name];
}

/* ---------- HELPER FUNCTIONS ----------------------------------- */

function validateEmail(email) {
  if (!email) return false;
  const re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(String(email).toLowerCase());
}

function calcAge_(dob) {
  Logger.log(`calcAge_(): dob="${dob}"`);
  if (!(dob instanceof Date) || isNaN(dob.getTime())) { return ''; }
  const today = new Date();
  let years = today.getFullYear() - dob.getFullYear();
  const m = today.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) { years--; }
  return years >= 0 ? years : '';
}

function nameSimilarity(a, b) {
  a = (a || '').trim().toLowerCase();
  b = (b || '').trim().toLowerCase();
  if (a.length === 0 && b.length === 0) return 1;
  if (a.length === 0 || b.length === 0) return 0;
  const dist = levenshteinDistance(a, b);
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - (dist / maxLen);
}

function levenshteinDistance(s, t) {
  const rows = t.length + 1;
  const cols = s.length + 1;
  const mat = Array.from({ length: rows }, () => Array(cols).fill(0));
  for (let i = 0; i < rows; i++) mat[i][0] = i;
  for (let j = 0; j < cols; j++) mat[0][j] = j;
  for (let i = 1; i < rows; i++) {
    for (let j = 1; j < cols; j++) {
      const cost = (t.charAt(i - 1) === s.charAt(j - 1)) ? 0 : 1;
      mat[i][j] = Math.min( mat[i - 1][j] + 1, mat[i][j - 1] + 1, mat[i - 1][j - 1] + cost );
    }
  }
  return mat[rows - 1][cols - 1];
}

function findRowByUid_(sheet, uid, colIndex, headerRows) {
  if (!sheet) return 0;
  
  const lastRow = sheet.getLastRow();
  
  // Guard clause: If there is no data below the headers, stop immediately.
  if (lastRow <= headerRows) return 0;

  // Now it is safe to define the range
  // (lastRow - headerRows) will always be >= 1 here
  const data = sheet.getRange(headerRows + 1, colIndex, lastRow - headerRows, 1).getValues().flat();
  
  const idx = data.indexOf(uid);
  return (idx === -1) ? 0 : idx + headerRows + 1;
}

function isPlaceholderEmail(email) {
  if (!email) return true;
  if (email === CONFIG.DEFAULT_TL_EMAIL_INFO_NOTE) return true;
  return !validateEmail(email);
}

function isPlaceholderResponsibleParty(responsible) {
  if (!responsible) return true;
  if (responsible === CONFIG.DEFAULT_TL_RESPONSIBLE_PARTY_NOTE) return true;
  return false;
}

function isValidPatientAndEmail(patient, email) {
  return !!patient && validateEmail(email);
}