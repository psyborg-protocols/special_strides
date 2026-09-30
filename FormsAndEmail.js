/* ================================================================
 * FORMS & EMAILS
 * ================================================================ */

/* ---------- FORM RESPONSE LOOKUP & MATCHING --------------------- */

const UID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const NAME_MATCH_MIN = 0.85;     // name similarity needed to suggest a response
const MAX_SUGGESTIONS = 5;

/** All Form Responses rows as objects (row = sheet row number). */
function readFormResponses_() {
  const fs = SS.getSheetByName(CONFIG.FORM_RESPONSES);
  if (!fs || fs.getLastRow() <= 1) return [];

  const rows = fs.getRange(2, 1, fs.getLastRow() - 1, CONFIG.FR_TOTAL_COLUMNS).getValues();
  return rows.map((r, i) => {
    const uid = String(r[CONFIG.FR_COL_UID - 1] || '').trim();
    return {
      row:       i + 2,
      timestamp: r[CONFIG.FR_COL_TIMESTAMP - 1],
      name:      String(r[CONFIG.FR_COL_RESPONSIBLE_PARTY - 1] || '').trim(),
      uid,
      typedName: uid && !UID_PATTERN.test(uid) ? uid : '',  // families without the prefilled link sometimes type a name here
      email:     String(r[CONFIG.FR_COL_EMAIL - 1] || '').trim(),
      phone:     String(r[CONFIG.FR_COL_PHONE - 1] || '').trim()
    };
  }).filter(x => x.timestamp !== '' || x.uid || x.name);
}

/** UIDs (lower-case) that already have an intake tab, per the registry. */
function uidsWithIntake_() {
  const reg = SS.getSheetByName(CONFIG.REGISTRY);
  const taken = new Set();
  if (!reg || reg.getLastRow() <= 1) return taken;
  reg.getRange(2, 1, reg.getLastRow() - 1, CONFIG.REG_COL_HAS_INTAKE).getValues().forEach(r => {
    if (r[CONFIG.REG_COL_HAS_INTAKE - 1] === true && r[CONFIG.REG_COL_UID - 1]) taken.add(String(r[CONFIG.REG_COL_UID - 1]).trim().toLowerCase());
  });
  return taken;
}

/**
 * Finds the form response for a person.
 * person: { uid, patientName, responsibleParty, email, phone }
 * Returns { exact, suggestions, unlinked } — exact is the response carrying the person's UID (latest wins);
 * suggestions are strong email / phone / name matches among responses not already linked to another intake;
 * unlinked is every response not linked to an intake (for browsing).
 */
function findFormMatches_(person, responses, takenUids) {
  const uid = String(person.uid || '').trim().toLowerCase();
  const exact = uid ? responses.filter(r => r.uid.toLowerCase() === uid).pop() || null : null;

  const available = responses.filter(r => r !== exact && !(r.uid && r.uid.toLowerCase() !== uid && takenUids.has(r.uid.toLowerCase())));
  if (exact) return { exact, suggestions: [], unlinked: available };

  const email = String(person.email || '').trim().toLowerCase();
  const phone = phoneKey_(person.phone);
  const names = [person.patientName, person.responsibleParty].filter(Boolean);

  const suggestions = [];
  available.forEach(r => {
    const reasons = [];
    let score = 0;
    if (validateEmail(email) && r.email.toLowerCase() === email) { reasons.push('Same email'); score = Math.max(score, 100); }
    if (phone && phoneKey_(r.phone) === phone) { reasons.push('Same phone'); score = Math.max(score, 90); }

    let bestName = 0;
    names.forEach(n => [r.name, r.typedName].forEach(rn => { bestName = Math.max(bestName, nameScore_(n, rn)); }));
    if (bestName >= NAME_MATCH_MIN) {
      reasons.push(bestName === 1 ? 'Same name' : `Similar name (${Math.round(bestName * 100)}%)`);
      score = Math.max(score, Math.round(bestName * 85));
    }

    if (reasons.length) suggestions.push(Object.assign({}, r, { reasons, score: score + (reasons.length - 1) * 10 }));
  });

  suggestions.sort((a, b) => (b.score - a.score) || (timeOf_(b.timestamp) - timeOf_(a.timestamp)));
  return { exact: null, suggestions: suggestions.slice(0, MAX_SUGGESTIONS), unlinked: available };
}

/** Name similarity 0–1, ignoring case, accents, punctuation and word order ("Smith, Jane" = "Jane Smith"). */
function nameScore_(a, b) {
  const words = s => {
    let t = String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
    if (t.includes(',')) t = t.split(',').reverse().join(' ');
    return t.replace(/[^a-z\s]/g, ' ').split(/\s+/).filter(Boolean);
  };
  const x = words(a), y = words(b);
  if (!x.length || !y.length) return 0;
  return Math.max(
    nameSimilarity(x.join(' '), y.join(' ')),
    nameSimilarity(x.slice().sort().join(' '), y.slice().sort().join(' '))
  );
}

function phoneKey_(phone) {
  const digits = String(phone || '').replace(/\D/g, '');
  return digits.length >= 7 ? digits.slice(-10) : '';
}

function timeOf_(d) {
  return d && typeof d.getTime === 'function' && !isNaN(d.getTime()) ? d.getTime() : 0;
}

/** Puts the intake's UID on a form response so every UID-based lookup finds it. Keeps what the family typed as a note. */
function linkFormResponse_(row, uid) {
  const fs = sheet_(CONFIG.FORM_RESPONSES);
  const cell = fs.getRange(row, CONFIG.FR_COL_UID);
  const current = String(cell.getValue() || '').trim();
  if (current === uid) return;

  if (current) {
    const when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
    cell.setNote(`Originally entered: "${current}"\nLinked to an intake tab on ${when}.`);
  }
  cell.setValue(uid);
}

/** Links a form response to the intake, pastes its answers into the tab and marks the form as submitted. */
function importFormResponseRow_(sheet, row, uid, options = {}) {
  const fs = sheet_(CONFIG.FORM_RESPONSES);
  const answers = fs.getRange(row, 1, 1, CONFIG.FR_TOTAL_COLUMNS).getValues()[0];

  linkFormResponse_(row, uid);
  pasteFormAnswersToIntakeStructured_(sheet, { answers }, options);
  markIntakeFormSubmitted_(uid);
}

/** Ticks "form submitted" in the Telephone Log and Email History for this UID. */
function markIntakeFormSubmitted_(uid) {
  const tl = SS.getSheetByName(CONFIG.TELEPHONE_LOG);
  const tlRow = findRowByUid_(tl, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);
  if (tlRow) tl.getRange(tlRow, CONFIG.TL_COL_FORM_SUBMITTED).setValue(true);

  const hist = SS.getSheetByName(CONFIG.HISTORY);
  if (!hist) return;
  const data = hist.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.HISTORY_COL_UID - 1] === uid && data[i][CONFIG.HISTORY_COL_FORM - 1] === 'INTAKE') {
      hist.getRange(i + 1, CONFIG.HISTORY_COL_SUBMITTED).setValue(true);
    }
  }
}

/** Checks "Responded to inquiry" (column B) on every form response with this UID. */
function markFormResponseResponded_(uid) {
  const fs = SS.getSheetByName(CONFIG.FORM_RESPONSES);
  if (!uid || !fs || fs.getLastRow() <= 1) return;

  const uidCol = fs.getRange(2, CONFIG.FR_COL_UID, fs.getLastRow() - 1, 1).getValues().flat();
  uidCol.forEach((u, i) => {
    if (u === uid) fs.getRange(i + 2, CONFIG.FR_COL_RESPONDED).setValue(true);
  });
}

/** options.onlyEmpty: fill only empty cells (for tabs staff have already worked on). */
function pasteFormAnswersToIntakeStructured_(sheet, qa, options = {}) {
  const val  = colIndex => qa.answers[colIndex - 1] || '';
  const join = colIndexes => colIndexes.map(val).filter(Boolean).join('\n');

  const MAP = {
    [CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY]: val(CONFIG.FR_COL_RESPONSIBLE_PARTY),
    [CONFIG.INTAKE_CELL_PHONE]            : val(CONFIG.FR_COL_PHONE),
    [CONFIG.INTAKE_CELL_EMAIL]            : val(CONFIG.FR_COL_EMAIL),
    [CONFIG.INTAKE_CELL_DOB]              : val(CONFIG.FR_COL_DOB),
    [CONFIG.INTAKE_CELL_POTENTIAL_SERVICE]: join([CONFIG.FR_COL_INTEREST_CHILD, CONFIG.FR_COL_INTEREST_ADULT]),
    [CONFIG.INTAKE_CELL_DIAGNOSIS_NOTES]  : join([CONFIG.FR_COL_CHILD_DIAGNOSIS, CONFIG.FR_COL_ADULT_DIAGNOSIS]),
    [CONFIG.INTAKE_CELL_MED_HISTORY]      : join([CONFIG.FR_COL_CHILD_MED_HISTORY, CONFIG.FR_COL_ADULT_MED_HISTORY]),
    [CONFIG.INTAKE_CELL_CLASSROOM]        : val(CONFIG.FR_COL_CHILD_CLASSROOM),
    [CONFIG.INTAKE_CELL_THERAPIES]        : join([CONFIG.FR_COL_CHILD_THERAPY_SCHOOL, CONFIG.FR_COL_CHILD_THERAPY_OUTPATIENT, CONFIG.FR_COL_ADULT_THERAPY_OUTPATIENT]),
    [CONFIG.INTAKE_CELL_FUNCTION_LEVEL]   : join([CONFIG.FR_COL_CHILD_FUNCTION_LEVEL, CONFIG.FR_COL_ADULT_FUNCTION_LEVEL]),
    [CONFIG.INTAKE_CELL_GOALS]            : join([CONFIG.FR_COL_CHILD_GOALS, CONFIG.FR_COL_ADULT_GOALS]),
    [CONFIG.INTAKE_CELL_ADDL_INFO]        : join([CONFIG.FR_COL_CHILD_ADDL_INFO, CONFIG.FR_COL_ADULT_ADDL_INFO]),
    [CONFIG.INTAKE_CELL_BEST_CONTACT]     : join([CONFIG.FR_COL_CHILD_BEST_CONTACT, CONFIG.FR_COL_ADULT_BEST_CONTACT])
  };

  // --- NEW LOGIC: Adult Detection ---
  // If the "Child Interest" field is empty, we assume this is an adult filling it out for themselves.
  // In that case, we overwrite the Patient Name with the Responsible Party name.
  if (!val(CONFIG.FR_COL_INTEREST_CHILD)) {
    MAP[CONFIG.INTAKE_CELL_PATIENT_NAME] = val(CONFIG.FR_COL_RESPONSIBLE_PARTY);
  }
  // ----------------------------------

  Object.entries(MAP).forEach(([cell, value]) => {
    if (!value) return;
    const range = sheet.getRange(cell);
    if (options.onlyEmpty && range.getValue() !== '') return;
    range.setValue(value);
  });

  const marker = sheet.getRange('B40');
  if (options.onlyEmpty && marker.getValue() !== '') return;
  marker.setValue('Google-Form answers imported automatically').setFontStyle('italic').setFontSize(9).setBackground('#f5f5ff');
}

function sendForm_(formKey, { uid, email, patient = '', responsible = '', apptDate = '' }) {
  // 1. Get merged configuration (Code logic + Sheet URL)
  const cfg = getFormConfig_(formKey);
  if (!cfg) { 
    Logger.log(`sendForm_: Configuration not found for "${formKey}"`); 
    SpreadsheetApp.getUi().alert(`Error: Form configuration for "${formKey}" is missing. Check the System_Form_Links sheet.`);
    return false; 
  }

  const displayName = cfg.displayName || formKey;
  const confirm = confirmSend_(displayName, responsible || patient || email);
  if (!confirm.ok) return false;
  responsible = confirm.responsible;

  const history = sheet_(CONFIG.HISTORY);
  const rows = history.getDataRange().getValues();
  if (rows.some(r => r[CONFIG.HISTORY_COL_UID-1] === uid && r[CONFIG.HISTORY_COL_FORM-1] === formKey && r[CONFIG.HISTORY_COL_SENT-1] === true)) {
    return false; // Already sent
  }

  // 2. SMART URL CONSTRUCTION
  // Auto-protect: specific replacement ensures patients always get the VIEW link
  // even though we store the EDIT link in the sheet for the rollover script.
  let link = cfg.formUrl.replace(/\/edit.*$/, '/viewform');
  
  // Only append UID if a specific parameter name (e.g., 'entry.1234') is provided
  if (cfg.uidEntry) {
    // Check if the URL already has a '?' to decide between '?' and '&'
    const separator = link.includes('?') ? '&' : '?';
    link = `${link}${separator}${cfg.uidEntry}=${encodeURIComponent(uid)}`;
  }

  const html = cfg.mail.body({link, patient, responsible, apptDate});
  MailApp.sendEmail({to: email, subject: cfg.mail.subject, htmlBody: html});

  let found = false;
  for (let i=1;i<rows.length;i++){
    if (rows[i][CONFIG.HISTORY_COL_UID-1] === uid && rows[i][CONFIG.HISTORY_COL_FORM-1] === formKey) {
      history.getRange(i+1, CONFIG.HISTORY_COL_DATE).setValue(new Date());
      history.getRange(i+1, CONFIG.HISTORY_COL_SENT).setValue(true);
      found = true; break;
    }
  }
  if (!found) history.appendRow([uid, formKey, new Date(), email, true, false]);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`${displayName} sent to ${email}`, 'Email sent', 5);
  return true;
}

function confirmSend_(displayName, responsible) {
  const ui = SpreadsheetApp.getUi();
  responsible = (responsible || '').trim();

  if (responsible) {
    const yesNo = ui.alert(`Send ${displayName}?`, `Would you like to send “${displayName}” to “${responsible}”?`, ui.ButtonSet.YES_NO);
    return (yesNo === ui.Button.YES) ? { ok: true, responsible } : { ok: false };
  }

  const prompt = ui.prompt(`Send ${displayName}`, 'Enter the recipient’s name:', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return { ok: false };
  const newName = prompt.getResponseText().trim();
  return newName ? { ok: true, responsible: newName } : { ok: false };
}