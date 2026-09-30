/* ================================================================
 * INTAKE TAB CREATION & MANAGEMENT
 * ================================================================ */

/* ---------- Menu entry points ---------------------------------- */

function openIntakeCreator() {
  showIntakeDialog_({ mode: 'create' }, 'New Intake Tab');
}

function openFormAnswerImporter() {
  const sh = SS.getActiveSheet();
  if (SYSTEM_SHEET_NAMES.includes(sh.getName()) || !sh.getRange(CONFIG.INTAKE_CELL_UID).getValue()) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Open an intake tab first', 'Go to the intake tab you want to fill in, then choose this menu item again.', ui.ButtonSet.OK);
    return;
  }
  showIntakeDialog_({ mode: 'attach', sheetName: sh.getName() }, 'Import Form Answers');
}

function showIntakeDialog_(context, title) {
  const tpl = HtmlService.createTemplateFromFile('IntakeCreator');
  tpl.context = JSON.stringify(context).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(tpl.evaluate().setWidth(620).setHeight(640), title);
}

/* ---------- Called from IntakeCreator.html --------------------- */

/** Telephone Log calls without an intake tab (newest first), each with its form status. */
function getOpenCalls() {
  const responses = readFormResponses_();
  const taken = uidsWithIntake_();
  return readOpenCalls_(taken).map(call => {
    const m = findFormMatches_(call, responses, taken);
    call.form = m.exact ? 'exact' : (m.suggestions.length ? 'possible' : 'none');
    return call;
  });
}

/** Form response matches for a call, a new client or an existing intake tab. */
function getFormMatches(person) {
  const m = findFormMatches_(person, readFormResponses_(), uidsWithIntake_());
  return {
    exact:       m.exact ? responseSummary_(m.exact) : null,
    suggestions: m.suggestions.map(responseSummary_),
    unlinked:    m.unlinked.slice().sort((a, b) => timeOf_(b.timestamp) - timeOf_(a.timestamp)).map(responseSummary_)
  };
}

/** The client on an existing intake tab (for "Import Form Answers"). */
function getIntakeTabPerson(sheetName) {
  const sh = SS.getSheetByName(sheetName);
  if (!sh) throw new Error(`The tab "${sheetName}" no longer exists.`);
  const v = ref => String(sh.getRange(ref).getValue() || '').trim();
  return {
    sheetName,
    uid:              v(CONFIG.INTAKE_CELL_UID),
    patientName:      v(CONFIG.INTAKE_CELL_PATIENT_NAME),
    responsibleParty: v(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY),
    email:            v(CONFIG.INTAKE_CELL_EMAIL),
    phone:            v(CONFIG.INTAKE_CELL_PHONE)
  };
}

/**
 * Creates an intake tab.
 * request: { source: 'call', uid } or { source: 'new', patientName, responsibleParty, email, phone },
 *          plus responseRow (the form response to import) or skipImport.
 */
function createIntake(request) {
  let person;
  if (request.source === 'call') {
    person = readCall_(request.uid);
    if (!person) throw new Error('That call is no longer in the Telephone Log.');
    ensureRegistryEntry_(person.uid, person.patientName, person.email);
  } else {
    const t = v => String(v || '').trim();
    person = { patientName: t(request.patientName), responsibleParty: t(request.responsibleParty), email: t(request.email), phone: t(request.phone) };
    if (!person.patientName) throw new Error('Please enter the patient name.');
    if (person.email && !validateEmail(person.email)) throw new Error(`"${person.email}" is not a valid email address.`);
    person.uid = getOrCreateUID_(person.patientName, person.responsibleParty, person.email);
  }
  markHasIntake_(person.uid, true);

  const sh = createIntakeTab_(person);

  syncTelephoneLog_({
    uid: person.uid, patientName: person.patientName, email: person.email,
    responsibleParty: person.responsibleParty, phone: person.phone,
    addToWaitingList: false, isInitialCreation: true
  });

  let row = request.responseRow || null;
  if (!row && !request.skipImport) {
    const exact = findFormMatches_(person, readFormResponses_(), new Set()).exact;
    if (exact) row = exact.row;
  }
  if (row) importFormResponseRow_(sh, row, person.uid);

  SS.setActiveSheet(sh);
  return { sheetName: sh.getName(), imported: !!row };
}

/** Imports a form response into an existing intake tab, filling only empty cells. */
function importFormAnswers(sheetName, row) {
  const sh = SS.getSheetByName(sheetName);
  if (!sh) throw new Error(`The tab "${sheetName}" no longer exists.`);
  const uid = String(sh.getRange(CONFIG.INTAKE_CELL_UID).getValue() || '').trim();
  if (!uid) throw new Error(`The tab "${sheetName}" has no UID in ${CONFIG.INTAKE_CELL_UID}.`);

  importFormResponseRow_(sh, row, uid, { onlyEmpty: true });
  return { sheetName };
}

/* ---------- Helpers -------------------------------------------- */

function createIntakeTab_({ uid, patientName, responsibleParty, email, phone }) {
  const sh = sheet_(CONFIG.TEMPLATE).copyTo(SS).setName(uniqueSheetName_(patientName || `Intake ${uid.slice(0, 6)}`));
  placeAfterSystemSheets_(sh);
  sh.showSheet();
  setIntakeTabColor_(sh);

  sh.getRange(CONFIG.INTAKE_CELL_UID).setValue(uid).setFontWeight('normal').setFontSize(12);
  sh.getRange(CONFIG.INTAKE_CELL_DATE).setValue(new Date()).setNumberFormat('MM/dd/yyyy');
  if (patientName)      sh.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).setValue(patientName);
  if (responsibleParty) sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).setValue(responsibleParty);
  if (email)            sh.getRange(CONFIG.INTAKE_CELL_EMAIL).setValue(email);
  if (phone)            sh.getRange(CONFIG.INTAKE_CELL_PHONE).setValue(phone);
  return sh;
}

/** "Jane Smith", or "Jane Smith (2)" if that tab name is taken. */
function uniqueSheetName_(base) {
  const name = String(base).trim().slice(0, 90) || 'Intake';
  let candidate = name;
  let n = 2;
  while (SS.getSheetByName(candidate)) candidate = `${name} (${n++})`;
  return candidate;
}

function tlCallWidth_() {
  return Math.max(CONFIG.TL_COL_UID, CONFIG.TL_COL_DATE, CONFIG.TL_COL_RESPONSIBLE,
                  CONFIG.TL_COL_PATIENT_NAME, CONFIG.TL_COL_PHONE, CONFIG.TL_COL_EMAIL);
}

function readOpenCalls_(uidsWithIntake) {
  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  const first = CONFIG.TL_HEADER_ROWS + 1;
  if (tl.getLastRow() < first) return [];

  return tl.getRange(first, 1, tl.getLastRow() - first + 1, tlCallWidth_()).getValues()
    .map(r => callFromRow_(r))
    .filter(c => c.uid && !uidsWithIntake.has(c.uid.toLowerCase()))
    .reverse();
}

function readCall_(uid) {
  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  const row = findRowByUid_(tl, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);
  return row ? callFromRow_(tl.getRange(row, 1, 1, tlCallWidth_()).getValues()[0]) : null;
}

/** A Telephone Log row as a person, with the new-row placeholder text removed. */
function callFromRow_(r) {
  const text = col => String(r[col - 1] || '').trim();
  const date = r[CONFIG.TL_COL_DATE - 1];
  const rp = text(CONFIG.TL_COL_RESPONSIBLE);
  const email = text(CONFIG.TL_COL_EMAIL);
  return {
    uid:              text(CONFIG.TL_COL_UID),
    date:             timeOf_(date) ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yy') : '',
    patientName:      text(CONFIG.TL_COL_PATIENT_NAME),
    responsibleParty: rp === CONFIG.DEFAULT_TL_DISABLE_FORM_NOTE ? '' : rp,
    email:            validateEmail(email) ? email : '',
    phone:            text(CONFIG.TL_COL_PHONE)
  };
}

/** What the dialog shows for a form response. */
function responseSummary_(r) {
  return {
    row:       r.row,
    date:      timeOf_(r.timestamp) ? Utilities.formatDate(r.timestamp, Session.getScriptTimeZone(), 'MM/dd/yy') : '',
    name:      r.name,
    typedName: r.typedName,
    email:     r.email,
    phone:     r.phone,
    otherCall: UID_PATTERN.test(r.uid),   // carries another call's UID (that call has no intake yet)
    reasons:   r.reasons || []
  };
}

function findIntakeSheetsByUid_(uid) {
  return SS.getSheets().filter(sh => {
    if (SYSTEM_SHEET_NAMES.includes(sh.getName())) return false;
    return sh.getRange(CONFIG.INTAKE_CELL_UID).getValue() === uid;
  });
}

function setIntakeTabColor_(sheet) {
  const notInterested = sheet.getRange(CONFIG.INTAKE_CELL_NOT_INTERESTED).getValue() === true;
  const active     = sheet.getRange(CONFIG.INTAKE_CELL_ACTIVE    ).getValue() === true;
  const spotFound  = sheet.getRange(CONFIG.INTAKE_CELL_SPOT_FOUND).getValue() === true;
  const intakeCallCompleted = sheet.getRange(CONFIG.INTAKE_CELL_CALL_COMPLETED).getValue() === true;

  if (notInterested)            { sheet.setTabColor(COLORS.RED);          }
  else if (active)              { sheet.setTabColor(COLORS.GREEN);       }
  else if (spotFound)           { sheet.setTabColor(COLORS.LIGHT_GREEN); }
  else if (intakeCallCompleted) { sheet.setTabColor(COLORS.LIGHT_YELLOW);}
  else                          { sheet.setTabColor(null); }
}

function renameIntakeTabsForUid_(uid, patientName) {
  if (!patientName) return;
  const targetSheets = findIntakeSheetsByUid_(uid);
  targetSheets.forEach(sh => {
    if (sh.getName() === patientName) return;
    let newName   = patientName;
    let counter   = 2;
    while (SS.getSheetByName(newName) &&
           SS.getSheetByName(newName).getSheetId() !== sh.getSheetId()) {
      newName = `${patientName} (${counter++})`;
    }
    sh.setName(newName);
  });
}

function placeAfterSystemSheets_(sh) {
  const ss = SpreadsheetApp.getActive();
  const lastSysIdx = ss.getSheets()
                       .filter(s => SYSTEM_SHEET_NAMES.includes(s.getName()))
                       .reduce((max, s) => Math.max(max, s.getIndex()), 0);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(lastSysIdx + 1);
}