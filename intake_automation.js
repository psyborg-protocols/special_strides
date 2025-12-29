// =================================================================
//  SPECIAL STRIDES – 2025 INTAKE AUTOMATION SUITE
//  UID management • Intake-tab creator • Waiting-List & Telephone-Log sync
//  Multi-directional Patient Name & Email sync
//  editHandler trigger + one-time Google-Form e-mail
// =================================================================

/* ---------- GLOBAL CONFIG -------------------------------------- */
const SS = SpreadsheetApp.getActive();

const CONFIG = {
  // ---- Sheet names
  TELEPHONE_LOG:  'Telephone_Log_2025',
  WAITING_LIST:   'Waiting_List_2025',
  REGISTRY:       'Patient_Registry',
  TEMPLATE:       'Intake_Template',
  HISTORY:        'Email_History_2025',
  FORM_RESPONSES: 'Form Responses',

  // ---- Google Form
  FORM_URL: 'https://docs.google.com/forms/d/e/1FAIpQLScfdkxGN5ZAGittteFpYRn1y2_nCvi0fgxgEZfMCsyHlKl5bw/viewform?usp=pp_url',

  // ---- Patient Registry column indexes (1-based)
  REG_COL_UID:                1, // Column A
  REG_COL_PATIENT_NAME:       2, // Column B
  REG_COL_EMAIL:              3, // Column C
  REG_COL_HAS_INTAKE:         5, // Column E

  // ---- Email history column indexes (1-based)
  HISTORY_COL_UID          : 1, // A
  HISTORY_COL_FORM         : 2, // B
  HISTORY_COL_DATE         : 3, // C
  HISTORY_COL_EMAIL        : 4, // D
  HISTORY_COL_SENT         : 5, // E
  HISTORY_COL_SUBMITTED    : 6,  // F

  // ---- Waiting List column indexes (1-based)
  WL_HEADER_ROWS:             2, // Header rows in Waiting List
  WL_COL_UID:                 1, // Column A
  WL_COL_ACTIVE:              2, // Column B ('Active' from Intake D3)
  WL_COL_SPOT_FOUND:          3, // Column C ('Spot Found' from Intake D2)
  WL_COL_RESPONSIBLE:         4, // Column D  
  WL_COL_PATIENT_NAME:        5, // Column E
  WL_COL_DIAGNOSIS_NOTES:     6, // Column F
  WL_COL_AGE:                 7, // Column G (calculated from DOB)
  WL_COL_PHONE:               8, // Column H
  WL_COL_EMAIL:               9, // Column I
  WL_COL_POTENTIAL_SERVICE:   12, // Column L
  WL_SOCIAL_STRIDES:          13, // Column M 
  WL_NOT_INTERESTED:          15, // Column o
  WL_INTAKE_COMPLETED:        16, // Column P

  
  // ---- Telephone-Log column indexes (1-based)
  TL_HEADER_ROWS:            2,  // Header rows in Telephone Log
  TL_COL_UID:                1,  // A
  TL_COL_CALL_OUTCOME:       2,  // B
  TL_COL_DATE:               3,  // C
  TL_COL_FORM_SUBMITTED:     4,  // E 
  TL_COL_FORM_SENT:          5,  // D 
  TL_COL_RESPONSIBLE:        6,  // F
  TL_COL_PATIENT_NAME:       7,  // G
  TL_COL_PHONE:              8,  // H
  TL_COL_EMAIL:              9,  // I
  TL_COL_INFORMATION:       10,  // J
  TL_COL_TELEVISIT_SCHED:   11,  // K
  TL_COL_TELEVISIT_COMPLETE:12,  // L
  TL_COL_CONTACT_METHOD:    13,  // M
  TL_COL_WAITLIST_FLAG:     14,  // N
  TL_COL_ONSCHEDULE:        15,  // O
  TL_COL_NOT_INTERESTED:    16,  // P

  // ---- TL placeholders
  DEFAULT_TL_CALL_OUTCOME_PLACEHOLDER:'📝 FILL ME:',
  DEFAULT_TL_DISABLE_FORM_NOTE:      '<- Click here to disable automatic sending',
  DEFAULT_TL_EMAIL_INFO_NOTE:        'Adding an email in the email column will send a Google Form automatically',


  // ---- Intake Sheet Cell References (absolute A1 notation) ------------
  INTAKE_CELL_UID:                'A1',
  INTAKE_CELL_DATE:               'B2',
  INTAKE_CELL_RESPONSIBLE_PARTY:  'B3',
  INTAKE_CELL_PATIENT_NAME:       'B4',   // adjust when a real “patient name” field exists
  INTAKE_CELL_PHONE:              'B7',
  INTAKE_CELL_EMAIL:              'B8',
  INTAKE_CELL_DOB:                'B9',

  INTAKE_CELL_DIAGNOSIS_NOTES:    'B10',  // primary Dx (“notes” column on WL)
  INTAKE_CELL_MED_HISTORY:        'B12',

  INTAKE_CELL_CLASSROOM:          'B18',
  INTAKE_CELL_THERAPIES:          'B19',
  INTAKE_CELL_FUNCTION_LEVEL:     'B15',
  INTAKE_CELL_GOALS:              'B39',
  INTAKE_CELL_ADDL_INFO:          'B29',
  INTAKE_CELL_BEST_CONTACT:       'B30',

  INTAKE_CELL_TELEHEALTH_DATE:    'B32', // Telehealth appointment date
  INTAKE_CELL_TELEHEALTH_TIME:    'C32', // Telehealth appointment time
  INTAKE_CELL_POTENTIAL_SERVICE:  'B33',

  // status / flags
  INTAKE_CELL_RECREATIONAL:       'D1', // Recreational Therapy checkbox
  INTAKE_CELL_CALL_COMPLETED:     'D2',
  INTAKE_CELL_SPOT_FOUND:         'D3',
  INTAKE_CELL_ACTIVE:             'D4',
  INTAKE_CELL_NOT_INTERESTED:     'D5',
  INTAKE_CELL_TELEHEALTH_LINK:    'D32',
  INTAKE_CELL_FINANCIAL_AID:      'D36', // 2025 Financial Aid Form link

  // “Note to Waiting-List” cell
  INTAKE_TAB_WL_NOTE:             'E2',
  // “Note For Forms Sent” cell
  INTAKE_TAB_FORMS_SENT_NOTE:     'E4',
  // “Note For Telehealth Link” cell
  INTAKE_TAB_TELEHEALTH_LINK_NOTE: 'E32',
  // “Note For Financial Aid” cell
  INTAKE_TAB_FINANCIAL_AID_NOTE:  'E36',

  // ---- Form-Responses column indexes (1-based) ------------------------
  FR_COL_TIMESTAMP                :  1, // A
  FR_COL_RESPONSIBLE_PARTY        :  2, // B  (“Your Name”)
  FR_COL_UID                      :  3, // C
  FR_COL_EMAIL                    :  4, // D
  FR_COL_PHONE                    :  5, // E
  FR_COL_DOB                      :  6, // F  (Patient’s Date of Birth)
  FR_COL_PATIENT_TYPE             :  7, // G  (adult / child)
  FR_COL_INTEREST_CHILD           :  8, // H
  FR_COL_CHILD_DIAGNOSIS          :  9, // I
  FR_COL_CHILD_MED_HISTORY        : 10, // J
  FR_COL_CHILD_CLASSROOM          : 11, // K
  FR_COL_CHILD_THERAPY_SCHOOL     : 12, // L
  FR_COL_CHILD_THERAPY_OUTPATIENT : 13, // M
  FR_COL_CHILD_FUNCTION_LEVEL     : 14, // N
  FR_COL_CHILD_GOALS              : 15, // O
  FR_COL_CHILD_ADDL_INFO          : 16, // P
  FR_COL_CHILD_BEST_CONTACT       : 17, // Q
  FR_COL_INTEREST_ADULT           : 18, // R
  FR_COL_ADULT_DIAGNOSIS          : 19, // S
  FR_COL_ADULT_MED_HISTORY        : 20, // T
  FR_COL_ADULT_THERAPY_OUTPATIENT : 21, // U
  FR_COL_ADULT_FUNCTION_LEVEL     : 22, // V
  FR_COL_ADULT_GOALS              : 23, // W
  FR_COL_ADULT_ADDL_INFO          : 24, // X
  FR_COL_ADULT_BEST_CONTACT       : 25, // Y

};

/* ---------- TAB-COLOUR CONSTANTS -------------------------------- */
const COLORS = {
  GREEN:        '#00B050', // Active   → green
  LIGHT_GREEN:  '#92D050', // SpotFound w/out Active
  LIGHT_YELLOW: '#f1c232',  // default on creation
  RED         : '#EA4335'   // Not-interested
};

// ---- For editHandler identification of intake sheets and relevant cells
const SYSTEM_SHEET_NAMES = [
  CONFIG.TELEPHONE_LOG,
  CONFIG.WAITING_LIST,
  CONFIG.REGISTRY,
  CONFIG.TEMPLATE,
  CONFIG.HISTORY
];

// Map of intake sheet cells to parameter names for syncWaitingList_ and syncTelephoneLog_ (original sync)
// This is primarily for the one-way sync of *other* fields from intake to WL/TL
const INTAKE_SHEET_RELEVANT_CELLS_FOR_ORIGINAL_SYNC = {
  [CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY]: 'responsibleParty',
  [CONFIG.INTAKE_CELL_PATIENT_NAME]:      'patientName',      // Also handled by new global sync
  [CONFIG.INTAKE_CELL_PHONE]:             'phone',
  [CONFIG.INTAKE_CELL_EMAIL]:             'email',            // Also handled by new global sync
  [CONFIG.INTAKE_CELL_DOB]:               'dob',
  [CONFIG.INTAKE_CELL_DIAGNOSIS_NOTES]:   'diagnosisNotes',
  [CONFIG.INTAKE_CELL_POTENTIAL_SERVICE]: 'potentialService', 
  [CONFIG.INTAKE_CELL_SPOT_FOUND]:        'spotFound',        
  [CONFIG.INTAKE_CELL_ACTIVE]:            'active'            
};


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

/* ================================================================
 * 1.  UID REGISTRY HELPERS   (v2 – accepts partial data)
 * ================================================================ */
function markHasIntake_(uid, value = true) {
  const reg = sheet_(CONFIG.REGISTRY);
  const uidCol = reg.getRange(1, CONFIG.REG_COL_UID, reg.getLastRow(), 1).getValues().flat();
  const idx = uidCol.indexOf(uid);
  if (idx === -1) return; // not found (shouldn't happen if we call ensureRegistryEntry_ earlier)
  const row = idx + 1;
  reg.getRange(row, CONFIG.REG_COL_HAS_INTAKE).setValue(value);
}

function getOrCreateUID_(patient, responsible, email) {
  Logger.log(`getOrCreateUID_ v2: patient="${patient}", email="${email}"`);

  // ---------- 1) look-up logic (try to reuse an old UID) ----------
  const reg = sheet_(CONFIG.REGISTRY);
  const rows = reg.getDataRange().getValues();

  // Fast path – valid e-mail was supplied
  if (validateEmail(email)) {
    // a) exact email + exact name
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][CONFIG.REG_COL_EMAIL - 1] === email &&
          rows[i][CONFIG.REG_COL_PATIENT_NAME - 1] === patient) {
        return rows[i][CONFIG.REG_COL_UID - 1];
      }
    }
    // b) exact email + *similar* name
    for (let r of rows) {
      if (r[CONFIG.REG_COL_EMAIL - 1] === email &&
          nameSimilarity(r[CONFIG.REG_COL_PATIENT_NAME - 1], patient) >= 0.90) {
        return r[CONFIG.REG_COL_UID - 1];
      }
    }
  }

  // ---------- 2) could not reuse → make a brand-new UID ----------
  const uid = newUid_({patient, email});
  return uid;
}

/* ----------------------------------------------------------------
 * 1b.  Make sure a UID exists (and back-fill name/e-mail if known)
 * ---------------------------------------------------------------- */
function ensureRegistryEntry_(uid, patient, email) {
  const reg = sheet_(CONFIG.REGISTRY);
  const uidCol = reg.getRange(1, CONFIG.REG_COL_UID,
                              reg.getLastRow(), 1).getValues().flat();
  let idx = uidCol.indexOf(uid);

  if (idx === -1) {                         // totally new UID → append row
    reg.appendRow([
      uid,
      patient || '',
      validateEmail(email) ? email : '',
      new Date(),
      false
    ]);
    return;
  }

  // UID already exists – improve the record if we now have better data
  idx += 1;                                 // convert to sheet-row number
  if (patient &&
      !reg.getRange(idx, CONFIG.REG_COL_PATIENT_NAME).getValue()) {
    reg.getRange(idx, CONFIG.REG_COL_PATIENT_NAME).setValue(patient);
  }
  if (validateEmail(email) &&
      !reg.getRange(idx, CONFIG.REG_COL_EMAIL).getValue()) {
    reg.getRange(idx, CONFIG.REG_COL_EMAIL).setValue(email);
  }
}

/* ----------------------------------------------------------------
 * 1c.  UID factory – always returns a fresh UID that is already
 *      present in Patient Registry.
 * ---------------------------------------------------------------- */
function newUid_({ patient = '', email = '' } = {}) {
  const uid = Utilities.getUuid();
  ensureRegistryEntry_(uid, patient, email);
  return uid;
}

/* ================================================================
 * 2.  CUSTOM MENU – create intake tab & new telephone log entry
 * ================================================================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Intake Tools')
    .addItem('✙🧍‍♂️ New Intake Tab', 'openIntakeCreator')
    .addSeparator()
    .addItem('✙📒 New Telephone Log Entry', 'addNewTelephoneLogEntry')
    .addToUi();
}

function createPatientIntakeTab() {
  Logger.log('createPatientIntakeTab(): start');
  const ui = SpreadsheetApp.getUi();

  let r = ui.prompt('Patient name:', 'Please enter the patient’s full name (LastName, FirstName).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || !r.getResponseText()) {
    Logger.log('createPatientIntakeTab(): cancelled or empty patient name.'); return;
  }
  const patient = r.getResponseText().trim();
  Logger.log(`createPatientIntakeTab(): patient="${patient}"`);

  let email = '';
  let emailPromptResponse;
  while (true) {
    emailPromptResponse = ui.prompt('Email address:', 'Please enter a contact email address. (Leave blank and click OK if no email, or click Cancel to abort)', ui.ButtonSet.OK_CANCEL);
    if (emailPromptResponse.getSelectedButton() !== ui.Button.OK) {
      Logger.log('createPatientIntakeTab(): Cancelled at email prompt.'); return;
    }
    email = emailPromptResponse.getResponseText().trim();
    if (email === "") { Logger.log('createPatientIntakeTab(): Email intentionally left blank by user.'); break; }
    if (validateEmail(email)) { Logger.log(`createPatientIntakeTab(): valid email="${email}"`); break; }
    else {
      ui.alert('Invalid Email Format', `The email "${email}" is not a valid format. Please try again, leave blank, or click Cancel.`, ui.ButtonSet.OK);
      Logger.log(`createPatientIntakeTab(): invalid email format: "${email}", re-prompting.`);
    }
  }

  const uid = getOrCreateUID_(patient, undefined, email);
  Logger.log(`createPatientIntakeTab(): resolved UID="${uid}"`);
  markHasIntake_(uid, true);

  const template = sheet_(CONFIG.TEMPLATE);
  if (!template) return;

  let newSheetName = patient;
  let counter = 1;
  while (SS.getSheetByName(newSheetName)) { newSheetName = `${patient} (${counter++})`; }
  const sh = template.copyTo(SS).setName(newSheetName);
  placeAfterSystemSheets_(sh);
  sh.showSheet();
  setIntakeTabColor_(sh);  
  SS.setActiveSheet(sh);
  Logger.log(`createPatientIntakeTab(): copied template to new sheet "${newSheetName}"`);

  sh.getRange(CONFIG.INTAKE_CELL_UID).setValue(uid);
  sh.getRange(CONFIG.INTAKE_CELL_DATE).setValue(new Date()).setNumberFormat('MM/dd/yyyy');
  sh.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).setValue(patient);
  sh.getRange(CONFIG.INTAKE_CELL_EMAIL).setValue(email);

  // Set UID cell formatting: not bold, 12 pt
  const uidCell = sh.getRange(CONFIG.INTAKE_CELL_UID);
  uidCell.setFontWeight('normal').setFontSize(12);

  const currentDate = new Date();

  syncTelephoneLog_({
    uid: uid, patientName: patient, email: email, date: currentDate, addToWaitingList:false, isInitialCreation: true
  });
  Logger.log(`createPatientIntakeTab(): initial syncTelephoneLog_ done for UID "${uid}"`);

  // ---------- Import Google-Form answers ----------
  const existingForm = getFormResponseForUid_(uid);
  if (existingForm) {
    const uiResp = ui.alert(
      'Form Already Filled',
      'A Google Form submission already exists for this UID.\n' +
      'Would you like to import the answers onto the new Intake tab?',
      ui.ButtonSet.YES_NO
    );
    if (uiResp === ui.Button.YES) {
      pasteFormAnswersToIntakeStructured_(sh, existingForm);
      Logger.log(`createPatientIntakeTab(): Imported Google-Form Q/A for UID ${uid}.`);
    }
  }

  Logger.log('createPatientIntakeTab(): end');
  ui.alert('Intake Tab Created', `New intake tab "${newSheetName}" created for ${patient} (UID: ${uid}).\nA minimal entry has been added to the Telephone Log. Name and Email will be synced across sheets.`, ui.ButtonSet.OK);
}

/* ================================================================
 * 2b.  Modern Intake-creation dialog
 * ================================================================ */
function openIntakeCreator() {
  const html = HtmlService.createHtmlOutputFromFile('IntakeCreator')
               .setWidth(420).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create / Attach Intake');
}


/* ================================================================
 * Return up to 25 recent TL rows that do **not yet** have an Intake tab
 * Label format:  05/27 – Jane Parent (Kid, Person) – jane@example.com
 * ================================================================ */
function getRecentTLWithoutIntake_(limit = 25, scanWindow = 400) {
  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  const lastRow = tl.getLastRow();
  if (lastRow <= CONFIG.TL_HEADER_ROWS) return [];

  // Build a set of UIDs that already have intake
  const reg = sheet_(CONFIG.REGISTRY);
  const regLast = reg.getLastRow();
  const hasIntake = new Set();
  if (regLast > 1) {
    const uidCol = reg.getRange(2, CONFIG.REG_COL_UID, regLast - 1, 1).getValues().flat();
    const flagCol = reg.getRange(2, CONFIG.REG_COL_HAS_INTAKE, regLast - 1, 1).getValues().flat();
    for (let i = 0; i < uidCol.length; i++) if (flagCol[i] === true) hasIntake.add(uidCol[i]);
  }

  // Calculate a tail window so we don’t read the entire sheet
  const firstDataRow = CONFIG.TL_HEADER_ROWS + 1;
  const startRow = Math.max(firstDataRow, lastRow - scanWindow + 1);
  const numRows = lastRow - startRow + 1;

  // Read a single contiguous block once.
  // (We only need cols A..I, but maxCol keeps it simple/robust.)
  const maxCol = Math.max(
    CONFIG.TL_COL_UID, CONFIG.TL_COL_DATE, CONFIG.TL_COL_RESPONSIBLE,
    CONFIG.TL_COL_PATIENT_NAME, CONFIG.TL_COL_EMAIL
  );
  const block = tl.getRange(startRow, 1, numRows, maxCol).getValues();

  // Fast date formatter (avoid getDisplayValues)
  const tz = Session.getScriptTimeZone() || 'UTC';
  const fmtDate = (v) => {
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'MM/dd');
    const s = (v == null ? '' : String(v)).trim();
    return s || '(no date)';
  };

  const out = [];
  // Walk from bottom to top so we hit “recent” entries first
  for (let i = block.length - 1; i >= 0 && out.length < limit; i--) {
    const row = block[i];
    const uid = row[CONFIG.TL_COL_UID - 1];
    if (!uid || hasIntake.has(uid)) continue;

    const dateStr = fmtDate(row[CONFIG.TL_COL_DATE - 1]);
    const resp    = row[CONFIG.TL_COL_RESPONSIBLE - 1] || '(no RP)';
    const name    = row[CONFIG.TL_COL_PATIENT_NAME - 1] || '(no name)';
    const mail    = row[CONFIG.TL_COL_EMAIL - 1] || '(no e-mail)';

    out.push({ uid, label: `${dateStr} – ${resp} (${name}) – ${mail}` });
  }
  return out;
}




/* Called from the dialog: make a new Intake tab for the chosen UID
   or fall back to the classic prompt if uid === 'NEW'               */
function createIntakeFromDialog(uid) {
  if (uid === 'NEW') {
    // fall back to the classic prompt so we can capture name/email
    createPatientIntakeTab();          // existing function kept
    return;
  }
  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  const r = findRowByUid_(tl, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);

  if (!r) { SpreadsheetApp.getUi().alert('Selected call not found'); return; }
  const patient  = tl.getRange(r, CONFIG.TL_COL_PATIENT_NAME).getValue();
  const email    = tl.getRange(r, CONFIG.TL_COL_EMAIL).getValue();
  const responsible = tl.getRange(r, CONFIG.TL_COL_RESPONSIBLE).getValue();
  const phone    = tl.getRange(r, CONFIG.TL_COL_PHONE).getValue();

  // **reuse the relaxed UID logic to be safe**
  const finalUid = uid || getOrCreateUID_(patient, responsible, email);
  ensureRegistryEntry_(finalUid, patient, email);
  markHasIntake_(finalUid, true);

  // --- clone the template ---------------------------------------
  const template = sheet_(CONFIG.TEMPLATE);
  const shName   = patient || `Intake ${finalUid.slice(0, 6)}`;
  const sh       = template.copyTo(SS).setName(shName);
  placeAfterSystemSheets_(sh);
  sh.showSheet();
  setIntakeTabColor_(sh);
  SS.setActiveSheet(sh);

  // --- fill key cells -------------------------------------------
  sh.getRange(CONFIG.INTAKE_CELL_UID).setValue(finalUid);
  sh.getRange(CONFIG.INTAKE_CELL_DATE)
    .setValue(new Date())
    .setNumberFormat('MM/dd/yyyy');
  if (patient) sh.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).setValue(patient);
  if (email)   sh.getRange(CONFIG.INTAKE_CELL_EMAIL).setValue(email);

  // --- NEW: pull Google-Form answers (if any) --------------------
  const qa = getFormResponseForUid_(finalUid);
  if (qa) {
    console.log(`[createIntakeFromDialog] Form answers found for ${finalUid}; pasting`);
    pasteFormAnswersToIntakeStructured_(sh, qa);
  } else {
    console.warn(`[createIntakeFromDialog] No form answers for ${finalUid}`);
  }

  // --- TL sync – flag as *initial* creation ----------------------
  syncTelephoneLog_({
    uid              : finalUid,
    patientName      : patient,
    email            : email,
    responsibleParty : responsible,
    phone            : phone,
    isInitialCreation: true          // <<<<<< changed to TRUE
  });

  SpreadsheetApp.getUi()
    .alert('Intake tab created & linked!\nGoogle-Form answers were ' +
           (qa ? 'imported.' : 'NOT found for this UID.'));
}

/**
 * Adds a new blank row to the Telephone Log with current date and empty checkboxes.
 */
function addNewTelephoneLogEntry() {
  Logger.log('addNewTelephoneLogEntry(): start');
  const tlSheet = sheet_(CONFIG.TELEPHONE_LOG);
  if (!tlSheet) {
    Logger.log('addNewTelephoneLogEntry(): Telephone Log sheet not found. Aborting.');
    SpreadsheetApp.getUi().alert('Error', `Telephone Log sheet "${CONFIG.TELEPHONE_LOG}" not found.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const numCols = Math.max(tlSheet.getMaxColumns(), CONFIG.TL_COL_ONSCHEDULE); // Ensure enough columns for all defined ones
  const newRowData = new Array(numCols).fill(''); // Create an array of empty strings

  /*  📌  Column defaults  */
  newRowData[CONFIG.TL_COL_UID - 1]           = newUid_();
  newRowData[CONFIG.TL_COL_CALL_OUTCOME  - 1] = CONFIG.DEFAULT_TL_CALL_OUTCOME_PLACEHOLDER;   //  Col B
  newRowData[CONFIG.TL_COL_RESPONSIBLE   - 1] = CONFIG.DEFAULT_TL_DISABLE_FORM_NOTE; //  Col E
  newRowData[CONFIG.TL_COL_EMAIL         - 1] = CONFIG.DEFAULT_TL_EMAIL_INFO_NOTE;   //  Col H



  // Set current date in Date column (Column C, index 2 for 0-based array)
  newRowData[CONFIG.TL_COL_DATE - 1] = new Date();

  // Set false for checkboxes (unchecked)
  newRowData[CONFIG.TL_COL_FORM_SENT - 1] = false;
  newRowData[CONFIG.TL_COL_FORM_SUBMITTED - 1] = false;
  newRowData[CONFIG.TL_COL_WAITLIST_FLAG - 1] = false;
  newRowData[CONFIG.TL_COL_ONSCHEDULE - 1] = false;
  newRowData[CONFIG.TL_COL_NOT_INTERESTED - 1] = false;

  tlSheet.appendRow(newRowData);
  const newRowNum = tlSheet.getLastRow();
  Logger.log(`addNewTelephoneLogEntry(): Appended new row ${newRowNum} to Telephone Log.`);

  // Format the date cell
  tlSheet.getRange(newRowNum, CONFIG.TL_COL_DATE).setNumberFormat('MM/dd/yyyy');

  // Insert checkboxes
  const checkboxCols = [
    CONFIG.TL_COL_FORM_SENT,
    CONFIG.TL_COL_FORM_SUBMITTED,
    CONFIG.TL_COL_WAITLIST_FLAG,
    CONFIG.TL_COL_ONSCHEDULE,
    CONFIG.TL_COL_NOT_INTERESTED
  ];
  checkboxCols.forEach(colIdx => {
    try {
      const cell = tlSheet.getRange(newRowNum, colIdx);
      // Only insert if not already a checkbox (e.g. from sheet template having it)
      // However, appendRow with `false` should work fine even if it's already a checkbox.
      // This is more robust if the sheet template might not have checkboxes.
      if (cell.getDataValidation() == null || cell.getDataValidation().getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
        cell.insertCheckboxes();
      }
      // Ensure it's unchecked if we explicitly set `false`
      cell.setValue(false); 
    } catch (e) {
      Logger.log(`addNewTelephoneLogEntry(): Error inserting checkbox in col ${colIdx} for new row ${newRowNum}. Error: ${e}`);
    }
  });
  
  // Activate the new row or a specific cell in it for easy editing
  tlSheet.getRange(newRowNum, CONFIG.TL_COL_RESPONSIBLE).activate(); // Example: activate RP Name cell

  Logger.log('addNewTelephoneLogEntry(): end');
}

/**
 * Find the most-recent form-response row for a given UID
 * and return its headers & answers.
 *
 * The form now has a UID column inserted at **C**,
 * so every answer column from C onward has shifted one-to-the-right.
 */
/**
 * Find the most-recent form-response row for a given UID
 * and return its headers & answers.
 *
 * The form now has a UID column inserted at **C**,
 * so every answer column from C onward has shifted one-to-the-right.
 */
function getFormResponseForUid_(uid) {
  const fs = sheet_(CONFIG.FORM_RESPONSES);
  if (!fs) {
    console.warn('Form responses sheet not found.');
    return null;
  }

  if (fs.getLastRow() <= 1) {            // header only → nothing to read
    Logger.log('Form-responses sheet has no data yet.');
    return null;
 }

  // ── sheet layout ──────────────────────────────────────
  const UID_COL         = 3;                // column C (1-based)
  const TOTAL_COLUMNS   = 25;               // A-Y (was 24)

  // headers (row 1) – include the UID header so indices match answers
  const headers = fs.getRange(1, 1, 1, TOTAL_COLUMNS)
                    .getValues()[0];
  console.log('Headers loaded:', headers);

  // entire UID column (rows 2…n) – skip the header row
  const uidCol  = fs.getRange(2, UID_COL, fs.getLastRow() - 1, 1)
                    .getValues()
                    .flat();
  console.log('UID column values:', uidCol);

  // newest match if duplicates exist
  const idx = uidCol.lastIndexOf(uid);
  console.log(`UID "${uid}" found at index:`, idx);

  if (idx === -1) {
    console.warn(`UID "${uid}" not found in responses.`);
    return null;
  }

  // answers in that same row, full width A-Y
  const answers = fs.getRange(idx + 2, 1, 1, TOTAL_COLUMNS)
                    .getValues()[0];
  console.log('Matched response row:', answers);

  return { questions: headers, answers };
}



/**
 * Copy a single Google-Form response row into the Intake sheet.
 *
 * Relies exclusively on CONFIG.FR_COL_* (source columns)
 * and CONFIG.INTAKE_CELL_* (destination cells).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  – the Intake sheet
 * @param {Object} qa                                 – response wrapper
 *        └── qa.answers  : array of raw answers in column order
 */
function pasteFormAnswersToIntakeStructured_(sheet, qa) {

  /* ── helpers ───────────────────────────────────────── */

  // read one form answer (1-based → 0-based)
  const val  = colIndex => qa.answers[colIndex - 1] || '';

  // concatenate several answers, skipping blanks
  const join = colIndexes => colIndexes.map(val).filter(Boolean).join('\n');

  /* ── Intake-cell ← Form-column mapping ─────────────── */

  const MAP = {

    [CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY]: val(CONFIG.FR_COL_RESPONSIBLE_PARTY),

    [CONFIG.INTAKE_CELL_PHONE]            : val(CONFIG.FR_COL_PHONE),
    [CONFIG.INTAKE_CELL_EMAIL]            : val(CONFIG.FR_COL_EMAIL),
    [CONFIG.INTAKE_CELL_DOB]              : val(CONFIG.FR_COL_DOB),

    [CONFIG.INTAKE_CELL_POTENTIAL_SERVICE]: join([
                                              CONFIG.FR_COL_INTEREST_CHILD,
                                              CONFIG.FR_COL_INTEREST_ADULT]),

    [CONFIG.INTAKE_CELL_DIAGNOSIS_NOTES]  : join([
                                              CONFIG.FR_COL_CHILD_DIAGNOSIS,
                                              CONFIG.FR_COL_ADULT_DIAGNOSIS]),

    [CONFIG.INTAKE_CELL_MED_HISTORY]      : join([
                                              CONFIG.FR_COL_CHILD_MED_HISTORY,
                                              CONFIG.FR_COL_ADULT_MED_HISTORY]),

    [CONFIG.INTAKE_CELL_CLASSROOM]        : val(CONFIG.FR_COL_CHILD_CLASSROOM),

    [CONFIG.INTAKE_CELL_THERAPIES]        : join([
                                              CONFIG.FR_COL_CHILD_THERAPY_SCHOOL,
                                              CONFIG.FR_COL_CHILD_THERAPY_OUTPATIENT,
                                              CONFIG.FR_COL_ADULT_THERAPY_OUTPATIENT]),

    [CONFIG.INTAKE_CELL_FUNCTION_LEVEL]   : join([
                                              CONFIG.FR_COL_CHILD_FUNCTION_LEVEL,
                                              CONFIG.FR_COL_ADULT_FUNCTION_LEVEL]),

    [CONFIG.INTAKE_CELL_GOALS]            : join([
                                              CONFIG.FR_COL_CHILD_GOALS,
                                              CONFIG.FR_COL_ADULT_GOALS]),

    [CONFIG.INTAKE_CELL_ADDL_INFO]        : join([
                                              CONFIG.FR_COL_CHILD_ADDL_INFO,
                                              CONFIG.FR_COL_ADULT_ADDL_INFO]),

    [CONFIG.INTAKE_CELL_BEST_CONTACT]     : join([
                                              CONFIG.FR_COL_CHILD_BEST_CONTACT,
                                              CONFIG.FR_COL_ADULT_BEST_CONTACT])
  };

  /* ── one-shot batch write ──────────────────────────── */

  Object.entries(MAP).forEach(([cell, value]) => {
    if (value) sheet.getRange(cell).setValue(value);
  });

  // friendly footer (static cell, not in CONFIG on purpose)
  sheet.getRange('B40')
       .setValue('Google-Form answers imported automatically')
       .setFontStyle('italic')
       .setFontSize(9)
       .setBackground('#f5f5ff');
}




/* ================================================================
 * 3.  WAITING-LIST & TELEPHONE-LOG SYNC 
 * ================================================================ */
function calcAge_(dob) {
  Logger.log(`calcAge_(): dob="${dob}"`);
  if (!(dob instanceof Date) || isNaN(dob.getTime())) { return ''; }
  const today = new Date();
  let years = today.getFullYear() - dob.getFullYear();
  const m = today.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) { years--; }
  Logger.log(`calcAge_(): calculated age=${years}`);
  return years >= 0 ? years : '';
}

function syncWaitingList_(params) {
  Logger.log(`syncWaitingList_ (original for other fields): Called with params: ${JSON.stringify(params)}`);
  const ws = sheet_(CONFIG.WAITING_LIST);
  if (!ws) return;

  let sheetRow = findRowByUid_(ws, params.uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);


  if (params.isInitialCreation) {
    Logger.log(`syncWaitingList_ (original): Initial creation for UID "${params.uid}"`);
    const intakeDateToUse = (params.intakeDate instanceof Date && !isNaN(params.intakeDate))
      ? params.intakeDate.toLocaleDateString('en-US')
      : new Date().toLocaleDateString('en-US');
    const initialData = [ 
      params.uid, 
      false, // Active (WL Col B)
      false, // Spot Found (WL Col C)
      params.responsibleParty || '',
      params.patientName || '', 
      '',   //  Diagnosis/Notes
      calcAge_(params.dob), 
      params.phone || '', 
      params.email || '',
      '',    // Communication
      intakeDateToUse, 
      params.potentialService || '', // Potential Service: PT/OT/Speech
      false, // Social Strides (WL Col M)
      "", // WL Col N - empty for now
      false,  // Not Interested (WL Col O)
      true // Intake Completed (WL Col P)
    ];
    ws.appendRow(initialData);
    sheetRow = ws.getLastRow();
    Logger.log(`syncWaitingList_ (original): Appended new row ${sheetRow} for UID "${params.uid}"`);

    ws.getRange(sheetRow, 10).setNumberFormat('MM/dd/yyyy'); // Col J - Intake Date
    [CONFIG.WL_COL_ACTIVE, CONFIG.WL_COL_SPOT_FOUND, CONFIG.WL_SOCIAL_STRIDES, CONFIG.WL_NOT_INTERESTED, CONFIG.WL_INTAKE_COMPLETED].forEach(colIdx => { // Checkboxes B, C, M, O
        try { ws.getRange(sheetRow, colIdx).insertCheckboxes(); } catch (e) { Logger.log(`WL Checkbox Error (Initial Col ${colIdx}): ${e}`);}
    });
    if (!ws.isColumnHiddenByUser(CONFIG.WL_COL_UID)) { ws.hideColumns(CONFIG.WL_COL_UID); }

  } else { 
    if (sheetRow === 0) {
      Logger.log(`syncWaitingList_ (original): Update for UID "${params.uid}", but UID not found. Aborting.`); return;
    }
    Logger.log(`syncWaitingList_ (original): Updating specific fields on row ${sheetRow} for UID "${params.uid}"`);
    if (params.responsibleParty !== undefined) ws.getRange(sheetRow, CONFIG.WL_COL_RESPONSIBLE).setValue(params.responsibleParty);
    if (params.dob !== undefined)              ws.getRange(sheetRow, CONFIG.WL_COL_AGE).setValue(calcAge_(params.dob)); // Age Col G
    if (params.phone !== undefined)            ws.getRange(sheetRow, CONFIG.WL_COL_PHONE).setValue(params.phone);
    if (params.potentialService !== undefined) ws.getRange(sheetRow, CONFIG.WL_COL_POTENTIAL_SERVICE).setValue(params.potentialService); 
    if (params.active !== undefined)           ws.getRange(sheetRow, CONFIG.WL_COL_ACTIVE).setValue(params.active); 
    if (params.spotFound !== undefined)        ws.getRange(sheetRow, CONFIG.WL_COL_SPOT_FOUND).setValue(params.spotFound);
    if (params.diagnosisNotes !== undefined)   ws.getRange(sheetRow, CONFIG.WL_COL_DIAGNOSIS_NOTES).setValue(params.diagnosisNotes);
  }
  Logger.log(`syncWaitingList_ (original): Processing done for row=${sheetRow}, UID="${params.uid}"`);
}

function syncTelephoneLog_(params) {
  Logger.log(`syncTelephoneLog_ (for other fields/initial): Called with params: ${JSON.stringify(params)}`);
  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  if (!tl) return;

  let sheetRow = findRowByUid_(tl, params.uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);

  if (params.isInitialCreation) {
    const dateToUse = (params.date instanceof Date && !isNaN(params.date)) ? params.date : new Date();
    const patientNameForLog = params.patientName || "";
    const emailForLog = params.email || "";
    const responsiblePartyForLog = params.responsibleParty || ""; 
    const phoneForLog = params.phone || "";

    if (sheetRow !== 0) { 
      Logger.log(`syncTelephoneLog_ (Initial Call): UID "${params.uid}" already exists at row ${sheetRow}. Updating relevant fields.`);
      const currentDateInLog = tl.getRange(sheetRow, CONFIG.TL_COL_DATE).getValue();
      if (!currentDateInLog || (currentDateInLog instanceof Date && dateToUse.getTime() > currentDateInLog.getTime())) {
        tl.getRange(sheetRow, CONFIG.TL_COL_DATE).setValue(dateToUse).setNumberFormat('MM/dd/yyyy');
        Logger.log(`syncTelephoneLog_ (Initial Call): Updated Date for existing UID "${params.uid}" to ${dateToUse.toISOString().slice(0,10)}`);
      }
      if (tl.getRange(sheetRow, CONFIG.TL_COL_PATIENT_NAME).getValue() !== patientNameForLog) {
        tl.getRange(sheetRow, CONFIG.TL_COL_PATIENT_NAME).setValue(patientNameForLog);
      }
      if (tl.getRange(sheetRow, CONFIG.TL_COL_EMAIL).getValue() !== emailForLog) {
        tl.getRange(sheetRow, CONFIG.TL_COL_EMAIL).setValue(emailForLog);
      }
      if (!tl.getRange(sheetRow, CONFIG.TL_COL_RESPONSIBLE).getValue() && responsiblePartyForLog) {
        tl.getRange(sheetRow, CONFIG.TL_COL_RESPONSIBLE).setValue(responsiblePartyForLog);
      }
      if (!tl.getRange(sheetRow, CONFIG.TL_COL_PHONE).getValue() && phoneForLog) {
        tl.getRange(sheetRow, CONFIG.TL_COL_PHONE).setValue(phoneForLog);
      }
      const waitlistCell = tl.getRange(sheetRow, CONFIG.TL_COL_WAITLIST_FLAG);
      if (waitlistCell.getValue() !== true) {
        waitlistCell.setValue(true); 
        Logger.log(`syncTelephoneLog_ (Initial Call): Set Waitlist flag to TRUE for existing UID "${params.uid}"`);
      }
    } else { 
      Logger.log(`syncTelephoneLog_ (Initial Call): UID "${params.uid}" not found. Appending new row.`);
      const numColsTl = Math.max(tl.getLastColumn(), CONFIG.TL_COL_ONSCHEDULE);
      const initialTlData = new Array(numColsTl).fill('');
      initialTlData[CONFIG.TL_COL_UID - 1]                 = params.uid;
      initialTlData[CONFIG.TL_COL_CALL_OUTCOME - 1]        = ""; 
      initialTlData[CONFIG.TL_COL_DATE - 1]                = dateToUse; 
      initialTlData[CONFIG.TL_COL_FORM_SUBMITTED - 1]      = false; // CHECKBOX: Form not submitted initially
      initialTlData[CONFIG.TL_COL_FORM_SENT - 1]           = false; // CHECKBOX: Form not sent initially
      initialTlData[CONFIG.TL_COL_RESPONSIBLE - 1]         = responsiblePartyForLog; 
      initialTlData[CONFIG.TL_COL_PATIENT_NAME - 1]        = patientNameForLog; 
      initialTlData[CONFIG.TL_COL_PHONE - 1]               = phoneForLog; 
      initialTlData[CONFIG.TL_COL_EMAIL - 1]               = emailForLog; 
      initialTlData[CONFIG.TL_COL_INFORMATION - 1]         = ""; 
      initialTlData[CONFIG.TL_COL_TELEVISIT_SCHED - 1]     = ""; 
      initialTlData[CONFIG.TL_COL_TELEVISIT_COMPLETE - 1]  = ""; 
      initialTlData[CONFIG.TL_COL_CONTACT_METHOD - 1]      = ""; 
      initialTlData[CONFIG.TL_COL_WAITLIST_FLAG - 1]       = params.addToWaitingList; // CHECKBOX: Not on waiting list initially
      initialTlData[CONFIG.TL_COL_ONSCHEDULE - 1]          = false;// CHECKBOX: Not on schedule initially
      initialTlData[CONFIG.TL_COL_NOT_INTERESTED - 1]     = false; // CHECKBOX: Not marked as not interested initially

      tl.appendRow(initialTlData);
      sheetRow = tl.getLastRow();
      Logger.log(`syncTelephoneLog_ (Initial Call): Appended new row ${sheetRow} for UID "${params.uid}"`);
      
      tl.getRange(sheetRow, CONFIG.TL_COL_DATE).setNumberFormat('MM/dd/yyyy');
      [CONFIG.TL_COL_FORM_SENT, CONFIG.TL_COL_FORM_SUBMITTED, CONFIG.TL_COL_WAITLIST_FLAG, CONFIG.TL_COL_ONSCHEDULE, CONFIG.TL_COL_NOT_INTERESTED].forEach(colIdx => {
          try { tl.getRange(sheetRow, colIdx).insertCheckboxes(); } catch (e) { Logger.log(`TL Checkbox Error (Initial Col ${colIdx}): ${e}`); }
      });
      if (!tl.isColumnHiddenByUser(CONFIG.TL_COL_UID)) { tl.hideColumns(CONFIG.TL_COL_UID); }

    }
  } else { 
    if (sheetRow === 0) {
      Logger.log(`syncTelephoneLog_ (Update Call): UID "${params.uid}" not found in Tel.Log. Aborting update.`); return;
    }
    Logger.log(`syncTelephoneLog_ (Update Call): Updating RP/Phone on row ${sheetRow} for UID "${params.uid}" if provided.`);
    if (params.responsibleParty !== undefined) {
        const currentVal = tl.getRange(sheetRow, CONFIG.TL_COL_RESPONSIBLE).getValue();
        if (currentVal !== params.responsibleParty) {
            tl.getRange(sheetRow, CONFIG.TL_COL_RESPONSIBLE).setValue(params.responsibleParty);
        }
    }
    if (params.addToWaitingList !== undefined) {
        const currentWaitlistFlag = tl.getRange(sheetRow, CONFIG.TL_COL_WAITLIST_FLAG).getValue();
        if (currentWaitlistFlag !== params.addToWaitingList) {
            tl.getRange(sheetRow, CONFIG.TL_COL_WAITLIST_FLAG).setValue(params.addToWaitingList);
        }
    }
    if (params.phone !== undefined) {
        const currentVal = tl.getRange(sheetRow, CONFIG.TL_COL_PHONE).getValue();
        if (currentVal !== params.phone) {
            tl.getRange(sheetRow, CONFIG.TL_COL_PHONE).setValue(params.phone);
        }
    }
  }
  Logger.log(`syncTelephoneLog_ (Original): Processing done for row=${sheetRow}, UID="${params.uid}"`);
}


/* ================================================================
 * 4.  CENTRALIZED SYNCHRONIZATION FOR SHARED PATIENT INFO (Name & Email)
 * ================================================================ */
function updateSharedPatientInfo(uid, newData, sourceSheetName) {
  Logger.log(`updateSharedPatientInfo: UID='${uid}', Data='${JSON.stringify(newData)}', Source='${sourceSheetName}'`);
  let updatedPatientName = newData.patientName;
  let updatedEmail = newData.email;

  if (sourceSheetName !== CONFIG.REGISTRY) {
    const regSheet = sheet_(CONFIG.REGISTRY);
    if (regSheet) {
      const regUidCol = regSheet.getRange(1, CONFIG.REG_COL_UID, regSheet.getLastRow(), 1).getValues().flat();
      const regRowIndex = regUidCol.indexOf(uid);
      if (regRowIndex !== -1) {
        const sheetRow = regRowIndex + 1;
        if (updatedPatientName !== undefined) {
          const currentRegName = regSheet.getRange(sheetRow, CONFIG.REG_COL_PATIENT_NAME).getValue();
          if (currentRegName !== updatedPatientName) {
            regSheet.getRange(sheetRow, CONFIG.REG_COL_PATIENT_NAME).setValue(updatedPatientName);
            Logger.log(`Updated Registry: Patient Name for UID ${uid} to "${updatedPatientName}"`);
          }
        }
        if (updatedEmail !== undefined) {
          const currentRegEmail = regSheet.getRange(sheetRow, CONFIG.REG_COL_EMAIL).getValue();
          if (currentRegEmail !== updatedEmail) {
            regSheet.getRange(sheetRow, CONFIG.REG_COL_EMAIL).setValue(updatedEmail);
            Logger.log(`Updated Registry: Email for UID ${uid} to "${updatedEmail}"`);
          }
        }
      } else { Logger.log(`updateSharedPatientInfo: UID ${uid} not found in Registry.`); }
    }
  }

  if (sourceSheetName !== CONFIG.WAITING_LIST) {
    const wlSheet = sheet_(CONFIG.WAITING_LIST);
    if (wlSheet) {
    const wlRow = findRowByUid_(wlSheet, uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);
      if (wlRow) {
        if (updatedPatientName !== undefined) {
          const currentWlName = wlSheet.getRange(wlRow, CONFIG.WL_COL_PATIENT_NAME).getValue();
          if (currentWlName !== updatedPatientName) {
            wlSheet.getRange(wlRow, CONFIG.WL_COL_PATIENT_NAME).setValue(updatedPatientName);
            Logger.log(`Updated Waiting List: Patient Name for UID ${uid} to "${updatedPatientName}"`);
          }
        }
        if (updatedEmail !== undefined) {
          const currentWlEmail = wlSheet.getRange(sheetRow, CONFIG.WL_COL_EMAIL).getValue();
          if (currentWlEmail !== updatedEmail) {
            wlSheet.getRange(sheetRow, CONFIG.WL_COL_EMAIL).setValue(updatedEmail);
            Logger.log(`Updated Waiting List: Email for UID ${uid} to "${updatedEmail}"`);
          }
        }
      } else { Logger.log(`updateSharedPatientInfo: UID ${uid} not found in Waiting List.`); }
    }
  }

  if (sourceSheetName !== CONFIG.TELEPHONE_LOG) {
    const tlSheet = sheet_(CONFIG.TELEPHONE_LOG);
    if (tlSheet) {
      const tlRow = findRowByUid_(tlSheet, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);
      if (tlRow) {
        if (updatedPatientName !== undefined) {
          const currentTlName = tlSheet.getRange(tlRow, CONFIG.TL_COL_PATIENT_NAME).getValue();
          if (currentTlName !== updatedPatientName) {
            tlSheet.getRange(tlRow, CONFIG.TL_COL_PATIENT_NAME).setValue(updatedPatientName);
            Logger.log(`Updated Telephone Log: Patient Name for UID ${uid} to "${updatedPatientName}"`);
          }
        }
        if (updatedEmail !== undefined) {
          const currentTlEmail = tlSheet.getRange(tlRow, CONFIG.TL_COL_EMAIL).getValue();
          if (currentTlEmail !== updatedEmail) {
            tlSheet.getRange(tlRow, CONFIG.TL_COL_EMAIL).setValue(updatedEmail);
            Logger.log(`Updated Telephone Log: Email for UID ${uid} to "${updatedEmail}"`);
          }
        }
      } else { Logger.log(`updateSharedPatientInfo: UID ${uid} not found in Telephone Log.`); }
    }
  }

  const allSheets = SS.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const currentSheet = allSheets[i];
    const currentSheetName = currentSheet.getName();
    if (SYSTEM_SHEET_NAMES.includes(currentSheetName) || currentSheetName === sourceSheetName) { continue; }
    try {
      const intakeUid = currentSheet.getRange(CONFIG.INTAKE_CELL_UID).getValue();
      if (intakeUid === uid) {
        Logger.log(`updateSharedPatientInfo: Found matching Intake Tab: "${currentSheetName}" for UID ${uid}`);
        if (updatedPatientName !== undefined) {
          const currentIntakeName = currentSheet.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).getValue();
          if (currentIntakeName !== updatedPatientName) {
            currentSheet.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).setValue(updatedPatientName);
            Logger.log(`Updated Intake Tab "${currentSheetName}": Patient Name to "${updatedPatientName}"`);
          }
        }
        if (updatedEmail !== undefined) {
          const currentIntakeEmail = currentSheet.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue();
          if (currentIntakeEmail !== updatedEmail) {
            currentSheet.getRange(CONFIG.INTAKE_CELL_EMAIL).setValue(updatedEmail);
            Logger.log(`Updated Intake Tab "${currentSheetName}": Email to "${updatedEmail}"`);
          }
        }
      }
    } catch (e) { /* Likely not an intake sheet or C1 doesn't exist. */ }
  }
  // update the Intake tab name if the patient name was changed  
  if (updatedPatientName !== undefined && updatedPatientName !== '') {
    renameIntakeTabsForUid_(uid, updatedPatientName);
  }
}


/* ================================================================
 * 5.  editHandler – Main Trigger
 * ================================================================ */
function editHandler(e) {
  if (!e || !e.range) { Logger.log('editHandler(): Event object or range is undefined. Exiting.'); return; }

  const range       = e.range;
  const sh          = range.getSheet();
  const sheetName   = sh.getName();
  const row         = range.getRow();
  const col         = range.getColumn();
  const editedValue = e.value;            // for check-boxes this is the string "TRUE" / "FALSE"

  Logger.log(`editHandler(): Sheet="${sheetName}", Cell=${range.getA1Notation()}, Row=${row}, Col=${col}, NewValue="${editedValue}"`);

  let uid;                      // UID tied to this row/tab, if resolvable
  let newData = {};             // {patientName?, email?} for shared-info sync
  let syncSharedInfo = false;   // flip to true when we’ve gathered newData

  /* ----------------------------------------------------------------
   * 1)  INTAKE TABS  – any sheet that is *not* in SYSTEM_SHEET_NAMES
   * ---------------------------------------------------------------- */
  if (!SYSTEM_SHEET_NAMES.includes(sheetName)) {

    const uidFromTab = sh.getRange(CONFIG.INTAKE_CELL_UID).getValue();
    if (uidFromTab) {
      uid = uidFromTab;
      const cellA1 = range.getA1Notation().split(':')[0];

      /* —— AUTO-MAIL: Therapy or Recreational paperwork pack —— */
      if (cellA1 === CONFIG.INTAKE_CELL_ACTIVE && editedValue === 'TRUE') {

        const email        = sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue();
        const responsible  = sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue();
        const isRecreational = sh.getRange(CONFIG.INTAKE_CELL_RECREATIONAL).getValue() === true;

        if (validateEmail(email)) {

          // Decide which pack to send
          const formKey  = isRecreational ? 'RECREATIONAL_START' : 'THERAPY_START';
          const noteText = isRecreational ? 'Recreational Forms Sent!' : 'Therapy Forms Sent!';

          const ok = sendForm_(formKey, { uid, email, responsible });

          if (ok) {
            // Write a clear “sent” note for staff
            sh.getRange(CONFIG.INTAKE_TAB_FORMS_SENT_NOTE).setValue(noteText);
            Logger.log(`${formKey} paperwork sent for UID ${uid}`);
          } 
        } else {
          // If email is invalid, un-check the checkbox to prevent confusion
          sh.getRange(CONFIG.INTAKE_CELL_ACTIVE).setValue(false);
          SpreadsheetApp.getUi().alert(
            'Invalid Email',
            `The email "${email}" is missing or invalid.`,
            SpreadsheetApp.getUi().ButtonSet.OK
          );
        }
      }

      /* —— NEW TWO-WAY CHECKBOX SYNC —— */
      if (cellA1 === CONFIG.INTAKE_CELL_ACTIVE) {
        syncActive_(uid, editedValue === 'TRUE', sheetName);
      } else if (cellA1 === CONFIG.INTAKE_CELL_SPOT_FOUND) {
        syncSpotFound_(uid, editedValue === 'TRUE', sheetName);
      } else if (cellA1 === CONFIG.INTAKE_CELL_NOT_INTERESTED) {
        syncNotInterested_(uid, editedValue === 'TRUE', sheetName);
      }

      /* —— AUTO-MAIL: Financial-Aid application —— */
      if (cellA1 === CONFIG.INTAKE_CELL_FINANCIAL_AID && editedValue === 'TRUE') {

        const email       = sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue();
        const responsible = sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue();

        if (validateEmail(email)) {
          const ok = sendForm_('FINANCIAL_AID_2025', { uid, email, responsible });
          if (ok) {
            // leave the checkbox checked – it now serves as the "sent" flag
            Logger.log(`Financial-Aid link e-mailed for UID ${uid}`);
            sh.getRange(CONFIG.INTAKE_TAB_FINANCIAL_AID_NOTE).setValue('Form Sent!'); // Mark the link as sent
          } else {
            // If sending failed, un-check the checkbox to prevent confusion
            sh.getRange(CONFIG.INTAKE_CELL_FINANCIAL_AID).setValue(false);
          }
        }else {
          // If email is invalid, un-check the checkbox to prevent confusion
          sh.getRange(CONFIG.INTAKE_CELL_FINANCIAL_AID).setValue(false);
          SpreadsheetApp.getUi().alert(
            'Invalid Email',
            `The email "${email}" is missing or invalid.`,
            SpreadsheetApp.getUi().ButtonSet.OK
          );
        }
      }

      /* —— AUTO-MAIL: Tele-health appointment link —— */
      if (cellA1 === CONFIG.INTAKE_CELL_TELEHEALTH_LINK && editedValue === 'TRUE') {

        const email       = sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue();
        const responsible = sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue();
        const dateStr = sh.getRange(CONFIG.INTAKE_CELL_TELEHEALTH_DATE).getDisplayValue();
        const timeStr = sh.getRange(CONFIG.INTAKE_CELL_TELEHEALTH_TIME).getDisplayValue();

        // Placeholder texts that mean the user hasn’t filled the field yet
        const DATE_PLACEHOLDER = 'Telehealth Visit Date, (e.g. Wednesday, June 11)';
        const TIME_PLACEHOLDER = 'Visit Time';

        // Guard-rail: make sure both fields are real values
        if (dateStr === DATE_PLACEHOLDER || timeStr === TIME_PLACEHOLDER || !dateStr || !timeStr) {
          SpreadsheetApp.getUi().alert(
            'Please enter BOTH the Tele-health visit date *and* time before sending the link.'
          );
          // Un-check the “Send Link” box so the user can try again later
          sh.getRange(CONFIG.INTAKE_CELL_TELEHEALTH_LINK).setValue(false);
          return;                  // Abort the send
        }

        const apptDate = `${dateStr}, at ${timeStr}`;

        if (validateEmail(email)) {
          const ok = sendForm_('TELEHEALTH_APPT',
                              { uid, email, responsible, apptDate });
          if (ok) {
            Logger.log(`Tele-health link sent for UID ${uid}`);
            sh.getRange(CONFIG.INTAKE_TAB_TELEHEALTH_LINK_NOTE).setValue('Link Sent!'); // Mark the link as sent
          } else {
            // If sending failed, un-check the checkbox to prevent confusion
            sh.getRange(CONFIG.INTAKE_CELL_TELEHEALTH_LINK).setValue(false);
          }
        } else {
          // If email is invalid, un-check the checkbox to prevent confusion
          sh.getRange(CONFIG.INTAKE_CELL_TELEHEALTH_LINK).setValue(false);
          SpreadsheetApp.getUi().alert(
            'Invalid Email',
            `The email "${email}" is missing or invalid.`,
            SpreadsheetApp.getUi().ButtonSet.OK
          );
        }
      }



      /* —— INTAKE CALL COMPLETED - ADD TO WAITING LIST —— */
      if (cellA1 === CONFIG.INTAKE_CELL_CALL_COMPLETED && editedValue === 'TRUE') {

        // FIRST, make sure we don't duplicate Waiting List entry.
        const wl = sheet_(CONFIG.WAITING_LIST);
        const uidList = wl.getRange(2, CONFIG.WL_COL_UID, wl.getLastRow()-1, 1).getValues().flat();
        
        if (uidList.indexOf(uid) === -1) { // if patient isn't yet in WL
        
          // SECOND, collect patient details from Intake tab:
          const payload = {
            uid,
            patientName      : sh.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).getValue(),
            responsibleParty : sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue(),
            phone            : sh.getRange(CONFIG.INTAKE_CELL_PHONE).getValue(),
            email            : sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue(),
            dob              : sh.getRange(CONFIG.INTAKE_CELL_DOB).getValue(),
            potentialService : sh.getRange(CONFIG.INTAKE_CELL_POTENTIAL_SERVICE).getValue(),
            intakeDate       : new Date(),
            isInitialCreation: true
          };

          // THIRD, now actually write the entry to WL:
          syncWaitingList_(payload);

          // FOURTH, mark TL checkbox ("On waiting list"):
          syncTelephoneLog_({ uid, addToWaitingList:true, isInitialCreation:false });

          // FIFTH, clear E2 on the Intake tab:
          sh.getRange(CONFIG.INTAKE_TAB_WL_NOTE).clearContent();

          // SIXTH, color the tab yellow to indicate it’s now in the Waiting List:
          setIntakeTabColor_(sh);
        }
      }

      /* —— EXISTING NAME / EMAIL SYNC —— */
      if (cellA1 === CONFIG.INTAKE_CELL_PATIENT_NAME) {
        newData.patientName = editedValue; syncSharedInfo = true;
      } else if (cellA1 === CONFIG.INTAKE_CELL_EMAIL) {
        newData.email = editedValue;
        if (editedValue && !validateEmail(editedValue)) {
          SpreadsheetApp.getUi().alert("Invalid Email",
              `The email "${editedValue}" is not a valid format. It will not be synced.`,
              SpreadsheetApp.getUi().ButtonSet.OK);
          syncSharedInfo = false; newData.email = undefined;
        } else { syncSharedInfo = true; }
      }

      /* —— ORIGINAL INTAKE → WL/TL SYNC FOR OTHER FIELDS —— */
      if (INTAKE_SHEET_RELEVANT_CELLS_FOR_ORIGINAL_SYNC.hasOwnProperty(cellA1)) {
        const payload = {
          uid,
          patientName:      sh.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).getValue(),
          responsibleParty: sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue(),
          phone:            sh.getRange(CONFIG.INTAKE_CELL_PHONE).getValue(),
          email:            sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue(),
          dob:              null,
          potentialService: sh.getRange(CONFIG.INTAKE_CELL_POTENTIAL_SERVICE).getValue(),
          active:           sh.getRange(CONFIG.INTAKE_CELL_ACTIVE).getValue(),
          spotFound:        sh.getRange(CONFIG.INTAKE_CELL_SPOT_FOUND).getValue(),
          diagnosisNotes:   sh.getRange(CONFIG.INTAKE_CELL_DIAGNOSIS_NOTES).getValue(),
          isInitialCreation: false
        };
        const dobVal = sh.getRange(CONFIG.INTAKE_CELL_DOB).getValue();
        if (dobVal instanceof Date && !isNaN(dobVal.getTime())) payload.dob = dobVal;

        syncWaitingList_(payload);
        syncTelephoneLog_({
          uid,
          patientName: payload.patientName,
          responsibleParty: payload.responsibleParty,
          phone: payload.phone,
          email: payload.email,
          isInitialCreation: false
        });
      }

    } else {
      Logger.log(`editHandler(): Edit on "${sheetName}" but ${CONFIG.INTAKE_CELL_UID} is blank – skipping.`);
    }
  }

  /* ---------------------------------------------------------------
   * 2)  WAITING LIST
   * --------------------------------------------------------------- */
  else if (sheetName === CONFIG.WAITING_LIST) {
    if (row > 1) {                                      // skip header row
      uid = sh.getRange(row, CONFIG.WL_COL_UID).getValue();

      /* —— NEW CHECKBOX SYNC —— */
      if (uid) {
        if (col === CONFIG.WL_COL_ACTIVE) {
          syncActive_(uid, editedValue === 'TRUE', CONFIG.WAITING_LIST);
        } else if (col === CONFIG.WL_COL_SPOT_FOUND) {
          syncSpotFound_(uid, editedValue === 'TRUE', CONFIG.WAITING_LIST);
        } else if (col === CONFIG.WL_NOT_INTERESTED) {
          syncNotInterested_(uid, editedValue === 'TRUE', sheetName);
        }
      }

      /* —— NAME / EMAIL SYNC —— */
      if (uid) {
        if (col === CONFIG.WL_COL_PATIENT_NAME) {
          newData.patientName = editedValue; syncSharedInfo = true;
        } else if (col === CONFIG.WL_COL_EMAIL) {
          newData.email = editedValue;
          if (editedValue && !validateEmail(editedValue)) {
            SpreadsheetApp.getUi().alert("Invalid Email",
                `The email "${editedValue}" is not valid. Not synced.`,
                SpreadsheetApp.getUi().ButtonSet.OK);
            syncSharedInfo = false; newData.email = undefined;
          } else { syncSharedInfo = true; }
        }
      }
    }
  }

  /* ---------------------------------------------------------------
   * 3)  TELEPHONE LOG
   * --------------------------------------------------------------- */
  else if (sheetName === CONFIG.TELEPHONE_LOG) {
    if (row > 1) {
      uid = sh.getRange(row, CONFIG.TL_COL_UID).getValue();

      /* —— NEW CHECKBOX SYNC (On-Schedule) —— */
      if (uid) {
        if (col === CONFIG.TL_COL_ONSCHEDULE) {
          syncActive_(uid, editedValue === 'TRUE', CONFIG.TELEPHONE_LOG);
        } else if (col === CONFIG.TL_COL_NOT_INTERESTED) {
          syncNotInterested_(uid, editedValue === 'TRUE', CONFIG.TELEPHONE_LOG);
        }
      }

      /* —— NAME / EMAIL SYNC —— */
      if (uid) {
        if (col === CONFIG.TL_COL_PATIENT_NAME) {
          newData.patientName = editedValue; syncSharedInfo = true;
        } else if (col === CONFIG.TL_COL_EMAIL) {
          newData.email = editedValue;
          if (editedValue && !validateEmail(editedValue)) {
            SpreadsheetApp.getUi().alert("Invalid Email",
                `The email "${editedValue}" is not valid. Not synced.`,
                SpreadsheetApp.getUi().ButtonSet.OK);
            syncSharedInfo = false; newData.email = undefined;
          } else { syncSharedInfo = true; }
        }

        if (col === CONFIG.TL_COL_EMAIL) {
          /* Auto-send Google Form when an e-mail appears and Form not yet sent */
          const emailCellVal = sh.getRange(row, CONFIG.TL_COL_EMAIL).getValue();
          const sentCell     = sh.getRange(row, CONFIG.TL_COL_FORM_SENT);  // checkbox
          const responsibleCell = sh.getRange(row, CONFIG.TL_COL_RESPONSIBLE).getValue();
          
          let responsible = responsibleCell;
          if (isPlaceholderResponsibleParty(responsibleCell)) {
            responsible = 'thank you for contacting Special Strides';
          }

          
          if (emailCellVal) {

            /* ——— 2️⃣  Valid address? ——— */
            if (validateEmail(emailCellVal)) {

              /* ——— 3️⃣  Not already sent? ——— */
              if (sentCell.getValue() !== true) {
                if (sendForm_('INTAKE', { uid, email: emailCellVal, responsible })) {
                  sentCell.setValue(true);                        // tick the checkbox
                  Logger.log(`editHandler (Tel.Log): Form e-mailed and checkbox ticked for UID ${uid}.`);
                }
              }

            } else {
              /* ⚠️  Invalid address → tell the user and do nothing else */
              SpreadsheetApp.getUi().alert(
                'Invalid E-mail',
                `"${emailCellVal}" is not a valid address. Please correct it before sending.`,
                SpreadsheetApp.getUi().ButtonSet.OK
              );
            }
          }
        }  
      }
    }
  }

  /* ---------------------------------------------------------------
   * 4)  PATIENT REGISTRY
   * --------------------------------------------------------------- */
  else if (sheetName === CONFIG.REGISTRY) {
    if (row > 1) {
      uid = sh.getRange(row, CONFIG.REG_COL_UID).getValue();
      if (uid) {
        if (col === CONFIG.REG_COL_PATIENT_NAME) {
          newData.patientName = editedValue; syncSharedInfo = true;
        } else if (col === CONFIG.REG_COL_EMAIL) {
          newData.email = editedValue;
          if (editedValue && !validateEmail(editedValue)) {
            SpreadsheetApp.getUi().alert("Invalid Email",
                `The email "${editedValue}" is not valid. Not synced.`,
                SpreadsheetApp.getUi().ButtonSet.OK);
            syncSharedInfo = false; newData.email = undefined;
          } else { syncSharedInfo = true; }
        }
      }
    }
  }

  /* ---------------------------------------------------------------
   * 5)  FINAL – propagate shared NAME / E-MAIL updates
   * --------------------------------------------------------------- */
  if (syncSharedInfo && uid &&
      (newData.patientName !== undefined || newData.email !== undefined)) {
    updateSharedPatientInfo(uid, newData, sheetName);
  } else if (syncSharedInfo && !uid) {
    Logger.log(`editHandler: Wanted to propagate shared info from "${sheetName}" but UID was not resolved.`);
  }
}

/* ================================================================
 * 5b  Universal on-Submit handler
 * ================================================================ */
function onFormSubmitTrigger(e) {
  const sheetName = e.range.getSheet().getName();     // e.g. “Form Responses”
  const formKey = Object.keys(FORMS).find(k => FORMS[k].responseSheet === sheetName);
  if (!formKey) { Logger.log(`onFormSubmitTrigger: Sheet ${sheetName} not registered`); return; }

  const cfg = FORMS[formKey];

  /* --- locate UID in the just-added row ---------------------- */
  const header  = e.range.getSheet().getRange(1,1,1,e.range.getLastColumn()).getValues()[0];
  const uidCol  = header.indexOf('UID') + 1;
  const uid     = e.range.getSheet().getRange(e.range.getRow(), uidCol).getValue();
  if (!uid) { Logger.log('onFormSubmitTrigger: UID missing'); return; }

  /* --- flip the Submitted flag in History -------------------- */
  const hist = sheet_(CONFIG.HISTORY);
  const data = hist.getDataRange().getValues();
  for (let i=1;i<data.length;i++){
    if (data[i][CONFIG.HISTORY_COL_UID-1]  === uid &&
        data[i][CONFIG.HISTORY_COL_FORM-1] === formKey) {
      hist.getRange(i+1, CONFIG.HISTORY_COL_SUBMITTED).setValue(true);
      break;
    }
  }
  Logger.log(`onFormSubmitTrigger: marked ${formKey} submitted for UID ${uid}`);
}


/* ================================================================
 * 6.  Form-send + history log
 * ================================================================ */
function sendForm_(formKey, { uid, email, patient = '', responsible = '', apptDate = '' }) {
  const cfg = FORMS[formKey];
  if (!cfg) { Logger.log(`sendForm_: Unknown formKey "${formKey}"`); return false; }

  /* —— NEW CONFIRMATION DIALOG —— */
  const displayName = cfg.displayName || formKey;
  const confirm = confirmSend_(displayName, responsible || patient || email);
  if (!confirm.ok) { Logger.log(`sendForm_: User cancelled ${formKey}`); return false; }
  // If the user typed a new value, overwrite 'responsible'
  responsible = confirm.responsible;

  /* --- prevent duplicates ------------------------------------ */
  const history = sheet_(CONFIG.HISTORY);
  const rows = history.getDataRange().getValues();
  if (rows.some(r => r[CONFIG.HISTORY_COL_UID-1]  === uid &&
                     r[CONFIG.HISTORY_COL_FORM-1] === formKey &&
                     r[CONFIG.HISTORY_COL_SENT-1] === true)) {
    Logger.log(`sendForm_: ${formKey} already sent for UID ${uid}`); return false;
  }

  /* --- build the personalised link & e-mail ------------------ */
  const link = `${cfg.formUrl}&${cfg.uidEntry}=${encodeURIComponent(uid)}`;
  const html = cfg.mail.body({link, patient, responsible, ...arguments[1]});
  MailApp.sendEmail({to: email, subject: cfg.mail.subject, htmlBody: html});

  /* --- write / update history -------------------------------- */
  let found = false;
  for (let i=1;i<rows.length;i++){
    if (rows[i][CONFIG.HISTORY_COL_UID-1]  === uid &&
        rows[i][CONFIG.HISTORY_COL_FORM-1] === formKey) {
      history.getRange(i+1, CONFIG.HISTORY_COL_DATE).setValue(new Date());
      history.getRange(i+1, CONFIG.HISTORY_COL_SENT).setValue(true);
      found = true; break;
    }
  }
  if (!found) {
    history.appendRow([uid, formKey, new Date(), email, true, false]);
  }
  Logger.log(`sendForm_: ${formKey} sent for UID ${uid}`);
  SpreadsheetApp.getActiveSpreadsheet().toast(`${displayName} sent to ${email}`, 'Email sent', 5);
  
  return true;
}

/**
 * Confirmation dialog for automatic e-mails.
 * • When we already have a 'responsible' name → simple Yes/No.
 * • When it's blank  → prompt the user to enter one.
 * Returns { ok:Boolean, responsible:String }
 */
function confirmSend_(displayName, responsible) {
  const ui = SpreadsheetApp.getUi();
  responsible = (responsible || '').trim();

  /* ——— CASE 1: name already present ——— */
  if (responsible) {
    const yesNo = ui.alert(
      `Send ${displayName}?`,
      `Would you like to send “${displayName}” to “${responsible}”?`,
      ui.ButtonSet.YES_NO
    );
    return (yesNo === ui.Button.YES)
      ? { ok: true, responsible }
      : { ok: false };
  }

  /* ——— CASE 2: name missing ——— */
  const prompt = ui.prompt(
    `Send ${displayName}`,
    'Enter the recipient’s name:',
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return { ok: false };

  const newName = prompt.getResponseText().trim();
  return newName
    ? { ok: true, responsible: newName }
    : { ok: false };           // empty → abort
}

/* ================================================================
 * 7.  HELPERS
 * ================================================================ */
function validateEmail(email) {
  if (!email) return false;
  const re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(String(email).toLowerCase());
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
  else                          { sheet.setTabColor(null); } // default color     
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

// sync helpers:
// Helper – returns the absolute sheet row for a UID, or 0 if not found
function findRowByUid_(sh, uid, uidCol, headerRows) {
  const firstData = headerRows + 1;                          // first data row
  const uidArr = sh.getRange(firstData, uidCol,
                             sh.getLastRow() - headerRows, 1)
                   .getValues().flat();
  const idx = uidArr.indexOf(uid);                           // 0-based inside slice
  return (idx === -1) ? 0 : firstData + idx;                 // absolute row
}

/**
 * Keep Active / On-Schedule (WL B  ⇆  TL N  ⇆  Intake D3)
 */
function syncActive_(uid, newVal, source) {
  // Waiting List
  if (source !== CONFIG.WAITING_LIST) {
    const wl = sheet_(CONFIG.WAITING_LIST);
    const wlRow = findRowByUid_(wl, uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);
    if (wlRow) wl.getRange(wlRow, CONFIG.WL_COL_ACTIVE).setValue(newVal);
  }

  // Telephone Log
  if (source !== CONFIG.TELEPHONE_LOG) {
    const tl = sheet_(CONFIG.TELEPHONE_LOG);
    const tlRow = findRowByUid_(tl, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);
    if (tlRow) tl.getRange(tlRow, CONFIG.TL_COL_ONSCHEDULE).setValue(newVal);
  }

  // Intake tabs (may be >1 if copies were made)
  findIntakeSheetsByUid_(uid).forEach(sh => {
    if (source !== sh.getName())
      sh.getRange(CONFIG.INTAKE_CELL_ACTIVE).setValue(newVal);
    setIntakeTabColor_(sh);
  });
}

/**
 * Keep Spot-Found (WL C  ⇆  Intake D2)
 */
function syncSpotFound_(uid, newVal, source) {
  // Waiting List
  if (source !== CONFIG.WAITING_LIST) {
    const wl = sheet_(CONFIG.WAITING_LIST);
    const wlRow = findRowByUid_(wl, uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);
    if (wlRow) wl.getRange(wlRow, CONFIG.WL_COL_SPOT_FOUND).setValue(newVal);
  }

  // Intake tabs
  findIntakeSheetsByUid_(uid).forEach(sh => {
    if (source !== sh.getName())
      sh.getRange(CONFIG.INTAKE_CELL_SPOT_FOUND).setValue(newVal);
    setIntakeTabColor_(sh);
  });
}


/**
 * Keep Not-Interested
 *   (WL Col O  ⇆  TL Col P  ⇆  Intake D5)
 *   When TRUE it also clears the “active/on-schedule” & “spot-found” flags everywhere.
 */
function syncNotInterested_(uid, newVal, source) {
  /* ---- WAITING LIST ---- */
  const wl = sheet_(CONFIG.WAITING_LIST);
  const wlRow = findRowByUid_(wl, uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);
  if (wlRow) {
    if (source !== CONFIG.WAITING_LIST)
      wl.getRange(wlRow, CONFIG.WL_NOT_INTERESTED).setValue(newVal);
    if (newVal) {
      wl.getRange(wlRow, CONFIG.WL_COL_ACTIVE    ).setValue(false);
      wl.getRange(wlRow, CONFIG.WL_COL_SPOT_FOUND).setValue(false);
    }
  }

  /* ---- TELEPHONE LOG ---- */
  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  const tlRow = findRowByUid_(tl, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);
  if (tlRow) {
    if (source !== CONFIG.TELEPHONE_LOG)
      tl.getRange(tlRow, CONFIG.TL_COL_NOT_INTERESTED).setValue(newVal);
    if (newVal) tl.getRange(tlRow, CONFIG.TL_COL_ONSCHEDULE).setValue(false);
  }

  /* ---- INTAKE TABS ---- */
  findIntakeSheetsByUid_(uid).forEach(sh => {
    if (source !== sh.getName())       sh.getRange(5, 4).setValue(newVal);  // D5
    if (newVal) {
      sh.getRange(4, 4).setValue(false);  // D4
      sh.getRange(3, 4).setValue(false);  // D3
    }
    setIntakeTabColor_(sh);
  });
}

/**
 * Rename every Intake sheet that carries the given UID so the tab title
 * matches the (possibly updated) patient name.  
 * If the desired title is already taken we append “ (2)”, “ (3)”, …  
 * This guarantees unique tab names while still making them readable.
 */
function renameIntakeTabsForUid_(uid, patientName) {
  if (!patientName) return;                       // nothing to do
  const targetSheets = findIntakeSheetsByUid_(uid);   // already exists :contentReference[oaicite:0]{index=0}

  targetSheets.forEach(sh => {
    if (sh.getName() === patientName) return;     // already correct

    let newName   = patientName;
    let counter   = 2;
    /* keep trying until the candidate name is either unused
       or it refers to this very sheet (happens during rename-loops) */
    while (SS.getSheetByName(newName) &&
           SS.getSheetByName(newName).getSheetId() !== sh.getSheetId()) {
      newName = `${patientName} (${counter++})`;
    }
    sh.setName(newName);                          // 🔀 the actual rename
  });
}


function isPlaceholderEmail(email) {
  // Treat blank, default instructional note, or non‑RFC e‑mail as invalid
  if (!email) return true;
  if (email === CONFIG.DEFAULT_TL_EMAIL_INFO_NOTE) return true;
  return !validateEmail(email);
}

function isPlaceholderResponsibleParty(responsible) {
  // Treat blank, default instructional note, or non‑RFC e‑mail as invalid
  if (!responsible) return true;
  if (responsible === CONFIG.DEFAULT_TL_RESPONSIBLE_PARTY_NOTE) return true;
  return false;
}

function isValidPatientAndEmail(patient, email) {
  // Both patient & e‑mail must be present and e‑mail must be valid
  return !!patient && validateEmail(email);
}

/** Exposed to HtmlService – returns the same data as the private helper. */
function listOpenTLRows() {
  return getRecentTLWithoutIntake_();
}

/**
 * Put the freshly-copied Intake tab right after the last system sheet.
 * Works for createPatientIntakeTab and createIntakeFromDialog.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh  – the new Intake sheet
 */
function placeAfterSystemSheets_(sh) {
  const ss = SpreadsheetApp.getActive();

  // Highest (right-most) system-sheet index, or 0 if none
  const lastSysIdx = ss.getSheets()
                       .filter(s => SYSTEM_SHEET_NAMES.includes(s.getName()))
                       .reduce((max, s) => Math.max(max, s.getIndex()), 0);

  // `moveActiveSheet` requires the sheet to be active first
  ss.setActiveSheet(sh);            // make the new tab active
  ss.moveActiveSheet(lastSysIdx + 1);   // 1-based index → immediately after the system tabs
}

