/* ================================================================
 * INTAKE TAB CREATION & MANAGEMENT
 * ================================================================ */

function createPatientIntakeTab() {
  Logger.log('createPatientIntakeTab(): start');
  const ui = SpreadsheetApp.getUi();

  let r = ui.prompt('Patient name:', 'Please enter the patient’s full name (LastName, FirstName).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || !r.getResponseText()) return;
  const patient = r.getResponseText().trim();

  let email = '';
  let emailPromptResponse;
  while (true) {
    emailPromptResponse = ui.prompt('Email address:', 'Please enter a contact email address. (Leave blank and click OK if no email)', ui.ButtonSet.OK_CANCEL);
    if (emailPromptResponse.getSelectedButton() !== ui.Button.OK) return;
    email = emailPromptResponse.getResponseText().trim();
    if (email === "") break;
    if (validateEmail(email)) break;
    else ui.alert('Invalid Email Format', `The email "${email}" is not a valid format.`, ui.ButtonSet.OK);
  }

  const uid = getOrCreateUID_(patient, undefined, email);
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

  sh.getRange(CONFIG.INTAKE_CELL_UID).setValue(uid);
  sh.getRange(CONFIG.INTAKE_CELL_DATE).setValue(new Date()).setNumberFormat('MM/dd/yyyy');
  sh.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).setValue(patient);
  sh.getRange(CONFIG.INTAKE_CELL_EMAIL).setValue(email);
  sh.getRange(CONFIG.INTAKE_CELL_UID).setFontWeight('normal').setFontSize(12);

  syncTelephoneLog_({
    uid: uid, patientName: patient, email: email, date: new Date(), addToWaitingList:false, isInitialCreation: true
  });

  // Import Form answers if available
  const existingForm = getFormResponseForUid_(uid);
  if (existingForm) {
    const uiResp = ui.alert('Form Already Filled', 'Import answers?', ui.ButtonSet.YES_NO);
    if (uiResp === ui.Button.YES) pasteFormAnswersToIntakeStructured_(sh, existingForm);
  }

  ui.alert('Intake Tab Created', `New intake tab "${newSheetName}" created.`, ui.ButtonSet.OK);
}

function openIntakeCreator() {
  const html = HtmlService.createHtmlOutputFromFile('IntakeCreator').setWidth(420).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create / Attach Intake');
}

function createIntakeFromDialog(uid) {
  if (uid === 'NEW') {
    createPatientIntakeTab(); // Fallback to classic
    return;
  }
  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  const r = findRowByUid_(tl, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);

  if (!r) { SpreadsheetApp.getUi().alert('Selected call not found'); return; }
  const patient  = tl.getRange(r, CONFIG.TL_COL_PATIENT_NAME).getValue();
  const email    = tl.getRange(r, CONFIG.TL_COL_EMAIL).getValue();
  const responsible = tl.getRange(r, CONFIG.TL_COL_RESPONSIBLE).getValue();
  const phone    = tl.getRange(r, CONFIG.TL_COL_PHONE).getValue();

  const finalUid = uid || getOrCreateUID_(patient, responsible, email);
  ensureRegistryEntry_(finalUid, patient, email);
  markHasIntake_(finalUid, true);

  const template = sheet_(CONFIG.TEMPLATE);
  const shName   = patient || `Intake ${finalUid.slice(0, 6)}`;
  const sh       = template.copyTo(SS).setName(shName);
  placeAfterSystemSheets_(sh);
  sh.showSheet();
  setIntakeTabColor_(sh);
  SS.setActiveSheet(sh);

  sh.getRange(CONFIG.INTAKE_CELL_UID).setValue(finalUid);
  sh.getRange(CONFIG.INTAKE_CELL_DATE).setValue(new Date()).setNumberFormat('MM/dd/yyyy');
  if (patient) sh.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).setValue(patient);
  if (email)   sh.getRange(CONFIG.INTAKE_CELL_EMAIL).setValue(email);
  if (responsible) sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).setValue(responsible);

  const qa = getFormResponseForUid_(finalUid);
  if (qa) pasteFormAnswersToIntakeStructured_(sh, qa);

  syncTelephoneLog_({
    uid: finalUid, patientName: patient, email: email, responsibleParty: responsible, phone: phone, isInitialCreation: true
  });

  SpreadsheetApp.getUi().alert('Intake tab created & linked!');
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