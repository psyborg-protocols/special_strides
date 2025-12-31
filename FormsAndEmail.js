/* ================================================================
 * FORMS & EMAILS
 * ================================================================ */

function getFormResponseForUid_(uid) {
  const fs = sheet_(CONFIG.FORM_RESPONSES);
  if (!fs || fs.getLastRow() <= 1) return null;

  const UID_COL = 3; 
  const TOTAL_COLUMNS = 25;

  const headers = fs.getRange(1, 1, 1, TOTAL_COLUMNS).getValues()[0];
  const uidCol  = fs.getRange(2, UID_COL, fs.getLastRow() - 1, 1).getValues().flat();
  const idx = uidCol.lastIndexOf(uid);

  if (idx === -1) return null;

  const answers = fs.getRange(idx + 2, 1, 1, TOTAL_COLUMNS).getValues()[0];
  return { questions: headers, answers };
}

function pasteFormAnswersToIntakeStructured_(sheet, qa) {
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

  Object.entries(MAP).forEach(([cell, value]) => {
    if (value) sheet.getRange(cell).setValue(value);
  });
  
  sheet.getRange('B40').setValue('Google-Form answers imported automatically').setFontStyle('italic').setFontSize(9).setBackground('#f5f5ff');
}

function sendForm_(formKey, { uid, email, patient = '', responsible = '', apptDate = '' }) {
  // Use the new helper to get merged config (Code logic + Sheet URL)
  const cfg = getFormConfig_(formKey); 
  
  if (!cfg) { 
    SpreadsheetApp.getUi().alert(`Configuration Error: Could not load form details for "${formKey}". Check the System_Form_Links sheet.`);
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

  const link = `${cfg.formUrl}&${cfg.uidEntry}=${encodeURIComponent(uid)}`;
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