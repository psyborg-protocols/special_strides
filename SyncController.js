/* ================================================================
 * SYNC LOGIC (Waiting List <-> Telephone Log <-> Intake)
 * ================================================================ */

function syncWaitingList_(params) {
  Logger.log(`syncWaitingList_: ${JSON.stringify(params)}`);
  const ws = sheet_(CONFIG.WAITING_LIST);
  if (!ws) return;

  let sheetRow = findRowByUid_(ws, params.uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);

  if (params.isInitialCreation) {
    const intakeDateToUse = (params.intakeDate instanceof Date) ? params.intakeDate.toLocaleDateString('en-US') : new Date().toLocaleDateString('en-US');
    const initialData = [ 
      params.uid, false, false, params.responsibleParty || '', params.patientName || '', '',
      calcAge_(params.dob), params.phone || '', params.email || '', '',
      intakeDateToUse, params.potentialService || '', false, "", false, true
    ];
    ws.appendRow(initialData);
    sheetRow = ws.getLastRow();

    ws.getRange(sheetRow, 10).setNumberFormat('MM/dd/yyyy');
    [CONFIG.WL_COL_ACTIVE, CONFIG.WL_COL_SPOT_FOUND, CONFIG.WL_SOCIAL_STRIDES, CONFIG.WL_NOT_INTERESTED, CONFIG.WL_INTAKE_COMPLETED].forEach(colIdx => {
        try { ws.getRange(sheetRow, colIdx).insertCheckboxes(); } catch (e) {}
    });
    if (!ws.isColumnHiddenByUser(CONFIG.WL_COL_UID)) { ws.hideColumns(CONFIG.WL_COL_UID); }

  } else { 
    if (sheetRow === 0) return;
    if (params.responsibleParty !== undefined) ws.getRange(sheetRow, CONFIG.WL_COL_RESPONSIBLE).setValue(params.responsibleParty);
    if (params.dob !== undefined)              ws.getRange(sheetRow, CONFIG.WL_COL_AGE).setValue(calcAge_(params.dob));
    if (params.phone !== undefined)            ws.getRange(sheetRow, CONFIG.WL_COL_PHONE).setValue(params.phone);
    if (params.potentialService !== undefined) ws.getRange(sheetRow, CONFIG.WL_COL_POTENTIAL_SERVICE).setValue(params.potentialService); 
    if (params.active !== undefined)           ws.getRange(sheetRow, CONFIG.WL_COL_ACTIVE).setValue(params.active); 
    if (params.spotFound !== undefined)        ws.getRange(sheetRow, CONFIG.WL_COL_SPOT_FOUND).setValue(params.spotFound);
    if (params.diagnosisNotes !== undefined)   ws.getRange(sheetRow, CONFIG.WL_COL_DIAGNOSIS_NOTES).setValue(params.diagnosisNotes);
  }
}

function syncTelephoneLog_(params) {
  Logger.log(`syncTelephoneLog_: ${JSON.stringify(params)}`);
  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  if (!tl) return;

  let sheetRow = findRowByUid_(tl, params.uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);

  if (params.isInitialCreation) {
    const dateToUse = (params.date instanceof Date) ? params.date : new Date();
    const patientNameForLog = params.patientName || "";
    const emailForLog = params.email || "";
    const responsiblePartyForLog = params.responsibleParty || ""; 
    const phoneForLog = params.phone || "";

    if (sheetRow !== 0) { 
      // Update existing
      const currentDateInLog = tl.getRange(sheetRow, CONFIG.TL_COL_DATE).getValue();
      if (!currentDateInLog || (currentDateInLog instanceof Date && dateToUse.getTime() > currentDateInLog.getTime())) {
        tl.getRange(sheetRow, CONFIG.TL_COL_DATE).setValue(dateToUse).setNumberFormat('MM/dd/yyyy');
      }
      if (tl.getRange(sheetRow, CONFIG.TL_COL_PATIENT_NAME).getValue() !== patientNameForLog) tl.getRange(sheetRow, CONFIG.TL_COL_PATIENT_NAME).setValue(patientNameForLog);
      if (tl.getRange(sheetRow, CONFIG.TL_COL_EMAIL).getValue() !== emailForLog) tl.getRange(sheetRow, CONFIG.TL_COL_EMAIL).setValue(emailForLog);
      if (!tl.getRange(sheetRow, CONFIG.TL_COL_RESPONSIBLE).getValue() && responsiblePartyForLog) tl.getRange(sheetRow, CONFIG.TL_COL_RESPONSIBLE).setValue(responsiblePartyForLog);
      if (!tl.getRange(sheetRow, CONFIG.TL_COL_PHONE).getValue() && phoneForLog) tl.getRange(sheetRow, CONFIG.TL_COL_PHONE).setValue(phoneForLog);
      
      tl.getRange(sheetRow, CONFIG.TL_COL_WAITLIST_FLAG).setValue(true); 
    } else { 
      // Append new
      const numColsTl = Math.max(tl.getLastColumn(), CONFIG.TL_COL_ONSCHEDULE);
      const initialTlData = new Array(numColsTl).fill('');
      initialTlData[CONFIG.TL_COL_UID - 1]                 = params.uid;
      initialTlData[CONFIG.TL_COL_CALL_OUTCOME - 1]        = ""; 
      initialTlData[CONFIG.TL_COL_DATE - 1]                = dateToUse; 
      initialTlData[CONFIG.TL_COL_FORM_SUBMITTED - 1]      = false;
      initialTlData[CONFIG.TL_COL_FORM_SENT - 1]           = false;
      initialTlData[CONFIG.TL_COL_RESPONSIBLE - 1]         = responsiblePartyForLog; 
      initialTlData[CONFIG.TL_COL_PATIENT_NAME - 1]        = patientNameForLog; 
      initialTlData[CONFIG.TL_COL_PHONE - 1]               = phoneForLog; 
      initialTlData[CONFIG.TL_COL_EMAIL - 1]               = emailForLog; 
      initialTlData[CONFIG.TL_COL_WAITLIST_FLAG - 1]       = params.addToWaitingList;
      initialTlData[CONFIG.TL_COL_ONSCHEDULE - 1]          = false;
      initialTlData[CONFIG.TL_COL_NOT_INTERESTED - 1]     = false;

      tl.appendRow(initialTlData);
      sheetRow = tl.getLastRow();
      
      tl.getRange(sheetRow, CONFIG.TL_COL_DATE).setNumberFormat('MM/dd/yyyy');
      [CONFIG.TL_COL_FORM_SENT, CONFIG.TL_COL_FORM_SUBMITTED, CONFIG.TL_COL_WAITLIST_FLAG, CONFIG.TL_COL_ONSCHEDULE, CONFIG.TL_COL_NOT_INTERESTED].forEach(colIdx => {
          try { tl.getRange(sheetRow, colIdx).insertCheckboxes(); } catch (e) {}
      });
      if (!tl.isColumnHiddenByUser(CONFIG.TL_COL_UID)) { tl.hideColumns(CONFIG.TL_COL_UID); }
    }
  } else { 
    if (sheetRow === 0) return;
    if (params.responsibleParty !== undefined) tl.getRange(sheetRow, CONFIG.TL_COL_RESPONSIBLE).setValue(params.responsibleParty);
    if (params.addToWaitingList !== undefined) tl.getRange(sheetRow, CONFIG.TL_COL_WAITLIST_FLAG).setValue(params.addToWaitingList);
    if (params.phone !== undefined) tl.getRange(sheetRow, CONFIG.TL_COL_PHONE).setValue(params.phone);
  }
}

function updateSharedPatientInfo(uid, newData, sourceSheetName) {
  let updatedPatientName = newData.patientName;
  let updatedEmail = newData.email;

  // Registry Sync
  if (sourceSheetName !== CONFIG.REGISTRY) {
    const regSheet = sheet_(CONFIG.REGISTRY);
    const regRowIndex = regSheet.getRange(1, CONFIG.REG_COL_UID, regSheet.getLastRow(), 1).getValues().flat().indexOf(uid);
    if (regRowIndex !== -1) {
      if (updatedPatientName) regSheet.getRange(regRowIndex + 1, CONFIG.REG_COL_PATIENT_NAME).setValue(updatedPatientName);
      if (updatedEmail) regSheet.getRange(regRowIndex + 1, CONFIG.REG_COL_EMAIL).setValue(updatedEmail);
    }
  }

  // Waiting List Sync
  if (sourceSheetName !== CONFIG.WAITING_LIST) {
    const wlSheet = sheet_(CONFIG.WAITING_LIST);
    const wlRow = findRowByUid_(wlSheet, uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);
    if (wlRow) {
      if (updatedPatientName) wlSheet.getRange(wlRow, CONFIG.WL_COL_PATIENT_NAME).setValue(updatedPatientName);
      if (updatedEmail) wlSheet.getRange(wlRow, CONFIG.WL_COL_EMAIL).setValue(updatedEmail);
    }
  }

  // Telephone Log Sync
  if (sourceSheetName !== CONFIG.TELEPHONE_LOG) {
    const tlSheet = sheet_(CONFIG.TELEPHONE_LOG);
    const tlRow = findRowByUid_(tlSheet, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);
    if (tlRow) {
      if (updatedPatientName) tlSheet.getRange(tlRow, CONFIG.TL_COL_PATIENT_NAME).setValue(updatedPatientName);
      if (updatedEmail) tlSheet.getRange(tlRow, CONFIG.TL_COL_EMAIL).setValue(updatedEmail);
    }
  }

  // Intake Tabs Sync
  const allSheets = SS.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const s = allSheets[i];
    if (SYSTEM_SHEET_NAMES.includes(s.getName()) || s.getName() === sourceSheetName) continue;
    try {
      if (s.getRange(CONFIG.INTAKE_CELL_UID).getValue() === uid) {
        if (updatedPatientName) s.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).setValue(updatedPatientName);
        if (updatedEmail) s.getRange(CONFIG.INTAKE_CELL_EMAIL).setValue(updatedEmail);
      }
    } catch (e) {}
  }
  
  if (updatedPatientName) renameIntakeTabsForUid_(uid, updatedPatientName);
}

function syncActive_(uid, newVal, source) {
  if (source !== CONFIG.WAITING_LIST) {
    const wl = sheet_(CONFIG.WAITING_LIST);
    const wlRow = findRowByUid_(wl, uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);
    if (wlRow) wl.getRange(wlRow, CONFIG.WL_COL_ACTIVE).setValue(newVal);
  }
  if (source !== CONFIG.TELEPHONE_LOG) {
    const tl = sheet_(CONFIG.TELEPHONE_LOG);
    const tlRow = findRowByUid_(tl, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);
    if (tlRow) tl.getRange(tlRow, CONFIG.TL_COL_ONSCHEDULE).setValue(newVal);
  }
  findIntakeSheetsByUid_(uid).forEach(sh => {
    if (source !== sh.getName()) {
      sh.getRange(CONFIG.INTAKE_CELL_ACTIVE).setValue(newVal);
      setIntakeTabColor_(sh);
    }
  });
}

function syncSpotFound_(uid, newVal, source) {
  if (source !== CONFIG.WAITING_LIST) {
    const wl = sheet_(CONFIG.WAITING_LIST);
    const wlRow = findRowByUid_(wl, uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);
    if (wlRow) wl.getRange(wlRow, CONFIG.WL_COL_SPOT_FOUND).setValue(newVal);
  }
  findIntakeSheetsByUid_(uid).forEach(sh => {
    if (source !== sh.getName()) {
      sh.getRange(CONFIG.INTAKE_CELL_SPOT_FOUND).setValue(newVal);
      setIntakeTabColor_(sh);
    }
  });
}

function syncNotInterested_(uid, newVal, source) {
  const wl = sheet_(CONFIG.WAITING_LIST);
  const wlRow = findRowByUid_(wl, uid, CONFIG.WL_COL_UID, CONFIG.WL_HEADER_ROWS);
  if (wlRow) {
    if (source !== CONFIG.WAITING_LIST) wl.getRange(wlRow, CONFIG.WL_NOT_INTERESTED).setValue(newVal);
    if (newVal) {
      wl.getRange(wlRow, CONFIG.WL_COL_ACTIVE).setValue(false);
      wl.getRange(wlRow, CONFIG.WL_COL_SPOT_FOUND).setValue(false);
    }
  }

  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  const tlRow = findRowByUid_(tl, uid, CONFIG.TL_COL_UID, CONFIG.TL_HEADER_ROWS);
  if (tlRow) {
    if (source !== CONFIG.TELEPHONE_LOG) tl.getRange(tlRow, CONFIG.TL_COL_NOT_INTERESTED).setValue(newVal);
    if (newVal) tl.getRange(tlRow, CONFIG.TL_COL_ONSCHEDULE).setValue(false);
  }

  findIntakeSheetsByUid_(uid).forEach(sh => {
    if (source !== sh.getName()) sh.getRange(CONFIG.INTAKE_CELL_NOT_INTERESTED).setValue(newVal);
    if (newVal) {
      sh.getRange(CONFIG.INTAKE_CELL_ACTIVE).setValue(false);
      sh.getRange(CONFIG.INTAKE_CELL_SPOT_FOUND).setValue(false);
    }
    setIntakeTabColor_(sh);
  });
}

function addNewTelephoneLogEntry() {
  const tlSheet = sheet_(CONFIG.TELEPHONE_LOG);
  if (!tlSheet) return;

  const numCols = Math.max(tlSheet.getMaxColumns(), CONFIG.TL_COL_ONSCHEDULE);
  const newRowData = new Array(numCols).fill('');

  newRowData[CONFIG.TL_COL_UID - 1]           = newUid_();
  newRowData[CONFIG.TL_COL_CALL_OUTCOME  - 1] = CONFIG.DEFAULT_TL_CALL_OUTCOME_PLACEHOLDER;
  newRowData[CONFIG.TL_COL_RESPONSIBLE   - 1] = CONFIG.DEFAULT_TL_DISABLE_FORM_NOTE;
  newRowData[CONFIG.TL_COL_EMAIL         - 1] = CONFIG.DEFAULT_TL_EMAIL_INFO_NOTE;
  newRowData[CONFIG.TL_COL_DATE - 1] = new Date();
  
  [CONFIG.TL_COL_FORM_SENT, CONFIG.TL_COL_FORM_SUBMITTED, CONFIG.TL_COL_WAITLIST_FLAG, CONFIG.TL_COL_ONSCHEDULE, CONFIG.TL_COL_NOT_INTERESTED].forEach(i => {
    newRowData[i-1] = false;
  });

  tlSheet.appendRow(newRowData);
  const newRowNum = tlSheet.getLastRow();
  tlSheet.getRange(newRowNum, CONFIG.TL_COL_DATE).setNumberFormat('MM/dd/yyyy');

  [CONFIG.TL_COL_FORM_SENT, CONFIG.TL_COL_FORM_SUBMITTED, CONFIG.TL_COL_WAITLIST_FLAG, CONFIG.TL_COL_ONSCHEDULE, CONFIG.TL_COL_NOT_INTERESTED].forEach(colIdx => {
    try { tlSheet.getRange(newRowNum, colIdx).insertCheckboxes(); } catch (e) {}
  });
  
  tlSheet.getRange(newRowNum, CONFIG.TL_COL_RESPONSIBLE).activate();
}

function listOpenTLRows() {
  return getRecentTLWithoutIntake_();
}

function getRecentTLWithoutIntake_(limit = 25, scanWindow = 400) {
  const tl = sheet_(CONFIG.TELEPHONE_LOG);
  const lastRow = tl.getLastRow();
  if (lastRow <= CONFIG.TL_HEADER_ROWS) return [];

  const reg = sheet_(CONFIG.REGISTRY);
  const regLast = reg.getLastRow();
  const hasIntake = new Set();
  if (regLast > 1) {
    const uidCol = reg.getRange(2, CONFIG.REG_COL_UID, regLast - 1, 1).getValues().flat();
    const flagCol = reg.getRange(2, CONFIG.REG_COL_HAS_INTAKE, regLast - 1, 1).getValues().flat();
    for (let i = 0; i < uidCol.length; i++) if (flagCol[i] === true) hasIntake.add(uidCol[i]);
  }

  const firstDataRow = CONFIG.TL_HEADER_ROWS + 1;
  const startRow = Math.max(firstDataRow, lastRow - scanWindow + 1);
  const numRows = lastRow - startRow + 1;
  const maxCol = Math.max(CONFIG.TL_COL_UID, CONFIG.TL_COL_DATE, CONFIG.TL_COL_RESPONSIBLE, CONFIG.TL_COL_PATIENT_NAME, CONFIG.TL_COL_EMAIL);
  const block = tl.getRange(startRow, 1, numRows, maxCol).getValues();

  const tz = Session.getScriptTimeZone() || 'UTC';
  const out = [];
  
  for (let i = block.length - 1; i >= 0 && out.length < limit; i--) {
    const row = block[i];
    const uid = row[CONFIG.TL_COL_UID - 1];
    if (!uid || hasIntake.has(uid)) continue;

    const dateStr = (row[CONFIG.TL_COL_DATE - 1] instanceof Date) ? Utilities.formatDate(row[CONFIG.TL_COL_DATE - 1], tz, 'MM/dd') : '(no date)';
    const resp    = row[CONFIG.TL_COL_RESPONSIBLE - 1] || '(no RP)';
    const name    = row[CONFIG.TL_COL_PATIENT_NAME - 1] || '(no name)';
    const mail    = row[CONFIG.TL_COL_EMAIL - 1] || '(no e-mail)';

    out.push({ uid, label: `${dateStr} – ${resp} (${name}) – ${mail}` });
  }
  return out;
}