/* ================================================================
 * TRIGGER ENTRY POINTS
 * ================================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Intake Tools')
    .addItem('✙🧍‍♂️ New Intake Tab', 'openIntakeCreator')
    .addSeparator()
    .addItem('✙📒 New Telephone Log Entry', 'addNewTelephoneLogEntry')
    .addToUi();
}

function onFormSubmitTrigger(e) {
  const sheetName = e.range.getSheet().getName();
  const formKey = Object.keys(FORMS).find(k => FORMS[k].responseSheet === sheetName);
  if (!formKey) return;

  const header  = e.range.getSheet().getRange(1,1,1,e.range.getLastColumn()).getValues()[0];
  const uidCol  = header.indexOf('UID') + 1;
  const uid     = e.range.getSheet().getRange(e.range.getRow(), uidCol).getValue();
  if (!uid) return;

  const hist = sheet_(CONFIG.HISTORY);
  const data = hist.getDataRange().getValues();
  for (let i=1;i<data.length;i++){
    if (data[i][CONFIG.HISTORY_COL_UID-1] === uid && data[i][CONFIG.HISTORY_COL_FORM-1] === formKey) {
      hist.getRange(i+1, CONFIG.HISTORY_COL_SUBMITTED).setValue(true);
      break;
    }
  }
}

function editHandler(e) {
  if (!e || !e.range) return;

  const range       = e.range;
  const sh          = range.getSheet();
  const sheetName   = sh.getName();
  const row         = range.getRow();
  const col         = range.getColumn();
  const editedValue = e.value;
  let uid;
  let newData = {};
  let syncSharedInfo = false;

  // 1. INTAKE TABS
  if (!SYSTEM_SHEET_NAMES.includes(sheetName)) {
    const uidFromTab = sh.getRange(CONFIG.INTAKE_CELL_UID).getValue();
    if (uidFromTab) {
      uid = uidFromTab;
      const cellA1 = range.getA1Notation().split(':')[0];

      // Auto-Email Triggers
      if (cellA1 === CONFIG.INTAKE_CELL_ACTIVE && editedValue === 'TRUE') {
         const email = sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue();
         const isRec = sh.getRange(CONFIG.INTAKE_CELL_RECREATIONAL).getValue() === true;
         if (validateEmail(email)) {
            const formKey = isRec ? 'RECREATIONAL_START' : 'THERAPY_START';
            if (sendForm_(formKey, { uid, email, responsible: sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue() })) {
              sh.getRange(CONFIG.INTAKE_TAB_FORMS_SENT_NOTE).setValue(isRec ? 'Recreational Forms Sent!' : 'Therapy Forms Sent!');
            }
         } else { sh.getRange(CONFIG.INTAKE_CELL_ACTIVE).setValue(false); }
      }

      if (cellA1 === CONFIG.INTAKE_CELL_FINANCIAL_AID && editedValue === 'TRUE') {
         const email = sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue();
         if (validateEmail(email)) {
            if(sendForm_('FINANCIAL_AID_2025', { uid, email, responsible: sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue() })) {
               sh.getRange(CONFIG.INTAKE_TAB_FINANCIAL_AID_NOTE).setValue('Form Sent!');
            }
         } else { sh.getRange(CONFIG.INTAKE_CELL_FINANCIAL_AID).setValue(false); }
      }

      if (cellA1 === CONFIG.INTAKE_CELL_TELEHEALTH_LINK && editedValue === 'TRUE') {
         const dateStr = sh.getRange(CONFIG.INTAKE_CELL_TELEHEALTH_DATE).getDisplayValue();
         const timeStr = sh.getRange(CONFIG.INTAKE_CELL_TELEHEALTH_TIME).getDisplayValue();
         const email = sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue();

         // Check if Date/Time are empty OR contain the placeholder text "Date"/"Time"
         if (!dateStr || !timeStr || dateStr.includes('Date') || timeStr.includes('Time')) {
             sh.getRange(CONFIG.INTAKE_CELL_TELEHEALTH_LINK).setValue(false); // Uncheck the box
             SpreadsheetApp.getUi().alert(
               'Missing Appointment Info', 
               'Please enter a valid Date and Time for the appointment before sending the link.', 
               SpreadsheetApp.getUi().ButtonSet.OK
             );
         } else if (validateEmail(email)) {
             if (sendForm_('TELEHEALTH_APPT', { uid, email, responsible: sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue(), apptDate: `${dateStr}, at ${timeStr}` })) {
                 sh.getRange(CONFIG.INTAKE_TAB_TELEHEALTH_LINK_NOTE).setValue('Link Sent!');
             }
         } else { 
            sh.getRange(CONFIG.INTAKE_CELL_TELEHEALTH_LINK).setValue(false); 
            SpreadsheetApp.getUi().alert('Missing Email', 'Please enter a valid email address.', SpreadsheetApp.getUi().ButtonSet.OK);
         }
      }

      // Checkbox Sync
      if (cellA1 === CONFIG.INTAKE_CELL_ACTIVE) syncActive_(uid, editedValue === 'TRUE', sheetName);
      if (cellA1 === CONFIG.INTAKE_CELL_SPOT_FOUND) syncSpotFound_(uid, editedValue === 'TRUE', sheetName);
      if (cellA1 === CONFIG.INTAKE_CELL_NOT_INTERESTED) syncNotInterested_(uid, editedValue === 'TRUE', sheetName);

      // Call Completed -> Add to Waiting List
      if (cellA1 === CONFIG.INTAKE_CELL_CALL_COMPLETED && editedValue === 'TRUE') {
         const wl = sheet_(CONFIG.WAITING_LIST);
         const uidList = wl.getRange(2, CONFIG.WL_COL_UID, wl.getLastRow()-1, 1).getValues().flat();
         if (uidList.indexOf(uid) === -1) {
            syncWaitingList_({
                uid,
                patientName: sh.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).getValue(),
                responsibleParty: sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue(),
                phone: sh.getRange(CONFIG.INTAKE_CELL_PHONE).getValue(),
                email: sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue(),
                dob: sh.getRange(CONFIG.INTAKE_CELL_DOB).getValue(),
                potentialService: sh.getRange(CONFIG.INTAKE_CELL_POTENTIAL_SERVICE).getValue(),
                intakeDate: new Date(),
                isInitialCreation: true
            });
            syncTelephoneLog_({ uid, addToWaitingList:true, isInitialCreation:false });
            sh.getRange(CONFIG.INTAKE_TAB_WL_NOTE).clearContent();
            setIntakeTabColor_(sh);
         }
      }

      // Name/Email Sync
      if (cellA1 === CONFIG.INTAKE_CELL_PATIENT_NAME) { newData.patientName = editedValue; syncSharedInfo = true; }
      if (cellA1 === CONFIG.INTAKE_CELL_EMAIL) {
         if (editedValue && !validateEmail(editedValue)) { syncSharedInfo = false; } else { newData.email = editedValue; syncSharedInfo = true; }
      }

      // Sync other fields
      if (INTAKE_SHEET_RELEVANT_CELLS_FOR_ORIGINAL_SYNC.hasOwnProperty(cellA1)) {
        const payload = {
            uid,
            patientName: sh.getRange(CONFIG.INTAKE_CELL_PATIENT_NAME).getValue(),
            responsibleParty: sh.getRange(CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY).getValue(),
            phone: sh.getRange(CONFIG.INTAKE_CELL_PHONE).getValue(),
            email: sh.getRange(CONFIG.INTAKE_CELL_EMAIL).getValue(),
            dob: sh.getRange(CONFIG.INTAKE_CELL_DOB).getValue() instanceof Date ? sh.getRange(CONFIG.INTAKE_CELL_DOB).getValue() : null,
            potentialService: sh.getRange(CONFIG.INTAKE_CELL_POTENTIAL_SERVICE).getValue(),
            active: sh.getRange(CONFIG.INTAKE_CELL_ACTIVE).getValue(),
            spotFound: sh.getRange(CONFIG.INTAKE_CELL_SPOT_FOUND).getValue(),
            diagnosisNotes: sh.getRange(CONFIG.INTAKE_CELL_DIAGNOSIS_NOTES).getValue(),
            isInitialCreation: false
        };
        syncWaitingList_(payload);
        syncTelephoneLog_(payload);
      }
    }
  }

  // 2. WAITING LIST
  else if (sheetName === CONFIG.WAITING_LIST && row > 1) {
    uid = sh.getRange(row, CONFIG.WL_COL_UID).getValue();
    if (uid) {
        if (col === CONFIG.WL_COL_ACTIVE) syncActive_(uid, editedValue === 'TRUE', CONFIG.WAITING_LIST);
        if (col === CONFIG.WL_COL_SPOT_FOUND) syncSpotFound_(uid, editedValue === 'TRUE', CONFIG.WAITING_LIST);
        if (col === CONFIG.WL_NOT_INTERESTED) syncNotInterested_(uid, editedValue === 'TRUE', sheetName);
        if (col === CONFIG.WL_COL_PATIENT_NAME) { newData.patientName = editedValue; syncSharedInfo = true; }
        if (col === CONFIG.WL_COL_EMAIL) { if (validateEmail(editedValue)) { newData.email = editedValue; syncSharedInfo = true; } }
    }
  }

  // 3. TELEPHONE LOG
  else if (sheetName === CONFIG.TELEPHONE_LOG && row > 1) {
    uid = sh.getRange(row, CONFIG.TL_COL_UID).getValue();
    if (uid) {
        if (col === CONFIG.TL_COL_ONSCHEDULE) syncActive_(uid, editedValue === 'TRUE', CONFIG.TELEPHONE_LOG);
        if (col === CONFIG.TL_COL_NOT_INTERESTED) syncNotInterested_(uid, editedValue === 'TRUE', CONFIG.TELEPHONE_LOG);
        if (col === CONFIG.TL_COL_PATIENT_NAME) { newData.patientName = editedValue; syncSharedInfo = true; }
        if (col === CONFIG.TL_COL_EMAIL) { 
            if (validateEmail(editedValue)) { newData.email = editedValue; syncSharedInfo = true; }
            
            // Auto-Send Form on Email Entry
            const sentCell = sh.getRange(row, CONFIG.TL_COL_FORM_SENT);
            if (editedValue && validateEmail(editedValue) && sentCell.getValue() !== true) {
                let resp = sh.getRange(row, CONFIG.TL_COL_RESPONSIBLE).getValue();
                if (isPlaceholderResponsibleParty(resp)) resp = 'thank you for contacting Special Strides';
                if (sendForm_('INTAKE', { uid, email: editedValue, responsible: resp })) {
                    sentCell.setValue(true);
                }
            }
        }
    }
  }

  // 4. REGISTRY
  else if (sheetName === CONFIG.REGISTRY && row > 1) {
    uid = sh.getRange(row, CONFIG.REG_COL_UID).getValue();
    if (uid) {
        if (col === CONFIG.REG_COL_PATIENT_NAME) { newData.patientName = editedValue; syncSharedInfo = true; }
        if (col === CONFIG.REG_COL_EMAIL && validateEmail(editedValue)) { newData.email = editedValue; syncSharedInfo = true; }
    }
  }

  // 5. PROPAGATE SHARED INFO
  if (syncSharedInfo && uid) updateSharedPatientInfo(uid, newData, sheetName);
}