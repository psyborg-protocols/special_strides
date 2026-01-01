/* ================================================================
 * NEW YEAR ROLLOVER LOGIC
 * ================================================================ */

/* 1. ROLLOVER (CREATE NEW WORKBOOK) -------------------------- */

function openRolloverDialog() {
  const html = HtmlService.createHtmlOutputFromFile('RolloverDialog')
    .setWidth(450)
    .setHeight(400); // Increased height to fit inputs and spinner
  SpreadsheetApp.getUi().showModalDialog(html, 'Create New Year Workbook');
}

/**
 * Backend function called by RolloverDialog.html
 */
function processRollover(targetYear, finAidUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentName = ss.getName();
    
    // 1. Determine new name
    let newName = currentName.match(/20\d{2}/) 
      ? currentName.replace(/20\d{2}/, targetYear) 
      : `${targetYear} ${currentName}`;
      
    // 2. Make Copy
    const newFile = DriveApp.getFileById(ss.getId()).makeCopy(newName);
    const newSS = SpreadsheetApp.openById(newFile.getId());
    
    // 3. Clean up the NEW sheet
    cleanUpNewSheet_(newSS, targetYear);
    
// 4. Handle Google Form (Intake) - PATCHED: Uses System_Form_Links only
    let newFormUrl = '(Form not found - manual link required)';
    try {
      let formUrl = null;

      // A) Look up the 'INTAKE' URL in the System_Form_Links sheet
      const linkSheet = ss.getSheetByName(CONFIG.FORM_LINKS);
      if (linkSheet) {
        const data = linkSheet.getDataRange().getValues();
        // Loop rows (skipping header) to find the INTAKE key
        for (let i = 1; i < data.length; i++) {
          if (data[i][CONFIG.LINKS_COL_KEY - 1] === 'INTAKE') {
            formUrl = data[i][CONFIG.LINKS_COL_URL - 1];
            break;
          }
        }
      }

      // B) If URL found, process the copy
      if (formUrl) {
        const oldForm = FormApp.openByUrl(formUrl);
        const oldFormFile = DriveApp.getFileById(oldForm.getId());
        
        // Copy the form
        const newFormFile = oldFormFile.makeCopy(`${targetYear} Intake Form`);
        const newForm = FormApp.openById(newFormFile.getId());
        
        // Link new form to new spreadsheet
        newForm.setDestination(FormApp.DestinationType.SPREADSHEET, newSS.getId());
        fixFormDestinationTab_(newSS); // Fix the "Form Responses 2" issue

        // Update settings
        newFormUrl = newForm.getEditUrl(); // Store EDIT url so next year's rollover can find the ID
        newForm.setConfirmationMessage(`Thank you. Your ${targetYear} intake has been received.`);
      } else {
        Logger.log('Warning: No INTAKE row found in System_Form_Links.');
      }
    } catch (e) {
      console.error('Error handling forms: ' + e.message);
    }

    // 5. Update System_Form_Links in the NEW sheet
    updateSystemLinks_(newSS, targetYear, newFormUrl, finAidUrl);

    // Return success object to the HTML runner
    return {
      success: true,
      newUrl: newSS.getUrl(),
      newName: newName
    };

  } catch (e) {
    return { success: false, error: e.message };
  }
}


/* 2. TRIGGER INSTALLATION (RUN IN NEW FILE) ------------------ */

function initializeNewYearTriggers() {
  const html = HtmlService.createHtmlOutputFromFile('TriggerDialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Initialize System');
}

/**
 * Backend function called by TriggerDialog.html
 */
function processTriggerInstall() {
  try {
    // 1. Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => ScriptApp.deleteTrigger(t));
    
    // 2. Install Form Submit Trigger
    ScriptApp.newTrigger('onFormSubmitTrigger')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onFormSubmit()
      .create();
      
    // 3. Install Edit Trigger
    ScriptApp.newTrigger('editHandler')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/* -----------------------------------------------------------
 * HELPER FUNCTIONS (Logic remains largely the same)
 * ----------------------------------------------------------- */

function cleanUpNewSheet_(targetSS, targetYear) {
  const sheets = targetSS.getSheets();
  const deleteRequests = [];
  
  const systemKeywords = [
    'Patient_Registry', 'Intake_Template', 'System_Form_Links',
    'Telephone_Log', 'Waiting_List', 'Email_History', 'Form Responses'
  ];

  sheets.forEach(sh => {
    const name = sh.getName();
    const isSystem = systemKeywords.some(key => name.includes(key));

    if (!isSystem) {
      deleteRequests.push({ deleteSheet: { sheetId: sh.getSheetId() } });
    } else {
      // System Sheet Cleanup
      if (name.includes('Form Responses')) {
         if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
      } 
      else if (name.includes('Telephone_Log')) {
         if (sh.getLastRow() > CONFIG.TL_HEADER_ROWS) {
            sh.getRange(CONFIG.TL_HEADER_ROWS + 1, 1, sh.getLastRow() - CONFIG.TL_HEADER_ROWS, sh.getLastColumn()).clearContent();
         }
         sh.setName(`Telephone_Log_${targetYear}`); 
      }
      else if (name.includes('Waiting_List')) {
         if (sh.getLastRow() > CONFIG.WL_HEADER_ROWS) {
           sh.getRange(CONFIG.WL_HEADER_ROWS + 1, 1, sh.getLastRow() - CONFIG.WL_HEADER_ROWS, sh.getLastColumn()).clearContent();
         }
         sh.setName(`Waiting_List_${targetYear}`);
      }
      else if (name.includes('Email_History')) {
         if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
         sh.setName(`Email_History_${targetYear}`);
      }
      else if (name.includes('Patient_Registry')) {
         if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
      }
    }
  });

  if (deleteRequests.length > 0) {
    try {
      Sheets.Spreadsheets.batchUpdate({requests: deleteRequests}, targetSS.getId());
    } catch (e) { console.error(e); }
  }
}

function fixFormDestinationTab_(ss) {
  const sheets = ss.getSheets();
  let oldRespSheet = null;
  let newRespSheet = null;
  
  sheets.forEach(s => {
    if (s.getName() === CONFIG.FORM_RESPONSES) oldRespSheet = s;
    if (s.getName().startsWith('Form Responses') && s.getName() !== CONFIG.FORM_RESPONSES) newRespSheet = s;
  });

  if (oldRespSheet && newRespSheet) {
    ss.deleteSheet(oldRespSheet);
    newRespSheet.setName(CONFIG.FORM_RESPONSES);
  }
}

function updateSystemLinks_(ss, year, intakeUrl, finAidUrl) {
  const sheet = ss.getSheetByName(CONFIG.FORM_LINKS);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    if (key === 'INTAKE' && intakeUrl) {
      sheet.getRange(i + 1, 2).setValue(intakeUrl);
    }
    if (key.includes('FINANCIAL_AID')) {
      sheet.getRange(i + 1, 1).setValue(`FINANCIAL_AID_${year}`);
      sheet.getRange(i + 1, 4).setValue(`Financial Aid Application`); 
      if (finAidUrl) sheet.getRange(i + 1, 2).setValue(finAidUrl);
      else sheet.getRange(i + 1, 2).setValue('(Update this link)');
    }
  }
}