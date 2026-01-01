/* ================================================================
 * NEW YEAR ROLLOVER LOGIC
 * ================================================================ */

function openRolloverDialog() {
  const ui = SpreadsheetApp.getUi();
  const date = new Date();
  const currentYear = date.getFullYear();
  const month = date.getMonth(); // 0 = Jan
  
  // Suggest Next Year unless it is January
  const suggestedYear = (month === 0) ? currentYear : currentYear + 1;

  const result = ui.prompt(
    'Create New Year Workbook', 
    `This will create a brand new workbook for the upcoming year.\n\n` +
    `It will:\n` + 
    `1. Copy this spreadsheet\n` + 
    `2. Clear all patient tabs in the copy\n` + 
    `3. Clear history/logs in the copy\n` +
    `4. Duplicate the Intake Form and attach it\n\n` +
    `Enter the Year for the new workbook:`, 
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;
  
  const targetYear = result.getResponseText().trim();
  if (!targetYear.match(/^20\d{2}$/)) {
    ui.alert('Invalid Year', 'Please enter a valid 4-digit year (e.g. 2026).', ui.ButtonSet.OK);
    return;
  }

  // Ask for Financial Aid URL upfront
  const finAidResponse = ui.prompt(
    'Financial Aid Form',
    `Please paste the public URL (Link) for the ${targetYear} Financial Aid Form.\n`+
    `(If you don't have it yet, leave blank and update System_Form_Links later).`,
    ui.ButtonSet.OK
  );
  const finAidUrl = finAidResponse.getResponseText().trim();

  createYearlyCopy_(targetYear, finAidUrl);
}

function createYearlyCopy_(targetYear, finAidUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const currentName = ss.getName();
  
  // 1. Determine new name
  // Tries to replace "2025" with "2026". If no year in name, appends it.
  let newName = currentName.match(/20\d{2}/) 
    ? currentName.replace(/20\d{2}/, targetYear) 
    : `${targetYear} ${currentName}`;
    
  ss.toast('Creating copy... This may take a minute.', 'Processing', 30);

  // 2. Make Copy
  const newFile = DriveApp.getFileById(ss.getId()).makeCopy(newName);
  const newSS = SpreadsheetApp.openById(newFile.getId());
  
  // 3. Clean up the NEW sheet
  cleanUpNewSheet_(newSS, targetYear);
  
  // 4. Handle Google Form (Intake)
  // We look for the form linked to "Form Responses"
  let newFormUrl = '(Form not found - manual link required)';
  try {
    const formUrl = ss.getFormUrl(); // Gets form linked to THIS sheet
    if (formUrl) {
      const oldForm = FormApp.openByUrl(formUrl);
      const oldFormFile = DriveApp.getFileById(oldForm.getId());
      
      // Copy the form
      const newFormFile = oldFormFile.makeCopy(`${targetYear} Intake Form`);
      const newForm = FormApp.openById(newFormFile.getId());
      
      // Link new form to new spreadsheet
      newForm.setDestination(FormApp.DestinationType.SPREADSHEET, newSS.getId());
      
      // Important: Rename the response tab in the new sheet back to "Form Responses"
      // (Linking creates "Form Responses 2")
      fixFormDestinationTab_(newSS);

      newFormUrl = newForm.getPublishedUrl();
      newForm.setConfirmationMessage(`Thank you. Your ${targetYear} intake has been received.`);
    }
  } catch (e) {
    console.error('Error handling forms: ' + e.message);
  }

  // 5. Update System_Form_Links in the NEW sheet
  updateSystemLinks_(newSS, targetYear, newFormUrl, finAidUrl);

  // 6. Final User Message
  const htmlOutput = HtmlService.createHtmlOutput(
    `<p><strong>Success!</strong> The new workbook has been created.</p>` +
    `<p>New File: <a href="${newSS.getUrl()}" target="_blank">${newName}</a></p>` +
    `<p><strong>Next Steps:</strong></p>` +
    `<ol>` +
    `<li>Open the new workbook.</li>` +
    `<li>Click the "Intake Tools" menu (it may take a moment to appear).</li>` +
    `<li>Select <strong>"Initialize New Year Triggers"</strong> to activate automation.</li>` +
    `</ol>`
  ).setWidth(400).setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Rollover Complete');
}

function cleanUpNewSheet_(targetSS, targetYear) {
  const sheets = targetSS.getSheets();
  const keepNames = [
    CONFIG.REGISTRY,
    CONFIG.TEMPLATE,
    CONFIG.FORM_LINKS,
    // Note: We use the DYNAMIC variable keys from current config, 
    // but we will look for them using partial matches or standard names because
    // the sheet names in the COPY are still 2025 until we rename them.
    'Telephone_Log', 
    'Waiting_List', 
    'Email_History',
    'Form Responses'
  ];

  // A. Delete Patient Tabs & Clear Data
  sheets.forEach(sh => {
    const name = sh.getName();
    
    // Check if it's a system sheet (partial match allows finding "Telephone_Log_2025")
    const isSystem = keepNames.some(k => name.includes(k));

    if (!isSystem) {
      targetSS.deleteSheet(sh);
    } else {
      // It is a system sheet: Clear Content (keep headers)
      if (name.includes('Form Responses')) {
         // Clear all responses except header
         if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
      } 
      else if (name.includes('Telephone_Log')) {
        // Use a fixed column count from CONFIG so we don't depend on "last used" column
        clearBelowHeader_(sh, CONFIG.TL_HEADER_ROWS, CONFIG.TL_COL_NOT_INTERESTED);
        sh.setName(`Telephone_Log_${targetYear}`);
      }

      else if (name.includes('Waiting_List')) {
        clearBelowHeader_(sh, CONFIG.WL_HEADER_ROWS, CONFIG.WL_INTAKE_COMPLETED);
        sh.setName(`Waiting_List_${targetYear}`);
      }
      else if (name.includes('Email_History')) {
         if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
         sh.setName(`Email_History_${targetYear}`); // Rename
      }
      else if (name === CONFIG.REGISTRY) {
         if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
      }
    }
  });
}

function clearBelowHeader_(sh, headerRows, maxCol) {
  const maxRows = sh.getMaxRows();
  const numRows = maxRows - headerRows;
  if (numRows <= 0) return;

  const cols = Math.max(1, maxCol || sh.getMaxColumns());
  sh.getRange(headerRows + 1, 1, numRows, cols).clearContent();
}


function fixFormDestinationTab_(ss) {
  // Linking a form creates "Form Responses X". We need it to be "Form Responses".
  const sheets = ss.getSheets();
  let oldRespSheet = null;
  let newRespSheet = null;
  
  sheets.forEach(s => {
    if (s.getName() === CONFIG.FORM_RESPONSES) oldRespSheet = s;
    if (s.getName().startsWith('Form Responses') && s.getName() !== CONFIG.FORM_RESPONSES) newRespSheet = s;
  });

  if (oldRespSheet && newRespSheet) {
    // We already cleared the old one, but the Form is linked to the new one.
    // Safest: Delete the old one, rename the new one.
    ss.deleteSheet(oldRespSheet);
    newRespSheet.setName(CONFIG.FORM_RESPONSES);
  }
}

function updateSystemLinks_(ss, year, intakeUrl, finAidUrl) {
  const sheet = ss.getSheetByName(CONFIG.FORM_LINKS);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  // Columns: Key(1), URL(2), UID_Param(3), Name(4)
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    
    if (key === 'INTAKE' && intakeUrl) {
      sheet.getRange(i + 1, 2).setValue(intakeUrl);
    }
    
    if (key.includes('FINANCIAL_AID')) {
      // Update key name to new year
      sheet.getRange(i + 1, 1).setValue(`FINANCIAL_AID_${year}`);
      // Update Display Name
      sheet.getRange(i + 1, 4).setValue(`Financial Aid Application`); // Generic or Year specific
      // Update URL if provided
      if (finAidUrl) {
        sheet.getRange(i + 1, 2).setValue(finAidUrl);
      } else {
        sheet.getRange(i + 1, 2).setValue('(Update this link)');
      }
    }
  }
}

/* ------------------------------------------------
 * TRIGGER INSTALLER (Run this in the NEW file)
 * ------------------------------------------------ */
function initializeNewYearTriggers() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Initialize System',
    'This will install the automation triggers for this new workbook.\nAre you sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  // 1. Delete existing triggers (copies usually don't have them, but safety first)
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
    
  ui.alert('Success', 'System initialized for the new year.', ui.ButtonSet.OK);
}