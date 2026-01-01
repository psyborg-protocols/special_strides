/* ================================================================
 * NEW YEAR ROLLOVER LOGIC
 * ================================================================ */

const TEMPLATE_FOLDER_ID = '1lhgmsFF9FRdARoWs8y7tQjvDtS-bB1mq'

/* 1. ROLLOVER (CREATE NEW WORKBOOK) -------------------------- */

function openRolloverDialog() {
  const html = HtmlService.createHtmlOutputFromFile('RolloverDialog')
    .setWidth(450)
    .setHeight(400); // Increased height to fit inputs and spinner
  SpreadsheetApp.getUi().showModalDialog(html, 'Create New Year Workbook');
}


function processRollover(targetYear, finAidUrl) {
  try {
    // 1. COPY THE ENTIRE FOLDER
    // This performs the "magic" copy that preserves the Form<->Sheet link
    // and keeps the "Form Responses" tab intact (no "Form Responses 2").
    const templateFolder = DriveApp.getFolderById(TEMPLATE_FOLDER_ID);
    const newFolder = templateFolder.makeCopy(`Special Strides ${targetYear}`);
    
    // 2. LOCATE THE NEW FILES
    // We iterate through the new folder to find our new Sheet and Form
    let newSS = null;
    let newForm = null;
    
    const files = newFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const mime = file.getMimeType();
      
      if (mime === MimeType.GOOGLE_SHEETS) {
        newSS = SpreadsheetApp.openById(file.getId());
        file.setName(`${targetYear} Intake Communication Log`); // Rename file
      } 
      else if (mime === MimeType.GOOGLE_FORMS) {
        newForm = FormApp.openById(file.getId());
        file.setName(`New Client/Participant Intake ${targetYear}`); // Rename file
      }
    }

    if (!newSS || !newForm) throw new Error("Could not find Sheet or Form in template folder.");

    // 3. UPDATE SETTINGS 
    // The link already exists! We just need to update the text.
    const newFormUrl = newForm.getEditUrl(); // Store Edit URL for safety
    newForm.setConfirmationMessage(`Thank you. Your ${targetYear} intake has been received.`);
    
    setupNewYearTabs_(newSS, targetYear);

    // 4. UPDATE INTERNAL LINKS
    // Update the System_Form_Links tab so the new workbook knows its own form
    updateSystemLinks_(newSS, targetYear, newFormUrl, finAidUrl);

    return {
      success: true,
      newUrl: newSS.getUrl(),
      newName: newSS.getName()
    };

  } catch (e) {
    console.error(e);
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

/* ================================================================
 * HELPERS
 * ================================================================ */

function setupNewYearTabs_(ss, targetYear) {
  const sheets = ss.getSheets();
  
  // Define what the tabs are named in your MASTER TEMPLATE
  // vs. what they need to be named in the LIVE FILE.
  const mappings = [
    { contains: 'Telephone_Log',  prefix: 'Telephone_Log_' },
    { contains: 'Waiting_List',   prefix: 'Waiting_List_' },
    { contains: 'Email_History',  prefix: 'Email_History_' }
  ];

  sheets.forEach(sh => {
    const name = sh.getName();
    
    mappings.forEach(map => {
      // If the sheet name matches the template pattern (e.g. "Telephone_Log_Template")
      if (name.includes(map.contains)) {
        const newName = `${map.prefix}${targetYear}`;
        
        // Rename only if it's not already correct
        if (name !== newName) {
          sh.setName(newName);
        }
        
        // Ensure main tabs are visible
        if (map.contains !== 'Email_History') {
          sh.showSheet(); 
        }
      }
    });
  });
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
