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

const TEMPLATE_SS_ID = '11EfX1ZAlpcUM_JgKaaSXJ1C4SF7KT3rphbP6JrbFxl8';
const TEMPLATE_FORM_ID = '1SUz8-4a33F01t-N1Sq71XHIqugUsaKIBwYgRu_ovp1k';

function processRollover(targetYear, finAidUrl) {
  try {
    // 1. CREATE THE NEW FOLDER
    // We create a container folder so new year's files stay organized
    const newFolder = DriveApp.createFolder(`Special Strides Intake ${targetYear}`);

    // 2. COPY THE TEMPLATE FILES INTO THE NEW FOLDER
    const templateSSFile = DriveApp.getFileById(TEMPLATE_SS_ID);
    const templateFormFile = DriveApp.getFileById(TEMPLATE_FORM_ID);

    const newSSFile = templateSSFile.makeCopy(`${targetYear} Intake Communication Log`, newFolder);
    const newFormFile = templateFormFile.makeCopy(`New Client/Participant Intake ${targetYear}`, newFolder);

    // 3. OPEN THE NEW FILES
    const newSS = SpreadsheetApp.openById(newSSFile.getId());
    const newForm = FormApp.openById(newFormFile.getId());

    // Force the form to go live immediately
    newForm.setPublished(true);

    // 4. LINK THEM
    // Since we copied files individually, they are NOT linked yet.
    // We must link them, which will create the "Form Responses" tab mess.

    newForm.setDestination(FormApp.DestinationType.SPREADSHEET, newSS.getId());

    // 5. GET FORM URL
    const newFormUrl = newForm.getEditUrl(); 
    
    // 6. UPDATE INTERNAL LINKS
    updateSystemLinks_(newSS, targetYear, newFormUrl, finAidUrl);

    // 7. RENAME TABS
    setupNewYearTabs_(newSS, targetYear);

    return {
      success: true,
      newUrl: newSS.getUrl(),
      newName: newSSFile.getName()
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
    // --- Check Document Properties for an existing trigger lock ---
    const props = PropertiesService.getDocumentProperties();
    const installedBy = props.getProperty('TRIGGERS_INSTALLED_BY');

    if (installedBy) {
       // Abort installation and tell the user who currently owns the triggers
       return { 
         success: false, 
         error: `Triggers are already installed for this workbook by: ${installedBy}.\n\nTo prevent duplicate emails, you cannot install a second set of triggers. If ownership needs to be transferred, the current triggers must be deleted and the system lock must be reset.` 
       };
    }
    // -------------------------------------------------------------------

    const ss = SpreadsheetApp.getActive();
    
    // Rename 'Form Responses X' to 'Form Responses'
    const targetName = 'Form Responses';
    if (!ss.getSheetByName(targetName)) {
      const sheets = ss.getSheets();
      const sheetToRename = sheets.find(sheet => /^Form Responses \d+$/.test(sheet.getName()));
      if (sheetToRename) {
        sheetToRename.setName(targetName);
      }
    }
    
    // 1. Delete existing triggers (for the CURRENT user only)
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => ScriptApp.deleteTrigger(t));
    
    // 2. Install Form Submit Trigger
    ScriptApp.newTrigger('onFormSubmitTrigger')
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();
      
    // 3. Install Edit Trigger
    ScriptApp.newTrigger('editHandler')
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    // --- Lock the system and record the owner's email ---
    const currentUser = Session.getEffectiveUser().getEmail() || 'Unknown User';
    props.setProperty('TRIGGERS_INSTALLED_BY', currentUser);
    // ---------------------------------------------------------

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function resetTriggerLock() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('TRIGGERS_INSTALLED_BY');
  SpreadsheetApp.getUi().alert('Success', 'The trigger installation lock has been removed. A new user can now install triggers.', SpreadsheetApp.getUi().ButtonSet.OK);
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
    // UPDATED: Check for generic 'FINANCIAL_AID' instead of year-specific
    if (key === 'FINANCIAL_AID') {
      // Don't change the key in col 1. Keep it as FINANCIAL_AID.
      // Update Label (Col 4)
      sheet.getRange(i + 1, 4).setValue(`Financial Aid Application ${year}`); 
      // Update URL (Col 2)
      if (finAidUrl) sheet.getRange(i + 1, 2).setValue(finAidUrl);
      else sheet.getRange(i + 1, 2).setValue('(Update this link)');
    }
  }
}