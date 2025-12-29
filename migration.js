// =================================================================
//  ADMIN MIGRATION TOOL - NEW YEAR SYSTEM GENERATOR
// =================================================================

/**
 * 1. RUN THIS FUNCTION to generate next year's system.
 * It will:
 * - Copy the current spreadsheet (keeping folder location & permissions)
 * - Copy the Google Form
 * - Link the new Form to the new Sheet
 * - Clear all patient data
 * - Delete old patient intake tabs
 */
function createNewYearSystem() {
  const ui = SpreadsheetApp.getUi();
  const currentSS = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Ask for the new name (e.g., "Intake System 2026")
  const prompt = ui.prompt(
    "New Year Migration",
    "Enter the name for the new file (e.g., 'Intake System 2026'):",
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  const newName = prompt.getResponseText();

  ui.alert("Migration started. This may take 1-2 minutes. Please wait...");

  // 2. Copy the Spreadsheet
  // Find the folder where the current file lives to keep things organized
  const sourceFile = DriveApp.getFileById(currentSS.getId());
  const parents = sourceFile.getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  
  const newSSFile = sourceFile.makeCopy(newName, folder);
  const newSS = SpreadsheetApp.openById(newSSFile.getId());

  // 2b. Copy Permissions (Editors & Viewers) to ensure staff access is maintained
  try {
    const editors = sourceFile.getEditors().map(user => user.getEmail());
    const viewers = sourceFile.getViewers().map(user => user.getEmail());
    
    if (editors.length > 0) newSSFile.addEditors(editors);
    if (viewers.length > 0) newSSFile.addViewers(viewers);
  } catch (e) {
    Logger.log("Could not copy permissions (you might not have permission to add users): " + e.toString());
  }
  
  // 3. Handle the Google Form
  // We extract the form ID from the CONFIG url in the active sheet
  const formUrl = CONFIG.FORM_URL; 
  let formId;
  try {
    const form = FormApp.openByUrl(formUrl);
    formId = form.getId();
  } catch (e) {
    ui.alert("Error: Could not access the current Google Form. Check the URL in CONFIG.");
    return;
  }

  // Copy the Form and Link it
  // We place the form in the same folder as the new spreadsheet
  const newFormFile = DriveApp.getFileById(formId).makeCopy(`Intake Form - ${newName}`, folder);
  const newForm = FormApp.openById(newFormFile.getId());
  newForm.setDestination(FormApp.DestinationType.SPREADSHEET, newSS.getId());
  const newFormUrl = newForm.getPublishedUrl();

  // 4. Clean Data in the NEW Sheet
  // We use the GENERIC names because we assume you ran Phase 1 already.
  const sheetsToClean = [
    CONFIG.TELEPHONE_LOG, 
    CONFIG.WAITING_LIST, 
    CONFIG.HISTORY,
    CONFIG.REGISTRY
  ];

  sheetsToClean.forEach(name => {
    const sheet = newSS.getSheetByName(name);
    if (sheet) {
      // Keep header rows (assume 2 for most, 1 for Registry/History)
      const headerRows = (name === CONFIG.REGISTRY || name === CONFIG.HISTORY) ? 1 : 2;
      const lastRow = sheet.getLastRow();
      if (lastRow > headerRows) {
        sheet.getRange(headerRows + 1, 1, lastRow - headerRows, sheet.getLastColumn()).clearContent();
        // Clear checkboxes specifically if needed, but clearContent usually handles values
      }
    }
  });

  // 5. Delete Patient Tabs
  // We use the SYSTEM_SHEET_NAMES list to know what to KEEP.
  // Note: The new "Form Responses" tab created by the copy needs to be handled
  const sheets = newSS.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();
    // Delete if it's not a system sheet AND not the Form Responses sheet
    if (!SYSTEM_SHEET_NAMES.includes(name) && !name.includes("Form Responses")) {
      newSS.deleteSheet(sheet);
    }
  });

  // 6. Rename the new Form Responses tab to standard "Form Responses"
  // When we linked the form, it created a new tab like "Form Responses 2"
  const allSheets = newSS.getSheets();
  const newFormTab = allSheets.find(s => s.getName().startsWith("Form Responses") && s.getName() !== CONFIG.FORM_RESPONSES);
  if (newFormTab) {
    // Delete the old (copied) response tab first
    const oldTab = newSS.getSheetByName(CONFIG.FORM_RESPONSES);
    if (oldTab) newSS.deleteSheet(oldTab);
    // Rename the new one
    newFormTab.setName(CONFIG.FORM_RESPONSES);
  }

  // 7. Success Message with Instructions
  const msg = `
    ✅ MIGRATION COMPLETE!
    
    1. Open your Google Drive (check the same folder as this file) to find: "${newName}"
    2. Open the new spreadsheet.
    3. Go to Extensions > Apps Script > intake_automation.js
    4. UPDATE the 'FORM_URL' variable to this NEW link:
       ${newFormUrl}
    5. Save the script.
    6. Refresh the spreadsheet.
    7. Click 'Intake Tools' > 'Install Triggers (One Time)'.
  `;
  
  ui.alert(msg);
  Logger.log(msg); // Log it so they can copy/paste the URL
}