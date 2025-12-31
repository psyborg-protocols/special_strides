/* ================================================================
 * FORMS – Central Catalogue & Configuration
 * ================================================================ */

/* ─────────────────────────────
 * Re-usable HTML e-mail signature
 * ───────────────────────────── */
const EMAIL_SIGNATURE_HTML = `
  <br><br><span style="color:#808080">
  ––<br>
  <strong>Special&nbsp;Strides</strong><br>
  Office: 732-446-0945<br>
  Fax:&nbsp;&nbsp;&nbsp;&nbsp;732-446-5391<br>
  Email: information@specialstrides.com<br>
  <br>Improving&nbsp;Lives…One&nbsp;Stride&nbsp;at&nbsp;a&nbsp;Time
  </span>`;

/* ─────────────────────────────
 * 1. STATIC DEFINITIONS (Logic & Text)
 * URLs are now loaded dynamically from 'System_Form_Links'
 * ───────────────────────────── */
const FORM_DEFINITIONS = {

  /* A) Online Intake Form */
  INTAKE: {
    defaultName: 'Google Intake Form',
    responseSheet: 'Form Responses',
    mail: {
      subject: 'Special Strides – Online Intake Form',
      body: ({link, responsible = ''}) => `
        <p>Dear ${responsible || 'Friend'},</p>
        <p>Thank you for contacting <strong>Special&nbsp;Strides</strong>.
        Please complete our intake form before your first visit.</p>
        <p><a href="${link}" target="_blank">${link}</a></p>
        <p>Warm regards,<br>The Special Strides Team</p>${EMAIL_SIGNATURE_HTML}`
    }
  },

  /* B) Therapy Paperwork Pack */
  THERAPY_START: {
    defaultName: 'Therapy Paperwork Pack',
    responseSheet: null,
    mail: {
      subject: 'Please complete these forms for Therapy',
      body: ({responsible = '', link}) => `
        <p>Dear ${responsible || 'Client'},</p>
        <p>We are excited that you will be starting therapy soon.
        Kindly complete the necessary paperwork before your start date.
        If you have any questions please do not hesitate to contact us.</p>
        <p><a href="${link}" target="_blank">New Client Handbook&nbsp;– Therapy</a></p>
        <p>Warm regards,<br>The Special Strides Team</p>${EMAIL_SIGNATURE_HTML}`
    }
  },

  /* C) Recreational Paperwork Pack */
  RECREATIONAL_START: {
    defaultName: 'Recreational Paperwork Pack',
    responseSheet: null,
    mail: {
        subject: 'Please complete these forms for Recreational Riding',
        body: ({responsible = '', link}) => `
        <p>Dear ${responsible || 'Friend'},</p>
        <p>We are excited that you will be starting recreational riding soon.
        Kindly complete the attached paperwork pack before your first session.
        If you have any questions please let us know.</p>
        <p><a href="${link}" target="_blank">New Client Handbook</a></p>
        <p>Warm regards,<br>The Special Strides Team</p>${EMAIL_SIGNATURE_HTML}`
    }
  },

  /* D) Annual Financial-Aid Application */
  FINANCIAL_AID_2025: {
    defaultName: 'Financial Aid Application',
    responseSheet: 'FA 2025',
    mail: {
      subject: 'Financial Aid Application Link',
      body: ({responsible = '', link}) => `
        <p>Dear ${responsible || 'Family'},</p>
        <p>It was a pleasure meeting with you today. Please complete the
        Financial Aid Application using the link below.</p>
        <p><a href="${link}" target="_blank">Financial Aid Application</a></p>
        <p>If you have any questions please contact me at any time at
        732-446-0945.</p>${EMAIL_SIGNATURE_HTML}`
    }
  },

  /* E) Tele-health Appointment */
  TELEHEALTH_APPT: {
    defaultName: 'Tele-health Appointment',
    responseSheet: null,
    mail: {
      subject: 'Tele-health visit appointment link',
      body: ({responsible = '', apptDate = '', link}) => `
        <p>Dear ${responsible || 'Client'},</p>
        <p>It was a pleasure to speak with you today. Please use this link
        for our tele-health visit scheduled on <strong>${apptDate}</strong>.</p>
        <p><a href="${link}" target="_blank">${link}</a></p>
        <p>I look forward to speaking with you.</p>
        <p>Warm regards,<br>The Special Strides Team</p>${EMAIL_SIGNATURE_HTML}`
    }
  }
};

/* ─────────────────────────────
 * 2. DYNAMIC LOADER
 * ───────────────────────────── */
function getFormConfig_(key) {
  // 1. Get the static code definition (email body, etc.)
  const staticDef = FORM_DEFINITIONS[key];
  if (!staticDef) {
    Logger.log(`Error: Form key "${key}" not found in code definitions.`);
    return null;
  }

  // 2. Fetch live URLs from the hidden System sheet
  const linkData = getFormLinkDataFromSheet_(key);

  if (!linkData || !linkData.url) {
    Logger.log(`Error: URL for "${key}" not found in sheet "${CONFIG.FORM_LINKS}".`);
    // Fallback: If you haven't filled the sheet yet, this will fail safely.
    return null;
  }

  // 3. Merge them
  return {
    ...staticDef,
    formUrl: linkData.url,
    uidEntry: linkData.uidParam || null, // Overwrite if present in sheet
    displayName: linkData.name || staticDef.defaultName
  };
}

function getFormLinkDataFromSheet_(key) {
  const sheet = sheet_(CONFIG.FORM_LINKS);
  if (!sheet) return null;

  // Cache strategy: Read all links once to save time?
  // For simplicity and low volume, we can just find the row.
  const data = sheet.getDataRange().getValues();
  
  // Skip header (row 0)
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.LINKS_COL_KEY - 1] === key) {
      return {
        url:      data[i][CONFIG.LINKS_COL_URL - 1],
        uidParam: data[i][CONFIG.LINKS_COL_UID_PARAM - 1],
        name:     data[i][CONFIG.LINKS_COL_NAME - 1]
      };
    }
  }
  return null;
}