/* ================================================================
 *  FORMS – central catalogue for every automatic e-mail we send
 *          Add / edit objects only here – everything else is generic
 * ================================================================ */

/* ─────────────────────────────
 *  Re-usable HTML e-mail signature
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
 *  Form catalogue
 * ───────────────────────────── */
const FORMS = {

  /* =============================================================
   *  A)  Online Intake Form  (Google Form)
   * =========================================================== */
  INTAKE: {
    displayName: 'Google Intake Form',
    formUrl : 'https://docs.google.com/forms/d/e/1FAIpQLScfdkxGN5ZAGittteFpYRn1y2_nCvi0fgxgEZfMCsyHlKl5bw/viewform?usp=pp_url',
    uidEntry: 'entry.526811714',          // UID pre-fill parameter
    responseSheet: 'Form Responses',      // edit if your tab name differs
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

  /* =============================================================
   *  B)  Therapy Paperwork Pack  (Handbook + Forms)
   *     – PDF / shared-drive link, not a Google Form
   * =========================================================== */
  THERAPY_START: {
    displayName: 'Therapy Paperwork Pack',
    formUrl : 'https://specialstrides-my.sharepoint.com/:b:/g/personal/srehr_specialstrides_onmicrosoft_com/EQkzCPxTLXFErNeQOm9i1NoBemQObnTqQTwUO7-vNPj9Lg?e=cRjsVb',  // ↺ link to the pack
    uidEntry: null,                          // no UID injection
    responseSheet: null,                     // no on-submit tracking
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
    /* =============================================================
    *  C)  Recreational Paperwork Pack
    * =========================================================== */
    RECREATIONAL_START: {
    displayName: 'Recreational Paperwork Pack',
    formUrl : 'https://specialstrides-my.sharepoint.com/:b:/g/personal/srehr_specialstrides_onmicrosoft_com/ERsF08isSf5Pj7yo7q9z9pkBUOpd0kUx6rY8svqPg74T0Q?e=UZIWnb',
    uidEntry: null,                        // no UID injection
    responseSheet: null,                   // no on-submit tracking
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
  /* =============================================================
   *  D)  Annual Financial-Aid Application (Google Form)
   * =========================================================== */
  FINANCIAL_AID_2025: {
    displayName: '2025 Financial Aid Application',
    formUrl : 'https://forms.gle/yBEbWNKjiHuELKDp8',
    uidEntry: null,
    responseSheet: 'FA 2025',
    mail: {
      subject: '2025 Financial Aid Application Link',
      body: ({responsible = ''}) => `
        <p>Dear ${responsible || 'Family'},</p>
        <p>It was a pleasure meeting with you today. Please complete the
        2025 Financial Aid Application using the link below.</p>
        <p><a href="https://forms.gle/yBEbWNKjiHuELKDp8" target="_blank">
        2025 Financial Aid Application</a></p>
        <p>If you have any questions please contact me at any time at
        732-446-0945.</p>${EMAIL_SIGNATURE_HTML}`
    }
  },

  /* =============================================================
   *  E)  Tele-health Appointment (static doxy.me room link)
   * =========================================================== */
  TELEHEALTH_APPT: {
    displayName: 'Tele-health Appointment',
    formUrl : 'https://specialstrides.doxy.me/srehr',  // static room link
    uidEntry: null,
    responseSheet: null,          // no submission tracking
    mail: {
      subject: 'Tele-health visit appointment link',
      body: ({responsible = '', apptDate = ''}) => `
        <p>Dear ${responsible || 'Client'},</p>
        <p>It was a pleasure to speak with you today. Please use this link
        for our tele-health visit scheduled on <strong>${apptDate}</strong>.</p>
        <p><a href="https://specialstrides.doxy.me/srehr" target="_blank">
        https://specialstrides.doxy.me/srehr</a></p>
        <p>I look forward to speaking with you.</p>
        <p>Warm regards,<br>The Special Strides Team</p>${EMAIL_SIGNATURE_HTML}`
    }
  }

  /* ── Add more forms here in the same pattern ── */
};