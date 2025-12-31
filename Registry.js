/* ================================================================
 * UID REGISTRY HELPERS
 * ================================================================ */

function markHasIntake_(uid, value = true) {
  const reg = sheet_(CONFIG.REGISTRY);
  const uidCol = reg.getRange(1, CONFIG.REG_COL_UID, reg.getLastRow(), 1).getValues().flat();
  const idx = uidCol.indexOf(uid);
  if (idx === -1) return;
  const row = idx + 1;
  reg.getRange(row, CONFIG.REG_COL_HAS_INTAKE).setValue(value);
}

function getOrCreateUID_(patient, responsible, email) {
  Logger.log(`getOrCreateUID_ v2: patient="${patient}", email="${email}"`);

  // 1) look-up logic (try to reuse an old UID)
  const reg = sheet_(CONFIG.REGISTRY);
  const rows = reg.getDataRange().getValues();

  // Fast path – valid e-mail was supplied
  if (validateEmail(email)) {
    // a) exact email + exact name
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][CONFIG.REG_COL_EMAIL - 1] === email &&
          rows[i][CONFIG.REG_COL_PATIENT_NAME - 1] === patient) {
        return rows[i][CONFIG.REG_COL_UID - 1];
      }
    }
    // b) exact email + *similar* name
    for (let r of rows) {
      if (r[CONFIG.REG_COL_EMAIL - 1] === email &&
          nameSimilarity(r[CONFIG.REG_COL_PATIENT_NAME - 1], patient) >= 0.90) {
        return r[CONFIG.REG_COL_UID - 1];
      }
    }
  }

  // 2) could not reuse → make a brand-new UID
  const uid = newUid_({patient, email});
  return uid;
}

function ensureRegistryEntry_(uid, patient, email) {
  const reg = sheet_(CONFIG.REGISTRY);
  const uidCol = reg.getRange(1, CONFIG.REG_COL_UID, reg.getLastRow(), 1).getValues().flat();
  let idx = uidCol.indexOf(uid);

  if (idx === -1) { // totally new UID → append row
    reg.appendRow([
      uid,
      patient || '',
      validateEmail(email) ? email : '',
      new Date(),
      false
    ]);
    return;
  }

  // UID already exists – improve the record if we now have better data
  idx += 1;
  if (patient && !reg.getRange(idx, CONFIG.REG_COL_PATIENT_NAME).getValue()) {
    reg.getRange(idx, CONFIG.REG_COL_PATIENT_NAME).setValue(patient);
  }
  if (validateEmail(email) && !reg.getRange(idx, CONFIG.REG_COL_EMAIL).getValue()) {
    reg.getRange(idx, CONFIG.REG_COL_EMAIL).setValue(email);
  }
}

function newUid_({ patient = '', email = '' } = {}) {
  const uid = Utilities.getUuid();
  ensureRegistryEntry_(uid, patient, email);
  return uid;
}