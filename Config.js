/* ---------- GLOBAL CONFIG -------------------------------------- */
const SS = SpreadsheetApp.getActive();

const CONFIG = {
  // ---- Sheet names
  TELEPHONE_LOG:  'Telephone_Log_2025',
  WAITING_LIST:   'Waiting_List_2025',
  REGISTRY:       'Patient_Registry',
  TEMPLATE:       'Intake_Template',
  HISTORY:        'Email_History_2025',
  FORM_RESPONSES: 'Form Responses',

  // ---- Google Form
  FORM_URL: 'https://docs.google.com/forms/d/e/1FAIpQLScfdkxGN5ZAGittteFpYRn1y2_nCvi0fgxgEZfMCsyHlKl5bw/viewform?usp=pp_url',

  // ---- Patient Registry column indexes (1-based)
  REG_COL_UID:                1, // Column A
  REG_COL_PATIENT_NAME:       2, // Column B
  REG_COL_EMAIL:              3, // Column C
  REG_COL_HAS_INTAKE:         5, // Column E

  // ---- Email history column indexes (1-based)
  HISTORY_COL_UID          : 1, // A
  HISTORY_COL_FORM         : 2, // B
  HISTORY_COL_DATE         : 3, // C
  HISTORY_COL_EMAIL        : 4, // D
  HISTORY_COL_SENT         : 5, // E
  HISTORY_COL_SUBMITTED    : 6,  // F

  // ---- Waiting List column indexes (1-based)
  WL_HEADER_ROWS:             2, 
  WL_COL_UID:                 1, 
  WL_COL_ACTIVE:              2, 
  WL_COL_SPOT_FOUND:          3, 
  WL_COL_RESPONSIBLE:         4, 
  WL_COL_PATIENT_NAME:        5, 
  WL_COL_DIAGNOSIS_NOTES:     6, 
  WL_COL_AGE:                 7, 
  WL_COL_PHONE:               8, 
  WL_COL_EMAIL:               9, 
  WL_COL_POTENTIAL_SERVICE:   12, 
  WL_SOCIAL_STRIDES:          13, 
  WL_NOT_INTERESTED:          15, 
  WL_INTAKE_COMPLETED:        16, 

  // ---- Telephone-Log column indexes (1-based)
  TL_HEADER_ROWS:            2, 
  TL_COL_UID:                1, 
  TL_COL_CALL_OUTCOME:       2, 
  TL_COL_DATE:               3, 
  TL_COL_FORM_SUBMITTED:     4, 
  TL_COL_FORM_SENT:          5, 
  TL_COL_RESPONSIBLE:        6, 
  TL_COL_PATIENT_NAME:       7, 
  TL_COL_PHONE:              8, 
  TL_COL_EMAIL:              9, 
  TL_COL_INFORMATION:       10, 
  TL_COL_TELEVISIT_SCHED:   11, 
  TL_COL_TELEVISIT_COMPLETE:12, 
  TL_COL_CONTACT_METHOD:    13, 
  TL_COL_WAITLIST_FLAG:     14, 
  TL_COL_ONSCHEDULE:        15, 
  TL_COL_NOT_INTERESTED:    16, 

  // ---- TL placeholders
  DEFAULT_TL_CALL_OUTCOME_PLACEHOLDER:'📝 FILL ME:',
  DEFAULT_TL_DISABLE_FORM_NOTE:      '<- Click here to disable automatic sending',
  DEFAULT_TL_EMAIL_INFO_NOTE:        'Adding an email in the email column will send a Google Form automatically',

  // ---- Intake Sheet Cell References (absolute A1 notation) ------------
  INTAKE_CELL_UID:                'A1',
  INTAKE_CELL_DATE:               'B2',
  INTAKE_CELL_RESPONSIBLE_PARTY:  'B3',
  INTAKE_CELL_PATIENT_NAME:       'B4',
  INTAKE_CELL_PHONE:              'B7',
  INTAKE_CELL_EMAIL:              'B8',
  INTAKE_CELL_DOB:                'B9',
  INTAKE_CELL_DIAGNOSIS_NOTES:    'B10',
  INTAKE_CELL_MED_HISTORY:        'B12',
  INTAKE_CELL_CLASSROOM:          'B18',
  INTAKE_CELL_THERAPIES:          'B19',
  INTAKE_CELL_FUNCTION_LEVEL:     'B15',
  INTAKE_CELL_GOALS:              'B39',
  INTAKE_CELL_ADDL_INFO:          'B29',
  INTAKE_CELL_BEST_CONTACT:       'B30',
  INTAKE_CELL_TELEHEALTH_DATE:    'B32',
  INTAKE_CELL_TELEHEALTH_TIME:    'C32',
  INTAKE_CELL_POTENTIAL_SERVICE:  'B33',

  // status / flags
  INTAKE_CELL_RECREATIONAL:       'D1',
  INTAKE_CELL_CALL_COMPLETED:     'D2',
  INTAKE_CELL_SPOT_FOUND:         'D3',
  INTAKE_CELL_ACTIVE:             'D4',
  INTAKE_CELL_NOT_INTERESTED:     'D5',
  INTAKE_CELL_TELEHEALTH_LINK:    'D32',
  INTAKE_CELL_FINANCIAL_AID:      'D36',

  // “Notes” cells
  INTAKE_TAB_WL_NOTE:             'E2',
  INTAKE_TAB_FORMS_SENT_NOTE:     'E4',
  INTAKE_TAB_TELEHEALTH_LINK_NOTE: 'E32',
  INTAKE_TAB_FINANCIAL_AID_NOTE:  'E36',

  // ---- Form-Responses column indexes (1-based) ------------------------
  FR_COL_TIMESTAMP                :  1,
  FR_COL_RESPONSIBLE_PARTY        :  2,
  FR_COL_UID                      :  3,
  FR_COL_EMAIL                    :  4,
  FR_COL_PHONE                    :  5,
  FR_COL_DOB                      :  6,
  FR_COL_PATIENT_TYPE             :  7,
  FR_COL_INTEREST_CHILD           :  8,
  FR_COL_CHILD_DIAGNOSIS          :  9,
  FR_COL_CHILD_MED_HISTORY        : 10,
  FR_COL_CHILD_CLASSROOM          : 11,
  FR_COL_CHILD_THERAPY_SCHOOL     : 12,
  FR_COL_CHILD_THERAPY_OUTPATIENT : 13,
  FR_COL_CHILD_FUNCTION_LEVEL     : 14,
  FR_COL_CHILD_GOALS              : 15,
  FR_COL_CHILD_ADDL_INFO          : 16,
  FR_COL_CHILD_BEST_CONTACT       : 17,
  FR_COL_INTEREST_ADULT           : 18,
  FR_COL_ADULT_DIAGNOSIS          : 19,
  FR_COL_ADULT_MED_HISTORY        : 20,
  FR_COL_ADULT_THERAPY_OUTPATIENT : 21,
  FR_COL_ADULT_FUNCTION_LEVEL     : 22,
  FR_COL_ADULT_GOALS              : 23,
  FR_COL_ADULT_ADDL_INFO          : 24,
  FR_COL_ADULT_BEST_CONTACT       : 25, 
};

/* ---------- TAB-COLOUR CONSTANTS -------------------------------- */
const COLORS = {
  GREEN:        '#00B050', // Active
  LIGHT_GREEN:  '#92D050', // SpotFound
  LIGHT_YELLOW: '#f1c232', // default
  RED         : '#EA4335'  // Not-interested
};

// Map of intake sheet cells to parameter names for sync
const INTAKE_SHEET_RELEVANT_CELLS_FOR_ORIGINAL_SYNC = {
  [CONFIG.INTAKE_CELL_RESPONSIBLE_PARTY]: 'responsibleParty',
  [CONFIG.INTAKE_CELL_PATIENT_NAME]:      'patientName',
  [CONFIG.INTAKE_CELL_PHONE]:             'phone',
  [CONFIG.INTAKE_CELL_EMAIL]:             'email',
  [CONFIG.INTAKE_CELL_DOB]:               'dob',
  [CONFIG.INTAKE_CELL_DIAGNOSIS_NOTES]:   'diagnosisNotes',
  [CONFIG.INTAKE_CELL_POTENTIAL_SERVICE]: 'potentialService', 
  [CONFIG.INTAKE_CELL_SPOT_FOUND]:        'spotFound',        
  [CONFIG.INTAKE_CELL_ACTIVE]:            'active'            
};

const SYSTEM_SHEET_NAMES = [
  CONFIG.TELEPHONE_LOG,
  CONFIG.WAITING_LIST,
  CONFIG.REGISTRY,
  CONFIG.TEMPLATE,
  CONFIG.HISTORY
];
