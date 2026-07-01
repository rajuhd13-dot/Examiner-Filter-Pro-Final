/**
 * Google Apps Script Backend for Examiner Filter Pro (v3.1 - Hyper-Optimized with Training Date)
 * Paste this code into your Google Apps Script editor (Extensions > Apps Script).
 * 
 * 1. Replace the SPREADSHEET_ID below with your actual ID.
 * 2. Click "Deploy" > "New Deployment" > "Web App".
 * 3. Set "Execute as" to "Me" and "Who has access" to "Anyone".
 * 4. Copy the Web App URL and paste it into the "GAS_DEPLOYMENT_URL" secret in AI Studio / Vercel.
 */

/*************** CONFIG ***************/
let SPREADSHEET_ID = '1R_O4llA1K43Y97GAgkK97WMvWbqg-tftz_FXpcUSZPU';
let SHEET_NAME     = 'Examiner Information';

// Security: Use dynamic config if passed from Express server/proxy
function setConfig_(e, content) {
  const qId = e?.parameter?.ssId || content?.ssId;
  const qSh = e?.parameter?.sheetName || content?.sheetName;
  if (qId) SPREADSHEET_ID = qId;
  if (qSh) SHEET_NAME = qSh;
}

const ALLOW = {
  ENGLISH: 55,
  BANGLA: 48,
  PHYSICS: 48,
  CHEMISTRY: 48,
  MATH: 48,
  BIOLOGY: 48,
  ICT: 48
};

const BLANK_LABEL = '(Blank)';
const BLANK_KEY   = '__blank__';

const ALL_SUBJECT_KEYS = ['english','bangla','physics','chemistry','math','biology','ict'];

const COL = {
  SL:          0,   // A
  NAME:        1,   // B (Nick Name)
  STATUS:      2,   // C
  TPIN:        3,   // D
  INST:        4,   // E
  DEPT:        5,   // F
  BATCH:       6,   // G (HSC Batch)
  RM:          7,   // H
  REMARKED_BY: 8,   // I
  MOB1:        9,   // J (Mobile Number)
  ALT:         10,  // K (Alternate)
  NAGAD:       11,  // L (Mobile Banking / Nagad)
  
  EN:          61,  // BJ English(%)
  BN:          64,  // BM Bangla(%)
  PHY:         67,  // BP Physics(%)
  CHEM:        70,  // BS Chemistry(%)
  MATH:        73,  // BV Math(%)
  BIO:         76,  // BY Biology(%)
  ICT:         79,  // CB ICT(%)
  
  TRAIN:            82,  // CE Training Report
  TRAIN_DATE:       83,  // CF Training Date
  FORM_CAMPUS:      84,  // CG Form Fill Up Car
  ID_CHECKED:       85,  // CH ID Checked
  ENTRY_BY:         86,  // CI Entry By
  FORM_FILLUP_DATE: 87,  // CJ Form Fillup Date
  CAMPUS:           88,  // CK Campus
  REMARK_RAW:       92   // CQ Remark
};

/*************** MEMORY CACHE ***************/
// Separate TTLs: options cache longer, data cache shorter for nearly live experience
let MEM_STORE = {
  loadedAt:     0,
  dataTtlMs:    20000,   // 20s for row data
  optTtlMs:     60000,   // 60s for dropdown options
  lastRowCount: 0,
  header:       null,
  body:         null,
  options:      null,
  optLoadedAt:  0,
  // Pre-built filter indexes for O(1) filtering
  instIdx:      null,   // Map<normalizedKey, Set<rowIdx>>
  deptIdx:      null,
  batchIdx:     null,
  trainIdx:     null,
  trainDateIdx: null,
  campusIdx:    null,
  tpinIdx:      null
};

/*************** ENTRY POINTS ***************/
function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  try {
    const action = (e && e.parameter) ? e.parameter.action : null;
    let content = {};
    
    if (e && e.postData && e.postData.contents) {
      try { content = JSON.parse(e.postData.contents); } catch (err) {}
    }

    setConfig_(e, content);

    const finalAction = action || content.action;
    
    if (finalAction === 'ping') {
      return respond_({ success: true, pong: true, version: "3.1.0", time: new Date().toISOString() });
    }

    const filters  = content.filters  || (e.parameter.filters ? JSON.parse(e.parameter.filters) : null);
    const page     = content.page     || e.parameter.page;
    const pageSize = content.pageSize || e.parameter.pageSize;
    const query    = content.query    || e.parameter.query;

    let result;

    if (finalAction === 'options' || finalAction === 'filterOptions') {
      result = getFilterOptionsFast();
    } else if (finalAction === 'filter') {
      result = getFilteredDataFast(filters, page, pageSize);
    } else if (finalAction === 'lookup') {
      result = lookupByQuery(query);
    } else if (finalAction === 'sync') {
      result = getSheetRowCount();
    } else if (finalAction === 'clearCache') {
      result = clearFastCache();
    } else {
      if (!finalAction) return respond_({ success: true, message: "GAS Ready" });
      result = { success: false, error: 'Unknown action: ' + finalAction };
    }
    
    return respond_(result);
  } catch (err) {
    return respond_({ success: false, error: err.toString(), hint: "Verify Spreadsheet ID." });
  }
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/*************** HELPERS ***************/
function normalize(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function isBlankish_(v) {
  if (v === null || v === undefined) return true;
  const s = String(v).replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
  return s === '';
}

function openSheetStrict_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    const names = ss.getSheets().map(s => s.getName()).join(' | ');
    throw new Error(`Sheet "${SHEET_NAME}" not found. Available: ${names}`);
  }
  return sh;
}

function toNum_(v) {
  const s = String(v ?? '').trim();
  if (!s) return NaN;
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

function cleanMobile_(v) {
  let m = String(v ?? '').trim().replace(/\D/g, '');
  if (m.length === 10) m = '0' + m;
  if (m.length === 11 && !m.startsWith('0')) m = '0' + m.slice(-10);
  return m;
}

function subjectThresholdDynamic_(k, allowEnglish, allowOthers) {
  const en  = Number.isFinite(allowEnglish) ? allowEnglish : null;
  const oth = Number.isFinite(allowOthers)  ? allowOthers  : null;
  if (k === 'english') return en !== null ? en : ALLOW.ENGLISH;
  if (['bangla','physics','chemistry','math','biology','ict'].includes(k)) {
    if (oth !== null) return oth;
    switch (k) {
      case 'bangla':    return ALLOW.BANGLA;
      case 'physics':   return ALLOW.PHYSICS;
      case 'chemistry': return ALLOW.CHEMISTRY;
      case 'math':      return ALLOW.MATH;
      case 'biology':   return ALLOW.BIOLOGY;
      case 'ict':       return ALLOW.ICT;
    }
  }
  return null;
}

function subjectValue_(r, k) {
  switch (k) {
    case 'english':   return toNum_(r[COL.EN]);
    case 'bangla':    return toNum_(r[COL.BN]);
    case 'physics':   return toNum_(r[COL.PHY]);
    case 'chemistry': return toNum_(r[COL.CHEM]);
    case 'math':      return toNum_(r[COL.MATH]);
    case 'biology':   return toNum_(r[COL.BIO]);
    case 'ict':       return toNum_(r[COL.ICT]);
    default:          return NaN;
  }
}

function isAllowedBySubjectsDynamic_(r, subjectsSelected, subjectLogic, allowEnglish, allowOthers) {
  let keys = Array.isArray(subjectsSelected) ? subjectsSelected.filter(Boolean) : [];
  if (keys.length === 0) keys = ALL_SUBJECT_KEYS;
  const mode = (subjectLogic === 'all') ? 'all' : 'any';
  
  if (mode === 'all') {
    for (let i = 0; i < keys.length; i++) {
      const k  = keys[i];
      const th = subjectThresholdDynamic_(k, allowEnglish, allowOthers);
      const v  = subjectValue_(r, k);
      if (!Number.isFinite(v) || !Number.isFinite(th) || v < th) return false;
    }
    return true;
  }
  for (let i = 0; i < keys.length; i++) {
    const k  = keys[i];
    const th = subjectThresholdDynamic_(k, allowEnglish, allowOthers);
    const v  = subjectValue_(r, k);
    if (!Number.isFinite(v) || !Number.isFinite(th)) continue;
    if (v >= th) return true;
  }
  return false;
}

/*************** INDEX BUILDER ***************/
// Build inverted indexes once on data load for fast O(1) multi-select filtering
function buildIndexes_(body) {
  const instIdx      = new Map();
  const deptIdx      = new Map();
  const batchIdx     = new Map();
  const trainIdx     = new Map();
  const trainDateIdx = new Map();
  const campusIdx    = new Map();
  const tpinIdx      = new Map();

  function addIdx(map, key, i) {
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(i);
  }

  const timeZone = Session.getScriptTimeZone();

  for (let i = 0; i < body.length; i++) {
    const r = body[i];
    const inst   = normalize(r[COL.INST]   ?? '');
    const dept   = normalize(r[COL.DEPT]   ?? '');
    const batch  = normalize(r[COL.BATCH]  ?? '');
    const train  = isBlankish_(r[COL.TRAIN]  ?? '') ? BLANK_KEY : normalize(r[COL.TRAIN]  ?? '');
    const campus = isBlankish_(r[COL.CAMPUS] ?? '') ? BLANK_KEY : normalize(r[COL.CAMPUS] ?? '');
    const tpin   = isBlankish_(r[COL.TPIN]   ?? '') ? BLANK_KEY : normalize(r[COL.TPIN]   ?? '');

    let trainDate = r[COL.TRAIN_DATE];
    if (trainDate instanceof Date) {
      trainDate = Utilities.formatDate(trainDate, timeZone, "yyyy-MM-dd");
    }
    const trainDateStr = isBlankish_(trainDate ?? '') ? BLANK_KEY : normalize(trainDate ?? '');

    if (inst)  addIdx(instIdx,  inst,  i);
    if (dept)  addIdx(deptIdx,  dept,  i);
    if (batch) addIdx(batchIdx, batch, i);
    addIdx(trainIdx,     train,        i);
    addIdx(trainDateIdx, trainDateStr, i);
    addIdx(campusIdx,    campus,       i);
    addIdx(tpinIdx,      tpin,         i);
  }

  return { instIdx, deptIdx, batchIdx, trainIdx, trainDateIdx, campusIdx, tpinIdx };
}

/*************** RESULT COLUMNS ***************/
function buildKeepColsAndHeader_(subjectKeys) {
  const keepCols = [COL.SL, COL.NAME, COL.TPIN, COL.INST, COL.DEPT, COL.BATCH, COL.RM, COL.MOB1, COL.ALT];
  const header   = ['SL', 'Nick Name', 'T-PIN', 'Inst.', 'Dept.', 'HSC Batch', 'Rm', 'Mobile Number', 'Alternate'];

  const SUBJECT_MAP = {
    english:   { col: COL.EN,   label: 'English(%)'   },
    bangla:    { col: COL.BN,   label: 'Bangla(%)'    },
    physics:   { col: COL.PHY,  label: 'Physics(%)'   },
    chemistry: { col: COL.CHEM, label: 'Chemistry(%)'  },
    math:      { col: COL.MATH, label: 'Math(%)'      },
    biology:   { col: COL.BIO,  label: 'Biology(%)'   },
    ict:       { col: COL.ICT,  label: 'ICT(%)'       }
  };

  for (let i = 0; i < subjectKeys.length; i++) {
    const map = SUBJECT_MAP[subjectKeys[i]];
    if (map) { keepCols.push(map.col); header.push(map.label); }
  }

  keepCols.push(COL.TRAIN, COL.TRAIN_DATE, COL.CAMPUS, -1);
  header.push('Training Report', 'Training Date', 'Physical Campus', 'Allow Status');
  return { keepCols, header };
}

/*************** SHEET STORE — SMART CACHE ***************/
function getSheetStore_() {
  const now = Date.now();
  const sh  = openSheetStrict_();
  const currentRowCount = sh.getLastRow();

  const cacheValid =
    MEM_STORE.body &&
    MEM_STORE.header &&
    (now - MEM_STORE.loadedAt) < MEM_STORE.dataTtlMs &&
    MEM_STORE.lastRowCount === currentRowCount;

  if (cacheValid) {
    return { header: MEM_STORE.header, body: MEM_STORE.body };
  }

  const lastCol = sh.getLastColumn();

  if (currentRowCount < 2) {
    MEM_STORE = { ...MEM_STORE, header: [], body: [], options: null, optLoadedAt: 0, loadedAt: now, lastRowCount: currentRowCount, instIdx: null, deptIdx: null, batchIdx: null, trainIdx: null, trainDateIdx: null, campusIdx: null, tpinIdx: null };
    return { header: [], body: [] };
  }

  const values = sh.getRange(1, 1, currentRowCount, lastCol).getValues();
  const header = values[0];
  const body0  = values.slice(1);

  const body = new Array(body0.length);
  for (let i = 0; i < body0.length; i++) {
    const r = body0[i].slice();
    r[COL.MOB1] = cleanMobile_(r[COL.MOB1]);
    r[COL.ALT]  = cleanMobile_(r[COL.ALT]);
    body[i] = r;
  }

  // Build indexes
  const { instIdx, deptIdx, batchIdx, trainIdx, trainDateIdx, campusIdx, tpinIdx } = buildIndexes_(body);

  MEM_STORE.header       = header;
  MEM_STORE.body         = body;
  MEM_STORE.options      = null; // invalidate options too
  MEM_STORE.optLoadedAt  = 0;
  MEM_STORE.loadedAt     = now;
  MEM_STORE.lastRowCount = currentRowCount;
  MEM_STORE.instIdx      = instIdx;
  MEM_STORE.deptIdx      = deptIdx;
  MEM_STORE.batchIdx     = batchIdx;
  MEM_STORE.trainIdx     = trainIdx;
  MEM_STORE.trainDateIdx = trainDateIdx;
  MEM_STORE.campusIdx    = campusIdx;
  MEM_STORE.tpinIdx      = tpinIdx;

  return { header, body };
}

/*************** OPTIONS ***************/
function getFilterOptionsFast() {
  try {
    const now = Date.now();
    // Options have their own longer TTL
    if (MEM_STORE.options && (now - MEM_STORE.optLoadedAt) < MEM_STORE.optTtlMs) {
      return MEM_STORE.options;
    }

    const store = getSheetStore_();
    const body  = store.body;

    const institutes    = new Set();
    const departments   = new Set();
    const batches       = new Set();
    const trainings     = new Set();
    const trainingDates = new Set();
    const campuses      = new Set();
    const tpins         = new Set();

    const timeZone = Session.getScriptTimeZone();

    for (let i = 0; i < body.length; i++) {
      const r    = body[i];
      const inst = String(r[COL.INST]   ?? '').trim();
      const dept = String(r[COL.DEPT]   ?? '').trim();
      const bat  = String(r[COL.BATCH]  ?? '').trim();
      const trn  = String(r[COL.TRAIN]  ?? '').trim();
      const cam  = String(r[COL.CAMPUS] ?? '').trim();
      const tpin = String(r[COL.TPIN]   ?? '').trim();

      let tDate = r[COL.TRAIN_DATE];
      if (tDate instanceof Date) {
        tDate = Utilities.formatDate(tDate, timeZone, "yyyy-MM-dd");
      }
      const trnDate = String(tDate ?? '').trim();

      if (inst) institutes.add(inst);
      if (dept) departments.add(dept);
      if (bat)  batches.add(bat);
      trainings.add(trn  ? trn  : BLANK_LABEL);
      trainingDates.add(trnDate ? trnDate : BLANK_LABEL);
      campuses.add(cam   ? cam  : BLANK_LABEL);
      tpins.add(tpin     ? tpin : BLANK_LABEL);
    }

    const subjects = [
      { key: 'english',   label: 'English(%)'   },
      { key: 'bangla',    label: 'Bangla(%)'    },
      { key: 'physics',   label: 'Physics(%)'   },
      { key: 'chemistry', label: 'Chemistry(%)'  },
      { key: 'math',      label: 'Math(%)'      },
      { key: 'biology',   label: 'Biology(%)'   },
      { key: 'ict',       label: 'ICT(%)'       }
    ];

    const out = {
      success:       true,
      rowCount:      body.length,
      institutes:    [...institutes  ].sort((a, b) => a.localeCompare(b, 'en-US')),
      departments:   [...departments ].sort((a, b) => a.localeCompare(b, 'en-US')),
      batches:       [...batches     ].sort((a, b) => a.localeCompare(b, 'en-US', { numeric: true })),
      trainings:     [...trainings   ].sort((a, b) => a.localeCompare(b, 'en-US', { numeric: true })),
      trainingDates: [...trainingDates].sort((a, b) => a.localeCompare(b, 'en-US', { numeric: true })),
      campuses:      [...campuses    ].sort((a, b) => a.localeCompare(b, 'en-US', { numeric: true })),
      tpins:         [...tpins       ].sort((a, b) => a.localeCompare(b, 'en-US', { numeric: true })),
      subjects,
      allow: ALLOW
    };

    MEM_STORE.options     = out;
    MEM_STORE.optLoadedAt = now;
    return out;
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/*************** ROW COUNT CHECK ***************/
function getSheetRowCount() {
  try {
    const sh = openSheetStrict_();
    return { success: true, rowCount: sh.getLastRow() - 1 };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/*************** INDEX-BASED FAST ROW RESOLUTION ***************/
// Returns a candidate Set of row indices based on index lookups
// Much faster than iterating all rows when filters are applied
function getCandidateRows_(body, nf) {
  const { instIdx, deptIdx, batchIdx, trainIdx, trainDateIdx, campusIdx, tpinIdx } = MEM_STORE;
  let candidates = null; // null = all rows

  function intersect(map, selectedArr) {
    if (!selectedArr || selectedArr.length === 0) return candidates;
    const unionSet = new Set();
    for (let i = 0; i < selectedArr.length; i++) {
      const key = selectedArr[i] === BLANK_KEY ? BLANK_KEY : normalize(selectedArr[i]);
      const rows = map ? map.get(key) : null;
      if (rows) { for (let j = 0; j < rows.length; j++) unionSet.add(rows[j]); }
    }
    if (candidates === null) return unionSet;
    // Intersect existing candidates with new union
    const result = new Set();
    for (const idx of unionSet) { if (candidates.has(idx)) result.add(idx); }
    return result;
  }

  if (nf.institute.length  > 0 && instIdx)   candidates = intersect(instIdx,   nf.institute);
  if (nf.department.length > 0 && deptIdx)   candidates = intersect(deptIdx,   nf.department);
  if (nf.batch.length      > 0 && batchIdx)  candidates = intersect(batchIdx,  nf.batch);
  if (nf.trainingsSelected.length > 0 && trainIdx)  candidates = intersect(trainIdx,  nf.trainingsSelected);
  if (nf.trainingDatesSelected.length > 0 && trainDateIdx) candidates = intersect(trainDateIdx, nf.trainingDatesSelected);
  if (nf.campusesSelected.length  > 0 && campusIdx) candidates = intersect(campusIdx, nf.campusesSelected);
  if (nf.tpinsSelected.length     > 0 && tpinIdx)   candidates = intersect(tpinIdx,   nf.tpinsSelected);

  return candidates; // null means no index filter applied → iterate all
}

/*************** FILTER CORE ***************/
function buildNF_(filters) {
  return {
    institute:             (filters?.institute         || []).map(normalize).filter(Boolean),
    department:            (filters?.department        || []).map(normalize).filter(Boolean),
    batch:                 (filters?.batch             || []).map(normalize).filter(Boolean),
    trainingsSelected:     processMultiBlankAware_(filters?.trainingsSelected || []),
    trainingDatesSelected: processMultiBlankAware_(filters?.trainingDatesSelected || []),
    campusesSelected:      processMultiBlankAware_(filters?.campusesSelected  || []),
    tpinsSelected:         processMultiBlankAware_(filters?.tpinsSelected     || []),
    subjectsSelected:      Array.isArray(filters?.subjectsSelected)
                             ? filters.subjectsSelected.map(s => String(s || '').toLowerCase()).filter(Boolean)
                             : [],
    onlyAllowed:  (filters?.onlyAllowed !== false),
    subjectLogic: (filters?.subjectLogic === 'all') ? 'all' : 'any'
  };
}

function processMultiBlankAware_(arr) {
  if (!Array.isArray(arr)) return [];
  return arr
    .map(x => {
      const s = String(x ?? '');
      if (s === BLANK_LABEL || s === BLANK_KEY) return BLANK_KEY;
      const n = normalize(s);
      return n || null;
    })
    .filter(Boolean);
}

function filterCore_(filters, page, pageSize, returnAllRows) {
  const store = getSheetStore_();
  const body  = store.body;

  const allowEnglish = toNum_(filters?.allowEnglish);
  const allowOthers  = toNum_(filters?.allowOthers);
  const nf = buildNF_(filters);

  let subjectKeys = Array.isArray(nf.subjectsSelected) ? nf.subjectsSelected.filter(Boolean) : [];
  if (subjectKeys.length === 0) subjectKeys = ALL_SUBJECT_KEYS;

  const { keepCols, header: slimHeader } = buildKeepColsAndHeader_(subjectKeys);

  const ps         = pageSize || 200;
  const startIndex = Math.max(0, ((page || 1) - 1) * ps);
  const endIndex   = startIndex + ps;

  // Get candidate rows via index (fast path)
  const candidateSet = getCandidateRows_(body, nf);

  const rows = [];
  let totalMatched = 0;

  // Iterate only candidate rows (or all if no index filter)
  const useAll = candidateSet === null;
  const iterLen = useAll ? body.length : 0;
  const iterSet = useAll ? null : candidateSet;

  const timeZone = Session.getScriptTimeZone();

  function processRow(i) {
    const src = body[i];
    const ok = isAllowedBySubjectsDynamic_(src, nf.subjectsSelected, nf.subjectLogic, allowEnglish, allowOthers);
    if (nf.onlyAllowed && !ok) return;

    totalMatched++;

    if (returnAllRows || (totalMatched > startIndex && totalMatched <= endIndex)) {
      const row = new Array(keepCols.length);
      for (let j = 0; j < keepCols.length; j++) {
        const c = keepCols[j];
        if (c === -1) {
          row[j] = ok ? 'ALLOWED' : 'NOT ALLOWED';
        } else if (c === COL.TRAIN_DATE) {
          let tDate = src[c];
          if (tDate instanceof Date) {
            tDate = Utilities.formatDate(tDate, timeZone, "yyyy-MM-dd");
          }
          row[j] = tDate ? String(tDate).trim() : '';
        } else {
          row[j] = src[c] !== undefined && src[c] !== null ? String(src[c]).trim() : '';
        }
      }
      rows.push(row);
    }
  }

  if (useAll) {
    for (let i = 0; i < body.length; i++) processRow(i);
  } else {
    for (const i of iterSet) processRow(i);
  }

  return {
    success:    true,
    header:     slimHeader,
    rows,
    total:      totalMatched,
    page:       page || 1,
    pageSize:   ps,
    totalPages: Math.max(1, Math.ceil(totalMatched / ps)),
    allow: {
      ENGLISH: Number.isFinite(allowEnglish) ? allowEnglish : ALLOW.ENGLISH,
      OTHERS:  Number.isFinite(allowOthers)  ? allowOthers  : 48
    }
  };
}

/*************** PUBLIC ***************/
function getFilteredDataFast(filters, page, pageSize) {
  try {
    return filterCore_(filters, page || 1, pageSize || 200, false);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getAllFilteredDataForExport(filters) {
  try {
    return filterCore_(filters, 1, 999999999, true);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function lookupByQuery(query) {
  try {
    const q = normalize(query);
    if (!q) return { success: true, found: false };
    const { body } = getSheetStore_();
    
    const timeZone = Session.getScriptTimeZone();

    for (const r of body) {
      if (normalize(r[COL.TPIN]) === q || normalize(r[COL.MOB1]).slice(-10) === q.slice(-10) || normalize(r[COL.ALT]).slice(-10) === q.slice(-10)) {
        const { keepCols, header } = buildKeepColsAndHeader_(ALL_SUBJECT_KEYS);
        
        const row = keepCols.map(c => {
          if (c === -1) {
            return 'LOOKUP';
          } else if (c === COL.TRAIN_DATE) {
            let tDate = r[c];
            if (tDate instanceof Date) {
              tDate = Utilities.formatDate(tDate, timeZone, "yyyy-MM-dd");
            }
            return tDate ? String(tDate).trim() : '';
          } else {
            return r[c] !== undefined && r[c] !== null ? String(r[c]).trim() : '';
          }
        });

        return { success: true, found: true, header, row };
      }
    }
    return { success: true, found: false };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function clearFastCache() {
  MEM_STORE = {
    loadedAt:     0,
    dataTtlMs:    20000,
    optTtlMs:     60000,
    lastRowCount: 0,
    header:       null,
    body:         null,
    options:      null,
    optLoadedAt:  0,
    instIdx:      null,
    deptIdx:      null,
    batchIdx:     null,
    trainIdx:     null,
    trainDateIdx: null,
    campusIdx:    null,
    tpinIdx:      null
  };
  return { success: true };
}

function testSheetAccess() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  Logger.log(sh.getName() + ' | rows: ' + sh.getLastRow() + ' | cols: ' + sh.getLastColumn());
}
