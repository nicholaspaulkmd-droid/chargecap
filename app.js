// app.js — ChargeCap v1.0
// Vanilla JS, no framework. Uses idb-keyval (IndexedDB) for storage,
// Tesseract.js for on-device OCR, NLM ClinicalTables API for live ICD-10
// search, and the Google Identity Services + Sheets API for Drive sync.

// ---------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------

// Paste your Google OAuth 2.0 Web Client ID here after completing the
// Google Cloud Console steps in README.md. Leave blank to run with
// Drive sync disabled — every other feature works without it.
// This can also be set (and is persisted) from the in-app Settings panel,
// so you do not have to rebuild/redeploy just to add it later.
const DEFAULT_GOOGLE_CLIENT_ID = "";

const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file";
const SHEET_NAME = "ChargeCap Log";

// Tab layout inside the single "ChargeCap Log" spreadsheet:
//   - one tab per role (so primary/co-surgeon cases and assistant cases
//     never mix on the same page, per how case logs get counted)
//   - a hidden helper tab that unions both, so the tally only has to
//     read from one place
//   - the tally itself, entirely formula-driven so it stays live as new
//     rows land without the app having to recompute anything
const TAB_PRIMARY = "Primary & Co-Surgeon";
const TAB_ASSISTANT = "Assistant";
const TAB_ALL = "All Cases";
const TAB_TALLY = "Monthly Tally";
const DATA_TABS = [TAB_PRIMARY, TAB_ASSISTANT];
const ALL_TABS = [TAB_PRIMARY, TAB_ASSISTANT, TAB_ALL, TAB_TALLY];

const SHEET_HEADER = [
  "Date", "Patient", "MRN", "DOB", "Surgeon", "Role", "Modifier", "Facility",
  "Billing Status", "Billed Date", "CPT Codes", "CPT Descriptions",
  "ICD-10 Codes", "ICD-10 Descriptions", "Notes", "Category",
];

// Case category, used only for the monthly tally. A case counts as
// Bariatric or EGD if ANY of its CPT codes match those lists (Bariatric
// checked first, since a bariatric case that also includes an on-table
// EGD should still count as bariatric) — everything else falls through
// to General Surgery. Edit these two lists if the code sets change.
const BARIATRIC_CPT = new Set(["43633", "43860", "43644", "43659", "43775", "43845", "43774"]);
const EGD_CPT = new Set(["43266", "43235", "43239", "43245", "43247", "43233"]);

function caseCategory(c) {
  const codes = (c.cptCodes || []).map((x) => x.code);
  if (codes.some((code) => BARIATRIC_CPT.has(code))) return "Bariatric";
  if (codes.some((code) => EGD_CPT.has(code))) return "EGD";
  return "General Surgery";
}

function tabForRole(role) {
  return role === "assistant" ? TAB_ASSISTANT : TAB_PRIMARY;
}

// ---------------------------------------------------------------------
// Storage (IndexedDB via idb-keyval)
// ---------------------------------------------------------------------

// idb-keyval's createStore() opens the database itself, so calling it
// three times for three stores in the SAME database is a race: the
// browser only runs the "create the storage areas" step once per
// database, so only the first store to win that race actually gets
// created — the other two silently never exist, and any transaction
// against them throws NotFoundError. Fixed by opening the database
// ourselves, once, creating all three object stores together in a
// single upgrade — then wrapping each as an idb-keyval-compatible
// store function (same shape createStore() returns, so idbKeyval.get/
// set/entries all still work unchanged).
//
// DB_VERSION is bumped (from an implicit 1) so a browser that already
// has a broken "chargecap-db" from an earlier buggy deploy re-runs the
// upgrade and gets the missing stores created, instead of staying
// stuck in its broken state forever.
const DB_VERSION = 2;
const STORE_NAMES = ["cases", "meta", "codes"];

function openChargeCapDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open("chargecap-db", DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      STORE_NAMES.forEach((name) => {
        if (!db.objectStoreNames.contains(name)) db.createObjectStore(name);
      });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

const dbPromise = openChargeCapDB();

function makeStore(storeName) {
  return (txMode, callback) =>
    dbPromise.then((db) => callback(db.transaction(storeName, txMode).objectStore(storeName)));
}

const casesStore = makeStore("cases");
const metaStore = makeStore("meta");
const codesStore = makeStore("codes");

let CASES = []; // in-memory cache, source of truth is IndexedDB
let SYNC_QUEUE = [];

// Editable code libraries. Seeded once from data.js on first run, then
// live entirely in IndexedDB — editing/adding/deleting here never
// touches data.js, so changes persist without a redeploy.
let CPT_LIB = [];
let ICD10_LIB = [];

async function loadCodes() {
  let cpt = await idbKeyval.get("cpt", codesStore);
  let icd10 = await idbKeyval.get("icd10", codesStore);
  if (!cpt) {
    cpt = CPT_FAVORITES.map((c) => ({ ...c, id: uuid() }));
    await idbKeyval.set("cpt", cpt, codesStore);
  }
  if (!icd10) {
    icd10 = ICD10_FAVORITES.map((c) => ({ ...c, id: uuid() }));
    await idbKeyval.set("icd10", icd10, codesStore);
  }
  CPT_LIB = cpt;
  ICD10_LIB = icd10;
}

function codeLib(type) { return type === "cpt" ? CPT_LIB : ICD10_LIB; }

async function persistCodeLib(type) {
  await idbKeyval.set(type, codeLib(type), codesStore);
}

async function addCodeToLib(type, entry) {
  const lib = codeLib(type);
  const clean = { id: uuid(), code: entry.code.trim(), desc: entry.desc.trim(), category: (entry.category || "Custom").trim() || "Custom" };
  lib.unshift(clean);
  await persistCodeLib(type);
  return clean;
}

async function updateCodeInLib(type, id, patch) {
  const lib = codeLib(type);
  const item = lib.find((c) => c.id === id);
  if (!item) return;
  Object.assign(item, patch);
  await persistCodeLib(type);
}

async function deleteCodeFromLib(type, id) {
  const lib = codeLib(type);
  const idx = lib.findIndex((c) => c.id === id);
  if (idx >= 0) lib.splice(idx, 1);
  await persistCodeLib(type);
}

function isCodeInLib(type, code) {
  return codeLib(type).some((c) => c.code === code);
}

async function loadCases() {
  const all = await idbKeyval.entries(casesStore);
  CASES = all.map(([, v]) => v).sort((a, b) => b.createdAt - a.createdAt);
}

async function saveCase(c) {
  c.updatedAt = Date.now();
  await idbKeyval.set(c.id, c, casesStore);
  const idx = CASES.findIndex((x) => x.id === c.id);
  if (idx >= 0) CASES[idx] = c;
  else CASES.unshift(c);
}

async function getMeta(key, fallback) {
  const v = await idbKeyval.get(key, metaStore);
  return v === undefined ? fallback : v;
}
async function setMeta(key, val) {
  await idbKeyval.set(key, val, metaStore);
}

// ---------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------

function uuid() {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

function todayISO() {
  const d = new Date();
  return d.toISOString().slice(0, 10);
}

function fmtDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  if (!y) return iso;
  return `${m}/${d}/${y}`;
}

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function csvEscape(s) {
  const str = String(s ?? "");
  if (/[",\n]/.test(str)) return `"${str.replace(/"/g, '""')}"`;
  return str;
}

function toast(msg, isError) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.className = "toast show" + (isError ? " error" : "");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => (el.className = "toast"), 2600);
}

// ---------------------------------------------------------------------
// App state
// ---------------------------------------------------------------------

const state = {
  view: "capture", // 'capture' | 'cases' | 'settings'
  draft: null, // in-progress case being built on the Capture tab
  casesFilter: "all", // 'all' | 'pending' | 'billed'
  casesSearch: "",
  codePicker: null, // { type: 'cpt'|'icd10', category, search }
  cropSource: null, // { file, url, naturalWidth, naturalHeight } — set while the crop-before-scan screen is open
  editingCaseId: null,
  ocrBusy: false,
  showRawOcr: false,
  library: { type: "cpt", category: "All", search: "", editingId: null, adding: false, formCode: "", formDesc: "", formCategory: "" },
};

function newDraft() {
  return {
    id: uuid(),
    patientName: "",
    mrn: "",
    dob: "",
    dos: todayISO(),
    surgeon: "",
    facility: FACILITIES[0],
    role: "primary",
    modifier: "",
    cptCodes: [],
    icd10Codes: [],
    notes: "",
    status: "pending",
    billedAt: null,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    lowConfidenceFields: [],
    rawOcrText: "",
  };
}

// ---------------------------------------------------------------------
// OCR (Tesseract.js)
// ---------------------------------------------------------------------

let ocrWorkerPromise = null;
function getOcrWorker() {
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = Tesseract.createWorker("eng");
  }
  return ocrWorkerPromise;
}

async function runOcr(file) {
  state.ocrBusy = true;
  render();
  try {
    const worker = await getOcrWorker();
    const { data } = await worker.recognize(file);
    const parsed = parseOcrText(data.text, data.words || []);
    parsed.rawText = data.text;
    return parsed;
  } finally {
    state.ocrBusy = false;
    render();
  }
}

// Runs OCR on whatever image (original photo or a cropped Blob from
// the crop screen below) and applies the result to the in-progress
// draft. Shared by both the "use full photo" path and the "use this
// crop" path so they stay in sync.
async function runOcrAndApply(fileOrBlob) {
  const d = state.draft;
  try {
    const extracted = await runOcr(fileOrBlob);
    Object.assign(d, {
      patientName: extracted.patientName || d.patientName,
      mrn: extracted.mrn || d.mrn,
      dob: extracted.dob || d.dob,
      lowConfidenceFields: extracted.lowConfidenceFields,
      rawOcrText: extracted.rawText || "",
    });
    render();
    toast("Scan complete — review highlighted fields");
  } catch (err) {
    console.error(err);
    toast("OCR failed — enter details manually", true);
  }
}

// ---------------------------------------------------------------------
// Crop-before-scan
// ---------------------------------------------------------------------
//
// Opens the crop screen for a just-picked/taken photo instead of
// running OCR on it immediately. Reading the image's natural size
// first (rather than trusting CSS-driven layout alone) is what makes
// the crop rectangle's pixel math against the ORIGINAL photo exact —
// see confirmCrop below.
function startCropSession(file) {
  const url = URL.createObjectURL(file);
  const probe = new Image();
  probe.onload = () => {
    state.cropSource = { file, url, naturalWidth: probe.naturalWidth, naturalHeight: probe.naturalHeight };
    render();
  };
  probe.onerror = () => {
    // Can't even decode it as an image here — don't block the scan on
    // a crop step that can't work; just run OCR on the original file.
    URL.revokeObjectURL(url);
    runOcrAndApply(file);
  };
  probe.src = url;
}

function closeCropSession() {
  if (state.cropSource) URL.revokeObjectURL(state.cropSource.url);
  state.cropSource = null;
}

// Wires up the crop screen's drag-to-move / drag-to-resize rectangle.
// Deliberately does NOT go through state + render() on every pointer
// move — that would tear down and rebuild the image/canvas on every
// frame of a drag, which is both janky and would lose the gesture
// mid-drag. Instead the rectangle's own DOM element is updated
// directly, and state/render() are only touched at the start and end
// of the whole crop session (open / confirm / cancel).
function bindCropEvents() {
  const stage = document.getElementById("cropStage");
  const img = document.getElementById("cropImg");
  const rectEl = document.getElementById("cropRect");
  if (!stage || !img || !rectEl) return;

  const MIN_SIZE = 40;
  let rect = { x: 0, y: 0, w: 0, h: 0 };

  function applyRect() {
    rectEl.style.left = rect.x + "px";
    rectEl.style.top = rect.y + "px";
    rectEl.style.width = rect.w + "px";
    rectEl.style.height = rect.h + "px";
  }

  function initRect() {
    const stageW = stage.clientWidth;
    const stageH = stage.clientHeight;
    // Start inset ~8% from each edge rather than the full photo — on
    // all four example stickers the sticker itself sits well inside
    // the photo's outer border, so this default usually needs only a
    // small nudge instead of a drag from a corner-sized starting box.
    const insetX = stageW * 0.08;
    const insetY = stageH * 0.08;
    rect = { x: insetX, y: insetY, w: stageW - insetX * 2, h: stageH - insetY * 2 };
    applyRect();
  }

  if (img.complete && img.naturalWidth) initRect();
  else img.addEventListener("load", initRect, { once: true });

  function clamp(r) {
    const stageW = stage.clientWidth;
    const stageH = stage.clientHeight;
    r.w = Math.max(MIN_SIZE, Math.min(r.w, stageW));
    r.h = Math.max(MIN_SIZE, Math.min(r.h, stageH));
    r.x = Math.max(0, Math.min(r.x, stageW - r.w));
    r.y = Math.max(0, Math.min(r.y, stageH - r.h));
    return r;
  }

  let dragMode = null; // 'move' | 'nw' | 'ne' | 'sw' | 'se'
  let startPointer = { x: 0, y: 0 };
  let startRect = null;

  function beginDrag(mode) {
    return (e) => {
      dragMode = mode;
      startPointer = { x: e.clientX, y: e.clientY };
      startRect = { ...rect };
      if (e.target.setPointerCapture) e.target.setPointerCapture(e.pointerId);
      e.preventDefault();
      e.stopPropagation();
    };
  }

  function onPointerMove(e) {
    if (!dragMode) return;
    const dx = e.clientX - startPointer.x;
    const dy = e.clientY - startPointer.y;
    const r = { ...startRect };
    if (dragMode === "move") {
      r.x = startRect.x + dx;
      r.y = startRect.y + dy;
    } else {
      if (dragMode.includes("n")) { r.y = startRect.y + dy; r.h = startRect.h - dy; }
      if (dragMode.includes("s")) { r.h = startRect.h + dy; }
      if (dragMode.includes("w")) { r.x = startRect.x + dx; r.w = startRect.w - dx; }
      if (dragMode.includes("e")) { r.w = startRect.w + dx; }
    }
    rect = clamp(r);
    applyRect();
  }

  function onPointerUp() {
    dragMode = null;
  }

  rectEl.addEventListener("pointerdown", beginDrag("move"));
  rectEl.querySelectorAll(".crop-handle").forEach((h) =>
    h.addEventListener("pointerdown", beginDrag(h.dataset.handle))
  );
  stage.addEventListener("pointermove", onPointerMove);
  stage.addEventListener("pointerup", onPointerUp);
  stage.addEventListener("pointercancel", onPointerUp);

  const useCropBtn = document.getElementById("useCropBtn");
  if (useCropBtn) useCropBtn.addEventListener("click", () => confirmCrop(rect, img, stage));

  const useFullPhotoBtn = document.getElementById("useFullPhotoBtn");
  if (useFullPhotoBtn) useFullPhotoBtn.addEventListener("click", () => {
    const { file } = state.cropSource;
    closeCropSession();
    render();
    runOcrAndApply(file);
  });

  const cancelCropBtn = document.getElementById("cancelCropBtn");
  if (cancelCropBtn) cancelCropBtn.addEventListener("click", () => {
    closeCropSession();
    render();
  });
}

// Cuts the selected rectangle out of the ORIGINAL photo at full
// resolution (not the on-screen scaled-down display size) by scaling
// the displayed rect back up using the ratio between the image's
// natural size and its displayed size, then drawing just that region
// onto an offscreen canvas. The resulting Blob is what actually gets
// handed to Tesseract — the full original photo never does.
function confirmCrop(rect, img, stage) {
  const { file } = state.cropSource;
  const scaleX = img.naturalWidth / stage.clientWidth;
  const scaleY = img.naturalHeight / stage.clientHeight;
  const sx = Math.round(rect.x * scaleX);
  const sy = Math.round(rect.y * scaleY);
  const sw = Math.round(rect.w * scaleX);
  const sh = Math.round(rect.h * scaleY);

  const canvas = document.createElement("canvas");
  canvas.width = sw;
  canvas.height = sh;
  canvas.getContext("2d").drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh);

  canvas.toBlob((blob) => {
    closeCropSession();
    render();
    runOcrAndApply(blob || file); // fall back to the uncropped photo if toBlob ever fails
  }, "image/jpeg", 0.92);
}

// Lines that are almost certainly NOT the patient name, used to filter
// candidate lines out of the name-guessing fallback below (hospital
// stickers are full of other short caps-heavy lines: facility name,
// "PATIENT LABEL", room/bed, barcode text, etc.).
const NAME_EXCLUDE_WORDS = /\b(DOB|MRN|DATE|FACILITY|HOSPITAL|ROOM|BED|ACCOUNT|ACCT|SURGEON|PHYSICIAN|DOCTOR|ADDRESS|PHONE|ADMIT|PATIENT LABEL|CONFIDENTIAL|SPECIMEN|ALLERGY|ALLERGIES)\b/i;

// Face sheets and stickers routinely carry OTHER "Last, First"-shaped
// names beside the patient's — a guarantor, an emergency contact, the
// admitting/referring/attending physician — and any of those can look
// exactly like a plain patient name once OCR flattens the layout to
// text. This checks the matched line and the line right before it
// (OCR often splits a table's label onto its own line from the value
// beside it) for a role that disqualifies the match.
const NAME_DISQUALIFY_CONTEXT = /\b(GUARANTOR|GUARDIAN|RESPONSIBLE\s*PART(?:Y|IES)|EMERGENCY\s*CONTACT|NEXT\s*OF\s*KIN|INSURED|SUBSCRIBER|POLICY\s*HOLDER|ATTN|ATTENDING|ADMITTING|REFERRING|SURGEON|PHYSICIAN|PROVIDER|DOCTOR)\b/i;

// Some stickers print the attending physician as "LAST, MD, FIRST, M"
// (e.g. "PAULK, MD, NICHOLAS, J") — a credential sitting where a first
// name would in a plain "Last, First" match. Reject those.
const CREDENTIAL_WORD = /^(MD|DO|PA|PA-C|NP|DPM|RN|CRNA|PHD)$/i;

// Dates on a face sheet that are NOT the birthdate — encounter date,
// date of service, admission/discharge, etc. — checked around a date
// match when no recognized birth-date label was found at all, so the
// generic "any date on the page" fallback below doesn't just grab
// whichever date happens to print first (often the encounter date at
// the very top of the sheet).
const DOB_DISQUALIFY_CONTEXT = /\b(ENCOUNTER|SERVICE|ADMIT|ADMISSION|DISCHARGE|SURGERY|VISIT|DOS|CREATED|PRINTED|COLLECTED|SIGNED)\b/i;

// Heuristic extraction of name / MRN / DOB from raw OCR text. Hospital
// stickers and face sheets vary a lot by facility, so this looks for
// common label patterns and common date/ID shapes rather than assuming
// one fixed layout. Anything it can't find with confidence is left
// blank for manual entry, and fields that DID get a hit but from a
// low-confidence OCR read are flagged in lowConfidenceFields. The raw
// recognized text is always kept (see runOcr) so a scan that comes back
// wrong or incomplete can be inspected on-device via the "View scanned
// text" toggle in the Capture tab, without needing to share real
// patient data anywhere to debug it.
function parseOcrText(text, words) {
  const result = { patientName: "", mrn: "", dob: "", lowConfidenceFields: [] };
  const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);

  const avgConfidence = (snippet) => {
    if (!words.length) return 100;
    const hits = words.filter((w) => snippet.includes(w.text) && w.text.length > 1);
    if (!hits.length) return 100;
    return hits.reduce((s, w) => s + (w.confidence || 0), 0) / hits.length;
  };

  // DOB: three passes, most reliable first.
  //
  // 1. A date sharing a line with an age marker — "BD 7/6/1980 (44
  //    yrs)", "Male, 72 y.o., 3/14/1954" — since that pairing only
  //    ever describes a birthdate. This is checked FIRST and matched
  //    by shape rather than by label text on purpose: OCR frequently
  //    misreads short all-caps labels (this exact sheet reads "BD" as
  //    "8D" — B/8 is a very common OCR confusion), so anchoring to the
  //    label alone silently loses the read even when the date and age
  //    right next to it came through fine.
  // 2. Failing that, a recognized birth-date label (DOB / D.O.B. /
  //    Birth Date / BD / Born) — covers sheets with no age shown.
  // 3. Failing that, scan every date on the page and skip ones sitting
  //    next to an obviously different date field (encounter/service/
  //    admission/DOS/discharge, etc.) rather than just grabbing
  //    whichever date happens to appear first (often an encounter date
  //    up top).
  const AGE_MARKER_RE = /\b\d{1,2}\s*(?:yrs?\.?|y\.?o\.?|years?\s*old|years?)\b/i;
  const DATE_RE = /\b(\d{1,2}[\/\-]\d{1,2}[\/\-](?:19|20)\d{2})\b/;

  let dobRaw = null;
  for (const line of lines) {
    if (!AGE_MARKER_RE.test(line)) continue;
    const dm = line.match(DATE_RE);
    if (dm) { dobRaw = dm[1]; break; }
  }
  if (!dobRaw) {
    const dobLabelMatch = text.match(/(?:DOB|D\.?O\.?B\.?|Birth\s*Date|\bBD\b|\bBorn\b)\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i);
    dobRaw = dobLabelMatch && dobLabelMatch[1];
  }
  if (!dobRaw) {
    for (const m of text.matchAll(/\b(\d{1,2}[\/\-]\d{1,2}[\/\-](?:19|20)\d{2})\b/g)) {
      const context = text.slice(Math.max(0, m.index - 20), m.index);
      if (DOB_DISQUALIFY_CONTEXT.test(context)) continue;
      dobRaw = m[1];
      break;
    }
  }
  if (dobRaw) {
    const norm = normalizeDate(dobRaw);
    if (norm) {
      result.dob = norm;
      if (avgConfidence(dobRaw) < 70) result.lowConfidenceFields.push("dob");
    }
  }

  // MRN: "MRN" and "Account"/"Acct" are NOT interchangeable — a face
  // sheet routinely prints both, as two DIFFERENT numbers (the medical
  // record number vs. the encounter/billing account number). Matching
  // them with one regex meant whichever label happened to appear first
  // in reading order won, even when it was the wrong one. So: try
  // "MRN" (tolerating "MIRN" — OCR inserting a stray letter into a
  // short all-caps label, the same kind of noise "BD" -> "8D" was)
  // and "Med Rec" FIRST, and only fall back to "Account"/"Acct"/"ID"
  // if no MRN-specific label was found anywhere at all. Below that,
  // fall back to any standalone 6-10 digit run elsewhere on the
  // sticker (common when the MRN sits under a barcode with no text
  // label at all), skipping the digits already used as DOB.
  const mrnMatch =
    text.match(/(?:MRN|M\.?I?\.?RN|Med(?:ical)?\s*Rec(?:ord)?\s*(?:No\.?|#|Number)?)\s*[:\-]?\s*([A-Z0-9\-]{4,15})/i) ||
    text.match(/(?:Acct\.?\s*#?|Account\s*(?:No\.?|#)?|Patient\s*ID|ID\s*#)\s*[:\-]?\s*([A-Z0-9\-]{4,15})/i);
  if (mrnMatch) {
    result.mrn = mrnMatch[1];
    if (avgConfidence(mrnMatch[1]) < 70) result.lowConfidenceFields.push("mrn");
  } else {
    const dobDigits = dobRaw ? dobRaw.replace(/\D/g, "") : null;
    const bareNumber = (text.match(/\b\d{6,10}\b/g) || []).find((n) => !dobDigits || !dobDigits.includes(n));
    if (bareNumber) {
      result.mrn = bareNumber;
      result.lowConfidenceFields.push("mrn"); // unlabeled guess — always flag for review
    }
  }

  // Name: look for a "Patient"/"Name" LABEL first (most reliable when
  // present), then fall back to scanning every "Last, First"-shaped
  // line for the best patient candidate, then finally a plain 2-4 word
  // ALL-CAPS line (e.g. "SMITH JOHN A") that isn't one of the known
  // non-name labels above — common on stickers with no comma in the
  // name. Both the label search and the line scan skip anything tied
  // to a guarantor/contact/physician role (see NAME_DISQUALIFY_CONTEXT)
  // instead of just taking the first match in reading order, since
  // that role's name often appears ABOVE the patient's on a face sheet.
  const patientHeaderIdx = lines.findIndex((l) => /^PATIENT\b/i.test(l));

  let nameGuess = null;
  let nameFromExplicitLabel = false;
  for (const m of text.matchAll(/(?:Patient(?:\s*Name)?|Name)\s*[:\-]\s*([A-Za-z,'.\- ]{3,40})/gi)) {
    // Check a bit of text before the label too, not just the match
    // itself — "Guarantor Name:" and "Emergency Contact Name:" don't
    // contain "Patient"/"Name" at their start, so the match alone
    // would miss the disqualifying word in front of it.
    const context = text.slice(Math.max(0, m.index - 25), m.index + m[0].length);
    if (NAME_DISQUALIFY_CONTEXT.test(context)) continue;
    nameGuess = m[1].trim();
    nameFromExplicitLabel = true;
    break;
  }

  if (!nameGuess) {
    // Score every clean "Last, First" pair found ANYWHERE in each line
    // (not just at the start — a label like "Name" or "Referring"
    // commonly sits before the name on the same OCR line rather than
    // on a line of its own) and keep the best one: skip a whole line
    // when its own text names a non-patient role, and also pull in the
    // line before it as context ONLY when that previous line ITSELF
    // ends with one of those role words (optionally followed by a
    // colon) — a label that got flattened onto the tail of the
    // previous line by OCR, with its value continuing on this line
    // (e.g. "...Salt Lake City, Utah Guarantor:" / next line
    // "OLLERTON, JENNIFER KAYE"). This is deliberately NOT "does the
    // previous line contain a comma at all" — an address line like
    // "Salt Lake City, Utah Guarantor:" has its own unrelated comma,
    // and a line like "Attn: Paulk, MD, Nicholas, J" already carries
    // its own value, so pulling either of those wholesale into the
    // NEXT line's context would wrongly disqualify an unrelated real
    // patient name sitting right after them. Also skip a pair whose
    // captured "first name" is actually a credential (e.g. "PAULK,
    // MD, NICHOLAS, J"). When the sheet has a "PATIENT" section
    // header, prefer a candidate found at or after it over one found
    // above it (e.g. a guarantor/referring line higher up the page);
    // otherwise take the first clean candidate found, in reading
    // order.
    const labelTailRe = new RegExp(NAME_DISQUALIFY_CONTEXT.source.replace(/^\\b\(/, "(") + "\\s*:?\\s*$", "i");
    let best = null;
    lines.forEach((line, i) => {
      const prevLine = lines[i - 1] || "";
      const prevEndsWithLabel = labelTailRe.test(prevLine);
      const context = prevEndsWithLabel ? `${prevLine} ${line}` : line;
      if (NAME_DISQUALIFY_CONTEXT.test(context)) return;
      for (const lf of line.matchAll(/([A-Z][A-Za-z'\-]+),\s*([A-Z][A-Za-z'\-]+)/g)) {
        if (CREDENTIAL_WORD.test(lf[2])) continue;
        const score = patientHeaderIdx >= 0 && i < patientHeaderIdx ? 1 : 0;
        if (!best || score < best.score) best = { text: lf[0], score };
      }
    });
    if (best) nameGuess = best.text;
  }
  if (!nameGuess) {
    const capsLine = lines.find(
      (l) =>
        /^[A-Z][A-Z'\-]+(?:\s+[A-Z][A-Z'\-]*){1,3}$/.test(l) &&
        l.length <= 35 &&
        !NAME_EXCLUDE_WORDS.test(l) &&
        !/\d/.test(l)
    );
    if (capsLine) nameGuess = capsLine;
  }
  if (nameGuess) {
    result.patientName = nameGuess.replace(/\s{2,}/g, " ").trim();
    // A name pulled from an explicit "Name:"/"Patient:" label is only
    // flagged if the OCR read itself was shaky. A name found by the
    // "Last, First"-shape fallback or the bare ALL-CAPS fallback has no
    // such label to back it up — it's a structural guess, and nothing
    // stops it from confidently matching something that ISN'T the
    // patient at all (e.g. a crisp, clearly-legible watermark/stamp
    // elsewhere in the photo, which Tesseract may read with high
    // per-word confidence even though it's the wrong text entirely —
    // confidence score alone can't catch that). So any fallback-derived
    // guess is always flagged for review, regardless of how "confident"
    // OCR was about the characters themselves.
    if (!nameFromExplicitLabel || avgConfidence(nameGuess) < 70) {
      result.lowConfidenceFields.push("patientName");
    }
  }

  return result;
}

function normalizeDate(raw) {
  const m = raw.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!m) return null;
  let [, mo, da, yr] = m;
  if (yr.length === 2) yr = (Number(yr) > 30 ? "19" : "20") + yr;
  mo = mo.padStart(2, "0");
  da = da.padStart(2, "0");
  return `${yr}-${mo}-${da}`;
}

// ---------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------

const app = document.getElementById("app");

function render() {
  const tab = `
    <nav class="tabbar">
      <button class="tab ${state.view === "capture" ? "active" : ""}" data-nav="capture">
        <span class="icon">📷</span>Capture
      </button>
      <button class="tab ${state.view === "cases" ? "active" : ""}" data-nav="cases">
        <span class="icon">🗂️</span>Cases
        ${CASES.filter((c) => c.status === "pending").length ? `<span class="badge">${CASES.filter((c) => c.status === "pending").length}</span>` : ""}
      </button>
      <button class="tab ${state.view === "settings" ? "active" : ""}" data-nav="settings">
        <span class="icon">⚙️</span>Settings
      </button>
    </nav>`;

  let body = "";
  if (state.view === "capture") body = renderCapture();
  else if (state.view === "cases") body = renderCases();
  else if (state.view === "settings") body = renderSettings();
  else if (state.view === "library") body = renderLibrary();

  app.innerHTML = `<div class="screen">${body}</div>${tab}${state.codePicker ? renderCodePicker() : ""}${state.cropSource ? renderCropModal() : ""}`;
  bindEvents();
}

function renderCapture() {
  if (!state.draft) state.draft = newDraft();
  const d = state.draft;
  const lowConf = new Set(d.lowConfidenceFields || []);

  return `
    <header class="topbar"><h1>New Case</h1></header>
    <div class="content">
      <section class="card">
        <label class="capture-btn">
          📷 ${state.ocrBusy ? "Reading sticker…" : "Scan patient sticker / face sheet"}
          <input id="camInput" type="file" accept="image/*" capture="environment" ${state.ocrBusy ? "disabled" : ""} />
        </label>
        ${state.ocrBusy ? '<div class="spinner"></div>' : ""}
        ${d.rawOcrText ? `
          <button class="link-btn" data-toggle-raw-ocr="1">${state.showRawOcr ? "Hide" : "View"} scanned text</button>
          ${state.showRawOcr ? `<pre class="raw-ocr">${escapeHtml(d.rawOcrText)}</pre>` : ""}
        ` : ""}
      </section>

      <section class="card">
        <h2>Patient</h2>
        <div class="field ${lowConf.has("patientName") ? "low-conf" : ""}">
          <label>Name</label>
          <input id="f_patientName" type="text" value="${escapeHtml(d.patientName)}" placeholder="Last, First" />
        </div>
        <div class="row2">
          <div class="field ${lowConf.has("mrn") ? "low-conf" : ""}">
            <label>MRN</label>
            <input id="f_mrn" type="text" value="${escapeHtml(d.mrn)}" />
          </div>
          <div class="field ${lowConf.has("dob") ? "low-conf" : ""}">
            <label>DOB</label>
            <input id="f_dob" type="date" value="${escapeHtml(d.dob)}" />
          </div>
        </div>
        <div class="row2">
          <div class="field">
            <label>Date of service</label>
            <input id="f_dos" type="date" value="${escapeHtml(d.dos)}" />
          </div>
          <div class="field">
            <label>Facility</label>
            <select id="f_facility">
              ${FACILITIES.map((f) => `<option value="${f}" ${d.facility === f ? "selected" : ""}>${f}</option>`).join("")}
            </select>
          </div>
        </div>
        <div class="field">
          <label>Surgeon</label>
          <input id="f_surgeon" type="text" value="${escapeHtml(d.surgeon)}" placeholder="e.g. Paulk" />
        </div>
      </section>

      <section class="card">
        <h2>Role</h2>
        <div class="role-select">
          ${[
            ["primary", "Primary", ""],
            ["assistant", "Assistant", "Mod 80"],
            ["cosurgeon", "Co-surgeon", "Mod 62"],
          ].map(([val, label, mod]) => `
            <button class="role-btn ${d.role === val ? "active" : ""}" data-role="${val}">
              ${label}${mod ? `<small>${mod}</small>` : ""}
            </button>`).join("")}
        </div>
      </section>

      <section class="card">
        <h2>CPT codes <span class="count">${d.cptCodes.length}</span></h2>
        ${renderCodeChips(d.cptCodes, "cpt")}
        <button class="add-code-btn" data-open-picker="cpt">+ Add CPT code</button>
      </section>

      <section class="card">
        <h2>ICD-10 codes <span class="count">${d.icd10Codes.length}</span></h2>
        ${renderCodeChips(d.icd10Codes, "icd10")}
        <button class="add-code-btn" data-open-picker="icd10">+ Add ICD-10 code</button>
      </section>

      <section class="card">
        <h2>Notes / modifier</h2>
        <textarea id="f_notes" rows="3" placeholder="Case notes, additional modifiers…">${escapeHtml(d.notes)}</textarea>
      </section>

      <button id="saveCaseBtn" class="primary-btn" ${!d.patientName ? "disabled" : ""}>Save case</button>
    </div>`;
}

function renderCodeChips(list, type) {
  if (!list.length) return `<p class="empty-hint">No codes added yet.</p>`;
  return `<div class="chips">${list.map((c, i) => `
    <div class="chip">
      <div class="chip-main">
        <strong>${escapeHtml(c.code)}</strong>
        <span>${escapeHtml(c.label || c.desc)}</span>
      </div>
      <button class="chip-remove" data-remove-code="${type}:${i}">×</button>
    </div>`).join("")}</div>`;
}

function renderCases() {
  const q = state.casesSearch.trim().toLowerCase();
  let list = CASES.filter((c) => {
    if (state.casesFilter === "pending" && c.status !== "pending") return false;
    if (state.casesFilter === "billed" && c.status !== "billed") return false;
    if (q && !(`${c.patientName} ${c.mrn} ${c.facility}`.toLowerCase().includes(q))) return false;
    return true;
  });

  return `
    <header class="topbar"><h1>Cases</h1></header>
    <div class="content">
      <input id="casesSearch" class="search-input" type="search" placeholder="Search patient, MRN, facility…" value="${escapeHtml(state.casesSearch)}" />
      <div class="filter-row">
        ${["all", "pending", "billed"].map((f) => `
          <button class="filter-chip ${state.casesFilter === f ? "active" : ""}" data-filter="${f}">
            ${f[0].toUpperCase() + f.slice(1)}
          </button>`).join("")}
      </div>
      ${list.length ? "" : '<p class="empty-hint">No cases match.</p>'}
      <div class="case-list">
        ${list.map(renderCaseCard).join("")}
      </div>
      ${CASES.length ? `
      <div class="export-row">
        <button class="secondary-btn" data-export="biller">Copy biller summary</button>
        <button class="secondary-btn" data-export="csv-all">CSV: All</button>
        <button class="secondary-btn" data-export="csv-billed">CSV: Billed</button>
        <button class="secondary-btn" data-export="csv-pending">CSV: Pending</button>
      </div>` : ""}
    </div>`;
}

function renderCaseCard(c) {
  return `
    <div class="case-card" data-edit-case="${c.id}">
      <button class="status-dot ${c.status}" data-toggle-status="${c.id}" title="Tap to toggle billed status"></button>
      <div class="case-main">
        <div class="case-title">${escapeHtml(c.patientName) || "(no name)"} <span class="muted">${escapeHtml(c.facility)}</span></div>
        <div class="case-sub muted">${fmtDate(c.dos)} · ${c.cptCodes.length} CPT · ${c.icd10Codes.length} ICD-10</div>
      </div>
      <div class="case-chevron">›</div>
    </div>`;
}

function renderSettings() {
  const driveConnected = !!state._driveToken;
  return `
    <header class="topbar"><h1>Settings</h1></header>
    <div class="content">
      <section class="card">
        <h2>Google Drive sync</h2>
        <p class="muted">Every saved case backs up automatically to a spreadsheet called <strong>${SHEET_NAME}</strong> in your Google Drive — primary/co-surgeon cases on one tab, assistant cases on another, plus a Monthly Tally tab that auto-counts Bariatric/EGD/General Surgery cases per month. Data goes only to your own Google account — no other server is involved.</p>
        <div class="field">
          <label>Google OAuth Client ID</label>
          <input id="f_clientId" type="text" value="${escapeHtml(state._clientId || "")}" placeholder="xxxx.apps.googleusercontent.com" />
        </div>
        <button id="saveClientIdBtn" class="secondary-btn">Save client ID</button>
        <button id="driveConnectBtn" class="secondary-btn" ${state._clientId ? "" : "disabled"}>
          ${driveConnected ? "✓ Connected — reconnect" : "Connect Google Drive"}
        </button>
        ${!state._clientId ? '<p class="hint">Set up a Client ID in Google Cloud Console — see README.md. Everything else in the app works without this.</p>' : ""}
        ${SYNC_QUEUE.length ? `<p class="hint">${SYNC_QUEUE.length} case(s) queued to sync once connected/online.</p>` : ""}
      </section>

      <section class="card">
        <h2>Code library</h2>
        <p class="muted">${CPT_LIB.length} CPT · ${ICD10_LIB.length} ICD-10 codes. Add, edit, or remove codes any time — changes are saved on this device immediately, no redeploy needed.</p>
        <button class="secondary-btn" data-nav="library">Manage code library</button>
      </section>

      <section class="card">
        <h2>About</h2>
        <p class="muted">ChargeCap v1.0 · ${CASES.length} case(s) stored locally on this device (IndexedDB).</p>
        <p class="muted small">CPT/ICD-10 favorites were imported from your OR billing sheets — double-check codes against current documentation before relying on them for claims.</p>
      </section>
    </div>`;
}

function renderLibrary() {
  const lib = state.library;
  const source = codeLib(lib.type);
  const categories = ["All", ...new Set(source.map((c) => c.category))];
  const q = lib.search.trim().toLowerCase();
  const results = source
    .filter((c) => lib.category === "All" || c.category === lib.category)
    .filter((c) => !q || `${c.code} ${c.desc}`.toLowerCase().includes(q))
    .sort((a, b) => a.category.localeCompare(b.category) || a.code.localeCompare(b.code));

  const editing = lib.editingId ? source.find((c) => c.id === lib.editingId) : null;
  const showForm = lib.adding || editing;

  return `
    <header class="topbar">
      <button class="back-btn" data-nav="settings">‹ Settings</button>
      <h1>Code library</h1>
    </header>
    <div class="content">
      <div class="filter-row">
        ${["cpt", "icd10"].map((t) => `
          <button class="filter-chip ${lib.type === t ? "active" : ""}" data-lib-type="${t}">${t === "cpt" ? "CPT" : "ICD-10"}</button>`).join("")}
      </div>

      ${showForm ? renderLibraryForm(editing) : `
        <button class="add-code-btn" data-lib-add="1">+ Add ${lib.type === "cpt" ? "CPT" : "ICD-10"} code</button>
      `}

      <input id="librarySearch" class="search-input" type="search" placeholder="Search code or description…" value="${escapeHtml(lib.search)}" />
      <div class="cat-row">
        ${categories.map((c) => `<button class="filter-chip ${lib.category === c ? "active" : ""}" data-lib-cat="${escapeHtml(c)}">${escapeHtml(c)}</button>`).join("")}
      </div>

      <div class="picker-list">
        ${results.map((c) => `
          <div class="lib-row">
            <div class="lib-row-main">
              <strong>${escapeHtml(c.code)}</strong>
              <span>${escapeHtml(c.desc)}</span>
              <small class="muted">${escapeHtml(c.category)}</small>
            </div>
            <div class="lib-row-actions">
              <button class="icon-btn" data-lib-edit="${c.id}" title="Edit">✎</button>
              <button class="icon-btn danger" data-lib-delete="${c.id}" title="Delete">🗑</button>
            </div>
          </div>`).join("")}
        ${!results.length ? '<p class="empty-hint">No codes match.</p>' : ""}
      </div>
    </div>`;
}

function renderLibraryForm(editing) {
  const lib = state.library;
  const code = editing ? editing.code : lib.formCode;
  const desc = editing ? editing.desc : lib.formDesc;
  const category = editing ? editing.category : lib.formCategory;
  return `
    <section class="card">
      <h2>${editing ? "Edit code" : "Add code"}</h2>
      <div class="field">
        <label>Code</label>
        <input id="lf_code" type="text" value="${escapeHtml(code)}" placeholder="e.g. 43775" />
      </div>
      <div class="field">
        <label>Description</label>
        <input id="lf_desc" type="text" value="${escapeHtml(desc)}" placeholder="e.g. Sleeve gastrectomy — LAP" />
      </div>
      <div class="field">
        <label>Category</label>
        <input id="lf_category" type="text" value="${escapeHtml(category)}" placeholder="e.g. Bariatric" />
      </div>
      <div class="form-row">
        <button class="primary-btn" data-lib-save="${editing ? editing.id : "new"}">Save</button>
        <button class="secondary-btn" data-lib-cancel="1">Cancel</button>
      </div>
    </section>`;
}

function renderCodePicker() {
  const { type, category, search } = state.codePicker;
  const source = codeLib(type);
  const categories = ["All", ...new Set(source.map((c) => c.category))];
  const q = search.trim().toLowerCase();
  let results = source.filter((c) => {
    if (category !== "All" && c.category !== category) return false;
    if (q && !(`${c.code} ${c.desc}`.toLowerCase().includes(q))) return false;
    return true;
  });

  return `
    <div class="modal-backdrop" data-close-picker="1">
      <div class="modal" data-stop="1">
        <div class="modal-header">
          <h2>Add ${type === "cpt" ? "CPT" : "ICD-10"} code</h2>
          <button class="close-btn" data-close-picker="1">×</button>
        </div>
        <input id="pickerSearch" class="search-input" type="search" placeholder="Search code or description…" value="${escapeHtml(search)}" autofocus />
        <div class="cat-row">
          ${categories.map((c) => `<button class="filter-chip ${category === c ? "active" : ""}" data-picker-cat="${escapeHtml(c)}">${escapeHtml(c)}</button>`).join("")}
        </div>
        <div class="picker-list">
          ${results.map((c) => `
            <button class="picker-item" data-pick="${escapeHtml(c.id)}">
              <strong>${escapeHtml(c.code)}</strong>
              <span>${escapeHtml(c.desc)}</span>
            </button>`).join("")}
          ${!results.length ? '<p class="empty-hint">No matches in favorites.</p>' : ""}
        </div>
        ${type === "icd10" ? `
        <div id="liveSearchArea">
          <button id="liveSearchBtn" class="secondary-btn">Search full NLM ICD-10 database for "${escapeHtml(search)}"</button>
          <div id="liveResults"></div>
        </div>` : ""}
      </div>
    </div>`;
}

// Shown right after a photo is picked/taken, before OCR ever sees it —
// lets the user drag a crop box down to just the sticker/face sheet so
// stray text nearby (a watermark, another patient's sticker on the
// same sheet, a monitor showing a different chart) can't confuse the
// scan the way it did with the two example photos that had this issue.
// "Use full photo" skips straight to OCR on the original image for
// when a photo is already tightly framed.
function renderCropModal() {
  const { url } = state.cropSource;
  return `
    <div class="modal-backdrop">
      <div class="crop-modal" data-stop="1">
        <div class="crop-header">
          <h2>Crop to just the sticker</h2>
          <p class="crop-hint">Drag the corners to trim out anything that isn't the patient sticker or face sheet — this keeps stray text nearby from confusing the scan.</p>
        </div>
        <div id="cropStage" class="crop-stage">
          <img id="cropImg" class="crop-image" src="${url}" alt="Captured photo" />
          <div id="cropRect" class="crop-rect">
            <div class="crop-handle" data-handle="nw"></div>
            <div class="crop-handle" data-handle="ne"></div>
            <div class="crop-handle" data-handle="sw"></div>
            <div class="crop-handle" data-handle="se"></div>
          </div>
        </div>
        <div class="crop-actions">
          <button id="useCropBtn" class="primary-btn">Use this crop</button>
          <button id="useFullPhotoBtn" class="secondary-btn">Use full photo</button>
          <button id="cancelCropBtn" class="secondary-btn">Cancel</button>
        </div>
      </div>
    </div>`;
}

// ---------------------------------------------------------------------
// Event binding
// ---------------------------------------------------------------------

function bindEvents() {
  app.querySelectorAll("[data-nav]").forEach((el) =>
    el.addEventListener("click", () => {
      state.view = el.dataset.nav;
      state.codePicker = null;
      render();
    })
  );

  if (state.view === "capture") bindCaptureEvents();
  if (state.view === "cases") bindCasesEvents();
  if (state.view === "settings") bindSettingsEvents();
  if (state.view === "library") bindLibraryEvents();
  if (state.codePicker) bindPickerEvents();
  if (state.cropSource) bindCropEvents();
}

function bindCaptureEvents() {
  const d = state.draft;
  const camInput = document.getElementById("camInput");
  if (camInput) {
    camInput.addEventListener("change", (e) => {
      const file = e.target.files[0];
      // Reset so picking the exact same file again still fires "change".
      e.target.value = "";
      if (!file) return;
      startCropSession(file);
    });
  }

  const toggleRawBtn = document.querySelector("[data-toggle-raw-ocr]");
  if (toggleRawBtn) toggleRawBtn.addEventListener("click", () => {
    state.showRawOcr = !state.showRawOcr;
    render();
  });

  ["patientName", "mrn", "dob", "dos", "surgeon", "notes"].forEach((f) => {
    const el = document.getElementById("f_" + f);
    if (el) el.addEventListener("input", () => { d[f] = el.value; syncSaveButton(); });
  });
  const facilityEl = document.getElementById("f_facility");
  if (facilityEl) facilityEl.addEventListener("change", () => (d.facility = facilityEl.value));

  app.querySelectorAll("[data-role]").forEach((el) =>
    el.addEventListener("click", () => {
      d.role = el.dataset.role;
      d.modifier = d.role === "assistant" ? "80" : d.role === "cosurgeon" ? "62" : "";
      render();
    })
  );

  app.querySelectorAll("[data-open-picker]").forEach((el) =>
    el.addEventListener("click", () => {
      state.codePicker = { type: el.dataset.openPicker, category: "All", search: "" };
      render();
    })
  );

  app.querySelectorAll("[data-remove-code]").forEach((el) =>
    el.addEventListener("click", () => {
      const [type, idx] = el.dataset.removeCode.split(":");
      const key = type === "cpt" ? "cptCodes" : "icd10Codes";
      d[key].splice(Number(idx), 1);
      render();
    })
  );

  const saveBtn = document.getElementById("saveCaseBtn");
  if (saveBtn) saveBtn.addEventListener("click", async () => {
    await saveCase(d);
    toast("Case saved");
    state.draft = null;
    state.view = "cases";
    render();
    // Fire-and-forget: don't make the user wait on the network round
    // trip just to see their case in the list. syncCaseToDrive queues
    // itself silently if Drive isn't connected or the request fails.
    syncCaseToDrive(d);
  });
}

function syncSaveButton() {
  const btn = document.getElementById("saveCaseBtn");
  if (btn) btn.disabled = !state.draft.patientName;
}

function bindCasesEvents() {
  const search = document.getElementById("casesSearch");
  if (search) search.addEventListener("input", () => { state.casesSearch = search.value; render(); });

  app.querySelectorAll("[data-filter]").forEach((el) =>
    el.addEventListener("click", () => { state.casesFilter = el.dataset.filter; render(); })
  );

  app.querySelectorAll("[data-toggle-status]").forEach((el) =>
    el.addEventListener("click", async (e) => {
      e.stopPropagation();
      const c = CASES.find((x) => x.id === el.dataset.toggleStatus);
      if (!c) return;
      c.status = c.status === "billed" ? "pending" : "billed";
      c.billedAt = c.status === "billed" ? Date.now() : null;
      await saveCase(c);
      render();
      // Cases already sync on save; re-sync here just updates the
      // existing row's Billing Status / Billed Date in place.
      syncCaseToDrive(c);
    })
  );

  app.querySelectorAll("[data-edit-case]").forEach((el) =>
    el.addEventListener("click", () => {
      const c = CASES.find((x) => x.id === el.dataset.editCase);
      if (!c) return;
      state.draft = JSON.parse(JSON.stringify(c));
      state.view = "capture";
      render();
    })
  );

  app.querySelectorAll("[data-export]").forEach((el) =>
    el.addEventListener("click", () => handleExport(el.dataset.export))
  );
}

function bindSettingsEvents() {
  const saveBtn = document.getElementById("saveClientIdBtn");
  if (saveBtn) saveBtn.addEventListener("click", async () => {
    const val = document.getElementById("f_clientId").value.trim();
    state._clientId = val;
    await setMeta("googleClientId", val);
    toast("Client ID saved");
    render();
  });
  const connectBtn = document.getElementById("driveConnectBtn");
  if (connectBtn) connectBtn.addEventListener("click", connectGoogleDrive);
}

function bindLibraryEvents() {
  const lib = state.library;

  app.querySelectorAll("[data-lib-type]").forEach((el) =>
    el.addEventListener("click", () => {
      lib.type = el.dataset.libType;
      lib.category = "All";
      lib.editingId = null;
      lib.adding = false;
      render();
    })
  );

  const searchEl = document.getElementById("librarySearch");
  if (searchEl) searchEl.addEventListener("input", () => { lib.search = searchEl.value; render(); });

  app.querySelectorAll("[data-lib-cat]").forEach((el) =>
    el.addEventListener("click", () => { lib.category = el.dataset.libCat; render(); })
  );

  const addBtn = document.querySelector("[data-lib-add]");
  if (addBtn) addBtn.addEventListener("click", () => {
    lib.adding = true;
    lib.editingId = null;
    lib.formCode = "";
    lib.formDesc = "";
    lib.formCategory = lib.category !== "All" ? lib.category : "";
    render();
  });

  app.querySelectorAll("[data-lib-edit]").forEach((el) =>
    el.addEventListener("click", () => {
      lib.editingId = el.dataset.libEdit;
      lib.adding = false;
      render();
    })
  );

  app.querySelectorAll("[data-lib-delete]").forEach((el) =>
    el.addEventListener("click", async () => {
      await deleteCodeFromLib(lib.type, el.dataset.libDelete);
      toast("Code deleted");
      render();
    })
  );

  const cancelBtn = document.querySelector("[data-lib-cancel]");
  if (cancelBtn) cancelBtn.addEventListener("click", () => {
    lib.adding = false;
    lib.editingId = null;
    render();
  });

  const saveBtn = document.querySelector("[data-lib-save]");
  if (saveBtn) saveBtn.addEventListener("click", async () => {
    const code = document.getElementById("lf_code").value.trim();
    const desc = document.getElementById("lf_desc").value.trim();
    const category = document.getElementById("lf_category").value.trim() || "Custom";
    if (!code || !desc) { toast("Code and description are required", true); return; }
    const target = saveBtn.dataset.libSave;
    if (target === "new") {
      await addCodeToLib(lib.type, { code, desc, category });
      toast("Code added");
    } else {
      await updateCodeInLib(lib.type, target, { code, desc, category });
      toast("Code updated");
    }
    lib.adding = false;
    lib.editingId = null;
    render();
  });
}

function bindPickerEvents() {
  const backdrop = document.querySelector("[data-close-picker]");
  if (backdrop) {
    document.querySelectorAll("[data-close-picker]").forEach((el) =>
      el.addEventListener("click", () => { state.codePicker = null; render(); })
    );
  }
  const modal = document.querySelector("[data-stop]");
  if (modal) modal.addEventListener("click", (e) => e.stopPropagation());

  const searchEl = document.getElementById("pickerSearch");
  if (searchEl) {
    searchEl.addEventListener("input", () => {
      state.codePicker.search = searchEl.value;
      render();
      document.getElementById("pickerSearch").focus();
      document.getElementById("pickerSearch").selectionStart = document.getElementById("pickerSearch").value.length;
    });
  }

  app.querySelectorAll("[data-picker-cat]").forEach((el) =>
    el.addEventListener("click", () => { state.codePicker.category = el.dataset.pickerCat; render(); })
  );

  app.querySelectorAll("[data-pick]").forEach((el) =>
    el.addEventListener("click", () => {
      const { type } = state.codePicker;
      const source = codeLib(type);
      const entry = source.find((c) => c.id === el.dataset.pick);
      if (!entry) return;
      const key = type === "cpt" ? "cptCodes" : "icd10Codes";
      state.draft[key].push({ code: entry.code, desc: entry.desc, label: "", notes: "" });
      state.codePicker = null;
      render();
    })
  );

  const liveBtn = document.getElementById("liveSearchBtn");
  if (liveBtn) liveBtn.addEventListener("click", () => runLiveIcd10Search(state.codePicker.search));
}

async function runLiveIcd10Search(term) {
  if (!term.trim()) return;
  const resultsEl = document.getElementById("liveResults");
  resultsEl.innerHTML = '<div class="spinner"></div>';
  try {
    const url = `https://clinicaltables.nlm.nih.gov/api/icd10cm/v3/search?sf=code,name&terms=${encodeURIComponent(term)}`;
    const res = await fetch(url);
    const json = await res.json();
    const rows = json[3] || [];
    if (!rows.length) {
      resultsEl.innerHTML = '<p class="empty-hint">No matches.</p>';
      return;
    }
    resultsEl.innerHTML = `<div class="picker-list">${rows.map(([code, name]) => `
      <div class="picker-item live-item">
        <button class="live-pick" data-live-pick="${escapeHtml(code)}" data-live-desc="${escapeHtml(name)}">
          <strong>${escapeHtml(code)}</strong><span>${escapeHtml(name)}</span>
        </button>
        <button class="star-btn ${isCodeInLib("icd10", code) ? "starred" : ""}" data-live-star="${escapeHtml(code)}" data-live-desc="${escapeHtml(name)}" title="Save to favorites">
          ${isCodeInLib("icd10", code) ? "★" : "☆"}
        </button>
      </div>`).join("")}</div>`;
    resultsEl.querySelectorAll("[data-live-pick]").forEach((el) =>
      el.addEventListener("click", () => {
        state.draft.icd10Codes.push({ code: el.dataset.livePick, desc: el.dataset.liveDesc, label: "", notes: "" });
        state.codePicker = null;
        render();
      })
    );
    resultsEl.querySelectorAll("[data-live-star]").forEach((el) =>
      el.addEventListener("click", async (e) => {
        e.stopPropagation();
        const code = el.dataset.liveStar;
        if (isCodeInLib("icd10", code)) { toast("Already in favorites"); return; }
        await addCodeToLib("icd10", { code, desc: el.dataset.liveDesc, category: "Custom" });
        toast(`${code} saved to favorites`);
        runLiveIcd10Search(term);
      })
    );
  } catch (err) {
    console.error(err);
    resultsEl.innerHTML = '<p class="empty-hint">Search failed — check connection.</p>';
  }
}

// ---------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------

function handleExport(kind) {
  if (kind === "biller") return exportBillerText();
  if (kind === "csv-all") return exportCsv(CASES, "chargecap-all");
  if (kind === "csv-billed") return exportCsv(CASES.filter((c) => c.status === "billed"), "chargecap-billed");
  if (kind === "csv-pending") return exportCsv(CASES.filter((c) => c.status === "pending"), "chargecap-pending");
}

function exportBillerText() {
  const pending = CASES.filter((c) => c.status === "pending");
  const lines = pending.map((c) => {
    const cpt = c.cptCodes.map((x) => x.code + (x.modifier ? `-${x.modifier}` : "")).join(", ");
    const icd = c.icd10Codes.map((x) => x.code).join(", ");
    return `${fmtDate(c.dos)} | ${c.patientName} | MRN ${c.mrn} | DOB ${fmtDate(c.dob)} | ${c.facility} | ${c.surgeon} (${c.role}${c.modifier ? " mod " + c.modifier : ""}) | CPT: ${cpt} | ICD-10: ${icd}${c.notes ? " | Notes: " + c.notes : ""}`;
  });
  const text = lines.join("\n") || "No pending cases.";
  navigator.clipboard.writeText(text).then(
    () => toast(`Copied ${pending.length} pending case(s) to clipboard`),
    () => toast("Could not copy — clipboard permission denied", true)
  );
}

function exportCsv(list, filename) {
  const rows = [SHEET_HEADER, ...list.map(caseToRow)];
  const csv = rows.map((r) => r.map(csvEscape).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${filename}-${todayISO()}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  toast(`Exported ${list.length} case(s)`);
}

function caseToRow(c) {
  return [
    fmtDate(c.dos), c.patientName, c.mrn, fmtDate(c.dob), c.surgeon, c.role, c.modifier, c.facility,
    c.status, c.billedAt ? new Date(c.billedAt).toLocaleString() : "",
    c.cptCodes.map((x) => x.code).join("; "),
    c.cptCodes.map((x) => x.label || x.desc).join("; "),
    c.icd10Codes.map((x) => x.code).join("; "),
    c.icd10Codes.map((x) => x.label || x.desc).join("; "),
    c.notes,
    caseCategory(c),
  ];
}

// ---------------------------------------------------------------------
// Google Drive sync (OAuth via Google Identity Services + Sheets API)
// ---------------------------------------------------------------------

let tokenClient = null;

function initGoogleAuth() {
  if (!state._clientId || !window.google) return;
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: state._clientId,
    scope: SHEETS_SCOPE,
    callback: async (resp) => {
      if (resp.error) { toast("Google sign-in failed", true); return; }
      state._driveToken = resp.access_token;
      await setMeta("driveTokenExpiry", Date.now() + (resp.expires_in || 3600) * 1000);
      toast("Google Drive connected");
      render();
      flushSyncQueue();
    },
  });
}

function connectGoogleDrive() {
  if (!tokenClient) initGoogleAuth();
  if (!tokenClient) { toast("Add a Client ID first", true); return; }
  tokenClient.requestAccessToken();
}

async function driveFetch(url, opts = {}) {
  if (!state._driveToken) throw new Error("not-connected");
  return fetch(url, {
    ...opts,
    headers: { ...(opts.headers || {}), Authorization: `Bearer ${state._driveToken}` },
  });
}

async function findOrCreateSheet() {
  const cachedId = await getMeta("sheetId", null);
  if (cachedId) return cachedId;

  const q = encodeURIComponent(`name='${SHEET_NAME}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`);
  const listRes = await driveFetch(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`);
  const listJson = await listRes.json();
  if (listJson.files && listJson.files.length) {
    await setMeta("sheetId", listJson.files[0].id);
    return listJson.files[0].id;
  }

  const createRes = await driveFetch("https://sheets.googleapis.com/v4/spreadsheets", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ properties: { title: SHEET_NAME } }),
  });
  const createJson = await createRes.json();
  const sheetId = createJson.spreadsheetId;
  await setMeta("sheetId", sheetId);
  return sheetId;
}

// Makes sure the four tabs (Primary & Co-Surgeon, Assistant, All Cases,
// Monthly Tally) exist with headers + formulas in place. Runs once per
// spreadsheet (tracked via the "tabsReady" meta flag) and is safe to
// re-run — it only adds what's missing.
async function ensureTabs(sheetId) {
  const ready = await getMeta("tabsReady", false);
  if (ready) return;

  const metaRes = await driveFetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties`);
  const metaJson = await metaRes.json();
  const sheets = metaJson.sheets || [];
  const existingTitles = sheets.map((s) => s.properties.title);
  const requests = [];

  if (!existingTitles.includes(TAB_PRIMARY)) {
    // Rename Google's auto-created default tab (usually "Sheet1") to our
    // first data tab instead of leaving an empty orphan tab behind —
    // but only if that default tab isn't already one of ours.
    const firstSheet = sheets[0];
    if (firstSheet && !ALL_TABS.includes(firstSheet.properties.title)) {
      requests.push({
        updateSheetProperties: {
          properties: { sheetId: firstSheet.properties.sheetId, title: TAB_PRIMARY },
          fields: "title",
        },
      });
    } else {
      requests.push({ addSheet: { properties: { title: TAB_PRIMARY } } });
    }
  }
  [TAB_ASSISTANT, TAB_ALL, TAB_TALLY].forEach((title) => {
    if (!existingTitles.includes(title)) requests.push({ addSheet: { properties: { title } } });
  });

  if (requests.length) {
    await driveFetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}:batchUpdate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ requests }),
    });
  }

  // Header rows + the formulas that drive "All Cases" and "Monthly Tally".
  // USER_ENTERED (not RAW) so date strings like "8/27/2026" get parsed
  // into real Sheets dates — the tally's date-range math depends on that.
  const allDateRange = `'${TAB_ALL}'!A2:A5000`;
  await driveFetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchUpdate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      valueInputOption: "USER_ENTERED",
      data: [
        { range: `'${TAB_PRIMARY}'!A1:P1`, values: [SHEET_HEADER] },
        { range: `'${TAB_ASSISTANT}'!A1:P1`, values: [SHEET_HEADER] },
        { range: `'${TAB_ALL}'!A1:P1`, values: [SHEET_HEADER] },
        {
          // Stacks both role tabs into one sorted-by-date table so the
          // tally only ever has to read from one range.
          range: `'${TAB_ALL}'!A2`,
          values: [[`=SORT({'${TAB_PRIMARY}'!A2:P5000;'${TAB_ASSISTANT}'!A2:P5000},1,TRUE)`]],
        },
        { range: `'${TAB_TALLY}'!A1:E1`, values: [["Month", "Bariatric", "EGD", "General Surgery", "Total"]] },
        {
          // One first-of-month row per month that actually has a case,
          // newest formulas spill down automatically as rows are added.
          range: `'${TAB_TALLY}'!A2`,
          values: [[`=SORT(UNIQUE(FILTER(EOMONTH(${allDateRange},-1)+1,${allDateRange}<>"")))`]],
        },
        {
          range: `'${TAB_TALLY}'!B2`,
          values: [[`=ARRAYFORMULA(IF(A2:A="","",COUNTIFS('${TAB_ALL}'!$A$2:$A$5000,">="&A2:A,'${TAB_ALL}'!$A$2:$A$5000,"<"&EDATE(A2:A,1),'${TAB_ALL}'!$P$2:$P$5000,"Bariatric")))`]],
        },
        {
          range: `'${TAB_TALLY}'!C2`,
          values: [[`=ARRAYFORMULA(IF(A2:A="","",COUNTIFS('${TAB_ALL}'!$A$2:$A$5000,">="&A2:A,'${TAB_ALL}'!$A$2:$A$5000,"<"&EDATE(A2:A,1),'${TAB_ALL}'!$P$2:$P$5000,"EGD")))`]],
        },
        {
          range: `'${TAB_TALLY}'!D2`,
          values: [[`=ARRAYFORMULA(IF(A2:A="","",COUNTIFS('${TAB_ALL}'!$A$2:$A$5000,">="&A2:A,'${TAB_ALL}'!$A$2:$A$5000,"<"&EDATE(A2:A,1),'${TAB_ALL}'!$P$2:$P$5000,"General Surgery")))`]],
        },
        { range: `'${TAB_TALLY}'!E2`, values: [[`=ARRAYFORMULA(IF(A2:A="","",B2:B+C2:C+D2:D))`]] },
      ],
    }),
  });

  await setMeta("tabsReady", true);
}

async function ensureSpreadsheet() {
  const sheetId = await findOrCreateSheet();
  await ensureTabs(sheetId);
  return sheetId;
}

// Extracts the 1-based row number Sheets actually wrote to from an
// append response's updatedRange, e.g. "'Assistant'!A5:P5" -> 5.
function rowFromUpdatedRange(range) {
  if (!range) return null;
  const m = range.match(/![A-Z]+(\d+):/);
  return m ? Number(m[1]) : null;
}

async function syncCaseToDrive(c) {
  if (!state._driveToken) {
    if (!SYNC_QUEUE.includes(c.id)) SYNC_QUEUE.push(c.id);
    await setMeta("syncQueue", SYNC_QUEUE);
    return;
  }
  try {
    const sheetId = await ensureSpreadsheet();
    const tab = tabForRole(c.role);
    const row = caseToRow(c);

    if (c._sheetSync && c._sheetSync.tab === tab && c._sheetSync.row) {
      // Already has a row on the right tab (from an earlier save, or a
      // billed-status change) — update it in place rather than
      // appending a duplicate.
      await driveFetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/'${tab}'!A${c._sheetSync.row}:P${c._sheetSync.row}?valueInputOption=USER_ENTERED`,
        {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ values: [row] }),
        }
      );
    } else {
      if (c._sheetSync && c._sheetSync.tab && c._sheetSync.row) {
        // Role was changed after an earlier sync (e.g. edited from
        // "assistant" to "primary") — clear the stale row on the old
        // tab so the same case doesn't show up twice.
        await driveFetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/'${c._sheetSync.tab}'!A${c._sheetSync.row}:P${c._sheetSync.row}:clear`,
          { method: "POST" }
        );
      }
      const appendRes = await driveFetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/'${tab}'!A1:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ values: [row] }),
        }
      );
      const appendJson = await appendRes.json();
      const rowNum = rowFromUpdatedRange(appendJson.updates && appendJson.updates.updatedRange);
      c._sheetSync = { tab, row: rowNum };
      await saveCase(c); // persist the row pointer so later edits update instead of re-appending
    }
    toast("Synced to Google Sheets");
  } catch (err) {
    console.error(err);
    if (!SYNC_QUEUE.includes(c.id)) SYNC_QUEUE.push(c.id);
    await setMeta("syncQueue", SYNC_QUEUE);
    toast("Sync failed — queued, will retry", true);
  }
}

async function flushSyncQueue() {
  if (!SYNC_QUEUE.length || !state._driveToken) return;
  const queue = [...SYNC_QUEUE];
  SYNC_QUEUE = [];
  for (const id of queue) {
    const c = CASES.find((x) => x.id === id);
    if (c) await syncCaseToDrive(c);
  }
  await setMeta("syncQueue", SYNC_QUEUE);
}

window.addEventListener("online", flushSyncQueue);

// ---------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------

async function boot() {
  await loadCases();
  await loadCodes();
  state._clientId = (await getMeta("googleClientId", DEFAULT_GOOGLE_CLIENT_ID)) || DEFAULT_GOOGLE_CLIENT_ID;
  SYNC_QUEUE = (await getMeta("syncQueue", [])) || [];
  render();
  initGoogleAuth();

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch((err) => console.error("SW registration failed", err));
  }
}

boot();
