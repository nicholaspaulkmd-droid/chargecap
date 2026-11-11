// app.js — ChargeCap v2.2
// Vanilla JS, no framework. Uses idb-keyval (IndexedDB) for storage,
// Tesseract.js for on-device OCR, NLM ClinicalTables API for live ICD-10
// search, and MSAL.js + Microsoft Graph for sync to an Excel workbook in
// the practice's Microsoft 365 OneDrive (v2.0; v1.x used Google Sheets).

// ---------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------

// Microsoft 365 app registration IDs are in the sync section near the
// bottom of this file (M365_CLIENT_ID / M365_TENANT_ID).
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
// v1.13: consult-only cases (see caseCategory) are STORED on their own
// tab — they no longer appear on Primary/Assistant at all. All Cases and
// the Monthly Tally read from it. (v1.12 had briefly made this a FILTER
// view of the role tabs; that was never deployed.)
const TAB_CONSULTS = "Non-Op Consults";
const DATA_TABS = [TAB_PRIMARY, TAB_ASSISTANT, TAB_CONSULTS];
const ALL_TABS = [TAB_PRIMARY, TAB_ASSISTANT, TAB_ALL, TAB_TALLY, TAB_CONSULTS];

// "Surgeon" removed 2026-09-11 — redundant with Role (Primary/Co-surgeon/
// Assistant already identifies who's who). Every column from Role onward
// shifted one letter left as a result — this sheet is now 15 columns
// (A-O, Category now in O) instead of 16 (A-P, Category was in P). See
// the "O" column references in ensureTabs()/syncCaseToDrive() below.
const SHEET_HEADER = [
  "Date", "Patient", "MRN", "DOB", "Role", "Modifier", "Facility",
  "Billing Status", "Billed Date", "CPT Codes", "CPT Descriptions",
  "ICD-10 Codes", "ICD-10 Descriptions", "Notes", "Category",
];

// Case category, used only for the monthly tally. A case counts as
// Bariatric, EGD, or Back if ANY of its CPT codes match that list —
// checked in this order (Bariatric, then EGD, then Back), so a case
// with codes from more than one list still only counts once, under
// the first list it matches (e.g. a bariatric case that also includes
// an on-table EGD still counts as bariatric). Everything else falls
// through to General Surgery. Edit these lists if the code sets change.
// 43846/43843 (open bypass/sleeve) and 43848 (revision VBG) added 2026-09-23 —
// they were on the billing sheet but missing here, so they fell to General Surgery.
const BARIATRIC_CPT = new Set(["43633", "43860", "43644", "43846", "43659", "43775", "43843", "43845", "43774", "43848"]);
const EGD_CPT = new Set(["43266", "43235", "43239", "43245", "43247", "43233"]);
const BACK_CPT = new Set(["22558", "22585"]);
const TUMMY_TUCK_CPT = new Set(["15830", "15847"]);
// Consult / E&M codes (added 2026-09-23). A case whose CPT codes are ALL
// from this list means the patient was seen but not operated on — it's
// categorized "Non-Op Consult" and kept out of every surgical category
// (and out of the tally's Total). If any procedure code is also on the
// case, the consult code is ignored and the procedure decides the
// category as usual.
const CONSULT_CPT = new Set(["99221", "99222", "99223", "99232", "99238", "99252", "99253", "99254", "99255"]);
const NON_OP_CATEGORY = "Non-Op Consult";

function caseCategory(c) {
  // Normalize with .trim() — CPT entries transcribed from the source
  // Excel sheets have had stray whitespace before, and a bare string
  // mismatch here silently falls through to "General Surgery" with no
  // error, which is exactly the kind of miscategorization this
  // function exists to avoid.
  const codes = (c.cptCodes || []).map((x) => (x.code || "").trim());
  if (codes.length && codes.every((code) => CONSULT_CPT.has(code))) return NON_OP_CATEGORY;
  if (codes.some((code) => BARIATRIC_CPT.has(code))) return "Bariatric";
  if (codes.some((code) => EGD_CPT.has(code))) return "EGD";
  if (codes.some((code) => BACK_CPT.has(code))) return "Back";
  if (codes.some((code) => TUMMY_TUCK_CPT.has(code))) return "Tummy Tuck";
  return "General Surgery";
}

function tabForRole(role) {
  return role === "assistant" ? TAB_ASSISTANT : TAB_PRIMARY;
}

// The tab a case's row belongs on: consult-only cases go to the
// Non-Op Consults tab regardless of role; everything else by role.
function tabForCase(c) {
  return caseCategory(c) === NON_OP_CATEGORY ? TAB_CONSULTS : tabForRole(c.role);
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

// Codes that were in the ORIGINAL built-in favorites (data.js before
// FAVORITES_VERSION 2). Used once, when upgrading a phone's saved library
// to a new FAVORITES_VERSION: any saved entry whose code is NOT in this
// list (or is category "Custom") was added by the user in the app, so it's
// kept; everything else is replaced by the new data.js lists.
const LEGACY_SEED_CODES = {
  cpt: new Set(["15271", "15734", "34266", "36556", "38100", "38120", "43235", "43239", "43245", "43247", "43280", "43281", "43282", "43324", "43360", "43631 + 43659 + 44202", "43633", "43644", "43653", "43659", "43770", "43771", "43772", "43773", "43774", "43775", "43830", "43832", "43840", "43843", "43845", "43846", "43848", "43860", "44005", "44050", "44120", "44121", "44130", "44140", "44143", "44145", "44160", "44180", "44186", "44202", "44203", "44205", "44602", "44950", "44960", "44970", "47001", "47100", "47562", "47563", "47564", "47600", "47605", "47610", "49020", "49320", "49321", "49322", "49326", "49505", "49507", "49520", "49550", "49560", "49565", "49568", "49570", "49585", "49587", "49650", "49651", "49652", "49653", "49654", "49655", "49656", "49657", "58805", "60210", "60240", "60500", "64488", "99221", "99222", "99223", "99232", "99238", "99252", "99253", "99254", "99255"]),
  icd10: new Set(["D17", "D51.0", "D73.89", "E11.9", "E43", "E66.01", "E66.3", "E66.9", "E78.0", "F32.9", "G47.30", "I10", "I25.10", "I25.9", "I26.09", "I50.20", "I51.9", "I82.4", "I87.2", "J44.9", "J45.909", "K21.0", "K21.9", "K27.3", "K27.7", "K27.9", "K28", "K28.0", "K28.1", "K28.3", "K28.4", "K29.00", "K31.1", "K35.2", "K35.3", "K35.80", "K40.20", "K40.30", "K40.90", "K40.91", "K41.9", "K42.0", "K42.9", "K43.0", "K43.2", "K43.9", "K44.9", "K45.8", "K46.9", "K52.9", "K56.1", "K56.5", "K58.0", "K58.9", "K59.0", "K63.2", "K65.1", "K66.0", "K66.1", "K76.0", "K80.00", "K80.18", "K80.66", "K80.80", "K81.0", "K81.1", "K81.2", "K81.9", "K82.4", "K82.8", "K85.10", "K91.89", "L72.3", "L98.9", "M15.0", "M25.50", "M54.5", "N83.20", "R06.00", "R06.83", "R10.0", "R11.0", "R11.10", "R11.2", "R13.10", "R19.7", "S36.00XA", "Z30.2"]),
};

function mergeFavorites(type, saved, favorites) {
  const legacy = LEGACY_SEED_CODES[type];
  const fresh = favorites.map((c) => ({ ...c, id: uuid() }));
  const have = new Set(fresh.map((c) => `${c.code}|${c.category}`));
  const freshCodes = new Set(fresh.map((c) => c.code));
  const userAdded = (saved || []).filter((c) => {
    const code = (c.code || "").trim();
    const isUserAdded = c.category === "Custom" || !legacy.has(code);
    return isUserAdded && !freshCodes.has(code) && !have.has(`${code}|${c.category}`);
  });
  return [...userAdded, ...fresh];
}

async function loadCodes() {
  let cpt = await idbKeyval.get("cpt", codesStore);
  let icd10 = await idbKeyval.get("icd10", codesStore);
  const favVersion = await getMeta("favoritesVersion", 1);
  if (typeof FAVORITES_VERSION !== "undefined" && favVersion < FAVORITES_VERSION) {
    // Built-in favorites changed (e.g. billing sheet updated) — refresh the
    // saved copy, keeping user-added codes. Brand-new installs just seed.
    cpt = cpt ? mergeFavorites("cpt", cpt, CPT_FAVORITES) : null;
    icd10 = icd10 ? mergeFavorites("icd10", icd10, ICD10_FAVORITES) : null;
    if (cpt) await idbKeyval.set("cpt", cpt, codesStore);
    if (icd10) await idbKeyval.set("icd10", icd10, codesStore);
    await setMeta("favoritesVersion", FAVORITES_VERSION);
  }
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
  await queueCodeOp({ op: "add", type, code: clean.code, desc: clean.desc, category: clean.category });
  return clean;
}

async function updateCodeInLib(type, id, patch) {
  const lib = codeLib(type);
  const item = lib.find((c) => c.id === id);
  if (!item) return;
  const from = { code: item.code, category: item.category, desc: item.desc };
  Object.assign(item, patch);
  await persistCodeLib(type);
  await queueCodeOp({ op: "update", type, from, code: item.code, desc: item.desc, category: item.category });
}

async function deleteCodeFromLib(type, id) {
  const lib = codeLib(type);
  const idx = lib.findIndex((c) => c.id === id);
  if (idx < 0) return;
  const [gone] = lib.splice(idx, 1);
  await persistCodeLib(type);
  await queueCodeOp({ op: "delete", type, code: gone.code, category: gone.category, desc: gone.desc });
}

// v2.1: every library change is also logged as an operation and replayed
// onto the "Code Library" tab of the OneDrive workbook (syncCodeLibrary),
// so the library survives app updates, reinstalls, and a new phone.
async function queueCodeOp(op) {
  const ops = (await getMeta("codeOps", [])) || [];
  ops.push({ ...op, at: Date.now() });
  await setMeta("codeOps", ops);
  scheduleCodeLibrarySync();
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
  // v2.0: LOCAL date. toISOString() is UTC, so after 6pm in Utah it
  // returned tomorrow's date as the default date of service.
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
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
  liveScanOpen: false, // true while the live camera screen is open (see "Live camera scan" below)
  editingCaseId: null,
  ocrBusy: false,
  showRawOcr: false,
  // Last facility picked on the Capture screen (persisted in IndexedDB
  // meta "lastFacility") — every new case defaults to it, since a whole
  // operating day is usually at one location.
  lastFacility: null,
  // v1.12 sheet reconciliation (see reconcileWithSheet)
  missingIds: [],
  missingSelected: new Set(),
  reconcileBusy: false,
  reconcileCheckedAt: null,
  library: { type: "cpt", category: "All", search: "", editingId: null, adding: false, formCode: "", formDesc: "", formCategory: "" },
};

function newDraft() {
  return {
    id: uuid(),
    patientName: "",
    mrn: "",
    dob: "",
    dos: todayISO(),
    facility: state.lastFacility && FACILITIES.includes(state.lastFacility) ? state.lastFacility : FACILITIES[0],
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
    return await ocrExtract(file);
  } finally {
    state.ocrBusy = false;
    render();
  }
}

// Photos of an EHR screen are usually light text on a dark background
// (e.g. white-on-blue patient banner). Tesseract reads dark-on-light
// more reliably, so if the image is mostly dark, convert it to an
// inverted grayscale copy first. Stickers (dark text on a white label)
// are left untouched. Any failure just falls back to the original image.
const DARK_IMAGE_MEAN_MAX = 100; // 0-255 average brightness below which we invert
async function prepareForOcr(fileOrBlob) {
  try {
    if (typeof createImageBitmap !== "function") return fileOrBlob;
    const bmp = await createImageBitmap(fileOrBlob);
    const probe = document.createElement("canvas");
    probe.width = 64;
    probe.height = 32;
    const pctx = probe.getContext("2d", { willReadFrequently: true });
    pctx.drawImage(bmp, 0, 0, 64, 32);
    const px = pctx.getImageData(0, 0, 64, 32).data;
    let sum = 0;
    for (let i = 0; i < px.length; i += 4) sum += px[i] * 0.299 + px[i + 1] * 0.587 + px[i + 2] * 0.114;
    const mean = sum / (px.length / 4);
    if (mean >= DARK_IMAGE_MEAN_MAX) {
      if (bmp.close) bmp.close();
      return fileOrBlob;
    }
    const canvas = document.createElement("canvas");
    canvas.width = bmp.width;
    canvas.height = bmp.height;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    ctx.drawImage(bmp, 0, 0);
    if (bmp.close) bmp.close();
    // Pixel loop rather than ctx.filter — canvas filters aren't supported
    // on older iOS Safari and would silently do nothing there.
    const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const d = img.data;
    for (let i = 0; i < d.length; i += 4) {
      const v = 255 - ((d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) | 0);
      d[i] = d[i + 1] = d[i + 2] = v;
    }
    ctx.putImageData(img, 0, 0);
    return canvas;
  } catch (err) {
    console.error("OCR preprocessing skipped", err);
    return fileOrBlob;
  }
}

// v1.12: OCR + parse, with a second pass when the first one misses
// any field. Pass 2 uses prepareAdaptive() — every region is normalized
// to dark-text-on-light on its own, which helps a phone photo of a
// monitor where a light-on-dark banner sits inside a mostly-white
// screen (too bright overall for prepareForOcr to invert), or where
// glare washes out part of it. Pass 2 only fills fields pass 1 left
// blank; it never overrides a pass-1 read.
async function ocrExtract(fileOrBlob) {
  const worker = await getOcrWorker();
  const { data } = await worker.recognize(await prepareForOcr(fileOrBlob));
  const parsed = parseOcrText(data.text, data.words || []);
  parsed.rawText = data.text;
  if (parsed.patientName && parsed.mrn && parsed.dob) return parsed;
  try {
    const alt = await prepareAdaptive(fileOrBlob);
    if (!alt) return parsed;
    const { data: d2 } = await worker.recognize(alt);
    const p2 = parseOcrText(d2.text, d2.words || []);
    for (const f of ["patientName", "mrn", "dob"]) {
      if (!parsed[f] && p2[f]) {
        parsed[f] = p2[f];
        if (p2.lowConfidenceFields.includes(f) && !parsed.lowConfidenceFields.includes(f)) parsed.lowConfidenceFields.push(f);
      }
    }
    parsed.rawText = `${data.text}\n--- second pass ---\n${d2.text}`;
  } catch (err) {
    console.error("Second OCR pass skipped", err);
  }
  return parsed;
}

// Grayscale + local inversion: each pixel is compared with the average
// brightness of its neighborhood (box ~1/25 of the image width, via an
// integral image), and pixels in dark neighborhoods are inverted — so
// light-on-dark banners and dark-on-light text both end up dark-on-
// light. Large photos are scaled to <= 2000px wide first for speed.
async function prepareAdaptive(fileOrBlob) {
  if (typeof createImageBitmap !== "function") return null;
  const bmp = await createImageBitmap(fileOrBlob);
  const scale = Math.min(1, 2000 / bmp.width);
  const w = Math.max(1, Math.round(bmp.width * scale));
  const h = Math.max(1, Math.round(bmp.height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.drawImage(bmp, 0, 0, w, h);
  if (bmp.close) bmp.close();
  const img = ctx.getImageData(0, 0, w, h);
  const d = img.data;
  const gray = new Uint8ClampedArray(w * h);
  for (let i = 0, j = 0; j < gray.length; i += 4, j++) gray[j] = (d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) | 0;
  const integ = new Float64Array((w + 1) * (h + 1));
  for (let y = 1; y <= h; y++) {
    let rowSum = 0;
    for (let x = 1; x <= w; x++) {
      rowSum += gray[(y - 1) * w + (x - 1)];
      integ[y * (w + 1) + x] = integ[(y - 1) * (w + 1) + x] + rowSum;
    }
  }
  const r = Math.max(8, Math.round(w / 50));
  for (let y = 0; y < h; y++) {
    const y0 = Math.max(0, y - r), y1 = Math.min(h, y + r + 1);
    for (let x = 0; x < w; x++) {
      const x0 = Math.max(0, x - r), x1 = Math.min(w, x + r + 1);
      const sum = integ[y1 * (w + 1) + x1] - integ[y0 * (w + 1) + x1] - integ[y1 * (w + 1) + x0] + integ[y0 * (w + 1) + x0];
      const mean = sum / ((x1 - x0) * (y1 - y0));
      const g = gray[y * w + x];
      const v = mean < 128 ? 255 - g : g;
      const k = (y * w + x) * 4;
      d[k] = d[k + 1] = d[k + 2] = v;
    }
  }
  ctx.putImageData(img, 0, 0);
  return canvas;
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

// ---------------------------------------------------------------------
// Live camera scan
// ---------------------------------------------------------------------
//
// "Scan patient sticker" opens an in-page live camera preview with a
// pre-framed guide box. The user fits the sticker in the box and taps
// "Capture"; the photo is cropped to exactly what the box was framing
// and handed to OCR — no separate crop step needed.
//
// v1.11: the old automatic capture (a sharpness/steadiness heuristic
// that snapped the photo on its own) was removed at the user's request
// after real-world use — it wasn't reliable. Capture is now manual only;
// the guide box / auto-crop-to-box behavior is unchanged.
//
// `liveScan` is deliberately kept OUTSIDE `state`/render() while a scan
// is running: nothing should call render() while the camera is live,
// since a full innerHTML rebuild would tear down the <video> element and
// orphan the MediaStream. render() is only called to open the modal
// (before the camera is requested), and to close it.
const liveScan = {
  started: false, // guards against requesting the camera twice for one modal session
  stream: null,
  capturing: false, // true once Capture was tapped, to ignore double-taps while it wraps up
};

// iOS doesn't persist the camera-access grant for a home-screen web app
// the way it does for a regular Safari tab, so a fresh getUserMedia()
// call can re-trigger the "Allow Camera Access" system prompt. Fully
// releasing the camera after every single scan (the old behavior) meant
// that prompt reappeared for every patient, even back-to-back ones a
// few seconds apart. Instead, closeLiveScan() now keeps a just-used
// stream warm for CAMERA_IDLE_RELEASE_MS: openLiveScan() reuses it
// (skipping getUserMedia entirely, so no new prompt) if the next scan
// starts within that window, and releaseCamera() tears it down for real
// if nothing reuses it in time, or immediately if the app is backgrounded.
const CAMERA_IDLE_RELEASE_MS = 90000;
let idleReleaseTimer = null;

function releaseCamera() {
  if (idleReleaseTimer) {
    clearTimeout(idleReleaseTimer);
    idleReleaseTimer = null;
  }
  if (liveScan.stream) {
    liveScan.stream.getTracks().forEach((t) => t.stop());
    liveScan.stream = null;
  }
}

// Safety net: never let a warm-but-idle stream (or an actively-scanning
// one) keep the camera running once ChargeCap is backgrounded. iOS
// suspends it anyway, but releasing it ourselves keeps app state honest
// and avoids relying on the OS to clean up after us.
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) return;
  if (state.liveScanOpen) {
    closeLiveScan({ immediate: true });
    render();
  } else {
    releaseCamera();
  }
});

// Maps an on-screen element's rect (the guide box) to pixel coordinates
// in the VIDEO's native resolution, accounting for `object-fit: cover`
// (the preview fills its container and may crop/letterbox the actual
// feed) — mirrors what confirmCrop() above does for a plain <img>, just
// with the extra cover-vs-contain math a live video preview needs.
function getVideoContentRect(video) {
  const rect = video.getBoundingClientRect();
  const vw = video.videoWidth, vh = video.videoHeight;
  if (!vw || !vh) return rect;
  const elAspect = rect.width / rect.height;
  const vAspect = vw / vh;
  let contentW, contentH, offsetX, offsetY;
  if (vAspect > elAspect) {
    contentH = rect.height;
    contentW = rect.height * vAspect;
    offsetX = (rect.width - contentW) / 2;
    offsetY = 0;
  } else {
    contentW = rect.width;
    contentH = rect.width / vAspect;
    offsetX = 0;
    offsetY = (rect.height - contentH) / 2;
  }
  return { left: rect.left + offsetX, top: rect.top + offsetY, width: contentW, height: contentH };
}

function guideRectToVideoPixels(video, guideEl) {
  const content = getVideoContentRect(video);
  const g = guideEl.getBoundingClientRect();
  const scaleX = video.videoWidth / content.width;
  const scaleY = video.videoHeight / content.height;
  const sx = Math.max(0, Math.round((g.left - content.left) * scaleX));
  const sy = Math.max(0, Math.round((g.top - content.top) * scaleY));
  const sw = Math.max(1, Math.min(Math.round(g.width * scaleX), video.videoWidth - sx));
  const sh = Math.max(1, Math.min(Math.round(g.height * scaleY), video.videoHeight - sy));
  return { sx, sy, sw, sh };
}

// Opens the live-scan modal and requests camera access. Called right
// after the modal's first render (see bindLiveScanEvents), so the
// <video>/guide elements already exist by the time this runs.
async function openLiveScan() {
  // Yield one tick before doing anything that might call render() —
  // guarantees this never re-enters render()/bindEvents() from within
  // the very call stack that's still finishing binding them.
  await Promise.resolve();

  // Reuse a still-live stream left warm by a recent scan (see
  // closeLiveScan()/CAMERA_IDLE_RELEASE_MS above) instead of requesting
  // the camera again — this is what avoids re-triggering the OS
  // permission prompt for back-to-back patients.
  if (liveScan.stream && liveScan.stream.getVideoTracks().some((t) => t.readyState === "live")) {
    if (idleReleaseTimer) {
      clearTimeout(idleReleaseTimer);
      idleReleaseTimer = null;
    }
    const video = document.getElementById("liveScanVideo");
    if (!video) {
      releaseCamera();
      return;
    }
    video.srcObject = liveScan.stream;
    await video.play().catch(() => {});
    showScanReady();
    return;
  }

  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    fallbackToFilePicker("Live camera not supported on this browser — using your camera app instead");
    return;
  }

  try {
    const stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: "environment" }, width: { ideal: 1920 }, height: { ideal: 1080 } },
      audio: false,
    });
    if (!state.liveScanOpen) {
      // Modal was cancelled while the permission prompt was up.
      stream.getTracks().forEach((t) => t.stop());
      return;
    }
    const video = document.getElementById("liveScanVideo");
    if (!video) {
      stream.getTracks().forEach((t) => t.stop());
      return;
    }
    liveScan.stream = stream;
    video.srcObject = stream;
    await video.play().catch(() => {});
    showScanReady();
  } catch (err) {
    console.error("getUserMedia failed", err);
    fallbackToFilePicker("Live camera unavailable — using your camera app instead");
  }
}

function fallbackToFilePicker(message) {
  closeLiveScan();
  render();
  toast(message, true);
  const camInput = document.getElementById("camInput");
  if (camInput) camInput.click();
}

// Camera is live — prompt the user to frame the sticker and tap Capture.
function showScanReady() {
  liveScan.capturing = false;
  const statusEl = document.getElementById("scanStatus");
  if (statusEl) statusEl.textContent = "Fit the sticker in the box, then tap Capture";
}

// Crops the guide box's region out of the live video at full camera
// resolution (not the on-screen preview size) and hands it to OCR —
// the same shared runOcrAndApply() path the manual crop screen uses.
function captureLiveFrame(video, guideEl) {
  if (liveScan.capturing) return;
  liveScan.capturing = true;
  const statusEl = document.getElementById("scanStatus");
  if (statusEl) statusEl.textContent = "Captured!";

  const { sx, sy, sw, sh } = guideRectToVideoPixels(video, guideEl);
  const canvas = document.createElement("canvas");
  canvas.width = sw;
  canvas.height = sh;
  canvas.getContext("2d").drawImage(video, sx, sy, sw, sh, 0, 0, sw, sh);

  canvas.toBlob((blob) => {
    closeLiveScan();
    render();
    if (blob) runOcrAndApply(blob);
    else toast("Capture failed — try again", true);
  }, "image/jpeg", 0.92);
}

// Closes the live-scan modal. Every path that can end a live-scan
// session (Cancel, "use camera app instead", a capture, or
// getUserMedia failing) goes through this. It does NOT stop
// the camera stream itself by default — the stream is left warm for
// CAMERA_IDLE_RELEASE_MS so the next scan can reuse it without a new
// permission prompt (see openLiveScan()); pass {immediate: true} (used
// when the app is backgrounded) to release it right away instead.
function closeLiveScan({ immediate = false } = {}) {
  liveScan.capturing = false;
  liveScan.started = false;
  state.liveScanOpen = false;

  if (!liveScan.stream) return;
  if (immediate) {
    releaseCamera();
  } else if (!idleReleaseTimer) {
    idleReleaseTimer = setTimeout(releaseCamera, CAMERA_IDLE_RELEASE_MS);
  }
}

// Lines that are almost certainly NOT the patient name, used to filter
// candidate lines out of the name-guessing fallback below (hospital
// stickers are full of other short caps-heavy lines: facility name,
// "PATIENT LABEL", room/bed, barcode text, etc.).
const NAME_EXCLUDE_WORDS = /\b(DOB|MRN|DATE|FACILITY|HOSPITAL|ROOM|BED|ACCOUNT|ACCT|SURGEON|PHYSICIAN|DOCTOR|ADDRESS|PHONE|ADMIT|PATIENT LABEL|CONFIDENTIAL|SPECIMEN|ALLERGY|ALLERGIES)\b/i;

// A patient ID sticker sometimes also prints the ordering/attending
// physician's name in the same "Last, First" shape as the patient's —
// this disqualifies a match tied to that role so it isn't mistaken for
// the patient. This checks the matched line and the line right before
// it (OCR often splits a label onto its own line from the value beside
// it) for a role word that disqualifies the match.
const NAME_DISQUALIFY_CONTEXT = /\b(ATTENDING|ADMITTING|REFERRING|ORDERING|SURGEON|PHYSICIAN|PROVIDER|DOCTOR)\b/i;

// Some stickers print the attending physician as "LAST, MD, FIRST, M"
// (e.g. "PAULK, MD, NICHOLAS, J") — a credential sitting where a first
// name would in a plain "Last, First" match. Reject those.
// Words that show up as short Title-Case lines on EHR screen banners but
// are never a patient's name — keeps the "First M. Last" screen fallback
// in parseOcrText() from grabbing them.
const SCREEN_NON_NAME_WORDS = /\b(Male|Female|Bed|Room|Unit|Floor|Code|Status|Full|Allergies|Allergy|Isolation|Precautions|Admitted|Admission|Admit|Discharge|Location|Inpatient|Outpatient|Observation|Emergency|Surgery|Pre|Post|Op|Medical|Center|Hospital|Health|Clinic|Intermountain|Chart|Summary|Orders|Order|Results|Notes|Review|Medications?|Administration|Record|Problem|List|Vitals|Encounters?|Plan|History|Flowsheets?|Labs?|Imaging|Patient|Provider|Care|Team|Nursing|Documentation|Overview|Visit|Today|Schedule|Weight|Height|Attending|Primary|Insurance|Penicillin|Sulfa|None|Known|NKDA)\b/i;

// v1.12: text that, when it sits just before a "Word, Word" pair on the
// same line, means the pair is a list (allergies, diagnoses…) rather
// than "Last, First" — e.g. "Allergies: Penicillin, Sulfa".
const NAME_LIST_CONTEXT = /\b(ALLERG\w*|DX|DIAG\w*|PROBLEM\w*|HISTORY|HX|MEDS?|MEDICATIONS?|PROCEDURES?|ISOLATION|PRECAUTIONS?|CODE\s*STATUS)\b/i;
const NAME_PARTICLES = new Set(["de", "del", "della", "la", "las", "los", "van", "von", "der", "den", "da", "di", "du", "le", "st", "st.", "bin", "al", "y"]);

const CREDENTIAL_WORD = /^(MD|DO|PA|PA-C|NP|DPM|RN|CRNA|PHD)$/i;

// A date on a sticker that is NOT the birthdate — most commonly the
// sticker's own print timestamp — checked around a date match when no
// recognized birth-date label was found at all, so the generic "any
// date on the sticker" fallback below doesn't just grab whichever date
// happens to print first.
const DOB_DISQUALIFY_CONTEXT = /\b(PRINTED|CREATED)\b/i;

// Heuristic extraction of name / MRN / DOB from the raw OCR text of a
// patient ID sticker. Sticker layouts vary a lot by facility, so this
// looks for common label patterns and common date/ID shapes rather than
// assuming one fixed layout. Anything it can't find with confidence is
// left blank for manual entry, and fields that DID get a hit but from a
// low-confidence OCR read are flagged in lowConfidenceFields. The raw
// recognized text is always kept (see runOcr) so a scan that comes back
// wrong or incomplete can be inspected on-device via the "View scanned
// text" toggle in the Capture tab, without needing to share real
// patient data anywhere to debug it.
// v1.12 — see the screen-banner step in parseOcrText().
function nameSegment(line) {
  let seg = line.replace(/^[^A-Za-z]+/, ""); // leading icon/OCR junk ("★", "+", "%")
  const cut = seg.search(/\s{2,}|[|•·~©®]|\s[-–—]\s|\b(?:Male|Female|Sex|Age|DOB|D\.O\.B|MRN|Born|BD|Bed|Room)\b|\d/);
  if (cut >= 0) seg = seg.slice(0, cut);
  return seg.replace(/[^A-Za-z.]+$/, "").trim();
}

function looksLikeName(seg) {
  if (!seg || seg.length > 40) return false;
  if (NAME_EXCLUDE_WORDS.test(seg) || NAME_DISQUALIFY_CONTEXT.test(seg) || SCREEN_NON_NAME_WORDS.test(seg)) return false;
  const toks = seg.split(/\s+/);
  if (toks.length < 2 || toks.length > 5) return false;
  const WORD = /^(?:[A-Z][a-z]+|[A-Z][a-z]*['’][A-Z][a-z]+|Mc[A-Z][a-z]+|Mac[A-Z][a-z]+|[A-Z]{2,})(?:-(?:[A-Z][a-z]+|[A-Z][a-z]*['’][A-Z][a-z]+|[A-Z]{2,}))?$/;
  const INITIAL = /^[A-Z]\.?$/;
  let words = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    const inner = i > 0 && i < toks.length - 1;
    if (WORD.test(t)) { words++; continue; }
    if (INITIAL.test(t) && inner) continue;
    if (NAME_PARTICLES.has(t.toLowerCase()) && inner) continue;
    return false;
  }
  return words >= 2 && WORD.test(toks[0]) && WORD.test(toks[toks.length - 1]);
}

// "First M. de la Last" -> "de la Last, First M." (the app's format).
function flipToLastFirst(seg) {
  const toks = seg.split(/\s+/);
  let i = toks.length - 1;
  while (i - 1 > 0 && NAME_PARTICLES.has(toks[i - 1].toLowerCase())) i--;
  return `${toks.slice(i).join(" ")}, ${toks.slice(0, i).join(" ")}`;
}

function findScreenName(lines) {
  const DEMO_RE = /\b\d{1,3}\s*(?:yrs?\.?|y\.?[o0]\.?|years?)(?=\W|$)|\b(?:Male|Female|DOB|D\.O\.B|MRN)\b|\b\d{1,2}[\/\-]\d{1,2}[\/\-](?:19|20)\d{2}\b/i;
  const order = [];
  const demoIdx = lines.findIndex((l) => DEMO_RE.test(l));
  if (demoIdx >= 0) {
    order.push(demoIdx);
    if (demoIdx > 0) order.push(demoIdx - 1);
  }
  lines.forEach((_, i) => { if (!order.includes(i)) order.push(i); });
  for (const i of order) {
    const seg = nameSegment(lines[i]);
    if (looksLikeName(seg)) return flipToLastFirst(seg);
  }
  return null;
}

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
  const AGE_MARKER_RE = /\b\d{1,3}\s*(?:yrs?\.?|y\.?[o0]\.?|years?\s*old|years?)(?=\W|$)/i;
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
  // line for the first clean patient candidate, then finally a plain
  // 2-4 word ALL-CAPS line (e.g. "SMITH JOHN A") that isn't one of the
  // known non-name labels above — common on stickers with no comma in
  // the name. Both the label search and the line scan skip anything
  // tied to a physician/provider role (see NAME_DISQUALIFY_CONTEXT)
  // instead of just taking the first match blindly, since some
  // stickers also print the ordering/attending physician's name in the
  // same shape.
  let nameGuess = null;
  let nameFromExplicitLabel = false;
  for (const m of text.matchAll(/(?:Patient(?:\s*Name)?|Name)\s*[:\-]\s*([A-Za-z,'.\- ]{3,40})/gi)) {
    // Check a bit of text before the label too, not just the match
    // itself, in case a role word sits just in front of it.
    const context = text.slice(Math.max(0, m.index - 25), m.index + m[0].length);
    if (NAME_DISQUALIFY_CONTEXT.test(context)) continue;
    nameGuess = m[1].trim();
    nameFromExplicitLabel = true;
    break;
  }

  if (!nameGuess) {
    // Scan every clean "Last, First" pair found ANYWHERE in each line
    // (not just at the start — a label like "Name" or "Surgeon"
    // commonly sits before the name on the same OCR line rather than
    // on a line of its own) and take the FIRST one, in reading order:
    // skip a whole line when its own text names a physician/provider
    // role, and also pull in the line before it as context ONLY when
    // that previous line ITSELF ends with one of those role words
    // (optionally followed by a colon) — a label that got flattened
    // onto the tail of the previous line by OCR, with its value
    // continuing on this line (e.g. "Surgeon:" / next line "PAULK,
    // NICHOLAS"). Also skip a pair whose captured "first name" is
    // actually a credential (e.g. "PAULK, MD, NICHOLAS, J").
    const labelTailRe = new RegExp(NAME_DISQUALIFY_CONTEXT.source.replace(/^\\b\(/, "(") + "\\s*:?\\s*$", "i");
    outer: for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const prevLine = lines[i - 1] || "";
      const prevEndsWithLabel = labelTailRe.test(prevLine);
      const context = prevEndsWithLabel ? `${prevLine} ${line}` : line;
      if (NAME_DISQUALIFY_CONTEXT.test(context)) continue;
      for (const lf of line.matchAll(/([A-Z][A-Za-z'\-]+),\s*([A-Z][A-Za-z'\-]+)(?:\s+([A-Z])\b\.?)?/g)) {
        if (CREDENTIAL_WORD.test(lf[2])) continue;
        // v1.12: "Allergies: Penicillin, Sulfa" is a list, not a name.
        if (NAME_LIST_CONTEXT.test(line.slice(0, lf.index))) continue;
        if (SCREEN_NON_NAME_WORDS.test(lf[1]) || SCREEN_NON_NAME_WORDS.test(lf[2])) continue;
        nameGuess = lf[3] ? `${lf[1]}, ${lf[2]} ${lf[3]}` : `${lf[1]}, ${lf[2]}`;
        break outer;
      }
    }
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
  if (!nameGuess) {
    // EHR screen banner style (v1.11, reworked v1.12). The name has no
    // label and no comma, in normal "First M. Last" capitalization. v1.11
    // only accepted a line that was EXACTLY a name, which fails on the
    // very common one-row banner where OCR returns the whole row as one
    // line ("Lesley M. Webster Female 77 y.o. DOB: 7/23/1949 MRN: …"),
    // and on hyphenated / multi-part names ("Mary-Jo O'Neil-Garcia",
    // "Ana Lucia de la Cruz"). Now: take the text at the START of a line
    // up to the first sex/age/DOB/MRN token, digit or column gap, and
    // test that for a name shape. Lines are tried in order of how likely
    // they are to hold the name: the line with the DOB/age on it first,
    // then the line just above it, then everything else.
    nameGuess = findScreenName(lines);
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

  // Shown app-wide (except on Settings, which already has its own
  // connect control) whenever cases are stuck waiting to sync and we're
  // not currently signed in — a silent reauth already failed by the
  // time this shows (see ensureFreshToken), so this is the one-tap
  // fallback instead of having to notice sync stopped and go find
  // Settings → Connect on your own.
  const needsReconnect = !state._m365Token && SYNC_QUEUE.length && state.view !== "settings";
  const missingCount = (state.missingIds || []).length;
  const syncBanner = needsReconnect
    ? `<div class="sync-banner" id="syncBanner">⚠️ ${SYNC_QUEUE.length} case(s) waiting to sync — tap to reconnect Microsoft 365</div>`
    : missingCount && state.view !== "settings"
    ? `<div class="sync-banner" id="missingBanner">⚠️ ${missingCount} case(s) on this phone aren't in the workbook — tap to review</div>`
    : "";

  app.innerHTML = `<div class="screen">${syncBanner}${body}</div>${tab}${state.codePicker ? renderCodePicker() : ""}${state.cropSource ? renderCropModal() : ""}${state.liveScanOpen ? renderLiveScanModal() : ""}`;
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
        <button type="button" id="scanBtn" class="capture-btn" ${state.ocrBusy ? "disabled" : ""}>
          📷 ${state.ocrBusy ? "Reading sticker…" : "Scan patient sticker"}
        </button>
        <input id="camInput" type="file" accept="image/*" capture="environment" class="visually-hidden" ${state.ocrBusy ? "disabled" : ""} />
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
  return `
    <header class="topbar"><h1>Settings</h1></header>
    <div class="content">
      <section class="card">
        <h2>Microsoft 365 sync</h2>
        <p class="muted">Every saved case is written to the Excel workbook <strong>${WORKBOOK_FOLDER}/${WORKBOOK_FILE}</strong> in your OneDrive (practice Microsoft 365 account), newest cases at the top. Primary/co-surgeon, assistant and non-op consult cases each have their own tab, plus All Cases and a Monthly Tally. Data goes only to your practice's Microsoft 365. No other server is involved.</p>
        ${state._m365Account ? `<p class="hint">Signed in as <strong>${escapeHtml(state._m365Account)}</strong>${state._m365Token ? "" : " (sign-in expired; tap Reconnect)"}</p>` : ""}
        <button id="driveConnectBtn" class="secondary-btn">
          ${state._m365Account ? (state._m365Token ? "✓ Connected — reconnect" : "Reconnect Microsoft 365") : "Connect Microsoft 365"}
        </button>
        ${state._m365Account ? '<button id="signOutBtn" class="secondary-btn">Sign out</button>' : ""}
        ${SYNC_QUEUE.length ? `<p class="hint">${SYNC_QUEUE.length} case(s) queued to sync once connected/online.</p>` : ""}
      </section>

      <section class="card" id="missingCard">
        <h2>Missing from the workbook</h2>
        <p class="muted">Compares every case saved on this phone with the ${WORKBOOK_FILE} workbook and lists any that never made it there.</p>
        <button id="reconcileBtn" class="secondary-btn" ${state.reconcileBusy || !state._m365Account ? "disabled" : ""}>
          ${state.reconcileBusy ? "Checking…" : "Check for missing cases"}
        </button>
        ${(state.missingIds || []).length ? `
          <div class="missing-list">
            ${state.missingIds.map((id) => {
              const c = CASES.find((x) => x.id === id);
              if (!c) return "";
              return `<label class="missing-row">
                <input type="checkbox" data-missing-toggle="${c.id}" ${state.missingSelected.has(c.id) ? "checked" : ""} />
                <span><strong>${fmtDate(c.dos)}</strong> · ${escapeHtml(c.patientName) || "(no name)"} · ${escapeHtml(c.facility)} · ${escapeHtml(caseCategory(c))}${c.role === "assistant" ? " · Assist" : ""}</span>
              </label>`;
            }).join("")}
          </div>
          <button id="addMissingBtn" class="primary-btn" ${state.reconcileBusy || !state.missingSelected.size ? "disabled" : ""}>Add ${state.missingSelected.size} to workbook</button>
          <p class="hint">Uncheck anything you don't want in the workbook (e.g. old test cases).</p>
        ` : state.reconcileCheckedAt ? `<p class="hint">✓ Last check: everything on this phone is in the workbook.</p>` : ""}
      </section>

      <section class="card">
        <h2>Code library</h2>
        <p class="muted">${CPT_LIB.length} CPT · ${ICD10_LIB.length} ICD-10 codes. Changes are saved on this phone right away and backed up to the <strong>Code Library</strong> tab of your workbook, so app updates can't erase them. Codes you add on that tab in Excel show up here too.</p>
        <p class="hint">${state.codeLibPending ? `${state.codeLibPending} change(s) waiting to back up.` : state.codeLibSyncedAt ? `✓ Backed up ${new Date(state.codeLibSyncedAt).toLocaleString()}` : state._m365Account ? "Not backed up yet." : "Connect Microsoft 365 to back up your codes."}</p>
        <button class="secondary-btn" data-nav="library">Manage code library</button>
        <button id="codeSyncBtn" class="secondary-btn" ${state.codeLibBusy || !state._m365Account ? "disabled" : ""}>${state.codeLibBusy ? "Syncing…" : "Sync code library now"}</button>
      </section>

      <section class="card">
        <h2>Diagnostics</h2>
        <p class="muted">Tests each step of the Microsoft 365 connection. The results contain no patient information (only file names, tab names and row counts), so they're safe to copy and paste to whoever is helping you.</p>
        ${state.lastSyncError ? `<p class="hint"><strong>Last error</strong> (${escapeHtml(state.lastSyncError.where)}, ${new Date(state.lastSyncError.at).toLocaleString()}):<br><code class="small">${escapeHtml(state.lastSyncError.msg)}</code></p>` : ""}
        <button id="diagBtn" class="secondary-btn" ${state.diagBusy ? "disabled" : ""}>${state.diagBusy ? "Testing…" : "Run connection test"}</button>
        ${state.diagLines ? `<pre class="small" style="white-space:pre-wrap;word-break:break-word;">${escapeHtml(state.diagLines.join("\n"))}</pre>
        <button id="diagCopyBtn" class="secondary-btn">Copy results</button>` : ""}
      </section>

      <section class="card">
        <h2>About</h2>
        <p class="muted">ChargeCap v2.2 · ${CASES.length} case(s) stored locally on this device (IndexedDB).</p>
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
// lets the user drag a crop box down to just the sticker so stray text
// nearby (a watermark, another patient's sticker on the same sheet, a
// monitor showing a different chart) can't confuse the scan the way it
// did with the two example photos that had this issue. "Use full
// photo" skips straight to OCR on the original image for when a photo
// is already tightly framed.
function renderCropModal() {
  const { url } = state.cropSource;
  return `
    <div class="modal-backdrop">
      <div class="crop-modal" data-stop="1">
        <div class="crop-header">
          <h2>Crop to just the sticker</h2>
          <p class="crop-hint">Drag the corners to trim out anything that isn't the patient sticker — this keeps stray text nearby from confusing the scan.</p>
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

// Live in-page camera preview with the pre-framed guide box — see the
// "Live camera scan" section above. "Use camera app instead" is always
// offered as a manual escape hatch if the live preview misbehaves on a
// given phone.
function renderLiveScanModal() {
  return `
    <div class="modal-backdrop scan-backdrop">
      <div class="scan-modal" data-stop="1">
        <div class="scan-viewport">
          <video id="liveScanVideo" class="scan-video" autoplay playsinline muted></video>
          <div id="scanGuide" class="scan-guide"></div>
          <p id="scanStatus" class="scan-status">Starting camera…</p>
        </div>
        <div class="scan-actions">
          <button type="button" id="scanCaptureBtn" class="primary-btn">📷 Capture</button>
          <div class="scan-actions-row">
            <button type="button" id="scanUseFilePickerBtn" class="secondary-btn">Use camera app instead</button>
            <button type="button" id="scanCancelBtn" class="secondary-btn">Cancel</button>
          </div>
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
      // Defensive: the live-scan modal is a full-screen overlay so the
      // tab bar underneath shouldn't be reachable while it's open, but
      // never leave a camera stream running in the background either way.
      if (state.liveScanOpen) closeLiveScan();
      state.view = el.dataset.nav;
      state.codePicker = null;
      render();
    })
  );

  const syncBannerEl = document.getElementById("syncBanner");
  if (syncBannerEl) syncBannerEl.addEventListener("click", connectM365);
  const missingBannerEl = document.getElementById("missingBanner");
  if (missingBannerEl) missingBannerEl.addEventListener("click", () => {
    state.view = "settings";
    render();
    const card = document.getElementById("missingCard");
    if (card) card.scrollIntoView({ block: "start" });
  });

  if (state.view === "capture") bindCaptureEvents();
  if (state.view === "cases") bindCasesEvents();
  if (state.view === "settings") bindSettingsEvents();
  if (state.view === "library") bindLibraryEvents();
  if (state.codePicker) bindPickerEvents();
  if (state.cropSource) bindCropEvents();
  if (state.liveScanOpen) bindLiveScanEvents();
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

  const scanBtn = document.getElementById("scanBtn");
  if (scanBtn) {
    scanBtn.addEventListener("click", () => {
      if (state.ocrBusy) return;
      liveScan.started = false;
      state.liveScanOpen = true;
      render();
    });
  }

  const toggleRawBtn = document.querySelector("[data-toggle-raw-ocr]");
  if (toggleRawBtn) toggleRawBtn.addEventListener("click", () => {
    state.showRawOcr = !state.showRawOcr;
    render();
  });

  ["patientName", "mrn", "dob", "dos", "notes"].forEach((f) => {
    const el = document.getElementById("f_" + f);
    if (el) el.addEventListener("input", () => { d[f] = el.value; syncSaveButton(); });
  });
  const facilityEl = document.getElementById("f_facility");
  if (facilityEl) facilityEl.addEventListener("change", () => {
    d.facility = facilityEl.value;
    rememberFacility(d.facility);
  });

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
    rememberFacility(d.facility);
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

// Wires the live-scan modal's buttons and kicks off the actual camera
// request — but only the first time this modal is bound. bindEvents()
// re-runs after every render(), but nothing calls render() while a scan
// is in progress (see the "Live camera scan" section above), so
// in practice this only ever fires once per modal session; the
// `liveScan.started` guard just makes that explicit instead of relying
// on it.
function bindLiveScanEvents() {
  const modal = document.querySelector(".scan-modal[data-stop]");
  if (modal) modal.addEventListener("click", (e) => e.stopPropagation());

  const cancelBtn = document.getElementById("scanCancelBtn");
  if (cancelBtn) cancelBtn.addEventListener("click", () => { closeLiveScan(); render(); });

  const filePickerBtn = document.getElementById("scanUseFilePickerBtn");
  if (filePickerBtn) {
    filePickerBtn.addEventListener("click", () => {
      closeLiveScan();
      render();
      const camInput = document.getElementById("camInput");
      if (camInput) camInput.click();
    });
  }

  const captureBtn = document.getElementById("scanCaptureBtn");
  if (captureBtn) {
    captureBtn.addEventListener("click", () => {
      const video = document.getElementById("liveScanVideo");
      const guideEl = document.getElementById("scanGuide");
      if (video && guideEl && video.videoWidth) captureLiveFrame(video, guideEl);
    });
  }

  if (!liveScan.started) {
    liveScan.started = true;
    openLiveScan();
  }
}

function rememberFacility(f) {
  if (!f) return;
  state.lastFacility = f;
  setMeta("lastFacility", f).catch((err) => console.error("lastFacility save failed", err));
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
  const connectBtn = document.getElementById("driveConnectBtn");
  if (connectBtn) connectBtn.addEventListener("click", connectM365);
  const signOutBtn = document.getElementById("signOutBtn");
  if (signOutBtn) signOutBtn.addEventListener("click", signOutM365);
  const codeSyncBtn = document.getElementById("codeSyncBtn");
  if (codeSyncBtn) codeSyncBtn.addEventListener("click", () => syncCodeLibrary({ manual: true }));
  const reconcileBtn = document.getElementById("reconcileBtn");
  if (reconcileBtn) reconcileBtn.addEventListener("click", () => reconcileWithSheet({ manual: true }));
  const diagBtn = document.getElementById("diagBtn");
  if (diagBtn) diagBtn.addEventListener("click", runDiagnostics);
  const diagCopyBtn = document.getElementById("diagCopyBtn");
  if (diagCopyBtn) diagCopyBtn.addEventListener("click", () =>
    navigator.clipboard.writeText(diagReportText()).then(
      () => toast("Test results copied"),
      () => toast("Could not copy — select the text and copy it by hand", true)
    ));
  const addMissingBtn = document.getElementById("addMissingBtn");
  if (addMissingBtn) addMissingBtn.addEventListener("click", addMissingToSheet);
  app.querySelectorAll("[data-missing-toggle]").forEach((el) =>
    el.addEventListener("change", () => {
      const id = el.dataset.missingToggle;
      if (el.checked) state.missingSelected.add(id); else state.missingSelected.delete(id);
      render();
    })
  );
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
    return `${fmtDate(c.dos)} | ${c.patientName} | MRN ${c.mrn} | DOB ${fmtDate(c.dob)} | ${c.facility} | ${c.role}${c.modifier ? " mod " + c.modifier : ""} | CPT: ${cpt} | ICD-10: ${icd}${c.notes ? " | Notes: " + c.notes : ""}`;
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
    fmtDate(c.dos), c.patientName, c.mrn, fmtDate(c.dob), c.role, c.modifier, c.facility,
    // The sheet's own word for billed is "done" (also what the weekly
    // biller-email script writes), so write that rather than "billed".
    c.status === "billed" ? "done" : c.status, c.billedAt ? new Date(c.billedAt).toLocaleString() : "",
    c.cptCodes.map((x) => x.code).join("; "),
    c.cptCodes.map((x) => x.label || x.desc).join("; "),
    c.icd10Codes.map((x) => x.code).join("; "),
    c.icd10Codes.map((x) => x.label || x.desc).join("; "),
    c.notes,
    caseCategory(c),
  ];
}

// ---------------------------------------------------------------------
// Microsoft 365 sync (v2.0) — MSAL.js sign-in + Microsoft Graph Excel API
// ---------------------------------------------------------------------
// Every saved case is written to an Excel workbook in the user's own
// OneDrive for Business: /ChargeCap/ChargeCap Log.xlsx. This replaces
// the Google Identity Services + Sheets API code from v1.x so the PHI
// stays inside the practice's Microsoft 365 tenant (covered by its BAA).
//
// Row order (v2.0): NEWEST FIRST. A new case is inserted at its date
// position — above every row with the same or an older date — so the
// most recent cases are always at the top. Because inserts shift every
// row below, a stored row number is only a hint: each sync re-reads the
// tab and finds the case by date of service + MRN (or name), with the
// CPT list and the old row number breaking ties (see findCaseInTabs).

const M365_CLIENT_ID = "8296bb04-e9c9-482a-a6a8-9d141b25e85d";
const M365_TENANT_ID = "3602a33d-51c9-4967-bcfa-5dba0f27735b";
const GRAPH_SCOPES = ["User.Read", "Files.ReadWrite"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const WORKBOOK_FOLDER = "ChargeCap";
const WORKBOOK_FILE = "ChargeCap Log.xlsx";
// Blank workbook (tabs, headers, column formats, month-break formatting)
// shipped with the app. Used only if the workbook doesn't exist yet in
// OneDrive. Contains no patient data.
const WORKBOOK_TEMPLATE_URL = "chargecap-template.xlsx";

let msalApp = null;
let msalReady = null; // Promise, resolves once MSAL has processed any redirect

function initMsal() {
  if (msalReady) return msalReady;
  msalReady = (async () => {
    if (!window.msal) return false;
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: M365_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${M365_TENANT_ID}`,
        // Must match the SPA redirect URI registered in Entra exactly.
        redirectUri: location.origin + "/",
        navigateToLoginRequestUrl: false,
      },
      // localStorage so the sign-in survives the app being closed (an
      // iOS home-screen app is unloaded from memory constantly).
      cache: { cacheLocation: "localStorage" },
    });
    await msalApp.initialize();
    try {
      const resp = await msalApp.handleRedirectPromise();
      if (resp && resp.account) {
        msalApp.setActiveAccount(resp.account);
        state._m365Account = resp.account.username;
        toast("Microsoft 365 connected");
      }
    } catch (err) {
      console.error(err);
      toast("Microsoft sign-in didn't finish — tap Connect to try again", true);
    }
    const acct = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
    if (acct) {
      msalApp.setActiveAccount(acct);
      state._m365Account = acct.username;
    }
    return true;
  })();
  return msalReady;
}

// User-initiated connect/reconnect (Settings button, sync banner tap).
// Uses a full-page redirect rather than a popup: popups don't work
// reliably inside an iOS home-screen app (that is what broke the Google
// version's silent sign-in). The app reloads afterwards; every case is
// already saved in IndexedDB, so nothing in progress is lost.
async function connectM365() {
  if (!(await initMsal())) { toast("Microsoft sign-in isn't available offline", true); return; }
  const account = msalApp.getActiveAccount();
  const req = { scopes: GRAPH_SCOPES };
  if (account) await msalApp.acquireTokenRedirect({ ...req, account });
  else await msalApp.loginRedirect({ ...req, prompt: "select_account" });
}

async function signOutM365() {
  if (!(await initMsal())) return;
  const account = msalApp.getActiveAccount();
  // Clears this app's sign-in only — doesn't sign the phone out of
  // Microsoft everywhere, and doesn't touch any saved case.
  if (account) await msalApp.clearCache({ account });
  state._m365Token = null;
  state._m365Account = null;
  await setMeta("workbookItemId", null);
  toast("Signed out of Microsoft 365 on this app");
  render();
}

// Returns an access token without any UI, or null (caller queues the
// case and shows the reconnect banner). MSAL caches the ~1hr access
// token and refreshes it with its refresh token. Microsoft limits a web
// app's refresh token to 24 hours, so expect one "tap to reconnect" a
// day; that tap normally goes straight back without asking for a
// password.
const TOKEN_REQUEST_TIMEOUT_MS = 20000;
async function getGraphToken({ forceRefresh = false } = {}) {
  if (!(await initMsal())) return null;
  const account = msalApp.getActiveAccount();
  if (!account) return null;
  try {
    const resp = await Promise.race([
      msalApp.acquireTokenSilent({ scopes: GRAPH_SCOPES, account, forceRefresh }),
      new Promise((_, reject) => setTimeout(() => reject(new Error("token-timeout")), TOKEN_REQUEST_TIMEOUT_MS)),
    ]);
    state._m365Token = resp.accessToken;
    return resp.accessToken;
  } catch (err) {
    console.warn("Silent token failed", err && (err.errorCode || err.message));
    state._m365Token = null;
    return null;
  }
}

async function ensureFreshToken() {
  return !!(await getGraphToken());
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// fetch() wrapper for Graph: adds the token, retries once after a 401
// (with a forced token refresh) and up to three times on throttling /
// "workbook busy" responses, and throws on any other non-2xx so a failed
// write can never be mistaken for a success.
async function graphFetch(path, opts = {}, attempt = 0) {
  const token = await getGraphToken({ forceRefresh: attempt > 0 && opts._after401 });
  if (!token) throw new Error("not-connected");
  const url = path.startsWith("http") ? path : GRAPH + path;
  const res = await fetch(url, {
    method: opts.method || "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      ...(opts.body && !(opts.body instanceof Blob) && !(opts.body instanceof ArrayBuffer) ? { "Content-Type": "application/json" } : {}),
      ...(opts.headers || {}),
    },
    body: opts.body,
  });
  if (res.status === 401 && !opts._after401) return graphFetch(path, { ...opts, _after401: true }, attempt + 1);
  if ([429, 503, 504, 409].includes(res.status) && attempt < 3) {
    const wait = Math.min(Number(res.headers.get("Retry-After")) * 1000 || 1500 * (attempt + 1), 10000);
    await sleep(wait);
    return graphFetch(path, opts, attempt + 1);
  }
  if (!res.ok) {
    let detail = "";
    try { detail = JSON.stringify((await res.json()).error || {}); } catch { /* ignore */ }
    const err = new Error(`Graph ${res.status} on ${url.split("?")[0]}: ${detail}`);
    err.status = res.status;
    throw err;
  }
  return res;
}
const graphJson = async (path, opts) => (await graphFetch(path, opts)).json();

// Runs several Graph requests in one round trip, strictly in order
// (each depends on the one before). Throws if any of them failed.
async function graphBatch(reqs) {
  const requests = reqs.map((r, i) => ({
    id: String(i + 1),
    method: r.method,
    url: r.url,
    ...(r.body !== undefined ? { body: r.body, headers: { "Content-Type": "application/json" } } : {}),
    ...(i ? { dependsOn: [String(i)] } : {}),
  }));
  const json = await graphJson("/$batch", { method: "POST", body: JSON.stringify({ requests }) });
  const byId = Object.fromEntries((json.responses || []).map((r) => [r.id, r]));
  for (const r of requests) {
    const resp = byId[r.id];
    if (!resp || resp.status < 200 || resp.status >= 300) {
      throw new Error(`Graph batch step ${r.id} (${r.method} ${r.url.split("?")[0]}) failed: ${resp ? resp.status + " " + JSON.stringify(resp.body && resp.body.error || {}) : "no response"}`);
    }
  }
  return requests.map((r) => byId[r.id].body);
}

// ---- Workbook location ----------------------------------------------

async function findOrCreateWorkbook() {
  const cachedId = await getMeta("workbookItemId", null);
  if (cachedId) return cachedId;
  const path = `/me/drive/root:/${encodeURIComponent(WORKBOOK_FOLDER)}/${encodeURIComponent(WORKBOOK_FILE)}`;
  let item = null;
  try {
    item = await graphJson(`${path}?$select=id,name`);
  } catch (err) {
    if (err.status !== 404) throw err;
  }
  if (!item) {
    // Not there yet: upload the blank template (creates the ChargeCap
    // folder too). conflictBehavior=fail so this can never overwrite a
    // workbook that appeared in the meantime.
    const tpl = await fetch(WORKBOOK_TEMPLATE_URL, { cache: "no-store" });
    if (!tpl.ok) throw new Error("Couldn't load the blank workbook template");
    item = await graphJson(`${path}:/content?@microsoft.graph.conflictBehavior=fail`, {
      method: "PUT",
      body: await tpl.arrayBuffer(),
      headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
    });
  }
  await setMeta("workbookItemId", item.id);
  return item.id;
}

const wbPath = (itemId) => `/me/drive/items/${itemId}/workbook`;
// v2.2: worksheets are addressed by their ID once known, not by name.
// "Primary & Co-Surgeon" contains an "&", and a name in the URL is the
// fragile part; the ID ({00000000-0001-...}) has no special characters.
let SHEET_IDS = { itemId: null, byName: {} };
const wsPathByName = (itemId, tab) => `${wbPath(itemId)}/worksheets('${encodeURIComponent(tab.replace(/'/g, "''"))}')`;
const wsPath = (itemId, tab) => {
  const id = SHEET_IDS.itemId === itemId && SHEET_IDS.byName[tab];
  return id ? `${wbPath(itemId)}/worksheets/${encodeURIComponent(id)}` : wsPathByName(itemId, tab);
};
async function loadSheetIds(itemId) {
  const sheets = await graphJson(`${wbPath(itemId)}/worksheets?$select=id,name`);
  const list = sheets.value || [];
  SHEET_IDS = { itemId, byName: Object.fromEntries(list.map((x) => [x.name, x.id])) };
  return list;
}
const rangePath = (itemId, tab, addr) => `${wsPath(itemId, tab)}/range(address='${addr}')`;

// ---- Formulas (All Cases + Monthly Tally) ---------------------------
// The formulas live IN the workbook file (ChargeCap Log.xlsx and
// chargecap-template.xlsx are built with them), so the app normally
// never writes them. They are only (re)written here if one of those two
// tabs is missing, e.g. deleted by hand. Keep these in step with
// build_workbook.py.
const WB_SCHEMA_VERSION = 3;
// Every reference into a data tab starts at ROW 1 (the header): the app
// inserts new cases at row 2, and Excel shifts (rather than grows) a
// range whose first row is the insert point.
const FORMULA_ROWS = 20000;
const TALLY_ROWS = 240; // 20 years of months

function q(tab) { return `'${tab.replace(/'/g, "''")}'`; }

// All Cases (ordinary formulas, no dynamic arrays, so it works in any
// Excel): merges the three newest-first data tabs row by row. Hidden
// columns: Q/R/S = next unread row on each tab, T = tab chosen for this
// row, U/V/W = those rows' dates; Q1/R1/S1 = rows on each tab.
const ALL_CASES_ROWS = 5000;
function allCasesCounts() {
  const N = FORMULA_ROWS;
  return [q(TAB_PRIMARY), q(TAB_ASSISTANT), q(TAB_CONSULTS)].map(
    (t) => `=IFERROR(LOOKUP(2,1/(${t}!$A$1:$A$${N}<>""),ROW(${t}!$A$1:$A$${N})-1),0)`
  );
}
// Returns 23 formulas for columns A..W (P, index 15, left empty).
function allCasesRowFormulas(r) {
  const N = FORMULA_ROWS;
  const tabs = [q(TAB_PRIMARY), q(TAB_ASSISTANT), q(TAB_CONSULTS)];
  const out = new Array(23).fill("");
  "QRS".split("").forEach((col, k) => {
    out[16 + k] = r === 2 ? "=1" : `=${col}${r - 1}+($T${r - 1}=${k + 1})`;
    out[20 + k] = `=IF(${col}${r}>$${col}$1,-1,N(INDEX(${tabs[k]}!$A$1:$A$${N},${col}${r}+1)))`;
  });
  out[19] = `=IF(MAX(U${r}:W${r})<0,0,MATCH(MAX(U${r}:W${r}),U${r}:W${r},0))`;
  "ABCDEFGHIJKLMNO".split("").forEach((L, c) => {
    const cell = `INDEX(CHOOSE($T${r},${tabs[0]}!${L}$1:${L}$${N},${tabs[1]}!${L}$1:${L}$${N},${tabs[2]}!${L}$1:${L}$${N}),CHOOSE($T${r},$Q${r},$R${r},$S${r})+1)`;
    out[c] = `=IF($T${r}=0,"",IF(${cell}="","",${cell}))`;
  });
  return out;
}

// Monthly Tally row r (ordinary formulas, no spill): newest month first,
// back to the oldest month with a Primary/Co-Surgeon or Non-Op case.
function tallyRowFormulas(r) {
  const N = FORMULA_ROWS;
  const [P, C] = [q(TAB_PRIMARY), q(TAB_CONSULTS)];
  const dates = `${P}!$A$1:$A$${N},${C}!$A$1:$A$${N}`;
  const a = r === 2
    ? `=IF(COUNT(${dates})=0,"",EOMONTH(MAX(${dates}),-1)+1)`
    : `=IF(A${r - 1}="","",IF(EDATE(A${r - 1},-1)<EOMONTH(MIN(${dates}),-1)+1,"",EDATE(A${r - 1},-1)))`;
  const pa = `${P}!$A$1:$A$${N}`, po = `${P}!$O$1:$O$${N}`, na = `${C}!$A$1:$A$${N}`;
  const cnt = (cat) => `=IF($A${r}="","",COUNTIFS(${pa},">="&$A${r},${pa},"<"&EDATE($A${r},1),${po},"${cat}"))`;
  return [a, cnt("Bariatric"), cnt("EGD"), cnt("Back"), cnt("General Surgery"), cnt("Tummy Tuck"),
    `=IF($A${r}="","",SUM(B${r}:F${r}))`,
    `=IF($A${r}="","",COUNTIFS(${na},">="&$A${r},${na},"<"&EDATE($A${r},1)))`];
}

async function ensureWorkbookSchema(itemId) {
  const ready = await getMeta("wbSchemaReady", null);
  if (ready && ready.itemId === itemId && ready.version >= WB_SCHEMA_VERSION) return;
  const sheets = await graphJson(`${wbPath(itemId)}/worksheets?$select=name`);
  const names = (sheets.value || []).map((s) => s.name);
  const missing = ALL_TABS.filter((t) => !names.includes(t));
  const reqs = [];
  for (const tab of missing) reqs.push({ method: "POST", url: `${wbPath(itemId)}/worksheets/add`, body: { name: tab } });
  for (const tab of [...DATA_TABS, TAB_ALL]) {
    if (missing.includes(tab)) reqs.push({ method: "PATCH", url: rangePath(itemId, tab, "A1:O1"), body: { values: [SHEET_HEADER] } });
  }
  if (missing.includes(TAB_ALL)) {
    reqs.push({ method: "PATCH", url: rangePath(itemId, TAB_ALL, "Q1:S1"), body: { formulas: [allCasesCounts()] } });
    // 5,000 rows of formulas, sent 250 rows per request.
    for (let r0 = 2; r0 < ALL_CASES_ROWS + 2; r0 += 250) {
      const rows = [];
      for (let r = r0; r < Math.min(r0 + 250, ALL_CASES_ROWS + 2); r++) rows.push(allCasesRowFormulas(r));
      reqs.push({ method: "PATCH", url: rangePath(itemId, TAB_ALL, `A${r0}:W${r0 + rows.length - 1}`), body: { formulas: rows } });
    }
    reqs.push({ method: "PATCH", url: rangePath(itemId, TAB_ALL, "P:W"), body: { columnHidden: true } });
  }
  if (missing.includes(TAB_TALLY)) {
    reqs.push({
      method: "PATCH",
      url: rangePath(itemId, TAB_TALLY, "A1:H1"),
      body: { values: [["Month", "Bariatric", "EGD", "Back", "General Surgery", "Tummy Tuck", "Total", "Non-Op Consults"]] },
    });
    const rows = [];
    for (let r = 2; r < TALLY_ROWS + 2; r++) rows.push(tallyRowFormulas(r));
    reqs.push({ method: "PATCH", url: rangePath(itemId, TAB_TALLY, `A2:H${TALLY_ROWS + 1}`), body: { formulas: rows } });
    reqs.push({ method: "PATCH", url: rangePath(itemId, TAB_TALLY, `A2:A${TALLY_ROWS + 1}`), body: { numberFormat: Array.from({ length: TALLY_ROWS }, () => ["mmm yyyy"]) } });
  }
  for (let i = 0; i < reqs.length; i += 20) await graphBatch(reqs.slice(i, i + 20));
  await setMeta("wbSchemaReady", { itemId, version: WB_SCHEMA_VERSION });
}

async function ensureWorkbook() {
  let itemId = await findOrCreateWorkbook();
  try {
    // One cheap request per sync. It also proves the stored file ID still
    // points at a real workbook: in v2.1 a replaced/re-uploaded file left
    // the phone asking for the old ID forever.
    await loadSheetIds(itemId);
  } catch (err) {
    if (err.status !== 404 && !/itemNotFound/i.test(err.message)) throw err;
    await setMeta("workbookItemId", null);
    await setMeta("wbSchemaReady", null);
    itemId = await findOrCreateWorkbook();
    await loadSheetIds(itemId);
  }
  await ensureWorkbookSchema(itemId);
  if (ALL_TABS.some((t) => !SHEET_IDS.byName[t])) await loadSheetIds(itemId);
  return itemId;
}

// ---- Cell values ----------------------------------------------------

// Excel date serial (days since 1899-12-30) from "YYYY-MM-DD".
function excelSerialFromISO(iso) {
  const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 + 25569;
}
function excelSerialFromTimestamp(ms) {
  if (!ms) return null;
  const local = ms - new Date(ms).getTimezoneOffset() * 60000;
  return Math.round((local / 86400000 + 25569) * 100000) / 100000;
}

// Column number formats: A Date, D DOB, I Billed Date are real Excel
// dates; everything else is text ("@") so MRNs keep leading zeros and
// codes like "99222-57" are never reinterpreted.
const ROW_NUMBER_FORMAT = SHEET_HEADER.map((_, i) => (i === 0 || i === 3 ? "m/d/yyyy" : i === 8 ? "m/d/yyyy h:mm AM/PM" : "@"));

function caseToExcelRow(c) {
  const row = caseToRow(c).map((v) => (v == null ? "" : String(v)));
  const dos = excelSerialFromISO(c.dos);
  const dob = excelSerialFromISO(c.dob);
  const billed = c.status === "billed" ? excelSerialFromTimestamp(c.billedAt) : null;
  row[0] = dos != null ? dos : row[0];
  row[3] = dob != null ? dob : row[3];
  row[8] = billed != null ? billed : "";
  return row;
}

// ---- Reading the data tabs ------------------------------------------

function rowFromAddress(addr) {
  const m = String(addr || "").match(/![A-Z]+(\d+)/);
  return m ? Number(m[1]) : 1;
}

// Returns { [tab]: { rows: [...values], lastRow } } where rows[i] is
// sheet row i + 2 (row 1 = header).
async function readDataTabs(itemId) {
  const bodies = await graphBatch(
    DATA_TABS.map((tab) => ({ method: "GET", url: `${wsPath(itemId, tab)}/usedRange(valuesOnly=true)?$select=address,values` }))
  ).catch(async () => {
    // $batch requests don't depend on each other here; fall back to one
    // at a time if the batch itself failed for an unrelated reason.
    const out = [];
    for (const tab of DATA_TABS) out.push(await graphJson(`${wsPath(itemId, tab)}/usedRange(valuesOnly=true)?$select=address,values`));
    return out;
  });
  const out = {};
  DATA_TABS.forEach((tab, i) => {
    const body = bodies[i] || {};
    const start = rowFromAddress(body.address);
    const values = body.values || [];
    const rows = [];
    values.forEach((r, j) => {
      const rowNum = start + j;
      if (rowNum >= 2) rows[rowNum - 2] = r;
    });
    for (let k = 0; k < rows.length; k++) if (!rows[k]) rows[k] = [];
    out[tab] = { rows, lastRow: Math.max(1, start + values.length - 1) };
  });
  return out;
}

function sheetDateToISO(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number" || /^\d{5}(\.\d+)?$/.test(String(v).trim())) {
    const d = new Date(Math.round((Number(v) - 25569) * 86400000));
    return isNaN(d) ? "" : d.toISOString().slice(0, 10);
  }
  const str = String(v).trim();
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  m = str.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})/);
  if (m) {
    const y = m[3].length === 2 ? "20" + m[3] : m[3];
    return `${y}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`;
  }
  return str;
}
const normName = (n) => String(n || "").toLowerCase().replace(/[^a-z]/g, "");
const normMrn = (n) => String(n || "").replace(/\D/g, "").replace(/^0+/, "");
const normCpt = (s) => String(s || "").replace(/\s/g, "");

function caseMatchesRow(c, row) {
  if (!row || !row.length) return false;
  if (sheetDateToISO(row[0]) !== c.dos) return false;
  const cm = normMrn(c.mrn), rm = normMrn(row[2]);
  if (cm && rm) return cm === rm;
  const cn = normName(c.patientName), rn = normName(row[1]);
  if (cn && rn) return cn === rn;
  return normCpt(row[9]) === c.cptCodes.map((x) => x.code).join(";");
}

// Finds the row holding case c. Several rows can share a date + MRN (two
// operations on one patient the same day), so among the candidates
// prefer: same CPT list, then the row the case was last written to, then
// the tab the case belongs on. `claimed` (a Set of "tab|row") keeps two
// phone cases from both mapping to the same row.
function findCaseInTabs(c, tabs, claimed) {
  const target = tabForCase(c);
  const cpt = c.cptCodes.map((x) => x.code).join(";");
  let best = null, bestScore = -1;
  for (const t of DATA_TABS) {
    const rows = (tabs[t] && tabs[t].rows) || [];
    for (let i = 0; i < rows.length; i++) {
      if (!caseMatchesRow(c, rows[i])) continue;
      const key = `${t}|${i + 2}`;
      if (claimed && claimed.has(key)) continue;
      let score = 0;
      if (normCpt(rows[i][9]) === cpt) score += 4;
      if (c._sheetSync && c._sheetSync.tab === t && c._sheetSync.row === i + 2) score += 2;
      if (t === target) score += 1;
      if (score > bestScore) { best = { tab: t, row: i + 2 }; bestScore = score; }
    }
  }
  return best;
}

// The weekly biller report marks rows "done". If the workbook says done
// (or billed) and the phone still says pending, the phone adopts it, so
// a later re-sync can't flip the row back to "pending".
async function adoptSheetStatus(c, row) {
  if (!row || c.status !== "pending") return false;
  if (!/^(done|billed)$/i.test(String(row[7] || "").trim())) return false;
  c.status = "billed";
  c.billedAt = Date.now();
  await saveCase(c);
  return true;
}

// Row a new case should be inserted at so the tab stays newest-first:
// directly above the first row dated the same day or earlier. If every
// row is newer, it goes after the last row.
function insertRowFor(c, tabData) {
  const serial = excelSerialFromISO(c.dos);
  const rows = (tabData && tabData.rows) || [];
  if (serial != null) {
    for (let i = 0; i < rows.length; i++) {
      const v = rows[i] && rows[i][0];
      const n = typeof v === "number" ? v : excelSerialFromISO(sheetDateToISO(v));
      if (n != null && n <= serial) return { row: i + 2, insert: true };
    }
  }
  return { row: Math.max(2, ((tabData && tabData.lastRow) || 1) + 1), insert: false };
}

async function insertCaseRow(itemId, tab, c, tabData) {
  const { row, insert } = insertRowFor(c, tabData);
  const addr = `A${row}:O${row}`;
  const url = rangePath(itemId, tab, addr);
  const reqs = [];
  if (insert) reqs.push({ method: "POST", url: `${url}/insert`, body: { shift: "Down" } });
  // A freshly inserted row copies the formatting of the row above it
  // (the bold header, when it's row 2) — reset it to plain data style.
  reqs.push({ method: "POST", url: `${url}/format/fill/clear`, body: {} });
  reqs.push({ method: "PATCH", url: `${url}/format/font`, body: { bold: false, italic: false, color: "#000000" } });
  reqs.push({ method: "PATCH", url, body: { numberFormat: [ROW_NUMBER_FORMAT] } });
  reqs.push({ method: "PATCH", url, body: { values: [caseToExcelRow(c)] } });
  await graphBatch(reqs);
  return { tab, row };
}

async function updateCaseRow(itemId, tab, rowNum, c) {
  await graphFetch(rangePath(itemId, tab, `A${rowNum}:O${rowNum}`), {
    method: "PATCH",
    body: JSON.stringify({ values: [caseToExcelRow(c)] }),
  });
}

async function deleteCaseRow(itemId, tab, rowNum) {
  await graphFetch(`${rangePath(itemId, tab, `A${rowNum}:O${rowNum}`)}/delete`, {
    method: "POST",
    body: JSON.stringify({ shift: "Up" }),
  });
}

// ---- Sync queue -----------------------------------------------------
// A case is queued BEFORE any network work and only removed once the
// workbook write succeeds, so a failed or interrupted sync is retried.
const syncInFlight = new Set();
async function enqueueSync(id) {
  if (!SYNC_QUEUE.includes(id)) SYNC_QUEUE.push(id);
  await setMeta("syncQueue", SYNC_QUEUE);
}
async function dequeueSync(id) {
  SYNC_QUEUE = SYNC_QUEUE.filter((x) => x !== id);
  await setMeta("syncQueue", SYNC_QUEUE);
}

// v2.2: keep the real error so it can be seen in Settings → Diagnostics
// (it used to go only to the browser console, which a phone can't show).
async function recordSyncError(where, err) {
  const msg = String((err && (err.message || err.errorCode)) || err || "unknown error").slice(0, 600);
  state.lastSyncError = { where, at: Date.now(), msg };
  try { await setMeta("lastSyncError", state.lastSyncError); } catch { /* ignore */ }
}

async function syncCaseToDrive(c, { quiet = false } = {}) {
  await enqueueSync(c.id);
  if (syncInFlight.has(c.id)) return;
  syncInFlight.add(c.id);
  try {
    await syncCaseInner(c, { quiet });
  } finally {
    syncInFlight.delete(c.id);
  }
}

async function syncCaseInner(c, { quiet }) {
  if (!(await ensureFreshToken())) {
    render(); // reconnect banner
    return;
  }
  try {
    const itemId = await ensureWorkbook();
    const tab = tabForCase(c);
    const tabs = await readDataTabs(itemId);
    const found = findCaseInTabs(c, tabs);
    if (found) await adoptSheetStatus(c, tabs[found.tab].rows[found.row - 2]);

    if (found && found.tab === tab) {
      await updateCaseRow(itemId, tab, found.row, c);
      c._sheetSync = found;
    } else {
      if (found) {
        // The case now belongs on another tab (role changed, or it became
        // / stopped being consult-only): remove its old row first.
        await deleteCaseRow(itemId, found.tab, found.row);
      }
      // Re-read the target tab if a delete just shifted rows on it.
      const fresh = found && found.tab === tab ? await readDataTabs(itemId) : tabs;
      c._sheetSync = await insertCaseRow(itemId, tab, c, fresh[tab]);
    }
    await saveCase(c);
    await dequeueSync(c.id);
    if (!quiet) toast("Saved to Excel (OneDrive)");
    render();
  } catch (err) {
    console.error(err);
    await recordSyncError("Saving a case", err);
    if (!quiet) toast("Sync failed — queued, will retry (details: Settings → Diagnostics)", true);
  }
}

// ---- Diagnostics (v2.2) ---------------------------------------------
// Runs every step of the sync path one at a time and reports which one
// fails, with Microsoft's own error text. No PHI: names, sizes, counts.
function diagReportText() {
  const lines = [...(state.diagLines || [])];
  if (state.lastSyncError) lines.push("", `Last error (${state.lastSyncError.where}, ${new Date(state.lastSyncError.at).toISOString()}): ${state.lastSyncError.msg}`);
  return lines.join("\n");
}

async function runDiagnostics() {
  const out = [`ChargeCap v2.2 connection test — ${new Date().toLocaleString()}`];
  const ok = (m) => out.push("✓ " + m);
  const bad = (m, err) => out.push("✗ " + m + (err ? ` — ${String(err.message || err).slice(0, 400)}` : ""));
  const show = () => { state.diagLines = [...out]; render(); };
  state.diagBusy = true; show();
  try {
    // 1. Sign-in
    if (!(await initMsal())) { bad("Microsoft sign-in library didn't load (offline?)"); return; }
    const acct = msalApp.getActiveAccount();
    if (!acct) { bad("Not signed in — tap Connect Microsoft 365"); return; }
    ok(`Signed in as ${acct.username} (tenant ${acct.tenantId || "?"})`);
    if (acct.tenantId && acct.tenantId !== M365_TENANT_ID) bad(`That account is in a different tenant than the app (${M365_TENANT_ID})`);
    const token = await getGraphToken({ forceRefresh: true });
    if (!token) { bad("Couldn't get an access token — tap Reconnect"); return; }
    try {
      const claims = JSON.parse(atob(token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
      ok(`Token scopes: ${claims.scp || "(none)"}`);
    } catch { ok("Got an access token"); }
    show();

    // 2. OneDrive + folder contents
    try {
      const drive = await graphJson("/me/drive?$select=driveType,webUrl");
      ok(`OneDrive reachable (${drive.driveType})`);
    } catch (err) { bad("Couldn't open your OneDrive", err); return; }
    try {
      const kids = await graphJson(`/me/drive/root:/${encodeURIComponent(WORKBOOK_FOLDER)}:/children?$select=name,size,lastModifiedDateTime,file`);
      const files = (kids.value || []).map((f) => `${f.name} (${Math.round((f.size || 0) / 1024)} KB, modified ${f.lastModifiedDateTime})`);
      ok(`Files in /${WORKBOOK_FOLDER}: ${files.length ? files.join("; ") : "(empty)"}`);
    } catch (err) { bad(`Couldn't list the /${WORKBOOK_FOLDER} folder`, err); }
    show();

    // 3. The workbook file
    let item;
    try {
      item = await graphJson(`/me/drive/root:/${encodeURIComponent(WORKBOOK_FOLDER)}/${encodeURIComponent(WORKBOOK_FILE)}?$select=id,name,size`);
      ok(`Found ${item.name}`);
    } catch (err) { bad(`No file named exactly "${WORKBOOK_FILE}" in /${WORKBOOK_FOLDER}`, err); return; }
    const cached = await getMeta("workbookItemId", null);
    if (cached && cached !== item.id) {
      bad("The phone was pointing at an older copy of the workbook — fixed now");
      await setMeta("workbookItemId", item.id);
      await setMeta("wbSchemaReady", null);
    } else ok(cached ? "Phone is pointing at this file" : "Phone hadn't linked a file yet — linked now");
    if (!cached) await setMeta("workbookItemId", item.id);
    show();

    // 4. Open it with Excel Online
    let sheets;
    const t0 = Date.now();
    try {
      sheets = await loadSheetIds(item.id);
      ok(`Excel opened the workbook in ${((Date.now() - t0) / 1000).toFixed(1)} s. Tabs: ${sheets.map((x) => x.name).join(", ")}`);
    } catch (err) { bad("Excel Online couldn't open the workbook", err); return; }
    const missing = ALL_TABS.filter((t) => !sheets.some((x) => x.name === t));
    if (missing.length) bad(`Missing tabs (names must match exactly): ${missing.join(", ")}`);
    show();

    // 5. Read each data tab, by ID and (for comparison) by name
    for (const tab of DATA_TABS) {
      if (!SHEET_IDS.byName[tab]) continue;
      try {
        const b = await graphJson(`${wsPath(item.id, tab)}/usedRange(valuesOnly=true)?$select=address`);
        ok(`Read "${tab}": ${b.address}`);
      } catch (err) { bad(`Couldn't read "${tab}"`, err); }
      try {
        await graphJson(`${wsPathByName(item.id, tab)}/usedRange(valuesOnly=true)?$select=address`);
      } catch (err) { out.push(`  (by name it fails: ${String(err.message).slice(0, 200)})`); }
      show();
    }

    // 6. The whole read the app actually does
    try {
      const tabs = await readDataTabs(item.id);
      ok("Full read used by sync: " + DATA_TABS.map((t) => `${t} ${tabs[t].rows.length} rows`).join(", "));
    } catch (err) { bad("The sync's own read failed", err); }
  } catch (err) {
    bad("Unexpected error", err);
  } finally {
    state.diagBusy = false;
    show();
  }
}

// ---- Reconciliation -------------------------------------------------
// The phone is the source of truth for the cases it has. Compare every
// phone case with the workbook and list the ones with no matching row;
// nothing is added without the user's tap ("Add N to sheet").
async function reconcileWithSheet({ manual = false } = {}) {
  if (!(await ensureFreshToken())) {
    if (manual) toast("Connect Microsoft 365 first", true);
    render();
    return;
  }
  state.reconcileBusy = true;
  if (manual) render();
  try {
    const itemId = await ensureWorkbook();
    const tabs = await readDataTabs(itemId);
    const claimed = new Set();
    const missing = [];
    // Exact-CPT matches first so they claim their rows before looser ones.
    const ordered = [...CASES].sort((a, b) => {
      const fa = findCaseInTabs(a, tabs), fb = findCaseInTabs(b, tabs);
      const sa = fa && normCpt(tabs[fa.tab].rows[fa.row - 2][9]) === a.cptCodes.map((x) => x.code).join(";") ? 0 : 1;
      const sb = fb && normCpt(tabs[fb.tab].rows[fb.row - 2][9]) === b.cptCodes.map((x) => x.code).join(";") ? 0 : 1;
      return sa - sb;
    });
    for (const c of ordered) {
      const found = findCaseInTabs(c, tabs, claimed);
      if (found) {
        claimed.add(`${found.tab}|${found.row}`);
        await adoptSheetStatus(c, tabs[found.tab].rows[found.row - 2]);
        if (!c._sheetSync || c._sheetSync.tab !== found.tab || c._sheetSync.row !== found.row) {
          c._sheetSync = found;
          await saveCase(c);
        }
      } else if (!SYNC_QUEUE.includes(c.id)) {
        missing.push(c.id);
      }
    }
    missing.sort((a, b) => {
      const ca = CASES.find((x) => x.id === a), cb = CASES.find((x) => x.id === b);
      return String(cb.dos).localeCompare(String(ca.dos));
    });
    state.missingIds = missing;
    state.missingSelected = new Set(missing);
    state.reconcileCheckedAt = Date.now();
    if (manual) toast(missing.length ? `${missing.length} case(s) on this phone are not in the workbook` : "Everything on this phone is in the workbook ✓", !!missing.length);
  } catch (err) {
    console.error(err);
    await recordSyncError("Check for missing cases", err);
    if (manual) toast("Couldn't read the workbook — details in Settings → Diagnostics", true);
  } finally {
    state.reconcileBusy = false;
    render();
  }
}

async function addMissingToSheet() {
  const ids = [...(state.missingSelected || [])];
  if (!ids.length) return;
  state.reconcileBusy = true;
  render();
  for (const id of ids) {
    const c = CASES.find((x) => x.id === id);
    if (!c) continue;
    c._sheetSync = null;
    await saveCase(c);
    await syncCaseToDrive(c, { quiet: true });
  }
  state.reconcileBusy = false;
  await reconcileWithSheet({ manual: true });
}


// ---- Code library backup (v2.1) --------------------------------------
// The library lives on the phone (IndexedDB) AND on the "Code Library"
// tab of ChargeCap Log.xlsx. Phone edits are logged as operations
// (queueCodeOp) and replayed onto the tab's current contents, so codes
// typed into the tab in Excel are kept too. After each sync the phone's
// library is replaced by the tab, which is the master copy.
const TAB_CODES = "Code Library";
const CODES_HEADER = ["Type", "Code", "Description", "Category"];
const TYPE_LABEL = { cpt: "CPT", icd10: "ICD-10" };

function typeFromLabel(v) {
  const t = String(v || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (t === "CPT") return "cpt";
  if (t.startsWith("ICD")) return "icd10";
  return null;
}
function cellText(v) {
  if (v == null) return "";
  if (typeof v === "number") return String(Number.isInteger(v) ? v : v);
  return String(v).trim();
}
const codeKey = (code) => String(code || "").toUpperCase().replace(/\s+/g, "");
const sameEntry = (e, code, category) =>
  codeKey(e.code) === codeKey(code) && (category == null || String(e.category || "").toLowerCase() === String(category || "").toLowerCase());

function applyCodeOps(lists, ops) {
  for (const o of ops) {
    const list = lists[o.type];
    if (!list) continue;
    if (o.op === "add") {
      if (!list.some((e) => sameEntry(e, o.code, o.category))) list.unshift({ code: o.code, desc: o.desc, category: o.category });
    } else if (o.op === "update") {
      let i = list.findIndex((e) => sameEntry(e, o.from.code, o.from.category));
      if (i < 0) i = list.findIndex((e) => sameEntry(e, o.from.code, null));
      const next = { code: o.code, desc: o.desc, category: o.category };
      if (i >= 0) list[i] = next;
      else if (!list.some((e) => sameEntry(e, o.code, o.category))) list.unshift(next);
    } else if (o.op === "delete") {
      let i = list.findIndex((e) => sameEntry(e, o.code, o.category));
      if (i < 0) i = list.findIndex((e) => sameEntry(e, o.code, null) && e.desc === o.desc);
      if (i >= 0) list.splice(i, 1);
    }
  }
  return lists;
}

let codeSyncTimer = null;
function scheduleCodeLibrarySync() {
  refreshCodeLibState();
  clearTimeout(codeSyncTimer);
  codeSyncTimer = setTimeout(() => { if (state._m365Account) syncCodeLibrary({ manual: false }); }, 2000);
}
async function refreshCodeLibState() {
  state.codeLibPending = ((await getMeta("codeOps", [])) || []).length;
  state.codeLibSyncedAt = await getMeta("codeLibSyncedAt", null);
}

let codeSyncing = false;
async function syncCodeLibrary({ manual = false } = {}) {
  if (codeSyncing) return;
  codeSyncing = true;
  state.codeLibBusy = true;
  if (manual) render();
  try {
    if (!(await ensureFreshToken())) { if (manual) toast("Connect Microsoft 365 first", true); return; }
    const itemId = await ensureWorkbook();
    const ops = (await getMeta("codeOps", [])) || [];
    const opCount = ops.length;
    const sheets = await graphJson(`${wbPath(itemId)}/worksheets?$select=name`);
    const hasTab = (sheets.value || []).some((s) => s.name === TAB_CODES);

    let remote = null; // { cpt: [...], icd10: [...] }
    let oldLastRow = 1;
    if (hasTab) {
      const body = await graphJson(`${wsPath(itemId, TAB_CODES)}/usedRange(valuesOnly=true)?$select=address,values`);
      const start = rowFromAddress(body.address);
      const values = body.values || [];
      oldLastRow = Math.max(1, start + values.length - 1);
      remote = { cpt: [], icd10: [] };
      values.forEach((r, j) => {
        if (start + j < 2) return; // header
        const type = typeFromLabel(r[0]);
        const code = cellText(r[1]);
        if (!type || !code) return;
        remote[type].push({ code, desc: cellText(r[2]) || code, category: cellText(r[3]) || "Custom" });
      });
      if (!remote.cpt.length && !remote.icd10.length) remote = null;
    }

    let result;
    let mustWrite = opCount > 0;
    if (!remote) {
      // No tab (or an empty one): this phone's library becomes the master.
      result = { cpt: CPT_LIB.map(({ code, desc, category }) => ({ code, desc, category })), icd10: ICD10_LIB.map(({ code, desc, category }) => ({ code, desc, category })) };
      mustWrite = true;
    } else {
      result = applyCodeOps(remote, ops);
      // First sync from this phone: keep any codes that exist only here
      // (added before v2.1), instead of letting the tab overwrite them.
      if (!(await getMeta("codeLibLinked", false))) {
        for (const type of ["cpt", "icd10"]) {
          for (const e of codeLib(type)) {
            if (!result[type].some((r) => sameEntry(r, e.code, null))) {
              result[type].unshift({ code: e.code, desc: e.desc, category: e.category });
              mustWrite = true;
            }
          }
        }
      }
    }

    if (mustWrite) {
      const rows = [
        ...result.cpt.map((e) => ["CPT", e.code, e.desc, e.category]),
        ...result.icd10.map((e) => ["ICD-10", e.code, e.desc, e.category]),
      ];
      const reqs = [];
      if (!hasTab) reqs.push({ method: "POST", url: `${wbPath(itemId)}/worksheets/add`, body: { name: TAB_CODES } });
      const lastNeeded = Math.max(oldLastRow, rows.length + 1);
      if (hasTab && oldLastRow >= 2) reqs.push({ method: "POST", url: `${rangePath(itemId, TAB_CODES, `A2:D${lastNeeded}`)}/clear`, body: { applyTo: "Contents" } });
      reqs.push({ method: "PATCH", url: rangePath(itemId, TAB_CODES, "A1:D1"), body: { values: [CODES_HEADER] } });
      for (let i = 0; i < rows.length; i += 200) {
        const chunk = rows.slice(i, i + 200);
        const r0 = i + 2;
        reqs.push({
          method: "PATCH",
          url: rangePath(itemId, TAB_CODES, `A${r0}:D${r0 + chunk.length - 1}`),
          body: { numberFormat: chunk.map(() => ["@", "@", "@", "@"]), values: chunk },
        });
      }
      for (let i = 0; i < reqs.length; i += 20) await graphBatch(reqs.slice(i, i + 20));
    }

    // The tab is now the master: replace the phone's copy, keeping ids
    // for entries that didn't change so an open edit form isn't broken.
    for (const type of ["cpt", "icd10"]) {
      const old = codeLib(type);
      const next = result[type].map((e) => {
        const prev = old.find((o) => sameEntry(o, e.code, e.category) && o.desc === e.desc);
        return { ...e, id: prev ? prev.id : uuid() };
      });
      if (type === "cpt") CPT_LIB = next; else ICD10_LIB = next;
      await persistCodeLib(type);
    }
    // Drop only the operations this sync replayed; edits made while it
    // was running stay queued for the next one.
    const now = (await getMeta("codeOps", [])) || [];
    await setMeta("codeOps", now.slice(opCount));
    await setMeta("codeLibLinked", true);
    await setMeta("codeLibSyncedAt", Date.now());
    await refreshCodeLibState();
    if (manual) toast("Code library synced");
    if (state.codeLibPending) scheduleCodeLibrarySync();
  } catch (err) {
    console.error("Code library sync failed", err);
    await recordSyncError("Code library sync", err);
    if (manual) toast("Code library sync failed — your changes are saved on this phone and will retry", true);
  } finally {
    codeSyncing = false;
    state.codeLibBusy = false;
    render();
  }
}

let flushing = false;
async function flushSyncQueue() {
  if (flushing || !SYNC_QUEUE.length) return;
  flushing = true;
  try {
    if (!(await ensureFreshToken())) { render(); return; }
    const queue = [...SYNC_QUEUE];
    let ok = 0;
    for (const id of queue) {
      const c = CASES.find((x) => x.id === id);
      if (!c) { await dequeueSync(id); continue; }
      await syncCaseToDrive(c, { quiet: true });
      if (!SYNC_QUEUE.includes(id)) ok++;
    }
    if (ok) toast(`Saved ${ok} queued case(s) to Excel`);
    if (SYNC_QUEUE.length) toast(`${SYNC_QUEUE.length} case(s) still waiting to sync`, true);
  } finally {
    flushing = false;
    render();
  }
}

window.addEventListener("online", () => { flushSyncQueue(); if (state.codeLibPending) syncCodeLibrary(); });
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") {
    flushSyncQueue();
    if (state.codeLibPending) syncCodeLibrary();
  }
});

// ---------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------

// One-time switch from the Google Sheet (v1.x) to the OneDrive workbook:
// old row pointers refer to the Google sheet, so drop them (the first
// reconcile re-links every case to its row in the workbook) and forget
// the Google token/sheet id. Cases themselves are untouched.
async function migrateToM365() {
  if ((await getMeta("syncBackend", null)) === "m365") return;
  for (const c of CASES) {
    if (c._sheetSync) {
      delete c._sheetSync;
      await saveCase(c);
    }
  }
  for (const k of ["driveToken", "driveTokenExpiry", "sheetId", "tabsReady", "googleClientId"]) await setMeta(k, null);
  await setMeta("syncBackend", "m365");
}

async function boot() {
  await loadCases();
  await loadCodes();
  state.lastFacility = await getMeta("lastFacility", null);
  SYNC_QUEUE = (await getMeta("syncQueue", [])) || [];
  state.lastSyncError = await getMeta("lastSyncError", null);
  await migrateToM365();
  render();

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch((err) => console.error("SW registration failed", err));
  }

  // Processes the sign-in redirect if we just came back from Microsoft.
  await initMsal();
  await refreshCodeLibState();
  if (state._m365Account && (await ensureFreshToken())) {
    await flushSyncQueue();
    await syncCodeLibrary({ manual: false });
    reconcileWithSheet({ manual: false });
  } else {
    render();
  }
}

boot();
