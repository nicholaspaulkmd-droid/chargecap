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
  liveScanOpen: false, // true while the live camera screen is open (see "Live camera scan" below)
  editingCaseId: null,
  ocrBusy: false,
  showRawOcr: false,
  // Last facility picked on the Capture screen (persisted in IndexedDB
  // meta "lastFacility") — every new case defaults to it, since a whole
  // operating day is usually at one location.
  lastFacility: null,
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
    const worker = await getOcrWorker();
    const { data } = await worker.recognize(await prepareForOcr(file));
    const parsed = parseOcrText(data.text, data.words || []);
    parsed.rawText = data.text;
    return parsed;
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
const SCREEN_NON_NAME_WORDS = /\b(Male|Female|Bed|Room|Unit|Floor|Code|Status|Full|Allergies|Isolation|Precautions|Admitted|Admission|Location|Inpatient|Outpatient|Observation|Emergency|Surgery|Pre|Post|Op|Medical|Center|Hospital|Health|Clinic|Intermountain|Chart|Summary|Orders|Results|Notes|Review)\b/;

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
      for (const lf of line.matchAll(/([A-Z][A-Za-z'\-]+),\s*([A-Z][A-Za-z'\-]+)/g)) {
        if (CREDENTIAL_WORD.test(lf[2])) continue;
        nameGuess = lf[0];
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
    // EHR screen banner style (added 2026-09-23): the name sits alone on
    // its own line in normal "First M. Last" capitalization with no label
    // and no comma, e.g. "Lesley M. Webster" above a "Female, 77 y.o.,
    // 7/23/1949" line. Take the first such line and flip it to the
    // "Last, First M." format the rest of the app uses.
    const titleLine = lines.find(
      (l) =>
        /^[A-Z][a-z'\-]+(?:\s+(?:[A-Z]\.?|[A-Z][A-Za-z'\-]*[a-z][A-Za-z'\-]*)){1,3}$/.test(l) &&
        l.length <= 35 &&
        !NAME_EXCLUDE_WORDS.test(l) &&
        !NAME_DISQUALIFY_CONTEXT.test(l) &&
        !SCREEN_NON_NAME_WORDS.test(l)
    );
    if (titleLine) {
      const parts = titleLine.split(/\s+/);
      const last = parts[parts.length - 1];
      nameGuess = `${last}, ${parts.slice(0, -1).join(" ")}`;
    }
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
  const needsReconnect = !state._driveToken && state._clientId && SYNC_QUEUE.length && state.view !== "settings";
  const syncBanner = needsReconnect
    ? `<div class="sync-banner" id="syncBanner">⚠️ ${SYNC_QUEUE.length} case(s) waiting to sync — tap to reconnect Google Drive</div>`
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
  const driveConnected = !!state._driveToken;
  return `
    <header class="topbar"><h1>Settings</h1></header>
    <div class="content">
      <section class="card">
        <h2>Google Drive sync</h2>
        <p class="muted">Every saved case backs up automatically to a spreadsheet called <strong>${SHEET_NAME}</strong> in your Google Drive — primary/co-surgeon cases on one tab, assistant cases on another, plus a Monthly Tally tab that auto-counts Bariatric/EGD/Back/General Surgery/Tummy Tuck cases per month, with Non-Op Consults (consult codes only, no procedure) counted separately. Data goes only to your own Google account — no other server is involved.</p>
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
  if (syncBannerEl) syncBannerEl.addEventListener("click", connectGoogleDrive);

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

// Access tokens from Google's implicit/token-client flow expire (~1hr)
// and there's no refresh token in this browser-only flow — that's
// inherent to the security model, not something we can change. What WAS
// a bug: state._driveToken only ever lived in memory, never written to
// IndexedDB, so it was lost on every app restart (and on iOS a PWA's JS
// gets evicted from memory constantly just from backgrounding), forcing
// a manual "Connect" in Settings far more often than the token's actual
// 1hr lifetime should require. Fixed by persisting the token alongside
// its expiry and restoring it on boot (see restoreDriveToken()), plus
// attempting a silent (no-UI) reauth via prompt:"" before falling back
// to asking the user to tap reconnect.
let pendingTokenResolvers = [];
let silentAttemptInFlight = false;

function initGoogleAuth() {
  if (!state._clientId || !window.google) return;
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: state._clientId,
    scope: SHEETS_SCOPE,
    callback: async (resp) => {
      const resolvers = pendingTokenResolvers.splice(0);
      const wasSilent = silentAttemptInFlight;
      silentAttemptInFlight = false;
      if (resp.error) {
        // A silent (background) attempt failing is expected whenever
        // there's no live Google session to reuse — that's not an error
        // worth interrupting the user for. Only a user-initiated
        // "Connect"/"reconnect" tap surfaces a toast on failure.
        if (!wasSilent) toast("Google sign-in failed", true);
        resolvers.forEach((r) => r(false));
        return;
      }
      state._driveToken = resp.access_token;
      const expiresAt = Date.now() + (resp.expires_in || 3600) * 1000;
      await setMeta("driveToken", resp.access_token);
      await setMeta("driveTokenExpiry", expiresAt);
      if (!wasSilent) toast("Google Drive connected");
      render();
      resolvers.forEach((r) => r(true));
      flushSyncQueue();
    },
  });
}

// Restores a still-valid token from IndexedDB on boot so the app doesn't
// need to re-authenticate every time it's relaunched — only once the
// token has actually expired.
async function restoreDriveToken() {
  const token = await getMeta("driveToken", null);
  const expiresAt = await getMeta("driveTokenExpiry", 0);
  if (token && Date.now() < expiresAt - 60 * 1000) {
    state._driveToken = token;
  }
}

// Wraps tokenClient.requestAccessToken() in a Promise. silent:true uses
// prompt:"" (no popup/consent UI — succeeds only if Google can reissue a
// token without asking the user anything; fails quietly otherwise).
function requestToken({ silent = false } = {}) {
  return new Promise((resolve) => {
    if (!tokenClient) { resolve(false); return; }
    pendingTokenResolvers.push(resolve);
    if (silent) silentAttemptInFlight = true;
    // Silent: prompt:"" — no UI, succeeds only if Google can reissue
    // without asking anything. Manual: no override, same as before —
    // let Google show the minimum it needs (often just an instant
    // account-picker tap for a previously-granted user, not a full
    // consent screen every time).
    tokenClient.requestAccessToken(silent ? { prompt: "" } : {});
  });
}

// Google's GSI script (accounts.google.com/gsi/client) loads async over
// the network — on a fresh app launch it's often not ready yet even
// though app.js already is. Anything that needs tokenClient waits for
// it here first instead of silently no-op'ing.
function waitForGoogleIdentity(timeoutMs = 15000, intervalMs = 250) {
  return new Promise((resolve) => {
    if (window.google && window.google.accounts && window.google.accounts.oauth2) { resolve(true); return; }
    const start = Date.now();
    const id = setInterval(() => {
      if (window.google && window.google.accounts && window.google.accounts.oauth2) {
        clearInterval(id);
        resolve(true);
      } else if (Date.now() - start > timeoutMs) {
        clearInterval(id);
        resolve(false);
      }
    }, intervalMs);
  });
}

// Explicit, user-initiated connect/reconnect (Settings button, sync
// banner tap). Always shows the Google UI so it works even the very
// first time, before any session/consent exists to reuse silently.
async function connectGoogleDrive() {
  if (!state._clientId) { toast("Add a Client ID first", true); return; }
  if (!tokenClient) {
    if (!(await waitForGoogleIdentity(5000))) {
      toast("Google sign-in isn't ready yet — try again in a moment", true);
      return;
    }
    initGoogleAuth();
  }
  if (!tokenClient) { toast("Add a Client ID first", true); return; }
  requestToken({ silent: false });
}

// Called before anything that needs Drive access. Returns true if
// state._driveToken is set and good to use — reusing the persisted
// token when it's still valid, otherwise trying a silent reauth first
// so most syncs never need a tap at all. Only returns false (caller
// should queue + show the reconnect banner) when even that fails,
// which mainly happens after ~1hr idle with no live Google session to
// reuse, or before the very first connect.
async function ensureFreshToken() {
  const expiresAt = await getMeta("driveTokenExpiry", 0);
  if (state._driveToken && Date.now() < expiresAt - 60 * 1000) return true;
  if (!state._clientId) return false;
  if (!tokenClient) initGoogleAuth();
  if (!tokenClient) return false;
  return requestToken({ silent: true });
}

async function driveFetch(url, opts = {}) {
  if (!state._driveToken) throw new Error("not-connected");
  return fetch(url, {
    ...opts,
    headers: { ...(opts.headers || {}), Authorization: `Bearer ${state._driveToken}` },
  });
}

// driveFetch() only rejects on a network-level failure — a 400/403/etc.
// HTTP response still resolves normally, so callers that don't check
// res.ok can silently treat a failed write as a success (this is how
// the "All Cases"/"Monthly Tally" formulas ended up never actually
// written: the setup call failed server-side but ensureTabs() marked
// itself done anyway and never retried). Use this wrapper for any
// Sheets/Drive write whose success matters.
async function driveFetchOk(url, opts = {}, _retried = false) {
  const res = await driveFetch(url, opts);
  // A stored token can go stale mid-session (revoked, or our expiry
  // estimate was slightly optimistic) — a 401 here means Google itself
  // rejected it, not just our local clock. Try one silent reauth + retry
  // before giving up, so an in-progress sync recovers on its own instead
  // of surfacing an avoidable failure.
  if (res.status === 401 && !_retried) {
    state._driveToken = null;
    const refreshed = await requestToken({ silent: true });
    if (refreshed) return driveFetchOk(url, opts, true);
  }
  if (!res.ok) {
    let detail = "";
    try { detail = JSON.stringify((await res.json()).error || {}); } catch { /* ignore */ }
    throw new Error(`Sheets API ${res.status} on ${url.split("?")[0]}: ${detail}`);
  }
  return res;
}

async function findOrCreateSheet() {
  const cachedId = await getMeta("sheetId", null);
  if (cachedId) return cachedId;

  const q = encodeURIComponent(`name='${SHEET_NAME}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`);
  const listRes = await driveFetchOk(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`);
  const listJson = await listRes.json();
  if (listJson.files && listJson.files.length) {
    await setMeta("sheetId", listJson.files[0].id);
    return listJson.files[0].id;
  }

  const createRes = await driveFetchOk("https://sheets.googleapis.com/v4/spreadsheets", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ properties: { title: SHEET_NAME } }),
  });
  const createJson = await createRes.json();
  const sheetId = createJson.spreadsheetId;
  await setMeta("sheetId", sheetId);
  return sheetId;
}

// Bump whenever the tab/formula structure changes (e.g. adding a new
// tally category) so an already-connected user's spreadsheet picks up
// the new columns/formulas automatically on their next sync, instead of
// staying stuck with whatever schema existed when they first connected.
// Mirrors DB_VERSION's self-repair pattern above. "tabsReady" used to
// store a plain `true`; comparing with >= still does the right thing
// for a legacy `true` value (coerces to 1, which is < any bumped
// version here) so old installs re-run this once automatically.
//
// v3: added the "Tummy Tuck" tally category (CPT 15830/15847) and,
// separately, fixed a real bug from v2 — the setup calls below weren't
// checked for failure, so a rejected write (e.g. a malformed request)
// could still mark tabsReady done and permanently skip the "All Cases"/
// "Monthly Tally" formulas. Bumping the version forces everyone to
// redo this once with the fix in place.
// v4: removed the "Surgeon" column (redundant with Role) — sheet is now
// 15 columns (A-O) instead of 16 (A-P). Rewriting headers/formulas here
// only fixes the TOP of the sheet (header row + tally); it does NOT
// re-shift any data rows already written under the old 16-column
// layout. Existing rows must be migrated by hand — see README.
// v5 (2026-09-16): Monthly Tally was counting every case in "All Cases"
// — which deliberately includes Assistant-role cases too, for the
// combined log — so assist cases were inflating the user's own tally.
// Monthly Tally's A2/B2:F2 formulas now read from "Primary & Co-Surgeon"
// only. Bumping the version so an already-connected install picks up
// the corrected formulas on next sync instead of keeping the old ones
// (the live spreadsheet was also hand-fixed directly on 2026-09-16 —
// this bump just makes sure a fresh ensureTabs() run agrees with it).
// v6 (2026-09-23): added the "Non-Op Consults" tally column (H) — cases
// whose only CPT codes are consult/E&M codes (see CONSULT_CPT). Placed
// AFTER Total on purpose: Total (G) stays the operative-case count and
// does NOT include consults. Also re-tags any already-synced consult-only
// rows' Category cell (column O) from "General Surgery" to
// "Non-Op Consult" — see retagNonOpRows().
const TABS_SCHEMA_VERSION = 6;

// Makes sure the four tabs (Primary & Co-Surgeon, Assistant, All Cases,
// Monthly Tally) exist with headers + formulas in place, and that the
// Monthly Tally formulas match TABS_SCHEMA_VERSION. Runs once per schema
// version (tracked via the "tabsReady" meta flag) and is safe to re-run
// — it only adds what's missing, and every Monthly Tally cell is always
// a live formula anyway, so overwriting them in place is harmless (see
// TABS_SCHEMA_VERSION above).
async function ensureTabs(sheetId) {
  const readyVersion = await getMeta("tabsReady", 0);
  if (readyVersion >= TABS_SCHEMA_VERSION) return;

  const metaRes = await driveFetchOk(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties`);
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
    await driveFetchOk(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}:batchUpdate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ requests }),
    });
  }

  // Header rows + the formulas that drive "All Cases" and "Monthly Tally".
  // USER_ENTERED (not RAW) so date strings like "8/27/2026" get parsed
  // into real Sheets dates — the tally's date-range math depends on that.
  //
  // Monthly Tally counts ONLY Primary & Co-Surgeon cases — it must not
  // read from "All Cases" (which deliberately also includes Assistant-role
  // cases, for the full combined log view). tallyDateRange therefore
  // points at TAB_PRIMARY, not TAB_ALL, even though TAB_ASSISTANT-role
  // cases never appear in TAB_PRIMARY in the first place (see
  // tabForRole()), so no separate role filter is needed on top of this.
  const allDateRange = `'${TAB_ALL}'!A2:A5000`;
  const tallyDateRange = `'${TAB_PRIMARY}'!A2:A5000`;
  await driveFetchOk(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchUpdate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      valueInputOption: "USER_ENTERED",
      data: [
        { range: `'${TAB_PRIMARY}'!A1:O1`, values: [SHEET_HEADER] },
        { range: `'${TAB_ASSISTANT}'!A1:O1`, values: [SHEET_HEADER] },
        { range: `'${TAB_ALL}'!A1:O1`, values: [SHEET_HEADER] },
        {
          // Stacks both role tabs into one sorted-by-date table so
          // there's a single combined log to look at — this tab is NOT
          // what the tally counts from (see tallyDateRange above).
          range: `'${TAB_ALL}'!A2`,
          values: [[`=SORT({'${TAB_PRIMARY}'!A2:O5000;'${TAB_ASSISTANT}'!A2:O5000},1,TRUE)`]],
        },
        {
          range: `'${TAB_TALLY}'!A1:H1`,
          values: [["Month", "Bariatric", "EGD", "Back", "General Surgery", "Tummy Tuck", "Total", "Non-Op Consults"]],
        },
        {
          // One first-of-month row per month that actually has a
          // Primary/Co-Surgeon case, newest formulas spill down
          // automatically as rows are added.
          range: `'${TAB_TALLY}'!A2`,
          values: [[`=SORT(UNIQUE(FILTER(EOMONTH(${tallyDateRange},-1)+1,${tallyDateRange}<>"")))`]],
        },
        {
          range: `'${TAB_TALLY}'!B2`,
          values: [[`=ARRAYFORMULA(IF(A2:A="","",COUNTIFS('${TAB_PRIMARY}'!$A$2:$A$5000,">="&A2:A,'${TAB_PRIMARY}'!$A$2:$A$5000,"<"&EDATE(A2:A,1),'${TAB_PRIMARY}'!$O$2:$O$5000,"Bariatric")))`]],
        },
        {
          range: `'${TAB_TALLY}'!C2`,
          values: [[`=ARRAYFORMULA(IF(A2:A="","",COUNTIFS('${TAB_PRIMARY}'!$A$2:$A$5000,">="&A2:A,'${TAB_PRIMARY}'!$A$2:$A$5000,"<"&EDATE(A2:A,1),'${TAB_PRIMARY}'!$O$2:$O$5000,"EGD")))`]],
        },
        {
          range: `'${TAB_TALLY}'!D2`,
          values: [[`=ARRAYFORMULA(IF(A2:A="","",COUNTIFS('${TAB_PRIMARY}'!$A$2:$A$5000,">="&A2:A,'${TAB_PRIMARY}'!$A$2:$A$5000,"<"&EDATE(A2:A,1),'${TAB_PRIMARY}'!$O$2:$O$5000,"Back")))`]],
        },
        {
          range: `'${TAB_TALLY}'!E2`,
          values: [[`=ARRAYFORMULA(IF(A2:A="","",COUNTIFS('${TAB_PRIMARY}'!$A$2:$A$5000,">="&A2:A,'${TAB_PRIMARY}'!$A$2:$A$5000,"<"&EDATE(A2:A,1),'${TAB_PRIMARY}'!$O$2:$O$5000,"General Surgery")))`]],
        },
        {
          range: `'${TAB_TALLY}'!F2`,
          values: [[`=ARRAYFORMULA(IF(A2:A="","",COUNTIFS('${TAB_PRIMARY}'!$A$2:$A$5000,">="&A2:A,'${TAB_PRIMARY}'!$A$2:$A$5000,"<"&EDATE(A2:A,1),'${TAB_PRIMARY}'!$O$2:$O$5000,"Tummy Tuck")))`]],
        },
        { range: `'${TAB_TALLY}'!G2`, values: [[`=ARRAYFORMULA(IF(A2:A="","",B2:B+C2:C+D2:D+E2:E+F2:F))`]] },
        {
          range: `'${TAB_TALLY}'!H2`,
          values: [[`=ARRAYFORMULA(IF(A2:A="","",COUNTIFS('${TAB_PRIMARY}'!$A$2:$A$5000,">="&A2:A,'${TAB_PRIMARY}'!$A$2:$A$5000,"<"&EDATE(A2:A,1),'${TAB_PRIMARY}'!$O$2:$O$5000,"${NON_OP_CATEGORY}")))`]],
        },
      ],
    }),
  });

  await retagNonOpRows(sheetId);
  await setMeta("tabsReady", TABS_SCHEMA_VERSION);
}

// One-time fix-up for the v6 "Non-Op Consult" category: cases already
// synced before v6 had their Category cell (column O) written as
// "General Surgery". For each such case we know the row of (via
// _sheetSync), re-read that row first and only rewrite column O if the
// row's MRN + patient name still match the case — so a row that's been
// moved/sorted by hand is left alone rather than overwritten blindly.
async function retagNonOpRows(sheetId) {
  const targets = CASES.filter(
    (c) => c._sheetSync && c._sheetSync.tab && c._sheetSync.row && caseCategory(c) === NON_OP_CATEGORY
  );
  if (!targets.length) return;
  const ranges = targets.map((c) => `'${c._sheetSync.tab}'!A${c._sheetSync.row}:O${c._sheetSync.row}`);
  const qs = ranges.map((r) => "ranges=" + encodeURIComponent(r)).join("&");
  const res = await driveFetchOk(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchGet?${qs}`);
  const json = await res.json();
  const data = [];
  (json.valueRanges || []).forEach((vr, i) => {
    const c = targets[i];
    const row = (vr.values && vr.values[0]) || [];
    if (String(row[1] || "") !== String(c.patientName || "")) return;
    if (String(row[2] || "") !== String(c.mrn || "")) return;
    if (row[14] === NON_OP_CATEGORY) return;
    data.push({ range: `'${c._sheetSync.tab}'!O${c._sheetSync.row}`, values: [[NON_OP_CATEGORY]] });
  });
  if (!data.length) return;
  await driveFetchOk(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchUpdate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ valueInputOption: "USER_ENTERED", data }),
  });
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
  // Tries a silent reauth first (see ensureFreshToken) so a case synced
  // shortly after the token expired still goes straight through instead
  // of landing in the queue and waiting on a manual reconnect.
  const ready = await ensureFreshToken();
  if (!ready) {
    if (!SYNC_QUEUE.includes(c.id)) SYNC_QUEUE.push(c.id);
    await setMeta("syncQueue", SYNC_QUEUE);
    render(); // surface the reconnect banner right away, not just on next nav
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
      await driveFetchOk(
        `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/'${tab}'!A${c._sheetSync.row}:O${c._sheetSync.row}?valueInputOption=USER_ENTERED`,
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
        await driveFetchOk(
          `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/'${c._sheetSync.tab}'!A${c._sheetSync.row}:O${c._sheetSync.row}:clear`,
          { method: "POST" }
        );
      }
      const appendRes = await driveFetchOk(
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
  if (!SYNC_QUEUE.length) return;
  if (!(await ensureFreshToken())) return;
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
  state.lastFacility = await getMeta("lastFacility", null);
  state._clientId = (await getMeta("googleClientId", DEFAULT_GOOGLE_CLIENT_ID)) || DEFAULT_GOOGLE_CLIENT_ID;
  SYNC_QUEUE = (await getMeta("syncQueue", [])) || [];
  await restoreDriveToken(); // reuse a still-valid token instead of forcing reconnect on every launch
  render();

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch((err) => console.error("SW registration failed", err));
  }

  if (state._clientId) {
    // Google's GSI script loads async over the network and usually
    // isn't ready this early — wait for it before touching tokenClient,
    // otherwise this whole attempt just silently no-ops (see
    // waitForGoogleIdentity) and it's back to needing a manual tap for
    // no real reason.
    if (await waitForGoogleIdentity()) {
      initGoogleAuth();
      // If the restored token was missing/expired, try a silent (no-UI)
      // reauth — on a device that's stayed signed into Google this
      // often succeeds with no tap at all. flushSyncQueue() covers the
      // case where a valid token was already restored; if the silent
      // attempt fails too, the reconnect banner (see render()) is the
      // fallback.
      if (state._driveToken) {
        flushSyncQueue();
      } else {
        const ok = await ensureFreshToken();
        if (ok) render();
      }
    }
  }
}

boot();
