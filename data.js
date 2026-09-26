// data.js
// CPT and ICD-10 favorite code libraries for ChargeCap.
// Rebuilt 2026-09-23 from Dr. Paulk's updated "Billing Sheet - OR.xlsx"
// (the sheet is the source of truth — codes match it as written;
// descriptions lightly cleaned up for spelling).
//
// IMPORTANT: Verify against current CPT/ICD-10-CM documentation before
// relying on these for claims submission. Codes can be edited any time in
// the app under Settings > Manage Code Library (stored in IndexedDB).
//
// FAVORITES_VERSION: bump whenever these lists change. On next launch the
// app replaces its saved copy of the favorites with these lists, keeping
// any codes the user added themselves in the app (see loadCodes()).
const FAVORITES_VERSION = 2;

const CPT_FAVORITES = [
  // --- Bariatric (primary surgeries) ---
  { code: "43644", desc: "Gastric bypass, Roux-en-Y — LAP", category: "Bariatric" },
  { code: "43846", desc: "Gastric bypass, Roux-en-Y — OPEN", category: "Bariatric" },
  { code: "43775", desc: "Gastric sleeve — LAP", category: "Bariatric" },
  { code: "43843", desc: "Gastric sleeve — OPEN", category: "Bariatric" },
  { code: "43659", desc: "Duodenal switch (unlisted) — LAP", category: "Bariatric" },
  { code: "43845", desc: "Duodenal switch — OPEN", category: "Bariatric" },

  // --- Revision ---
  { code: "43774", desc: "Lap band removal, including components", category: "Revision" },
  { code: "43860", desc: "Revision of GJ anastomosis — OPEN", category: "Revision" },
  { code: "44202 + 44203", desc: "Revision JJ", category: "Revision" },
  { code: "43848", desc: "Revision VBG", category: "Revision" },
  { code: "43631 + 43659 + 44202", desc: "Bypass to DS — subtotal gastrectomy, DS, small bowel resection", category: "Revision" },
  { code: "43659", desc: "Sleeve to DS", category: "Revision" },

  // --- Endoscopy ---
  { code: "43266", desc: "EGD with stent", category: "Endoscopy" },
  { code: "43235", desc: "EGD — no biopsy or intra-op", category: "Endoscopy" },
  { code: "43239", desc: "EGD with biopsy", category: "Endoscopy" },
  { code: "43245", desc: "EGD with dilation", category: "Endoscopy" },
  { code: "43247", desc: "EGD with foreign body removal", category: "Endoscopy" },

  // --- Upper GI ---
  { code: "43280", desc: "Fundoplication only — LAP", category: "Upper GI" },
  { code: "43633", desc: "Subtotal gastrectomy, Roux-en-Y reconstruction", category: "Upper GI" },
  { code: "43653", desc: "Gastrostomy — LAP", category: "Upper GI" },
  { code: "43832", desc: "Gastrostomy tube (G tube)", category: "Upper GI" },
  { code: "43840", desc: "Repair gastric / duodenal ulcer", category: "Upper GI" },

  // --- Bowel ---
  { code: "44050", desc: "Internal hernia / intussusception", category: "Bowel" },
  { code: "44120", desc: "Enterectomy (single)", category: "Bowel" },
  { code: "44121", desc: "Enterectomy — each additional segment", category: "Bowel" },
  { code: "44130", desc: "Enteroenterostomy, extra anastomosis", category: "Bowel" },
  { code: "44180", desc: "Enterolysis — LAP", category: "Bowel" },
  { code: "44005", desc: "Enterolysis — OPEN", category: "Bowel" },
  { code: "44186", desc: "Jejunostomy tube", category: "Bowel" },
  { code: "44202", desc: "Small bowel resection — LAP", category: "Bowel" },
  { code: "44203", desc: "Small bowel resection, each additional — LAP", category: "Bowel" },
  { code: "49320", desc: "Exploratory laparoscopy", category: "Bowel" },
  { code: "44602", desc: "Enterorrhaphy — wound / injury / ulcer", category: "Bowel" },
  { code: "49326", desc: "Omentopexy (Graham patch) — LAP", category: "Bowel" },

  // --- Gallbladder ---
  { code: "47562", desc: "Cholecystectomy — LAP", category: "Gallbladder" },
  { code: "47563", desc: "Cholecystectomy with IOC — LAP", category: "Gallbladder" },
  { code: "47564", desc: "Cholecystectomy with common duct exploration — LAP", category: "Gallbladder" },

  // --- Appendix ---
  { code: "44970", desc: "Appendectomy — LAP", category: "Appendix" },
  { code: "44960", desc: "Appendectomy, ruptured", category: "Appendix" },
  { code: "44950", desc: "Appendectomy — OPEN", category: "Appendix" },

  // --- Liver / Spleen / Pancreas ---
  { code: "47100", desc: "Wedge liver biopsy", category: "Liver / Spleen / Pancreas" },
  { code: "47001", desc: "Liver biopsy, needle", category: "Liver / Spleen / Pancreas" },
  { code: "38120", desc: "Splenectomy — LAP", category: "Liver / Spleen / Pancreas" },
  { code: "38100", desc: "Splenectomy — OPEN", category: "Liver / Spleen / Pancreas" },
  { code: "48140", desc: "Distal pancreatectomy ± spleen", category: "Liver / Spleen / Pancreas" },

  // --- Hernia: hiatal ---
  { code: "43281", desc: "Paraesophageal / hiatal hernia — LAP", category: "Hernia" },
  { code: "43282", desc: "Paraesophageal / hiatal hernia with mesh — LAP", category: "Hernia" },
  // --- Hernia: ventral / incisional (2023+ codes, mesh included) ---
  { code: "49591", desc: "Ventral, initial, reducible, <3 cm", category: "Hernia" },
  { code: "49592", desc: "Ventral, initial, incarcerated, <3 cm", category: "Hernia" },
  { code: "49593", desc: "Ventral, initial, reducible, 3–10 cm", category: "Hernia" },
  { code: "49594", desc: "Ventral, initial, incarcerated, 3–10 cm", category: "Hernia" },
  { code: "49595", desc: "Ventral, initial, reducible, >10 cm", category: "Hernia" },
  { code: "49596", desc: "Ventral, initial, incarcerated, >10 cm", category: "Hernia" },
  { code: "49613", desc: "Ventral, recurrent, reducible, <3 cm", category: "Hernia" },
  { code: "49614", desc: "Ventral, recurrent, incarcerated, <3 cm", category: "Hernia" },
  { code: "49615", desc: "Ventral, recurrent, reducible, 3–10 cm", category: "Hernia" },
  { code: "49616", desc: "Ventral, recurrent, incarcerated, 3–10 cm", category: "Hernia" },
  { code: "49617", desc: "Ventral, recurrent, reducible, >10 cm", category: "Hernia" },
  { code: "49618", desc: "Ventral, recurrent, incarcerated, >10 cm", category: "Hernia" },
  { code: "49621", desc: "Parastomal, reducible", category: "Hernia" },
  { code: "49622", desc: "Parastomal, incarcerated", category: "Hernia" },
  { code: "49623", desc: "Mesh removal (add-on)", category: "Hernia" },
  { code: "15778", desc: "Abdominal wall (non-hernia) absorbable mesh", category: "Hernia" },
  // --- Hernia: inguinal ---
  { code: "49505", desc: "Inguinal, open, reducible, initial", category: "Hernia" },
  { code: "49507", desc: "Inguinal, open, incarcerated, initial", category: "Hernia" },
  { code: "49520", desc: "Inguinal, open, recurrent, reducible", category: "Hernia" },
  { code: "49521", desc: "Inguinal, open, recurrent, incarcerated", category: "Hernia" },
  { code: "49525", desc: "Inguinal, open, sliding", category: "Hernia" },
  { code: "49650", desc: "Inguinal, initial — LAP / robotic", category: "Hernia" },
  { code: "49651", desc: "Inguinal, recurrent — LAP / robotic", category: "Hernia" },

  // --- Component separation ---
  { code: "15273", desc: "Biologic implant >100 cm²", category: "Component Separation" },
  { code: "15734", desc: "Fascial flap", category: "Component Separation" },

  // --- Other procedures ---
  { code: "36556", desc: "Central line", category: "Other" },

  // --- Spine ---
  { code: "22558", desc: "ALIF", category: "Spine" },
  { code: "22585", desc: "ALIF, each additional level", category: "Spine" },

  // --- Plastics ---
  { code: "15830", desc: "Panniculectomy", category: "Plastics" },
  { code: "15847", desc: "Abdominoplasty", category: "Plastics" },

  // --- E&M / consults ---
  { code: "99221", desc: "Admit H&P — low MDM, 20 min", category: "E&M" },
  { code: "99222", desc: "Admit H&P — moderate MDM, 55 min", category: "E&M" },
  { code: "99223", desc: "Admit H&P — high MDM, 75 min", category: "E&M" },
  { code: "99232", desc: "Daily care", category: "E&M" },
  { code: "99238", desc: "Discharge", category: "E&M" },
  { code: "99252", desc: "Consult — limited", category: "E&M" },
  { code: "99253", desc: "Consult — low MDM, 45 min", category: "E&M" },
  { code: "99254", desc: "Consult — moderate MDM, 60 min", category: "E&M" },
  { code: "99255", desc: "Consult — high MDM, 80 min", category: "E&M" },
];

// CPT modifiers (sheet: 50, 22, 57, 62; 80/LT/RT kept — 80 is auto-set for Assistant role)
const CPT_MODIFIERS = [
  { code: "80", label: "Assistant surgeon" },
  { code: "62", label: "Co-surgeon" },
  { code: "22", label: "Complicated procedure" },
  { code: "50", label: "Bilateral procedure" },
  { code: "57", label: "Decision for surgery" },
  { code: "LT", label: "Left side" },
  { code: "RT", label: "Right side" },
];

const ICD10_FAVORITES = [
  // --- Abdominal pain ---
  { code: "R10.0", desc: "Abdominal pain — RUQ / RLQ / LUQ / LLQ / epigastric / umbilical / general", category: "GI" },

  // --- Gallbladder ---
  { code: "K81.0", desc: "Cholecystitis, acute", category: "Gallbladder" },
  { code: "K80.80", desc: "Cholelithiasis, no obstruction", category: "Gallbladder" },
  { code: "K82.4", desc: "Cholesterolosis", category: "Gallbladder" },
  { code: "K82.8", desc: "Biliary dyskinesia", category: "Gallbladder" },

  // --- Metabolic ---
  { code: "E11.9", desc: "DM II, uncomplicated", category: "Metabolic" },
  { code: "E78.0", desc: "Hypercholesterolemia", category: "Metabolic" },
  { code: "E66.01", desc: "Morbid obesity", category: "Metabolic" },
  { code: "E66.9", desc: "Obesity, unspecified", category: "Metabolic" },
  { code: "K76.0", desc: "Fatty liver", category: "Metabolic" },
  { code: "E43", desc: "Malnutrition, severe", category: "Metabolic" },

  // --- Musculoskeletal ---
  { code: "M54.5", desc: "Lumbago (back pain)", category: "MSK" },
  { code: "M25.50", desc: "Arthralgia", category: "MSK" },
  { code: "M15.0", desc: "Degenerative joint disease", category: "MSK" },

  // --- Cardiac ---
  { code: "I10", desc: "Hypertension", category: "Cardiac" },
  { code: "I51.9", desc: "Organic heart disease", category: "Cardiac" },
  { code: "I50.20", desc: "Congestive heart failure", category: "Cardiac" },
  { code: "I25.9", desc: "Ischemic heart disease", category: "Cardiac" },
  { code: "I25.10", desc: "Coronary artery disease", category: "Cardiac" },

  // --- Hernias ---
  { code: "K45.8", desc: "Internal hernia", category: "Hernia" },
  { code: "K44.9", desc: "Hiatal / paraesophageal hernia", category: "Hernia" },
  { code: "K42.9", desc: "Umbilical hernia", category: "Hernia" },
  { code: "K42.0", desc: "Umbilical hernia, incarcerated", category: "Hernia" },
  { code: "K43.9", desc: "Ventral / epigastric hernia", category: "Hernia" },
  { code: "K43.2", desc: "Incisional hernia", category: "Hernia" },
  { code: "K43.0", desc: "Incisional hernia with obstruction", category: "Hernia" },
  { code: "K40.90", desc: "Inguinal hernia, without obstruction", category: "Hernia" },
  { code: "K40.20", desc: "Inguinal hernia, bilateral", category: "Hernia" },
  { code: "K40.91", desc: "Inguinal hernia, recurrent", category: "Hernia" },
  { code: "K41.9", desc: "Femoral hernia, unilateral", category: "Hernia" },

  // --- Ulcers ---
  { code: "K27.7", desc: "Peptic ulcer, chronic, no obstruction", category: "Ulcer" },
  { code: "K27.9", desc: "Peptic ulcer, acute or chronic, without bleed/perforation", category: "Ulcer" },
  { code: "K27.3", desc: "Peptic ulcer, acute, without perforation", category: "Ulcer" },
  { code: "K28", desc: "GJ ulcer", category: "Ulcer" },
  { code: "K28.0", desc: "GJ ulcer with bleed", category: "Ulcer" },
  { code: "K28.1", desc: "GJ ulcer with perforation", category: "Ulcer" },

  // --- Respiratory ---
  { code: "J45.909", desc: "Asthma", category: "Respiratory" },
  { code: "J44.9", desc: "COPD", category: "Respiratory" },
  { code: "G47.30", desc: "Sleep apnea", category: "Respiratory" },
  { code: "I26.09", desc: "Pulmonary embolus", category: "Respiratory" },
  { code: "I82.4", desc: "DVT, lower extremity", category: "Respiratory" },

  // --- GI ---
  { code: "K21.0", desc: "Reflux esophagitis", category: "GI" },
  { code: "K21.9", desc: "GERD", category: "GI" },
  { code: "K29.00", desc: "Gastritis / duodenitis", category: "GI" },
  { code: "K31.1", desc: "Gastric outlet obstruction / GJ stricture or dysfunction", category: "GI" },
  { code: "K31.84", desc: "Gastroparesis", category: "GI" },
  { code: "K35.2", desc: "Appendicitis, ruptured", category: "GI" },
  { code: "K35.80", desc: "Appendicitis, acute", category: "GI" },
  { code: "K52.9", desc: "Gastroenteritis", category: "GI" },
  { code: "K56.1", desc: "Intussusception", category: "GI" },
  { code: "K56.51", desc: "SBO, partial", category: "GI" },
  { code: "K56.52", desc: "SBO, complete", category: "GI" },
  { code: "K59.0", desc: "Constipation", category: "GI" },
  { code: "K63.2", desc: "Intestinal fistula", category: "GI" },
  { code: "K65.1", desc: "Peritoneal abscess", category: "GI" },
  { code: "K66.1", desc: "Hemoperitoneum", category: "GI" },
  { code: "K85.1", desc: "Pancreatitis, biliary", category: "GI" },
  { code: "K91.89", desc: "Anastomotic leak (post-procedure complication)", category: "GI" },
  { code: "R11.2", desc: "Nausea & vomiting", category: "GI" },
  { code: "R13.10", desc: "Dysphagia", category: "GI" },
  { code: "R19.7", desc: "Diarrhea", category: "GI" },

  // --- Spleen ---
  { code: "D73.89", desc: "Ruptured spleen", category: "Spleen" },
  { code: "S36.0", desc: "Injured spleen (unspecified)", category: "Spleen" },

  // --- Skin ---
  { code: "L98.9", desc: "Skin lesion", category: "Skin" },
  { code: "L72.3", desc: "Sebaceous cyst", category: "Skin" },
  { code: "D17", desc: "Lipoma", category: "Skin" },
];

// Facility list (from billing sheet header)
const FACILITIES = ["SM", "SMOPS", "IMC", "LDS", "SLR"];

if (typeof module !== "undefined" && module.exports) {
  module.exports = { FAVORITES_VERSION, CPT_FAVORITES, CPT_MODIFIERS, ICD10_FAVORITES, FACILITIES };
}
