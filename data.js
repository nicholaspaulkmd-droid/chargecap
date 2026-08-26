// data.js
// CPT and ICD-10 favorite code libraries for ChargeCap.
// Transcribed directly from Dr. Paulk's OR billing sheets
// ("Billing Sheet - OR.xlsx" and "Billing Sheet - 11.04.22.xlsx").
//
// IMPORTANT: These codes were imported verbatim from existing paper/Excel
// billing sheets that may contain transcription errors or outdated codes.
// Verify against current CPT/ICD-10-CM documentation before relying on
// them for claims submission. Codes can also be edited any time in the
// app under Settings > Manage Code Library (stored in IndexedDB, so
// edits persist locally).

const CPT_FAVORITES = [
  // --- Bariatric ---
  { code: "43644", desc: "Gastric bypass, Roux-en-Y — LAP", category: "Bariatric" },
  { code: "43846", desc: "Gastric bypass, Roux-en-Y — OPEN", category: "Bariatric" },
  { code: "43775", desc: "Sleeve gastrectomy — LAP", category: "Bariatric" },
  { code: "43843", desc: "Gastric restrictive procedure — OPEN", category: "Bariatric" },
  { code: "43659", desc: "Duodenal switch (unlisted code) — LAP", category: "Bariatric" },
  { code: "43845", desc: "Duodenal switch / BPD — OPEN", category: "Bariatric" },
  { code: "43770", desc: "Lap band placement", category: "Bariatric" },
  { code: "43771", desc: "Lap band revision only", category: "Bariatric" },
  { code: "43772", desc: "Lap band removal only", category: "Bariatric" },
  { code: "43773", desc: "Lap band removal & replacement", category: "Bariatric" },
  { code: "43774", desc: "Lap band removal, including components", category: "Bariatric" },

  // --- Revision ---
  { code: "43860", desc: "Revision of gastrojejunal anastomosis — OPEN", category: "Revision" },
  { code: "44120", desc: "Revision JJ — enterectomy, single resection", category: "Revision" },
  { code: "44130", desc: "Revision JJ — enteroenterostomy", category: "Revision" },
  { code: "43360", desc: "Revision of gastric restrictive procedure (VBG) — OPEN", category: "Revision" },
  { code: "43659", desc: "Gastric bypass revision (unlisted code) — LAP", category: "Revision" },
  { code: "43848", desc: "Gastric restrictive procedure revision — OPEN", category: "Revision" },
  { code: "43631 + 43659 + 44202", desc: "Bypass-to-DS conversion: subtotal gastrectomy, DS, small bowel resection", category: "Revision" },

  // --- Hernia (all LAP include mesh per billing sheet) ---
  { code: "49650", desc: "Inguinal hernia repair — LAP", category: "Hernia" },
  { code: "49505", desc: "Inguinal hernia repair — OPEN", category: "Hernia" },
  { code: "49507", desc: "Inguinal hernia repair, incarcerated — OPEN", category: "Hernia" },
  { code: "49651", desc: "Inguinal hernia repair, recurrent — LAP", category: "Hernia" },
  { code: "49520", desc: "Inguinal hernia repair, recurrent — OPEN", category: "Hernia" },
  { code: "49550", desc: "Femoral hernia repair — OPEN", category: "Hernia" },
  { code: "49652", desc: "Ventral / umbilical / spigelian / epigastric hernia repair — LAP", category: "Hernia" },
  { code: "49653", desc: "Ventral / umbilical / spigelian hernia repair, incarcerated — LAP", category: "Hernia" },
  { code: "49560", desc: "Ventral or incisional hernia repair — OPEN", category: "Hernia" },
  { code: "49565", desc: "Ventral or incisional hernia repair, recurrent — OPEN", category: "Hernia" },
  { code: "49570", desc: "Epigastric hernia repair — OPEN", category: "Hernia" },
  { code: "49585", desc: "Umbilical hernia repair — OPEN", category: "Hernia" },
  { code: "49587", desc: "Umbilical hernia repair, incarcerated — OPEN", category: "Hernia" },
  { code: "49654", desc: "Incisional hernia repair — LAP", category: "Hernia" },
  { code: "49655", desc: "Incisional hernia repair, incarcerated — LAP", category: "Hernia" },
  { code: "49656", desc: "Recurrent incisional hernia repair — LAP", category: "Hernia" },
  { code: "49657", desc: "Recurrent incisional hernia repair, incarcerated — LAP", category: "Hernia" },
  { code: "43281", desc: "Paraesophageal / hiatal hernia repair — LAP", category: "Hernia" },
  { code: "43282", desc: "Paraesophageal / hiatal hernia repair with mesh — LAP", category: "Hernia" },
  { code: "49568", desc: "Mesh implantation (incisional/ventral hernia, add-on)", category: "Hernia" },

  // --- Cholecystectomy ---
  { code: "47562", desc: "Cholecystectomy — LAP", category: "Cholecystectomy" },
  { code: "47563", desc: "Cholecystectomy with cholangiography — LAP", category: "Cholecystectomy" },
  { code: "47564", desc: "Cholecystectomy with common duct exploration — LAP", category: "Cholecystectomy" },
  { code: "47600", desc: "Cholecystectomy — OPEN", category: "Cholecystectomy" },
  { code: "47605", desc: "Cholecystectomy with cholangiography — OPEN", category: "Cholecystectomy" },
  { code: "47610", desc: "Cholecystectomy with common duct exploration — OPEN", category: "Cholecystectomy" },

  // --- Appendix ---
  { code: "44970", desc: "Appendectomy — LAP", category: "Appendix" },
  { code: "44960", desc: "Appendectomy, ruptured with abscess — OPEN", category: "Appendix" },
  { code: "44950", desc: "Appendectomy — OPEN", category: "Appendix" },

  // --- Bowel ---
  { code: "44005", desc: "Enterolysis (freeing of intestinal adhesions) — OPEN", category: "Bowel" },
  { code: "44180", desc: "Enterolysis (freeing of intestinal adhesions) — LAP", category: "Bowel" },
  { code: "44050", desc: "Reduction of intestinal intussusception / internal hernia repair — OPEN", category: "Bowel" },
  { code: "44120", desc: "Enterectomy, single resection with anastomosis", category: "Bowel" },
  { code: "44121", desc: "Enterectomy, each additional resection (add-on)", category: "Bowel" },
  { code: "44130", desc: "Enteroenterostomy", category: "Bowel" },
  { code: "44202", desc: "Small bowel resection with anastomosis — LAP", category: "Bowel" },
  { code: "44203", desc: "Small bowel resection, each additional (add-on) — LAP", category: "Bowel" },
  { code: "44160", desc: "Colectomy, partial, with anastomosis — OPEN", category: "Bowel" },
  { code: "44140", desc: "Colectomy, partial — OPEN", category: "Bowel" },
  { code: "44143", desc: "Colectomy, partial, with end colostomy (Hartmann)", category: "Bowel" },
  { code: "44145", desc: "Colectomy, partial, with coloproctostomy", category: "Bowel" },
  { code: "44205", desc: "Colectomy, partial, with anastomosis — LAP", category: "Bowel" },
  { code: "49320", desc: "Exploratory laparoscopy, diagnostic", category: "Bowel" },
  { code: "49321", desc: "Exploratory laparoscopy with biopsy", category: "Bowel" },
  { code: "49322", desc: "Laparoscopic aspiration of cavity or cyst", category: "Bowel" },
  { code: "49326", desc: "Laparoscopic omentopexy (Graham patch)", category: "Bowel" },
  { code: "49020", desc: "Drainage of peritoneal abscess — OPEN", category: "Bowel" },
  { code: "44602", desc: "Enterorrhaphy for wound / injury / ulcer perforation", category: "Bowel" },
  { code: "58805", desc: "Ovarian cyst excision", category: "Bowel" },
  { code: "38100", desc: "Splenectomy — OPEN", category: "Bowel" },
  { code: "38120", desc: "Splenectomy — LAP", category: "Bowel" },

  // --- Upper GI ---
  { code: "43280", desc: "Nissen fundoplication — LAP", category: "Upper GI" },
  { code: "43324", desc: "Fundoplication — OPEN", category: "Upper GI" },
  { code: "43633", desc: "Gastrectomy, subtotal, with Roux-en-Y reconstruction", category: "Upper GI" },
  { code: "43653", desc: "Gastrostomy — LAP", category: "Upper GI" },
  { code: "43830", desc: "Gastrostomy — OPEN", category: "Upper GI" },
  { code: "43832", desc: "Gastrostomy tube placement", category: "Upper GI" },
  { code: "44186", desc: "Jejunostomy tube placement — LAP", category: "Upper GI" },
  { code: "36556", desc: "Central line placement", category: "Upper GI" },
  { code: "43840", desc: "Repair of gastric / duodenal ulcer", category: "Upper GI" },

  // --- Endoscopy ---
  { code: "43235", desc: "EGD — diagnostic / intraoperative", category: "Endoscopy" },
  { code: "43239", desc: "EGD with biopsy", category: "Endoscopy" },
  { code: "43245", desc: "EGD with dilation", category: "Endoscopy" },
  { code: "43247", desc: "EGD with foreign body removal", category: "Endoscopy" },
  { code: "34266", desc: "EGD with stent placement", category: "Endoscopy" },

  // --- HPB (Hepatopancreatobiliary) ---
  { code: "47001", desc: "Liver biopsy, needle", category: "HPB" },
  { code: "47100", desc: "Liver biopsy, wedge", category: "HPB" },
  { code: "15271", desc: "Biologic implant, first 100 sq cm", category: "HPB" },
  { code: "15734", desc: "Fascial / falciform flap", category: "HPB" },

  // --- Thyroid / Parathyroid ---
  { code: "60240", desc: "Thyroidectomy, total", category: "Thyroid / Parathyroid" },
  { code: "60210", desc: "Thyroidectomy, partial", category: "Thyroid / Parathyroid" },
  { code: "60500", desc: "Parathyroidectomy", category: "Thyroid / Parathyroid" },

  // --- E&M ---
  { code: "99221", desc: "Initial hospital admit — limited", category: "E&M" },
  { code: "99222", desc: "Initial hospital admit — moderate", category: "E&M" },
  { code: "99223", desc: "Initial hospital admit — complex", category: "E&M" },
  { code: "99232", desc: "Subsequent hospital daily care", category: "E&M" },
  { code: "99238", desc: "Hospital discharge day management", category: "E&M" },
  { code: "99252", desc: "Inpatient consult — limited", category: "E&M" },
  { code: "99253", desc: "Inpatient consult — detailed", category: "E&M" },
  { code: "99254", desc: "Inpatient consult — extended", category: "E&M" },
  { code: "99255", desc: "Inpatient consult — comprehensive", category: "E&M" },

  // --- Component separation / central line ---
  { code: "15734", desc: "Fascial flap (component separation)", category: "Component Separation" },
  { code: "15271", desc: "Biologic implant, incisional/ventral hernia, >100 cm2", category: "Component Separation" },
  { code: "64488", desc: "TAP block", category: "Component Separation" },
  { code: "36556", desc: "Central line placement", category: "Component Separation" },
];

// CPT modifiers (surgeon role + common billing modifiers, from billing sheet "Modifiers" section)
const CPT_MODIFIERS = [
  { code: "80", label: "Assistant surgeon" },
  { code: "62", label: "Co-surgeon" },
  { code: "22", label: "Complicated procedure" },
  { code: "50", label: "Bilateral procedure" },
  { code: "LT", label: "Left side" },
  { code: "RT", label: "Right side" },
];

const ICD10_FAVORITES = [
  // --- GI ---
  { code: "R10.0", desc: "Abdominal pain, RUQ/RLQ/LUQ/LLQ", category: "GI" },
  { code: "K52.9", desc: "Gastroenteritis, unspecified", category: "GI" },
  { code: "K29.00", desc: "Acute gastritis without bleeding", category: "GI" },
  { code: "K21.0", desc: "Reflux esophagitis (GERD with esophagitis)", category: "GI" },
  { code: "K21.9", desc: "GERD without esophagitis", category: "GI" },
  { code: "K31.1", desc: "Gastric outlet / GJ obstruction or dysfunction", category: "GI" },
  { code: "K35.80", desc: "Acute appendicitis, unspecified", category: "GI" },
  { code: "K35.2", desc: "Acute appendicitis, ruptured, generalized peritonitis", category: "GI" },
  { code: "K35.3", desc: "Acute appendicitis with localized peritonitis", category: "GI" },
  { code: "K56.5", desc: "Intestinal adhesions with obstruction (small bowel obstruction)", category: "GI" },
  { code: "K56.1", desc: "Intussusception", category: "GI" },
  { code: "K59.0", desc: "Constipation", category: "GI" },
  { code: "K63.2", desc: "Fistula of intestine", category: "GI" },
  { code: "K65.1", desc: "Peritoneal abscess", category: "GI" },
  { code: "K66.0", desc: "Peritoneal adhesions (abdominal)", category: "GI" },
  { code: "K66.1", desc: "Hemoperitoneum", category: "GI" },
  { code: "K91.89", desc: "Anastomotic leak (postprocedural complication)", category: "GI" },
  { code: "R11.2", desc: "Nausea with vomiting, unspecified", category: "GI" },
  { code: "R11.0", desc: "Nausea", category: "GI" },
  { code: "R11.10", desc: "Vomiting, unspecified, persistent", category: "GI" },
  { code: "R13.10", desc: "Dysphagia, unspecified", category: "GI" },
  { code: "R19.7", desc: "Diarrhea, unspecified", category: "GI" },
  { code: "K58.0", desc: "Irritable bowel syndrome with diarrhea", category: "GI" },
  { code: "K58.9", desc: "Irritable bowel syndrome without diarrhea", category: "GI" },
  { code: "D51.0", desc: "Vitamin B12 deficiency anemia", category: "GI" },
  { code: "K76.0", desc: "Fatty liver, not elsewhere classified", category: "GI" },

  // --- Gallbladder ---
  { code: "K80.80", desc: "Cholelithiasis without obstruction", category: "Gallbladder" },
  { code: "K80.18", desc: "Cholelithiasis with acute cholecystitis, no obstruction", category: "Gallbladder" },
  { code: "K80.00", desc: "Cholelithiasis with acute cholecystitis, with obstruction", category: "Gallbladder" },
  { code: "K80.66", desc: "Cholelithiasis with acute and chronic cholecystitis", category: "Gallbladder" },
  { code: "K81.0", desc: "Acute cholecystitis", category: "Gallbladder" },
  { code: "K81.1", desc: "Chronic cholecystitis", category: "Gallbladder" },
  { code: "K81.2", desc: "Acute and chronic cholecystitis", category: "Gallbladder" },
  { code: "K81.9", desc: "Cholecystitis, unspecified", category: "Gallbladder" },
  { code: "K82.4", desc: "Cholesterolosis of gallbladder", category: "Gallbladder" },
  { code: "K82.8", desc: "Biliary dyskinesia", category: "Gallbladder" },
  { code: "K85.10", desc: "Biliary acute pancreatitis", category: "Gallbladder" },

  // --- Hernia ---
  { code: "K40.90", desc: "Inguinal hernia, unilateral, without obstruction", category: "Hernia" },
  { code: "K40.30", desc: "Inguinal hernia, unilateral, with obstruction", category: "Hernia" },
  { code: "K40.91", desc: "Inguinal hernia, unilateral, recurrent", category: "Hernia" },
  { code: "K40.20", desc: "Inguinal hernia, bilateral", category: "Hernia" },
  { code: "K41.9", desc: "Femoral hernia, unilateral", category: "Hernia" },
  { code: "K42.9", desc: "Umbilical hernia without obstruction", category: "Hernia" },
  { code: "K42.0", desc: "Umbilical hernia, incarcerated", category: "Hernia" },
  { code: "K43.9", desc: "Ventral / epigastric hernia without obstruction", category: "Hernia" },
  { code: "K43.2", desc: "Incisional hernia, recurrent", category: "Hernia" },
  { code: "K43.0", desc: "Incisional hernia with obstruction", category: "Hernia" },
  { code: "K44.9", desc: "Hiatal / paraesophageal hernia without obstruction", category: "Hernia" },
  { code: "K45.8", desc: "Internal hernia (other specified abdominal hernia)", category: "Hernia" },
  { code: "K46.9", desc: "Epigastric hernia, unspecified", category: "Hernia" },

  // --- Metabolic ---
  { code: "E66.01", desc: "Morbid (severe) obesity due to excess calories", category: "Metabolic" },
  { code: "E66.9", desc: "Obesity, unspecified", category: "Metabolic" },
  { code: "E66.3", desc: "Overweight (BMI 25–29.9)", category: "Metabolic" },
  { code: "E11.9", desc: "Type 2 diabetes mellitus without complications", category: "Metabolic" },
  { code: "E78.0", desc: "Pure hypercholesterolemia", category: "Metabolic" },
  { code: "E43", desc: "Severe protein-calorie malnutrition", category: "Metabolic" },

  // --- Cardiac ---
  { code: "I10", desc: "Essential (primary) hypertension", category: "Cardiac" },
  { code: "I51.9", desc: "Heart disease, unspecified", category: "Cardiac" },
  { code: "I50.20", desc: "Congestive heart failure, unspecified", category: "Cardiac" },
  { code: "I25.9", desc: "Chronic ischemic heart disease, unspecified", category: "Cardiac" },
  { code: "I25.10", desc: "Coronary artery disease, native vessel", category: "Cardiac" },
  { code: "I82.4", desc: "Deep vein thrombosis, lower extremity", category: "Cardiac" },
  { code: "I26.09", desc: "Pulmonary embolism with acute cor pulmonale", category: "Cardiac" },
  { code: "I87.2", desc: "Venous insufficiency (chronic, peripheral)", category: "Cardiac" },

  // --- Respiratory ---
  { code: "J45.909", desc: "Asthma, unspecified, uncomplicated", category: "Respiratory" },
  { code: "J44.9", desc: "Chronic obstructive pulmonary disease, unspecified", category: "Respiratory" },
  { code: "G47.30", desc: "Sleep apnea, unspecified", category: "Respiratory" },
  { code: "R06.00", desc: "Dyspnea, unspecified", category: "Respiratory" },
  { code: "R06.83", desc: "Snoring", category: "Respiratory" },

  // --- Ulcer ---
  { code: "K27.9", desc: "Peptic ulcer, unspecified, without hemorrhage/perforation", category: "Ulcer" },
  { code: "K27.7", desc: "Peptic ulcer, chronic, without obstruction", category: "Ulcer" },
  { code: "K27.3", desc: "Peptic ulcer, acute, without hemorrhage/perforation", category: "Ulcer" },
  { code: "K28", desc: "Gastrojejunal ulcer, unspecified", category: "Ulcer" },
  { code: "K28.0", desc: "Gastrojejunal ulcer, acute, with hemorrhage", category: "Ulcer" },
  { code: "K28.1", desc: "Gastrojejunal ulcer, acute, with perforation", category: "Ulcer" },
  { code: "K28.3", desc: "Gastrojejunal ulcer, acute, with obstruction", category: "Ulcer" },
  { code: "K28.4", desc: "Gastrojejunal ulcer, chronic, with obstruction", category: "Ulcer" },

  // --- MSK ---
  { code: "M54.5", desc: "Low back pain", category: "MSK" },
  { code: "M25.50", desc: "Arthralgia, unspecified joint", category: "MSK" },
  { code: "M15.0", desc: "Primary generalized (osteo)arthritis / degenerative joint disease", category: "MSK" },

  // --- Other ---
  { code: "L98.9", desc: "Skin lesion, unspecified", category: "Other" },
  { code: "L72.3", desc: "Sebaceous cyst", category: "Other" },
  { code: "D17", desc: "Benign lipomatous neoplasm (lipoma)", category: "Other" },
  { code: "N83.20", desc: "Ovarian cyst, unspecified", category: "Other" },
  { code: "Z30.2", desc: "Encounter for sterilization (desired)", category: "Other" },
  { code: "F32.9", desc: "Major depressive disorder, single episode, unspecified", category: "Other" },
  { code: "D73.89", desc: "Ruptured spleen (other diseases of spleen)", category: "Other" },
  { code: "S36.00XA", desc: "Injury of spleen, unspecified, initial encounter", category: "Other" },
];

// Facility list (from billing sheet header)
const FACILITIES = ["SM", "SMOPS", "IMC", "LDS", "SLR"];

if (typeof module !== "undefined" && module.exports) {
  module.exports = { CPT_FAVORITES, CPT_MODIFIERS, ICD10_FAVORITES, FACILITIES };
}
