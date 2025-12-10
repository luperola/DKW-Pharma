// catalogLoader.js
import xlsx from "xlsx";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

/// cartella dati (risolta rispetto a questo file, non alla cwd)
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const dataDir = path.join(__dirname, "data");
const asmePath = path.join(dataDir, "ASME_BPE.xlsx");
const teesReducersPath = path.join(dataDir, "Tees and reducers.xlsx");

// ---------- utils comuni ----------

function parseNum(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  let s = String(v).trim();
  if (!s) return 0;
  const hasDot = s.includes(".");
  const hasComma = s.includes(",");
  if (hasDot && hasComma) {
    if (s.lastIndexOf(",") > s.lastIndexOf("."))
      s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (hasComma) s = s.replace(",", ".");
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function sheetJSON(wb, name) {
  return wb.Sheets[name]
    ? xlsx.utils.sheet_to_json(wb.Sheets[name], { defval: null, range: 1 })
    : [];
}

function pick(row, keys) {
  for (const k of keys) {
    if (Object.prototype.hasOwnProperty.call(row, k)) return row[k];
  }
  return null;
}

function formatNumber(num) {
  const n = parseNum(num);
  if (!n) return (num ?? "").toString().trim();
  if (Number.isInteger(n)) return String(n);
  return String(Number.parseFloat(n.toFixed(3))).replace(/\.0+$/, "");
}

function formatDimension(mm, inch) {
  const mmStr = formatNumber(mm);
  const inchStr = (inch ?? "").toString().trim();

  if (!mmStr && !inchStr) return "";
  if (mmStr && inchStr) return `${mmStr} mm (${inchStr})`;
  if (mmStr) return `${mmStr} mm`;
  return inchStr;
}

// Finiture disponibili
export const FINISHES = [
  {
    key: "ASME BPE SF1",
    tubePriceKeys: [
      "SF1 €/m",
      "SF1 €/mt",
      "SF1 €/m ",
      "SF1 €/MT",
      " ASME BPE SF1 (BF) Price in € / m",
    ],
    tubeCodeKeys: ["Code", "Item Code", "Codice", " ASME BPE SF1 (BF) Code"],
    fittingPriceKeys: [
      "SF1 €/pc",
      "SF1 €/pz",
      "SF1 €/piece",
      " ASME BPE SF1 (BF) Price in € / m",
    ],
    fittingCodeKeys: ["Code", "Item Code", "Codice", " ASME BPE SF1 (BF) Code"],
  },
  {
    key: "ASME BPE SF4",
    tubePriceKeys: [
      "SF4 €/m",
      "SF4 €/mt",
      "SF4 €/m ",
      "SF4 €/MT",
      " ASME BPE SF4 (BF) Price in € / m",
    ],
    tubeCodeKeys: ["Code", "Item Code", "Codice", " ASME BPE SF4 (BF) Code"],
    fittingPriceKeys: [
      "SF4 €/pc",
      "SF4 €/pz",
      "SF4 €/piece",
      " ASME BPE SF4 (BF) Price in € / m",
    ],
    fittingCodeKeys: ["Code", "Item Code", "Codice", " ASME BPE SF4 (BF) Code"],
  },
];

// ---------- TUBES (ND, peso kg/m, €/m per finitura) ----------

function loadTubesCatalog() {
  try {
    if (!fs.existsSync(asmePath)) return [];
    const wb = xlsx.readFile(asmePath);

    // prova prima "Tubes", altrimenti primo foglio
    const sheetName = wb.Sheets["Tubes"]
      ? "Tubes"
      : wb.SheetNames.find((n) => n.toLowerCase().includes("tube")) ||
        wb.SheetNames[0];

    const rows = sheetJSON(wb, sheetName);
    const out = [];

    for (const r of rows) {
      const inchRaw = pick(r, ["Inch", "DN inch", "ND inch"]);
      const mmRaw = pick(r, ["mm", "DN mm", "ND mm"]);
      const ND = formatDimension(mmRaw, inchRaw);
      if (!ND) continue;

      for (const finish of FINISHES) {
        const price = parseNum(pick(r, finish.tubePriceKeys));
        if (price <= 0) continue;

        const codeRaw = pick(r, finish.tubeCodeKeys);
        const code = codeRaw != null ? String(codeRaw).trim() : "";

        const pesoKgM = parseNum(
          pick(r, ["Peso Kg/m", "Peso kg/m", "Weight kg/m", "Kg/m"])
        );

        out.push({
          itemType: "Tubes",
          finish: finish.key,
          ND,
          code,
          pesoKgM,
          pricePerM: price,
        });
      }
    }

    return out;
  } catch (err) {
    console.error("Errore nel caricamento del catalogo Tubes:", err);
    return [];
  }
}

// ---------- FITTINGS SEMPLICI (solo ND) ----------
// Elbows 90°, Elbows 45°, End Caps, Ferrule A/B/C

const SIMPLE_SHEETS = [
  {
    itemType: "Elbows 90°",
    sheetNames: [
      "Elbows 90°",
      "Elbows 90",
      "90° Elbows",
      "Elbow 90°",
      "Elbow 90",
      "90° Elbow",
    ],
  },
  {
    itemType: "Elbows 45°",
    sheetNames: [
      "Elbows 45°",
      "Elbows 45",
      "45° Elbows",
      "Elbow 45°",
      "Elbow 45",
      "45° Elbow",
    ],
  },
  {
    itemType: "End Caps",
    sheetNames: ["End Caps", "Caps", "EndCap"],
  },
  {
    itemType: "Ferrule A (Long)",
    sheetNames: ["Ferrule A (Long)", "Ferrule A Long", "Ferrule A"],
  },
  {
    itemType: "Ferrule B (Medium)",
    sheetNames: ["Ferrule B (Medium)", "Ferrule B Medium", "Ferrule B"],
  },
  {
    itemType: "Ferrule C (Short)",
    sheetNames: ["Ferrule C (Short)", "Ferrule C Short", "Ferrule C"],
  },
  {
    itemType: "Clamps",
    sheetNames: ["Clamps", "Clamp"],
  },
];

// Fogli da cui leggere la lunghezza in mm per le ferrule
const FERRULE_LENGTH_SHEETS = {
  "Ferrule A (Long)": "Ferrule A (long)",
  "Ferrule B (Medium)": "Ferrule B (medium)",
  "Ferrule C (Short)": "Ferrule C (short)",
};

const FERRULE_PRICE_COLUMNS = {
  "ASME BPE SF1": "E", // prezzo in €/pz su colonna E del foglio lunghezze
  "ASME BPE SF4": "G", // prezzo in €/pz su colonna G del foglio lunghezze
};

function findSheetName(wb, targetName) {
  if (!targetName) return null;
  if (wb.Sheets[targetName]) return targetName;
  const targetLower = targetName.toLowerCase();
  return wb.SheetNames.find((n) => n.toLowerCase() === targetLower) || null;
}

function loadSimpleFittingsCatalog() {
  try {
    if (!fs.existsSync(asmePath)) return [];
    const wb = xlsx.readFile(asmePath);
    const out = [];

    for (const def of SIMPLE_SHEETS) {
      // trova il primo foglio che esiste tra quelli indicati
      const sheetName =
        def.sheetNames.find((n) => wb.Sheets[n]) ||
        wb.SheetNames.find((n) =>
          n.toLowerCase().includes(def.itemType.split(" ")[0].toLowerCase())
        );

      if (!sheetName) continue;

      const ws = wb.Sheets[sheetName];
      const rows = sheetJSON(wb, sheetName);

      for (let idx = 0; idx < rows.length; idx++) {
        const r = rows[idx];
        const inchRaw = pick(r, ["Inch", "DN inch", "ND inch"]);
        const mmRaw = pick(r, ["mm", "DN mm", "ND mm"]);
        const ND = formatDimension(mmRaw, inchRaw);
        if (!ND) continue;
        const isFerrule = def.itemType.startsWith("Ferrule");
        const lengthSheetName = isFerrule
          ? findSheetName(wb, FERRULE_LENGTH_SHEETS[def.itemType])
          : null;
        const lengthSheet = lengthSheetName ? wb.Sheets[lengthSheetName] : null;
        const lengthRaw = isFerrule
          ? lengthSheet?.[`C${idx + 3}`]?.v ||
            pick(r, [
              "Length",
              "Lenght",
              "Length mm",
              "Lenght mm",
              "Length (mm)",
            ])
          : null;
        const lengthMm = isFerrule ? parseNum(lengthRaw) : null;

        const flangeSizeRaw =
          def.itemType === "Clamps"
            ? pick(r, ["Flange size mm", "Flange size"])
            : null;
        const flangeSizeMm =
          def.itemType === "Clamps" ? parseNum(flangeSizeRaw) : null;

        for (const finish of FINISHES) {
          let priceRaw = null;
          if (isFerrule && lengthSheet && ferruleRow) {
            const priceColumn = FERRULE_PRICE_COLUMNS[finish.key];
            if (priceColumn) {
              priceRaw = lengthSheet[`${priceColumn}${ferruleRow}`]?.v ?? null;
            }
          }

          if (priceRaw == null) priceRaw = pick(r, finish.fittingPriceKeys);

          const price = parseNum(priceRaw);
          if (price <= 0) continue;

          let codeRaw = null;
          if (isFerrule && lengthSheet) {
            const ferruleColumn =
              finish.key === "ASME BPE SF1"
                ? "D"
                : finish.key === "ASME BPE SF4"
                ? "F"
                : null;

            if (ferruleColumn) {
              const ferruleCell = lengthSheet[`${ferruleColumn}${ferruleRow}`];
              if (ferruleCell && ferruleCell.v != null) {
                codeRaw = ferruleCell.v;
              }
            }
          }

          if (codeRaw == null) codeRaw = pick(r, finish.fittingCodeKeys);
          const code = codeRaw != null ? String(codeRaw).trim() : "";
          out.push({
            itemType: def.itemType,
            finish: finish.key,
            ND,
            code,
            pricePerPc: price,
            lengthMm: isFerrule && lengthMm ? lengthMm : null,
            flangeSizeMm:
              def.itemType === "Clamps" && flangeSizeMm ? flangeSizeMm : null,
          });
        }
      }
    }

    return out;
  } catch (err) {
    console.error("Errore nel caricamento del catalogo fittings (ND):", err);
    return [];
  }
}

// ---------- TEES + REDUCERS (OD1 più piccolo, OD2 più grande) ----------

const COMPLEX_SHEETS = [
  {
    itemType: "Tees",
    sheetNames: ["Tees", "Tee"],
  },
  {
    itemType: "Conc. Reducers",
    sheetNames: ["Conc. Reducers", "Concentric Reducers", "Reducers Conc."],
  },
  {
    itemType: "Ecc. Reducers",
    sheetNames: ["Ecc. Reducers", "Eccentric Reducers", "Reducers Ecc."],
  },
];

function loadComplexCatalog() {
  try {
    if (!fs.existsSync(teesReducersPath)) return [];
    const wb = xlsx.readFile(teesReducersPath);
    const out = [];

    for (const def of COMPLEX_SHEETS) {
      const sheetName =
        def.sheetNames.find((n) => wb.Sheets[n]) ||
        wb.SheetNames.find((n) =>
          n.toLowerCase().includes(def.itemType.split(" ")[0].toLowerCase())
        );

      if (!sheetName) continue;

      const rows = sheetJSON(wb, sheetName);

      for (const r of rows) {
        const inchOD1raw = pick(r, ["Inch OD1", "OD1 inch", "Inch od1"]);
        const mmOD1raw = pick(r, ["mm OD1", "OD1 mm", "mm od1"]);
        const inchOD2raw = pick(r, ["Inch OD2", "OD2 inch", "Inch od2"]);
        const mmOD2raw = pick(r, ["mm OD2", "OD2 mm", "mm od2"]);
        if (mmOD1raw == null || mmOD2raw == null) continue;

        let mm1 = parseNum(mmOD1raw);
        let mm2 = parseNum(mmOD2raw);
        let inch1 = inchOD1raw;
        let inch2 = inchOD2raw;

        if (mm1 && mm2 && mm1 > mm2) {
          [mm1, mm2] = [mm2, mm1];
          [inch1, inch2] = [inch2, inch1];
        }

        const OD1 = formatDimension(mm1 || mmOD1raw, inch1);
        const OD2 = formatDimension(mm2 || mmOD2raw, inch2);
        if (!OD1 || !OD2) continue;

        for (const finish of FINISHES) {
          const price = parseNum(pick(r, finish.fittingPriceKeys));
          if (price <= 0) continue;

          const codeRaw = pick(r, finish.fittingCodeKeys);
          const code = codeRaw != null ? String(codeRaw).trim() : "";

          out.push({
            itemType: def.itemType,
            finish: finish.key,
            OD1,
            OD2,
            code,
            pricePerPc: price,
          });
        }
      }
    }

    return out;
  } catch (err) {
    console.error("Errore nel caricamento del catalogo Tees/Reducers:", err);
    return [];
  }
}

// ---------- entry point per l'app ----------

export function loadCatalog() {
  try {
    const tubes = loadTubesCatalog();
    const simple = loadSimpleFittingsCatalog();
    const complex = loadComplexCatalog();

    return {
      tubes, // Tubes (solo ND)
      simple, // Elbows 90°, Elbows 45°, End Caps, Ferrule A/B/C (solo ND)
      complex, // Tees, Conc. Reducers, Ecc. Reducers (OD1/OD2)
    };
  } catch (err) {
    console.error("Errore generale nel caricamento del catalogo:", err);
    return { tubes: [], simple: [], complex: [] };
  }
}
