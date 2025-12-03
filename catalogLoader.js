// catalogLoader.js
import xlsx from "xlsx";
import path from "path";
import fs from "fs";

// cartella dati
const dataDir = path.join(process.cwd(), "data");
const asmePath = path.join(dataDir, "ASME_BPE.xlsx");
const teesReducersPath = path.join(dataDir, "Tees_Reducers_ASME_BPE.xlsx");

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
    ? xlsx.utils.sheet_to_json(wb.Sheets[name], { defval: null })
    : [];
}

function pick(row, keys) {
  for (const k of keys) {
    if (Object.prototype.hasOwnProperty.call(row, k)) return row[k];
  }
  return null;
}

// Finiture disponibili
export const FINISHES = [
  {
    key: "ASME BPE SF1",
    tubePriceKeys: ["SF1 €/m", "SF1 €/mt", "SF1 €/m ", "SF1 €/MT"],
    fittingPriceKeys: ["SF1 €/pc", "SF1 €/pz", "SF1 €/piece"],
  },
  {
    key: "ASME BPE SF4",
    tubePriceKeys: ["SF4 €/m", "SF4 €/mt", "SF4 €/m ", "SF4 €/MT"],
    fittingPriceKeys: ["SF4 €/pc", "SF4 €/pz", "SF4 €/piece"],
  },
];

// ---------- TUBES (ND, peso kg/m, €/m per finitura) ----------

function loadTubesCatalog() {
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
    const NDraw = pick(r, ["ND", "DN", "Nominal Diameter"]);
    if (!NDraw) continue;
    const ND = String(NDraw).trim();
    if (!ND) continue;

    const codeRaw = pick(r, ["Code", "Item Code", "Codice"]);
    const code = codeRaw != null ? String(codeRaw).trim() : "";

    const pesoKgM = parseNum(pick(r, ["Peso Kg/m", "Weight kg/m", "Kg/m"]));

    for (const finish of FINISHES) {
      const price = parseNum(pick(r, finish.tubePriceKeys));
      if (price <= 0) continue;

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
}

// ---------- FITTINGS SEMPLICI (solo ND) ----------
// Elbows 90°, Elbows 45°, End Caps, Ferrule A/B/C

const SIMPLE_SHEETS = [
  {
    itemType: "Elbows 90°",
    sheetNames: ["Elbows 90°", "Elbows 90", "90° Elbows"],
  },
  {
    itemType: "Elbows 45°",
    sheetNames: ["Elbows 45°", "Elbows 45", "45° Elbows"],
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
];

function loadSimpleFittingsCatalog() {
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

    const rows = sheetJSON(wb, sheetName);

    for (const r of rows) {
      const NDraw = pick(r, ["ND", "DN", "Nominal Diameter"]);
      if (!NDraw) continue;
      const ND = String(NDraw).trim();
      if (!ND) continue;

      const codeRaw = pick(r, ["Code", "Item Code", "Codice"]);
      const code = codeRaw != null ? String(codeRaw).trim() : "";

      for (const finish of FINISHES) {
        const price = parseNum(pick(r, finish.fittingPriceKeys));
        if (price <= 0) continue;

        out.push({
          itemType: def.itemType,
          finish: finish.key,
          ND,
          code,
          pricePerPc: price,
        });
      }
    }
  }

  return out;
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
      const OD1raw = pick(r, ["OD1", "OD min", "OD smaller", "DN1"]);
      const OD2raw = pick(r, ["OD2", "OD max", "OD bigger", "DN2"]);
      if (OD1raw == null || OD2raw == null) continue;

      let OD1 = String(OD1raw).trim();
      let OD2 = String(OD2raw).trim();
      if (!OD1 || !OD2) continue;

      // assicuro OD1 = diametro minore e OD2 = maggiore (se numerici)
      const n1 = parseNum(OD1);
      const n2 = parseNum(OD2);
      if (n1 && n2 && n1 > n2) {
        const tmp = OD1;
        OD1 = OD2;
        OD2 = tmp;
      }

      const codeRaw = pick(r, ["Code", "Item Code", "Codice"]);
      const code = codeRaw != null ? String(codeRaw).trim() : "";

      for (const finish of FINISHES) {
        const price = parseNum(pick(r, finish.fittingPriceKeys));
        if (price <= 0) continue;

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
}

// ---------- entry point per l'app ----------

export function loadCatalog() {
  const tubes = loadTubesCatalog();
  const simple = loadSimpleFittingsCatalog();
  const complex = loadComplexCatalog();

  return {
    tubes, // Tubes (solo ND)
    simple, // Elbows 90°, Elbows 45°, End Caps, Ferrule A/B/C (solo ND)
    complex, // Tees, Conc. Reducers, Ecc. Reducers (OD1/OD2)
  };
}
