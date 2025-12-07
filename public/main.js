// public/main.js

// Stato corrente righe
let currentRows = [];

// Item che usano solo ND
const ND_ITEMS = new Set([
  "Tubes",
  "Elbows 90°",
  "Elbows 45°",
  "End Caps",
  "Clamps",
  "Ferrule A (Long)",
  "Ferrule B (Medium)",
  "Ferrule C (Short)",
]);

// Item che usano OD1 / OD2
const OD_ITEMS = new Set(["Tees", "Conc. Reducers", "Ecc. Reducers"]);

let otherItems = [];

const STILMAS_OLSA_DISCOUNT = 51.87;

function qs(id) {
  return document.getElementById(id);
}

function sortByNumericString(a, b) {
  const na = parseFloat(String(a).replace(",", "."));
  const nb = parseFloat(String(b).replace(",", "."));
  if (!isNaN(na) && !isNaN(nb)) return na - nb;
  return String(a).localeCompare(String(b));
}

function roundToDecimals(value, decimals) {
  return (
    Math.round((Number(value || 0) + Number.EPSILON) * 10 ** decimals) /
    10 ** decimals
  );
}

function computeRowUnitPrice(row) {
  const base =
    row.itemType === "Tubes"
      ? Number(row.basePricePerM || 0)
      : Number(row.basePricePerPc || 0);

  if (row.itemType !== "Tubes") return base;

  const peso = roundToDecimals(row.pesoKgM || 0, 3);
  const asKg = Number(row.alloySurchargePerKg || 0);
  return base + peso * asKg;
}

function formatCurrency(value) {
  return Number(value || 0).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

// cache temporanea per OD1 -> OD2
let currentComplexMap = {}; // { OD1: Set(OD2, ...) }

document.addEventListener("DOMContentLoaded", () => {
  const finishSelect = qs("finishSelect");
  const itemTypeSelect = qs("itemType");
  const qtyInput = qs("qtyInput");
  const qtyUnitSpan = qs("qtyUnitSpan");
  const qtyGroup = qs("qtyGroup");

  const ndGroup = qs("ndGroup");
  const ndSelect = qs("ndSelect");
  const od1Group = qs("od1Group");
  const od2Group = qs("od2Group");
  const od1Select = qs("od1Select");
  const od2Select = qs("od2Select");

  const alloyGroup = qs("alloyGroup");
  const alloyInput = qs("alloyInput");

  const otherItemsSection = qs("otherItemsSection");
  const otherItemsGrid = qs("otherItemsGrid");
  const addOtherItemsBtn = qs("addOtherItemsBtn");
  const otherItemInputs = new Map();

  const addRowBtn = qs("addBtn");
  const exportBtn = qs("exportBtn");
  const importBtn = qs("importBtn");
  const importFile = qs("importFile");
  const reloadCatalogBtn = qs("reloadCatalogBtn");

  const exportModalElement = qs("exportModal");
  const exportModal = new bootstrap.Modal(exportModalElement);

  const discountSuggestedRadio = qs("discountSuggestedRadio");
  const discountSuggestedValue = qs("discountSuggestedValue");
  const discountStilmasRadio = qs("discountStilmasRadio");
  const discountOtherRadio = qs("discountOtherRadio");
  const discountCustomInput = qs("discountCustomInput");
  const transportSelectModal = qs("transportSelectModal");
  const fileNameInput = qs("fileNameInput");
  const confirmExportBtn = qs("confirmExportBtn");

  function getDefaultFileName() {
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const dd = String(today.getDate()).padStart(2, "0");
    return `Offerta_${yyyy}-${mm}-${dd}`;
  }

  function computeOfferTotal(rows) {
    let tot = 0;
    for (const row of rows) {
      const unit = computeRowUnitPrice(row);
      const qty = Number(row.quantity || 0);
      tot += unit * qty;
    }
    return roundToDecimals(tot, 2);
  }

  function getSuggestedDiscount(totalValue) {
    if (totalValue < 20000) return 35.83;
    if (totalValue < 50000) return 41.18;
    if (totalValue < 100000) return 46.52;
    return 51.87;
  }

  function updateDiscountUI() {
    const totalValue = computeOfferTotal(currentRows);
    const suggested = getSuggestedDiscount(totalValue);
    discountSuggestedValue.textContent = `${suggested.toFixed(2)}%`;
    discountSuggestedRadio.value = suggested.toString();
    discountStilmasRadio.value = STILMAS_OSLA_DISCOUNT.toString();

    if (!fileNameInput.value) fileNameInput.value = getDefaultFileName();
    discountSuggestedRadio.checked = true;
    discountStilmasRadio.checked = false;
    discountOtherRadio.checked = false;
    discountCustomInput.value = "";
    discountCustomInput.disabled = true;
    transportSelectModal.value = "nord|4";
  }

  function getSelectedDiscount() {
    if (discountSuggestedRadio.checked)
      return parseFloat(discountSuggestedRadio.value) || 0;
    if (discountStilmasRadio.checked)
      return parseFloat(discountStilmasRadio.value) || STILMAS_OSLA_DISCOUNT;

    if (discountOtherRadio.checked) {
      const custom = parseFloat(discountCustomInput.value || "0");
      if (isNaN(custom) || custom < 0 || custom > 100) {
        alert("Inserisci uno sconto valido tra 0 e 100.");
        return null;
      }
      return custom;
    }
    return 0;
  }

  function handleDiscountOptionChange() {
    discountCustomInput.disabled = !discountOtherRadio.checked;
    if (!discountOtherRadio.checked) discountCustomInput.value = "";
  }

  function getSelectedTransportPercent() {
    const raw = transportSelectModal.value || "";
    const parts = raw.split("|");
    const percent = parts.length > 1 ? parts[1] : parts[0];
    const parsed = parseFloat(percent);
    return isNaN(parsed) ? 0 : parsed;
  }

  function renderOtherItemsGrid() {
    if (!otherItemsGrid) return;
    otherItemsGrid.innerHTML = "";
    otherItemInputs.clear();

    if (!otherItems.length) {
      const empty = document.createElement("div");
      empty.className = "col-12 text-center text-muted";
      empty.textContent = "Nessuna immagine trovata nella cartella 'immagini'.";
      otherItemsGrid.appendChild(empty);
      return;
    }

    otherItems.forEach((item) => {
      const col = document.createElement("div");
      col.className = "col-12";

      const card = document.createElement("div");
      card.className = "card h-100 other-item-card";

      const header = document.createElement("div");
      header.className = "card-header d-flex align-items-center gap-2";

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "form-check-input";
      checkbox.id = `${item.id}-checkbox`;

      const label = document.createElement("label");
      label.className = "form-check-label fw-semibold";
      label.htmlFor = checkbox.id;
      label.textContent = item.label;

      header.appendChild(checkbox);
      header.appendChild(label);
      card.appendChild(header);

      const body = document.createElement("div");
      body.className = "card-body d-flex flex-column align-items-start";

      const imgWrapper = document.createElement("div");
      imgWrapper.className = "other-item-img mb-2";

      if (item.url) {
        const imgEl = document.createElement("img");
        imgEl.src = item.url;
        imgEl.alt = item.label;
        imgWrapper.appendChild(imgEl);
      } else {
        imgWrapper.textContent = item.label;
      }
      body.appendChild(imgWrapper);

      const caption = document.createElement("div");
      caption.className = "other-item-caption mb-2";
      caption.textContent = item.fileName;
      body.appendChild(caption);

      card.appendChild(body);
      col.appendChild(card);
      otherItemsGrid.appendChild(col);

      otherItemInputs.set(item.id, {
        checkbox,
      });
    });
  }

  async function loadOtherItemsFromServer() {
    if (!otherItemsGrid) return;

    otherItemsGrid.innerHTML = "";
    const loading = document.createElement("div");
    loading.className = "col-12 text-center text-muted";
    loading.textContent = "Caricamento immagini...";
    otherItemsGrid.appendChild(loading);

    try {
      const res = await fetch("/api/other-items/images");
      if (!res.ok) throw new Error("Impossibile recuperare le immagini");
      const data = await res.json();
      otherItems = (data.images || []).map((img, idx) => ({
        id: `other-${idx}-${img.fileName}`,
        number: idx + 1,
        label: img.fileName,
        fileName: img.fileName,
        url: img.url,
      }));
    } catch (err) {
      console.error("Errore caricamento Other Items:", err);
      otherItems = [];
    }

    renderOtherItemsGrid();
  }

  function addSelectedOtherItemsToTable() {
    const selected = [];

    otherItems.forEach((item) => {
      const refs = otherItemInputs.get(item.id);
      if (!refs || !refs.checkbox.checked) return;

      selected.push({ item });
    });

    if (!selected.length) {
      alert("Seleziona almeno un Other Item.");
      return;
    }

    const finishValue = finishSelect.value || "N/A";

    selected.forEach(({ item }) => {
      currentRows.push({
        finish: finishValue,
        itemType: `Other Item - ${item.label}`,
        description: `${item.label}`,
        code: "",
        quantity: 1,
        basePricePerPc: 0,
        pesoKgM: null,
        alloySurchargePerKg: null,
        size: "",
      });

      // reset selections per ogni item aggiunto
      const refs = otherItemInputs.get(item.id);
      if (refs) {
        refs.checkbox.checked = false;
        refs.dnInput.value = "";
        refs.qtyField.value = "";
        refs.priceField.value = "";
      }
    });

    renderTable();
  }

  discountSuggestedRadio.addEventListener("change", handleDiscountOptionChange);
  discountStilmasRadio.addEventListener("change", handleDiscountOptionChange);
  discountOtherRadio.addEventListener("change", handleDiscountOptionChange);

  // Carica finiture dal backend (ASME BPE SF1 / SF4)
  fetch("/api/catalog/finishes")
    .then((r) => r.json())
    .then((data) => {
      (data.finishes || []).forEach((f) => {
        const opt = document.createElement("option");
        opt.value = f;
        opt.textContent = f;
        finishSelect.appendChild(opt);
      });
    })
    .catch((err) => console.error("Errore caricamento finiture:", err));

  function populateOd2Options(od1, previousOd2 = "") {
    od2Select.innerHTML = '<option value="">Seleziona</option>';
    od2Select.disabled = true;
    if (!od1 || !currentComplexMap[od1]) return;

    const od2Vals = Array.from(currentComplexMap[od1]).sort((a, b) => {
      const na = parseFloat(a.replace(",", "."));
      const nb = parseFloat(b.replace(",", "."));
      if (isNaN(na) || isNaN(nb)) return a.localeCompare(b);
      return na - nb;
    });
    for (const v of od2Vals) {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      od2Select.appendChild(opt);
    }
    if (previousOd2 && od2Vals.includes(previousOd2)) {
      od2Select.value = previousOd2;
      od2Select.disabled = false;
    } else if (od2Vals.length) {
      od2Select.disabled = false;
    }
  }

  // funzione per aggiornare ND / OD1-OD2 quando cambia finitura o item
  async function updateSizeOptions() {
    const finish = finishSelect.value;
    const itemType = itemTypeSelect.value;

    const previousND = ndSelect.value.trim();
    const previousOD1 = od1Select.value.trim();
    const previousOD2 = od2Select.value.trim();

    // reset selezioni
    ndSelect.innerHTML = '<option value="">Seleziona</option>';
    od1Select.innerHTML = '<option value="">Seleziona</option>';
    od2Select.innerHTML = '<option value="">Seleziona</option>';
    currentComplexMap = {};

    if (!finish || !itemType) {
      // nulla da caricare
      ndGroup.classList.remove("d-none");
      od1Group.classList.add("d-none");
      od2Group.classList.add("d-none");
      return;
    }

    try {
      if (itemType === "Tubes") {
        // Tubes -> ND da /api/catalog/tubes
        const res = await fetch(
          `/api/catalog/tubes?finish=${encodeURIComponent(finish)}`
        );
        const data = await res.json();
        const items = data.items || [];
        const ndSet = new Set(
          items.map((it) => (it.ND != null ? String(it.ND).trim() : ""))
        );
        const ndValues = Array.from(ndSet)
          .filter(Boolean)
          .sort(sortByNumericString);
        for (const nd of ndValues) {
          const opt = document.createElement("option");
          opt.value = nd;
          opt.textContent = nd;
          ndSelect.appendChild(opt);
        }
        if (previousND && ndValues.includes(previousND)) {
          ndSelect.value = previousND;
        }
        ndGroup.classList.remove("d-none");
        od1Group.classList.add("d-none");
        od2Group.classList.add("d-none");
      } else if (ND_ITEMS.has(itemType)) {
        // Fittings ND -> /api/catalog/simple
        const params = new URLSearchParams();
        params.set("type", itemType);
        params.set("finish", finish);
        const res = await fetch(`/api/catalog/simple?${params.toString()}`);
        const data = await res.json();
        const items = data.items || [];
        const ndSet = new Set(
          items.map((it) => (it.ND != null ? String(it.ND).trim() : ""))
        );
        const ndValues = Array.from(ndSet)
          .filter(Boolean)
          .sort(sortByNumericString);
        for (const nd of ndValues) {
          const opt = document.createElement("option");
          opt.value = nd;
          opt.textContent = nd;
          ndSelect.appendChild(opt);
        }
        if (previousND && ndValues.includes(previousND)) {
          ndSelect.value = previousND;
        }
        ndGroup.classList.remove("d-none");
        od1Group.classList.add("d-none");
        od2Group.classList.add("d-none");
      } else if (OD_ITEMS.has(itemType)) {
        // Tees / Reducers -> /api/catalog/complex
        const params = new URLSearchParams();
        params.set("type", itemType);
        params.set("finish", finish);
        const res = await fetch(`/api/catalog/complex?${params.toString()}`);
        const data = await res.json();
        const items = data.items || [];

        currentComplexMap = {};
        for (const it of items) {
          const OD1 = it.OD1 != null ? String(it.OD1).trim() : "";
          const OD2 = it.OD2 != null ? String(it.OD2).trim() : "";
          if (!OD1 || !OD2) continue;
          if (!currentComplexMap[OD1]) currentComplexMap[OD1] = new Set();
          currentComplexMap[OD1].add(OD2);
        }

        const od1Values = Object.keys(currentComplexMap).sort((a, b) => {
          const na = parseFloat(a.replace(",", "."));
          const nb = parseFloat(b.replace(",", "."));
          if (isNaN(na) || isNaN(nb)) return a.localeCompare(b);
          return na - nb;
        });

        for (const v of od1Values) {
          const opt = document.createElement("option");
          opt.value = v;
          opt.textContent = v;
          od1Select.appendChild(opt);
        }

        if (previousOD1 && od1Values.includes(previousOD1)) {
          od1Select.value = previousOD1;
          populateOd2Options(previousOD1, previousOD2);
        }

        ndGroup.classList.add("d-none");
        od1Group.classList.remove("d-none");
        od2Group.classList.remove("d-none");
        if (!previousOD1 || !od1Values.includes(previousOD1)) {
          od2Select.disabled = true;
        }
      } else {
        // fallback
        ndGroup.classList.remove("d-none");
        od1Group.classList.add("d-none");
        od2Group.classList.add("d-none");
      }
    } catch (err) {
      console.error("Errore caricamento ND/OD1-OD2:", err);
      alert("Errore nel caricamento dei diametri dal catalogo.");
    }
  }

  // Popola OD2 quando scelgo OD1
  od1Select.addEventListener("change", () => {
    populateOd2Options(od1Select.value);
  });

  // Cambia UI in base all'item scelto (Q.ty unità, Alloy abilitato solo Tubes, ND/OD)
  itemTypeSelect.addEventListener("change", () => {
    const itemType = itemTypeSelect.value;
    const isOtherItem = itemType === "Other Items";

    if (isOtherItem) {
      otherItemsSection.classList.remove("d-none");
      qtyGroup.classList.add("d-none");
      ndGroup.classList.add("d-none");
      od1Group.classList.add("d-none");
      od2Group.classList.add("d-none");
      alloyGroup.classList.add("d-none");
    } else {
      otherItemsSection.classList.add("d-none");
      qtyGroup.classList.remove("d-none");
      alloyGroup.classList.remove("d-none");
    }
    // Q.ty unità
    if (itemType === "Tubes") {
      qtyUnitSpan.textContent = "(m)";
    } else {
      qtyUnitSpan.textContent = "(pcs)";
    }

    // Alloy solo Tubes
    if (itemType === "Tubes") {
      alloyInput.disabled = false;
    } else {
      alloyInput.disabled = true;
      alloyInput.value = "";
    }

    // aggiorna ND / OD1-OD2
    if (!isOtherItem) updateSizeOptions();
  });

  // Se cambia finitura, ricarico ND / OD1-OD2
  finishSelect.addEventListener("change", () => {
    updateSizeOptions();
  });

  // Aggiungi riga
  addRowBtn.addEventListener("click", async () => {
    const finish = finishSelect.value;
    const itemType = itemTypeSelect.value;
    const alloy = parseFloat(alloyInput.value ?? "0") || 0;
    if (itemType === "Other Items") {
      addSelectedOtherItemsToTable();
      return;
    }

    const prevND = ndSelect.value;
    const prevOD1 = od1Select.value;
    const prevOD2 = od2Select.value;

    let qtyRaw = qtyInput.value ?? "0";
    let qty = parseInt(qtyRaw, 10);
    if (isNaN(qty)) qty = 0;

    if (!finish) {
      alert("Seleziona la finitura (ASME BPE SF1 o SF4).");
      return;
    }
    if (!itemType) {
      alert("Seleziona l'Item.");
      return;
    }
    if (!qty || qty <= 0) {
      alert("Inserisci una Q.ty valida (intera).");
      return;
    }

    let endpoint = "";
    const params = new URLSearchParams();

    let sizeText = "";
    let nd = "";
    let od1 = "";
    let od2 = "";

    if (itemType === "Tubes") {
      nd = ndSelect.value.trim();
      if (!nd) {
        alert("Seleziona il ND per Tubes.");
        return;
      }
      endpoint = "/api/catalog/tubes";
      params.set("finish", finish);
      params.set("ND", nd);
      sizeText = nd;
    } else if (ND_ITEMS.has(itemType)) {
      nd = ndSelect.value.trim();
      if (!nd) {
        alert("Seleziona il ND.");
        return;
      }
      endpoint = "/api/catalog/simple";
      params.set("type", itemType);
      params.set("finish", finish);
      params.set("ND", nd);
      sizeText = nd;
    } else if (OD_ITEMS.has(itemType)) {
      od1 = od1Select.value.trim();
      od2 = od2Select.value.trim();
      if (!od1 || !od2) {
        alert("Seleziona OD1 e OD2.");
        return;
      }
      endpoint = "/api/catalog/complex";
      params.set("type", itemType);
      params.set("finish", finish);
      params.set("OD1", od1);
      params.set("OD2", od2);
      sizeText = `${od1} x ${od2}`;
    } else {
      alert("Item non gestito.");
      return;
    }

    try {
      const res = await fetch(`${endpoint}?${params.toString()}`);
      const raw = await res.text();

      let data;
      try {
        data = raw ? JSON.parse(raw) : {};
      } catch (parseErr) {
        console.error("Risposta catalogo non valida:", raw, parseErr);
        throw new Error(
          "Catalogo non disponibile: risposta non valida dal server."
        );
      }

      if (!res.ok) {
        const serverMsg =
          (typeof data?.error === "string" && data.error.trim()) ||
          (raw && raw.trim()) ||
          `Catalogo non disponibile (codice ${res.status}).`;
        throw new Error(serverMsg);
      }

      const items = data.items || [];

      if (!items.length) {
        alert(
          "Nessun articolo trovato nel catalogo per i parametri selezionati."
        );
        return;
      }

      const catItem = items[0]; // se più di uno, prendo il primo

      let description = `${itemType} ${finish} ${sizeText}`;
      if (itemType === "Clamps") {
        const flangeSize =
          catItem.flangeSizeMm != null
            ? String(catItem.flangeSizeMm).trim()
            : "";
        if (flangeSize) {
          description += ` - Flange size mm: ${flangeSize}`;
        }
      }
      if (
        itemType === "Ferrule A (Long)" ||
        itemType === "Ferrule B (Medium)" ||
        itemType === "Ferrule C (Short)"
      ) {
        const lengthVal =
          catItem.lengthMm != null ? String(catItem.lengthMm).trim() : "";
        if (lengthVal) {
          description += ` - L= ${lengthVal} mm`;
        }
      }

      const row = {
        finish,
        itemType,
        description,
        code: catItem.code || "",
        quantity: qty,
      };

      if (itemType === "Tubes") {
        row.basePricePerM = catItem.pricePerM || 0;
        row.pesoKgM = catItem.pesoKgM || 0;
        row.alloySurchargePerKg = alloy || 0;
      } else {
        row.basePricePerPc = catItem.pricePerPc || 0;
        row.pesoKgM = null;
        row.alloySurchargePerKg = null;
      }

      row.size = sizeText;

      currentRows.push(row);
      renderTable();

      // Mantiene le scelte precedenti per facilitare l'inserimento della riga successiva
      ndSelect.value = prevND;
      od1Select.value = prevOD1;
      od2Select.value = prevOD2;
    } catch (err) {
      console.error("Errore durante la lettura del catalogo:", err);
      alert(err?.message || "Errore durante la lettura del catalogo.");
    }
  });

  // EXPORT: apre la modale (Sconto / Trasporto)
  exportBtn.addEventListener("click", () => {
    if (!currentRows.length) {
      alert("Non ci sono righe da esportare.");
      return;
    }
    updateDiscountUI();
    exportModal.show();
  });

  // Conferma export dalla modale
  confirmExportBtn.addEventListener("click", async () => {
    if (!currentRows.length) {
      alert("Non ci sono righe da esportare.");
      return;
    }

    const discountPercent = getSelectedDiscount();
    if (discountPercent === null) return;

    const transportPercent = getSelectedTransportPercent();

    const fileNameSafe =
      (fileNameInput.value || getDefaultFileName()).trim() ||
      getDefaultFileName();
    try {
      const res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          rows: currentRows,
          discountPercent,
          transport: { percent: transportPercent },
        }),
      });

      if (!res.ok) throw new Error("Errore export");

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${fileNameSafe}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      exportModal.hide();
    } catch (err) {
      console.error(err);
      alert("Errore durante l'esportazione.");
    }
  });

  // Import Excel
  importBtn.addEventListener("click", () => {
    importFile.click();
  });

  importFile.addEventListener("change", async () => {
    const file = importFile.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch("/api/import", {
        method: "POST",
        body: formData,
      });
      if (!res.ok) throw new Error("Errore import");
      const data = await res.json();
      const rows = data.rows || [];

      currentRows = rows.map((r) => {
        const isTube = r.itemType === "Tubes";
        const row = {
          finish: "", // non ricostruibile dall'Excel
          itemType: isTube ? "Tubes" : "Imported",
          description: r.description || "",
          code: r.code || "",
          quantity: r.quantity || 0,
        };
        if (isTube) {
          row.basePricePerM = r.base || r.unitPrice || 0;
          row.pesoKgM = r.peso || 0;
          row.alloySurchargePerKg = r.asKg || 0;
        } else {
          row.basePricePerPc = r.base || r.unitPrice || 0;
          row.pesoKgM = null;
          row.alloySurchargePerKg = null;
        }
        row.size = ""; // ND/OD1-OD2 non recuperabili dall'Excel
        return row;
      });
      renderTable();
      importFile.value = "";
    } catch (err) {
      console.error(err);
      alert("Errore durante l'importazione.");
    }
  });
  addOtherItemsBtn.addEventListener("click", addSelectedOtherItemsToTable);
  loadOtherItemsFromServer();
});

// Render tabella righe
function renderTable() {
  const tbody = document.querySelector("#offerTable tbody");
  tbody.innerHTML = "";

  let grandTotal = 0;

  currentRows.forEach((row, idx) => {
    const tr = document.createElement("tr");

    const colIndex = document.createElement("td");
    colIndex.textContent = idx + 1;
    tr.appendChild(colIndex);

    const colItem = document.createElement("td");
    colItem.textContent = row.itemType || "";
    tr.appendChild(colItem);

    const colDesc = document.createElement("td");
    colDesc.textContent = row.description || "";
    tr.appendChild(colDesc);

    const colCode = document.createElement("td");
    colCode.textContent = row.code || "";
    tr.appendChild(colCode);

    const colBase = document.createElement("td");
    const isTube = row.itemType === "Tubes";
    colBase.classList.add("text-end");
    if (isTube) {
      const base = row.basePricePerM ?? 0;
      colBase.textContent = base.toFixed(2);
    } else {
      const input = document.createElement("input");
      input.type = "number";
      input.step = "0.01";
      input.min = "0";
      input.className = "form-control form-control-sm text-end";
      input.value = (row.basePricePerPc ?? 0).toFixed(2);
      input.addEventListener("change", () => {
        const newVal = parseFloat(input.value.replace(",", "."));
        row.basePricePerPc = isNaN(newVal) || newVal < 0 ? 0 : newVal;
        renderTable();
      });
      colBase.appendChild(input);
    }
    tr.appendChild(colBase);

    const colAlloy = document.createElement("td");
    if (row.itemType === "Tubes") {
      colAlloy.classList.add("col-alloy");
      const input = document.createElement("input");
      input.type = "number";
      input.step = "0.01";
      input.min = "0";
      input.className = "form-control form-control-sm text-end";
      input.value = (row.alloySurchargePerKg ?? 0).toFixed(2);
      input.addEventListener("change", () => {
        const newVal = parseFloat(input.value.replace(",", "."));
        row.alloySurchargePerKg = isNaN(newVal) || newVal < 0 ? 0 : newVal;
        renderTable();
      });
      colAlloy.appendChild(input);
    } else {
      colAlloy.textContent = "-";
      colAlloy.classList.add("text-center", "col-alloy");
    }
    tr.appendChild(colAlloy);

    const colUnit = document.createElement("td");
    colUnit.classList.add("text-end", "col-unit");
    const unitPrice = computeRowUnitPrice(row);
    colUnit.textContent = unitPrice.toFixed(2);
    tr.appendChild(colUnit);

    const colQty = document.createElement("td");
    colQty.classList.add("col-qty");
    const qtyValue = Number(row.quantity ?? 0);
    const qtyInput = document.createElement("input");
    qtyInput.type = "number";
    qtyInput.min = "0";
    qtyInput.step = "1";
    qtyInput.className = "form-control form-control-sm text-end";
    qtyInput.value = qtyValue.toString();
    qtyInput.addEventListener("change", () => {
      const newQty = parseInt(qtyInput.value, 10);
      row.quantity = isNaN(newQty) || newQty < 0 ? 0 : newQty;
      renderTable();
    });
    colQty.appendChild(qtyInput);
    tr.appendChild(colQty);

    const colTotal = document.createElement("td");
    colTotal.classList.add("text-end");
    const rowTotal = unitPrice * qtyValue;
    colTotal.textContent = formatCurrency(rowTotal);
    grandTotal += rowTotal;
    tr.appendChild(colTotal);

    const colActions = document.createElement("td");
    colActions.classList.add("text-center");
    const btnRemove = document.createElement("button");
    btnRemove.type = "button";
    btnRemove.className = "btn btn-sm btn-danger btn-icon";
    btnRemove.textContent = "X";
    btnRemove.addEventListener("click", () => {
      currentRows.splice(idx, 1);
      renderTable();
    });
    colActions.appendChild(btnRemove);
    tr.appendChild(colActions);

    tbody.appendChild(tr);
  });
  if (currentRows.length) {
    const totalRow = document.createElement("tr");
    totalRow.classList.add("table-secondary", "fw-semibold");

    const labelCell = document.createElement("td");
    labelCell.colSpan = 8;
    labelCell.classList.add("text-end");
    labelCell.textContent = "Grand Total";
    totalRow.appendChild(labelCell);

    const valueCell = document.createElement("td");
    valueCell.classList.add("text-end");
    valueCell.textContent = formatCurrency(grandTotal);
    totalRow.appendChild(valueCell);

    const emptyCell = document.createElement("td");
    totalRow.appendChild(emptyCell);

    tbody.appendChild(totalRow);
  }
}
