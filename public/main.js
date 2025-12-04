// public/main.js

// Stato corrente righe
let currentRows = [];

// Item che usano solo ND
const ND_ITEMS = new Set([
  "Tubes",
  "Elbows 90°",
  "Elbows 45°",
  "End Caps",
  "Ferrule A (Long)",
  "Ferrule B (Medium)",
  "Ferrule C (Short)",
]);

// Item che usano OD1 / OD2
const OD_ITEMS = new Set(["Tees", "Conc. Reducers", "Ecc. Reducers"]);

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

// cache temporanea per OD1 -> OD2
let currentComplexMap = {}; // { OD1: Set(OD2, ...) }

document.addEventListener("DOMContentLoaded", () => {
  const finishSelect = qs("finishSelect");
  const itemTypeSelect = qs("itemType");
  const qtyInput = qs("qtyInput");
  const qtyUnitSpan = qs("qtyUnitSpan");

  const ndGroup = qs("ndGroup");
  const ndSelect = qs("ndSelect");
  const od1Group = qs("od1Group");
  const od2Group = qs("od2Group");
  const od1Select = qs("od1Select");
  const od2Select = qs("od2Select");

  const alloyGroup = qs("alloyGroup");
  const alloyInput = qs("alloyInput");

  const addRowBtn = qs("addBtn");
  const exportBtn = qs("exportBtn");
  const importBtn = qs("importBtn");
  const importFile = qs("importFile");
  const reloadCatalogBtn = qs("reloadCatalogBtn");

  const exportModalElement = qs("exportModal");
  const exportModal = new bootstrap.Modal(exportModalElement);

  const richiestaInputModal = qs("richiestaInputModal");
  const clienteInputModal = qs("clienteInputModal");
  const cantiereInputModal = qs("cantiereInputModal");
  const discountSuggestedRadio = qs("discountSuggestedRadio");
  const discountSuggestedValue = qs("discountSuggestedValue");
  const discountNoneRadio = qs("discountNoneRadio");
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

    if (!fileNameInput.value) fileNameInput.value = getDefaultFileName();
    discountSuggestedRadio.checked = true;
    discountNoneRadio.checked = false;
    discountOtherRadio.checked = false;
    discountCustomInput.value = "";
    discountCustomInput.disabled = true;
    transportSelectModal.value = "4";
  }

  function getSelectedDiscount() {
    if (discountSuggestedRadio.checked)
      return parseFloat(discountSuggestedRadio.value) || 0;
    if (discountNoneRadio.checked) return 0;
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

  discountSuggestedRadio.addEventListener("change", handleDiscountOptionChange);
  discountNoneRadio.addEventListener("change", handleDiscountOptionChange);
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

  // funzione per aggiornare ND / OD1-OD2 quando cambia finitura o item
  async function updateSizeOptions() {
    const finish = finishSelect.value;
    const itemType = itemTypeSelect.value;

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
        for (const nd of Array.from(ndSet)
          .filter(Boolean)
          .sort(sortByNumericString)) {
          const opt = document.createElement("option");
          opt.value = nd;
          opt.textContent = nd;
          ndSelect.appendChild(opt);
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
        for (const nd of Array.from(ndSet)
          .filter(Boolean)
          .sort(sortByNumericString)) {
          const opt = document.createElement("option");
          opt.value = nd;
          opt.textContent = nd;
          ndSelect.appendChild(opt);
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

        ndGroup.classList.add("d-none");
        od1Group.classList.remove("d-none");
        od2Group.classList.remove("d-none");
        od2Select.disabled = true;
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
    const od1 = od1Select.value;
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
      od2Select.disabled = false;
    }
  });

  // Cambia UI in base all'item scelto (Q.ty unità, Alloy abilitato solo Tubes, ND/OD)
  itemTypeSelect.addEventListener("change", () => {
    const itemType = itemTypeSelect.value;

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
    updateSizeOptions();
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

      const description = `${itemType} ${finish} ${sizeText}`;

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

      // reset Q.ty ma non gli altri
      qtyInput.value = "";
    } catch (err) {
      console.error("Errore durante la lettura del catalogo:", err);
      alert(err?.message || "Errore durante la lettura del catalogo.");
    }
  });

  // EXPORT: apre la modale (Richiesta / Cantiere / Sconto / Trasporto)
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

    const transportPercent = parseFloat(transportSelectModal.value ?? "0") || 0;

    const meta = {
      richiesta: richiestaInputModal.value || "",
      cliente: clienteInputModal.value || "",
      cantiere: cantiereInputModal.value || "",
    };
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
          meta,
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
});

// Render tabella righe
function renderTable() {
  const tbody = document.querySelector("#offerTable tbody");
  tbody.innerHTML = "";

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
    colBase.classList.add("text-end");
    const base =
      row.itemType === "Tubes"
        ? row.basePricePerM ?? 0
        : row.basePricePerPc ?? 0;
    colBase.textContent = base.toFixed(2);
    tr.appendChild(colBase);

    const colAlloy = document.createElement("td");
    if (row.itemType === "Tubes") {
      colAlloy.classList.add("text-end");
      colAlloy.textContent = (row.alloySurchargePerKg ?? 0).toFixed(2);
    } else {
      colAlloy.textContent = "-";
      colAlloy.classList.add("text-center");
    }
    tr.appendChild(colAlloy);

    const colUnit = document.createElement("td");
    colUnit.classList.add("text-end");
    const unitPrice = computeRowUnitPrice(row);
    colUnit.textContent = unitPrice.toFixed(2);
    tr.appendChild(colUnit);

    const colQty = document.createElement("td");
    colQty.classList.add("text-end");
    const qtyValue = Number(row.quantity ?? 0);
    colQty.textContent = qtyValue.toString();
    tr.appendChild(colQty);

    const colTotal = document.createElement("td");
    colTotal.classList.add("text-end");
    colTotal.textContent = (unitPrice * qtyValue).toFixed(2);
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
}
