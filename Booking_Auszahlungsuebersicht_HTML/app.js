const DEFAULT_MAPPING = [
  { Keyword: "PrimeTimeSuites | Nahe Messe&Hbf", Beherbergungsidentifikationsnummer: "00631", Strasse: "Opernplatz 15" },
  { Keyword: "PrimeTimeSuite:Kuche+WiFi+Modern", Beherbergungsidentifikationsnummer: "00632", Strasse: "Mulheimer Str. 15" },
  { Keyword: "NewYork Suite", Beherbergungsidentifikationsnummer: "00633", Strasse: "Haskenstr. 44" },
  { Keyword: "Paris Suite", Beherbergungsidentifikationsnummer: "00633", Strasse: "Haskenstr. 44" },
  { Keyword: "London Suite", Beherbergungsidentifikationsnummer: "00633", Strasse: "Haskenstr. 44" },
  { Keyword: "Art Gallery", Beherbergungsidentifikationsnummer: "00634", Strasse: "Mulheimer Str. 17" }
];

const DEFAULT_RENAMES = {
  "Beherbergungsidentifikationsnummer": "Beherbergungsidentifikationsnummer",
  "Strasse": "Strasse",
  "Bruttobetrag": "SP1 Bruttoeinkuenfte",
  "Reinigungsgebuehr": "SP2 Reinigungskosten",
  "Zwischensumme (BmGI Beherbergungssteuer)": "SP4 Zwischensumme (BmGI Beherbergungssteuer)",
  "Kommissionen": "SP5 Servicegebuehren",
  "Auszahlungsbetrag": "SP6 Auszahlungsbetrag",
  "Bruttobetrag Inklusive Beherbergungssteuer": "Bruttobetrag inkl. Beherbergungssteuer",
  "Beherbergungssteuer": "SP7 Beherbergungssteuer (5%)"
};

const DEFAULT_GROUP_LABELS = {
  "Beherbergungsidentifikationsnummer": "",
  "Strasse": "",
  "Bruttobetrag": "SP1",
  "Reinigungsgebuehr": "SP2",
  "Zwischensumme (BmGI Beherbergungssteuer)": "SP4",
  "Kommissionen": "SP5",
  "Auszahlungsbetrag": "SP6",
  "Bruttobetrag Inklusive Beherbergungssteuer": "",
  "Beherbergungssteuer": "SP7"
};

const DEFAULT_COLUMN_ORDER = [
  "Beherbergungsidentifikationsnummer",
  "Strasse",
  "Bruttobetrag",
  "Reinigungsgebuehr",
  "Zwischensumme (BmGI Beherbergungssteuer)",
  "Kommissionen",
  "Auszahlungsbetrag",
  "Bruttobetrag Inklusive Beherbergungssteuer",
  "Beherbergungssteuer"
];

const DEFAULT_HEADER = {
  titleLine1: "Booking Auszahlungsuebersicht",
  titleLine2: "",
  showGroupRow: false
};

const state = {
  bookingFile: null,
  payoutFile: null,
  mapping: [],
  renames: {},
  groupLabels: {},
  columnOrder: [],
  header: { ...DEFAULT_HEADER },
  output: null
};

const els = {
  bookingFile: document.getElementById("bookingFile"),
  payoutFile: document.getElementById("payoutFile"),
  runBtn: document.getElementById("runBtn"),
  downloadBtn: document.getElementById("downloadBtn"),
  status: document.getElementById("status"),
  mappingTableBody: document.querySelector("#mappingTable tbody"),
  addMappingRowBtn: document.getElementById("addMappingRowBtn"),
  saveMappingBtn: document.getElementById("saveMappingBtn"),
  resetMappingBtn: document.getElementById("resetMappingBtn"),
  renamesTableBody: document.querySelector("#renamesTable tbody"),
  saveRenamesBtn: document.getElementById("saveRenamesBtn"),
  resetRenamesBtn: document.getElementById("resetRenamesBtn"),
  resultTableHead: document.querySelector("#resultTable thead"),
  resultTableBody: document.querySelector("#resultTable tbody"),
  bookingName: document.getElementById("bookingName"),
  payoutName: document.getElementById("payoutName"),
  bookingProgress: document.getElementById("bookingProgress"),
  payoutProgress: document.getElementById("payoutProgress"),
  titleLine1: document.getElementById("titleLine1"),
  titleLine2: document.getElementById("titleLine2"),
  showGroupRow: document.getElementById("showGroupRow"),
  overlay: document.getElementById("mappingOverlay"),
  openOverlayBtn: document.getElementById("openOverlayBtn"),
  closeOverlayBtn: document.getElementById("closeOverlayBtn"),
  closeOverlayBtnFooter: document.getElementById("closeOverlayBtnFooter")
};

function loadState() {
  const mappingRaw = localStorage.getItem("booking_mapping");
  const renamesRaw = localStorage.getItem("booking_renames");
  const groupsRaw = localStorage.getItem("booking_groups");
  const orderRaw = localStorage.getItem("booking_column_order");
  const headerRaw = localStorage.getItem("booking_header");
  state.mapping = mappingRaw ? JSON.parse(mappingRaw) : DEFAULT_MAPPING.slice();
  const renamesParsed = renamesRaw ? JSON.parse(renamesRaw) : {};
  const groupsParsed = groupsRaw ? JSON.parse(groupsRaw) : {};
  state.renames = { ...DEFAULT_RENAMES, ...renamesParsed };
  state.groupLabels = { ...DEFAULT_GROUP_LABELS, ...groupsParsed };
  state.columnOrder = orderRaw ? JSON.parse(orderRaw) : DEFAULT_COLUMN_ORDER.slice();
  state.header = headerRaw ? JSON.parse(headerRaw) : { ...DEFAULT_HEADER };
}

function saveMapping() {
  localStorage.setItem("booking_mapping", JSON.stringify(state.mapping));
}

function saveRenames() {
  localStorage.setItem("booking_renames", JSON.stringify(state.renames));
}

function saveGroupLabels() {
  localStorage.setItem("booking_groups", JSON.stringify(state.groupLabels));
}

function saveColumnOrder() {
  localStorage.setItem("booking_column_order", JSON.stringify(state.columnOrder));
}

function saveHeader() {
  localStorage.setItem("booking_header", JSON.stringify(state.header));
}

function renderHeaderInputs() {
  els.titleLine1.value = state.header.titleLine1 || "";
  els.titleLine2.value = state.header.titleLine2 || "";
  els.showGroupRow.checked = state.header.showGroupRow === true;
}

function openOverlay() {
  els.overlay.classList.add("show");
  els.overlay.setAttribute("aria-hidden", "false");
}

function closeOverlay() {
  els.overlay.classList.remove("show");
  els.overlay.setAttribute("aria-hidden", "true");
}

function renderMappingTable() {
  els.mappingTableBody.innerHTML = "";
  state.mapping.forEach((row, idx) => {
    const tr = document.createElement("tr");
    ["Keyword", "Beherbergungsidentifikationsnummer", "Strasse"].forEach((key) => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.value = row[key] || "";
      input.addEventListener("input", () => {
        state.mapping[idx][key] = input.value;
      });
      td.appendChild(input);
      tr.appendChild(td);
    });
    els.mappingTableBody.appendChild(tr);
  });
}

function renderRenamesTable() {
  els.renamesTableBody.innerHTML = "";
  const ordered = getOrderedKeys();
  ordered.forEach((orig) => {
    const tr = document.createElement("tr");
    tr.setAttribute("draggable", "true");
    tr.dataset.key = orig;
    const tdOrig = document.createElement("td");
    tdOrig.textContent = orig;
    const tdNew = document.createElement("td");
    const input = document.createElement("input");
    input.value = state.renames[orig] || orig;
    input.addEventListener("input", () => {
      state.renames[orig] = input.value.trim() || orig;
    });
    tdNew.appendChild(input);
    const tdGroup = document.createElement("td");
    const groupInput = document.createElement("input");
    groupInput.value = state.groupLabels[orig] || "";
    groupInput.addEventListener("input", () => {
      state.groupLabels[orig] = groupInput.value.trim();
    });
    tdGroup.appendChild(groupInput);
    const tdDrag = document.createElement("td");
    const dragHandle = document.createElement("span");
    dragHandle.className = "drag-handle";
    dragHandle.textContent = "::";
    tdDrag.appendChild(dragHandle);
    tr.appendChild(tdDrag);
    tr.appendChild(tdOrig);
    tr.appendChild(tdNew);
    tr.appendChild(tdGroup);
    tr.addEventListener("dragstart", () => {
      tr.classList.add("dragging");
      state.dragKey = orig;
    });
    tr.addEventListener("dragend", () => {
      tr.classList.remove("dragging");
      state.dragKey = null;
    });
    tr.addEventListener("dragover", (event) => {
      event.preventDefault();
    });
    tr.addEventListener("drop", (event) => {
      event.preventDefault();
      const fromKey = state.dragKey;
      const toKey = orig;
      if (!fromKey || fromKey === toKey) return;
      reorderColumns(fromKey, toKey);
    });
    els.renamesTableBody.appendChild(tr);
  });
}

function setStatus(msg, type = "") {
  els.status.textContent = msg;
  els.status.style.color = type === "error" ? "#f87171" : "#9aa4b2";
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9\s]/g, " ");
}

function parseMoneySmart(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number" && !Number.isNaN(value)) return value;
  const str = String(value).trim();
  if (!str) return 0;
  const cleaned = str.replace(/[^0-9,\.\-]/g, "");
  if (!cleaned || cleaned === "-" || cleaned === "." || cleaned === ",") return 0;
  const lastDot = cleaned.lastIndexOf(".");
  const lastComma = cleaned.lastIndexOf(",");
  let dec = ".";
  if (lastDot === -1 && lastComma !== -1) dec = ",";
  else if (lastDot !== -1 && lastComma !== -1) dec = lastDot > lastComma ? "." : ",";
  const thou = dec === "." ? "," : ".";
  let out = cleaned.split(thou).join("");
  if (dec === ",") out = out.replace(/,/g, ".");
  const num = Number(out);
  return Number.isFinite(num) ? num : 0;
}

function formatCurrency(value) {
  const num = typeof value === "number" ? value : parseMoneySmart(value);
  return new Intl.NumberFormat("de-DE", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(num);
}

function extractBookingNumber(notes) {
  const marker = "Buchungsnummer:";
  const idx = notes.indexOf(marker);
  if (idx === -1) return null;
  const after = notes.slice(idx + marker.length).trim();
  const match = after.match(/^(\d+)/);
  return match ? match[1] : null;
}

function extractSumFromColumn(rows, key, needles) {
  const list = Array.isArray(needles) ? needles : [needles];
  let total = 0;
  rows.forEach((row) => {
    const val = normalizeText(row[key] || "").toLowerCase();
    const hit = list.some((needle) => val.includes(needle));
    if (hit) {
      total += Math.abs(parseMoneySmart(row[key]));
    }
  });
  return total;
}

function buildBlocksFromBookingList(rows) {
  const blocks = [];
  let currentBlock = [];
  rows.forEach((row) => {
    const hasStart = row.Anreise && row.Abreise && row.Unterkunft && row.Position;
    if (hasStart && currentBlock.length) {
      blocks.push(currentBlock);
      currentBlock = [];
    }
    currentBlock.push(row);
  });
  if (currentBlock.length) blocks.push(currentBlock);

  return blocks
    .map((group) => {
      const notes = group.map((r) => r.Notizen || "").join(" ");
      const bookingNo = extractBookingNumber(notes);
      const cleaning = extractSumFromColumn(group, "Abreise", ["reinigungsgeb", "cleaning"]);
      const overnight = extractSumFromColumn(group, "Position", [
        "uebernachtungssteuer",
        "ubernachtungssteuer",
        "nachtungssteuer",
        "overnight"
      ]);
      return {
        Buchungsnummer_clean: bookingNo,
        Reinigungsgebuehr_num: cleaning,
        Uebernachtungssteuer_num: overnight
      };
    })
    .filter((row) => row.Buchungsnummer_clean);
}

function buildPayoutMini(rows) {
  const filtered = rows.filter((row) => {
    return !row["Art/Transaktionsart"] || String(row["Art/Transaktionsart"]).trim() === "Buchung";
  });
  const feeCol = findFeeColumn(filtered);

  return filtered.map((row) => {
    const gross = parseMoneySmart(row.Bruttobetrag);
    const commission = Math.abs(parseMoneySmart(row.Kommission));
    const fee = feeCol ? Math.abs(parseMoneySmart(row[feeCol])) : 0;
    return {
      Referenznummer: String(row.Referenznummer || "").trim(),
      Unterkunftsname: String(row.Unterkunftsname || ""),
      Bruttobetrag_brutto: gross,
      Kommissionen: commission + fee
    };
  });
}

function findFeeColumn(rows) {
  if (!rows.length) return null;
  const headers = Object.keys(rows[0]);
  const normalized = headers.map((header) => ({
    header,
    normalized: normalizeText(header).toLowerCase()
  }));
  const byPaymentFee = normalized.find((item) => item.normalized.includes("payment fee"));
  if (byPaymentFee) return byPaymentFee.header;
  const byBookingFee = normalized.find(
    (item) => item.normalized.includes("zahlung") && item.normalized.includes("geb")
  );
  return byBookingFee ? byBookingFee.header : null;
}

function mapIdAndStreet(name) {
  const lowered = String(name || "").toLowerCase();
  for (const row of state.mapping) {
    const kw = String(row.Keyword || "").toLowerCase();
    if (kw && lowered.includes(kw)) {
      return {
        id: String(row.Beherbergungsidentifikationsnummer || "").trim(),
        street: String(row.Strasse || "").trim()
      };
    }
  }
  return { id: "ID-BITTE_ANPASSEN", street: "" };
}

function getOrderedKeys() {
  const known = new Set(DEFAULT_COLUMN_ORDER);
  const order = (state.columnOrder || []).filter((key) => known.has(key));
  const merged = order.slice();
  DEFAULT_COLUMN_ORDER.forEach((key) => {
    if (!merged.includes(key)) merged.push(key);
  });
  state.columnOrder = merged;
  return merged;
}

function reorderColumns(fromKey, toKey) {
  const order = getOrderedKeys();
  const fromIndex = order.indexOf(fromKey);
  const toIndex = order.indexOf(toKey);
  if (fromIndex === -1 || toIndex === -1) return;
  order.splice(fromIndex, 1);
  order.splice(toIndex, 0, fromKey);
  state.columnOrder = order;
  saveColumnOrder();
  renderRenamesTable();
  if (state.output) {
    state.output.columns = buildColumns();
    renderResultTable(state.output);
  }
}

function buildColumns() {
  const keys = getOrderedKeys();
  return keys.map((key) => ({
    key,
    label: state.renames[key] || key,
    type: key === "Beherbergungsidentifikationsnummer" || key === "Strasse" ? "string" : "number",
    groupLabel: state.groupLabels[key] || ""
  }));
}

function computeResult(bookingRows, payoutRows) {
  const blocks = buildBlocksFromBookingList(bookingRows);
  if (!blocks.length) throw new Error("Keine Buchungsbloecke gefunden.");

  const payoutMini = buildPayoutMini(payoutRows);
  const payoutMap = new Map(payoutMini.map((r) => [r.Referenznummer, r]));
  const merged = blocks
    .map((block) => {
      const payout = payoutMap.get(block.Buchungsnummer_clean);
      if (!payout) return null;
      const mergedRow = {
        ...block,
        ...payout
      };
      mergedRow.Bruttobetrag = mergedRow.Bruttobetrag_brutto - mergedRow.Uebernachtungssteuer_num;
      mergedRow["Zwischensumme (BmGI Beherbergungssteuer)"] = mergedRow.Bruttobetrag - mergedRow.Reinigungsgebuehr_num;
      mergedRow.Auszahlungsbetrag =
        mergedRow.Bruttobetrag - mergedRow.Kommissionen + mergedRow.Uebernachtungssteuer_num;
      mergedRow.Beherbergungssteuer = mergedRow["Zwischensumme (BmGI Beherbergungssteuer)"] * 0.05;
      const mapped = mapIdAndStreet(mergedRow.Unterkunftsname);
      mergedRow.Beherbergungsidentifikationsnummer = mapped.id;
      mergedRow.Strasse = mapped.street;
      return mergedRow;
    })
    .filter(Boolean);

  if (!merged.length) throw new Error("Join ergab 0 Zeilen.");

  const grouped = groupByIdAndStreet(merged).map((row) => {
    return {
      Beherbergungsidentifikationsnummer: row.Beherbergungsidentifikationsnummer,
      Strasse: row.Strasse,
      Bruttobetrag: row.Bruttobetrag,
      Reinigungsgebuehr: row.Reinigungsgebuehr_num,
      "Zwischensumme (BmGI Beherbergungssteuer)": row["Zwischensumme (BmGI Beherbergungssteuer)"],
      Kommissionen: row.Kommissionen,
      Auszahlungsbetrag: row.Auszahlungsbetrag,
      "Bruttobetrag Inklusive Beherbergungssteuer": row.Bruttobetrag + row.Beherbergungssteuer,
      Beherbergungssteuer: row.Beherbergungssteuer
    };
  });

  const total = buildTotalRow(grouped);
  const columns = buildColumns();
  return {
    columns,
    rows: [...grouped, total],
    dataRowCount: grouped.length
  };
}

function buildTotalRow(rows) {
  return rows.reduce(
    (acc, row) => {
      Object.keys(row).forEach((key) => {
        if (typeof row[key] === "number") acc[key] += row[key];
      });
      return acc;
    },
    {
      Beherbergungsidentifikationsnummer: "GESAMT",
      Strasse: "",
      Bruttobetrag: 0,
      Reinigungsgebuehr: 0,
      "Zwischensumme (BmGI Beherbergungssteuer)": 0,
      Kommissionen: 0,
      Auszahlungsbetrag: 0,
      "Bruttobetrag Inklusive Beherbergungssteuer": 0,
      Beherbergungssteuer: 0
    }
  );
}

function updateTotals() {
  if (!state.output) return;
  const dataRows = state.output.rows.slice(0, state.output.dataRowCount);
  const total = buildTotalRow(dataRows);
  state.output.rows = [...dataRows, total];
}

function groupByIdAndStreet(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const key = `${row.Beherbergungsidentifikationsnummer}||${row.Strasse}`;
    if (!grouped.has(key)) {
      grouped.set(key, { ...row });
      return;
    }
    const current = grouped.get(key);
    Object.keys(row).forEach((col) => {
      if (typeof row[col] === "number") {
        current[col] += row[col];
      }
    });
  });
  return Array.from(grouped.values());
}

function buildGroupHeaderRow(columns) {
  const cells = [];
  let index = 0;

  while (index < columns.length) {
    const label = columns[index].groupLabel || "";
    let span = 1;
    while (
      index + span < columns.length &&
      (columns[index + span].groupLabel || "") === label
    ) {
      span += 1;
    }
    cells.push(`<th colspan="${span}" class="group-row">${label}</th>`);
    index += span;
  }

  return `<tr>${cells.join("")}</tr>`;
}

function renderResultTable(output) {
  els.resultTableHead.innerHTML = "";
  els.resultTableBody.innerHTML = "";
  if (!output || !output.rows.length) return;

  const headers = output.columns;
  const widths = computeColumnWidths(headers);
  const titleRow1 = document.createElement("tr");
  const titleCell1 = document.createElement("th");
  titleCell1.className = "title-row";
  titleCell1.colSpan = headers.length;
  titleCell1.textContent = state.header.titleLine1 || "";
  titleRow1.appendChild(titleCell1);

  const titleRow2 = document.createElement("tr");
  const titleCell2 = document.createElement("th");
  titleCell2.className = "title-row";
  titleCell2.colSpan = headers.length;
  titleCell2.textContent = state.header.titleLine2 || "";
  titleRow2.appendChild(titleCell2);

  const trHead = document.createElement("tr");
  headers.forEach((column, index) => {
    const th = document.createElement("th");
    th.textContent = column.label;
    th.style.minWidth = `${widths[index]}ch`;
    trHead.appendChild(th);
  });
  els.resultTableHead.appendChild(titleRow1);
  els.resultTableHead.appendChild(titleRow2);
  if (state.header.showGroupRow === true) {
    const groupWrap = document.createElement("tbody");
    groupWrap.innerHTML = buildGroupHeaderRow(headers);
    els.resultTableHead.appendChild(groupWrap.firstChild);
  }
  els.resultTableHead.appendChild(trHead);

  output.rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    const isTotal = rowIndex === output.rows.length - 1;
    headers.forEach((column, colIndex) => {
      const td = document.createElement("td");
      const val = row[column.key];
      const display = column.type === "number" ? formatCurrency(val) : val;
      td.textContent = display;
      td.style.minWidth = `${widths[colIndex]}ch`;
      if (!isTotal) {
        td.setAttribute("contenteditable", "true");
        td.dataset.rowIndex = rowIndex.toString();
        td.dataset.colKey = column.key;
        td.dataset.colType = column.type;
        td.addEventListener("focus", () => {
          if (column.type === "number") {
            td.textContent = Number(val || 0).toFixed(2);
          }
        });
        td.addEventListener("blur", () => {
          const rowIdx = Number(td.dataset.rowIndex);
          const colKey = td.dataset.colKey;
          const colType = td.dataset.colType;
          if (Number.isNaN(rowIdx) || rowIdx < 0) return;
          const raw = td.textContent || "";
          if (colType === "number") {
            state.output.rows[rowIdx][colKey] = parseMoneySmart(raw);
          } else {
            state.output.rows[rowIdx][colKey] = raw.trim();
          }
          updateTotals();
          renderResultTable(state.output);
        });
      }
      tr.appendChild(td);
    });
    els.resultTableBody.appendChild(tr);
  });
}

function parseCsv(text, delimiter) {
  const cleaned = text.replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < cleaned.length; i += 1) {
    const char = cleaned[i];
    const next = cleaned[i + 1];
    if (inQuotes) {
      if (char === "\"" && next === "\"") {
        field += "\"";
        i += 1;
      } else if (char === "\"") {
        inQuotes = false;
      } else {
        field += char;
      }
    } else {
      if (char === "\"") {
        inQuotes = true;
      } else if (char === delimiter) {
        row.push(field);
        field = "";
      } else if (char === "\n") {
        row.push(field);
        field = "";
        if (!row.every((cell) => cell.trim() === "")) {
          rows.push(row);
        }
        row = [];
      } else if (char !== "\r") {
        field += char;
      }
    }
  }

  if (field.length || row.length) {
    row.push(field);
    if (!row.every((cell) => cell.trim() === "")) {
      rows.push(row);
    }
  }

  if (!rows.length) return [];
  const headers = rows[0].map((h) => h.trim());
  return rows.slice(1).map((cols) => {
    const rowObj = {};
    headers.forEach((header, index) => {
      rowObj[header] = cols[index] !== undefined ? cols[index].trim() : "";
    });
    return rowObj;
  });
}

function readFileAsTextWithProgress(file, progressEl, nameEl) {
  return new Promise((resolve, reject) => {
    if (!file) {
      resolve("");
      return;
    }
    const reader = new FileReader();
    reader.onprogress = (event) => {
      if (event.lengthComputable) {
        const percent = Math.round((event.loaded / event.total) * 100);
        progressEl.style.width = `${percent}%`;
      }
    };
    reader.onload = () => {
      progressEl.style.width = "100%";
      nameEl.textContent = file.name;
      resolve(reader.result || "");
    };
    reader.onerror = () => {
      progressEl.style.width = "0%";
      reject(new Error("Datei konnte nicht gelesen werden."));
    };
    reader.readAsText(file);
  });
}

function setFile(type, file) {
  if (type === "booking") {
    state.bookingFile = file;
    els.bookingName.textContent = file ? file.name : "Keine Datei";
    els.bookingProgress.style.width = "0%";
  } else {
    state.payoutFile = file;
    els.payoutName.textContent = file ? file.name : "Keine Datei";
    els.payoutProgress.style.width = "0%";
  }
}

function wireDropZones() {
  document.querySelectorAll(".drop-zone").forEach((zone) => {
    zone.addEventListener("dragover", (event) => {
      event.preventDefault();
      zone.classList.add("dragover");
    });
    zone.addEventListener("dragleave", () => {
      zone.classList.remove("dragover");
    });
    zone.addEventListener("drop", (event) => {
      event.preventDefault();
      zone.classList.remove("dragover");
      const type = zone.dataset.drop;
      const file = event.dataTransfer.files[0];
      if (file) setFile(type, file);
    });
  });
}

function computeColumnWidths(columns) {
  return columns.map((column) => {
    const label = column.label || "";
    const base = Math.max(label.length, 6);
    return Math.min(base + 2, 60);
  });
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
function buildXlsx(output) {
  const sheetData = buildSheetXml(output);
  const sheetXml = sheetData.xml;
  const stylesXml = buildStylesXml();
  const tableXml = buildTableXml(output, sheetData.tableRange);
  const sheetRelsXml = buildSheetRelsXml();
  const workbookXml =
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"\n` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n` +
    `<sheets>\n` +
    `<sheet name="Auswertung" sheetId="1" r:id="rId1"/>\n` +
    `</sheets>\n` +
    `</workbook>`;

  const relsXml =
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n` +
    `<Relationship Id="rId1" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ` +
    `Target="worksheets/sheet1.xml"/>\n` +
    `<Relationship Id="rId2" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" ` +
    `Target="styles.xml"/>\n` +
    `</Relationships>`;

  const rootRelsXml =
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n` +
    `<Relationship Id="rId1" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ` +
    `Target="xl/workbook.xml"/>\n` +
    `</Relationships>`;

  const contentTypesXml =
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n` +
    `<Default Extension="xml" ContentType="application/xml"/>\n` +
    `<Override PartName="/xl/workbook.xml" ` +
    `ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n` +
    `<Override PartName="/xl/styles.xml" ` +
    `ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ` +
    `ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n` +
    `<Override PartName="/xl/tables/table1.xml" ` +
    `ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>\n` +
    `</Types>`;

  const files = [
    { name: "[Content_Types].xml", data: contentTypesXml },
    { name: "_rels/.rels", data: rootRelsXml },
    { name: "xl/workbook.xml", data: workbookXml },
    { name: "xl/_rels/workbook.xml.rels", data: relsXml },
    { name: "xl/styles.xml", data: stylesXml },
    { name: "xl/worksheets/sheet1.xml", data: sheetXml },
    { name: "xl/worksheets/_rels/sheet1.xml.rels", data: sheetRelsXml },
    { name: "xl/tables/table1.xml", data: tableXml }
  ];

  return makeZip(files);
}

function buildSheetXml(output) {
  if (!output.rows.length) {
    return { xml: "", tableRange: "A1:A1" };
  }

  const columnCount = output.columns.length;
  const lastColumn = columnName(columnCount);
  const titleLine1 = state.header.titleLine1 || "";
  const titleLine2 = state.header.titleLine2 || "";
  const showGroupRow = state.header.showGroupRow === true;

  const rows = [];
  const merges = [];

  rows.push({
    cells: [
      {
        value: titleLine1,
        type: "string",
        style: 2
      }
    ]
  });
  merges.push(`A1:${lastColumn}1`);

  rows.push({
    cells: [
      {
        value: titleLine2,
        type: "string",
        style: 2
      }
    ]
  });
  merges.push(`A2:${lastColumn}2`);

  if (showGroupRow) {
    const groupCells = [];
    let colIndex = 0;
    while (colIndex < output.columns.length) {
      const label = output.columns[colIndex].groupLabel || "";
      let span = 1;
      while (
        colIndex + span < output.columns.length &&
        (output.columns[colIndex + span].groupLabel || "") === label
      ) {
        span += 1;
      }
      if (label) {
        merges.push(`${columnName(colIndex + 1)}3:${columnName(colIndex + span)}3`);
      }
      for (let i = 0; i < span; i += 1) {
        groupCells.push({
          value: i === 0 ? label : "",
          type: "string",
          style: 1
        });
      }
      colIndex += span;
    }
    rows.push({ cells: groupCells });
  }

  const headerRow = {
    cells: output.columns.map((column) => ({
      value: column.label,
      type: "string",
      style: 1
    }))
  };
  rows.push(headerRow);

  output.rows.forEach((row) => {
    const cells = output.columns.map((column) => {
      const value = row[column.key];
      if (column.type === "number" && typeof value === "number" && !Number.isNaN(value)) {
        return { value, type: "number", style: 3 };
      }
      return { value, type: "string", style: 0 };
    });
    rows.push({ cells });
  });

  const startRowIndex = showGroupRow ? 4 : 3;
  const headerRowIndex = startRowIndex;
  const lastRowIndex = headerRowIndex + output.rows.length;
  const tableRange = `A${headerRowIndex}:${lastColumn}${lastRowIndex}`;

  const xmlRows = rows
    .map((row, rowIndex) => {
      const rowNumber = rowIndex + 1;
      const cells = row.cells
        .map((cell, cellIndex) => {
          const cellRef = `${columnName(cellIndex + 1)}${rowNumber}`;
          const styleAttr = cell.style !== undefined ? ` s="${cell.style}"` : "";
          if (cell.type === "number") {
            return `<c r="${cellRef}"${styleAttr}><v>${Number(cell.value).toFixed(2)}</v></c>`;
          }
          return (
            `<c r="${cellRef}" t="inlineStr"${styleAttr}>` +
            `<is><t>${escapeXml(cell.value)}</t></is></c>`
          );
        })
        .join("");
      return `<row r="${rowNumber}">${cells}</row>`;
    })
    .join("");

  const widths = computeColumnWidths(output.columns);
  const colsXml =
    `<cols>` +
    widths
      .map(
        (width, index) =>
          `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`
      )
      .join("") +
    `</cols>`;

  const mergeXml =
    merges.length > 0
      ? `<mergeCells count="${merges.length}">` +
        merges.map((ref) => `<mergeCell ref="${ref}"/>`).join("") +
        `</mergeCells>`
      : "";

  const tablePartsXml =
    `<tableParts count="1"><tablePart r:id="rId1"/></tableParts>`;

  return {
    xml:
      `<?xml version="1.0" encoding="UTF-8"?>\n` +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"\n` +
      ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n` +
      `${colsXml}\n` +
      `<sheetData>\n` +
      `${xmlRows}\n` +
      `</sheetData>\n` +
      `${mergeXml}\n` +
      `${tablePartsXml}\n` +
      `</worksheet>`,
    tableRange
  };
}

function buildStylesXml() {
  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n` +
    `<numFmts count="1">\n` +
    `<numFmt numFmtId="164" formatCode="[$-de-DE] #,##0.00 &quot;EUR&quot;"/>\n` +
    `</numFmts>\n` +
    `<fonts count="3">\n` +
    `<font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>\n` +
    `<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font>\n` +
    `<font><b/><sz val="22"/><color theme="1"/><name val="Calibri"/></font>\n` +
    `</fonts>\n` +
    `<fills count="2">\n` +
    `<fill><patternFill patternType="none"/></fill>\n` +
    `<fill><patternFill patternType="gray125"/></fill>\n` +
    `</fills>\n` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n` +
    `<cellXfs count="4">\n` +
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n` +
    `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1">` +
    `<alignment horizontal="center" vertical="center"/></xf>\n` +
    `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1">` +
    `<alignment horizontal="center" vertical="center"/></xf>\n` +
    `<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>\n` +
    `</cellXfs>\n` +
    `</styleSheet>`
  );
}

function buildTableXml(output, tableRange) {
  const columnsXml = output.columns
    .map(
      (column, index) =>
        `<tableColumn id="${index + 1}" name="${escapeXml(column.label)}"/>`
    )
    .join("");

  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
    `id="1" name="Auswertung" displayName="Auswertung" ref="${tableRange}" ` +
    `headerRowCount="1" totalsRowCount="0">\n` +
    `<autoFilter ref="${tableRange}"/>\n` +
    `<tableColumns count="${output.columns.length}">${columnsXml}</tableColumns>\n` +
    `<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" ` +
    `showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>\n` +
    `</table>`
  );
}

function buildSheetRelsXml() {
  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n` +
    `<Relationship Id="rId1" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" ` +
    `Target="../tables/table1.xml"/>\n` +
    `</Relationships>`
  );
}

function columnName(index) {
  let name = "";
  let current = index;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    current = Math.floor((current - 1) / 26);
  }
  return name;
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) {
    crc ^= bytes[i];
    for (let j = 0; j < 8; j += 1) {
      const mask = -(crc & 1);
      crc = (crc >>> 1) ^ (0xedb88320 & mask);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function makeZip(files) {
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  files.forEach((file) => {
    const nameBytes = encoder.encode(file.name);
    const dataBytes = typeof file.data === "string" ? encoder.encode(file.data) : file.data;
    const crc = crc32(dataBytes);
    const localHeader = buildLocalHeader(nameBytes.length, dataBytes.length, crc);
    const localRecord = concatArrays([localHeader, nameBytes, dataBytes]);
    localParts.push(localRecord);

    const centralHeader = buildCentralHeader(
      nameBytes.length,
      dataBytes.length,
      crc,
      offset
    );
    const centralRecord = concatArrays([centralHeader, nameBytes]);
    centralParts.push(centralRecord);

    offset += localRecord.length;
  });

  const centralDir = concatArrays(centralParts);
  const endRecord = buildEndHeader(centralParts.length, centralDir.length, offset);

  return concatArrays([...localParts, centralDir, endRecord]);
}

function buildLocalHeader(nameLength, dataLength, crc) {
  const header = [];
  pushUint32(header, 0x04034b50);
  pushUint16(header, 20);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint32(header, crc);
  pushUint32(header, dataLength);
  pushUint32(header, dataLength);
  pushUint16(header, nameLength);
  pushUint16(header, 0);
  return new Uint8Array(header);
}

function buildCentralHeader(nameLength, dataLength, crc, offset) {
  const header = [];
  pushUint32(header, 0x02014b50);
  pushUint16(header, 20);
  pushUint16(header, 20);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint32(header, crc);
  pushUint32(header, dataLength);
  pushUint32(header, dataLength);
  pushUint16(header, nameLength);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint32(header, 0);
  pushUint32(header, offset);
  return new Uint8Array(header);
}

function buildEndHeader(recordCount, centralSize, centralOffset) {
  const header = [];
  pushUint32(header, 0x06054b50);
  pushUint16(header, 0);
  pushUint16(header, 0);
  pushUint16(header, recordCount);
  pushUint16(header, recordCount);
  pushUint32(header, centralSize);
  pushUint32(header, centralOffset);
  pushUint16(header, 0);
  return new Uint8Array(header);
}

function pushUint16(buffer, value) {
  buffer.push(value & 0xff, (value >> 8) & 0xff);
}

function pushUint32(buffer, value) {
  buffer.push(
    value & 0xff,
    (value >> 8) & 0xff,
    (value >> 16) & 0xff,
    (value >> 24) & 0xff
  );
}

function concatArrays(arrays) {
  const total = arrays.reduce((sum, array) => sum + array.length, 0);
  const result = new Uint8Array(total);
  let offset = 0;
  arrays.forEach((array) => {
    result.set(array, offset);
    offset += array.length;
  });
  return result;
}

function downloadExcel(output) {
  const xlsxData = buildXlsx(output);
  const blob = new Blob([xlsxData], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "booking_auszahlungsuebersicht.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
els.bookingFile.addEventListener("change", (event) => {
  setFile("booking", event.target.files[0] || null);
});

els.payoutFile.addEventListener("change", (event) => {
  setFile("payout", event.target.files[0] || null);
});

els.addMappingRowBtn.addEventListener("click", () => {
  state.mapping.push({ Keyword: "", Beherbergungsidentifikationsnummer: "", Strasse: "" });
  renderMappingTable();
});

els.saveMappingBtn.addEventListener("click", () => {
  const invalid = state.mapping.some(
    (row) =>
      !String(row.Keyword || "").trim() ||
      !String(row.Beherbergungsidentifikationsnummer || "").trim()
  );
  if (invalid) {
    setStatus("Bitte: Keyword und Beherbergungsidentifikationsnummer duerfen nicht leer sein.", "error");
    return;
  }
  saveMapping();
  setStatus("Mapping gespeichert.");
});

els.resetMappingBtn.addEventListener("click", () => {
  state.mapping = DEFAULT_MAPPING.slice();
  saveMapping();
  renderMappingTable();
  setStatus("Mapping zurueckgesetzt.");
});

els.saveRenamesBtn.addEventListener("click", () => {
  saveRenames();
  saveGroupLabels();
  saveColumnOrder();
  if (state.output) {
    state.output.columns = buildColumns();
    renderResultTable(state.output);
  }
  setStatus("Spaltennamen gespeichert.");
});

els.resetRenamesBtn.addEventListener("click", () => {
  state.renames = { ...DEFAULT_RENAMES };
  state.groupLabels = { ...DEFAULT_GROUP_LABELS };
  state.columnOrder = DEFAULT_COLUMN_ORDER.slice();
  saveRenames();
  saveGroupLabels();
  saveColumnOrder();
  renderRenamesTable();
  if (state.output) {
    state.output.columns = buildColumns();
    renderResultTable(state.output);
  }
  setStatus("Spaltennamen zurueckgesetzt.");
});

els.titleLine1.addEventListener("input", () => {
  state.header.titleLine1 = els.titleLine1.value;
  saveHeader();
  if (state.output) {
    renderResultTable(state.output);
  }
});

els.titleLine2.addEventListener("input", () => {
  state.header.titleLine2 = els.titleLine2.value;
  saveHeader();
  if (state.output) {
    renderResultTable(state.output);
  }
});

els.showGroupRow.addEventListener("change", () => {
  state.header.showGroupRow = els.showGroupRow.checked;
  saveHeader();
  if (state.output) {
    renderResultTable(state.output);
  }
});

els.openOverlayBtn.addEventListener("click", () => {
  openOverlay();
});

els.closeOverlayBtn.addEventListener("click", () => {
  closeOverlay();
});

els.closeOverlayBtnFooter.addEventListener("click", () => {
  closeOverlay();
});

els.overlay.addEventListener("click", (event) => {
  if (event.target === els.overlay) {
    closeOverlay();
  }
});

els.runBtn.addEventListener("click", async () => {
  if (!state.bookingFile || !state.payoutFile) {
    setStatus("Bitte BookingList CSV und Payout CSV hochladen.", "error");
    return;
  }
  try {
    const bookingText = await readFileAsTextWithProgress(
      state.bookingFile,
      els.bookingProgress,
      els.bookingName
    );
    const payoutText = await readFileAsTextWithProgress(
      state.payoutFile,
      els.payoutProgress,
      els.payoutName
    );
    const bookingRows = parseCsv(bookingText, ";");
    const payoutRows = parseCsv(payoutText, ",");
    state.output = computeResult(bookingRows, payoutRows);
    renderResultTable(state.output);
    els.downloadBtn.disabled = false;
    setStatus(`Fertig. Zeilen: ${state.output.rows.length}`);
  } catch (err) {
    setStatus(err.message || "Fehler bei der Auswertung.", "error");
  }
});

els.downloadBtn.addEventListener("click", () => {
  if (!state.output || !state.output.rows.length) return;
  downloadExcel(state.output);
});

loadState();
renderHeaderInputs();
renderMappingTable();
renderRenamesTable();
wireDropZones();
