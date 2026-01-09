const pdfInput = document.getElementById("pdfFiles");
const dropZone = document.getElementById("dropZone");
const pickPdfBtn = document.getElementById("pickPdfBtn");
const fileList = document.getElementById("fileList");
const progressBar = document.getElementById("progressBar");
const statusEl = document.getElementById("status");
const previewBtn = document.getElementById("previewBtn");
const downloadBtn = document.getElementById("downloadBtn");
const snapshotBtn = document.getElementById("snapshotBtn");
const mappingFile = document.getElementById("mappingFile");
const mappingList = document.getElementById("mappingList");
const applyMappingBtn = document.getElementById("applyMappingBtn");
const resetMappingBtn = document.getElementById("resetMappingBtn");
const addMappingBtn = document.getElementById("addMappingBtn");
const titleLine1Input = document.getElementById("titleLine1");
const titleLine2Input = document.getElementById("titleLine2");
const groupRowToggle = document.getElementById("groupRowToggle");
const previewTable = document.getElementById("previewTable");
const tableHead = previewTable.querySelector("thead");
const tableBody = previewTable.querySelector("tbody");

const MAPPING_KEY = "direktbuchungenMapping";
const HEADER_KEY = "direktbuchungenHeader";
const DASHBOARD_KEY = "dashboard_snapshots_v1";
const SYSTEM_ID = "direkt";

const firebase = window.firebaseServices || {};
const {
  auth,
  db,
  onAuthStateChanged,
  doc,
  getDoc,
  setDoc,
  serverTimestamp
} = firebase;

const columns = [
  { key: "id", label: "BeherbergungsID", type: "text", groupLabel: "" },
  { key: "street", label: "Strasse", type: "text", groupLabel: "" },
  { key: "brutto", label: "Bruttoeinkuenfte", type: "number", groupLabel: "SP1" },
  { key: "reinigung", label: "Reinigungskosten", type: "number", groupLabel: "SP2" },
  { key: "bmgl", label: "BmGl Beherbergungssteuer", type: "number", groupLabel: "SP3" },
  { key: "steuer", label: "Beherbergungssteuer 5% von SP3", type: "number", groupLabel: "SP4" }
];

const state = {
  mapping: {},
  mappingKeys: [],
  output: { columns, rows: [] },
  header: {
    titleLine1: "Direktbuchungen Auswertung",
    titleLine2: "",
    showGroupRow: true
  }
};
let currentUser = null;

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = "lib/pdf.worker.min.js";
  pdfjsLib.disableWorker = true;
}

function setStatus(message, tone = "") {
  statusEl.textContent = message;
  statusEl.dataset.tone = tone;
}

function getSystemDoc() {
  if (!db || !currentUser) return null;
  return doc(db, "users", currentUser.uid, "systems", SYSTEM_ID);
}

async function loadRemoteState() {
  const ref = getSystemDoc();
  if (!ref) return;
  const snapshot = await getDoc(ref);
  const data = snapshot.exists() ? snapshot.data() : null;
  if (data && data.payload) {
    applyStoragePayload(data.payload);
  }
}

async function saveRemoteState() {
  const ref = getSystemDoc();
  if (!ref) return;
  const payload = buildStoragePayload();
  await setDoc(
    ref,
    { payload, updatedAt: serverTimestamp() },
    { merge: true }
  );
}

let saveTimer = null;
function scheduleRemoteSave() {
  if (!currentUser) return;
  if (saveTimer) {
    clearTimeout(saveTimer);
  }
  saveTimer = setTimeout(() => {
    saveRemoteState().catch(() => {});
  }, 350);
}

async function appendDashboardEntry(metrics) {
  if (!db || !currentUser) return;
  const ref = doc(db, "users", currentUser.uid, "dashboard", "summary");
  const snapshot = await getDoc(ref);
  const data = snapshot.exists() ? snapshot.data() : null;
  const history = Array.isArray(data && data.snapshots) ? data.snapshots : [];
  history.push({
    system: SYSTEM_ID,
    ts: Date.now(),
    metrics
  });
  await setDoc(
    ref,
    { snapshots: history.slice(-200), updatedAt: serverTimestamp() },
    { merge: true }
  );
}

function setProgress(percent) {
  const safe = Math.max(0, Math.min(100, percent));
  progressBar.style.width = `${safe}%`;
}

function updateFileList(files) {
  fileList.innerHTML = "";
  if (!files || !files.length) {
    fileList.textContent = "Keine PDFs ausgew\u00e4hlt.";
    return;
  }
  Array.from(files).forEach((file) => {
    const item = document.createElement("div");
    item.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
    fileList.appendChild(item);
  });
}

function normalize(text) {
  return String(text || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ß/g, "ss")
    .replace(/ue/g, "u")
    .replace(/oe/g, "o")
    .replace(/ae/g, "a")
    .replace(/strasse/g, "str")
    .replace(/straße/g, "str")
    .replace(/[^a-z0-9]/g, "");
}

function parseGermanFloat(value) {
  if (value === null || value === undefined) {
    return 0;
  }
  const rawMatch = String(value).match(/[\d.]+,\d{2}/);
  let raw = rawMatch ? rawMatch[0] : String(value);
  raw = raw.replace(/[^\d,.-]/g, "");
  raw = raw.replace(/\./g, "").replace(",", ".");
  const number = Number.parseFloat(raw);
  return Number.isNaN(number) ? 0 : number;
}

function findAddressInLine(line) {
  const lineNorm = normalize(line);
  for (const key of state.mappingKeys) {
    if (lineNorm.includes(normalize(key))) {
      return key;
    }
  }
  return null;
}

function findIdForAddress(address) {
  const normalized = normalize(address);
  for (const key of state.mappingKeys) {
    if (normalized.includes(normalize(key))) {
      return state.mapping[key];
    }
  }
  return "ID FEHLT";
}

function detectDelimiter(text) {
  const firstLine = text.split(/\r?\n/)[0] || "";
  const commas = (firstLine.match(/,/g) || []).length;
  const semicolons = (firstLine.match(/;/g) || []).length;
  return semicolons > commas ? ";" : ",";
}

function normalizeHeader(header) {
  return String(header || "")
    .toLowerCase()
    .replace(/ä/g, "ae")
    .replace(/ö/g, "oe")
    .replace(/ü/g, "ue")
    .replace(/ß/g, "ss")
    .replace(/[^a-z0-9]/g, "");
}

function setMapping(mapping) {
  state.mapping = mapping;
  state.mappingKeys = Object.keys(mapping).sort((a, b) => b.length - a.length);
  localStorage.setItem(MAPPING_KEY, JSON.stringify(mapping));
  renderMappingList(mapping);
  scheduleRemoteSave();
}

function loadStoredMapping() {
  try {
    const stored = JSON.parse(localStorage.getItem(MAPPING_KEY));
    if (stored && typeof stored === "object") {
      setMapping(stored);
      return;
    }
  } catch (error) {
    // ignore
  }
  renderMappingList({});
}

function setHeaderConfig(config) {
  state.header = {
    titleLine1: config.titleLine1 || "",
    titleLine2: config.titleLine2 || "",
    showGroupRow: config.showGroupRow !== false
  };
  localStorage.setItem(HEADER_KEY, JSON.stringify(state.header));
  titleLine1Input.value = state.header.titleLine1;
  titleLine2Input.value = state.header.titleLine2;
  groupRowToggle.checked = state.header.showGroupRow;
  scheduleRemoteSave();
}

function buildStoragePayload() {
  return {
    storage: {
      [MAPPING_KEY]: localStorage.getItem(MAPPING_KEY) || "",
      [HEADER_KEY]: localStorage.getItem(HEADER_KEY) || "",
      [DASHBOARD_KEY]: localStorage.getItem(DASHBOARD_KEY) || ""
    }
  };
}

function applyStoragePayload(payload) {
  if (!payload || !payload.storage) return;
  Object.entries(payload.storage).forEach(([key, value]) => {
    if (value === null || value === undefined) return;
    localStorage.setItem(key, value);
  });
}

async function loadServerState() {
  await loadRemoteState();
  loadStoredMapping();
  loadStoredHeader();
}

async function saveServerState() {
  await saveRemoteState();
}

function loadStoredHeader() {
  try {
    const stored = JSON.parse(localStorage.getItem(HEADER_KEY));
    if (stored && typeof stored === "object") {
      setHeaderConfig(stored);
      return;
    }
  } catch (error) {
    // ignore
  }
  setHeaderConfig(state.header);
}

async function parseMappingFile(file) {
  const text = await file.text();
  const delimiter = detectDelimiter(text);
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const mapping = {};
  if (!lines.length) {
    return mapping;
  }

  const headerLine = normalizeHeader(lines[0]);
  const startIndex =
    headerLine.includes("adresse") && headerLine.includes("id") ? 1 : 0;

  for (let i = startIndex; i < lines.length; i += 1) {
    const parts = lines[i].split(delimiter).map((part) => part.trim());
    if (parts.length >= 2 && parts[0] && parts[1]) {
      mapping[parts[0]] = parts[1];
    }
  }

  return mapping;
}

function createMappingRow(address = "", id = "") {
  const row = document.createElement("div");
  row.className = "mapping-row";

  const addressInput = document.createElement("input");
  addressInput.type = "text";
  addressInput.placeholder = "Adresse";
  addressInput.value = address;

  const idInput = document.createElement("input");
  idInput.type = "text";
  idInput.placeholder = "ID";
  idInput.value = id;

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.textContent = "Entfernen";
  removeBtn.addEventListener("click", () => row.remove());

  row.appendChild(addressInput);
  row.appendChild(idInput);
  row.appendChild(removeBtn);

  return row;
}

function renderMappingList(mapping) {
  mappingList.innerHTML = "";
  const entries = Object.entries(mapping);
  if (!entries.length) {
    mappingList.appendChild(createMappingRow("", ""));
    return;
  }
  entries.forEach(([address, id]) => {
    mappingList.appendChild(createMappingRow(address, id));
  });
}

function collectMappingFromList() {
  const mapping = {};
  Array.from(mappingList.querySelectorAll(".mapping-row")).forEach((row) => {
    const inputs = row.querySelectorAll("input");
    const address = inputs[0] ? inputs[0].value.trim() : "";
    const id = inputs[1] ? inputs[1].value.trim() : "";
    if (address && id) {
      mapping[address] = id;
    }
  });
  return mapping;
}

async function extractLinesFromPdf(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const allLines = [];
  let fullText = "";

  for (let pageIndex = 1; pageIndex <= pdf.numPages; pageIndex += 1) {
    const page = await pdf.getPage(pageIndex);
    const textContent = await page.getTextContent();
    const items = textContent.items || [];
    const chunks = items
      .map((item) => ({
        str: item.str || "",
        x: item.transform ? item.transform[4] : 0,
        y: item.transform ? item.transform[5] : 0
      }))
      .filter((item) => item.str.trim().length > 0);

    chunks.sort((a, b) => (b.y - a.y) || (a.x - b.x));

    const lines = [];
    let current = null;
    const threshold = 2;

    chunks.forEach((item) => {
      if (!current || Math.abs(item.y - current.y) > threshold) {
        if (current) {
          lines.push(current.text.join(" "));
        }
        current = { y: item.y, text: [item.str] };
      } else {
        current.text.push(item.str);
      }
    });

    if (current) {
      lines.push(current.text.join(" "));
    }

    let pageText = lines.join("\n");
    fullText += `${pageText}\n`;
    pageText = pageText.replace(/PrimeTimeSuite/g, "\nPrimeTimeSuite");
    pageText = pageText.replace(/Prime TimeSuite/g, "\nPrime TimeSuite");

    const pageLines = pageText
      .split(/\n+/)
      .map((line) => line.trim())
      .filter(Boolean);

    allLines.push(...pageLines);
  }

  return { lines: allLines, text: fullText };
}

function parseLines(lines, vatMode) {
  const bookings = [];
  let totalCleaningValue = 0;
  let cleaningVatRate = 0;
  let singleEndbetrag = 0;
  let singleReinigungValue = 0;
  let singleSteuer = 0;

  lines.forEach((line) => {
    const endMatch = line.match(/Endbetrag:.*?([\d.,]+)/i);
    if (endMatch) {
      singleEndbetrag = parseGermanFloat(endMatch[1]);
    }

    const address = findAddressInLine(line);
    const amounts = line.match(/[\d.]+,\d{2}/g) || [];

    if (address && amounts.length) {
      const preis = parseGermanFloat(amounts[amounts.length - 1]);
      if (
        preis > 0 &&
        !/Rein\w*/i.test(line) &&
        !line.includes("Beherbergungssteuer")
      ) {
        bookings.push({ street: address, netto: preis });
      }
    }

    if (/Rein\w*/i.test(line)) {
      const values = amounts.map((amount) => parseGermanFloat(amount)).filter((val) => val > 0);
      if (values.length) {
        const lineCleaningTotal = Math.max(...values);
        totalCleaningValue += lineCleaningTotal;
        const pctMatch = line.match(/(\d{1,2}(?:,\d{1,2})?)\s*%/);
        if (pctMatch) {
          cleaningVatRate = parseGermanFloat(pctMatch[1]);
        }
        if (cleaningVatRate === 0) {
          cleaningVatRate = 19;
        }
        if (vatMode === "netto") {
          singleReinigungValue = lineCleaningTotal * (1 + cleaningVatRate / 100);
        } else {
          singleReinigungValue = lineCleaningTotal;
        }
      }
    }

    if (line.includes("Beherbergungssteuer")) {
      if (!line.includes("5%")) {
        if (amounts.length) {
          singleSteuer = parseGermanFloat(amounts[amounts.length - 1]);
        }
      }
    }
  });

  const extracted = [];
  const numBookings = bookings.length;

  if (numBookings <= 1) {
    const finalAddress = numBookings === 1 ? bookings[0].street : "Unbekannt";
    if (singleEndbetrag === 0 && numBookings === 1) {
      const rentGross = bookings[0].netto * (vatMode === "netto" ? 1.07 : 1.0);
      singleEndbetrag = rentGross + singleReinigungValue + singleSteuer;
    }
    const bruttoEinkuenfte = singleEndbetrag - singleSteuer;
    extracted.push({
      street: finalAddress,
      brutto: bruttoEinkuenfte,
      reinigung: singleReinigungValue
    });
  } else {
    if (cleaningVatRate === 0) {
      cleaningVatRate = 19;
    }
    const totalCleaningBrutto =
      vatMode === "netto"
        ? totalCleaningValue * (1 + cleaningVatRate / 100)
        : totalCleaningValue;
    const cleaningPerBooking = totalCleaningBrutto / numBookings;
    bookings.forEach((booking) => {
      const rentGross = booking.netto * (vatMode === "netto" ? 1.07 : 1.0);
      const revenueGross = rentGross + cleaningPerBooking;
      extracted.push({
        street: booking.street,
        brutto: revenueGross,
        reinigung: cleaningPerBooking
      });
    });
  }

  return extracted;
}

function buildOutput(entries) {
  const enriched = entries.map((entry) => {
    const bmgl = entry.brutto - entry.reinigung;
    const steuer = bmgl > 0 ? bmgl * 0.05 : 0;
    return { ...entry, bmgl, steuer };
  });

  const grouped = new Map();
  enriched.forEach((entry) => {
    if (!grouped.has(entry.street)) {
      grouped.set(entry.street, {
        street: entry.street,
        brutto: 0,
        reinigung: 0,
        bmgl: 0,
        steuer: 0
      });
    }
    const item = grouped.get(entry.street);
    item.brutto += entry.brutto;
    item.reinigung += entry.reinigung;
    item.bmgl += entry.bmgl;
    item.steuer += entry.steuer;
  });

  const rows = [];
  const totals = { brutto: 0, reinigung: 0, bmgl: 0, steuer: 0 };

  grouped.forEach((value) => {
    const id = findIdForAddress(value.street);
    rows.push({ id, street: value.street, ...value });
    totals.brutto += value.brutto;
    totals.reinigung += value.reinigung;
    totals.bmgl += value.bmgl;
    totals.steuer += value.steuer;
  });

  rows.push({
    id: "GESAMT",
    street: "",
    brutto: totals.brutto,
    reinigung: totals.reinigung,
    bmgl: totals.bmgl,
    steuer: totals.steuer
  });

  return { columns, rows };
}

function formatCurrency(value) {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return "0,00 €";
  }
  return value.toLocaleString("de-DE", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function renderTable(output) {
  if (!output.rows.length) {
    tableHead.innerHTML = "";
    tableBody.innerHTML = "";
    return;
  }

  const columnCount = output.columns.length;
  const titleLine1 = state.header.titleLine1 || "";
  const titleLine2 = state.header.titleLine2 || "";
  const showGroupRow = state.header.showGroupRow !== false;

  const groupRow = showGroupRow
    ? buildGroupHeaderRow(output.columns)
    : "";

  tableHead.innerHTML =
    `<tr><th colspan="${columnCount}" class="title-row">${titleLine1}</th></tr>` +
    `<tr><th colspan="${columnCount}" class="title-row">${titleLine2}</th></tr>` +
    groupRow +
    `<tr>${output.columns
      .map((column) => `<th>${column.label}</th>`)
      .join("")}</tr>`;

  tableBody.innerHTML = output.rows
    .map((row, rowIndex) => {
      const isTotal = rowIndex === output.rows.length - 1;
      const cells = output.columns
        .map((column) => {
          const value = row[column.key];
          if (column.type === "number") {
            return `<td>${formatCurrency(value)}</td>`;
          }
          return `<td>${value || ""}</td>`;
        })
        .join("");
      return `<tr class="${isTotal ? "total" : ""}">${cells}</tr>`;
    })
    .join("");
}

function findColumnKey(output, matchers) {
  const list = Array.isArray(matchers) ? matchers : [matchers];
  const column = output.columns.find((col) =>
    list.some((matcher) => matcher.test(col.label))
  );
  return column ? column.key : null;
}

function pushDashboardEntry(metrics) {
  try {
    const raw = localStorage.getItem(DASHBOARD_KEY);
    const history = raw ? JSON.parse(raw) : [];
    const safe = Array.isArray(history) ? history : [];
    safe.push({
      system: "direkt",
      ts: Date.now(),
      metrics
    });
    localStorage.setItem(DASHBOARD_KEY, JSON.stringify(safe.slice(-200)));
  } catch (error) {
    // ignore storage errors
  }
  appendDashboardEntry(metrics).catch(() => {});
}

function updateDashboardFromOutput(output) {
  if (!output || !output.rows || output.rows.length < 2) {
    return;
  }
  const totalRow = output.rows[output.rows.length - 1];
  const revenueKey = findColumnKey(output, [/\\bSP1\\b/i, /Brutto/i]) || "brutto";
  const feesKey = findColumnKey(output, [/\\bSP5\\b/i, /Service/i]);
  const taxKey =
    findColumnKey(output, [/\\bSP4\\b/i, /5%/i, /von SP3/i]) || "steuer";
  const revenue = Number(totalRow[revenueKey] || 0);
  const fees = feesKey ? Math.abs(Number(totalRow[feesKey] || 0)) : 0;
  const tax = Number(totalRow[taxKey] || 0);
  const payout = revenue;
  pushDashboardEntry({
    revenue,
    fees,
    tax,
    payout,
    period: state.header.titleLine2 || ""
  });
}

function buildGroupHeaderRow(columnsList) {
  const cells = [];
  let index = 0;

  while (index < columnsList.length) {
    const label = columnsList[index].groupLabel || "";
    let span = 1;
    while (
      index + span < columnsList.length &&
      (columnsList[index + span].groupLabel || "") === label
    ) {
      span += 1;
    }
    const safeLabel = label || "";
    cells.push(`<th colspan="${span}" class="group-row">${safeLabel}</th>`);
    index += span;
  }

  return `<tr>${cells.join("")}</tr>`;
}

async function buildPreview() {
  if (!pdfInput.files.length) {
    setStatus("Bitte zuerst PDFs ausw\u00e4hlen.");
    return;
  }
  if (!window.pdfjsLib) {
    setStatus("PDF.js konnte nicht geladen werden.");
    return;
  }

  previewBtn.disabled = true;
  downloadBtn.disabled = true;
  setStatus("PDFs werden verarbeitet...");
  setProgress(0);

  const entries = [];
  const files = Array.from(pdfInput.files);

  for (let i = 0; i < files.length; i += 1) {
    const file = files[i];
    const { lines, text } = await extractLinesFromPdf(file);
    const hasUstIncluded =
      text.includes("USt. enthalten") || text.includes("USt enthalten");
    const vatMode = hasUstIncluded ? "brutto" : "netto";
    const extracted = parseLines(lines, vatMode);
    entries.push(...extracted);
    setProgress(Math.round(((i + 1) / files.length) * 100));
  }

  state.output = buildOutput(entries);
  renderTable(state.output);
  if (snapshotBtn) {
    snapshotBtn.disabled = false;
  }
  downloadBtn.disabled = state.output.rows.length === 0;
  previewBtn.disabled = false;
  setStatus(`Fertig. ${state.output.rows.length} Zeilen erzeugt.`);
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
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

function computeColumnWidths(output) {
  return output.columns.map((column) => {
    let width = (column.label || "").length;
    output.rows.forEach((row) => {
      const raw = row[column.key];
      const text =
        column.type === "number" ? formatCurrency(raw) : String(raw || "");
      width = Math.max(width, text.length);
    });
    return Math.min(Math.max(width + 2, 10), 40);
  });
}

function buildSheetXml(output) {
  if (!output.rows.length) {
    return { xml: "", tableRange: "A1:A1" };
  }

  const rows = [];
  const columnCount = output.columns.length;
  const lastColumn = columnName(columnCount);
  const titleLine1 = state.header.titleLine1 || "";
  const titleLine2 = state.header.titleLine2 || "";
  const showGroupRow = state.header.showGroupRow !== false;
  const merges = [];

  rows.push({
    cells: [{ value: titleLine1, type: "string", style: 1 }]
  });
  merges.push(`A1:${lastColumn}1`);

  rows.push({
    cells: [{ value: titleLine2, type: "string", style: 1 }]
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
        merges.push(
          `${columnName(colIndex + 1)}3:${columnName(colIndex + span)}3`
        );
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

  rows.push({
    cells: output.columns.map((column) => ({
      value: column.label,
      type: "string",
      style: 1
    }))
  });

  output.rows.forEach((row) => {
    const cells = output.columns.map((column) => {
      const value = row[column.key];
      if (column.type === "number" && typeof value === "number") {
        return { value, type: "number", style: 2 };
      }
      return { value, type: "string", style: 0 };
    });
    rows.push({ cells });
  });

  const headerRowIndex = showGroupRow ? 4 : 3;
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

  const widths = computeColumnWidths(output);
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
    `<numFmt numFmtId="164" formatCode="[$€-de-DE] #,##0.00"/>\n` +
    `</numFmts>\n` +
    `<fonts count="2">\n` +
    `<font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>\n` +
    `<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font>\n` +
    `</fonts>\n` +
    `<fills count="2">\n` +
    `<fill><patternFill patternType="none"/></fill>\n` +
    `<fill><patternFill patternType="gray125"/></fill>\n` +
    `</fills>\n` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n` +
    `<cellXfs count="3">\n` +
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n` +
    `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1">` +
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

function downloadExcel(output) {
  const xlsxData = buildXlsx(output);
  const blob = new Blob([xlsxData], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "direktbuchungen_auswertung.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
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

previewBtn.addEventListener("click", () => {
  buildPreview().catch((error) => {
    setStatus(`Fehler: ${error.message}`);
    previewBtn.disabled = false;
  });
});

downloadBtn.addEventListener("click", () => {
  if (!state.output.rows.length) {
    setStatus("Bitte zuerst eine Vorschau erzeugen.");
    return;
  }
  downloadExcel(state.output);
});

if (snapshotBtn) {
  snapshotBtn.addEventListener("click", () => {
    if (!state.output.rows.length) {
      setStatus("Bitte zuerst eine Vorschau erzeugen.");
      return;
    }
    if (!state.header.titleLine2) {
      setStatus("Bitte Titelzeile 2 (Zeitraum) setzen, bevor du speicherst.");
      return;
    }
    updateDashboardFromOutput(state.output);
    setStatus("Snapshot gespeichert.");
  });
}

dropZone.addEventListener("click", () => {
  pdfInput.click();
});

pickPdfBtn.addEventListener("click", (event) => {
  event.stopPropagation();
  pdfInput.click();
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("drag");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("drag");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("drag");
  const files = event.dataTransfer.files;
  if (files && files.length) {
    const dataTransfer = new DataTransfer();
    Array.from(files).forEach((file) => dataTransfer.items.add(file));
    pdfInput.files = dataTransfer.files;
    updateFileList(dataTransfer.files);
  }
});

pdfInput.addEventListener("change", () => {
  updateFileList(pdfInput.files);
  setProgress(0);
});

mappingFile.addEventListener("change", async () => {
  const file = mappingFile.files[0];
  if (!file) {
    return;
  }
  try {
    const mapping = await parseMappingFile(file);
    setMapping(mapping);
    setStatus("Mapping geladen.");
  } catch (error) {
    setStatus(`Mapping Fehler: ${error.message}`);
  }
});

applyMappingBtn.addEventListener("click", () => {
  const mapping = collectMappingFromList();
  setMapping(mapping);
  setStatus("Mapping gespeichert.");
});

resetMappingBtn.addEventListener("click", () => {
  localStorage.removeItem(MAPPING_KEY);
  state.mapping = {};
  state.mappingKeys = [];
  renderMappingList({});
  setStatus("Mapping zur\u00fcckgesetzt.");
  scheduleRemoteSave();
});

addMappingBtn.addEventListener("click", () => {
  mappingList.appendChild(createMappingRow("", ""));
});

function handleHeaderChange() {
  setHeaderConfig({
    titleLine1: titleLine1Input.value.trim(),
    titleLine2: titleLine2Input.value.trim(),
    showGroupRow: groupRowToggle.checked
  });
  renderTable(state.output);
}

titleLine1Input.addEventListener("input", handleHeaderChange);
titleLine2Input.addEventListener("input", handleHeaderChange);
groupRowToggle.addEventListener("change", handleHeaderChange);

loadStoredMapping();
loadStoredHeader();
updateFileList(pdfInput.files);

if (auth && onAuthStateChanged) {
  onAuthStateChanged(auth, async (user) => {
    currentUser = user;
    if (!user) return;
    await loadRemoteState();
    loadStoredMapping();
    loadStoredHeader();
    updateFileList(pdfInput.files);
  });
}
