const defaultMappingRules = [
  { match: "PrimeTime Suite f\u00fcr 8+K\u00fcche/WIFI", id: "00632" },
  { match: "PrimeTime: Suite f\u00fcr 2 + K\u00fcche", id: "00632" },
  { match: "PrimeTimeSuite:K\u00fcche+WiFi+Modern", id: "00632" },
  { match: "PrimeTimeSute:K\u00fcche+WiFi+modern", id: "00632" },
  { match: "PrimeTimeSuites | N\u00e4he Messe&Hbf", id: "00631" },
  { match: "NewYork Suite f\u00fcr 4+K\u00fcche/WIFI", id: "00633" },
  { match: "Paris Suite f\u00fcr 4+K\u00fcche/WIFI", id: "00633" },
  { match: "London Suite f\u00fcr 4+K\u00fcche/WIFI", id: "00633" },
  { match: "Prime Art Gallery Apartment | Zentrale Lage Essen", id: "00634" },
  { match: "PrimeTime Suite f\u00fcr 4+K\u00fcche/WIFI/Parkplatz", id: "00635" },
  { match: "Modern & zentral | N\u00e4he Messe | K\u00fcche + WLAN", id: "00636" }
];

const defaultStreetMap = {
  "00631": "Opernplatz 15",
  "00632": "M\u00fclheimer Str. 15",
  "00633": "Haskenstr. 44",
  "00634": "M\u00fclheimer Str. 17",
  "00635": "Heibauerfeld 13",
  "00636": "Osnabr\u00fccker Str. 5"
};

const defaultColumns = [
  {
    key: "id",
    label: "Beherbergungsidentifikationsnummer",
    groupLabel: "",
    type: "text",
    enabled: true
  },
  { key: "street", label: "Strasse", groupLabel: "", type: "text", enabled: true },
  {
    key: "sp1",
    label: "SP1 Bruttoeink\u00fcnfte",
    groupLabel: "SP1",
    type: "number",
    formula: "brutto + adjustment_service",
    enabled: true
  },
  {
    key: "sp2",
    label: "SP2 Reinigungskosten",
    groupLabel: "SP2",
    type: "number",
    formula: "reinigung",
    enabled: true
  },
  {
    key: "sp3",
    label: "SP3 Anpassungen",
    groupLabel: "SP3",
    type: "number",
    formula: "anpassungen",
    enabled: true
  },
  {
    key: "sp4",
    label: "SP4 Zwischensumme (BmGI Beherbergungssteuer)",
    groupLabel: "SP4",
    type: "number",
    formula: "sp1 - sp2 + sp3",
    enabled: true
  },
  {
    key: "sp5",
    label: "SP5 Servicegeb\u00fchren",
    groupLabel: "SP5",
    type: "number",
    formula: "-service",
    enabled: true
  },
  {
    key: "sp6",
    label: "SP6 Auszahlung n. Abz\u00fcgen SP 1 - SP3 - SP5",
    groupLabel: "SP6",
    type: "number",
    formula: "payout",
    enabled: true
  },
  {
    key: "sp7",
    label: "SP7 Beherbergungssteuer 5% von SP4",
    groupLabel: "SP7",
    type: "number",
    formula: "sp4 * 0.05",
    enabled: true
  }
];

const defaultHeaderConfig = {
  titleLine1: "Airbnb Auszahlungsuebersicht",
  titleLine2: "Dezember 2025",
  showGroupRow: true
};

const STORAGE_KEYS = {
  rules: "airbnbMappingRules",
  streets: "airbnbStreetMap",
  columns: "airbnbColumns",
  header: "airbnbHeaderConfig"
};

const state = {
  rows: [],
  output: { columns: [], rows: [] },
  mappings: {
    rules: [],
    streets: {}
  },
  columns: [],
  header: { ...defaultHeaderConfig }
};

const csvInput = document.getElementById("csvFile");
const dropZone = document.getElementById("dropZone");
const pickFileBtn = document.getElementById("pickFileBtn");
const fileNameEl = document.getElementById("fileName");
const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");

const previewBtn = document.getElementById("previewBtn");
const downloadBtn = document.getElementById("downloadBtn");
const statusEl = document.getElementById("status");
const previewTable = document.getElementById("previewTable");
const tableHead = document.querySelector("#previewTable thead");
const tableBody = document.querySelector("#previewTable tbody");

const openMappingBtn = document.getElementById("openMappingBtn");
const resetMappingBtn = document.getElementById("resetMappingBtn");
const mappingOverlay = document.getElementById("mappingOverlay");
const closeOverlayBtn = document.getElementById("closeOverlayBtn");
const cancelMappingBtn = document.getElementById("cancelMappingBtn");
const saveMappingBtn = document.getElementById("saveMappingBtn");
const addRuleBtn = document.getElementById("addRuleBtn");
const addStreetBtn = document.getElementById("addStreetBtn");
const addColumnBtn = document.getElementById("addColumnBtn");
const titleLine1Input = document.getElementById("titleLine1");
const titleLine2Input = document.getElementById("titleLine2");
const groupRowToggle = document.getElementById("groupRowToggle");
const ruleList = document.getElementById("ruleList");
const streetList = document.getElementById("streetList");
const columnList = document.getElementById("columnList");

function setStatus(message, tone = "") {
  statusEl.textContent = message;
  statusEl.dataset.tone = tone;
}

function setProgress(percent) {
  const safe = Math.max(0, Math.min(100, percent));
  progressBar.style.width = `${safe}%`;
  progressText.textContent = `${safe}%`;
}

function updateFileName(file) {
  if (!file) {
    fileNameEl.textContent = "Keine Datei ausgewaehlt";
    return;
  }
  fileNameEl.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
}

function setFile(file) {
  if (!file) {
    return;
  }
  const dataTransfer = new DataTransfer();
  dataTransfer.items.add(file);
  csvInput.files = dataTransfer.files;
  updateFileName(file);
  setProgress(0);
}

function loadStoredMappings() {
  let rules = defaultMappingRules;
  let streets = defaultStreetMap;
  let columns = defaultColumns;
  let header = defaultHeaderConfig;

  try {
    const storedRules = JSON.parse(localStorage.getItem(STORAGE_KEYS.rules));
    if (Array.isArray(storedRules) && storedRules.length) {
      rules = storedRules;
    }
  } catch (error) {
    rules = defaultMappingRules;
  }

  try {
    const storedStreets = JSON.parse(localStorage.getItem(STORAGE_KEYS.streets));
    if (storedStreets && typeof storedStreets === "object") {
      streets = storedStreets;
    }
  } catch (error) {
    streets = defaultStreetMap;
  }

  try {
    const storedColumns = JSON.parse(localStorage.getItem(STORAGE_KEYS.columns));
    if (Array.isArray(storedColumns) && storedColumns.length) {
      columns = storedColumns;
    }
  } catch (error) {
    columns = defaultColumns;
  }

  try {
    const storedHeader = JSON.parse(localStorage.getItem(STORAGE_KEYS.header));
    if (storedHeader && typeof storedHeader === "object") {
      header = { ...defaultHeaderConfig, ...storedHeader };
    }
  } catch (error) {
    header = defaultHeaderConfig;
  }

  state.mappings = {
    rules: rules.map((rule) => ({ match: rule.match || "", id: rule.id || "" })),
    streets: { ...streets }
  };
  state.columns = sanitizeColumns(columns);
  state.header = { ...defaultHeaderConfig, ...header };
}

function saveStoredMappings() {
  localStorage.setItem(STORAGE_KEYS.rules, JSON.stringify(state.mappings.rules));
  localStorage.setItem(STORAGE_KEYS.streets, JSON.stringify(state.mappings.streets));
  localStorage.setItem(STORAGE_KEYS.columns, JSON.stringify(state.columns));
  localStorage.setItem(STORAGE_KEYS.header, JSON.stringify(state.header));
}

function sanitizeColumns(columns) {
  const defaultsByKey = new Map(defaultColumns.map((column) => [column.key, column]));
  const result = [];

  if (Array.isArray(columns)) {
    columns.forEach((column) => {
      if (!column || !column.key) {
        return;
      }
      const base = defaultsByKey.get(column.key);
      if (base) {
        const merged = {
          ...base,
          ...column,
          key: base.key,
          type: base.type,
          groupLabel: column.groupLabel || base.groupLabel || ""
        };
        if (merged.key === "sp1" && merged.formula === "brutto") {
          merged.formula = "brutto + adjustment_service";
        }
        if (merged.key === "sp6" && merged.formula === "sp1 + sp3 + sp5") {
          merged.formula = "payout";
        }
        result.push(merged);
      } else {
        result.push({ ...column, groupLabel: column.groupLabel || "" });
      }
      defaultsByKey.delete(column.key);
    });
  }

  defaultsByKey.forEach((value) => {
    result.push({ ...value });
  });

  return result;
}

function parseCsv(text) {
  const rows = [];
  const cleanedText = String(text || "").replace(/^\uFEFF/, "");
  const delimiter = detectDelimiter(cleanedText);
  let current = "";
  let inQuotes = false;
  let row = [];

  for (let i = 0; i < cleanedText.length; i += 1) {
    const char = cleanedText[i];
    const next = cleanedText[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      row.push(current);
      current = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(current);
      current = "";
      if (row.length > 1 || row[0] !== "") {
        rows.push(row);
      }
      row = [];
    } else {
      current += char;
    }
  }

  if (current.length || row.length) {
    row.push(current);
    rows.push(row);
  }

  if (!rows.length) {
    return { headers: [], rows: [] };
  }

  const headers = rows[0].map((header) => header.trim());
  const dataRows = rows.slice(1).map((data) => {
    const entry = {};
    headers.forEach((header, index) => {
      entry[header] = data[index] !== undefined ? data[index] : "";
    });
    return entry;
  });

  return { headers, rows: dataRows };
}

function parseCurrency(value) {
  if (value === null || value === undefined) {
    return 0;
  }
  const raw = String(value)
    .replace(/[^0-9,.-]/g, "")
    .replace(/\s+/g, "")
    .trim();

  if (!raw) {
    return 0;
  }

  let normalized = raw;
  const hasComma = raw.includes(",");
  const hasDot = raw.includes(".");

  if (hasComma && hasDot) {
    normalized = raw.replace(/\./g, "").replace(/,/g, ".");
  } else if (hasComma && !hasDot) {
    normalized = raw.replace(/,/g, ".");
  }

  const number = Number.parseFloat(normalized);
  return Number.isNaN(number) ? 0 : number;
}

function detectDelimiter(text) {
  let comma = 0;
  let semicolon = 0;
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (char === '"') {
      inQuotes = !inQuotes;
    }
    if (!inQuotes) {
      if (char === ",") {
        comma += 1;
      }
      if (char === ";") {
        semicolon += 1;
      }
      if (char === "\n") {
        break;
      }
    }
  }

  if (semicolon > comma) {
    return ";";
  }
  return ",";
}

function normalizeHeader(header) {
  return String(header || "")
    .toLowerCase()
    .replace(/\u00e4/g, "ae")
    .replace(/\u00f6/g, "oe")
    .replace(/\u00fc/g, "ue")
    .replace(/\u00df/g, "ss")
    .replace(/[^a-z0-9]/g, "");
}

function standardizeRows(rows, headers) {
  const normalizedMap = {};
  headers.forEach((header) => {
    normalizedMap[normalizeHeader(header)] = header;
  });

  const expected = {
    typ: "Typ",
    inserat: "Inserat",
    bruttoeinkuenfte: "Bruttoeink\u00fcnfte",
    betrag: "Betrag",
    reinigungsgebuehr: "Reinigungsgeb\u00fchr",
    servicegebuehr: "Servicegeb\u00fchr"
  };

  const missing = Object.keys(expected).filter((key) => !normalizedMap[key]);
  if (missing.length) {
    throw new Error(`CSV-Header fehlt: ${missing.map((key) => expected[key]).join(", ")}`);
  }

  return rows.map((row) => ({
    Typ: row[normalizedMap.typ] || "",
    Inserat: row[normalizedMap.inserat] || "",
    "Bruttoeink\u00fcnfte": row[normalizedMap.bruttoeinkuenfte] || "",
    Betrag: row[normalizedMap.betrag] || "",
    "Reinigungsgeb\u00fchr": row[normalizedMap.reinigungsgebuehr] || "",
    "Servicegeb\u00fchr": row[normalizedMap.servicegebuehr] || ""
  }));
}

function findId(inserat, rules) {
  const text = String(inserat || "").toLowerCase();
  for (const rule of rules) {
    if (!rule || !rule.match) {
      continue;
    }
    const match = String(rule.match).toLowerCase();
    if (match && text.includes(match)) {
      return String(rule.id || "").trim() || "ID-BITTE_ANPASSEN";
    }
  }
  return "ID-BITTE_ANPASSEN";
}

function evaluateFormula(formula, context) {
  if (!formula) {
    return 0;
  }
  const normalized = String(formula).replace(/,/g, ".").trim();
  if (!/^[0-9+\-*/().\sA-Za-z_]+$/.test(normalized)) {
    throw new Error(`Ungueltige Formel: ${formula}`);
  }

  const identifiers = normalized.match(/[A-Za-z_][A-Za-z0-9_]*/g) || [];
  let expression = normalized;

  identifiers.forEach((identifier) => {
    if (!(identifier in context)) {
      throw new Error(`Unbekannte Variable: ${identifier}`);
    }
    const value = Number(context[identifier] || 0);
    expression = expression.replace(new RegExp(`\\b${identifier}\\b`, "g"), `${value}`);
  });

  const result = Function(`"use strict";return (${expression});`)();
  return Number.isFinite(result) ? result : 0;
}

function buildOutput(rows, rules, streets, columns) {
  const groups = new Map();

  rows.forEach((row) => {
    const typ = String(row.Typ || "").trim();
    if (!typ || typ === "Payout") {
      return;
    }

    const inserat = row.Inserat || "";
    const id = findId(inserat, rules);

    if (!groups.has(id)) {
      groups.set(id, {
        id,
        bruttos: 0,
        reinigung: 0,
        anpassungen: 0,
        servicePositiv: 0,
        adjustmentService: 0,
        payout: 0
      });
    }

    const entry = groups.get(id);
    const brutto = parseCurrency(row["Bruttoeink\u00fcnfte"]);
    const betrag = parseCurrency(row["Betrag"]);
    const reinigung = parseCurrency(row["Reinigungsgeb\u00fchr"]);
    const service = parseCurrency(row["Servicegeb\u00fchr"]);

    if (typ !== "Payout") {
      entry.payout += betrag;
    }

    if (typ === "Mediations-Anpassung") {
      entry.anpassungen += brutto;
      return;
    }

    if (typ === "Anpassung") {
      entry.anpassungen += betrag;
      entry.adjustmentService += service;
      return;
    }

    if (typ === "Durchlaufposten gesamt" || typ === "Buchung") {
      entry.bruttos += brutto;
      entry.reinigung += reinigung;
      entry.servicePositiv += service;
    }
  });

  const activeColumns = columns.filter((column) => column.enabled !== false);
  const totals = {};
  activeColumns.forEach((column) => {
    totals[column.key] = 0;
  });

  const outputRows = [];

  groups.forEach((entry) => {
    const context = {
      brutto: entry.bruttos,
      reinigung: entry.reinigung,
      anpassungen: entry.anpassungen,
      service: entry.servicePositiv,
      adjustment_service: entry.adjustmentService,
      payout: entry.payout
    };

    const row = {};

    activeColumns.forEach((column) => {
      let value = "";
      if (column.key === "id") {
        value = entry.id;
      } else if (column.key === "street") {
        value = streets[entry.id] || "";
      } else if (column.type === "number") {
        value = evaluateFormula(column.formula, context);
      }

      row[column.key] = value;
      context[column.key] = value;

      if (column.type === "number") {
        totals[column.key] += Number(value || 0);
      }
    });

    outputRows.push(row);
  });

  const totalRow = {};
  activeColumns.forEach((column) => {
    if (column.key === "id") {
      totalRow[column.key] = "GESAMT";
    } else if (column.type === "number") {
      totalRow[column.key] = totals[column.key] || 0;
    } else {
      totalRow[column.key] = "";
    }
  });

  outputRows.push(totalRow);

  return { columns: activeColumns, rows: outputRows };
}

function formatNumber(value) {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return "0.00";
  }
  return value.toLocaleString("de-DE", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function formatCurrency(value) {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return "€ 0,00";
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

  const widths = computeColumnWidths(output.columns);
  const colGroup = `<colgroup>${widths
    .map((width) => `<col style="width:${width}ch">`)
    .join("")}</colgroup>`;
  const existingColGroup = previewTable.querySelector("colgroup");
  if (existingColGroup) {
    existingColGroup.outerHTML = colGroup;
  } else {
    previewTable.insertAdjacentHTML("afterbegin", colGroup);
  }

  const columnCount = output.columns.length;
  const titleLine1 = state.header.titleLine1 || "";
  const titleLine2 = state.header.titleLine2 || "";
  const titleRows = [
    `<tr><th colspan="${columnCount}" class="title-row">${titleLine1}</th></tr>`,
    `<tr><th colspan="${columnCount}" class="title-row">${titleLine2}</th></tr>`
  ];

  const groupRow = state.header.showGroupRow
    ? buildGroupHeaderRow(output.columns)
    : "";

  const headerRow = `<tr>${output.columns
    .map((column) => `<th>${column.label}</th>`)
    .join("")}</tr>`;

  tableHead.innerHTML = `${titleRows.join("")}${groupRow}${headerRow}`;

  tableBody.innerHTML = output.rows
    .map((row, rowIndex) => {
      const isTotal = rowIndex === output.rows.length - 1;
      const cells = output.columns.map((column) => {
        const value = row[column.key];
        const display = column.type === "number" ? formatCurrency(value) : value;
        const editable = !isTotal;
        return (
          `<td data-row-index="${rowIndex}" data-col-key="${column.key}" ` +
          `contenteditable="${editable}">${display}</td>`
        );
      });
      return `<tr>${cells.join("")}</tr>`;
    })
    .join("");
}

function computeColumnWidths(columns) {
  return columns.map((column) => {
    const label = column.label || "";
    const base = Math.max(label.length, 6);
    return Math.min(base + 2, 60);
  });
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
    const safeLabel = label || "";
    cells.push(`<th colspan="${span}" class="group-row">${safeLabel}</th>`);
    index += span;
  }

  return `<tr>${cells.join("")}</tr>`;
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
  const showGroupRow = state.header.showGroupRow !== false;

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
    `<numFmt numFmtId="164" formatCode="[$€-de-DE] #,##0.00"/>\n` +
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
  link.download = "airbnb_auswertung.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function clearList(container) {
  while (container.firstChild) {
    container.removeChild(container.firstChild);
  }
}

function createDragHandle() {
  const handle = document.createElement("span");
  handle.className = "drag-handle";
  handle.textContent = "::";
  handle.setAttribute("draggable", "true");
  return handle;
}

function createRuleRow(rule) {
  const row = document.createElement("div");
  row.className = "list-row";
  row.dataset.type = "rule";

  const handle = createDragHandle();
  const matchInput = document.createElement("input");
  matchInput.type = "text";
  matchInput.placeholder = "Text im Inserat";
  matchInput.value = rule.match || "";

  const idInput = document.createElement("input");
  idInput.type = "text";
  idInput.placeholder = "Beherbergungs-ID";
  idInput.value = rule.id || "";

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.textContent = "Entfernen";
  removeBtn.addEventListener("click", () => {
    row.remove();
  });

  row.appendChild(handle);
  row.appendChild(matchInput);
  row.appendChild(idInput);
  row.appendChild(removeBtn);

  return row;
}

function createStreetRow(entry) {
  const row = document.createElement("div");
  row.className = "list-row";
  row.dataset.type = "street";

  const handle = createDragHandle();
  const idInput = document.createElement("input");
  idInput.type = "text";
  idInput.placeholder = "Beherbergungs-ID";
  idInput.value = entry.id || "";

  const streetInput = document.createElement("input");
  streetInput.type = "text";
  streetInput.placeholder = "Strasse";
  streetInput.value = entry.street || "";

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.textContent = "Entfernen";
  removeBtn.addEventListener("click", () => {
    row.remove();
  });

  row.appendChild(handle);
  row.appendChild(idInput);
  row.appendChild(streetInput);
  row.appendChild(removeBtn);

  return row;
}

function createColumnRow(column) {
  const row = document.createElement("div");
  row.className = "list-row columns";
  row.dataset.key = column.key;
  row.dataset.type = column.type;

  const handle = createDragHandle();

  const keyInput = document.createElement("input");
  keyInput.type = "text";
  keyInput.placeholder = "key";
  keyInput.value = column.key || "";
  keyInput.className = "col-key";
  keyInput.addEventListener("input", () => {
    row.dataset.key = keyInput.value.trim();
  });

  const labelInput = document.createElement("input");
  labelInput.type = "text";
  labelInput.placeholder = "Spaltenname";
  labelInput.value = column.label || column.key;
  labelInput.className = "col-label";

  const groupInput = document.createElement("input");
  groupInput.type = "text";
  groupInput.placeholder = "Gruppenlabel";
  groupInput.value = column.groupLabel || "";
  groupInput.className = "col-group";

  const formulaInput = document.createElement("input");
  formulaInput.type = "text";
  formulaInput.placeholder = column.type === "number" ? "Formel" : "Kein Formelwert";
  formulaInput.value = column.formula || "";
  formulaInput.className = "col-formula";
  if (column.type !== "number") {
    formulaInput.disabled = true;
  }

  const typeSelect = document.createElement("select");
  typeSelect.className = "col-type";
  [
    { value: "number", label: "Zahl" },
    { value: "text", label: "Text" }
  ].forEach((option) => {
    const item = document.createElement("option");
    item.value = option.value;
    item.textContent = option.label;
    if (option.value === column.type) {
      item.selected = true;
    }
    typeSelect.appendChild(item);
  });
  typeSelect.addEventListener("change", () => {
    row.dataset.type = typeSelect.value;
    if (typeSelect.value === "number") {
      formulaInput.disabled = false;
    } else {
      formulaInput.disabled = true;
      formulaInput.value = "";
    }
  });

  const toggle = document.createElement("label");
  toggle.className = "toggle";
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.className = "col-enabled";
  checkbox.checked = column.enabled !== false;
  const toggleText = document.createElement("span");
  toggleText.textContent = "Anzeigen";
  toggle.appendChild(checkbox);
  toggle.appendChild(toggleText);

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "remove-btn";
  removeBtn.textContent = "Entfernen";
  removeBtn.addEventListener("click", () => {
    row.remove();
  });

  row.appendChild(handle);
  row.appendChild(keyInput);
  row.appendChild(labelInput);
  row.appendChild(groupInput);
  row.appendChild(formulaInput);
  row.appendChild(typeSelect);
  row.appendChild(toggle);
  row.appendChild(removeBtn);

  return row;
}

function enableDragList(container) {
  let dragging = null;

  container.addEventListener("dragstart", (event) => {
    const handle = event.target.closest(".drag-handle");
    if (!handle) {
      event.preventDefault();
      return;
    }
    const row = handle.closest(".list-row");
    if (!row) {
      return;
    }
    dragging = row;
    row.classList.add("dragging");
    event.dataTransfer.effectAllowed = "move";
  });

  container.addEventListener("dragover", (event) => {
    event.preventDefault();
    if (!dragging) {
      return;
    }
    const target = event.target.closest(".list-row");
    if (!target || target === dragging) {
      return;
    }
    const rect = target.getBoundingClientRect();
    const after = event.clientY - rect.top > rect.height / 2;
    container.insertBefore(dragging, after ? target.nextSibling : target);
  });

  container.addEventListener("dragend", () => {
    if (dragging) {
      dragging.classList.remove("dragging");
    }
    dragging = null;
  });
}

function renderOverlay() {
  clearList(ruleList);
  clearList(streetList);
  clearList(columnList);

  titleLine1Input.value = state.header.titleLine1 || "";
  titleLine2Input.value = state.header.titleLine2 || "";
  groupRowToggle.checked = state.header.showGroupRow !== false;

  state.mappings.rules.forEach((rule) => {
    ruleList.appendChild(createRuleRow(rule));
  });

  const streetEntries = Object.entries(state.mappings.streets).map(([id, street]) => ({
    id,
    street
  }));

  streetEntries.forEach((entry) => {
    streetList.appendChild(createStreetRow(entry));
  });

  state.columns.forEach((column) => {
    columnList.appendChild(createColumnRow(column));
  });
}

function openOverlay() {
  renderOverlay();
  mappingOverlay.classList.add("show");
  mappingOverlay.setAttribute("aria-hidden", "false");
}

function closeOverlay() {
  mappingOverlay.classList.remove("show");
  mappingOverlay.setAttribute("aria-hidden", "true");
}

function collectOverlayMappings() {
  const rules = [];
  const streets = {};
  const columns = [];

  Array.from(ruleList.children).forEach((row) => {
    const inputs = row.querySelectorAll("input");
    const match = inputs[0] ? inputs[0].value.trim() : "";
    const id = inputs[1] ? inputs[1].value.trim() : "";
    if (match && id) {
      rules.push({ match, id });
    }
  });

  Array.from(streetList.children).forEach((row) => {
    const inputs = row.querySelectorAll("input");
    const id = inputs[0] ? inputs[0].value.trim() : "";
    const street = inputs[1] ? inputs[1].value.trim() : "";
    if (id) {
      streets[id] = street;
    }
  });

  Array.from(columnList.children).forEach((row) => {
    const key = row.querySelector(".col-key").value.trim();
    const type = row.querySelector(".col-type").value;
    const label = row.querySelector(".col-label").value.trim() || key;
    const groupLabel = row.querySelector(".col-group").value.trim();
    const formulaInput = row.querySelector(".col-formula");
    const enabled = row.querySelector(".col-enabled").checked;
    columns.push({
      key,
      type,
      label,
      groupLabel,
      formula: formulaInput ? formulaInput.value.trim() : "",
      enabled
    });
  });

  return {
    rules,
    streets,
    columns,
    header: {
      titleLine1: titleLine1Input.value.trim(),
      titleLine2: titleLine2Input.value.trim(),
      showGroupRow: groupRowToggle.checked
    }
  };
}

function validateColumns(columns) {
  const enabled = columns.filter((column) => column.enabled !== false);
  if (!enabled.length) {
    throw new Error("Bitte mindestens eine Spalte aktivieren.");
  }
  const seenKeys = new Set();
  const context = {
    brutto: 1,
    reinigung: 1,
    anpassungen: 1,
    service: 1,
    adjustment_service: 1,
    payout: 1
  };
  columns.forEach((column) => {
    if (!column.key) {
      throw new Error("Jede Spalte braucht einen eindeutigen key.");
    }
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(column.key)) {
      throw new Error(`Ungueltiger key: ${column.key}`);
    }
    if (seenKeys.has(column.key)) {
      throw new Error(`Key doppelt vorhanden: ${column.key}`);
    }
    seenKeys.add(column.key);
    if (column.type === "number" && column.enabled !== false && column.formula) {
      const tempContext = { ...context };
      columns.forEach((col) => {
        if (col.key && !(col.key in tempContext)) {
          tempContext[col.key] = 1;
        }
      });
      evaluateFormula(column.formula, tempContext);
    }
  });
}

function updateTotals(output) {
  if (!output.rows.length) {
    return;
  }
  const totalIndex = output.rows.length - 1;
  const totals = {};
  output.columns.forEach((column) => {
    if (column.type === "number") {
      totals[column.key] = 0;
    }
  });

  output.rows.slice(0, totalIndex).forEach((row) => {
    output.columns.forEach((column) => {
      if (column.type === "number") {
        totals[column.key] += Number(row[column.key] || 0);
      }
    });
  });

  output.columns.forEach((column) => {
    if (column.key === "id") {
      output.rows[totalIndex][column.key] = "GESAMT";
    } else if (column.type === "number") {
      output.rows[totalIndex][column.key] = totals[column.key] || 0;
    } else {
      output.rows[totalIndex][column.key] = "";
    }
  });
}

function nextCustomKey(existingKeys) {
  let index = 1;
  while (existingKeys.has(`custom_${index}`)) {
    index += 1;
  }
  return `custom_${index}`;
}

previewBtn.addEventListener("click", () => {
  if (!csvInput.files.length) {
    setStatus("Bitte zuerst eine CSV-Datei auswaehlen.");
    return;
  }

  const file = csvInput.files[0];
  const reader = new FileReader();
  setProgress(0);
  reader.onprogress = (event) => {
    if (event.lengthComputable) {
      setProgress(Math.round((event.loaded / event.total) * 100));
    }
  };
  reader.onload = () => {
    try {
      const parsed = parseCsv(reader.result);
      const standardized = standardizeRows(parsed.rows, parsed.headers);
      state.rows = standardized;
      state.output = buildOutput(
        standardized,
        state.mappings.rules,
        state.mappings.streets,
        state.columns
      );
      renderTable(state.output);
      downloadBtn.disabled = state.output.rows.length === 0;
      setProgress(100);
      setStatus(`Fertig. ${state.output.rows.length} Zeilen erzeugt.`);
    } catch (error) {
      setStatus(`Fehler beim Verarbeiten: ${error.message}`);
    }
  };
  reader.onerror = () => {
    setStatus("Datei konnte nicht gelesen werden.");
  };

  reader.readAsText(file, "utf-8");
});

downloadBtn.addEventListener("click", () => {
  if (!state.output.rows.length) {
    setStatus("Bitte zuerst eine Vorschau erzeugen.");
    return;
  }
  downloadExcel(state.output);
});

tableBody.addEventListener(
  "blur",
  (event) => {
    const cell = event.target.closest("td");
    if (!cell || cell.getAttribute("contenteditable") !== "true") {
      return;
    }
    const rowIndex = Number(cell.dataset.rowIndex);
    const key = cell.dataset.colKey;
    if (!Number.isFinite(rowIndex) || !key) {
      return;
    }
    if (rowIndex >= state.output.rows.length - 1) {
      return;
    }
    const column = state.output.columns.find((col) => col.key === key);
    if (!column) {
      return;
    }
    const raw = cell.textContent.trim();
    const value = column.type === "number" ? parseCurrency(raw) : raw;
    state.output.rows[rowIndex][key] = value;
    updateTotals(state.output);
    renderTable(state.output);
  },
  true
);

openMappingBtn.addEventListener("click", () => {
  openOverlay();
});

closeOverlayBtn.addEventListener("click", () => {
  closeOverlay();
});

cancelMappingBtn.addEventListener("click", () => {
  closeOverlay();
});

saveMappingBtn.addEventListener("click", () => {
  try {
    const updated = collectOverlayMappings();
    if (!updated.rules.length) {
      setStatus("Bitte mindestens eine Regel speichern.");
      return;
    }
    validateColumns(updated.columns);
    state.mappings = { rules: updated.rules, streets: updated.streets };
    state.columns = sanitizeColumns(updated.columns);
    state.header = { ...defaultHeaderConfig, ...updated.header };
    saveStoredMappings();
    setStatus("Konfiguration gespeichert.");
    closeOverlay();
  } catch (error) {
    setStatus(`Fehler: ${error.message}`);
  }
});

addRuleBtn.addEventListener("click", () => {
  ruleList.appendChild(createRuleRow({ match: "", id: "" }));
});

addStreetBtn.addEventListener("click", () => {
  streetList.appendChild(createStreetRow({ id: "", street: "" }));
});

addColumnBtn.addEventListener("click", () => {
  const existingKeys = new Set(
    Array.from(columnList.children).map((row) => row.dataset.key).filter(Boolean)
  );
  const key = nextCustomKey(existingKeys);
  columnList.appendChild(
    createColumnRow({
      key,
      label: "Neue Spalte",
      groupLabel: "",
      type: "number",
      formula: "0",
      enabled: true
    })
  );
});

resetMappingBtn.addEventListener("click", () => {
  state.mappings = {
    rules: defaultMappingRules.map((rule) => ({ ...rule })),
    streets: { ...defaultStreetMap }
  };
  state.columns = defaultColumns.map((column) => ({ ...column }));
  state.header = { ...defaultHeaderConfig };
  saveStoredMappings();
  setStatus("Konfiguration zurueckgesetzt.");
});

mappingOverlay.addEventListener("click", (event) => {
  if (event.target === mappingOverlay) {
    closeOverlay();
  }
});

dropZone.addEventListener("click", () => {
  csvInput.click();
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
  const file = event.dataTransfer.files[0];
  if (file) {
    setFile(file);
  }
});

csvInput.addEventListener("change", () => {
  const file = csvInput.files[0];
  updateFileName(file);
  setProgress(0);
});

pickFileBtn.addEventListener("click", (event) => {
  event.stopPropagation();
  csvInput.click();
});

enableDragList(ruleList);
enableDragList(streetList);
enableDragList(columnList);

loadStoredMappings();

