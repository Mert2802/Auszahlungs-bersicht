const DEFAULT_SETTINGS = {
  tolerance: 1.0,
  windowStart: 1,
  windowEnd: 5,
  allowPartial: true,
  accounts: [
    {
      id: "konto-1",
      name: "Hauptkonto",
      iban: ""
    }
  ],
  tenants: [
    {
      name: "Max Mustermann",
      unit: "Wohnung 1",
      rent: 750.0,
      dueDay: 3,
      iban: ""
    }
  ]
};

const STORAGE_KEY = "miete_settings_v1";
const CSV_CACHE_KEY = "miete_csv_cache_v1";
const MONTH_KEY = "miete_month_v1";
const ACCOUNT_KEY = "miete_account_v1";
const DASHBOARD_KEY = "dashboard_snapshots_v1";
const BACKUP_DB = "miete_backup_db_v1";
const BACKUP_STORE = "handles";
const BACKUP_KEY = "backupFile";

const els = {
  csvInput: document.getElementById("csvInput"),
  monthSelect: document.getElementById("monthSelect"),
  accountSelect: document.getElementById("accountSelect"),
  tenantTable: document.querySelector("#tenantTable tbody"),
  unmatchedTable: document.querySelector("#unmatchedTable tbody"),
  statIncome: document.getElementById("statIncome"),
  statExpected: document.getElementById("statExpected"),
  statMissing: document.getElementById("statMissing"),
  statPartial: document.getElementById("statPartial"),
  exportEurBtn: document.getElementById("exportEurBtn"),
  exportEurXlsBtn: document.getElementById("exportEurXlsBtn"),
  clearImportsBtn: document.getElementById("clearImportsBtn"),
  snapshotBtn: document.getElementById("snapshotBtn"),
  snapshotStatus: document.getElementById("snapshotStatus"),
  settingsOverlay: document.getElementById("settingsOverlay"),
  openSettingsBtn: document.getElementById("openSettingsBtn"),
  closeSettingsBtn: document.getElementById("closeSettingsBtn"),
  closeSettingsBtnFooter: document.getElementById("closeSettingsBtnFooter"),
  saveSettingsBtn: document.getElementById("saveSettingsBtn"),
  addTenantBtn: document.getElementById("addTenantBtn"),
  addAccountBtn: document.getElementById("addAccountBtn"),
  settingsTable: document.querySelector("#settingsTable tbody"),
  accountsTable: document.querySelector("#accountsTable tbody"),
  toleranceInput: document.getElementById("toleranceInput"),
  windowStartInput: document.getElementById("windowStartInput"),
  windowEndInput: document.getElementById("windowEndInput"),
  allowPartialInput: document.getElementById("allowPartialInput"),
  backupStatus: document.getElementById("backupStatus"),
  backupSelectBtn: document.getElementById("backupSelectBtn"),
  backupNowBtn: document.getElementById("backupNowBtn"),
  authOverlay: document.getElementById("authOverlay"),
  loginUser: document.getElementById("loginUser"),
  loginPass: document.getElementById("loginPass"),
  loginBtn: document.getElementById("loginBtn"),
  loginError: document.getElementById("loginError")
};

let settings = loadSettings();
let transactions = [];
let currentUser = null;
let currentProfile = null;
let remoteLoaded = false;

const firebase = window.firebaseServices || {};
const {
  auth,
  db,
  onAuthStateChanged,
  signInWithEmailAndPassword,
  signOut,
  doc,
  getDoc,
  setDoc,
  serverTimestamp
} = firebase;

function loadSettings() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return structuredClone(DEFAULT_SETTINGS);
  try {
    const parsed = JSON.parse(raw);
    const accounts = Array.isArray(parsed.accounts)
      ? parsed.accounts
      : DEFAULT_SETTINGS.accounts;
    return {
      ...DEFAULT_SETTINGS,
      ...parsed,
      accounts: accounts.length ? accounts : DEFAULT_SETTINGS.accounts,
      tenants: Array.isArray(parsed.tenants) ? parsed.tenants : []
    };
  } catch (error) {
    return structuredClone(DEFAULT_SETTINGS);
  }
}

function saveSettings(next) {
  settings = next;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

function openBackupDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(BACKUP_DB, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(BACKUP_STORE);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveBackupHandle(handle) {
  const db = await openBackupDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(BACKUP_STORE, "readwrite");
    tx.objectStore(BACKUP_STORE).put(handle, BACKUP_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadBackupHandle() {
  const db = await openBackupDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(BACKUP_STORE, "readonly");
    const request = tx.objectStore(BACKUP_STORE).get(BACKUP_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

function serializeTransactions(items) {
  return items.map((entry) => ({
    ...entry,
    date: entry.date ? entry.date.toISOString() : null,
    bookingDate: entry.bookingDate ? entry.bookingDate.toISOString() : null
  }));
}

function deserializeTransactions(items) {
  if (!Array.isArray(items)) return [];
  return items.map((entry) => ({
    ...entry,
    date: entry.date ? new Date(entry.date) : null,
    bookingDate: entry.bookingDate ? new Date(entry.bookingDate) : null
  }));
}

function buildBackupPayload() {
  return {
    savedAt: new Date().toISOString(),
    settings,
    month: els.monthSelect ? els.monthSelect.value : "",
    account: els.accountSelect ? els.accountSelect.value : "",
    transactions: serializeTransactions(transactions)
  };
}

async function writeBackup(handle) {
  if (!handle) return false;
  const payload = JSON.stringify(buildBackupPayload(), null, 2);
  const writable = await handle.createWritable();
  await writable.write(payload);
  await writable.close();
  return true;
}

async function readBackup(handle) {
  if (!handle) return null;
  const file = await handle.getFile();
  const text = await file.text();
  return JSON.parse(text);
}

async function updateBackupStatus(handle) {
  if (!els.backupStatus) return;
  if (!handle) {
    els.backupStatus.textContent = "Kein Speicherort gew\u00e4hlt.";
    return;
  }
  els.backupStatus.textContent = "Auto-Backup aktiv.";
}

function toEmail(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  return raw.includes("@") ? raw : `${raw}@cockpit.local`;
}

function usernameFromEmail(email) {
  const raw = String(email || "");
  return raw.includes("@") ? raw.split("@")[0] : raw;
}

async function ensureUserProfile(user) {
  if (!db || !user) return null;
  const ref = doc(db, "users", user.uid);
  const snapshot = await getDoc(ref);
  if (snapshot.exists()) {
    return snapshot.data();
  }
  const username = usernameFromEmail(user.email) || "user";
  const role = username === "admin" ? "admin" : "user";
  const profile = {
    username,
    role,
    createdAt: serverTimestamp()
  };
  await setDoc(ref, profile, { merge: true });
  return { ...profile };
}

function buildRemotePayload() {
  return {
    settings,
    month: els.monthSelect ? els.monthSelect.value : "",
    account: els.accountSelect ? els.accountSelect.value : "",
    transactions: serializeTransactions(transactions)
  };
}

async function saveRemoteData() {
  if (!currentUser || !db) return;
  const ref = doc(db, "users", currentUser.uid, "systems", "miete");
  await setDoc(
    ref,
    { payload: buildRemotePayload(), updatedAt: serverTimestamp() },
    { merge: true }
  );
}

let saveTimer = null;
function scheduleRemoteSave() {
  if (!currentUser || !db) return;
  if (saveTimer) {
    clearTimeout(saveTimer);
  }
  saveTimer = setTimeout(() => {
    saveRemoteData().catch(() => {});
  }, 400);
}

async function loadRemoteData() {
  if (!currentUser || !db) return;
  const ref = doc(db, "users", currentUser.uid, "systems", "miete");
  const snapshot = await getDoc(ref);
  const data = snapshot.exists() ? snapshot.data() : null;
  if (data && data.payload) {
    applyPayload(data.payload);
    remoteLoaded = true;
  }
}

async function appendDashboardEntry(metrics) {
  if (!currentUser || !db) return;
  const ref = doc(db, "users", currentUser.uid, "dashboard", "summary");
  const snapshot = await getDoc(ref);
  const data = snapshot.exists() ? snapshot.data() : null;
  const history = Array.isArray(data && data.snapshots) ? data.snapshots : [];
  history.push({
    system: "miete",
    ts: Date.now(),
    metrics
  });
  await setDoc(
    ref,
    { snapshots: history.slice(-200), updatedAt: serverTimestamp() },
    { merge: true }
  );
}

function showAuthOverlay() {
  if (!els.authOverlay) return;
  els.authOverlay.classList.add("show");
  els.authOverlay.setAttribute("aria-hidden", "false");
}

function hideAuthOverlay() {
  if (!els.authOverlay) return;
  els.authOverlay.classList.remove("show");
  els.authOverlay.setAttribute("aria-hidden", "true");
}


function formatCurrency(value) {
  if (!Number.isFinite(value)) return "-";
  return value.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
}

function parseAmount(value) {
  if (value === null || value === undefined) return 0;
  const normalized = String(value)
    .replace(/\./g, "")
    .replace(",", ".")
    .replace(/[^\d.-]/g, "");
  const amount = Number.parseFloat(normalized);
  return Number.isFinite(amount) ? amount : 0;
}

function parseDate(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const match = raw.match(/(\d{2})\.(\d{2})\.(\d{2,4})/);
  if (!match) return null;
  let year = Number(match[3]);
  if (year < 100) {
    year += 2000;
  }
  const month = Number(match[2]) - 1;
  const day = Number(match[1]);
  return new Date(year, month, day);
}

function formatCsvValue(value) {
  const raw = String(value ?? "");
  const needsQuotes = /[;"\n]/.test(raw);
  if (!needsQuotes) {
    return raw;
  }
  const safe = raw.replace(/"/g, "\"\"");
  return `"${safe}"`;
}

function formatAmount(value) {
  if (!Number.isFinite(value)) return "";
  return value.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatAmountEuro(value) {
  if (!Number.isFinite(value)) return "";
  return value.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
}

function formatPeriodLabel(monthKey) {
  if (!monthKey) return "";
  const [year, month] = monthKey.split("-").map(Number);
  const names = [
    "Januar",
    "Februar",
    "Maerz",
    "April",
    "Mai",
    "Juni",
    "Juli",
    "August",
    "September",
    "Oktober",
    "November",
    "Dezember"
  ];
  if (!month || month < 1 || month > 12 || !year) return monthKey;
  return `${names[month - 1]} ${year}`;
}

function isInMonth(date, monthKey) {
  if (!date || !monthKey) return false;
  const [year, month] = monthKey.split("-").map(Number);
  return date.getFullYear() === year && date.getMonth() + 1 === month;
}

function categorizeTransaction(entry) {
  const text = normalizeText(`${entry.text} ${entry.purpose} ${entry.partner}`);
  const amount = entry.amount;
  if (amount >= 0) {
    if (text.includes("miete")) return "Mieteinnahmen";
    if (text.includes("kaution")) return "Kaution";
    if (text.includes("airbnb")) return "Airbnb Auszahlungen";
    if (text.includes("booking")) return "Booking Auszahlungen";
    if (text.includes("erstattung")) return "Erstattungen";
    return "Sonstige Einnahmen";
  }
  if (text.includes("entgeltabschluss") || text.includes("gebuehr")) {
    return "Bankgebuehren";
  }
  if (text.includes("lastschrift") || text.includes("sepa")) {
    return "Lastschriften";
  }
  return "Sonstige Ausgaben";
}

function parseCSV(text) {
  const rows = [];
  let current = "";
  let insideQuotes = false;
  const lines = [];
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (char === "\"") {
      insideQuotes = !insideQuotes;
      current += char;
      continue;
    }
    if (char === "\n" && !insideQuotes) {
      lines.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  if (current) lines.push(current);
  lines.forEach((line) => {
    if (!line.trim()) return;
    const cells = [];
    let cell = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i += 1) {
      const char = line[i];
      if (char === "\"") {
        inQuotes = !inQuotes;
        continue;
      }
      if (char === ";" && !inQuotes) {
        cells.push(cell);
        cell = "";
        continue;
      }
      cell += char;
    }
    cells.push(cell);
    rows.push(cells);
  });
  return rows;
}

function buildTransactions(rows) {
  if (!rows.length) return [];
  let headerIndex = 0;
  if (
    rows[0].length === 1 &&
    /^sep=/.test(String(rows[0][0] || "").trim().toLowerCase())
  ) {
    headerIndex = 1;
  }
  const headerProbe = rows.slice(headerIndex, headerIndex + 5);
  const detectedIndex = headerProbe.findIndex((row) =>
    row.some((cell) => {
      const norm = normalizeHeaderName(cell);
      return norm.includes("buchungstag") || norm.includes("valutadatum");
    })
  );
  if (detectedIndex !== -1) {
    headerIndex += detectedIndex;
  }
  if (!rows[headerIndex]) return [];
  const headers = rows[headerIndex].map((header) => String(header || "").trim());
  const normalized = headers.map((header) => normalizeHeaderName(header));
  const findIndexExact = (needle) =>
    normalized.findIndex((header) => header === needle);
  const findIndex = (needles) => {
    const list = Array.isArray(needles) ? needles : [needles];
    return normalized.findIndex((header) =>
      list.some((needle) => header.includes(needle))
    );
  };
  const idxDate = findIndex(["valutadatum"]);
  const idxBookingDate = findIndex(["buchungstag", "buchungsdatum"]);
  const idxText = findIndex(["buchungstext"]);
  const idxPurpose = findIndex(["verwendungszweck"]);
  const idxPartner = findIndex(["beguenstigter", "zahlungspflichtiger", "empfaenger"]);
  const idxIban = findIndex(["kontonummer", "iban"]);
  let idxAmount = findIndexExact("betrag");
  if (idxAmount === -1) {
    idxAmount = normalized.findIndex(
      (header) => header.includes("betrag") && !header.includes("ursprungsbetrag")
    );
  }
  const dataRows = rows.slice(headerIndex + 1);
  return dataRows
    .map((row) => {
      const bookingDate = parseDate(row[idxBookingDate]);
      const valueDate = parseDate(row[idxDate]);
      const date = bookingDate || valueDate;
      return {
        date,
        bookingDate,
        text: row[idxText] || "",
        purpose: row[idxPurpose] || "",
        partner: row[idxPartner] || "",
        iban: row[idxIban] || "",
        amount: parseAmount(row[idxAmount])
      };
    })
    .filter((entry) => entry.date);
}

function normalizeHeaderName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/\//g, "")
    .replace(/\u00e4/g, "ae")
    .replace(/\u00f6/g, "oe")
    .replace(/\u00fc/g, "ue")
    .replace(/\u00df/g, "ss");
}

function buildMonthOptions(items) {
  const months = Array.from(
    new Set(
      items
        .map((entry) => {
          if (!entry.date) return "";
          const year = entry.date.getFullYear();
          const month = String(entry.date.getMonth() + 1).padStart(2, "0");
          return `${year}-${month}`;
        })
        .filter(Boolean)
    )
  ).sort();
  els.monthSelect.innerHTML = "<option value=\"\">Zeitraum</option>";
  months.forEach((month) => {
    const option = document.createElement("option");
    option.value = month;
    option.textContent = month;
    els.monthSelect.appendChild(option);
  });
  if (months.length) {
    els.monthSelect.value = months[months.length - 1];
  }
  els.monthSelect.disabled = months.length === 0;
  if (els.exportEurBtn) {
    els.exportEurBtn.disabled = months.length === 0;
  }
  if (els.exportEurXlsBtn) {
    els.exportEurXlsBtn.disabled = months.length === 0 || !canExportXlsx();
  }
  if (els.snapshotBtn) {
    els.snapshotBtn.disabled = months.length === 0;
  }
  if (els.accountSelect) {
    els.accountSelect.disabled = false;
  }
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function isWithinWindow(date, monthKey, startDay, endDay) {
  if (!date) return false;
  const [year, month] = monthKey.split("-").map(Number);
  if (date.getFullYear() !== year) return false;
  if (date.getMonth() + 1 !== month) return false;
  const day = date.getDate();
  return day >= startDay && day <= endDay;
}

function matchTenant(tenant, items, monthKey) {
  const tolerance = Number(settings.tolerance) || 0;
  const startDay = Number(settings.windowStart) || 1;
  const endDay = Number(settings.windowEnd) || 31;
  const name = normalizeText(tenant.name);
  const iban = normalizeText(tenant.iban);
  const matches = items.filter((entry) => {
    if (entry.amount <= 0) return false;
    if (!isWithinWindow(entry.date, monthKey, startDay, endDay)) return false;
    const haystack = `${entry.text} ${entry.purpose} ${entry.partner}`.toLowerCase();
    const ibanMatch = iban && normalizeText(entry.iban).includes(iban);
    const nameMatch = name && haystack.includes(name);
    return ibanMatch || nameMatch;
  });
  const sum = matches.reduce((acc, entry) => acc + entry.amount, 0);
  const rent = Number(tenant.rent) || 0;
  const diff = sum - rent;
  let status = "missing";
  if (sum === 0) {
    status = "missing";
  } else if (Math.abs(diff) <= tolerance) {
    status = "ok";
  } else if (sum > rent + tolerance) {
    status = "over";
  } else if (settings.allowPartial) {
    status = "partial";
  }
  return { sum, status, matches };
}

function renderTenantTable(monthKey) {
  els.tenantTable.innerHTML = "";
  if (!monthKey) return;
  const accountId = getSelectedAccountId();
  const scopedTransactions = filterTransactionsByAccount(transactions, accountId);
  const rows = [];
  const usedMatches = new Set();
  settings.tenants.forEach((tenant) => {
    const result = matchTenant(tenant, scopedTransactions, monthKey);
    result.matches.forEach((entry) => usedMatches.add(entry));
    rows.push({ tenant, result });
  });

  rows.forEach(({ tenant, result }) => {
    const tr = document.createElement("tr");
    const label = result.status;
    const statusClass =
      label === "ok"
        ? "status-ok"
        : label === "partial"
          ? "status-partial"
          : label === "over"
            ? "status-over"
            : "status-missing";
    const statusText =
      label === "ok"
        ? "Bezahlt"
        : label === "partial"
          ? "Teilzahlung"
          : label === "over"
            ? "Ueberzahlt"
            : "Fehlt";
    const latestMatchDate =
      result.matches.length > 0
        ? result.matches
            .slice()
            .sort((a, b) => b.date - a.date)[0]
            .date.toLocaleDateString("de-DE")
        : "-";
    tr.innerHTML = `
      <td>${tenant.name || "-"}</td>
      <td>${tenant.unit || "-"}</td>
      <td>${formatCurrency(Number(tenant.rent) || 0)}</td>
      <td>${tenant.dueDay ? String(tenant.dueDay) : "-"}</td>
      <td>${latestMatchDate}</td>
      <td>${formatCurrency(result.sum)}</td>
      <td><span class="status-pill ${statusClass}">${statusText}</span></td>
    `;
    els.tenantTable.appendChild(tr);
  });

  const unmatched = scopedTransactions.filter(
    (entry) =>
      entry.amount > 0 &&
      isWithinWindow(
        entry.date,
        monthKey,
        Number(settings.windowStart) || 1,
        Number(settings.windowEnd) || 31
      ) &&
      !usedMatches.has(entry)
  );

  renderUnmatched(unmatched);
  updateStats(rows, monthKey, scopedTransactions);
}

function pushDashboardEntry(metrics) {
  try {
    const raw = localStorage.getItem(DASHBOARD_KEY);
    const history = raw ? JSON.parse(raw) : [];
    const safe = Array.isArray(history) ? history : [];
    safe.push({
      system: "miete",
      ts: Date.now(),
      metrics
    });
    localStorage.setItem(DASHBOARD_KEY, JSON.stringify(safe.slice(-200)));
  } catch (error) {
    // ignore
  }
  appendDashboardEntry(metrics).catch(() => {});
}

function computePaidSum(monthKey) {
  const accountId = getSelectedAccountId();
  const scopedTransactions = filterTransactionsByAccount(transactions, accountId);
  return settings.tenants.reduce((total, tenant) => {
    const result = matchTenant(tenant, scopedTransactions, monthKey);
    if (result.status === "ok" || result.status === "over") {
      return total + result.sum;
    }
    return total;
  }, 0);
}

function computeTotalIncome(monthKey) {
  const accountId = getSelectedAccountId();
  return filterTransactionsByAccount(transactions, accountId)
    .filter((entry) => isInMonth(entry.date, monthKey) && entry.amount > 0)
    .reduce((sum, entry) => sum + entry.amount, 0);
}

function setSnapshotStatus(message) {
  if (els.snapshotStatus) {
    els.snapshotStatus.textContent = message;
  }
}

function saveSnapshot() {
  const monthKey = els.monthSelect ? els.monthSelect.value : "";
  if (!monthKey) {
    setSnapshotStatus("Bitte zuerst einen Zeitraum w\u00e4hlen.");
    return;
  }
  const periodLabel = formatPeriodLabel(monthKey);
  const payout = computePaidSum(monthKey);
  const revenue = payout;
  pushDashboardEntry({
    revenue,
    fees: 0,
    tax: 0,
    payout,
    period: periodLabel
  });
  setSnapshotStatus("Snapshot gespeichert.");
}

function applyPayload(payload) {
  if (!payload) return;
  if (payload.settings) {
    settings = {
      ...DEFAULT_SETTINGS,
      ...payload.settings,
      accounts: Array.isArray(payload.settings.accounts)
        ? payload.settings.accounts
        : DEFAULT_SETTINGS.accounts,
      tenants: Array.isArray(payload.settings.tenants)
        ? payload.settings.tenants
        : []
    };
    saveSettings(settings);
  }
  if (Array.isArray(payload.transactions)) {
    transactions = deserializeTransactions(payload.transactions);
    localStorage.setItem(CSV_CACHE_KEY, JSON.stringify(transactions));
  }
  if (payload.month && els.monthSelect) {
    els.monthSelect.value = payload.month;
    localStorage.setItem(MONTH_KEY, payload.month);
  }
  if (payload.account && els.accountSelect) {
    localStorage.setItem(ACCOUNT_KEY, payload.account);
    els.accountSelect.value = payload.account;
  }
  buildMonthOptions(transactions);
  populateAccountSelect();
  renderTenantTable(els.monthSelect.value);
}

function renderUnmatched(items) {
  els.unmatchedTable.innerHTML = "";
  if (!items.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = "<td colspan=\"4\" class=\"muted\">Keine Eintraege gefunden.</td>";
    els.unmatchedTable.appendChild(tr);
    return;
  }
  items.forEach((entry) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${entry.date.toLocaleDateString("de-DE")}</td>
      <td>${entry.purpose || entry.text || "-"}</td>
      <td>${entry.partner || "-"}</td>
      <td>${formatCurrency(entry.amount)}</td>
    `;
    els.unmatchedTable.appendChild(tr);
  });
}

function exportEurCsv() {
  const monthKey = els.monthSelect ? els.monthSelect.value : "";
  if (!monthKey) return;
  const entries = filterTransactionsByAccount(
    transactions,
    getSelectedAccountId()
  ).filter((entry) => isInMonth(entry.date, monthKey));
  if (!entries.length) return;
  const categoryTotals = new Map();
  let sumIncome = 0;
  let sumExpense = 0;
  entries.forEach((entry) => {
    const category = categorizeTransaction(entry);
    const isIncome = entry.amount >= 0;
    const income = isIncome ? entry.amount : 0;
    const expense = isIncome ? 0 : Math.abs(entry.amount);
    sumIncome += income;
    sumExpense += expense;
    const current = categoryTotals.get(category) || { income: 0, expense: 0 };
    current.income += income;
    current.expense += expense;
    categoryTotals.set(category, current);
  });

  const rows = [
    [
      "Zeitraum",
      "Bereich",
      "Kategorie",
      "Einnahmen",
      "Ausgaben",
      "Saldo",
      "Datum",
      "Buchungstext",
      "Empf\u00e4nger",
      "Verwendungszweck"
    ]
  ];
  Array.from(categoryTotals.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([category, totals]) => {
      const saldo = totals.income - totals.expense;
      rows.push([
        monthKey,
        "Zusammenfassung",
        category,
        totals.income ? formatAmountEuro(totals.income) : "",
        totals.expense ? formatAmountEuro(totals.expense) : "",
        formatAmountEuro(saldo),
        "",
        "",
        "",
        ""
      ]);
    });
  rows.push([
    monthKey,
    "Zusammenfassung",
    "SUMME",
    formatAmountEuro(sumIncome),
    formatAmountEuro(sumExpense),
    formatAmountEuro(sumIncome - sumExpense),
    "",
    "",
    "",
    ""
  ]);
  entries
    .slice()
    .sort((a, b) => a.date - b.date)
    .forEach((entry) => {
      const category = categorizeTransaction(entry);
      const isIncome = entry.amount >= 0;
      const bookingDate = entry.bookingDate || entry.date;
      rows.push([
        monthKey,
        "Details",
        category,
        isIncome ? formatAmountEuro(entry.amount) : "",
        !isIncome ? formatAmountEuro(Math.abs(entry.amount)) : "",
        "",
        bookingDate ? bookingDate.toLocaleDateString("de-DE") : "",
        entry.text || "",
        entry.partner || "",
        entry.purpose || ""
      ]);
    });
  const content = rows.map((row) => row.map(formatCsvValue).join(";")).join("\n");
  const blob = new Blob([content], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `euer-${monthKey}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function buildEurData(monthKey) {
  const entries = filterTransactionsByAccount(
    transactions,
    getSelectedAccountId()
  ).filter((entry) => isInMonth(entry.date, monthKey));
  const categoryTotals = new Map();
  let sumIncome = 0;
  let sumExpense = 0;
  entries.forEach((entry) => {
    const category = categorizeTransaction(entry);
    const isIncome = entry.amount >= 0;
    const income = isIncome ? entry.amount : 0;
    const expense = isIncome ? 0 : Math.abs(entry.amount);
    sumIncome += income;
    sumExpense += expense;
    const current = categoryTotals.get(category) || { income: 0, expense: 0 };
    current.income += income;
    current.expense += expense;
    categoryTotals.set(category, current);
  });
  const summaryRows = Array.from(categoryTotals.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([category, totals]) => ({
      category,
      income: totals.income,
      expense: totals.expense,
      saldo: totals.income - totals.expense
    }));
  const detailsRows = entries
    .slice()
    .sort((a, b) => a.date - b.date)
    .map((entry) => {
      const category = categorizeTransaction(entry);
      const isIncome = entry.amount >= 0;
      const bookingDate = entry.bookingDate || entry.date;
      return {
        date: bookingDate ? bookingDate.toLocaleDateString("de-DE") : "",
        text: entry.text || "",
        partner: entry.partner || "",
        purpose: entry.purpose || "",
        category,
        income: isIncome ? entry.amount : 0,
        expense: !isIncome ? Math.abs(entry.amount) : 0
      };
    });
  return {
    summaryRows,
    detailsRows,
    sumIncome,
    sumExpense,
    saldo: sumIncome - sumExpense
  };
}

function canExportXlsx() {
  return typeof window !== "undefined" && typeof window.XLSX !== "undefined";
}

function buildEurSheetData(monthKey, summaryRows, detailsRows, sumIncome, sumExpense, saldo) {
  const summaryTable = [
    ["Kategorie", "Einnahmen", "Ausgaben", "Saldo"],
    ...summaryRows.map((row) => [
      row.category,
      row.income ? formatAmountEuro(row.income) : "",
      row.expense ? formatAmountEuro(row.expense) : "",
      formatAmountEuro(row.saldo)
    ]),
    ["SUMME", formatAmountEuro(sumIncome), formatAmountEuro(sumExpense), formatAmountEuro(saldo)]
  ];
  const detailsTable = [
    ["Datum", "Buchungstext", "Empf\u00e4nger", "Verwendungszweck", "Kategorie", "Einnahme", "Ausgabe"],
    ...detailsRows.map((row) => [
      row.date,
      row.text,
      row.partner,
      row.purpose,
      row.category,
      row.income ? formatAmountEuro(row.income) : "",
      row.expense ? formatAmountEuro(row.expense) : ""
    ])
  ];
  const title = `Einnahmen-\u00dcberschuss-Rechnung ${monthKey}`;
  const labelRow = 1;
  const headerRow = 2;
  const detailStartCol = 6;
  const totalRows = Math.max(summaryTable.length, detailsTable.length);
  const grid = Array.from({ length: headerRow + 1 + totalRows }, () => []);
  grid[0][0] = title;
  grid[labelRow][0] = "Zusammenfassung";
  grid[labelRow][detailStartCol] = "Details";
  for (let i = 0; i < totalRows; i += 1) {
    const rowIndex = headerRow + i;
    const summaryRow = summaryTable[i] || [];
    const detailsRow = detailsTable[i] || [];
    summaryRow.forEach((cell, colIndex) => {
      grid[rowIndex][colIndex] = cell;
    });
    detailsRow.forEach((cell, colIndex) => {
      grid[rowIndex][detailStartCol + colIndex] = cell;
    });
  }
  return { grid, detailStartCol, headerRow };
}

function exportEurXls() {
  const monthKey = els.monthSelect ? els.monthSelect.value : "";
  if (!monthKey) return;
  if (!canExportXlsx()) {
    window.alert("XLSX Export nicht verfuegbar. Bitte xlsx.full.min.js im Ordner lib ablegen.");
    return;
  }
  const { summaryRows, detailsRows, sumIncome, sumExpense, saldo } =
    buildEurData(monthKey);
  if (!summaryRows.length && !detailsRows.length) return;
  const sheetInfo = buildEurSheetData(
    monthKey,
    summaryRows,
    detailsRows,
    sumIncome,
    sumExpense,
    saldo
  );
  const ws = XLSX.utils.aoa_to_sheet(sheetInfo.grid);
  const detailStartCol = sheetInfo.detailStartCol;
  const detailHeaderRow = sheetInfo.headerRow;
  ws["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { r: detailHeaderRow, c: detailStartCol },
      e: { r: detailHeaderRow, c: detailStartCol + 6 }
    })
  };
  ws["!cols"] = [
    { wch: 16 },
    { wch: 14 },
    { wch: 22 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 12 },
    { wch: 28 },
    { wch: 24 },
    { wch: 28 },
    { wch: 18 },
    { wch: 14 },
    { wch: 14 }
  ];
  const wb = XLSX.utils.book_new();
  wb.Workbook = {
    Views: [
      {
        state: "frozen",
        xSplit: 0,
        ySplit: 4,
        topLeftCell: "A5",
        activePane: "bottomLeft"
      }
    ]
  };
  XLSX.utils.book_append_sheet(wb, ws, "EUER");
  XLSX.writeFile(wb, `euer-${monthKey}.xlsx`);
}

function clearImports() {
  const selectedMonth = els.monthSelect ? els.monthSelect.value : "";
  const selectedAccount = getSelectedAccountId();
  const scopeLabel =
    selectedMonth && selectedAccount && selectedAccount !== "Alle"
      ? `Monat ${selectedMonth} (nur gew\u00e4hltes Konto)`
      : selectedMonth
        ? `Monat ${selectedMonth} (alle Konten)`
        : "alle Daten";
  const ok = window.confirm(`Imports l\u00f6schen: ${scopeLabel}?`);
  if (!ok) return;
  if (selectedMonth) {
    transactions = transactions.filter((entry) => {
      if (!isInMonth(entry.date, selectedMonth)) return true;
      if (!selectedAccount || selectedAccount === "Alle") return false;
      return entry.accountId !== selectedAccount;
    });
  } else {
    transactions = [];
  }
  localStorage.setItem(CSV_CACHE_KEY, JSON.stringify(transactions));
  buildMonthOptions(transactions);
  if (!transactions.length) {
    localStorage.removeItem(MONTH_KEY);
    if (els.monthSelect) {
      els.monthSelect.value = "";
    }
    if (els.exportEurBtn) {
      els.exportEurBtn.disabled = true;
    }
    if (els.exportEurXlsBtn) {
      els.exportEurXlsBtn.disabled = true;
    }
  }
  renderTenantTable(els.monthSelect.value);
  renderUnmatched([]);
  els.statIncome.textContent = "-";
  els.statExpected.textContent = "-";
  els.statMissing.textContent = "-";
  els.statPartial.textContent = "-";
  triggerAutoBackup();
  scheduleRemoteSave();
}

function updateStats(rows, monthKey, scopedTransactions) {
  const source = scopedTransactions || transactions;
  const income = source
    .filter(
      (entry) =>
        entry.amount > 0 &&
        isWithinWindow(
          entry.date,
          monthKey,
          Number(settings.windowStart) || 1,
          Number(settings.windowEnd) || 31
        )
    )
    .reduce((acc, entry) => acc + entry.amount, 0);
  const expected = settings.tenants.reduce(
    (acc, tenant) => acc + (Number(tenant.rent) || 0),
    0
  );
  const missing = rows
    .filter((row) => row.result.status === "missing")
    .reduce((acc, row) => acc + (Number(row.tenant.rent) || 0), 0);
  const partial = rows.filter((row) => row.result.status === "partial").length;

  els.statIncome.textContent = formatCurrency(income);
  els.statExpected.textContent = formatCurrency(expected);
  els.statMissing.textContent = formatCurrency(missing);
  els.statPartial.textContent = String(partial);
}

function populateAccountSelect() {
  if (!els.accountSelect) return;
  els.accountSelect.innerHTML = "";
  const optionAll = document.createElement("option");
  optionAll.value = "Alle";
  optionAll.textContent = "Alle Konten";
  els.accountSelect.appendChild(optionAll);
  settings.accounts.forEach((account) => {
    const option = document.createElement("option");
    option.value = account.id;
    option.textContent = account.name || account.id;
    els.accountSelect.appendChild(option);
  });
  const stored = localStorage.getItem(ACCOUNT_KEY);
  if (stored && els.accountSelect.querySelector(`option[value=\"${stored}\"]`)) {
    els.accountSelect.value = stored;
  } else {
    els.accountSelect.value = "Alle";
  }
  els.accountSelect.disabled = false;
}

function getSelectedAccountId() {
  if (!els.accountSelect) return "Alle";
  return els.accountSelect.value || "Alle";
}

function filterTransactionsByAccount(items, accountId) {
  if (!accountId || accountId === "Alle") return items;
  return items.filter((entry) => entry.accountId === accountId);
}

function openSettings() {
  populateSettingsForm();
  els.settingsOverlay.classList.add("show");
  els.settingsOverlay.setAttribute("aria-hidden", "false");
}

function closeSettings() {
  els.settingsOverlay.classList.remove("show");
  els.settingsOverlay.setAttribute("aria-hidden", "true");
}

function populateSettingsForm() {
  els.toleranceInput.value = settings.tolerance;
  els.windowStartInput.value = settings.windowStart;
  els.windowEndInput.value = settings.windowEnd;
  els.allowPartialInput.checked = settings.allowPartial;
  renderSettingsTable();
  renderAccountsTable();
}

function renderSettingsTable() {
  els.settingsTable.innerHTML = "";
  settings.tenants.forEach((tenant, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input data-field="name" value="${tenant.name || ""}"></td>
      <td><input data-field="unit" value="${tenant.unit || ""}"></td>
      <td><input data-field="rent" type="number" step="0.01" value="${tenant.rent || 0}"></td>
      <td><input data-field="dueDay" type="number" min="1" max="31" value="${tenant.dueDay || ""}"></td>
      <td><input data-field="iban" value="${tenant.iban || ""}"></td>
      <td><button class="ghost" data-action="remove">Entfernen</button></td>
    `;
    tr.querySelectorAll("input").forEach((input) => {
      input.addEventListener("input", (event) => {
        const field = event.target.getAttribute("data-field");
        settings.tenants[index][field] = event.target.value;
      });
    });
    tr.querySelector("[data-action=\"remove\"]").addEventListener("click", () => {
      settings.tenants.splice(index, 1);
      renderSettingsTable();
    });
    els.settingsTable.appendChild(tr);
  });
}

function renderAccountsTable() {
  if (!els.accountsTable) return;
  els.accountsTable.innerHTML = "";
  settings.accounts.forEach((account, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input data-field="name" value="${account.name || ""}"></td>
      <td><input data-field="iban" value="${account.iban || ""}"></td>
      <td><button class="ghost" data-action="remove">Entfernen</button></td>
    `;
    tr.querySelectorAll("input").forEach((input) => {
      input.addEventListener("input", (event) => {
        const field = event.target.getAttribute("data-field");
        settings.accounts[index][field] = event.target.value;
      });
    });
    tr.querySelector("[data-action=\"remove\"]").addEventListener("click", () => {
      settings.accounts.splice(index, 1);
      renderAccountsTable();
      populateAccountSelect();
    });
    els.accountsTable.appendChild(tr);
  });
}

function addTenantRow() {
  settings.tenants.push({
    name: "",
    unit: "",
    rent: 0,
    dueDay: settings.windowStart,
    iban: ""
  });
  renderSettingsTable();
}

function addAccountRow() {
  const nextId = `konto-${Date.now()}`;
  settings.accounts.push({
    id: nextId,
    name: "",
    iban: ""
  });
  renderAccountsTable();
  populateAccountSelect();
}

function handleSaveSettings() {
  const next = {
    tolerance: Number(els.toleranceInput.value) || 0,
    windowStart: Number(els.windowStartInput.value) || 1,
    windowEnd: Number(els.windowEndInput.value) || 31,
    allowPartial: els.allowPartialInput.checked,
    accounts: settings.accounts.map((account) => ({
      id: account.id || `konto-${Date.now()}`,
      name: account.name || "",
      iban: account.iban || ""
    })),
    tenants: settings.tenants.map((tenant) => ({
      ...tenant,
      rent: Number(tenant.rent) || 0,
      dueDay: Number(tenant.dueDay) || 0
    }))
  };
  saveSettings(next);
  closeSettings();
  populateAccountSelect();
  renderTenantTable(els.monthSelect.value);
  triggerAutoBackup();
  scheduleRemoteSave();
}

function handleCsvUpload(file) {
  const reader = new FileReader();
  reader.onload = () => {
    const accountId = getSelectedAccountId();
    if (!accountId || accountId === "Alle") {
      window.alert("Bitte zuerst ein Konto ausw\u00e4hlen.");
      return;
    }
    const rows = parseCSV(String(reader.result || ""));
    const newEntries = buildTransactions(rows).map((entry) => ({
      ...entry,
      accountId
    }));
    transactions = transactions.concat(newEntries);
    localStorage.setItem(CSV_CACHE_KEY, JSON.stringify(transactions));
    buildMonthOptions(transactions);
    if (els.monthSelect.value) {
      localStorage.setItem(MONTH_KEY, els.monthSelect.value);
    }
    renderTenantTable(els.monthSelect.value);
    triggerAutoBackup();
    scheduleRemoteSave();
  };
  reader.readAsText(file, "ISO-8859-1");
}

if (els.csvInput) {
  els.csvInput.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      handleCsvUpload(file);
    }
  });
}

if (els.monthSelect) {
  els.monthSelect.addEventListener("change", () => {
    if (els.monthSelect.value) {
      localStorage.setItem(MONTH_KEY, els.monthSelect.value);
    }
    renderTenantTable(els.monthSelect.value);
    triggerAutoBackup();
    scheduleRemoteSave();
  });
}

if (els.accountSelect) {
  els.accountSelect.addEventListener("change", () => {
    if (els.accountSelect.value) {
      localStorage.setItem(ACCOUNT_KEY, els.accountSelect.value);
    }
    renderTenantTable(els.monthSelect.value);
    scheduleRemoteSave();
  });
}

if (els.exportEurBtn) {
  els.exportEurBtn.addEventListener("click", exportEurCsv);
}

if (els.exportEurXlsBtn) {
  els.exportEurXlsBtn.addEventListener("click", exportEurXls);
}

if (els.clearImportsBtn) {
  els.clearImportsBtn.addEventListener("click", clearImports);
}

if (els.snapshotBtn) {
  els.snapshotBtn.addEventListener("click", saveSnapshot);
}

if (els.openSettingsBtn) {
  els.openSettingsBtn.addEventListener("click", openSettings);
}

if (els.closeSettingsBtn) {
  els.closeSettingsBtn.addEventListener("click", closeSettings);
}

if (els.closeSettingsBtnFooter) {
  els.closeSettingsBtnFooter.addEventListener("click", closeSettings);
}

if (els.addTenantBtn) {
  els.addTenantBtn.addEventListener("click", addTenantRow);
}

if (els.addAccountBtn) {
  els.addAccountBtn.addEventListener("click", addAccountRow);
}

if (els.saveSettingsBtn) {
  els.saveSettingsBtn.addEventListener("click", handleSaveSettings);
}

if (els.settingsOverlay) {
  els.settingsOverlay.addEventListener("click", (event) => {
    if (event.target === els.settingsOverlay) {
      closeSettings();
    }
  });
}

populateSettingsForm();
populateAccountSelect();

function restoreCachedCsv() {
  if (remoteLoaded) return;
  const raw = localStorage.getItem(CSV_CACHE_KEY);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return;
    transactions = parsed.map((entry) => ({
      ...entry,
      date: entry.date ? new Date(entry.date) : null,
      bookingDate: entry.bookingDate ? new Date(entry.bookingDate) : null
    }));
    buildMonthOptions(transactions);
    const storedMonth = localStorage.getItem(MONTH_KEY);
    if (storedMonth) {
      els.monthSelect.value = storedMonth;
    }
    populateAccountSelect();
    const storedAccount = localStorage.getItem(ACCOUNT_KEY);
    if (storedAccount && els.accountSelect) {
      els.accountSelect.value = storedAccount;
    }
    renderTenantTable(els.monthSelect.value);
  } catch (error) {
    return;
  }
}

async function restoreBackupIfNeeded() {
  const handle = await loadBackupHandle().catch(() => null);
  await updateBackupStatus(handle);
  if (remoteLoaded) return;
  const cached = localStorage.getItem(CSV_CACHE_KEY);
  if (cached) return;
  if (!handle) return;
  try {
    const payload = await readBackup(handle);
    if (!payload) return;
      if (payload.settings) {
        const accounts = Array.isArray(payload.settings.accounts)
          ? payload.settings.accounts
          : DEFAULT_SETTINGS.accounts;
        settings = {
          ...DEFAULT_SETTINGS,
          ...payload.settings,
          accounts: accounts.length ? accounts : DEFAULT_SETTINGS.accounts,
          tenants: Array.isArray(payload.settings.tenants)
            ? payload.settings.tenants
            : []
        };
        saveSettings(settings);
      }
    transactions = deserializeTransactions(payload.transactions);
    if (transactions.length) {
      localStorage.setItem(CSV_CACHE_KEY, JSON.stringify(transactions));
      buildMonthOptions(transactions);
      if (payload.month) {
        els.monthSelect.value = payload.month;
        localStorage.setItem(MONTH_KEY, payload.month);
      }
      if (payload.account) {
        localStorage.setItem(ACCOUNT_KEY, payload.account);
      }
      populateAccountSelect();
      const storedAccount = localStorage.getItem(ACCOUNT_KEY);
      if (storedAccount && els.accountSelect) {
        els.accountSelect.value = storedAccount;
      }
      renderTenantTable(els.monthSelect.value);
    }
  } catch (error) {
    return;
  }
}

let backupHandleCache = null;

async function triggerAutoBackup() {
  if (!backupHandleCache) return;
  try {
    await writeBackup(backupHandleCache);
    updateBackupStatus(backupHandleCache);
  } catch (error) {
    return;
  }
}

async function chooseBackupFile() {
  if (!window.showSaveFilePicker) {
    if (els.backupStatus) {
      els.backupStatus.textContent = "Browser unterstuetzt kein Auto-Backup.";
    }
    return;
  }
  const handle = await window.showSaveFilePicker({
    suggestedName: "miete-backup.json",
    types: [
      {
        description: "JSON",
        accept: { "application/json": [".json"] }
      }
    ]
  });
  await saveBackupHandle(handle);
  backupHandleCache = handle;
  await updateBackupStatus(handle);
  await triggerAutoBackup();
}

async function initBackupHandle() {
  backupHandleCache = await loadBackupHandle().catch(() => null);
  await updateBackupStatus(backupHandleCache);
}

if (els.backupSelectBtn) {
  els.backupSelectBtn.addEventListener("click", () => {
    chooseBackupFile();
  });
}

if (els.backupNowBtn) {
  els.backupNowBtn.addEventListener("click", () => {
    triggerAutoBackup();
  });
}

initBackupHandle();
async function handleLogin() {
  if (!els.loginUser || !els.loginPass) return;
  const username = els.loginUser.value.trim();
  const password = els.loginPass.value.trim();
  const email = toEmail(username);
  if (!username || !password) {
    if (els.loginError) {
      els.loginError.textContent = "Bitte Benutzername und Passwort eingeben.";
    }
    return;
  }
  try {
    if (!auth) {
      throw new Error("auth_missing");
    }
    await signInWithEmailAndPassword(auth, email, password);
    hideAuthOverlay();
  } catch (error) {
    if (els.loginError) {
      els.loginError.textContent = "Login fehlgeschlagen.";
    }
  }
}

if (els.loginBtn) {
  els.loginBtn.addEventListener("click", () => {
    handleLogin();
  });
}

if (auth && onAuthStateChanged) {
  onAuthStateChanged(auth, async (user) => {
    currentUser = user;
    if (!user) {
      currentProfile = null;
      showAuthOverlay();
      return;
    }
    hideAuthOverlay();
    const profile = await ensureUserProfile(user);
    if (profile && profile.role === "disabled") {
      if (els.loginError) {
        els.loginError.textContent = "Konto deaktiviert.";
      }
      await signOut(auth);
      return;
    }
    currentProfile = profile;
    await loadRemoteData();
    restoreCachedCsv();
    restoreBackupIfNeeded();
  });
} else {
  showAuthOverlay();
  restoreCachedCsv();
  restoreBackupIfNeeded();
}
