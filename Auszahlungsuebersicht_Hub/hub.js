const frame = document.getElementById("appFrame");
const buttons = document.querySelectorAll("button[data-target]");
const welcome = document.getElementById("welcome");
const menuToggle = document.getElementById("menuToggle");
const menuPanel = document.getElementById("menuPanel");
const menuClose = document.getElementById("menuClose");
const PERIOD_KEY = "dashboard_period";
const SYSTEM_KEY = "dashboard_system";
const ADMIN_CREDS_KEY = "admin_creds_v1";
const dashboardEls = {
  revenue: document.getElementById("kpiRevenue"),
  fees: document.getElementById("kpiFees"),
  tax: document.getElementById("kpiTax"),
  payout: document.getElementById("kpiPayout"),
  revenueTrend: document.getElementById("kpiRevenueTrend"),
  feesTrend: document.getElementById("kpiFeesTrend"),
  taxTrend: document.getElementById("kpiTaxTrend"),
  payoutTrend: document.getElementById("kpiPayoutTrend"),
  updated: document.getElementById("kpiUpdated"),
  periodSelect: document.getElementById("periodSelect"),
  systemOptions: document.getElementById("systemOptions"),
  toggleBlurBtn: document.getElementById("toggleBlurBtn"),
  kpiGrid: document.getElementById("kpiGrid"),
  openFilterBtn: document.getElementById("openFilterBtn"),
  filterOverlay: document.getElementById("filterOverlay"),
  closeFilterBtn: document.getElementById("closeFilterBtn"),
  closeFilterBtnFooter: document.getElementById("closeFilterBtnFooter"),
  openSnapshotsBtn: document.getElementById("openSnapshotsBtn"),
  snapshotsOverlay: document.getElementById("snapshotsOverlay"),
  closeSnapshotsBtn: document.getElementById("closeSnapshotsBtn"),
  closeSnapshotsBtnFooter: document.getElementById("closeSnapshotsBtnFooter"),
  snapshotListModal: document.getElementById("snapshotListModal"),
  snapshotPeriodSelect: document.getElementById("snapshotPeriodSelect"),
  snapshotSystemSelect: document.getElementById("snapshotSystemSelect")
};
const loginGate = document.getElementById("loginGate");
const loginForm = document.getElementById("loginForm");
const loginId = document.getElementById("loginId");
const loginPw = document.getElementById("loginPw");
const loginError = document.getElementById("loginError");
const openAdminBtn = document.getElementById("openAdminBtn");
const adminOverlay = document.getElementById("adminOverlay");
const closeAdminBtn = document.getElementById("closeAdminBtn");
const closeAdminBtnFooter = document.getElementById("closeAdminBtnFooter");
const logoutBtn = document.getElementById("logoutBtn");
const goDashboardBtn = document.getElementById("goDashboardBtn");
const newUserName = document.getElementById("newUserName");
const newUserPass = document.getElementById("newUserPass");
const newUserAdmin = document.getElementById("newUserAdmin");
const addUserBtn = document.getElementById("addUserBtn");
const usersTable = document.querySelector("#usersTable tbody");
const adminError = document.getElementById("adminError");

const firebase = window.firebaseServices || {};
const {
  auth,
  db,
  onAuthStateChanged,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut,
  doc,
  getDoc,
  setDoc,
  getDocs,
  collection,
  serverTimestamp
} = firebase;

let currentUser = null;
let currentProfile = null;
let dashboardHistory = [];

const targets = {
  airbnb: "../Airbnb to XLS/index.html",
  booking: "../Booking_Auszahlungsuebersicht_HTML/index.html",
  direkt: "../Direktbuchungen/index.html",
  miete: "../Mieteingangskontrolle/index.html"
};

const menuTargetButtons = menuPanel
  ? Array.from(menuPanel.querySelectorAll("button[data-target]"))
  : [];
const menuDashboardBtn = document.getElementById("goDashboardBtn");

function updateMenuTargets(activeKey) {
  if (!menuTargetButtons.length) return;
  menuTargetButtons.forEach((btn) => {
    if (!activeKey) {
      btn.style.display = "";
      return;
    }
    const isActive = btn.dataset.target === activeKey;
    btn.style.display = isActive ? "none" : "";
  });
  if (menuDashboardBtn) {
    menuDashboardBtn.style.display = activeKey ? "" : "none";
  }
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

function storeAdminCreds(email, password) {
  if (!email || !password) return;
  sessionStorage.setItem(ADMIN_CREDS_KEY, JSON.stringify({ email, password }));
}

function readAdminCreds() {
  const raw = sessionStorage.getItem(ADMIN_CREDS_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function clearAdminCreds() {
  sessionStorage.removeItem(ADMIN_CREDS_KEY);
}

function setTarget(key) {
  frame.src = targets[key] || "";
  frame.style.display = key ? "block" : "none";
  if (welcome) {
    welcome.style.display = key ? "none" : "flex";
  }
  if (menuToggle) {
    menuToggle.classList.toggle("visible", true);
  }
  if (menuPanel) {
    menuPanel.classList.remove("show");
  }
  updateMenuTargets(key);
}

buttons.forEach((btn) => {
  btn.addEventListener("click", () => {
    setTarget(btn.dataset.target);
  });
});

function unlock() {
  document.body.classList.remove("locked");
  loginGate.style.display = "none";
  setTarget("");
}

function lock() {
  document.body.classList.add("locked");
  loginGate.style.display = "flex";
  frame.src = "";
  frame.style.display = "none";
  if (welcome) {
    welcome.style.display = "flex";
  }
  if (menuToggle) {
    menuToggle.classList.remove("visible");
  }
}

async function logout() {
  loginError.textContent = "";
  currentUser = null;
  currentProfile = null;
  dashboardHistory = [];
  setAdminVisible(false);
  closeAdmin();
  if (menuPanel) {
    menuPanel.classList.remove("show");
  }
  if (menuToggle) {
    menuToggle.setAttribute("aria-expanded", "false");
  }
  clearAdminCreds();
  if (auth) {
    await signOut(auth);
  }
  lock();
}

function setAdminVisible(isAdmin) {
  const blocks = document.querySelectorAll(".admin-only");
  blocks.forEach((block) => block.classList.toggle("show", Boolean(isAdmin)));
}

function openAdmin() {
  if (!adminOverlay) return;
  if (!currentProfile || currentProfile.role !== "admin") return;
  if (adminError) adminError.textContent = "";
  adminOverlay.classList.add("show");
  adminOverlay.setAttribute("aria-hidden", "false");
  loadUsers();
}

function closeAdmin() {
  if (!adminOverlay) return;
  adminOverlay.classList.remove("show");
  adminOverlay.setAttribute("aria-hidden", "true");
}

async function ensureUserProfile(user, usernameHint) {
  if (!db || !user) return null;
  const ref = doc(db, "users", user.uid);
  const snapshot = await getDoc(ref);
  if (snapshot.exists()) {
    return snapshot.data();
  }
  const email = user.email || "";
  const username = usernameHint || usernameFromEmail(email) || "user";
  const role = username === "admin" ? "admin" : "user";
  const profile = {
    username,
    role,
    createdAt: serverTimestamp()
  };
  await setDoc(ref, profile, { merge: true });
  return { ...profile };
}

function renderUsers(users) {
  if (!usersTable) return;
  usersTable.innerHTML = "";
  if (!users.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = "<td colspan=\"3\" class=\"muted\">Keine Benutzer.</td>";
    usersTable.appendChild(tr);
    return;
  }
  users.forEach((user) => {
    const roleLabel =
      user.role === "admin"
        ? "Admin"
        : user.role === "disabled"
          ? "Deaktiviert"
          : "Benutzer";
    const tr = document.createElement("tr");
    const disabled = currentUser && user.id === currentUser.uid;
    tr.innerHTML = `
      <td>${user.username || "-"}</td>
      <td>${roleLabel}</td>
      <td><button class="ghost" data-id="${user.id}" ${disabled ? "disabled" : ""}>Entfernen</button></td>
    `;
    const btn = tr.querySelector("button");
    btn.addEventListener("click", async () => {
      if (!confirm("Benutzer wirklich entfernen?")) return;
      await setDoc(
        doc(db, "users", user.id),
        { role: "disabled", disabled: true, updatedAt: serverTimestamp() },
        { merge: true }
      );
      loadUsers();
    });
    usersTable.appendChild(tr);
  });
}

async function loadUsers() {
  if (!db || !currentProfile || currentProfile.role !== "admin") return;
  const snapshot = await getDocs(collection(db, "users"));
  const users = snapshot.docs.map((docSnap) => ({
    id: docSnap.id,
    ...docSnap.data()
  }));
  renderUsers(users);
}

function dashboardDocRef() {
  if (!currentUser || !db) return null;
  return doc(db, "users", currentUser.uid, "dashboard", "summary");
}

async function loadDashboardHistory() {
  const ref = dashboardDocRef();
  if (!ref) {
    dashboardHistory = [];
    return;
  }
  const snapshot = await getDoc(ref);
  const data = snapshot.exists() ? snapshot.data() : null;
  dashboardHistory = Array.isArray(data && data.snapshots) ? data.snapshots : [];
}

async function saveDashboardHistory() {
  const ref = dashboardDocRef();
  if (!ref) return;
  await setDoc(
    ref,
    { snapshots: dashboardHistory.slice(-200), updatedAt: serverTimestamp() },
    { merge: true }
  );
}

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const id = loginId.value.trim();
  const pw = loginPw.value.trim();
  const email = toEmail(id);
  if (!email || !pw) {
    loginError.textContent = "Bitte ID und Passwort eingeben.";
    return;
  }
  if (!auth) {
    loginError.textContent = "Firebase ist nicht geladen.";
    return;
  }
  try {
    await signInWithEmailAndPassword(auth, email, pw);
    storeAdminCreds(email, pw);
    loginError.textContent = "";
  } catch (error) {
    loginError.textContent = "ID oder Passwort ist falsch.";
  }
});

if (logoutBtn) {
  logoutBtn.addEventListener("click", () => {
    logout();
  });
}

if (openAdminBtn) {
  openAdminBtn.addEventListener("click", openAdmin);
}

if (closeAdminBtn) {
  closeAdminBtn.addEventListener("click", closeAdmin);
}

if (closeAdminBtnFooter) {
  closeAdminBtnFooter.addEventListener("click", closeAdmin);
}

if (adminOverlay) {
  adminOverlay.addEventListener("click", (event) => {
    if (event.target === adminOverlay) {
      closeAdmin();
    }
  });
}

if (goDashboardBtn) {
  goDashboardBtn.addEventListener("click", () => {
    setTarget("");
  });
}

if (addUserBtn) {
  addUserBtn.addEventListener("click", async () => {
    if (!currentProfile || currentProfile.role !== "admin") return;
    const username = newUserName.value.trim();
    const password = newUserPass.value.trim();
    const isAdmin = newUserAdmin.checked;
    const email = toEmail(username);
    if (!email || !password) return;
    if (adminError) adminError.textContent = "";
    if (password.length < 6) {
      if (adminError) {
        adminError.textContent = "Passwort muss mindestens 6 Zeichen haben.";
      }
      return;
    }
    const adminCreds = readAdminCreds();
    if (!adminCreds) {
      if (adminError) {
        adminError.textContent = "Bitte als Admin neu anmelden.";
      }
      return;
    }
    try {
      await createUserWithEmailAndPassword(auth, email, password);
      const newUser = auth.currentUser;
      if (newUser) {
        await setDoc(
          doc(db, "users", newUser.uid),
          {
            username: usernameFromEmail(email),
            role: isAdmin ? "admin" : "user",
            createdAt: serverTimestamp()
          },
          { merge: true }
        );
      }
      await signOut(auth);
      await signInWithEmailAndPassword(
        auth,
        adminCreds.email,
        adminCreds.password
      );
      newUserName.value = "";
      newUserPass.value = "";
      newUserAdmin.checked = false;
      if (adminError) adminError.textContent = "Benutzer angelegt.";
      loadUsers();
    } catch (error) {
      if (adminError) {
        const message = String(error && error.message ? error.message : "");
        if (message.includes("email-already-in-use")) {
          adminError.textContent = "Benutzername ist bereits vergeben.";
        } else if (message.includes("invalid-email")) {
          adminError.textContent = "Ung\u00fcltige E-Mail/Benutzername.";
        } else {
          adminError.textContent = "Benutzer konnte nicht angelegt werden.";
        }
      }
    }
  });
}

function formatCurrency(value) {
  return new Intl.NumberFormat("de-DE", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(value || 0);
}

function readDashboardHistory() {
  return dashboardHistory;
}

function normalizePeriod(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const normalized = raw
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
  const yearMatch = normalized.match(/(20\d{2})/);
  const year = yearMatch ? yearMatch[1] : "";
  const monthMap = [
    ["januar", "Januar"],
    ["februar", "Februar"],
    ["marz", "M\u00e4rz"],
    ["maerz", "M\u00e4rz"],
    ["april", "April"],
    ["mai", "Mai"],
    ["juni", "Juni"],
    ["juli", "Juli"],
    ["august", "August"],
    ["september", "September"],
    ["oktober", "Oktober"],
    ["november", "November"],
    ["dezember", "Dezember"]
  ];
  const month = monthMap.find(([key]) => normalized.includes(key));
  if (month && year) {
    return `${month[1]} ${year}`;
  }
  const numericMatch = normalized.match(/(\d{1,2})\s*[./-]\s*(20\d{2})/);
  if (numericMatch) {
    const mm = numericMatch[1].padStart(2, "0");
    return `${mm}/${numericMatch[2]}`;
  }
  return raw;
}

function isSystemMatch(systemFilter, systemValue) {
  if (!systemFilter || systemFilter === "Alle") return true;
  if (Array.isArray(systemFilter)) {
    if (systemFilter.includes("Alle")) return true;
    return systemFilter.includes(systemValue);
  }
  return systemFilter === systemValue;
}

function ensureSystemSelection() {
  const checkboxes = getSystemCheckboxes();
  if (!checkboxes.length) return;
  const hasAny = checkboxes.some((box) => box.checked);
  if (!hasAny) {
    const allBox = checkboxes.find((box) => box.value === "Alle");
    if (allBox) {
      allBox.checked = true;
    }
  }
}

function getLatestBySystemAndPeriod(history, period, systemFilter) {
  const latest = new Map();
  history.forEach((entry) => {
    if (!entry || !entry.system || !entry.metrics) return;
    const entryPeriod = normalizePeriod(entry.metrics.period);
    if (period && period !== "Alle" && entryPeriod !== period) return;
    if (!isSystemMatch(systemFilter, entry.system)) return;
    const existing = latest.get(entry.system);
    if (!existing || entry.ts > existing.ts) {
      latest.set(entry.system, entry);
    }
  });
  return Array.from(latest.values());
}

function sumMetrics(entries) {
  return entries.reduce(
    (acc, entry) => {
      const m = entry.metrics || {};
      acc.revenue += m.revenue || 0;
      acc.fees += m.fees || 0;
      acc.tax += m.tax || 0;
      acc.payout += m.payout || 0;
      return acc;
    },
    { revenue: 0, fees: 0, tax: 0, payout: 0 }
  );
}

function parsePeriodLabel(label) {
  const raw = String(label || "").trim();
  if (!raw) return null;
  const numericMatch = raw.match(/(\d{1,2})\s*[./-]\s*(20\d{2})/);
  if (numericMatch) {
    const month = Number(numericMatch[1]);
    const year = Number(numericMatch[2]);
    if (month >= 1 && month <= 12) return { year, month };
  }
  const normalized = raw.toLowerCase();
  const yearMatch = normalized.match(/(20\d{2})/);
  const year = yearMatch ? Number(yearMatch[1]) : null;
  if (!year) return null;
  const monthKeys = [
    ["januar", 1],
    ["februar", 2],
    ["marz", 3],
    ["maerz", 3],
    ["april", 4],
    ["mai", 5],
    ["juni", 6],
    ["juli", 7],
    ["august", 8],
    ["september", 9],
    ["oktober", 10],
    ["november", 11],
    ["dezember", 12]
  ];
  const monthMatch = monthKeys.find(([name]) => normalized.includes(name));
  if (!monthMatch) return null;
  return { year, month: monthMatch[1] };
}

function getPreviousPeriodLabel(label) {
  const parsed = parsePeriodLabel(normalizePeriod(label));
  if (!parsed) return "";
  let year = parsed.year;
  let month = parsed.month - 1;
  if (month === 0) {
    month = 12;
    year -= 1;
  }
  const monthNames = [
    "Januar",
    "Februar",
    "M\u00e4rz",
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
  return `${monthNames[month - 1]} ${year}`;
}

function getLatestEntryForSystemAndPeriod(history, system, periodLabel) {
  let latest = null;
  history.forEach((entry) => {
    if (!entry || entry.system !== system || !entry.metrics) return;
    const entryPeriod = normalizePeriod(entry.metrics.period);
    if (entryPeriod !== periodLabel) return;
    if (!latest || entry.ts > latest.ts) {
      latest = entry;
    }
  });
  return latest;
}

function getPreviousMonthTotals(history, latestEntries, selectedPeriod, selectedSystem) {
  if (!latestEntries.length) return null;
  if (selectedPeriod && selectedPeriod !== "Alle") {
    const prevPeriod = getPreviousPeriodLabel(selectedPeriod);
    if (!prevPeriod) return null;
    const prevEntries = getLatestBySystemAndPeriod(
      history,
      prevPeriod,
      selectedSystem
    );
    return prevEntries.length ? sumMetrics(prevEntries) : null;
  }
  const prevEntries = [];
  latestEntries.forEach((entry) => {
    const currentPeriod = entry.metrics ? entry.metrics.period : "";
    const prevPeriod = getPreviousPeriodLabel(currentPeriod);
    if (!prevPeriod) return;
    const prevEntry = getLatestEntryForSystemAndPeriod(
      history,
      entry.system,
      prevPeriod
    );
    if (prevEntry) {
      prevEntries.push(prevEntry);
    }
  });
  return prevEntries.length ? sumMetrics(prevEntries) : null;
}

function formatDeltaCurrency(current, previous) {
  if (previous === null || previous === undefined) return "-";
  if (!Number.isFinite(previous)) return "-";
  const diff = current - previous;
  if (diff === 0) return formatCurrency(0);
  const sign = diff > 0 ? "+" : "-";
  return `${sign}${formatCurrency(Math.abs(diff))}`;
}

function populatePeriods(history) {
  if (!dashboardEls.periodSelect) return;
  const periods = Array.from(
    new Set(
      history
        .map((entry) => entry.metrics && normalizePeriod(entry.metrics.period))
        .filter(Boolean)
    )
  ).sort();
  dashboardEls.periodSelect.innerHTML = "";
  ["Alle", ...periods].forEach((period) => {
    const option = document.createElement("option");
    option.value = period;
    option.textContent = period;
    dashboardEls.periodSelect.appendChild(option);
  });
  const stored = localStorage.getItem(PERIOD_KEY);
  if (stored && periods.includes(stored)) {
    dashboardEls.periodSelect.value = stored;
  } else {
    dashboardEls.periodSelect.value = "Alle";
  }
}

function getSystemCheckboxes() {
  if (!dashboardEls.systemOptions) return [];
  return Array.from(
    dashboardEls.systemOptions.querySelectorAll("input[type=\"checkbox\"]")
  );
}

function readStoredSystems() {
  const stored = localStorage.getItem(SYSTEM_KEY);
  if (!stored) return ["Alle"];
  try {
    const parsed = JSON.parse(stored);
    if (Array.isArray(parsed) && parsed.length) return parsed;
  } catch (error) {
    // ignore
  }
  return [stored];
}

function applySystemSelection(systems) {
  const list = Array.isArray(systems) && systems.length ? systems : ["Alle"];
  const hasAll = list.includes("Alle");
  const checkboxes = getSystemCheckboxes();
  checkboxes.forEach((checkbox) => {
    if (checkbox.value === "Alle") {
      checkbox.checked = hasAll;
      return;
    }
    checkbox.checked = !hasAll && list.includes(checkbox.value);
  });
}

function getSelectedSystems() {
  const checkboxes = getSystemCheckboxes();
  if (!checkboxes.length) return ["Alle"];
  const selected = checkboxes.filter((box) => box.checked).map((box) => box.value);
  if (!selected.length || selected.includes("Alle")) return ["Alle"];
  return selected;
}

function restoreSystemFilter() {
  const stored = readStoredSystems();
  applySystemSelection(stored);
  ensureSystemSelection();
}

function populateSnapshotFilters(history) {
  if (!dashboardEls.snapshotPeriodSelect) return;
  const currentPeriod = dashboardEls.snapshotPeriodSelect.value;
  const currentSystem = dashboardEls.snapshotSystemSelect
    ? dashboardEls.snapshotSystemSelect.value
    : "Alle";
  const periods = Array.from(
    new Set(
      history
        .map((entry) => entry.metrics && normalizePeriod(entry.metrics.period))
        .filter(Boolean)
    )
  ).sort();
  dashboardEls.snapshotPeriodSelect.innerHTML = "";
  ["Alle", ...periods].forEach((period) => {
    const option = document.createElement("option");
    option.value = period;
    option.textContent = period;
    dashboardEls.snapshotPeriodSelect.appendChild(option);
  });
  const nextPeriod =
    currentPeriod && ["Alle", ...periods].includes(currentPeriod)
      ? currentPeriod
      : "Alle";
  dashboardEls.snapshotPeriodSelect.value = nextPeriod;
  if (dashboardEls.snapshotSystemSelect) {
    dashboardEls.snapshotSystemSelect.value = currentSystem || "Alle";
  }
}

function renderDashboard() {
  if (!dashboardEls.revenue) return;
  const history = readDashboardHistory();
  if (!history.length) {
    dashboardEls.revenue.textContent = "-";
    dashboardEls.fees.textContent = "-";
    dashboardEls.tax.textContent = "-";
    dashboardEls.payout.textContent = "-";
    dashboardEls.revenueTrend.textContent = "Vormonat: -";
    dashboardEls.feesTrend.textContent = "Vormonat: -";
    dashboardEls.taxTrend.textContent = "Vormonat: -";
    dashboardEls.payoutTrend.textContent = "Vormonat: -";
    dashboardEls.updated.textContent = "Keine Daten vorhanden.";
    return;
  }

  populatePeriods(history);
  restoreSystemFilter();
  const selectedPeriod = dashboardEls.periodSelect
    ? dashboardEls.periodSelect.value
    : "Alle";
  const selectedSystem = getSelectedSystems();
  const latestEntries = getLatestBySystemAndPeriod(
    history,
    selectedPeriod,
    selectedSystem
  );
  if (!latestEntries.length) {
    dashboardEls.revenue.textContent = "-";
    dashboardEls.fees.textContent = "-";
    dashboardEls.tax.textContent = "-";
    dashboardEls.payout.textContent = "-";
    dashboardEls.revenueTrend.textContent = "Vormonat: -";
    dashboardEls.feesTrend.textContent = "Vormonat: -";
    dashboardEls.taxTrend.textContent = "Vormonat: -";
    dashboardEls.payoutTrend.textContent = "Vormonat: -";
    dashboardEls.updated.textContent = "Keine Daten f\u00fcr die Auswahl.";
    return;
  }
  const totals = sumMetrics(latestEntries);

  const prevTotals = getPreviousMonthTotals(
    history,
    latestEntries,
    selectedPeriod,
    selectedSystem
  );

  const trend = {
    revenue: `Vormonat: ${formatDeltaCurrency(totals.revenue, prevTotals && prevTotals.revenue)}`,
    fees: `Vormonat: ${formatDeltaCurrency(totals.fees, prevTotals && prevTotals.fees)}`,
    tax: `Vormonat: ${formatDeltaCurrency(totals.tax, prevTotals && prevTotals.tax)}`,
    payout: `Vormonat: ${formatDeltaCurrency(totals.payout, prevTotals && prevTotals.payout)}`
  };

  dashboardEls.revenue.textContent = formatCurrency(totals.revenue);
  dashboardEls.fees.textContent = formatCurrency(totals.fees);
  dashboardEls.tax.textContent = formatCurrency(totals.tax);
  dashboardEls.payout.textContent = formatCurrency(totals.payout);
  dashboardEls.revenueTrend.textContent = trend.revenue;
  dashboardEls.feesTrend.textContent = trend.fees;
  dashboardEls.taxTrend.textContent = trend.tax;
  dashboardEls.payoutTrend.textContent = trend.payout;

  const latestTs = Math.max(...latestEntries.map((entry) => entry.ts));
  const date = new Date(latestTs);
  const periodText =
    selectedPeriod && selectedPeriod !== "Alle" ? ` | Zeitraum: ${selectedPeriod}` : "";
  const systemLabelText =
    Array.isArray(selectedSystem) && !selectedSystem.includes("Alle")
      ? selectedSystem.map(systemLabel).join(", ")
      : "";
  const systemText = systemLabelText ? ` | Portal: ${systemLabelText}` : "";
  dashboardEls.updated.textContent = `Zuletzt aktualisiert: ${date.toLocaleDateString("de-DE")} ${date.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" })}${periodText}${systemText}`;
}

if (dashboardEls.periodSelect) {
  dashboardEls.periodSelect.addEventListener("change", () => {
    localStorage.setItem(PERIOD_KEY, dashboardEls.periodSelect.value);
    renderDashboard();
  });
}

if (dashboardEls.systemOptions) {
  dashboardEls.systemOptions.addEventListener("change", (event) => {
    const target = event.target;
    if (!target || target.type !== "checkbox") return;
    const checkboxes = getSystemCheckboxes();
    if (target.value === "Alle" && target.checked) {
      checkboxes.forEach((box) => {
        if (box.value !== "Alle") {
          box.checked = false;
        }
      });
    } else if (target.value !== "Alle" && target.checked) {
      checkboxes.forEach((box) => {
        if (box.value === "Alle") {
          box.checked = false;
        }
      });
    }
    const selected = getSelectedSystems();
    localStorage.setItem(SYSTEM_KEY, JSON.stringify(selected));
    renderDashboard();
  });
}

if (dashboardEls.toggleBlurBtn && dashboardEls.kpiGrid) {
  dashboardEls.toggleBlurBtn.addEventListener("click", () => {
    const isBlurred = dashboardEls.kpiGrid.classList.toggle("blurred");
    dashboardEls.toggleBlurBtn.setAttribute("aria-pressed", String(isBlurred));
  });
}

function openFilter() {
  if (!dashboardEls.filterOverlay) return;
  dashboardEls.filterOverlay.classList.add("show");
  dashboardEls.filterOverlay.setAttribute("aria-hidden", "false");
}

function closeFilter() {
  if (!dashboardEls.filterOverlay) return;
  dashboardEls.filterOverlay.classList.remove("show");
  dashboardEls.filterOverlay.setAttribute("aria-hidden", "true");
}

if (dashboardEls.openFilterBtn) {
  dashboardEls.openFilterBtn.addEventListener("click", openFilter);
}

if (dashboardEls.closeFilterBtn) {
  dashboardEls.closeFilterBtn.addEventListener("click", closeFilter);
}

if (dashboardEls.closeFilterBtnFooter) {
  dashboardEls.closeFilterBtnFooter.addEventListener("click", closeFilter);
}

if (dashboardEls.filterOverlay) {
  dashboardEls.filterOverlay.addEventListener("click", (event) => {
    if (event.target === dashboardEls.filterOverlay) {
      closeFilter();
    }
  });
}

function systemLabel(key) {
  if (key === "airbnb") return "Airbnb";
  if (key === "booking") return "Booking";
  if (key === "direkt") return "Direktbuchungen";
  if (key === "miete") return "Mieteingangskontrolle";
  return key || "-";
}

function renderSnapshotModal() {
  if (!dashboardEls.snapshotListModal) return;
  const history = readDashboardHistory();
  populateSnapshotFilters(history);
  const period = dashboardEls.snapshotPeriodSelect
    ? dashboardEls.snapshotPeriodSelect.value
    : "Alle";
  const system = dashboardEls.snapshotSystemSelect
    ? dashboardEls.snapshotSystemSelect.value
    : "Alle";
  const filtered = history.filter((entry) => {
    const entryPeriod = normalizePeriod(entry.metrics && entry.metrics.period);
    if (period !== "Alle" && entryPeriod !== period) return false;
    if (system !== "Alle" && entry.system !== system) return false;
    return true;
  });
  dashboardEls.snapshotListModal.innerHTML = "";
  filtered
    .slice()
    .sort((a, b) => b.ts - a.ts)
    .forEach((entry) => {
      const row = document.createElement("div");
      row.className = "snapshot-row";
      const rawPeriod = entry.metrics && entry.metrics.period;
      const periodLabel = normalizePeriod(rawPeriod);
      const date = new Date(entry.ts);
      row.innerHTML =
        `<span>${periodLabel || "-"}</span>` +
        `<span class="muted">${systemLabel(entry.system)}</span>` +
        `<span class="muted">${date.toLocaleDateString("de-DE")} ${date.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" })}</span>` +
        `<button type="button" class="ghost">L\u00f6schen</button>`;
      const btn = row.querySelector("button");
      btn.addEventListener("click", async () => {
        const ok = window.confirm("Diesen Snapshot wirklich l\u00f6schen?");
        if (!ok) return;
        const idx = dashboardHistory.findIndex(
          (item) =>
            item &&
            item.ts === entry.ts &&
            item.system === entry.system &&
            (item.metrics ? item.metrics.period : undefined) === rawPeriod
        );
        if (idx !== -1) {
          dashboardHistory.splice(idx, 1);
          await saveDashboardHistory();
        }
        renderDashboard();
        renderSnapshotModal();
      });
      dashboardEls.snapshotListModal.appendChild(row);
    });
}

function openSnapshots() {
  if (!dashboardEls.snapshotsOverlay) return;
  renderSnapshotModal();
  dashboardEls.snapshotsOverlay.classList.add("show");
  dashboardEls.snapshotsOverlay.setAttribute("aria-hidden", "false");
}

function closeSnapshots() {
  if (!dashboardEls.snapshotsOverlay) return;
  dashboardEls.snapshotsOverlay.classList.remove("show");
  dashboardEls.snapshotsOverlay.setAttribute("aria-hidden", "true");
}

if (dashboardEls.openSnapshotsBtn) {
  dashboardEls.openSnapshotsBtn.addEventListener("click", openSnapshots);
}

if (dashboardEls.closeSnapshotsBtn) {
  dashboardEls.closeSnapshotsBtn.addEventListener("click", closeSnapshots);
}

if (dashboardEls.closeSnapshotsBtnFooter) {
  dashboardEls.closeSnapshotsBtnFooter.addEventListener("click", closeSnapshots);
}

if (dashboardEls.snapshotsOverlay) {
  dashboardEls.snapshotsOverlay.addEventListener("click", (event) => {
    if (event.target === dashboardEls.snapshotsOverlay) {
      closeSnapshots();
    }
  });
}

if (dashboardEls.snapshotPeriodSelect) {
  dashboardEls.snapshotPeriodSelect.addEventListener("change", renderSnapshotModal);
}

if (dashboardEls.snapshotSystemSelect) {
  dashboardEls.snapshotSystemSelect.addEventListener("change", renderSnapshotModal);
}

if (menuToggle && menuPanel) {
  menuToggle.addEventListener("click", () => {
    const isOpen = menuPanel.classList.toggle("show");
    menuToggle.setAttribute("aria-expanded", isOpen ? "true" : "false");
  });
}

if (menuClose && menuPanel) {
  menuClose.addEventListener("click", () => {
    menuPanel.classList.remove("show");
    menuToggle.setAttribute("aria-expanded", "false");
  });
}

document.addEventListener("click", (event) => {
  if (!menuPanel || !menuToggle) {
    return;
  }
  const target = event.target;
  if (
    menuPanel.classList.contains("show") &&
    !menuPanel.contains(target) &&
    !menuToggle.contains(target)
  ) {
    menuPanel.classList.remove("show");
    menuToggle.setAttribute("aria-expanded", "false");
  }
});

const logo = document.querySelector(".logo");
if (logo) {
  logo.addEventListener("click", () => {
    setTarget("");
  });
}

if (auth && onAuthStateChanged) {
  onAuthStateChanged(auth, async (user) => {
    currentUser = user;
    if (!user) {
      currentProfile = null;
      setAdminVisible(false);
      closeAdmin();
      dashboardHistory = [];
      renderDashboard();
      lock();
      return;
    }
    const profile = await ensureUserProfile(user, usernameFromEmail(user.email));
    if (profile && profile.role === "disabled") {
      loginError.textContent = "Konto deaktiviert.";
      await signOut(auth);
      return;
    }
    currentProfile = profile;
    setAdminVisible(profile && profile.role === "admin");
    unlock();
    await loadDashboardHistory();
    renderDashboard();
    loadUsers();
  });
} else {
  lock();
}
