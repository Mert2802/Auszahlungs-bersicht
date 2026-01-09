const frame = document.getElementById("appFrame");
const buttons = document.querySelectorAll("button[data-target]");
const welcome = document.getElementById("welcome");
const menuToggle = document.getElementById("menuToggle");
const menuPanel = document.getElementById("menuPanel");
const menuClose = document.getElementById("menuClose");
const DASHBOARD_KEY = "dashboard_snapshots_v1";
const PERIOD_KEY = "dashboard_period";
const SYSTEM_KEY = "dashboard_system";
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
  systemSelect: document.getElementById("systemSelect"),
  clearSnapshotsBtn: document.getElementById("clearSnapshotsBtn"),
  toggleBlurBtn: document.getElementById("toggleBlurBtn"),
  kpiGrid: document.getElementById("kpiGrid")
};
const loginGate = document.getElementById("loginGate");
const loginForm = document.getElementById("loginForm");
const loginId = document.getElementById("loginId");
const loginPw = document.getElementById("loginPw");
const loginError = document.getElementById("loginError");
const logoutBtn = document.getElementById("logoutBtn");

const targets = {
  airbnb: "../Airbnb to XLS/index.html",
  booking: "../Booking_Auszahlungsuebersicht_HTML/index.html",
  direkt: "../Direktbuchungen/index.html"
};

const AUTH_KEY = "hub_auth";

function setTarget(key) {
  frame.src = targets[key] || "";
  frame.style.display = key ? "block" : "none";
  if (welcome) {
    welcome.style.display = key ? "none" : "flex";
  }
  if (menuToggle) {
    menuToggle.classList.toggle("visible", Boolean(key));
  }
  if (menuPanel) {
    menuPanel.classList.remove("show");
  }
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

function checkStoredLogin() {
  return sessionStorage.getItem(AUTH_KEY) === "ok";
}

function logout() {
  sessionStorage.removeItem(AUTH_KEY);
  loginError.textContent = "";
  lock();
}

loginForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const id = loginId.value.trim();
  const pw = loginPw.value.trim();
  if (id === "admin" && pw === "admin123") {
    sessionStorage.setItem(AUTH_KEY, "ok");
    loginError.textContent = "";
    unlock();
  } else {
    loginError.textContent = "ID oder Passwort ist falsch.";
  }
});

if (logoutBtn) {
  logoutBtn.addEventListener("click", logout);
}

if (checkStoredLogin()) {
  unlock();
} else {
  lock();
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
  try {
    const raw = localStorage.getItem(DASHBOARD_KEY);
    const list = raw ? JSON.parse(raw) : [];
    return Array.isArray(list) ? list : [];
  } catch (error) {
    return [];
  }
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
    ["marz", "Maerz"],
    ["maerz", "Maerz"],
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

function getLatestBySystem(history) {
  const latest = new Map();
  history.forEach((entry) => {
    if (!entry || !entry.system || !entry.metrics) return;
    const existing = latest.get(entry.system);
    if (!existing || entry.ts > existing.ts) {
      latest.set(entry.system, entry);
    }
  });
  return Array.from(latest.values());
}

function getLatestBySystemAndPeriod(history, period, systemFilter) {
  const latest = new Map();
  history.forEach((entry) => {
    if (!entry || !entry.system || !entry.metrics) return;
    const entryPeriod = normalizePeriod(entry.metrics.period);
    if (period && period !== "Alle" && entryPeriod !== period) return;
    if (systemFilter && systemFilter !== "Alle" && entry.system !== systemFilter) return;
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

function getTotalsBefore(history, cutoffTs, period, systemFilter) {
  const latestBefore = new Map();
  history.forEach((entry) => {
    if (!entry || !entry.system || !entry.metrics) return;
    if (entry.ts >= cutoffTs) return;
    const entryPeriod = normalizePeriod(entry.metrics.period);
    if (period && period !== "Alle" && entryPeriod !== period) return;
    if (systemFilter && systemFilter !== "Alle" && entry.system !== systemFilter) return;
    const existing = latestBefore.get(entry.system);
    if (!existing || entry.ts > existing.ts) {
      latestBefore.set(entry.system, entry);
    }
  });
  return sumMetrics(Array.from(latestBefore.values()));
}

function formatTrend(current, previous) {
  if (!previous || previous <= 0) return "-";
  const diff = ((current - previous) / previous) * 100;
  const sign = diff > 0 ? "+" : "";
  return `${sign}${diff.toFixed(1)}%`;
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

function restoreSystemFilter() {
  if (!dashboardEls.systemSelect) return;
  const stored = localStorage.getItem(SYSTEM_KEY);
  if (stored) {
    dashboardEls.systemSelect.value = stored;
  } else {
    dashboardEls.systemSelect.value = "Alle";
  }
}

function renderDashboard() {
  if (!dashboardEls.revenue) return;
  const history = readDashboardHistory();
  if (!history.length) {
    dashboardEls.updated.textContent = "Keine Daten vorhanden.";
    return;
  }

  populatePeriods(history);
  restoreSystemFilter();
  const selectedPeriod = dashboardEls.periodSelect
    ? dashboardEls.periodSelect.value
    : "Alle";
  const selectedSystem = dashboardEls.systemSelect
    ? dashboardEls.systemSelect.value
    : "Alle";
  const latestEntries = getLatestBySystemAndPeriod(
    history,
    selectedPeriod,
    selectedSystem
  );
  const totals = sumMetrics(latestEntries);

  const now = Date.now();
  const prev7 = getTotalsBefore(
    history,
    now - 7 * 24 * 60 * 60 * 1000,
    selectedPeriod,
    selectedSystem
  );
  const prev30 = getTotalsBefore(
    history,
    now - 30 * 24 * 60 * 60 * 1000,
    selectedPeriod,
    selectedSystem
  );

  const trend = {
    revenue: `7T: ${formatTrend(totals.revenue, prev7.revenue)} · 30T: ${formatTrend(totals.revenue, prev30.revenue)}`,
    fees: `7T: ${formatTrend(totals.fees, prev7.fees)} · 30T: ${formatTrend(totals.fees, prev30.fees)}`,
    tax: `7T: ${formatTrend(totals.tax, prev7.tax)} · 30T: ${formatTrend(totals.tax, prev30.tax)}`,
    payout: `7T: ${formatTrend(totals.payout, prev7.payout)} · 30T: ${formatTrend(totals.payout, prev30.payout)}`
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
    selectedPeriod && selectedPeriod !== "Alle" ? ` · Zeitraum: ${selectedPeriod}` : "";
  const systemText =
    selectedSystem && selectedSystem !== "Alle" ? ` · Portal: ${selectedSystem}` : "";
  dashboardEls.updated.textContent = `Zuletzt aktualisiert: ${date.toLocaleDateString("de-DE")} ${date.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" })}${periodText}${systemText}`;
}

renderDashboard();

if (dashboardEls.periodSelect) {
  dashboardEls.periodSelect.addEventListener("change", () => {
    localStorage.setItem(PERIOD_KEY, dashboardEls.periodSelect.value);
    renderDashboard();
  });
}

if (dashboardEls.systemSelect) {
  dashboardEls.systemSelect.addEventListener("change", () => {
    localStorage.setItem(SYSTEM_KEY, dashboardEls.systemSelect.value);
    renderDashboard();
  });
}

if (dashboardEls.clearSnapshotsBtn) {
  dashboardEls.clearSnapshotsBtn.addEventListener("click", () => {
    const ok = window.confirm("Alle Snapshots wirklich loeschen?");
    if (!ok) return;
    localStorage.removeItem(DASHBOARD_KEY);
    renderDashboard();
  });
}

if (dashboardEls.toggleBlurBtn && dashboardEls.kpiGrid) {
  dashboardEls.toggleBlurBtn.addEventListener("click", () => {
    const isBlurred = dashboardEls.kpiGrid.classList.toggle("blurred");
    dashboardEls.toggleBlurBtn.setAttribute("aria-pressed", String(isBlurred));
  });
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
