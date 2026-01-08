const frame = document.getElementById("appFrame");
const buttons = document.querySelectorAll("button[data-target]");
const welcome = document.getElementById("welcome");
const menuToggle = document.getElementById("menuToggle");
const menuPanel = document.getElementById("menuPanel");
const menuClose = document.getElementById("menuClose");
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
