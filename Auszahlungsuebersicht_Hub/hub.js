const frame = document.getElementById("appFrame");
const buttons = document.querySelectorAll("button[data-target]");
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
  buttons.forEach((btn) => {
    btn.classList.toggle("primary", btn.dataset.target === key);
    btn.classList.toggle("secondary", btn.dataset.target !== key);
  });
}

buttons.forEach((btn) => {
  btn.addEventListener("click", () => {
    setTarget(btn.dataset.target);
  });
});

function unlock() {
  document.body.classList.remove("locked");
  loginGate.style.display = "none";
}

function lock() {
  document.body.classList.add("locked");
  loginGate.style.display = "flex";
  frame.src = "";
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
    setTarget("airbnb");
  } else {
    loginError.textContent = "ID oder Passwort ist falsch.";
  }
});

if (logoutBtn) {
  logoutBtn.addEventListener("click", logout);
}

if (checkStoredLogin()) {
  unlock();
  setTarget("airbnb");
} else {
  lock();
}
