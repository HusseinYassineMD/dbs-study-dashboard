/** Simple username + password gate for the static dashboard (session-based). */

const AUTH_SESSION_KEY = "dbs_dashboard_auth_v1";
const SESSION_MS = 7 * 24 * 60 * 60 * 1000;

const DEFAULT_AUTH = {
  username: "dbs-operations",
  passwordHash: "a85c9ab08e5d4218b4b8699aa575844d57b96d3e41a5f29b2a939db7e4bda19d",
};

function authConfig() {
  return window.DBS_AUTH || DEFAULT_AUTH;
}

async function sha256Hex(text) {
  const data = new TextEncoder().encode(text);
  const hash = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(hash))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

function readSession() {
  try {
    const raw = sessionStorage.getItem(AUTH_SESSION_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed?.exp || Date.now() > parsed.exp) {
      sessionStorage.removeItem(AUTH_SESSION_KEY);
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

function writeSession(username) {
  sessionStorage.setItem(
    AUTH_SESSION_KEY,
    JSON.stringify({ username, exp: Date.now() + SESSION_MS })
  );
}

function clearSession() {
  sessionStorage.removeItem(AUTH_SESSION_KEY);
}

function setAppVisible(visible) {
  const gate = document.getElementById("login-gate");
  const app = document.getElementById("app-root");
  if (gate) gate.hidden = visible;
  if (app) app.hidden = !visible;
}

function loadDashboardScripts() {
  if (window.__dbsDashboardScriptsLoaded) return Promise.resolve();
  window.__dbsDashboardScriptsLoaded = true;
  const v = "20261015b";
  const urls = [`report.js?v=${v}`, `app.js?v=${v}`];
  return urls.reduce(
    (chain, src) =>
      chain.then(
        () =>
          new Promise((resolve, reject) => {
            const s = document.createElement("script");
            s.src = src;
            s.onload = resolve;
            s.onerror = () => reject(new Error(`Failed to load ${src}`));
            document.body.appendChild(s);
          })
      ),
    Promise.resolve()
  );
}

async function attemptLogin(username, password) {
  const cfg = authConfig();
  const user = String(username || "").trim();
  const hash = await sha256Hex(String(password || ""));
  if (user === cfg.username && hash === cfg.passwordHash) {
    writeSession(user);
    return true;
  }
  return false;
}

function bindLoginForm() {
  const form = document.getElementById("login-form");
  const errorEl = document.getElementById("login-error");
  if (!form) return;

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (errorEl) errorEl.textContent = "";
    const user = document.getElementById("login-username")?.value;
    const pass = document.getElementById("login-password")?.value;
    const ok = await attemptLogin(user, pass);
    if (!ok) {
      if (errorEl) errorEl.textContent = "Incorrect username or password.";
      return;
    }
    await enterDashboard();
  });
}

async function enterDashboard() {
  setAppVisible(true);
  await loadDashboardScripts();
}

function bindSignOut() {
  const btn = document.getElementById("sign-out-btn");
  if (!btn) return;
  btn.addEventListener("click", () => {
    clearSession();
    window.location.reload();
  });
}

async function initAuth() {
  bindLoginForm();
  bindSignOut();
  if (readSession()) {
    await enterDashboard();
    return;
  }
  setAppVisible(false);
}

document.addEventListener("DOMContentLoaded", () => {
  initAuth().catch((err) => {
    console.error(err);
    const errorEl = document.getElementById("login-error");
    if (errorEl) errorEl.textContent = "Login failed to initialize.";
  });
});

window.dbsSignOut = () => {
  clearSession();
  window.location.reload();
};
