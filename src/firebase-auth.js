import {
  initializeApp,
  deleteApp,
  getApps,
  getApp
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

import { loginUser, initAuthAndDb } from "../firebase-auth.js";

/* ─── Load full Firebase config from Firestore (via temp app) ─── */
async function fetchFirebaseConfig() {
  // create / reuse a named temp app
  const tmpName = "tmp-login";
  const oldTmp = getApps().find(a => a.name === tmpName);
  if (oldTmp) await deleteApp(oldTmp);

  const tmp = initializeApp({ projectId: "excel-addin-auth" }, tmpName);
  const cfgSnap = await getDoc(doc(getFirestore(tmp), "config", "firebase"));
  if (!cfgSnap.exists()) throw new Error("❌ Firebase config missing in Firestore");
  const cfg = cfgSnap.data();
  await deleteApp(tmp);
  return cfg;
}

/* ─── Main login handler ─── */
async function handleLogin(email, password) {
  const status = document.getElementById("status");

  try {
    status.textContent = "🔄 Loading Firebase config…";
    const cfg = await fetchFirebaseConfig();

    // ⚡ Safe default‑app initialisation
    let app;
    if (getApps().length === 0) {
      app = initializeApp(cfg);                   // first time
    } else {
      app = getApp();                             // already initialised
      // (optional) sanity‑check: configs must match projectId
      if (app.options.projectId !== cfg.projectId) {
        throw new Error("❌ Firebase already initialised with a different project.");
      }
    }

    // Attach auth/db to helper module
    initAuthAndDb(app);

    status.textContent = "🔐 Signing in…";
    const ok = await loginUser(email, password);  // single‑session enforced
    if (!ok) return;                              // error message set inside

    // Save email for logout mail
    localStorage.setItem("email", email);

    // 🔗 Get taskpane URL from Firestore, then redirect
    const urlsSnap = await getDoc(doc(getFirestore(), "config", "urls"));
    const redirectUrl = urlsSnap.data()?.taskpane;
    if (!redirectUrl) throw new Error("❌ 'taskpane' URL missing in Firestore.");
    window.location.href = redirectUrl;

  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + err.message;
  }
}

/* ─── Wire the button once DOM is ready ─── */
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("loginBtn");
  if (!btn) {
    console.error("⚠️ Login button not found in DOM");
    return;
  }
  btn.addEventListener("click", () => {
    const email    = document.getElementById("emailInput").value.trim();
    const password = document.getElementById("passwordInput").value.trim();
    const status   = document.getElementById("status");

    if (!email || !password) {
      status.textContent = "❌ Enter both email and password.";
      return;
    }
    handleLogin(email, password);
  });
});
