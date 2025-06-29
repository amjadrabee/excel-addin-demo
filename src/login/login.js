// login.js  –  loads Firebase config from Firestore, logs user in, redirects
import {
  initializeApp,
  deleteApp,
  getApps
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { loginUser } from "../firebase-auth.js";

/* ── fetch Firebase config via temp‑app ── */
async function fetchFirebaseConfig() {
  // ensure any stale tmp app from previous load is gone
  const old = getApps().find(a => a.name === "tmpCfg-login");
  if (old) await deleteApp(old);

  const temp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg-login");
  const cfg  = await getDoc(doc(getFirestore(temp), "config", "firebase"))
                    .then(s => { if (!s.exists()) throw new Error("Config doc missing"); return s.data(); });
  await deleteApp(temp);
  return cfg;
}

/* ── full login flow ── */
async function handleLogin(email, password) {
  const status = document.getElementById("status");
  try {
    status.textContent = "🔄 Loading config…";
    const cfg = await fetchFirebaseConfig();

    if (getApps().length === 0) initializeApp(cfg);   // init main app

    status.textContent = "🔐 Signing in…";
    const ok = await loginUser(email, password);
    if (!ok) return;                                  // loginUser already set message

    // keep email for logout‑request
    localStorage.setItem("email", email);

    // redirect to Taskpane UI (URL stored in Firestore)
    const urlSnap = await getDoc(doc(getFirestore(), "config", "urls"));
    if (!urlSnap.exists() || !urlSnap.data().taskpane) {
      throw new Error("taskpane URL not found in Firestore");
    }
    window.location.href = urlSnap.data().taskpane;   // 🎯 go!
  } catch (err) {
    console.error("Login:", err);
    status.textContent = "❌ " + err.message;
  }
}

/* ── wire button after DOM ready ── */
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("loginBtn");
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
