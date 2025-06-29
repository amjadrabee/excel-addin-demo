import { initializeApp, deleteApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { loginUser } from "../firebase-auth.js";

/* ── fetch full Firebase config (via temp app) ── */
async function fetchFirebaseConfig() {
  const tmp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const cfg = await getDoc(doc(getFirestore(tmp), "config", "firebase"))
                    .then(s => { if (!s.exists()) throw new Error("Config missing"); return s.data(); });
  await deleteApp(tmp);
  return cfg;
}

/* ── fetch redirect URL from Firestore ── */
async function fetchTaskpaneUrl() {
  const snap = await getDoc(doc(getFirestore(), "config", "urls"));
  if (!snap.exists() || !snap.data().taskpane) throw new Error("Taskpane URL missing in Firestore");
  return snap.data().taskpane;
}

/* ── main login handler ── */
async function handleLogin(email, password) {
  const status = document.getElementById("status");
  try {
    status.textContent = "🔄 Loading config…";
    const cfg = await fetchFirebaseConfig();

    if (getApps().length === 0) initializeApp(cfg);   // set default app

    status.textContent = "🔐 Signing in…";
    if (!(await loginUser(email, password))) return;

    const redirectUrl = await fetchTaskpaneUrl();
    window.location.href = redirectUrl;               // 🎯 go to taskpane
  } catch (e) {
    console.error(e);
    status.textContent = "❌ " + e.message;
  }
}

/* ── button click ── */
document.getElementById("loginBtn").onclick = () => {
  const email    = document.getElementById("emailInput").value.trim();
  const password = document.getElementById("passwordInput").value.trim();
  if (!email || !password) {
    document.getElementById("status").textContent = "❌ Enter both fields.";
    return;
  }
  handleLogin(email, password);
};
