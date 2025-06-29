// login.js  —  uses loginUser() above, then injects UI HTML
import { loginUser } from "../firebase-auth.js";
import {
  initializeApp,
  deleteApp
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// ───────────────────────────────────────────────────
// 1. Fetch full Firebase config from Firestore
// ───────────────────────────────────────────────────
async function fetchFullConfig() {
  const tmpApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const db     = getFirestore(tmpApp);
  const snap   = await getDoc(doc(db, "config", "firebase"));
  if (!snap.exists()) throw new Error("Firebase config missing");
  await deleteApp(tmpApp);
  return snap.data();
}

// ───────────────────────────────────────────────────
// 2. Perform login, then inject UI HTML
// ───────────────────────────────────────────────────
async function handleLogin(email, password) {
  try {
    // Full config (needed only for Firestore reads after login)
    await fetchFullConfig();

    const ok = await loginUser(email, password);
    if (!ok) return;

    // Load UI HTML
    const { getFirestore, doc, getDoc } = await import(
      "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js"
    );
    const db = getFirestore();
    const uiSnap = await getDoc(doc(db, "config", "ui"));
    if (!uiSnap.exists()) throw new Error("UI HTML not found.");

    document.open();
    document.write(uiSnap.data().html);
    document.close();
  } catch (err) {
    console.error(err);
    document.getElementById("status").textContent = "❌ " + err.message;
  }
}

// ───────────────────────────────────────────────────
// 3. Wire the login form
// ───────────────────────────────────────────────────
document.getElementById("loginBtn").onclick = () => {
  const email    = document.getElementById("emailInput").value.trim();
  const password = document.getElementById("passwordInput").value.trim();
  if (!email || !password) {
    document.getElementById("status").textContent = "❌ Enter both fields.";
    return;
  }
  document.getElementById("status").textContent = "🔐 Logging in…";
  handleLogin(email, password);
};
