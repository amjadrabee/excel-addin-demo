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

import { loginUser, initAuthAndDb } from "../firebase-auth.js";

async function fetchFirebaseConfig() {
  const tmpApp = initializeApp({ projectId: "excel-addin-auth" }, "tmp-login");
  const tmpDb = getFirestore(tmpApp);

  const snap = await getDoc(doc(tmpDb, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found.");
  const cfg = snap.data();

  await deleteApp(tmpApp);
  return cfg;
}

async function handleLogin(email, password) {
  const status = document.getElementById("status");

  try {
    status.textContent = "🔄 Loading config…";

    const cfg = await fetchFirebaseConfig();
    const app = initializeApp(cfg);

    // Setup auth and db in our module
    initAuthAndDb(app);

    status.textContent = "🔐 Signing in...";
    const ok = await loginUser(email, password);
    if (!ok) return;

    localStorage.setItem("email", email);

    const urlSnap = await getDoc(doc(getFirestore(), "config", "urls"));
    const redirectUrl = urlSnap.data()?.taskpane;
    if (!redirectUrl) throw new Error("❌ UI redirect URL missing.");
    window.location.href = redirectUrl;
  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + err.message;
  }
}

document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("loginBtn");
  btn?.addEventListener("click", () => {
    const email = document.getElementById("emailInput").value.trim();
    const password = document.getElementById("passwordInput").value.trim();
    const status = document.getElementById("status");

    if (!email || !password) {
      status.textContent = "❌ Enter both fields.";
      return;
    }

    handleLogin(email, password);
  });
});
