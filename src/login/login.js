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

// Step 1: Load config from Firestore via temporary app
async function fetchFirebaseConfig() {
  const tmpApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const tmpDb = getFirestore(tmpApp);

  const snap = await getDoc(doc(tmpDb, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found in Firestore.");

  const config = snap.data();
  await deleteApp(tmpApp); // ✅ Correct order
  return config;
}

// Step 2: Attempt login
async function handleLogin(email, password) {
  const status = document.getElementById("status");
  try {
    status.textContent = "🔄 Preparing Firebase…";

    const config = await fetchFirebaseConfig();
    if (getApps().length === 0) {
      initializeApp(config); // Initialize main app
    }

    status.textContent = "🔐 Signing in...";
    const success = await loginUser(email, password);
    if (!success) return;

    localStorage.setItem("email", email);

    // 🔁 Option A: Hardcoded taskpane redirect (recommended)
    window.location.href = "../ui/taskpane.html";

    // 🔁 Option B (alt): If using Firestore to store taskpane.html URL:
    // const urlsDoc = await getDoc(doc(getFirestore(), "config", "urls"));
    // const redirectUrl = urlsDoc.data()?.taskpane;
    // if (!redirectUrl) throw new Error("❌ Missing taskpane URL in Firestore.");
    // window.location.href = redirectUrl;

  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + (err.message || "Login failed.");
  }
}

// Step 3: Setup click listener
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("loginBtn");
  if (btn) {
    btn.addEventListener("click", () => {
      const email = document.getElementById("emailInput").value.trim();
      const password = document.getElementById("passwordInput").value.trim();
      const status = document.getElementById("status");

      if (!email || !password) {
        status.textContent = "❌ Enter both fields.";
        return;
      }

      handleLogin(email, password);
    });
  } else {
    console.warn("⚠️ Login button not found.");
  }
});
