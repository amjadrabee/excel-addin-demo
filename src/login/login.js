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

import { loginUser } from "../firebase-auth.js";

// 🔧 Load Firebase config from Firestore (temporary app)
async function fetchFirebaseConfig() {
  const tmpApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const tmpDb = getFirestore(tmpApp);

  const snap = await getDoc(doc(tmpDb, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found in Firestore.");

  const config = snap.data();
  await deleteApp(tmpApp);
  return config;
}

// 🔐 Handle login logic
async function handleLogin(email, password) {
  const status = document.getElementById("status");
  try {
    status.textContent = "🔄 Preparing Firebase…";

    const config = await fetchFirebaseConfig();

    // 🧠 Safe Firebase init (avoid app conflict)
    const apps = getApps();
    if (apps.length === 0) {
      initializeApp(config); // Default unnamed app
    } else {
      const currentApp = getApp();
      const currentOptions = currentApp.options;
      if (JSON.stringify(currentOptions) !== JSON.stringify(config)) {
        throw new Error("❌ Firebase already initialized with different config.");
      }
    }

    status.textContent = "🔐 Signing in...";
    const ok = await loginUser(email, password);
    if (!ok) return;

    localStorage.setItem("email", email);

    // 🚀 Redirect to taskpane
    window.location.href = "../ui/taskpane.html";

  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + err.message;
  }
}

// 🎯 Bind login button
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
  }
});
