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

// Fetch Firebase config from Firestore using a temporary app
async function fetchFirebaseConfig() {
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const tempDb = getFirestore(tempApp);

  const snap = await getDoc(doc(tempDb, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found.");

  const config = snap.data();
  await deleteApp(tempApp);
  return config;
}

// Handle login process
async function handleLogin(email, password) {
  const status = document.getElementById("status");
  status.textContent = "🔄 Preparing login...";

  try {
    const config = await fetchFirebaseConfig();
    console.log("✅ Firebase config loaded");

    if (getApps().length === 0) {
      initializeApp(config);
      console.log("✅ Firebase default app initialized");
    }

    status.textContent = "🔐 Signing in...";
    const ok = await loginUser(email, password);
    if (!ok) return;

    localStorage.setItem("email", email); // Store email for later use

    // Redirect to the taskpane UI
    window.location.href = "../ui/taskpane.html";
  } catch (err) {
    console.error("❌ Login error:", err);
    status.textContent = "❌ " + (err.message || "Login failed.");
  }
}

// Attach to login button
document.getElementById("loginBtn").onclick = () => {
  const email = document.getElementById("emailInput").value.trim();
  const password = document.getElementById("passwordInput").value.trim();
  const status = document.getElementById("status");

  if (!email || !password) {
    status.textContent = "❌ Enter both email and password.";
    return;
  }

  handleLogin(email, password);
};
