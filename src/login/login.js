import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

import { loginUser } from "../firebase-auth.js";

// Load full config from Firestore and return it
async function fetchFirebaseConfig() {
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const tempDb = getFirestore(tempApp);

  const configSnap = await getDoc(doc(tempDb, "config", "firebase"));
  if (!configSnap.exists()) {
    await deleteApp(tempApp);
    throw new Error("❌ Firebase config missing.");
  }

  const config = configSnap.data();
  await deleteApp(tempApp);
  return config;
}

// Full login flow
async function handleLogin(email, password) {
  const statusEl = document.getElementById("status");

  try {
    statusEl.textContent = "🔄 Loading config...";
    const config = await fetchFirebaseConfig();

    // Init default app with real config
    initializeApp(config);

    statusEl.textContent = "🔐 Signing in...";
    const result = await loginUser(email, password);
    if (!result) return;

    // Save email for future reference (e.g., logout request)
    localStorage.setItem("email", email);

    // Redirect to taskpane UI (do NOT fetch UI from Firestore here)
    window.location.href = "/src/ui/taskpane.html";
  } catch (err) {
    console.error(err);
    statusEl.textContent = "❌ " + (err.message || "Login failed.");
  }
}

// Handle login button
document.getElementById("loginBtn").onclick = () => {
  const email = document.getElementById("emailInput").value.trim();
  const password = document.getElementById("passwordInput").value.trim();

  const statusEl = document.getElementById("status");
  if (!email || !password) {
    statusEl.textContent = "❌ Enter both fields.";
    return;
  }

  statusEl.textContent = "🔐 Logging in...";
  handleLogin(email, password);
};
