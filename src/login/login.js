import { initializeApp, deleteApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { loginUser } from "../firebase-auth.js";

// Step 1: Load config from Firestore via temporary app
async function fetchFirebaseConfig() {
  // Prevent multiple initializations of tmp app
  if (getApps().some(app => app.name === "tmpCfg")) {
    await deleteApp(getApps().find(app => app.name === "tmpCfg"));
  }

  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const tempDb = getFirestore(tempApp);

  const snap = await getDoc(doc(tempDb, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found in Firestore.");

  const config = snap.data();
  await deleteApp(tempApp);
  return config;
}

// Step 2: Handle login
async function handleLogin(email, password) {
  const status = document.getElementById("status");
  status.textContent = "🔄 Preparing...";

  try {
    const config = await fetchFirebaseConfig();

    if (getApps().length === 0) {
      initializeApp(config);
    }

    status.textContent = "🔐 Logging in...";
    const ok = await loginUser(email, password);
    if (!ok) return;

    // Store email locally
    localStorage.setItem("email", email);

    // Redirect to main UI
    window.location.href = "../ui/taskpane.html";
  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + (err.message || "Login failed.");
  }
}

// Step 3: Attach login button event
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("loginBtn");
  if (btn) {
    btn.onclick = () => {
      const email = document.getElementById("emailInput").value.trim();
      const password = document.getElementById("passwordInput").value.trim();
      const status = document.getElementById("status");

      if (!email || !password) {
        status.textContent = "❌ Enter both fields.";
        return;
      }

      handleLogin(email, password);
    };
  } else {
    console.error("❌ Login button not found.");
  }
});
