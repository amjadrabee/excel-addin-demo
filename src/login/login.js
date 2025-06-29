import { initializeApp, deleteApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { loginUser } from "../firebase-auth.js";

// Step 1: Load config from Firestore
async function fetchFirebaseConfig() {
  // Prevent duplicate app errors
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg-login");
  const tempDb = getFirestore(tempApp);

  const snap = await getDoc(doc(tempDb, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found.");
  const config = snap.data();

  await deleteApp(tempApp);
  return config;
}

// Step 2: Handle login
async function handleLogin(email, password) {
  const status = document.getElementById("status");
  status.textContent = "🔄 Loading Firebase config...";

  try {
    const config = await fetchFirebaseConfig();

    if (getApps().length === 0) {
      initializeApp(config);
    }

    status.textContent = "🔐 Logging in...";
    const ok = await loginUser(email, password);
    if (!ok) return;

    // Store email for future use
    localStorage.setItem("email", email);

    // Redirect to actual UI
    window.location.href = "../ui/taskpane.html";
  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + (err.message || "Login failed.");
  }
}

// Step 3: Hook up the button
document.addEventListener("DOMContentLoaded", () => {
  const loginBtn = document.getElementById("loginBtn");
  if (!loginBtn) {
    console.error("❌ Login button not found!");
    return;
  }

  loginBtn.onclick = () => {
    const email = document.getElementById("emailInput").value.trim();
    const password = document.getElementById("passwordInput").value.trim();
    const status = document.getElementById("status");

    if (!email || !password) {
      status.textContent = "❌ Enter both email and password.";
      return;
    }

    handleLogin(email, password);
  };
});
