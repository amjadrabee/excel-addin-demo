import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { loginUser } from "../firebase-auth.js";

// Step 1: Load config from Firestore via temporary app
async function fetchFirebaseConfig() {
  let tempApp;
  try {
    tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
    const tempDb = getFirestore(tempApp);

    const snap = await getDoc(doc(tempDb, "config", "firebase"));
    if (!snap.exists()) throw new Error("❌ Firebase config not found in Firestore.");
    
    return snap.data();
  } finally {
    if (tempApp) await deleteApp(tempApp);
  }
}

// Step 2: Handle login
async function handleLogin(email, password) {
  const status = document.getElementById("status");
  status.textContent = "🔄 Preparing...";

  try {
    const config = await fetchFirebaseConfig();

    // Initialize default app (only one per page session)
    const app = initializeApp(config);

    status.textContent = "🔐 Logging in...";
    const ok = await loginUser(email, password);
    if (!ok) return;

    // Store email in localStorage
    localStorage.setItem("email", email);

    // Redirect to actual UI
    window.location.href = "/src/ui/taskpane.html";
  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + (err.message || "Login failed.");
  }
}

// Step 3: Hook login button
document.getElementById("loginBtn").onclick = () => {
  const email = document.getElementById("emailInput").value.trim();
  const password = document.getElementById("passwordInput").value.trim();

  const status = document.getElementById("status");
  if (!email || !password) {
    status.textContent = "❌ Enter both fields.";
    return;
  }

  handleLogin(email, password);
};
