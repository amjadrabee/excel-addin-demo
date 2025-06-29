import { initializeApp, deleteApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { loginUser } from "../firebase-auth.js";

// Load config from Firestore
async function fetchFirebaseConfig() {
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const tempDb = getFirestore(tempApp);
  const snap = await getDoc(doc(tempDb, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found.");
  const config = snap.data();
  await deleteApp(tempApp);
  return config;
}

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
    localStorage.setItem("email", email);
    window.location.href = "../ui/taskpane.html";
  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + (err.message || "Login failed.");
  }
}

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
