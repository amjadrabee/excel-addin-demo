import { initializeApp, deleteApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { loginUser } from "../firebase-auth.js";

// Step 1: Load config from Firestore via temporary app
async function fetchFirebaseConfig() {
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const tempDb = getFirestore(tempApp);

  const configSnap = await getDoc(doc(tempDb, "config", "firebase"));
  if (!configSnap.exists()) throw new Error("❌ Firebase config not found in Firestore.");
  const config = configSnap.data();

  await deleteApp(tempApp); // ✅ Cleanup after config fetch
  return config;
}

// Step 2: Load URLs from Firestore
async function fetchRedirectUrl(config) {
  const tempApp = initializeApp(config, "urlApp");
  const db = getFirestore(tempApp);

  const urlSnap = await getDoc(doc(db, "config", "urls"));
  if (!urlSnap.exists()) throw new Error("❌ URLs config not found.");
  const taskpaneUrl = urlSnap.data().taskpane;
  if (!taskpaneUrl) throw new Error("❌ 'taskpane' URL missing in Firestore.");

  await deleteApp(tempApp); // ✅ Clean up
  return taskpaneUrl;
}

// Step 3: Handle login and redirect
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

    // Store email
    localStorage.setItem("email", email);

    // Redirect to taskpane URL from Firestore
    const taskpaneUrl = await fetchRedirectUrl(config);
    window.location.href = taskpaneUrl;
  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + (err.message || "Login failed.");
  }
}

// Step 4: Hook login button
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
