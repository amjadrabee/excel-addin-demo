import { loginUser } from "../firebase-auth.js";
import {
  initializeApp,
  deleteApp
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// Fetch full Firebase config from Firestore and init default app once.
let defaultAppInitialised = false;
async function ensureDefaultApp() {
  if (defaultAppInitialised) return;

  const tmp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const tmpDb = getFirestore(tmp);
  const cfgSnap = await getDoc(doc(tmpDb, "config", "firebase"));
  if (!cfgSnap.exists()) throw new Error("❌ Firebase config missing in Firestore");
  const fullCfg = cfgSnap.data();

  await deleteApp(tmp);
  initializeApp(fullCfg);
  defaultAppInitialised = true;
}

// Handle login button click (DOM ready)
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("loginBtn");
  if (!btn) return console.error("loginBtn not found in DOM");

  btn.onclick = async () => {
    const email = document.getElementById("emailInput").value.trim();
    const password = document.getElementById("passwordInput").value.trim();
    const statusEl = document.getElementById("status") || document.getElementById("login-status");

    if (!email || !password) {
      statusEl.textContent = "❌ Enter both fields.";
      return;
    }

    statusEl.textContent = "🔐 Initialising Firebase…";
    try {
      await ensureDefaultApp();
      statusEl.textContent = "🔐 Logging in…";

      const ok = await loginUser(email, password);
      if (!ok) return;

      statusEl.textContent = "✅ Login successful. Redirecting…";
      window.location.href = "../ui/taskpane.html";
    } catch (err) {
      console.error(err);
      statusEl.textContent = "❌ " + err.message;
    }
  };
});
