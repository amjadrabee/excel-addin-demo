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

  // Temporary app just with projectId to read config
  const tmp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const tmpDb = getFirestore(tmp);
  const cfgSnap = await getDoc(doc(tmpDb, "config", "firebase"));
  if (!cfgSnap.exists()) throw new Error("❌ Firebase config missing in Firestore");
  const fullCfg = cfgSnap.data();

  await deleteApp(tmp);
  initializeApp(fullCfg); // default (unnamed) app — now getAuth()/getFirestore() work globally
  defaultAppInitialised = true;
}

// Inject UI HTML (stored in /config/ui) after successful login
async function injectUI() {
  const { getFirestore, doc, getDoc } = await import(
    "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js"
  );
  const db = getFirestore();
  const uiSnap = await getDoc(doc(db, "config", "ui"));
  if (!uiSnap.exists()) throw new Error("❌ UI HTML not found in Firestore");

  document.open();
  document.write(uiSnap.data().html);
  document.close();
}

// Handle login button click (DOM ready)
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("loginBtn");
  if (!btn) return console.error("loginBtn not found in DOM");

  btn.onclick = async () => {
    const email    = document.getElementById("emailInput").value.trim();
    const password = document.getElementById("passwordInput").value.trim();
    const statusEl = document.getElementById("status");

    if (!email || !password) {
      statusEl.textContent = "❌ Enter both fields.";
      return;
    }

    statusEl.textContent = "🔐 Initialising Firebase…";
    try {
      await ensureDefaultApp();
      statusEl.textContent = "🔐 Logging in…";

      const ok = await loginUser(email, password);
      if (!ok) return; // error message already shown by loginUser

      statusEl.textContent = "✅ Login successful. Loading add‑in UI…";
      await injectUI();
    } catch (err) {
      console.error(err);
      statusEl.textContent = "❌ " + err.message;
    }
  };
});
