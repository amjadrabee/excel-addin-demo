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

// ───── Load full Firebase config from Firestore ─────
async function fetchFullConfig() {
  const tmpApp = initializeApp({ projectId: "excel-addin-auth" }, "tmpCfg");
  const db = getFirestore(tmpApp);

  const snap = await getDoc(doc(db, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config missing.");

  await deleteApp(tmpApp);
  return snap.data();
}

// ───── Login then redirect to taskpane.html ─────
async function handleLogin(email, password) {
  try {
    document.getElementById("status").textContent = "🔐 Logging in…";

    await fetchFullConfig();  // ensures Firebase is initialized
    const ok = await loginUser(email, password);

    if (ok) {
      window.location.href = "/src/ui/taskpane.html"; // 🎯 redirect after login
    }
  } catch (err) {
    console.error(err);
    document.getElementById("status").textContent = "❌ " + err.message;
  }
}

// ───── Login Button Event ─────
document.getElementById("loginBtn").onclick = () => {
  const email = document.getElementById("emailInput").value.trim();
  const password = document.getElementById("passwordInput").value.trim();

  if (!email || !password) {
    document.getElementById("status").textContent = "❌ Enter both fields.";
    return;
  }

  handleLogin(email, password);
};
