import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// ─────────────────────────────────────────────
// 1. Load Firebase config (stored in Firestore)
// ─────────────────────────────────────────────
async function loadFirebaseConfig() {
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "temp");
  const db = getFirestore(tempApp);

  const snap = await getDoc(doc(db, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found.");

  const config = snap.data();
  await deleteApp(tempApp);
  return config;
}

// ─────────────────────────────────────────────
// 2. Handle login and inject UI HTML
// ─────────────────────────────────────────────
async function loginUser(email, password) {
  try {
    // 2‑a. Initialize Firebase with remote config
    const firebaseConfig = await loadFirebaseConfig();
    const app  = initializeApp(firebaseConfig, "main");
    const auth = getAuth(app);
    const db   = getFirestore(app);

    // 2‑b. Sign in
    await signInWithEmailAndPassword(auth, email, password);

    // 2‑c. Fetch UI HTML from Firestore
    const uiSnap = await getDoc(doc(db, "config", "ui"));
    if (!uiSnap.exists()) throw new Error("❌ UI HTML not found in Firestore.");

    // 2‑d. Render the UI inside the current (empty) taskpane
    const html = uiSnap.data().html;
    document.open();
    document.write(html);
    document.close();
  } catch (err) {
    console.error(err);
    document.getElementById("status").textContent = "❌ " + err.message;
  }
}

// ─────────────────────────────────────────────
// 3. Wire up the login button
// ─────────────────────────────────────────────
document.getElementById("loginBtn").onclick = () => {
  const email    = document.getElementById("emailInput").value.trim();
  const password = document.getElementById("passwordInput").value.trim();

  if (!email || !password) {
    document.getElementById("status").textContent = "❌ Enter both email and password.";
    return;
  }

  document.getElementById("status").textContent = "🔐 Logging in…";
  loginUser(email, password);
};
