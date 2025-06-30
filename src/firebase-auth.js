import {
  initializeApp,
  getApps,
  getApp
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";

import {
  getAuth,
  signInWithEmailAndPassword,
  signOut
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";

import {
  getFirestore,
  doc,
  getDoc,
  setDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// ─────────────────────────────────────────────
// ✅ Safe default app initialization (once only)
// ─────────────────────────────────────────────
function ensureFirebaseInitialized() {
  const apps = getApps();
  if (apps.length === 0) {
    // Minimal fallback config (used early if full not loaded yet)
    initializeApp({ projectId: "excel-addin-auth" });
  }
}

// ⏪ Ensure initialized now
ensureFirebaseInitialized();

const auth = getAuth();
const db = getFirestore();

// ─────────────────────────────────────────────
// 🔐 Login user (single session enforced)
// ─────────────────────────────────────────────
export async function loginUser(email, password) {
  const status = document.getElementById("status") || { textContent: "" };

  try {
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid = cred.user.uid;

    const sessRef = doc(db, "sessions", uid);
    const sessSnap = await getDoc(sessRef);
    const existingSession = sessSnap.exists() ? sessSnap.data().sessionId : null;

    const localSessionId = localStorage.getItem("sessionId");

    if (existingSession && existingSession !== localSessionId) {
      await signOut(auth);
      status.textContent = "❌ You're already signed in on another device.";
      return false;
    }

    const sessionId = localSessionId || crypto.randomUUID();

    await setDoc(sessRef, {
      sessionId,
      timestamp: Date.now()
    });

    localStorage.setItem("uid", uid);
    localStorage.setItem("sessionId", sessionId);
    localStorage.setItem("email", email);

    status.textContent = "✅ Login successful";
    return true;
  } catch (err) {
    console.error("Login failed:", err);
    status.textContent = "❌ Login failed.";
    return false;
  }
}

// ─────────────────────────────────────────────
// 🔒 Check if current session is valid
// ─────────────────────────────────────────────
export async function isSessionValid() {
  const uid = localStorage.getItem("uid");
  const sessionId = localStorage.getItem("sessionId");
  if (!uid || !sessionId) return false;

  try {
    const sessSnap = await getDoc(doc(db, "sessions", uid));
    return sessSnap.exists() && sessSnap.data().sessionId === sessionId;
  } catch (err) {
    console.error("Session check failed:", err);
    return false;
  }
}

// ─────────────────────────────────────────────
// 🔓 Logout locally after request logout
// ─────────────────────────────────────────────
export async function logoutRequestLocal() {
  try {
    await signOut(auth);
  } catch (err) {
    console.warn("Sign out warning:", err.message);
  }

  localStorage.removeItem("uid");
  localStorage.removeItem("sessionId");
  localStorage.removeItem("email");
}
