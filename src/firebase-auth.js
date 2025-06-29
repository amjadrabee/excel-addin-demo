import { initializeApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc,
  deleteDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// Initialize only if needed
if (getApps().length === 0) {
  initializeApp({ projectId: "excel-addin-auth" }); // Basic for early calls
}

const auth = getAuth();
const db = getFirestore();

// ─────────────────────────────────────────────────────────────
// 🔐 LOGIN USER (single session enforced)
// ─────────────────────────────────────────────────────────────
export async function loginUser(email, password) {
  const status = document.getElementById("status") || { textContent: "" };

  try {
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid = cred.user.uid;

    const sessRef = doc(db, "sessions", uid);
    const snap = await getDoc(sessRef);

    if (snap.exists()) {
      const existing = snap.data().sessionId;
      if (existing && existing !== localStorage.getItem("sessionId")) {
        await signOut(auth);
        status.textContent = "❌ You're already signed in on another device.";
        return false;
      }
    }

    const sessionId = crypto.randomUUID();

    // Save session to Firestore
    await setDoc(sessRef, {
      sessionId,
      timestamp: Date.now()
    });

    // Save locally
    localStorage.setItem("uid", uid);
    localStorage.setItem("email", email);
    localStorage.setItem("sessionId", sessionId);

    status.textContent = "✅ Login successful";
    return true;
  } catch (err) {
    console.error("Login failed:", err);
    status.textContent = "❌ Login failed.";
    return false;
  }
}

// ─────────────────────────────────────────────────────────────
// 🔒 CHECK SESSION VALIDITY
// ─────────────────────────────────────────────────────────────
export async function isSessionValid() {
  const uid = localStorage.getItem("uid");
  const sessionId = localStorage.getItem("sessionId");

  if (!uid || !sessionId) return false;

  try {
    const ref = doc(db, "sessions", uid);
    const snap = await getDoc(ref);
    return snap.exists() && snap.data().sessionId === sessionId;
  } catch {
    return false;
  }
}

// ─────────────────────────────────────────────────────────────
// 🔓 LOCAL LOGOUT ONLY (after logout request sent)
// ─────────────────────────────────────────────────────────────
export async function logoutRequestLocal() {
  localStorage.removeItem("uid");
  localStorage.removeItem("sessionId");
  localStorage.removeItem("email");

  try {
    await signOut(auth);
  } catch (err) {
    console.warn("Sign out warning:", err.message);
  }
}
