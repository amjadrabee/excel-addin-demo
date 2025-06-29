import { initializeApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc,
  deleteDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// Initialize app only once
if (getApps().length === 0) {
  initializeApp({ projectId: "excel-addin-auth" });
}

const auth = getAuth();
const db = getFirestore();

// ───── Login with single-session enforcement ─────
export async function loginUser(email, password) {
  const statusEl = document.getElementById("status") || { textContent: "" };

  try {
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid = cred.user.uid;

    const sessRef = doc(db, "sessions", uid);
    const sessSnap = await getDoc(sessRef);
    const existing = sessSnap.exists() ? sessSnap.data().sessionId : null;

    if (existing) {
      await signOut(auth);
      statusEl.textContent = "❌ Account is already active on another device.";
      throw new Error("Active session already exists.");
    }

    const sessionId = crypto.randomUUID();
    await setDoc(sessRef, {
      sessionId,
      timestamp: Date.now()
    });

    localStorage.setItem("uid", uid);
    localStorage.setItem("sessionId", sessionId);
    localStorage.setItem("email", email);

    statusEl.textContent = "✅ Logged in!";
    return true;
  } catch (err) {
    console.error(err);
    if (!statusEl.textContent) statusEl.textContent = "❌ Login failed.";
    return false;
  }
}

// ───── Check if session is valid ─────
export async function isSessionValid() {
  const uid = localStorage.getItem("uid");
  const sessionId = localStorage.getItem("sessionId");
  if (!uid || !sessionId) return false;

  try {
    const snap = await getDoc(doc(db, "sessions", uid));
    return snap.exists() && snap.data().sessionId === sessionId;
  } catch {
    return false;
  }
}

// ───── Local logout (client side only) ─────
export async function logoutRequestLocal() {
  localStorage.removeItem("uid");
  localStorage.removeItem("sessionId");
  localStorage.removeItem("email");
  await signOut(auth);
}

// ───── Export Firebase auth state listener ─────
export function onUserChanged(callback) {
  onAuthStateChanged(auth, callback);
}
