// firebase-auth.js
import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
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

// ───── Initialize basic app (projectId only) ─────
const basicApp = initializeApp({ projectId: "excel-addin-auth" });
const auth = getAuth(basicApp);
const db = getFirestore(basicApp);

// ───── Login with single-session enforcement ─────
export async function loginUser(email, password) {
  const statusEl = document.getElementById("login-status") || { textContent: "" };

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

    statusEl.textContent = "✅ Logged in!";
    return true;
  } catch (err) {
    console.error(err);
    if (!statusEl.textContent) statusEl.textContent = "❌ Login failed.";
    return false;
  }
}

// ───── Local logout (used after request logout) ─────
export async function logoutRequestLocal() {
  localStorage.removeItem("uid");
  localStorage.removeItem("sessionId");
  await signOut(auth);
}

// ───── Check if the session is still valid ─────
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
