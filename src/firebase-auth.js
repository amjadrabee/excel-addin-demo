// firebase-auth.js  —  secure single‑session support
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

// ─────────────────────────────────────────────────────
// 0. Firebase >> basic (public) config  • projectId only
//    Full config is fetched at runtime in login.js
// ─────────────────────────────────────────────────────
const basicApp  = initializeApp({ projectId: "excel-addin-auth" });
const auth      = getAuth(basicApp);
const db        = getFirestore(basicApp);

// ─────────────────────────────────────────────────────
// 1. Login user (creates session if none exists)
// ─────────────────────────────────────────────────────
export async function loginUser(email, password) {
  const statusEl = document.getElementById("login-status") || { textContent: "" };

  try {
    // 1‑a. Auth
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid  = cred.user.uid;

    // 1‑b. Check existing session
    const sessRef   = doc(db, "sessions", uid);
    const sessSnap  = await getDoc(sessRef);
    const existing  = sessSnap.exists() ? sessSnap.data().sessionId : null;

    if (existing) {
      // Someone already logged in on another device
      await signOut(auth);
      statusEl.textContent = "❌ Account already active on another device.";
      throw new Error("Active session exists");
    }

    // 1‑c. Create a new session
    const sessionId = crypto.randomUUID();
    await setDoc(sessRef, {
      sessionId,
      timestamp: Date.now()
    });

    localStorage.setItem("uid",        uid);
    localStorage.setItem("sessionId",  sessionId);

    statusEl.textContent = "✅ Logged in!";
    return true;
  } catch (err) {
    console.error(err);
    if (!statusEl.textContent) statusEl.textContent = "❌ Login failed.";
    return false;
  }
}

// ─────────────────────────────────────────────────────
// 2. Logout request (email sent elsewhere – handled in taskpane.js)
//    Here we ONLY remove local state; admin removes Firestore doc.
// ─────────────────────────────────────────────────────
export async function logoutRequestLocal() {
  localStorage.removeItem("uid");
  localStorage.removeItem("sessionId");
  await signOut(auth);            // firebase sign‑out (optional)
}

// ─────────────────────────────────────────────────────
// 3. Session validity helper  (true ⇢ still same device)
// ─────────────────────────────────────────────────────
export async function isSessionValid() {
  const uid       = localStorage.getItem("uid");
  const sessionId = localStorage.getItem("sessionId");
  if (!uid || !sessionId) return false;

  try {
    const snap = await getDoc(doc(db, "sessions", uid));
    return snap.exists() && snap.data().sessionId === sessionId;
  } catch {
    return false;
  }
}
