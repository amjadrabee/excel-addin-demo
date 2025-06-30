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

let auth;
let db;

export function initAuthAndDb(app) {
  auth = getAuth(app);
  db = getFirestore(app);
}

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
    await setDoc(sessRef, {
      sessionId,
      timestamp: Date.now()
    });

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
