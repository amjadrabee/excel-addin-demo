import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  setPersistence,
  browserLocalPersistence,
  signOut
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

const app = initializeApp({ projectId: "excel-addin-auth" });
const auth = getAuth(app);
const db = getFirestore(app);

export async function loginUser(email, password) {
  const statusEl = document.getElementById("status") || { textContent: "" };

  try {
    await setPersistence(auth, browserLocalPersistence);
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid = cred.user.uid;

    const sessRef = doc(db, "sessions", uid);
    const snap = await getDoc(sessRef);
    const existing = snap.exists() ? snap.data().sessionId : null;

    if (existing) {
      await signOut(auth);
      statusEl.textContent = "❌ Account is already active on another device.";
      return false;
    }

    const sessionId = crypto.randomUUID();
    await setDoc(sessRef, {
      sessionId,
      timestamp: Date.now()
    });

    localStorage.setItem("uid", uid);
    localStorage.setItem("sessionId", sessionId);
    return true;
  } catch (err) {
    console.error(err);
    statusEl.textContent = "❌ Login failed.";
    return false;
  }
}

export async function logoutRequestLocal() {
  localStorage.removeItem("uid");
  localStorage.removeItem("sessionId");
  await signOut(auth);
}

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
