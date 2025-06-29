// firebase-auth.js  –  single‑session + persistent login
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut,
  setPersistence,
  browserLocalPersistence,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getApps,
  initializeApp
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc,
  setDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

/* ─── Ensure some Firebase app exists (projectId‑only placeholder) ─── */
if (getApps().length === 0) {
  initializeApp({ projectId: "excel-addin-auth" });  // temp; real config comes later
}
const auth = getAuth();
const db   = getFirestore();

/* ─── Login user, enforce single session, persist credentials ─── */
export async function loginUser(email, password) {
  const status = document.getElementById("status") || { textContent: "" };

  try {
    await setPersistence(auth, browserLocalPersistence);   // stay signed‑in
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid  = cred.user.uid;

    // single‑session check
    const sessRef  = doc(db, "sessions", uid);
    const sessSnap = await getDoc(sessRef);
    if (sessSnap.exists()) {
      await signOut(auth);
      status.textContent = "❌ Account already active elsewhere.";
      return false;
    }

    // create session
    const sessionId = crypto.randomUUID();
    await setDoc(sessRef, { sessionId, timestamp: Date.now() });

    // save locally
    localStorage.setItem("uid", uid);
    localStorage.setItem("sessionId", sessionId);
    localStorage.setItem("email", email);

    status.textContent = "✅ Logged in!";
    return true;
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Login failed.";
    return false;
  }
}

/* ─── Validate stored session ─── */
export async function isSessionValid() {
  const uid       = localStorage.getItem("uid");
  const sessionId = localStorage.getItem("sessionId");
  if (!uid || !sessionId) return false;

  // wait for auth to settle
  await new Promise(res => onAuthStateChanged(auth, () => res()));

  const user = auth.currentUser;
  if (!user || user.uid !== uid) return false;

  const snap = await getDoc(doc(db, "sessions", uid));
  return snap.exists() && snap.data().sessionId === sessionId;
}

/* ─── Client‑side logout helper ─── */
export async function logoutRequestLocal() {
  localStorage.removeItem("uid");
  localStorage.removeItem("sessionId");
  localStorage.removeItem("email");
  await signOut(auth);
}
