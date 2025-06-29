import {
  getAuth,
  signInWithEmailAndPassword,
  signOut,
  setPersistence,
  browserLocalPersistence
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";

import {
  getFirestore,
  doc,
  setDoc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

/*  loginUser(email, password)
    - Forces browserLocalPersistence so auth survives browser/app restarts
    - Enforces single‑device session via /sessions/{uid}
    - Saves uid, email, sessionId in localStorage
*/
export async function loginUser(email, password) {
  const statusEl =
    document.getElementById("login-status") ||
    document.getElementById("status")     ||
    { textContent: "" };

  try {
    const auth = getAuth();                    // default Firebase app
    await setPersistence(auth, browserLocalPersistence); // 🔐 persist login

    // Sign‑in
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid  = cred.user.uid;

    // Firestore handles single session
    const db       = getFirestore();
    const sessRef  = doc(db, "sessions", uid);
    const sessSnap = await getDoc(sessRef);

    if (sessSnap.exists()) {
      // Reject login if another device already holds the session
      await signOut(auth);
      statusEl.textContent = "❌ Account already active on another device.";
      return false;
    }

    // Create a new session
    const sessionId = crypto.randomUUID();
    await setDoc(sessRef, { sessionId, timestamp: Date.now() });

    // ✨ Store identifiers locally
    localStorage.setItem("uid",        uid);
    localStorage.setItem("email",      cred.user.email);
    localStorage.setItem("sessionId",  sessionId);

    statusEl.textContent = "✅ Logged in!";
    return true;
  } catch (err) {
    console.error(err);
    statusEl.textContent = "❌ Login failed.";
    return false;
  }
}

/*  logoutRequestLocal()
    - Clears local session storage (used after e‑mail logout request)
*/
export async function logoutRequestLocal() {
  localStorage.removeItem("uid");
  localStorage.removeItem("email");
  localStorage.removeItem("sessionId");
  await signOut(getAuth());
}

/*  isSessionValid()
    - Returns true if the local session matches the Firestore session
      and the Firebase Auth user is still signed‑in.
*/
export async function isSessionValid() {
  const uid       = localStorage.getItem("uid");
  const sessionId = localStorage.getItem("sessionId");
  if (!uid || !sessionId) return false;

  try {
    const auth = getAuth();
    const db   = getFirestore();

    // Wait for auth state to finish loading
    await new Promise(resolve => {
      const unsub = auth.onAuthStateChanged(() => { unsub(); resolve(); });
    });

    const curUser = auth.currentUser;
    if (!curUser || curUser.uid !== uid) return false;

    const snap = await getDoc(doc(db, "sessions", uid));
    return snap.exists() && snap.data().sessionId === sessionId;
  } catch {
    return false;
  }
}
