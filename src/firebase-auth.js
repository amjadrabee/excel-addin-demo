import {
  getAuth,
  signInWithEmailAndPassword,
  signOut
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// NOTE: The *default* Firebase app is initialised elsewhere (entry.js / login.js).
// Every call here simply grabs getAuth() / getFirestore() from that default app.

// ───── Login with single‑session guard ─────
export async function loginUser(email, password) {
  const statusEl = document.getElementById("login-status") || { textContent: "" };

  try {
    const auth = getAuth();              // default app auth (must already exist)
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid  = cred.user.uid;

    const db       = getFirestore();
    const sessRef  = doc(db, "sessions", uid);
    const sessSnap = await getDoc(sessRef);

    if (sessSnap.exists()) {
      // Someone already logged in somewhere else
      await signOut(auth);
      statusEl.textContent = "❌ Account is already active on another device.";
      return false;
    }

    // Create new session
    const sessionId = crypto.randomUUID();
    await setDoc(sessRef, { sessionId, timestamp: Date.now() });

    localStorage.setItem("uid", uid);
    localStorage.setItem("sessionId", sessionId);

    statusEl.textContent = "✅ Logged in!";
    return true;
  } catch (err) {
    console.error(err);
    statusEl.textContent = "❌ Login failed.";
    return false;
  }
}

// ───── Local logout (after request e‑mail) ─────
export async function logoutRequestLocal() {
  localStorage.removeItem("uid");
  localStorage.removeItem("sessionId");
  await signOut(getAuth());
}

// ───── Validate that the stored session is still active ─────
export async function isSessionValid() {
  const uid       = localStorage.getItem("uid");
  const sessionId = localStorage.getItem("sessionId");
  if (!uid || !sessionId) return false;

  try {
    const auth = getAuth();
    const db   = getFirestore();
    const currentUser = auth.currentUser;

    if (!currentUser || currentUser.uid !== uid) return false;

    const snap = await getDoc(doc(db, "sessions", uid));
    return snap.exists() && snap.data().sessionId === sessionId;
  } catch {
    return false;
  }
}
