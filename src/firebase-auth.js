
/*  firebase-auth.js
    — single‑session support (one device at a time)           */

import {
  getAuth,
  signInWithEmailAndPassword,
  signOut
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  updateDoc,
  deleteDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

let auth, db;

/*  called once from login.js after default app is initialised */
export function initAuthAndDb(app) {
  auth = getAuth(app);
  db   = getFirestore(app);
}

/* ────────────────────────────────────────────────
   LOGIN  (returns true  / false)                  */
export async function loginUser(email, password) {
  const status = document.getElementById("status") || { textContent: "" };

  try {
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid  = cred.user.uid;

    const sessRef = doc(db, "sessions", uid);
    const snap    = await getDoc(sessRef);
    const localId = localStorage.getItem("sessionId") || "";

    /* block sign‑in from second device */
    if (snap.exists() && snap.data().sessionId !== localId) {
      status.textContent = "❌ Already signed‑in on another device.";
      await signOut(auth);
      return false;
    }

    /* create/reuse session */
    const sessionId = localId || crypto.randomUUID();
    await setDoc(sessRef, { sessionId, timestamp: Date.now() });

    /* persist locally */
    localStorage.setItem("uid",        uid);
    localStorage.setItem("email",      email);
    localStorage.setItem("sessionId",  sessionId);

    status.textContent = "✅ Login successful";
    return true;

  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + (err.code || "Login failed.");
    return false;
  }
}

/*  session check used by taskpane.js */
export async function isSessionValid() {
  const uid = localStorage.getItem("uid");
  const id  = localStorage.getItem("sessionId");
  if (!uid || !id || !db) return false;

  try {
    const s = await getDoc(doc(db, "sessions", uid));
    return s.exists() && s.data().sessionId === id;
  } catch {
    return false;
  }
}

/* ────────────────────────────────────────────────
   LOCAL cleanup after user requests logout
   (taskpane.js sends the email, then calls this)  */
export async function logoutRequestLocal() {
  const uid = localStorage.getItem("uid");
  if (!uid) return;

  try {
    await deleteDoc(doc(db, "sessions", uid));
    localStorage.clear();
    await signOut(auth);
  } catch (err) {
    console.error("Error in logoutRequestLocal:", err);
  }
}


/////////////////////////

// /*  firebase-auth.js
//     — single‑session support (one device at a time)           */

// import {
//   getAuth,
//   signInWithEmailAndPassword,
//   signOut
// } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
// import {
//   getFirestore,
//   doc,
//   getDoc,
//   setDoc
// } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// let auth, db;

// /*  called once from login.js after default app is initialised */
// export function initAuthAndDb(app) {
//   auth = getAuth(app);
//   db   = getFirestore(app);
// }

// /* ────────────────────────────────────────────────
//    LOGIN  (returns true  / false)                  */
// export async function loginUser(email, password) {
//   const status = document.getElementById("status") || { textContent: "" };

//   try {
//     const cred = await signInWithEmailAndPassword(auth, email, password);
//     const uid  = cred.user.uid;

//     const sessRef = doc(db, "sessions", uid);
//     const snap    = await getDoc(sessRef);
//     const localId = localStorage.getItem("sessionId") || "";

//     /* block sign‑in from second device */
//     if (snap.exists() && snap.data().sessionId !== localId) {
//       status.textContent = "❌ Already signed‑in on another device.";
//       await signOut(auth);
//       return false;
//     }

//     /* create/reuse session */
//     const sessionId = localId || crypto.randomUUID();
//     await setDoc(sessRef, { sessionId, timestamp: Date.now() });

//     /* persist locally */
//     localStorage.setItem("uid",        uid);
//     localStorage.setItem("email",      email);
//     localStorage.setItem("sessionId",  sessionId);

//     status.textContent = "✅ Login successful";
//     return true;

//   } catch (err) {
//     console.error("Login error:", err);
//     status.textContent = "❌ " + (err.code || "Login failed.");
//     return false;
//   }
// }

// /*  session check used by taskpane.js */
// export async function isSessionValid() {
//   const uid = localStorage.getItem("uid");
//   const id  = localStorage.getItem("sessionId");
//   if (!uid || !id || !db) return false;

//   try {
//     const s = await getDoc(doc(db, "sessions", uid));
//     return s.exists() && s.data().sessionId === id;
//   } catch {
//     return false;
//   }
// }

// export async function logoutRequestLocal() {
//   const user = auth.currentUser;
//   if (!user) return;

//   try {
//     const sessionRef = doc(db, "sessions", user.email);
//     await updateDoc(sessionRef, { sessionId: null });
//     localStorage.clear();
//     await signOut(auth);
//   } catch (err) {
//     console.error("Error in logoutRequestLocal:", err);
//   }
// }

