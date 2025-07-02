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
//   setDoc,
//   updateDoc,
//   deleteDoc
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

// /* ────────────────────────────────────────────────
//    LOCAL cleanup after user requests logout
//    (taskpane.js sends the email, then calls this)  */
// export async function logoutRequestLocal() {
//   const uid = localStorage.getItem("uid");
//   if (!uid) return;

//   try {
//     await deleteDoc(doc(db, "sessions", uid));
//     localStorage.clear();
//     await signOut(auth);
//   } catch (err) {
//     console.error("Error in logoutRequestLocal:", err);
//   }
// }


/////////////////////////

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
  deleteDoc // <-- Added this import for deleting documents
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

let auth, db;

/* called once from login.js after default app is initialised */
export function initAuthAndDb(app) {
  auth = getAuth(app);
  db = getFirestore(app);
}

/* ────────────────────────────────────────────────
   LOGIN (returns true / false)
   ──────────────────────────────────────────────── */
export async function loginUser(email, password) {
  const status = document.getElementById("status") || {
    textContent: ""
  };

  try {
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid = cred.user.uid;

    const sessRef = doc(db, "sessions", uid); // Use uid as the document ID for sessions
    const snap = await getDoc(sessRef);
    const localId = localStorage.getItem("sessionId") || "";

    /* block sign‑in from second device */
    if (snap.exists() && snap.data().sessionId !== localId) {
      status.textContent = "❌ Already signed‑in on another device.";
      await signOut(auth); // Sign out the new attempted login
      return false;
    }

    /* create/reuse session */
    const sessionId = localId || crypto.randomUUID();
    await setDoc(sessRef, {
      sessionId,
      timestamp: Date.now(),
      email: email // Store email for easier lookup/debugging if needed
    });

    /* persist locally */
    localStorage.setItem("uid", uid);
    localStorage.setItem("email", email);
    localStorage.setItem("sessionId", sessionId);

    status.textContent = "✅ Login successful";
    return true;

  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + (err.code || "Login failed.");
    return false;
  }
}

/* session check used by taskpane.js */
export async function isSessionValid() {
  const uid = localStorage.getItem("uid");
  const id = localStorage.getItem("sessionId");
  // Ensure auth and db are initialized before trying to use them
  if (!uid || !id || !db || !auth) return false;

  try {
    const s = await getDoc(doc(db, "sessions", uid));
    return s.exists() && s.data().sessionId === id;
  } catch (err) {
    console.error("Error checking session validity:", err);
    return false;
  }
}

export async function logoutRequestLocal() {
  const uid = localStorage.getItem("uid"); // Get UID from local storage
  if (!uid) {
    // If no UID, nothing to clear in Firestore, just clear local storage and sign out.
    localStorage.clear();
    if (auth && auth.currentUser) {
      await signOut(auth);
    }
    console.log("No UID found, local data cleared and signed out.");
    return;
  }

  try {
    const sessionRef = doc(db, "sessions", uid); // Use UID as the document ID
    await deleteDoc(sessionRef); // Delete the session document from Firestore
    console.log("Firestore session document deleted for UID:", uid);

    localStorage.clear(); // Clear all local storage items related to the session
    console.log("Local storage cleared.");

    if (auth && auth.currentUser) {
      await signOut(auth); // Sign out from Firebase authentication
      console.log("Firebase user signed out.");
    }
  } catch (err) {
    console.error("Error in logoutRequestLocal:", err);
    // Even if Firestore deletion fails, attempt to clear local storage and sign out Firebase
    localStorage.clear();
    if (auth && auth.currentUser) {
      await signOut(auth);
    }
  }
}
