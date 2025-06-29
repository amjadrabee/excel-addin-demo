import {
  initializeApp,
  deleteApp,
  getApps
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";

import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

import {
  getAuth,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";

import { isSessionValid, logoutRequestLocal } from "../firebase-auth.js";
import { setupHandlers } from "./handlers.js";

let defaultAppInitialized = false;

async function ensureDefaultApp() {
  if (defaultAppInitialized || getApps().length > 0) return;

  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tempTaskpane");
  const tempDb = getFirestore(tempApp);

  const snap = await getDoc(doc(tempDb, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found.");
  const fullConfig = snap.data();

  await deleteApp(tempApp);
  initializeApp(fullConfig);
  defaultAppInitialized = true;
}

Office.onReady(async () => {
  const statusBox = document.getElementById("app");
  try {
    await ensureDefaultApp();

    const auth = getAuth();
    const validSession = await isSessionValid();

    onAuthStateChanged(auth, async (user) => {
      if (!user || !validSession) {
        statusBox.innerHTML = `<div style="color:red; font-weight: bold;">🔒 Session expired – reload Add‑in.</div>`;
        return;
      }

      // Load UI from Firestore
      const db = getFirestore();
      const uiSnap = await getDoc(doc(db, "config", "ui"));
      if (!uiSnap.exists()) throw new Error("❌ UI HTML not found in Firestore.");

      statusBox.innerHTML = uiSnap.data().html;
      setupHandlers();
    });
  } catch (err) {
    console.error(err);
    statusBox.innerHTML = `<pre style="color: red;">${err.message}</pre>`;
  }
});
