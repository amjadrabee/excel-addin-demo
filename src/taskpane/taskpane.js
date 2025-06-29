import {
  initializeApp,
  deleteApp
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";

import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

let defaultAppInitialized = false;

async function ensureDefaultApp() {
  if (defaultAppInitialized) return;

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

    const { isSessionValid } = await import("../firebase-auth.js");
    const valid = await isSessionValid();

    if (!valid) {
      statusBox.innerHTML = `
        <div style="color: red; font-weight: bold;">
          🔒 Session expired – reload Add‑in.
        </div>`;
      return;
    }

    // Load UI from Firestore
    const { getFirestore, doc, getDoc } = await import("https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js");
    const db = getFirestore();
    const uiSnap = await getDoc(doc(db, "config", "ui"));
    if (!uiSnap.exists()) throw new Error("❌ UI HTML not found in Firestore.");
    statusBox.innerHTML = uiSnap.data().html;

    // Attach handlers after UI is injected
    const { convertToPDF, handleLogoutRequest } = await import("./handlers.js");
    document.getElementById("convertBtn").onclick = convertToPDF;
    document.getElementById("requestLogoutBtn").onclick = handleLogoutRequest;

  } catch (err) {
    console.error(err);
    statusBox.innerHTML = `<pre style="color: red;">${err.message}</pre>`;
  }
});
