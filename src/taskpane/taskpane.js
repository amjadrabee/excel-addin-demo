import {
  initializeApp,
  deleteApp
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";

import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

import { getAuth } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";

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

async function waitForAuthReady() {
  const auth = getAuth();
  return new Promise(resolve => {
    const unsub = auth.onAuthStateChanged(() => {
      unsub();
      resolve();
    });
  });
}

Office.onReady(async () => {
  const statusBox = document.getElementById("app");
  try {
    await ensureDefaultApp();
    await waitForAuthReady(); // Wait for currentUser to be available

    const { isSessionValid, logoutRequestLocal } = await import("../firebase-auth.js");
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

    // Setup buttons after HTML is loaded
    document.getElementById("convertBtn").onclick = convertToPDF;
    document.getElementById("requestLogoutBtn").onclick = async () => {
      const user = localStorage.getItem("uid") || "Unknown";
      try {
        await fetch("https://your-logout-email-service.com/send", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            to: "support@yourcompany.com",
            subject: "Logout Request",
            message: `${user} has requested to log out from the Excel Add-in.`
          })
        });
        await logoutRequestLocal();
        alert("📩 Logout request sent. Please close this window.");
      } catch (err) {
        console.error(err);
        alert("❌ Failed to send logout request.");
      }
    };
  } catch (err) {
    console.error(err);
    statusBox.innerHTML = `<pre style="color: red;">${err.message}</pre>`;
  }
});
