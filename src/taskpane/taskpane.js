import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import { isSessionValid, logoutRequestLocal } from "../firebase-auth.js";
import { convertToPDF, handleLogoutRequest } from "./handlers.js";

let defaultAppInitialized = false;

async function ensureFirebaseInitialized() {
  if (defaultAppInitialized) return;

  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tempTaskpane");
  const tempDb = getFirestore(tempApp);

  const snap = await getDoc(doc(tempDb, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found.");

  const config = snap.data();
  await deleteApp(tempApp);

  initializeApp(config);
  defaultAppInitialized = true;
}

Office.onReady(async () => {
  const statusBox = document.getElementById("app");

  try {
    await ensureFirebaseInitialized();

    const valid = await isSessionValid();
    if (!valid) {
      statusBox.innerHTML = `
        <div style="color:red;font-weight:bold">
          🔒 Session expired – reload Add‑in and log in again.
        </div>`;
      return;
    }

    // Load UI HTML from Firestore
    const db = getFirestore();
    const uiSnap = await getDoc(doc(db, "config", "ui"));
    if (!uiSnap.exists()) throw new Error("❌ UI HTML not found in Firestore.");

    statusBox.innerHTML = uiSnap.data().html;

    // Ensure we store the user email for logout request
    const user = getAuth().currentUser;
    if (user && user.email) {
      localStorage.setItem("email", user.email);
    }

    // Hook up buttons
    document.getElementById("convertBtn").onclick = convertToPDF;
    document.getElementById("requestLogoutBtn").onclick = handleLogoutRequest;

  } catch (err) {
    console.error("Taskpane error:", err);
    statusBox.innerHTML = `<pre style="color:red">${err.message}</pre>`;
  }
});
