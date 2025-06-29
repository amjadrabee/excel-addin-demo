import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

async function loadTaskpaneUI() {
  try {
    // Temporary app just to get Firebase config
    const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tempUI");
    const tempDb = getFirestore(tempApp);

    const cfgSnap = await getDoc(doc(tempDb, "config", "firebase"));
    if (!cfgSnap.exists()) throw new Error("❌ Firebase config missing.");
    const firebaseConfig = cfgSnap.data();

    await deleteApp(tempApp);

    // Real default app
    const app = initializeApp(firebaseConfig);
    const db = getFirestore(app);

    // Load UI HTML from Firestore
    const uiSnap = await getDoc(doc(db, "config", "ui"));
    if (!uiSnap.exists()) throw new Error("❌ UI HTML not found in Firestore.");

    const html = uiSnap.data().html;
    document.getElementById("app").innerHTML = html;
  } catch (err) {
    console.error(err);
    document.getElementById("app").innerHTML = `<div style="color:red;">${err.message}</div>`;
  }
}

loadTaskpaneUI();
