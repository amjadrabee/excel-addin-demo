import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// Load config from Firestore and redirect to actual taskpane
async function start() {
  try {
    // Temp app just to read Firestore config
    const tempApp = initializeApp(
      { projectId: "excel-addin-auth" }, // Only projectId is needed here
      "temp"
    );
    const db = getFirestore(tempApp);

    // Fetch firebase config from Firestore
    const configSnap = await getDoc(doc(db, "config", "firebase"));
    if (!configSnap.exists()) throw new Error("❌ Firebase config not found.");
    const firebaseConfig = configSnap.data();

    // Delete temp app
    await deleteApp(tempApp);

    // Initialize actual app
    const realApp = initializeApp(firebaseConfig);
    const realDb = getFirestore(realApp);

    // Load taskpane URL
    const urlSnap = await getDoc(doc(realDb, "config", "urls"));
    if (!urlSnap.exists()) throw new Error("❌ Taskpane URL not found.");
    const taskpaneUrl = urlSnap.data().taskpane;

    // Redirect to the real UI
    window.location.href = taskpaneUrl;
  } catch (err) {
    document.body.innerHTML = `<h3 style="color:red">Startup failed:</h3><pre>${err.message}</pre>`;
    console.error("Startup error:", err);
  }
}

start();
