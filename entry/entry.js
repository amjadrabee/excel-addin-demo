import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

async function start() {
  document.body.innerHTML = `<p>Loading add-in...</p>`;

  try {
    const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "temp");
    const db = getFirestore(tempApp);

    const configSnap = await getDoc(doc(db, "config", "firebase"));
    if (!configSnap.exists()) throw new Error("❌ Firebase config not found.");
    const firebaseConfig = configSnap.data();

    await deleteApp(tempApp);

    const realApp = initializeApp(firebaseConfig);
    const realDb = getFirestore(realApp);

    const urlSnap = await getDoc(doc(realDb, "config", "urls"));
    if (!urlSnap.exists()) throw new Error("❌ Taskpane URL not found.");
    const taskpaneUrl = urlSnap.data().taskpane;

    window.location.href = taskpaneUrl;
  } catch (err) {
    document.body.innerHTML = `<h3 style="color:red">Startup failed:</h3><pre>${err.message}</pre>`;
    console.error("Startup error:", err);
  }
}

start();
