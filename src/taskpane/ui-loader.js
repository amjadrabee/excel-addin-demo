import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

async function loadTaskpaneUI() {
  try {
    /* fetch Firebase config with a temp‑app */
    const temp = initializeApp({ projectId: "excel-addin-auth" }, "tmpUI");
    const cfg  = await getDoc(doc(getFirestore(temp), "config", "firebase"));
    if (!cfg.exists()) throw new Error("❌ Firebase config missing.");
    await deleteApp(temp);

    /* default app */
    initializeApp(cfg.data());
    const db = getFirestore();

    /* fetch UI HTML */
    const uiSnap = await getDoc(doc(db, "config", "ui"));
    if (!uiSnap.exists()) throw new Error("❌ UI HTML not found.");
    document.getElementById("app").innerHTML = uiSnap.data().html;

    /* notify taskpane.js that HTML is in the DOM */
    window.dispatchEvent(new Event("TaskpaneUILoaded"));
  } catch (err) {
    console.error(err);
    document.getElementById("app").innerHTML =
      `<div style="color:red;">${err.message}</div>`;
  }
}

loadTaskpaneUI();
