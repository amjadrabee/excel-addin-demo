import { initializeApp, deleteApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { getAuth, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import { isSessionValid } from "../firebase-auth.js";
import { convertToPDF, handleLogoutRequest } from "./handlers.js";

async function initDefaultApp() {
  if (getApps().length) return;                       // already initialised
  const tmp = initializeApp({ projectId: "excel-addin-auth" }, "tmpTaskpane");
  const cfg = await getDoc(doc(getFirestore(tmp), "config", "firebase"))
                     .then(s => { if (!s.exists()) throw new Error("No config"); return s.data(); });
  await deleteApp(tmp);
  initializeApp(cfg);
}

async function loadUIHtml() {
  const snap = await getDoc(doc(getFirestore(), "config", "ui"));
  if (!snap.exists()) throw new Error("UI HTML missing");
  return snap.data().html;
}

Office.onReady(async () => {
  const box = document.getElementById("app");
  try {
    await initDefaultApp();

    const auth = getAuth();
    const sessionOk = await isSessionValid();

    onAuthStateChanged(auth, async user => {
      if (!user || !sessionOk) {
        box.innerHTML = `<div style="color:red;font-weight:bold">🔒 Session expired – log in again.</div>`;
        return;
      }

      /* store email (if missed) */
      localStorage.setItem("email", user.email);

      /* inject UI */
      box.innerHTML = await loadUIHtml();

      /* wire buttons */
      document.getElementById("convertBtn")      ?.addEventListener("click", convertToPDF);
      document.getElementById("requestLogoutBtn")?.addEventListener("click", handleLogoutRequest);
    });
  } catch (e) {
    console.error(e);
    box.innerHTML = `<pre style="color:red">${e.message}</pre>`;
  }
});
