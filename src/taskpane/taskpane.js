/*  taskpane.js – attaches listeners AFTER ui‑loader injects the HTML  */
import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import { isSessionValid, logoutRequestLocal } from "../firebase-auth.js";

let appReady = false;

/* make sure default app exists (one‑time) */
async function ensureDefault() {
  if (appReady) return;
  const tmp   = initializeApp({ projectId: "excel-addin-auth" }, "tmpTaskpane");
  const cfg   = await getDoc(doc(getFirestore(tmp), "config", "firebase"));
  if (!cfg.exists()) throw new Error("Firebase config doc missing");
  await deleteApp(tmp);
  initializeApp(cfg.data());
  appReady = true;
}

/* handle Convert button */
async function convertToPDF() {
  const status = document.getElementById("status");
  const file   = document.getElementById("uploadDocx").files[0];
  if (!file) { status.textContent = "❌ Select a .docx file."; return; }

  try {
    await ensureDefault();
    if (!(await isSessionValid())) throw new Error("Session expired");

    status.textContent = "🔑 Reading API key…";
    const db = getFirestore();
    const keySnap = await getDoc(doc(db, "config", "cloudconvert"));
    if (!keySnap.exists()) throw new Error("CloudConvert key not set in Firestore");
    /* field can be key  | apiKey – support either */
    const apiKey = keySnap.data().key || keySnap.data().apiKey;
    if (!apiKey) throw new Error("CloudConvert key is empty");

    status.textContent = "⬆ Uploading…";
    const job = await fetch("https://api.cloudconvert.com/v2/jobs", {
      method : "POST",
      headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
      body   : JSON.stringify({
        tasks: {
          upload : { operation: "import/upload" },
          convert: { operation: "convert", input: "upload", input_format: "docx", output_format: "pdf" },
          export : { operation: "export/url", input: "convert" }
        }
      })
    }).then(r => r.json());

    const upTask = Object.values(job.data.tasks).find(t => t.operation === "import/upload");
    const fd = new FormData();
    Object.entries(upTask.result.form.parameters).forEach(([k, v]) => fd.append(k, v));
    fd.append("file", file);
    await fetch(upTask.result.form.url, { method: "POST", body: fd });

    status.textContent = "⏳ Converting…";
    let exportTask;
    while (!exportTask) {
      await new Promise(r => setTimeout(r, 2500));
      const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`,
        { headers: { Authorization: `Bearer ${apiKey}` } }).then(r => r.json());
      if (poll.data.status === "finished") {
        exportTask = poll.data.tasks.find(t => t.name === "export");
      }
    }
    const url = exportTask.result.files[0].url;
    status.innerHTML = `✅ Done! <a href="${url}" target="_blank">Download PDF</a>`;
  } catch (e) {
    console.error(e);
    status.textContent = `❌ ${e.message}`;
  }
}

/* handle Logout request */
async function requestLogout() {
  try {
    await ensureDefault();
    const email = getAuth().currentUser?.email ?? "Unknown";
    /* Open an Outlook draft – no external API needed */
    const mailto =
      `mailto:support@yourcompany.com` +
      `?subject=Logout%20Request&body=${encodeURIComponent(email)}%20wants%20to%20log%20out.`;
    window.location.href = mailto;

    /* local cleanup */
    await logoutRequestLocal();
    alert("Logout request created; please send the e‑mail.");
    window.location.reload();
  } catch (e) {
    console.error(e);
    alert("❌ Logout failed: " + e.message);
  }
}

/* wait until ui‑loader injects HTML, then wire buttons */
window.addEventListener("TaskpaneUILoaded", async () => {
  try {
    await ensureDefault();
    if (!(await isSessionValid())) {
      document.getElementById("app").innerHTML =
        "<div style='color:red'>Session expired – reload Add‑in.</div>";
      return;
    }
    /* show UI container (if css had display:none) */
    document.getElementById("main-ui")?.style.setProperty("display", "block");

    document.getElementById("convertBtn")?.addEventListener("click", convertToPDF);
    document.getElementById("requestLogoutBtn")?.addEventListener("click", requestLogout);
  } catch (err) {
    console.error(err);
    document.getElementById("app").innerHTML =
      `<div style="color:red;">${err.message}</div>`;
  }
});
