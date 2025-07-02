/* ─────────────────────────────────────────────────────────────
   taskpane.js  –  FINAL WORKING VERSION
───────────────────────────────────────────────────────────── */

import { logoutRequestLocal } from "../firebase-auth.js";

import {
  initializeApp,
  getApps,
  deleteApp                     // ✅ correct import here
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";

import { getAuth } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

/* ─── Ensure Firebase default app is ready ───────────────── */
async function ensureFirebase() {
  if (getApps().length) return;                    // already initialised

  // 1) temp app (projectId‑only) just to fetch full config
  const tmp = initializeApp({ projectId: "excel-addin-auth" }, "tmp-taskpane");
  const tmpDb = getFirestore(tmp);

  const cfgSnap = await getDoc(doc(tmpDb, "config", "firebase"));
  if (!cfgSnap.exists()) throw new Error("❌ Firebase config missing in Firestore.");
  const fullCfg = cfgSnap.data();

  await deleteApp(tmp);                            // dispose temp app

  // 2) real default app
  initializeApp(fullCfg);
}
/* ─────────────────────────────────────────────────────────── */

Office.onReady(async () => {
  await ensureFirebase();                          // 🟢 Firebase ready

  /* show main UI */
  document.getElementById("main-ui").style.display = "block";

  /* wire buttons */
  document.getElementById("convertBtn").addEventListener("click", convertToPDF);
  document.getElementById("requestLogout").addEventListener("click", requestLogout);
});

/* ─── Convert DOCX → PDF via CloudConvert ─────────────────── */
async function convertToPDF() {
  const fileInput = document.getElementById("uploadDocx");
  const status    = document.getElementById("status");
  const file      = fileInput.files[0];

  if (!file) {
    status.textContent = "❌ Select a .docx file.";
    return;
  }

  try {
    status.textContent = "🔄 Fetching API key…";

    const auth = getAuth();
    const db   = getFirestore();

    if (!auth.currentUser) throw new Error("❌ Not logged in.");

    const keySnap = await getDoc(doc(db, "config", "cloudconvert"));
    if (!keySnap.exists()) throw new Error("❌ API key not found.");
    const apiKey = keySnap.data().key;

    /* create CloudConvert job */
    status.textContent = "🔄 Uploading…";
    const jobRes = await fetch("https://api.cloudconvert.com/v2/jobs", {
      method : "POST",
      headers: {
        Authorization : `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        tasks: {
          upload : { operation: "import/upload" },
          convert: { operation: "convert", input: "upload",
                     input_format: "docx", output_format: "pdf" },
          export : { operation: "export/url", input: "convert" }
        }
      })
    });

    const job = await jobRes.json();
    const uploadTask = Object.values(job.data.tasks)
                              .find(t => t.operation === "import/upload");

    /* upload file */
    const formData = new FormData();
    for (const k in uploadTask.result.form.parameters) {
      formData.append(k, uploadTask.result.form.parameters[k]);
    }
    formData.append("file", file);

    await fetch(uploadTask.result.form.url, { method: "POST", body: formData });

    /* poll until finished */
    status.textContent = "⏳ Converting…";
    let exportTask;
    while (true) {
      const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
        headers: { Authorization: `Bearer ${apiKey}` }
      }).then(r => r.json());

      if (poll.data.status === "finished") {
        exportTask = poll.data.tasks.find(t => t.name === "export");
        break;
      }
      await new Promise(r => setTimeout(r, 3000));
    }

    const fileUrl = exportTask.result.files[0].url;
    status.innerHTML = `✅ Done! <a href="${fileUrl}" target="_blank">Download PDF</a>`;
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Conversion failed. Check console.";
  }
}

/* ─── Request Logout (opens default mail client) ──────────── */
async function requestLogout() {
  const email   = localStorage.getItem("email") || "Unknown user";
  const subject = encodeURIComponent("Logout Request");
  const body    = encodeURIComponent(`${email} requests logout from Excel Add‑in.`);
  window.location.href =
    `mailto:support@yourcompany.com?subject=${subject}&body=${body}`;

  /* clear local session */
  await logoutRequestLocal();
  window.location.reload();
}
/* ─────────────────────────────────────────────────────────── */


//////////////////////////////////////////////////////////

// import { logoutRequestLocal } from "../firebase-auth.js";
// import { initializeApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
// import { getAuth } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
// import { getFirestore, doc, getDoc, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";


// async function ensureFirebase() {
//   if (getApps().length) return;

//   const tmpApp = initializeApp({ projectId: "excel-addin-auth" }, "tmp-taskpane");
//   const tmpDb = getFirestore(tmpApp);
//   const cfgSnap = await getDoc(doc(tmpDb, "config", "firebase"));
//   if (!cfgSnap.exists()) throw new Error("❌ Firebase config missing in Firestore.");
//   const fullCfg = cfgSnap.data();
//   await deleteApp(tmpApp);

//   initializeApp(fullCfg);
// }

// Office.onReady(async () => {
//   await ensureFirebase();
//   document.getElementById("main-ui").style.display = "block";
//   document.getElementById("convertBtn").addEventListener("click", convertToPDF);
//   document.getElementById("requestLogout").addEventListener("click", requestLogout);
//   document.getElementById("convertBtn").onclick = convertToPDF;
  
// });

// async function convertToPDF() {
//   const fileInput = document.getElementById("uploadDocx");
//   const status = document.getElementById("status");
//   const file = fileInput.files[0];

//   if (!file) {
//     status.innerText = "❌ Select a .docx file.";
//     return;
//   }

//   try {
//     status.innerText = "🔄 Fetching API key...";

//     const auth = getAuth();
//     const db = getFirestore();
//     const user = auth.currentUser;
//     if (!user) throw new Error("❌ Not logged in.");

//     const keySnap = await getDoc(doc(db, "config", "cloudconvert"));
//     if (!keySnap.exists()) throw new Error("❌ API key not found.");
//     const apiKey = keySnap.data().key;

//     status.innerText = "🔄 Uploading...";

//     const jobRes = await fetch("https://api.cloudconvert.com/v2/jobs", {
//       method: "POST",
//       headers: {
//         Authorization: `Bearer ${apiKey}`,
//         "Content-Type": "application/json"
//       },
//       body: JSON.stringify({
//         tasks: {
//           upload: { operation: "import/upload" },
//           convert: {
//             operation: "convert",
//             input: "upload",
//             input_format: "docx",
//             output_format: "pdf"
//           },
//           export: { operation: "export/url", input: "convert" }
//         }
//       })
//     });

//     const job = await jobRes.json();
//     const uploadTask = Object.values(job.data.tasks).find(t => t.operation === "import/upload");

//     const formData = new FormData();
//     for (const key in uploadTask.result.form.parameters) {
//       formData.append(key, uploadTask.result.form.parameters[key]);
//     }
//     formData.append("file", file);

//     await fetch(uploadTask.result.form.url, {
//       method: "POST",
//       body: formData
//     });

//     status.innerText = "⏳ Converting...";

//     let done = false;
//     let exportTask;
//     while (!done) {
//       const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
//         headers: { Authorization: `Bearer ${apiKey}` }
//       });
//       const updatedJob = await poll.json();
//       done = updatedJob.data.status === "finished";
//       exportTask = updatedJob.data.tasks.find(t => t.name === "export");
//       if (!done) await new Promise(r => setTimeout(r, 3000));
//     }

//     const fileUrl = exportTask.result.files[0].url;
//     status.innerHTML = `✅ Done! <a href="${fileUrl}" target="_blank">Download PDF</a>`;
//   } catch (err) {
//     console.error(err);
//     status.innerText = "❌ Conversion failed. Check the console for errors.";
//   }
// }

// async function requestLogout() {
//   const userEmail = localStorage.getItem("uid") || "Unknown User";
//   const subject = encodeURIComponent("Logout Request");
//   const body = encodeURIComponent(`${userEmail} has requested to log out from the Excel Add-in.`);
//   const mailtoLink = `mailto:support@yourcompany.com?subject=${subject}&body=${body}`;
//   window.location.href = mailtoLink;
//   await logoutRequestLocal();
//   window.location.reload();
// }
