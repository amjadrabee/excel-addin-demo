// import { logoutRequestLocal } from "../firebase-auth.js";
// import { initializeApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
// import { getAuth } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
// import {
//   getFirestore,
//   doc,
//   getDoc
// } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// // Load Firebase config from Firestore and initialize
// async function ensureFirebase() {
//   if (getApps().length) return; // already initialized

//   const tmpApp = initializeApp({ projectId: "excel-addin-auth" }, "tmp-taskpane");
//   const tmpDb = getFirestore(tmpApp);
//   const cfgSnap = await getDoc(doc(tmpDb, "config", "firebase"));
//   if (!cfgSnap.exists()) throw new Error("❌ Firebase config missing in Firestore.");
//   const fullCfg = cfgSnap.data();

//   // Initialize main app
//   initializeApp(fullCfg);
// }

// Office.onReady(async () => {
//   await ensureFirebase();

//   document.getElementById("main-ui").style.display = "block";
//   document.getElementById("convertBtn").addEventListener("click", convertToPDF);
//   document.getElementById("requestLogout").addEventListener("click", requestLogout);
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

//     if (!jobRes.ok) throw new Error("Failed to create CloudConvert job.");
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
//     console.error("❌ Convert Error:", err);
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

//////////////////////////////////////////////////////////

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


/*  src/taskpane/taskpane.js  */

/* ─── Firebase bootstrap (projectId‑only) ─── */
import {
  initializeApp,
  getApps,
  deleteApp         // <- correct import location
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";

async function ensureFirebase () {
  if (getApps().length) return;                   // already initialised
  const tmp   = initializeApp({ projectId: "excel-addin-auth" }, "tmp");
  const cfg   = (await import(
    "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js"
  )).getFirestore(tmp);
  const snap  = await (await import(
    "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js"
  )).getDoc((await import(
    "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js"
  )).doc(cfg, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config missing in Firestore");
  await deleteApp(tmp);
  initializeApp(snap.data());                      // default app
}

/* ─── Office entry point ─── */
Office.onReady(async () => {
  await ensureFirebase();

  /* show UI once loaded */
  document.getElementById("main-ui").style.display = "block";

  /* hook buttons safely */
  const convertBtn = document.getElementById("convertBtn");
  const logoutBtn  = document.getElementById("requestLogout");

  if (convertBtn) convertBtn.addEventListener("click", convertToPDF);
  if (logoutBtn)  logoutBtn .addEventListener("click", requestLogout);
});

/* ─── Convert Word → PDF using hard‑coded key ─── */
async function convertToPDF () {
  const fileInput = document.getElementById("uploadDocx");
  const status    = document.getElementById("status");
  const file      = fileInput.files[0];

  if (!file) { status.textContent = "❌ Select a .docx file."; return; }

  /* ⚠️ HARD‑CODED CloudConvert API key */
  const apiKey = "YOUR‑CLOUDCONVERT‑API‑KEY‑HERE";

  try {
    status.textContent = "🔄 Creating conversion job…";

    /* Step 1 — create job (upload → convert → export) */
    const jobRes = await fetch("https://api.cloudconvert.com/v2/jobs", {
      method : "POST",
      headers: {
        "Authorization" : `Bearer ${apiKey}`,
        "Content-Type"  : "application/json"
      },
      body: JSON.stringify({
        tasks: {
          upload : { operation: "import/upload" },
          convert: {
            operation    : "convert",
            input        : "upload",
            input_format : "docx",
            output_format: "pdf"
          },
          export : { operation: "export/url", input: "convert" }
        }
      })
    });

    if (!jobRes.ok) throw new Error(`CloudConvert error ${jobRes.status}`);

    const job       = await jobRes.json();
    const uploadTask= Object.values(job.data.tasks)
                             .find(t => t.operation === "import/upload");

    /* Step 2 — upload DOCX */
    status.textContent = "🔄 Uploading file…";
    const fd = new FormData();
    for (const k in uploadTask.result.form.parameters) {
      fd.append(k, uploadTask.result.form.parameters[k]);
    }
    fd.append("file", file);

    await fetch(uploadTask.result.form.url, { method: "POST", body: fd });

    /* Step 3 — poll until finished */
    status.textContent = "⏳ Converting…";
    let exportTask;
    while (true) {
      const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
        headers: { "Authorization": `Bearer ${apiKey}` }
      });
      const info = await poll.json();
      if (info.data.status === "finished") {
        exportTask = info.data.tasks.find(t => t.name === "export");
        break;
      }
      await new Promise(r => setTimeout(r, 3000));
    }

    /* Step 4 — show link */
    const url = exportTask.result.files[0].url;
    status.innerHTML = `✅ Done! <a href="${url}" target="_blank">Download PDF</a>`;
  }
  catch (err) {
    console.error(err);
    status.textContent = "❌ Conversion failed – see console.";
  }
}

/* ─── Request Logout (opens mail client) ─── */
import { logoutRequestLocal } from "../firebase-auth.js";
async function requestLogout () {
  const email   = localStorage.getItem("email") || "Unknown User";
  const subject = encodeURIComponent("Logout Request");
  const body    = encodeURIComponent(`${email} requests logout from Excel Add‑in.`);
  window.location.href = `mailto:support@yourcompany.com?subject=${subject}&body=${body}`;

  /* local clean‑up */
  await logoutRequestLocal();
  localStorage.clear();
  // optional: window.location.reload();
}

