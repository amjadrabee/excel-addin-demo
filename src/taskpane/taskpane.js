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
// src/taskpane/taskpane.js

import {
  initializeApp,
  getApps,
  deleteApp
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";

import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

import { logoutRequestLocal } from "../firebase-auth.js";

let cloudConvertApiKey = null;

/* ───── Ensure Firebase Initialized ───── */
async function ensureFirebase() {
  if (getApps().length > 0) {
    console.log("✅ Firebase already initialized");

    const db = getFirestore();
    try {
      const cloudConvertSnap = await getDoc(doc(db, "config", "cloudconvert"));
      if (cloudConvertSnap.exists() && cloudConvertSnap.data().key) {
        cloudConvertApiKey = cloudConvertSnap.data().key;
        console.log("✅ CloudConvert key loaded from existing app");
      } else {
        console.warn("⚠️ CloudConvert key missing or empty in Firestore.");
      }
    } catch (err) {
      console.error("❌ Error fetching CloudConvert key from existing app:", err);
    }
    return;
  }

  // Temp app for config loading
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "temp-taskpane");
  const tempDb = getFirestore(tempApp);

  try {
    const firebaseSnap = await getDoc(doc(tempDb, "config", "firebase"));
    if (!firebaseSnap.exists()) throw new Error("❌ Firebase config missing.");

    const firebaseConfig = firebaseSnap.data();
    initializeApp(firebaseConfig); // Init default app

    const cloudConvertSnap = await getDoc(doc(tempDb, "config", "cloudconvert"));
    if (!cloudConvertSnap.exists() || !cloudConvertSnap.data().key) {
      throw new Error("❌ CloudConvert API key missing in Firestore.");
    }

    cloudConvertApiKey = cloudConvertSnap.data().key;
    console.log("✅ CloudConvert key loaded via temp app");

  } catch (error) {
    console.error("❌ Config loading error:", error);
    document.getElementById("status").textContent = "❌ Config error. See console.";
  } finally {
    await deleteApp(tempApp);
  }
}

/* ───── Office Add-in Ready ───── */
Office.onReady(async () => {
  await ensureFirebase();

  document.getElementById("main-ui").style.display = "block";

  const convertBtn = document.getElementById("convertBtn");
  const logoutBtn = document.getElementById("requestLogout");

  if (convertBtn) convertBtn.onclick = convertToPDF;
  if (logoutBtn) logoutBtn.onclick = requestLogout;
});

/* ───── Convert DOCX → PDF ───── */
async function convertToPDF() {
  const fileInput = document.getElementById("uploadDocx");
  const status = document.getElementById("status");
  const file = fileInput.files[0];

  if (!file) {
    status.textContent = "❌ Select a .docx file.";
    return;
  }

  if (!cloudConvertApiKey) {
    status.textContent = "❌ CloudConvert API key not loaded. Check console.";
    console.error("❌ API key missing.");
    return;
  }

  try {
    status.textContent = "🔄 Creating job…";

    const jobRes = await fetch("https://api.cloudconvert.com/v2/jobs", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${cloudConvertApiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        tasks: {
          upload: { operation: "import/upload" },
          convert: {
            operation: "convert",
            input: "upload",
            input_format: "docx",
            output_format: "pdf"
          },
          export: { operation: "export/url", input: "convert" }
        }
      })
    });

    if (!jobRes.ok) {
      const errJson = await jobRes.json().catch(() => ({}));
      throw new Error(`CloudConvert error ${jobRes.status}: ${errJson.message || "Unknown error"}`);
    }

    const job = await jobRes.json();
    const uploadTask = Object.values(job.data.tasks).find(t => t.operation === "import/upload");

    /* Upload File */
    const formData = new FormData();
    for (const key in uploadTask.result.form.parameters) {
      formData.append(key, uploadTask.result.form.parameters[key]);
    }
    formData.append("file", file);

    status.textContent = "🔄 Uploading file...";
    const uploadRes = await fetch(uploadTask.result.form.url, {
      method: "POST",
      body: formData
    });

    if (!uploadRes.ok) {
      throw new Error(`Upload failed: ${uploadRes.status}`);
    }

    /* Poll Until Done */
    status.textContent = "⏳ Converting...";
    let exportTask;
    while (true) {
      const pollRes = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
        headers: { Authorization: `Bearer ${cloudConvertApiKey}` }
      });
      const pollData = await pollRes.json();

      if (pollData.data.status === "finished") {
        exportTask = pollData.data.tasks.find(t => t.name === "export");
        break;
      } else if (pollData.data.status === "error") {
        throw new Error("❌ CloudConvert job failed.");
      }

      await new Promise(r => setTimeout(r, 3000)); // Wait 3s
    }

    /* Show Download Link */
    const downloadUrl = exportTask?.result?.files?.[0]?.url;
    if (!downloadUrl) throw new Error("No download URL returned.");

    status.innerHTML = `✅ Done! <a href="${downloadUrl}" target="_blank">Download PDF</a>`;

  } catch (err) {
    console.error("❌ Conversion error:", err);
    status.textContent = `❌ Failed: ${err.message || "Check console."}`;
  }
}

/* ───── Request Logout ───── */
async function requestLogout() {
  try {
    const userEmail = localStorage.getItem("email") || localStorage.getItem("uid") || "Unknown User";

    await fetch("https://your-api.example.com/send-logout-request", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        to: "support@yourcompany.com",
        subject: "Logout Request",
        message: `${userEmail} has requested to log out from the Excel Add-in.`
      })
    });

    alert("📩 Logout request sent.");
    await logoutRequestLocal();
    localStorage.clear();
    window.location.href = "login.html";

  } catch (err) {
    console.error("❌ Logout error:", err);
    alert("❌ Logout failed.");
  }
}
