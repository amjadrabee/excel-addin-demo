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

// Adjust this path if firebase-auth.js is in a different location relative to taskpane.js
import {
  logoutRequestLocal
} from "../firebase-auth.js";


let cloudConvertApiKey = null; // Stores the API key fetched from Firestore

async function ensureFirebase() {
  if (getApps().length) {
    // If the default Firebase app is already initialized (e.g., on page refresh),
    // try to re-fetch the CloudConvert key if it's not already set.
    const defaultApp = getApps().find(app => app.name === '[DEFAULT]');
    if (defaultApp && !cloudConvertApiKey) {
      try {
        const cfgDb = getFirestore(defaultApp);
        const cloudConvertSnap = await getDoc(doc(cfgDb, "config", "cloudconvert"));
        if (cloudConvertSnap.exists() && cloudConvertSnap.data().key) {
          cloudConvertApiKey = cloudConvertSnap.data().key;
        } else {
          console.warn("CloudConvert API key still missing after re-fetch attempt from existing Firebase app.");
        }
      } catch (error) {
        console.error("Error during re-fetching CloudConvert API key from existing app:", error);
      }
    }
    return;
  }

  // Initialize a temporary Firebase app to fetch configurations (before default app setup)
  const tmp = initializeApp({
    projectId: "excel-addin-auth"
  }, "tmp");
  const cfgDb = getFirestore(tmp);

  try {
    // Fetch Firebase config
    const firebaseSnap = await getDoc(doc(cfgDb, "config", "firebase"));
    if (!firebaseSnap.exists()) {
      throw new Error("❌ Firebase config missing in Firestore.");
    }
    initializeApp(firebaseSnap.data()); // Initialize the default Firebase app

    // Fetch CloudConvert API key
    const cloudConvertSnap = await getDoc(doc(cfgDb, "config", "cloudconvert"));
    if (!cloudConvertSnap.exists() || !cloudConvertSnap.data().key) {
      throw new Error("❌ CloudConvert API key missing in Firestore or empty.");
    }
    cloudConvertApiKey = cloudConvertSnap.data().key;

  } catch (error) {
    console.error("Error during add-in configuration (Firebase/CloudConvert key fetch):", error);
    document.getElementById("status").textContent = `❌ Add-in configuration failed. Check console.`;
    cloudConvertApiKey = null; // Ensure key is null on error
  } finally {
    await deleteApp(tmp); // Delete the temporary app
  }
}

/* ─── Office entry point ─── */
Office.onReady(async () => {
  await ensureFirebase();

  // Show the UI only after Firebase and config are processed
  document.getElementById("main-ui").style.display = "block";

  /* Hook buttons safely */
  const convertBtn = document.getElementById("convertBtn");
  const logoutBtn = document.getElementById("requestLogout");

  if (convertBtn) {
    convertBtn.addEventListener("click", convertToPDF);
  }
  if (logoutBtn) {
    logoutBtn.addEventListener("click", requestLogout);
  }
});

/* ─── Convert Word → PDF ─── */
async function convertToPDF() {
  const fileInput = document.getElementById("uploadDocx");
  const status = document.getElementById("status");
  const file = fileInput.files[0];

  if (!file) {
    status.textContent = "❌ Select a .docx file.";
    return;
  }

  // Ensure CloudConvert API key is available before proceeding
  if (!cloudConvertApiKey) {
    status.textContent = "❌ CloudConvert API key not loaded. Check console for configuration errors.";
    console.error("CloudConvert API key is missing. Conversion cannot proceed.");
    return;
  }

  const apiKey = cloudConvertApiKey; // Use the fetched key

  try {
    status.textContent = "🔄 Creating conversion job…";

    /* Step 1 — create job (upload → convert → export) */
    const jobRes = await fetch("https://api.cloudconvert.com/v2/jobs", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        tasks: {
          upload: {
            operation: "import/upload"
          },
          convert: {
            operation: "convert",
            input: "upload",
            input_format: "docx",
            output_format: "pdf"
          },
          export: {
            operation: "export/url",
            input: "convert"
          }
        }
      })
    });

    if (!jobRes.ok) {
      // Enhanced error handling for CloudConvert API responses
      let errorInfo = `CloudConvert API error: ${jobRes.status}`;
      try {
        const errorData = await jobRes.json();
        if (errorData && errorData.message) {
          errorInfo += ` - ${errorData.message}`;
        } else if (errorData && errorData.errors && errorData.errors.length > 0) {
          errorInfo += ` - ${errorData.errors[0].message || 'Unknown API error details'}`;
        } else {
          errorInfo += ` - (Raw response: ${JSON.stringify(errorData)})`;
        }
      } catch (e) {
        errorInfo += ` - (Failed to parse error response body: ${e.message})`;
      }
      throw new Error(errorInfo);
    }

    const job = await jobRes.json();
    const uploadTask = Object.values(job.data.tasks)
      .find(t => t.operation === "import/upload");

    /* Step 2 — upload DOCX */
    status.textContent = "🔄 Uploading file…";
    const fd = new FormData();
    for (const k in uploadTask.result.form.parameters) {
      fd.append(k, uploadTask.result.form.parameters[k]);
    }
    fd.append("file", file);

    const uploadRes = await fetch(uploadTask.result.form.url, {
      method: "POST",
      body: fd
    });
    if (!uploadRes.ok) {
      throw new Error(`CloudConvert upload failed with status: ${uploadRes.status}`);
    }


    /* Step 3 — poll until finished */
    status.textContent = "⏳ Converting…";
    let exportTask;
    while (true) {
      const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
        headers: {
          "Authorization": `Bearer ${apiKey}`
        }
      });
      const info = await poll.json();
      if (info.data.status === "finished") {
        exportTask = info.data.tasks.find(t => t.name === "export");
        break;
      } else if (info.data.status === "error") {
        throw new Error(`CloudConvert job failed: ${info.data.message || 'Unknown error during conversion.'}`);
      }
      await new Promise(r => setTimeout(r, 3000)); // Wait 3 seconds before next poll
    }

    /* Step 4 — show link */
    if (!exportTask || !exportTask.result || !exportTask.result.files || exportTask.result.files.length === 0) {
        throw new Error("CloudConvert did not return an exported file URL.");
    }
    const url = exportTask.result.files[0].url;
    status.innerHTML = `✅ Done! <a href="${url}" target="_blank">Download PDF</a>`;

  } catch (err) {
    console.error("Conversion error:", err);
    status.textContent = `❌ Conversion failed: ${err.message || "See console for details."}`;
  }
}

/* ─── Request Logout (opens mail client) ─── */
async function requestLogout() {
  const email = localStorage.getItem("email") || "Unknown User";
  const subject = encodeURIComponent("Logout Request");
  const body = encodeURIComponent(`${email} requests logout from Excel Add‑in.`);
  window.location.href = `mailto:support@yourcompany.com?subject=${subject}&body=${body}`;

  /* Local clean‑up */
  await logoutRequestLocal();
  // Optional: window.location.reload(); // Uncomment if you want to force a page refresh after mailto link
}
