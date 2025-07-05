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
// src/taskpane/taskpane.js

import {
  initializeApp,
  getApps,
  deleteApp
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc,
  collection, // <--- ADD THIS
  addDoc      // <--- ADD THIS
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

import {
  logoutRequestLocal,
  getAuthInstance, // <--- ADD THIS
  getDbInstance    // <--- ADD THIS
} from "../firebase-auth.js"; // Adjust path if needed


let cloudConvertApiKey = null;
// Declare auth and db variables at a scope accessible by conversion functions
let authInstance = null; // To hold the Firebase Auth instance
let dbInstance = null;   // To hold the Firestore DB instance

async function ensureFirebaseAndConfig() {
  const statusElement = document.getElementById("status");
  if (statusElement) statusElement.textContent = "🔄 Initializing add-in...";

  console.log("ensureFirebaseAndConfig started.");

  let firebaseAppInstance;
  const existingApps = getApps();

  if (existingApps.length > 0) {
    firebaseAppInstance = existingApps.find(app => app.name === '[DEFAULT]');
    if (firebaseAppInstance) {
      console.log("Default Firebase app already initialized.");
      authInstance = getAuthInstance(firebaseAppInstance); // Get auth instance from existing app
      dbInstance = getDbInstance(firebaseAppInstance);     // Get db instance from existing app
    }
  }

  if (!firebaseAppInstance) {
    console.log("No default Firebase app found. Initializing temporary Firebase app to fetch configs.");
    const tmp = initializeApp({
      projectId: "excel-addin-auth"
    }, "tmp");
    const cfgDb = getFirestore(tmp);

    try {
      console.log("Fetching Firebase config from Firestore...");
      const firebaseSnap = await getDoc(doc(cfgDb, "config", "firebase"));
      if (!firebaseSnap.exists()) {
        throw new Error("❌ Firebase config missing in Firestore.");
      }
      firebaseAppInstance = initializeApp(firebaseSnap.data()); // Initialize the DEFAULT Firebase app
      authInstance = getAuthInstance(firebaseAppInstance); // Initialize auth instance for new default app
      dbInstance = getDbInstance(firebaseAppInstance);     // Initialize db instance for new default app
      console.log("Default Firebase app and instances initialized with fetched config.");

      console.log("Fetching CloudConvert API key from Firestore...");
      const cloudConvertSnap = await getDoc(doc(cfgDb, "config", "cloudconvert"));
      if (!cloudConvertSnap.exists() || !cloudConvertSnap.data().key) {
        throw new Error("❌ CloudConvert API key missing or empty in Firestore.");
      }
      cloudConvertApiKey = cloudConvertSnap.data().key;
      console.log("CloudConvert API key successfully fetched and assigned.");

    } catch (error) {
      console.error("Error during initial Firebase/config fetch:", error);
      if (statusElement) statusElement.textContent = `❌ Add-in failed to load config: ${error.message}`;
      cloudConvertApiKey = null;
      throw error; // Re-throw to prevent UI from showing if config failed
    } finally {
      await deleteApp(tmp);
      console.log("Temporary Firebase app deleted.");
    }
  } else {
    // If default Firebase app was already initialized, ensure cloudConvertApiKey is fetched
    if (!cloudConvertApiKey) {
      console.log("Default Firebase app already existed, but CloudConvert API key not set. Attempting to fetch key.");
      try {
        const cfgDb = getFirestore(firebaseAppInstance); // Get Firestore from the existing default app
        const cloudConvertSnap = await getDoc(doc(cfgDb, "config", "cloudconvert"));
        if (cloudConvertSnap.exists() && cloudConvertSnap.data().key) {
          cloudConvertApiKey = cloudConvertSnap.data().key;
          console.log("CloudConvert API key fetched from existing Firebase app.");
        } else {
          console.warn("CloudConvert API key still missing/empty after re-fetch attempt from existing Firebase app.");
          throw new Error("CloudConvert API key not found in Firestore after re-fetch.");
        }
      } catch (error) {
        console.error("Error fetching CloudConvert key from existing Firebase app:", error);
        if (statusElement) statusElement.textContent = `❌ CloudConvert key load failed: ${error.message}`;
        cloudConvertApiKey = null;
        throw error;
      }
    } else {
      console.log("CloudConvert API key already present.");
    }
  }

  if (!cloudConvertApiKey) {
    const errMsg = "CloudConvert API key is critical for conversion and could not be loaded.";
    console.error(errMsg);
    if (statusElement) statusElement.textContent = `❌ ${errMsg}`;
    throw new Error(errMsg);
  } else {
    if (statusElement) statusElement.textContent = "✅ Add-in ready.";
  }
}


/* ─── Office entry point ─── */
Office.onReady(async () => {
  console.log("Office.onReady has fired. Starting config ensureance.");
  try {
    await ensureFirebaseAndConfig();
    document.getElementById("main-ui").style.display = "block"; // Show UI only after successful config load

    const convertBtn = document.getElementById("convertBtn");
    const logoutBtn = document.getElementById("requestLogout");

    if (convertBtn) {
      convertBtn.addEventListener("click", convertToPDF);
      console.log("Convert button event listener attached.");
    }
    if (logoutBtn) {
      logoutBtn.addEventListener("click", requestLogout);
      console.log("Logout button event listener attached.");
    }

  } catch (error) {
    console.error("Add-in failed to initialize:", error);
    const statusElement = document.getElementById("status");
    if (statusElement) statusElement.textContent = `❌ Initialization failed: ${error.message || 'Unknown error'}. Check console.`;
    document.getElementById("main-ui").style.display = "block"; // Show UI to display error
  }
});

/* ─── Convert Word → PDF ─── */
async function convertToPDF() {
  console.log("convertToPDF function called.");
  const fileInput = document.getElementById("uploadDocx");
  const status = document.getElementById("status");
  const file = fileInput.files[0];

  if (!file) {
    status.textContent = "❌ Select a .docx file.";
    console.warn("No file selected for conversion.");
    return;
  }

  if (!cloudConvertApiKey) {
    status.textContent = "❌ CloudConvert API key not loaded. Cannot proceed with conversion.";
    console.error("CloudConvert API key is null or not fetched successfully. Check console for config errors.");
    return;
  }

  const apiKey = cloudConvertApiKey; // Use the fetched key

  try {
    status.textContent = "🔄 Creating conversion job…";
    console.log("Attempting to create CloudConvert job...");

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
    console.log("CloudConvert job created:", job);

    const uploadTask = Object.values(job.data.tasks)
      .find(t => t.operation === "import/upload");

    if (!uploadTask || !uploadTask.result || !uploadTask.result.form || !uploadTask.result.form.url) {
      throw new Error("CloudConvert did not return expected upload task details.");
    }

    /* Step 2 — upload DOCX */
    status.textContent = "🔄 Uploading file…";
    console.log("Attempting to upload file to CloudConvert...");
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
    console.log("File uploaded successfully.");


    /* Step 3 — poll until finished */
    status.textContent = "⏳ Converting…";
    console.log("Polling CloudConvert job status...");
    let exportTask;
    while (true) {
      const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
        headers: {
          "Authorization": `Bearer ${apiKey}`
        }
      });
      const info = await poll.json();
      console.log("CloudConvert poll status:", info.data.status);
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
    console.log("Conversion successful, PDF URL:", url);

    // --- NEW: Log conversion to Firestore ---
    await logConversion(file.name, url, "success");
    // --- END NEW ---

  } catch (err) {
    console.error("Conversion error:", err);
    status.textContent = `❌ Conversion failed: ${err.message || "See console for details."}`;
    // --- NEW: Log conversion failure to Firestore ---
    await logConversion(file ? file.name : "N/A", "N/A", "failed", err.message || "Unknown error");
    // --- END NEW ---
  }
}

/* ─── NEW FUNCTION: Log Conversion to Firestore ─── */
async function logConversion(fileName, downloadUrl, status, errorMessage = null) {
  if (!dbInstance || !authInstance || !authInstance.currentUser) {
    console.warn("Firestore or Auth not initialized, or no user logged in. Cannot log conversion.");
    return;
  }

  const userId = authInstance.currentUser.uid;
  const conversionData = {
    userId: userId,
    fileName: fileName,
    downloadUrl: downloadUrl,
    status: status,
    timestamp: new Date(),
    errorMessage: errorMessage
  };

  try {
    // Add a new document to a 'conversion_history' sub-collection under the user's session document
    // Alternatively, you could have a top-level 'conversions' collection with a userId field
    await addDoc(collection(dbInstance, "sessions", userId, "conversion_history"), conversionData);
    console.log("Conversion logged to Firestore:", conversionData);
  } catch (error) {
    console.error("Error logging conversion to Firestore:", error);
  }
}
/* ─── END NEW FUNCTION ─── */


/* ─── Request Logout (opens mail client) ─── */
async function requestLogout() {
  console.log("requestLogout function called.");
  const email = localStorage.getItem("email") || "Unknown User";
  const subject = encodeURIComponent("Logout Request");
  const body = encodeURIComponent(`${email} requests logout from Excel Add‑in.`);
  window.location.href = `mailto:support@yourcompany.com?subject=${subject}&body=${body}`;

  /* local clean‑up */
  await logoutRequestLocal();
  console.log("logoutRequestLocal completed.");
}
