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

import { logoutRequestLocal } from "../firebase-auth.js";
import { initializeApp, getApps } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import { getFirestore, doc, getDoc, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";


async function ensureFirebase() {
    if (getApps().length) return;

    const tmpApp = initializeApp({ projectId: "excel-addin-auth" }, "tmp-taskpane");
    const tmpDb = getFirestore(tmpApp);
    const cfgSnap = await getDoc(doc(tmpDb, "config", "firebase"));
    if (!cfgSnap.exists()) throw new Error("❌ Firebase config missing in Firestore.");
    const fullCfg = cfgSnap.data();
    await deleteApp(tmpApp);

    initializeApp(fullCfg);
}

Office.onReady(async () => {
    await ensureFirebase();
    document.getElementById("main-ui").style.display = "block";
    document.getElementById("convertBtn").addEventListener("click", convertToPDF);
    document.getElementById("requestLogout").addEventListener("click", requestLogout);
    document.getElementById("convertBtn").onclick = convertToPDF;

});

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
async function convertToPDF() {
    const fileInput = document.getElementById("uploadDocx");
    const status = document.getElementById("status");
    const file = fileInput.files[0];

    if (!file) {
        status.innerText = "❌ Select a .docx file.";
        return;
    }

    try {
        status.innerText = "🔄 Using hardcoded API key...";

        // ⛔️ Replace with your real API key
        const apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiMmY3OWNjYWIyMzBiZDZmYzg2NGRmZDFiZTJlY2FiNWRhYTJjZDhlNmM5NGQ0ZTkxODAxYzIyN2M4ZjRjM2IyNzRlZmE2YzJlODA0ZGJlYzAiLCJpYXQiOjE3NTEyMDUyMzkuMDY0NDQ0LCJuYmYiOjE3NTEyMDUyMzkuMDY0NDQ1LCJleHAiOjQ5MDY4Nzg4MzkuMDYxNTMyLCJzdWIiOiI3MjIzNTM3NCIsInNjb3BlcyI6WyJ0YXNrLnJlYWQiLCJ0YXNrLndyaXRlIl19.Qw3ocALIUnOt1vwliMVam0-IK1lGOwLuGsLiHKuCXi0QguE5SioeTjlg00RpTuzDl-YZrMvNFVSAKRa4rylttOHaLRA5E61qhc04qpfg-ryi5x_Cmo5dCUDoafD-kS1rEo22MHNiI9zaTztuJ5viVqlPIuObKc29pTTDujdYk6W8UxExukKkRLbA8hf56PP0khIQSEXy06-pE6oBNzdkJd7B9LzU-FB_tUXmkPVOnUHR5dEtdwrmZmBMhbQcZVdD18qjtX1w3JCp2vVA0xgzWiTasTF04jGc_bmc2u89yyslEHihA63hiuEWePYPcz4n8s-UGs13wHk8O0i8fCXCL7_xFMRWObElnOmxcqLYCoeJamNQyRPY7ad9c1H1OR1zGTmncPdXvupFiAcjzt9hsG7S7NeRb_5luKhxes3_utrSr2zcAZyQLmgYjVWICmGR2HQgrVCIiJ3IcYMOk_EcvVjlenx-w2vo2BCr9a4sw7SQG1RPCxPsWIwVP7f9AxkDb7fReN0rNcUvgmS-BvsvfxAAMc6npfcDaWSZl1S1JO1vaRFHtzXlKjvYmDNOHat2ERDQFBGO70oluqWrqmFVSPKrK1dNXrEsfXTFNTpuWKn1WSWff4sAYO70DKQwd6u6YsaWH9Lu_qGVLLBkMbKnAfGYHG-tndE6UJRbEPz0v-I";

        status.innerText = "🔄 Uploading...";

        const jobRes = await fetch("https://api.cloudconvert.com/v2/jobs", {
            method: "POST",
            headers: {
                Authorization: `Bearer ${apiKey}`,
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

        if (!jobRes.ok) throw new Error("❌ Failed to create conversion job.");
        const job = await jobRes.json();
        const uploadTask = Object.values(job.data.tasks).find(t => t.operation === "import/upload");

        const formData = new FormData();
        for (const key in uploadTask.result.form.parameters) {
            formData.append(key, uploadTask.result.form.parameters[key]);
        }
        formData.append("file", file);

        await fetch(uploadTask.result.form.url, {
            method: "POST",
            body: formData
        });

        status.innerText = "⏳ Converting...";

        let done = false;
        let exportTask;
        while (!done) {
            const pollRes = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
                headers: { Authorization: `Bearer ${apiKey}` }
            });
            const pollData = await pollRes.json();
            done = pollData.data.status === "finished";
            exportTask = pollData.data.tasks.find(t => t.name === "export");
            if (!done) await new Promise(r => setTimeout(r, 3000));
        }

        const fileUrl = exportTask.result.files[0].url;
        status.innerHTML = `✅ Done! <a href="${fileUrl}" target="_blank">Download PDF</a>`;

    } catch (err) {
        console.error("❌ Convert Error:", err);
        status.innerText = "❌ Conversion failed. Check the console for errors.";
    }
}

async function requestLogout() {
    const userEmail = localStorage.getItem("uid") || "Unknown User";
    const subject = encodeURIComponent("Logout Request");
    const body = encodeURIComponent(`${userEmail} has requested to log out from the Excel Add-in.`);
    const mailtoLink = `mailto:support@yourcompany.com?subject=${subject}&body=${body}`;
    window.location.href = mailtoLink;
    await logoutRequestLocal();
    window.location.reload();
}
