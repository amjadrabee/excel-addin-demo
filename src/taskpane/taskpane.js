import { loginUser, logoutUser, isSessionValid } from "../firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import {
  getAuth,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";

Office.onReady(async () => {
  const valid = await isSessionValid();
  if (!valid) {
    document.getElementById("login-container").style.display = "block";
    document.getElementById("main-ui").style.display = "none";
    return;
  }

  document.getElementById("login-container").style.display = "none";
  document.getElementById("main-ui").style.display = "block";

  document.getElementById("convertBtn").onclick = convertToPDF;
});

document.addEventListener("DOMContentLoaded", () => {
  document.querySelector("#login-container button").onclick = () => {
    const email = document.getElementById("emailInput").value;
    const password = document.getElementById("passwordInput").value;
    loginUser(email, password);
  };

  document.querySelector("#main-ui button").onclick = () => {
    logoutUser();
  };
});

// 🔐 Load CloudConvert API key from Firestore
async function getCloudConvertKey() {
  const db = getFirestore();
  const auth = getAuth();

  return new Promise((resolve, reject) => {
    onAuthStateChanged(auth, async user => {
      if (!user) return reject("❌ Not authenticated");

      try {
        const snap = await getDoc(doc(db, "config", "cloudconvert"));
        if (!snap.exists()) return reject("❌ CloudConvert key not found");
        resolve(snap.data().apiKey);
      } catch (e) {
        reject("❌ Error loading API key: " + e.message);
      }
    });
  });
}

async function convertToPDF() {
  const fileInput = document.getElementById("uploadDocx");
  const status = document.getElementById("status");
  const file = fileInput.files[0];

  if (!await isSessionValid()) {
    status.innerText = "❌ Session expired. Please log in again.";
    document.getElementById("main-ui").style.display = "none";
    document.getElementById("login-container").style.display = "block";
    return;
  }

  if (!file) {
    status.innerText = "❌ Select a .docx file.";
    return;
  }

  try {
    status.innerText = "🔄 Uploading...";

    // 🔐 Securely get API key
    const apiKey = await getCloudConvertKey();

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

    const job = await jobRes.json();
    const uploadTask = Object.values(job.data.tasks).find(t => t.name === "upload");

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
    let done = false, exportTask;
    while (!done) {
      const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
        headers: { Authorization: `Bearer ${apiKey}` }
      });
      const updatedJob = await poll.json();
      done = updatedJob.data.status === "finished";
      exportTask = updatedJob.data.tasks.find(t => t.name === "export");
      if (!done) await new Promise(r => setTimeout(r, 3000));
    }

    const fileUrl = exportTask.result.files[0].url;
    status.innerHTML = `✅ Done! <a href="${fileUrl}" target="_blank">Download PDF</a>`;
  } catch (err) {
    console.error(err);
    status.innerText = "❌ Conversion failed.";
  }
}
