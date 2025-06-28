import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

let apiKey;

async function getApiKey() {
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "taskpane-app");
  const db = getFirestore(tempApp);
  const snap = await getDoc(doc(db, "config", "cloudconvert"));
  if (!snap.exists()) throw new Error("No API key in Firestore");
  apiKey = snap.data().apiKey;
}

Office.onReady(async () => {
  await getApiKey();

  document.getElementById("convertBtn").onclick = async () => {
    const fileInput = document.getElementById("uploadDocx");
    const status = document.getElementById("status");
    const file = fileInput.files[0];

    if (!file) {
      status.textContent = "❌ Please select a .docx file.";
      return;
    }

    try {
      status.textContent = "🔄 Uploading...";
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
      Object.entries(uploadTask.result.form.parameters).forEach(([k, v]) => formData.append(k, v));
      formData.append("file", file);

      await fetch(uploadTask.result.form.url, {
        method: "POST",
        body: formData
      });

      let done = false, exportTask;
      status.textContent = "⏳ Converting...";
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
      status.textContent = "❌ Conversion failed.";
    }
  };
});
