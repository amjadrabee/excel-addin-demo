// taskpane.js
import { loginUser, logoutUser, isSessionValid } from "../firebase-auth";

Office.onReady(async () => {
  const valid = await isSessionValid();
  if (!valid) {
    document.getElementById("login-container").style.display = "block";
    document.getElementById("main-ui").style.display = "none";
    return;
  }
  document.getElementById("login-container").style.display = "none";
  document.getElementById("main-ui").style.display = "block";

  const convertBtn = document.getElementById("convertBtn");
  if (convertBtn) convertBtn.onclick = convertToPDF;
});

window.loginUser = loginUser;
window.logoutUser = logoutUser;

async function convertToPDF() {
  const fileInput = document.getElementById("uploadDocx");
  const status = document.getElementById("status");
  const file = fileInput.files[0];

  if (!await isSessionValid()) {
    status.innerText = "❌ Your session is invalid. Please log in again.";
    document.getElementById("main-ui").style.display = "none";
    document.getElementById("login-container").style.display = "block";
    return;
  }

  if (!file) {
    status.innerText = "❌ Please select a .docx file first.";
    return;
  }

  try {
    status.innerText = "🔄 Creating conversion job...";

    const apiKey = "YOUR_CLOUDCONVERT_API_KEY";

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

    status.innerText = "📤 Uploading DOCX file...";
    const formData = new FormData();
    for (const key in uploadTask.result.form.parameters) {
      formData.append(key, uploadTask.result.form.parameters[key]);
    }
    formData.append("file", file);

    await fetch(uploadTask.result.form.url, {
      method: "POST",
      body: formData
    });

    status.innerText = "⏳ Waiting for conversion...";
    let exportTask, done = false;

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
