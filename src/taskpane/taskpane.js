Office.onReady(async () => {
  const valid = await window.isSessionValid();
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
  const loginBtn = document.querySelector("#login-container button");
  if (loginBtn) {
    loginBtn.onclick = () => {
      const email = document.getElementById("emailInput").value;
      const password = document.getElementById("passwordInput").value;
      window.loginUser(email, password);
    };
  }

  const logoutBtn = document.querySelector("#main-ui button[onclick]");
  if (logoutBtn) {
    logoutBtn.onclick = () => {
      window.logoutUser();
    };
  }
});

async function convertToPDF() {
  const fileInput = document.getElementById("uploadDocx");
  const status = document.getElementById("status");
  const file = fileInput.files[0];

  if (!await window.isSessionValid()) {
    status.innerText = "❌ Session expired. Please login again.";
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
