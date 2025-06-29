export async function convertToPDF() {
  const fileInput = document.getElementById("uploadDocx");
  const status = document.getElementById("status");
  const file = fileInput?.files?.[0];

  if (!file) {
    status.innerText = "❌ Select a .docx file.";
    return;
  }

  try {
    status.innerText = "🔄 Fetching API key...";
    const { getAuth } = await import("https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js");
    const { getFirestore, doc, getDoc } = await import("https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js");

    const auth = getAuth();
    const db = getFirestore();
    const user = auth.currentUser;
    if (!user) throw new Error("❌ Not logged in.");

    const keySnap = await getDoc(doc(db, "config", "cloudconvert"));
    if (!keySnap.exists()) throw new Error("❌ API key not found.");
    const apiKey = keySnap.data().key;

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

    const job = await jobRes.json();
    const uploadTask = Object.values(job.data.tasks).find(t => t.operation === "import/upload");

    const formData = new FormData();
    for (const key in uploadTask.result.form.parameters) {
      formData.append(key, uploadTask.result.form.parameters[key]);
    }
    formData.append("file", file);
    await fetch(uploadTask.result.form.url, { method: "POST", body: formData });

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
    status.innerText = "❌ Conversion failed. Check the console.";
  }
}

export async function handleLogoutRequest() {
  const { logoutRequestLocal } = await import("../firebase-auth.js");
  const user = localStorage.getItem("uid") || "Unknown";
  try {
    await fetch("https://your-logout-email-service.com/send", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        to: "support@yourcompany.com",
        subject: "Logout Request",
        message: `${user} has requested to log out from the Excel Add-in.`
      })
    });
    await logoutRequestLocal();
    alert("📩 Logout request sent.");
  } catch (err) {
    console.error("Logout email error:", err);
    alert("❌ Failed to send logout request.");
  }
}
