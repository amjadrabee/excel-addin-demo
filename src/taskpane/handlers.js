// handlers.js
export async function handleLogoutRequest() {
  try {
    const email = localStorage.getItem("email") || "Unknown";

    // Compose email message
    const subject = encodeURIComponent("Logout Request");
    const body = encodeURIComponent(`${email} wants to log out from the Excel Add-in.`);

    // Open email client
    window.location.href = `mailto:support@yourcompany.com?subject=${subject}&body=${body}`;
  } catch (err) {
    console.error("Logout email error:", err);
    // Fallback: show error in status div (avoiding window.alert)
    const statusBox = document.getElementById("status") || document.getElementById("app");
    if (statusBox) {
      statusBox.innerHTML = `<span style="color: red;">❌ Failed to open email client.</span>`;
    }
  }
}

export async function convertToPDF() {
  const fileInput = document.getElementById("uploadDocx");
  const status = document.getElementById("status");
  const file = fileInput?.files?.[0];

  if (!file || file.name.slice(-5).toLowerCase() !== ".docx") {
    status.innerText = "❌ Please select a .docx file.";
    return;
  }

  try {
    status.innerText = "🔄 Fetching API key...";

    const { getAuth } = await import("https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js");
    const { getFirestore, doc, getDoc } = await import("https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js");

    const auth = getAuth();
    const user = auth.currentUser;
    if (!user) throw new Error("❌ Not signed in.");

    const db = getFirestore();
    const keySnap = await getDoc(doc(db, "config", "cloudconvert"));
    if (!keySnap.exists()) throw new Error("❌ API key not found.");

    const apiKey = keySnap.data().key;

    // Create CloudConvert job
    status.innerText = "🔄 Creating job...";
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

    // Upload file
    status.innerText = "⏫ Uploading file...";
    await fetch(uploadTask.result.form.url, {
      method: "POST",
      body: formData
    });

    // Wait for conversion to complete
    status.innerText = "⏳ Converting...";
    let done = false;
    let exportTask = null;

    while (!done) {
      const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
        headers: { Authorization: `Bearer ${apiKey}` }
      });
      const updatedJob = await poll.json();
      done = updatedJob.data.status === "finished";
      exportTask = updatedJob.data.tasks.find(t => t.name === "export");
      if (!done) await new Promise(resolve => setTimeout(resolve, 3000));
    }

    const fileUrl = exportTask.result.files[0].url;
    status.innerHTML = `✅ Done! <a href="${fileUrl}" target="_blank">Download PDF</a>`;
  } catch (err) {
    console.error(err);
    status.innerText = "❌ Conversion failed. See console for details.";
  }
}
