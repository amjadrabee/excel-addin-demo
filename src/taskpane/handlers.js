/* ── Logout request: opens default mail client ── */
export function handleLogoutRequest() {
  const email   = localStorage.getItem("email") || "Unknown";
  const subject = encodeURIComponent("Logout Request");
  const body    = encodeURIComponent(`${email} has requested to log out from the Excel Add‑in.`);
  window.location.href = `mailto:support@yourcompany.com?subject=${subject}&body=${body}`;
}

/* ── Convert DOCX→PDF using CloudConvert ── */
export async function convertToPDF() {
  const status = document.getElementById("status");
  const file   = document.getElementById("uploadDocx")?.files[0];

  if (!file || !file.name.toLowerCase().endsWith(".docx")) {
    status.textContent = "❌ Select a .docx file.";
    return;
  }

  try {
    status.textContent = "🔑 Getting API key…";
    const { getFirestore, doc, getDoc } = await import("https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js");
    const keySnap = await getDoc(doc(getFirestore(), "config", "cloudconvert"));
    if (!keySnap.exists()) throw new Error("API key missing.");
    const apiKey = keySnap.data().key;

    /* create job */
    status.textContent = "📄 Uploading…";
    const job = await fetch("https://api.cloudconvert.com/v2/jobs", {
      method : "POST",
      headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
      body   : JSON.stringify({
        tasks: {
          upload : { operation: "import/upload" },
          convert: { operation: "convert", input: "upload", input_format: "docx", output_format: "pdf" },
          export : { operation: "export/url", input: "convert" }
        }
      })
    }).then(r => r.json());

    const uploadTask = Object.values(job.data.tasks).find(t => t.operation === "import/upload");
    const formData   = new FormData();
    Object.entries(uploadTask.result.form.parameters).forEach(([k,v]) => formData.append(k,v));
    formData.append("file", file);
    await fetch(uploadTask.result.form.url, { method: "POST", body: formData });

    status.textContent = "⏳ Converting…";
    let exportTask;
    while (!exportTask) {
      await new Promise(r => setTimeout(r, 2500));
      const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
        headers: { Authorization: `Bearer ${apiKey}` }
      }).then(r => r.json());
      if (poll.data.status === "finished") {
        exportTask = poll.data.tasks.find(t => t.name === "export");
      }
    }
    const pdfUrl = exportTask.result.files[0].url;
    status.innerHTML = `✅ Done! <a href="${pdfUrl}" target="_blank">Download PDF</a>`;
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Conversion failed.";
  }
}
