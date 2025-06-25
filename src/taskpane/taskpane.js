Office.onReady(() => {
  const token = localStorage.getItem("emailForSignIn");
  if (!token && !window.location.href.includes("signIn")) {
    document.getElementById("login-container").style.display = "block";
    document.getElementById("main-ui").style.display = "none";
  }

  const convertBtn = document.getElementById("convertBtn");
  if (convertBtn) convertBtn.onclick = convertToPDF;
});

async function checkSessionMatch() {
  const sessionId = localStorage.getItem("sessionId");
  const email = localStorage.getItem("emailForSignIn");

  if (!email || !sessionId) return false;

  try {
    const res = await fetch(`https://excel-addin-auth-default-rtdb.firebaseio.com/sessions/${btoa(email)}.json`);
    const data = await res.json();
    return data && data.sessionId === sessionId;
  } catch (err) {
    console.error("Failed to check session:", err);
    return false;
  }
}

async function convertToPDF() {
  const fileInput = document.getElementById("uploadDocx");
  const status = document.getElementById("status");
  const file = fileInput.files[0];

  if (!file) {
    status.innerText = "❌ Please select a .docx file first.";
    return;
  }

  const sessionValid = await checkSessionMatch();
  if (!sessionValid) {
    status.innerText = "⛔ Invalid session. Please log in again.";
    document.getElementById("login-container").style.display = "block";
    document.getElementById("main-ui").style.display = "none";
    return;
  }

  try {
    status.innerText = "🔄 Creating conversion job...";

    const apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiZGJmZTliNWJiOTBlMTY5OGNlZTdmZmI3NzIyZTg1MjgwMzQyODU3ZTYxY2M1MmZlNDAzMGVmZjdmZGM3YjdmZmRhZmE3NjYyMGQ0ODQyNGIiLCJpYXQiOjE3NTA1OTk2MjYuMjI3MTMyLCJuYmYiOjE3NTA1OTk2MjYuMjI3MTM0LCJleHAiOjQ5MDYyNzMyMjYuMjIzNTEzLCJzdWIiOiI3MjIzNTM3NCIsInNjb3BlcyI6WyJ0YXNrLnJlYWQiLCJ0YXNrLndyaXRlIl19.hmJsJlC28R-XMn8PvV10ZTnZD425ZUh4cBTpIsXMpFJ4FijCqgoibTKofRr7XgXHt8lkeMj9a3IJf8tfuLxx_U8OavhlxMRAG7MPBefKoaJynmp_ZxgI1GVS0W7NLp2jyo3iFnxSQa-H9qkWL8_TTfot72uGQsurnHy3LcwC-hLZfE63YMnndjSn3jFCGSCNOSELIIWosRZNkoHTegdRWL6rcKoIYcfi4Ff7EfwLL2w7-oTnSzICu-Lg-XnmoXRp_jf2AYEpEkvXnlOAKM86pBp3h3DlAcZZOsYr_WIJ8e8Ti4hmlAgwqD9yN76DZeQ65Wq9kBF2LbKiV8fh2WLERHDlA-0OKN92mo4qNvCsfThYgrRXcnNY9oMnYPHewaPwwT6PHv7DJGdLTHZUrtWc6mLLq1XpBzIMEzv1_qbrguot01lwrkkVACDZvnXaUPJ3plM04j6XlKf6qFPyoDNxR-ThbDjxpXqnCM57RDjE1E5ow3cXW2a_5jMFhCjvpMTyUkapzmXPDKF9rYTog4hw5aA6PRzIyS0yMIZB4VaxGM4S-tPwCJNdSdFSz7bssdTVnSf_h0GvlgUcp3en6qT6KfIIMGP7KAGsSZkB339Fu-7U7pJj7wosd1myL1-FXhauZI_YmL5W34SD6tLzvLbEfYfZ0pm9-lsSVXpMqgHlHLM";

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
