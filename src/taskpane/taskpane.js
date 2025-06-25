import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";

import {
  getFirestore,
  doc,
  setDoc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyCjB5shAXVySxyEXiBfQNx3ifBHs0tGSq0",
  authDomain: "excel-addin-auth.firebaseapp.com",
  projectId: "excel-addin-auth",
  storageBucket: "excel-addin-auth.appspot.com",
  messagingSenderId: "1051103393339",
  appId: "1:1051103393339:web:9f89eda79f1698b25dce1e"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

window.loginUser = async () => {
  const email = document.getElementById("emailInput").value;
  const password = document.getElementById("passwordInput").value;
  const status = document.getElementById("login-status");

  try {
    const result = await signInWithEmailAndPassword(auth, email, password);
    const sessionId = crypto.randomUUID();

    await setDoc(doc(db, "sessions", result.user.uid), {
      sessionId: sessionId
    });

    localStorage.setItem("sessionId", sessionId);
    localStorage.setItem("uid", result.user.uid);
    localStorage.setItem("emailForSignIn", email);

    status.textContent = "✅ Logged in successfully!";
    document.getElementById("login-container").style.display = "none";
    document.getElementById("main-ui").style.display = "block";
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Login failed. Please check your credentials.";
  }
};

window.logoutUser = async () => {
  try {
    await signOut(auth);
    localStorage.removeItem("sessionId");
    localStorage.removeItem("uid");
    localStorage.removeItem("emailForSignIn");
    document.getElementById("main-ui").style.display = "none";
    document.getElementById("login-container").style.display = "block";
  } catch (err) {
    console.error(err);
    document.getElementById("login-status").textContent = "❌ Logout failed.";
  }
};

export async function isSessionValid() {
  const uid = localStorage.getItem("uid");
  const sessionId = localStorage.getItem("sessionId");

  if (!uid || !sessionId) return false;

  try {
    const docRef = doc(db, "sessions", uid);
    const docSnap = await getDoc(docRef);

    if (docSnap.exists()) {
      const storedSessionId = docSnap.data().sessionId;
      return storedSessionId === sessionId;
    } else {
      return false;
    }
  } catch (err) {
    console.error("Error checking session:", err);
    return false;
  }
}

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
