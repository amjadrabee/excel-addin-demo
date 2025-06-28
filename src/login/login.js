// /login/login.js
import { loginUser, isSessionValid } from "../firebase-auth.js";
import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

let app, db;

async function initFirebaseFromFirestore() {
  const tempApp = initializeApp({ projectId: "excel-addin-auth" });
  const tempDb = getFirestore(tempApp);

  const configDoc = await getDoc(doc(tempDb, "config", "firebase"));
  if (!configDoc.exists()) {
    throw new Error("❌ Firebase config not found in Firestore");
  }

  const firebaseConfig = configDoc.data();
  app = initializeApp(firebaseConfig, "main");
  db = getFirestore(app);
}

document.getElementById("loginBtn").onclick = async () => {
  const status = document.getElementById("login-status");
  try {
    if (!app) await initFirebaseFromFirestore();

    const email = document.getElementById("email").value;
    const password = document.getElementById("password").value;

    await loginUser(email, password);

    const valid = await isSessionValid();
    if (valid) {
      const urlsDoc = await getDoc(doc(db, "config", "urls"));
      const uiUrl = urlsDoc.exists() ? urlsDoc.data().ui : null;
      if (uiUrl) {
        window.location.href = uiUrl;
      } else {
        status.textContent = "❌ UI URL not found.";
      }
    } else {
      status.textContent = "❌ Login failed.";
    }
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Error initializing login.";
  }
};
