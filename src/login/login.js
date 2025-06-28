import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore, doc, getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";
import {
  getAuth, signInWithEmailAndPassword
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";

document.getElementById("loginBtn").onclick = async () => {
  const email = document.getElementById("emailInput").value;
  const password = document.getElementById("passwordInput").value;
  const status = document.getElementById("login-status");

  try {
    const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "temp-login");
    const db = getFirestore(tempApp);

    const configSnap = await getDoc(doc(db, "config", "firebase"));
    if (!configSnap.exists()) throw new Error("Missing config");

    const firebaseConfig = configSnap.data();
    await deleteApp(tempApp);

    const app = initializeApp(firebaseConfig, "real-login");
    const auth = getAuth(app);

    const userCred = await signInWithEmailAndPassword(auth, email, password);
    localStorage.setItem("uid", userCred.user.uid);

    status.textContent = "✅ Logged in!";
    window.location.href = "../ui/taskpane.html";
  } catch (err) {
    console.error(err);
    status.textContent = "❌ " + err.message;
  }
};
