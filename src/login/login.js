import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// Temporary app to read config
async function loadFirebaseConfig() {
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "temp");
  const db = getFirestore(tempApp);
  const snap = await getDoc(doc(db, "config", "firebase"));
  if (!snap.exists()) throw new Error("❌ Firebase config not found.");
  await deleteApp(tempApp);
  return snap.data();
}

// Login user and render UI
async function loginUser(email, password) {
  try {
    const firebaseConfig = await loadFirebaseConfig();
    const app = initializeApp(firebaseConfig);
    const auth = getAuth(app);

    const userCred = await signInWithEmailAndPassword(auth, email, password);
    console.log("✅ Login success:", userCred.user.email);

    const db = getFirestore(app);
    const uiSnap = await getDoc(doc(db, "config", "ui"));
    if (!uiSnap.exists()) throw new Error("❌ UI HTML not found.");
    
    const html = uiSnap.data().html;
    document.open();
    document.write(html);
    document.close();
  } catch (err) {
    console.error(err);
    document.getElementById("status").textContent = "❌ " + err.message;
  }
}

document.getElementById("loginBtn").onclick = () => {
  const email = document.getElementById("emailInput").value;
  const password = document.getElementById("passwordInput").value;
  if (!email || !password) {
    document.getElementById("status").textContent = "❌ Enter both fields";
    return;
  }
  document.getElementById("status").textContent = "🔐 Logging in...";
  loginUser(email, password);
};
