// firebase-auth.js
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
  getDoc,
  deleteDoc
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

window.loginUser = async function(email, password) {
  const status = document.getElementById("login-status");
  try {
    const userCredential = await signInWithEmailAndPassword(auth, email, password);
    const uid = userCredential.user.uid;
    const sessionId = crypto.randomUUID();

    await setDoc(doc(db, "sessions", uid), {
      sessionId,
      timestamp: Date.now()
    });

    localStorage.setItem("uid", uid);
    localStorage.setItem("sessionId", sessionId);

    status.textContent = "✅ Logged in successfully!";
    document.getElementById("login-container").style.display = "none";
    document.getElementById("main-ui").style.display = "block";
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Login failed. Please check your credentials.";
  }
};

window.logoutUser = async function() {
  const uid = localStorage.getItem("uid");
  if (uid) await deleteDoc(doc(db, "sessions", uid));

  localStorage.removeItem("uid");
  localStorage.removeItem("sessionId");

  await signOut(auth);
  document.getElementById("login-container").style.display = "block";
  document.getElementById("main-ui").style.display = "none";
};

window.isSessionValid = async function() {
  const uid = localStorage.getItem("uid");
  const sessionId = localStorage.getItem("sessionId");
  if (!uid || !sessionId) return false;

  try {
    const sessionDoc = await getDoc(doc(db, "sessions", uid));
    if (!sessionDoc.exists()) return false;
    return sessionDoc.data().sessionId === sessionId;
  } catch (err) {
    console.error("Session check failed:", err);
    return false;
  }
};
