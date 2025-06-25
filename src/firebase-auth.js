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
