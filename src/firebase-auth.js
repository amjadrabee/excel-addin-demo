// firebase-auth.js
import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc,
  updateDoc
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
export const auth = getAuth(app);
const db = getFirestore(app);

export async function loginUser(email, password) {
  const status = document.getElementById("login-status");
  try {
    const result = await signInWithEmailAndPassword(auth, email, password);
    const sessionId = crypto.randomUUID();
    await setDoc(doc(db, "sessions", email), { sessionId });
    localStorage.setItem("sessionId", sessionId);
    localStorage.setItem("userEmail", email);
    status.textContent = "✅ Logged in successfully!";
    document.getElementById("login-container").style.display = "none";
    document.getElementById("main-ui").style.display = "block";
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Login failed.";
  }
}

export async function logoutUser() {
  try {
    const email = localStorage.getItem("userEmail");
    if (email) {
      await updateDoc(doc(db, "sessions", email), { sessionId: "" });
    }
    localStorage.clear();
    await signOut(auth);
    document.getElementById("main-ui").style.display = "none";
    document.getElementById("login-container").style.display = "block";
  } catch (err) {
    console.error("Logout failed", err);
  }
}

export async function isSessionValid() {
  const email = localStorage.getItem("userEmail");
  const localSessionId = localStorage.getItem("sessionId");
  if (!email || !localSessionId) return false;
  try {
    const docRef = doc(db, "sessions", email);
    const docSnap = await getDoc(docRef);
    return docSnap.exists() && docSnap.data().sessionId === localSessionId;
  } catch (err) {
    console.error("Session check error", err);
    return false;
  }
}
