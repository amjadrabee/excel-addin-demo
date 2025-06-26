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

export async function loginUser(email, password) {
  const status = document.getElementById("login-status");

  try {
    // Step 1: Sign in user (get UID)
    const userCredential = await signInWithEmailAndPassword(auth, email, password);
    const uid = userCredential.user.uid;

    // Step 2: Check existing session before continuing
    const sessionDocRef = doc(db, "sessions", uid);
    const sessionDoc = await getDoc(sessionDocRef);
    const existingSession = sessionDoc.exists() ? sessionDoc.data().sessionId : null;

    if (existingSession) {
      // Active session already exists
      await signOut(auth);
      status.textContent = "❌ This account is already in use elsewhere.";
      return;
    }

    // Step 3: Register new session
    const sessionId = crypto.randomUUID();
    await setDoc(sessionDocRef, {
      sessionId,
      timestamp: Date.now()
    });

    // Step 4: Save session locally and update UI
    localStorage.setItem("uid", uid);
    localStorage.setItem("sessionId", sessionId);

    status.textContent = "✅ Logged in successfully!";
    document.getElementById("login-container").style.display = "none";
    document.getElementById("main-ui").style.display = "block";
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Login failed. Please check your credentials.";
  }
}

export async function logoutUser() {
  const email = localStorage.getItem("emailForSignIn");
  const message = `Logout request for user ${email}`;

  const mailtoLink = `mailto:aecoresolutions@gmail.com?subject=Logout Request&body=${encodeURIComponent(message)}`;
  window.open(mailtoLink, "_blank");

  // Optional: alert the user
  // alert("Logout request sent to admin. Access will be revoked by admin.");
}


export async function isSessionValid() {
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
}
