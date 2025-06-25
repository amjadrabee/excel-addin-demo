import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc,
  updateDoc,
  deleteField
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

const sessionId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;

window.loginUser = async () => {
  const email = document.getElementById("emailInput").value;
  const password = document.getElementById("passwordInput").value;
  const status = document.getElementById("login-status");

  try {
    const userCred = await signInWithEmailAndPassword(auth, email, password);
    const user = userCred.user;

    // Save session to Firestore
    await setDoc(doc(db, "sessions", user.uid), {
      activeSession: sessionId,
      timestamp: Date.now()
    });

    localStorage.setItem("sessionId", sessionId);

    status.textContent = "✅ Logged in successfully!";
    document.getElementById("login-container").style.display = "none";
    document.getElementById("main-ui").style.display = "block";
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Login failed. Please check your credentials.";
  }
};

// ✅ Validate session on every page load
onAuthStateChanged(auth, async (user) => {
  if (user) {
    const sessionRef = doc(db, "sessions", user.uid);
    const sessionSnap = await getDoc(sessionRef);

    const currentSessionId = localStorage.getItem("sessionId");

    if (sessionSnap.exists() && sessionSnap.data().activeSession !== currentSessionId) {
      // Another session is active
      alert("⚠️ Your account was logged in from another device.");
      await signOut(auth);
      localStorage.removeItem("sessionId");
      location.reload();
    } else {
      document.getElementById("login-container").style.display = "none";
      document.getElementById("main-ui").style.display = "block";
    }
  }
});

window.logoutUser = async () => {
  const user = auth.currentUser;
  if (user) {
    await updateDoc(doc(db, "sessions", user.uid), {
      activeSession: deleteField()
    });
  }
  await signOut(auth);
  localStorage.removeItem("sessionId");
  location.reload();
};
