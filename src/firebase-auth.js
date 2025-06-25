const firebaseConfig = {
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_PROJECT.firebaseapp.com",
  projectId: "YOUR_PROJECT",
  storageBucket: "YOUR_PROJECT.appspot.com",
  messagingSenderId: "YOUR_SENDER_ID",
  appId: "YOUR_APP_ID"
};

firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.firestore();

window.loginUser = async function (email, password) {
  const status = document.getElementById("login-status");
  try {
    const userCredential = await auth.signInWithEmailAndPassword(email, password);
    const user = userCredential.user;

    const sessionId = Date.now().toString();
    await db.collection("sessions").doc(user.uid).set({ sessionId });
    localStorage.setItem("sessionId", sessionId);
    localStorage.setItem("userId", user.uid);

    status.textContent = "✅ Logged in!";
    document.getElementById("login-container").style.display = "none";
    document.getElementById("main-ui").style.display = "block";
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Login failed.";
  }
};

window.logoutUser = async function () {
  localStorage.clear();
  await auth.signOut();
  document.getElementById("login-container").style.display = "block";
  document.getElementById("main-ui").style.display = "none";
};

window.isSessionValid = async function () {
  const userId = localStorage.getItem("userId");
  const localSession = localStorage.getItem("sessionId");
  if (!userId || !localSession) return false;

  const doc = await db.collection("sessions").doc(userId).get();
  return doc.exists && doc.data().sessionId === localSession;
};
