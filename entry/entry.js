import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

async function loadFirebaseConfig() {
  const tempApp = initializeApp({ projectId: "excel-addin-auth" }, "tempApp");
  const db = getFirestore(tempApp);
  const configRef = doc(db, "config", "firebase");
  const configSnap = await getDoc(configRef);

  if (!configSnap.exists()) throw new Error("Missing firebase config in Firestore");
  const config = configSnap.data();

  // Cleanup temp app
  tempApp.delete?.();

  // Init actual app
  const app = initializeApp(config);
  return app;
}

// Usage:
loadFirebaseConfig().then(app => {
  console.log("Firebase initialized:", app.name);
  // Now you can use `getAuth(app)` or `getFirestore(app)` safely
}).catch(err => {
  console.error("❌ Failed to load Firebase config:", err);
});
