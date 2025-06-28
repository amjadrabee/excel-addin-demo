import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
    import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

    const firebaseConfig = {
      apiKey: "AIzaSyCjB5shAXVySxyEXiBfQNx3ifBHs0tGSq0",
      authDomain: "excel-addin-auth.firebaseapp.com",
      projectId: "excel-addin-auth"
    };

    const app = initializeApp(firebaseConfig);
    const db = getFirestore(app);

    try {
      const configRef = doc(db, "config", "urls");
      const configSnap = await getDoc(configRef);

      if (configSnap.exists()) {
        const url = configSnap.data().taskpane;
        window.location.replace(url);
      } else {
        document.body.innerHTML = "❌ Taskpane URL not found in database.";
      }
    } catch (err) {
      document.body.innerHTML = "❌ Failed to load taskpane URL.";
      console.error(err);
    }