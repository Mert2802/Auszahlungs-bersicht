import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js";
import {
  getAuth,
  onAuthStateChanged,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut
} from "https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  getDocs,
  collection,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyA_3oYbI399yyA5pr2jSj5qy6IQdOxfjew",
  authDomain: "cockpit-8b6e7.firebaseapp.com",
  projectId: "cockpit-8b6e7",
  storageBucket: "cockpit-8b6e7.firebasestorage.app",
  messagingSenderId: "438555216400",
  appId: "1:438555216400:web:348d52349a27a069015895",
  measurementId: "G-0M6341D5D2"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

window.firebaseServices = {
  app,
  auth,
  db,
  onAuthStateChanged,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut,
  doc,
  getDoc,
  setDoc,
  getDocs,
  collection,
  serverTimestamp
};
