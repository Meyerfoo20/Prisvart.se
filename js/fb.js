import { firebaseConfig } from "./firebase-config.js";
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.14.1/firebase-app.js";
import { getAuth, onAuthStateChanged, createUserWithEmailAndPassword, signInWithEmailAndPassword, signOut, sendPasswordResetEmail, EmailAuthProvider, reauthenticateWithCredential } from "https://www.gstatic.com/firebasejs/10.14.1/firebase-auth.js";
import { getFirestore, collection, doc, getDoc, setDoc, addDoc, updateDoc, deleteDoc, query, where, onSnapshot, writeBatch, serverTimestamp } from "https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore.js";

export const configured = firebaseConfig.projectId !== "KLISTRA-IN";
const app = configured ? initializeApp(firebaseConfig) : null;
export const auth = configured ? getAuth(app) : null;
export const db = configured ? getFirestore(app) : null;
export { onAuthStateChanged, createUserWithEmailAndPassword, signInWithEmailAndPassword, signOut, sendPasswordResetEmail, EmailAuthProvider, reauthenticateWithCredential, collection, doc, getDoc, setDoc, addDoc, updateDoc, deleteDoc, query, where, onSnapshot, writeBatch, serverTimestamp };
