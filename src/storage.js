import { doc, getDoc, setDoc } from "firebase/firestore";
import { db } from "./firebase.js";

// Drop-in replacements for the Claude-artifact window.storage.get/set calls.
// Every key is stored as one document in a single "copo_store" collection,
// shared by everyone who opens the deployed site (same as shared:true worked
// in the original artifact version).

const COLLECTION = "copo_store";

export async function sget(key, fallback) {
  try {
    const ref = doc(db, COLLECTION, key);
    const snap = await getDoc(ref);
    return snap.exists() ? JSON.parse(snap.data().value) : fallback;
  } catch (e) {
    console.error("Firestore read failed for", key, e);
    return fallback;
  }
}

export async function sset(key, value) {
  try {
    const ref = doc(db, COLLECTION, key);
    await setDoc(ref, { value: JSON.stringify(value) });
    return true;
  } catch (e) {
    console.error("Firestore write failed for", key, e);
    return false;
  }
}
