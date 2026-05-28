import { initializeApp, getApps, getApp } from 'firebase/app';
import { initializeFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyAvf72Sg-1aqe8hOXCtkqHUYU8GHhjsg00",
  authDomain: "ql-yt-database.firebaseapp.com",
  projectId: "ql-yt-database",
  storageBucket: "ql-yt-database.firebasestorage.app",
  messagingSenderId: "674009076264",
  appId: "1:674009076264:web:1f64c1a0fa90b80e8a1b1d",
  measurementId: "G-SVVD1KPEKQ"
};

let app = null;
let firestore = null;
let isCloudActive = false;

try {
  if (getApps().length === 0) {
    app = initializeApp(firebaseConfig);
  } else {
    app = getApp();
  }
  
  // Initialize firestore with long-polling settings for compatibility with Electron's file:// protocol
  firestore = initializeFirestore(app, {
    experimentalForceLongPolling: true,
    useFetchStreams: false
  });
  
  isCloudActive = true;
  console.log('🔥 Firebase Cloud Mode is ACTIVE (Long-polling enabled for Electron)');
} catch (err) {
  console.error('Error initializing Firebase:', err);
}

export { app, firestore, isCloudActive };
