import { initializeApp, getApps } from 'firebase/app';
import {
  getAuth,
  GoogleAuthProvider,
  onAuthStateChanged,
  signInWithPopup,
  type User,
} from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json' with { type: 'json' };

const app = getApps().length ? getApps()[0] : initializeApp(firebaseConfig);
export const auth = getAuth(app);

export const googleAuthProvider = new GoogleAuthProvider();
googleAuthProvider.addScope('https://www.googleapis.com/auth/calendar.events');

// Keep OAuth Access Token strictly in memory (never in localStorage or sessionStorage)
let cachedGoogleAccessToken: string | null = null;

onAuthStateChanged(auth, (user: User | null) => {
  if (!user) {
    cachedGoogleAccessToken = null;
  }
});

export async function signInWithGooglePopup(): Promise<{
  user: User;
  calendarAccessToken: string | null;
}> {
  const result = await signInWithPopup(auth, googleAuthProvider);
  const credential = GoogleAuthProvider.credentialFromResult(result);
  cachedGoogleAccessToken = credential?.accessToken || null;
  return {
    user: result.user,
    calendarAccessToken: cachedGoogleAccessToken,
  };
}

export function getCachedCalendarAccessToken(): string | null {
  return cachedGoogleAccessToken;
}

export function clearCachedCalendarAccessToken(): void {
  cachedGoogleAccessToken = null;
}
