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

export const CALENDAR_OAUTH_SCOPES = [
  'https://www.googleapis.com/auth/calendar.events',
] as const;

export const googleAuthProvider = new GoogleAuthProvider();
for (const scope of CALENDAR_OAUTH_SCOPES) {
  googleAuthProvider.addScope(scope);
}

// Keep OAuth Access Token strictly in memory (never in localStorage or sessionStorage)
let cachedGoogleAccessToken: string | null = null;

onAuthStateChanged(auth, (user: User | null) => {
  if (!user) {
    cachedGoogleAccessToken = null;
  }
});

export function mapGoogleOAuthErrorToVietnamese(err: unknown): string {
  const code = String((err as { code?: string })?.code || '').toLowerCase();
  const rawMessage = err instanceof Error ? err.message : String(err || '');
  const lower = rawMessage.toLowerCase();

  if (code.includes('popup-blocked') || lower.includes('popup-blocked')) {
    return 'Trình duyệt đang chặn cửa sổ Pop-up đăng nhập Google. Vui lòng bật cho phép Pop-up trên thanh địa chỉ rồi thử lại.';
  }
  if (
    code.includes('popup-closed-by-user') ||
    code.includes('cancelled-popup-request') ||
    lower.includes('popup-closed-by-user')
  ) {
    return 'Bạn đã đóng cửa sổ xác thực Google trước khi hoàn tất cấp quyền.';
  }
  if (
    code.includes('access_denied') ||
    lower.includes('access_denied') ||
    lower.includes('403')
  ) {
    return 'Lỗi 403 (access_denied): Tài khoản Google này chưa được thêm vào Test Users trên Google Cloud Console hoặc chưa cấp quyền Google Calendar.';
  }
  if (code.includes('unauthorized-domain') || lower.includes('unauthorized-domain')) {
    return 'Lỗi cấu hình tên miền (auth/unauthorized-domain): Domain hiện tại chưa được thêm vào Authorized domains trong Firebase Authentication.';
  }
  if (lower.includes('network-request-failed')) {
    return 'Mất kết nối mạng khi xác thực Google. Vui lòng kiểm tra Internet và thử lại.';
  }
  return rawMessage || 'Không thể xác thực Google OAuth 2.0';
}

export async function signInWithGooglePopup(forceConsent = false): Promise<{
  user: User;
  calendarAccessToken: string | null;
}> {
  const provider = new GoogleAuthProvider();
  for (const scope of CALENDAR_OAUTH_SCOPES) {
    provider.addScope(scope);
  }
  if (forceConsent) {
    provider.setCustomParameters({ prompt: 'consent select_account' });
  }
  const result = await signInWithPopup(auth, provider);
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
