import { getGoogleSSO } from '../store/settings';

/**
 * Google (Gmail / Google Workspace) sign-in via Google Identity Services.
 *
 * One-time setup in Google Cloud Console (console.cloud.google.com):
 *  1. Create/select a project → APIs & Services → OAuth consent screen.
 *  2. Credentials → Create credentials → OAuth client ID → Web application.
 *  3. Authorized JavaScript origins: the exact portal URL (e.g. https://sorrento.cloud).
 *  4. Copy the client ID into Settings → Sign-in (SSO) → Google.
 */

declare global {
  interface Window {
    google?: {
      accounts: {
        oauth2: {
          initTokenClient(config: {
            client_id: string;
            scope: string;
            callback: (resp: { access_token?: string; error?: string; error_description?: string }) => void;
          }): { requestAccessToken(): void };
        };
      };
    };
  }
}

let gsiLoading: Promise<void> | null = null;

function loadGsi(): Promise<void> {
  if (window.google?.accounts) return Promise.resolve();
  if (!gsiLoading) {
    gsiLoading = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'https://accounts.google.com/gsi/client';
      s.async = true;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('Could not load Google Identity Services script.'));
      document.head.appendChild(s);
    });
  }
  return gsiLoading;
}

export async function googleSignIn(): Promise<{ email: string; name: string }> {
  const cfg = getGoogleSSO();
  if (!cfg.enabled || !cfg.clientId) throw new Error('Google sign-in is not configured. Add the OAuth client ID in Settings → Sign-in (SSO).');
  await loadGsi();
  if (!window.google?.accounts) throw new Error('Google Identity Services unavailable.');

  return new Promise((resolve, reject) => {
    const client = window.google!.accounts.oauth2.initTokenClient({
      client_id: cfg.clientId,
      scope: 'openid email profile',
      callback: async (resp) => {
        if (resp.error || !resp.access_token) {
          reject(new Error(resp.error_description || resp.error || 'Google sign-in cancelled.'));
          return;
        }
        try {
          const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
            headers: { authorization: `Bearer ${resp.access_token}` },
          });
          if (!r.ok) throw new Error(`userinfo ${r.status}`);
          const me = await r.json();
          resolve({ email: me.email, name: me.name ?? me.email });
        } catch (e) {
          reject(e instanceof Error ? e : new Error(String(e)));
        }
      },
    });
    client.requestAccessToken();
  });
}
