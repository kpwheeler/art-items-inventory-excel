import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
} from '@azure/msal-browser';

// Read and write the signed-in user's own OneDrive files; nothing else.
const SCOPES = ['Files.ReadWrite'];

let pca: PublicClientApplication;
let account: AccountInfo | null = null;

export async function initAuth(clientId: string, authority: string): Promise<AccountInfo | null> {
  pca = new PublicClientApplication({
    auth: {
      clientId,
      authority,
      redirectUri: window.location.origin + import.meta.env.BASE_URL,
    },
    cache: { cacheLocation: 'localStorage' },
  });
  await pca.initialize();
  const result = await pca.handleRedirectPromise();
  account = result?.account ?? pca.getAllAccounts()[0] ?? null;
  return account;
}

export function signIn(): Promise<void> {
  return pca.loginRedirect({ scopes: SCOPES });
}

export function signOut(): Promise<void> {
  return pca.logoutRedirect({ account: account ?? undefined });
}

export async function getToken(): Promise<string> {
  if (!account) throw new Error('Not signed in.');
  try {
    const result = await pca.acquireTokenSilent({ scopes: SCOPES, account });
    return result.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      await pca.acquireTokenRedirect({ scopes: SCOPES, account });
    }
    throw e;
  }
}
