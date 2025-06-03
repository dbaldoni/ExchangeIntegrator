// services/auth-manager.js

/**
 * AuthManager per Exchange/365: usa il Client ID pubblico di Thunderbird (trusted).
 * Chiede solo scope Graph Mail.ReadWrite e Mail.Send (non più IMAP/SMTP).
 */

const CLIENT_ID_TB = "9e5f94bc-e8a4-4e73-b8be-63364c29d753";
const TENANT_ENDPOINT = "common"; // “common” permette account AAD e Microsoft personali.
const OAUTH_STORAGE_KEY = "oauth2_token";
const OAUTH_SCOPE = [
  "openid",
  "profile",
  "offline_access",
  // Cambiati gli scope, ora chiediamo Graph Mail.ReadWrite e Mail.Send
  "https://graph.microsoft.com/Mail.ReadWrite",
  "https://graph.microsoft.com/Mail.Send"
].join(" ");

/**
 * Genera un code_verifier e il relativo code_challenge (PKCE).
 */
async function generatePKCEPair() {
  const array = new Uint8Array(64);
  crypto.getRandomValues(array);
  const codeVerifier = btoa(String.fromCharCode(...array))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");

  const buffer = new TextEncoder().encode(codeVerifier);
  const digest = await crypto.subtle.digest("SHA-256", buffer);
  const hashArray = Array.from(new Uint8Array(digest));
  const base64Hash = btoa(String.fromCharCode(...hashArray))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");

  const codeChallenge = base64Hash;
  return { codeVerifier, codeChallenge };
}

/**
 * Recupera il token da browser.storage.local (se esiste).
 */
async function getCachedToken() {
  const result = await browser.storage.local.get(OAUTH_STORAGE_KEY);
  return result[OAUTH_STORAGE_KEY] || null;
}

/**
 * Salva il token in storage.
 */
async function cacheToken(tokenObj) {
  await browser.storage.local.set({ [OAUTH_STORAGE_KEY]: tokenObj });
}

/**
 * Restituisce un access_token valido:
 *  - Se il token salvato non è scaduto, lo restituisce.  
 *  - Altrimenti prova a usare refresh_token.  
 *  - Se manca o non funziona, lancia startAuthFlow() e scambia il code.
 */
async function getAccessToken() {
  let token = await getCachedToken();
  const now = Date.now();

  if (token && token.access_token && token.expires_on > now) {
    return token.access_token;
  }

  if (token && token.refresh_token) {
    try {
      const newToken = await refreshToken(token.refresh_token);
      return newToken.access_token;
    } catch (err) {
      console.warn("Refresh token failed, doing full auth flow:", err);
    }
  }

  const code = await startAuthFlow();
  const tokenResp = await exchangeCodeForToken(code);
  return tokenResp.access_token;
}

/**
 * Avvia il flusso OAuth2 con PKCE:
 *  1. genera code_verifier/challenge,  
 *  2. costruisce l’URL di authorize con CLIENT_ID_TB e code_challenge,  
 *  3. lancia browser.identity.launchWebAuthFlow → ottiene redirect con code.
 */
async function startAuthFlow() {
  const { codeVerifier, codeChallenge } = await generatePKCEPair();

  // Salvo temporaneo del code_verifier in sessionStorage
  sessionStorage.setItem("oauth2_code_verifier", codeVerifier);

  const redirectUri = browser.identity.getRedirectURL("oauth-callback.html");
  const params = new URLSearchParams({
    client_id: CLIENT_ID_TB,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: OAUTH_SCOPE,
    code_challenge: codeChallenge,
    code_challenge_method: "S256"
  });

  const authUrl = `https://login.microsoftonline.com/${TENANT_ENDPOINT}/oauth2/v2.0/authorize?${params.toString()}`;

  const redirectResponse = await browser.identity.launchWebAuthFlow({
    interactive: true,
    url: authUrl
  });

  const urlObj = new URL(redirectResponse);
  const code = urlObj.searchParams.get("code");
  if (!code) {
    throw new Error("Authorization code not returned");
  }
  return code;
}

/**
 * Scambia l’authorization_code con access_token + refresh_token:
 *  POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token  
 *  Body form-url-encoded: client_id, grant_type=authorization_code, code,
 *  redirect_uri, code_verifier, scope.
 */
async function exchangeCodeForToken(code) {
  const redirectUri = browser.identity.getRedirectURL("oauth-callback.html");
  const tokenEndpoint = `https://login.microsoftonline.com/${TENANT_ENDPOINT}/oauth2/v2.0/token`;

  const codeVerifier = sessionStorage.getItem("oauth2_code_verifier");
  if (!codeVerifier) {
    throw new Error("PKCE code_verifier missing");
  }

  const bodyParams = new URLSearchParams({
    client_id: CLIENT_ID_TB,
    grant_type: "authorization_code",
    code: code,
    redirect_uri: redirectUri,
    code_verifier: codeVerifier,
    scope: OAUTH_SCOPE
  });

  const resp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: bodyParams.toString()
  });

  if (!resp.ok) {
    const txt = await resp.text();
    throw new Error(`Token exchange failed ${resp.status}: ${txt}`);
  }

  const tokenResp = await resp.json();
  tokenResp.expires_on = Date.now() + tokenResp.expires_in * 1000;
  await cacheToken(tokenResp);
  sessionStorage.removeItem("oauth2_code_verifier");
  return tokenResp;
}

/**
 * Scambia un refresh_token per un nuovo access_token:
 *  POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token  
 *  Body: client_id, grant_type=refresh_token, refresh_token, scope.
 */
async function refreshToken(existingRefreshToken) {
  const tokenEndpoint = `https://login.microsoftonline.com/${TENANT_ENDPOINT}/oauth2/v2.0/token`;

  const bodyParams = new URLSearchParams({
    client_id: CLIENT_ID_TB,
    grant_type: "refresh_token",
    refresh_token: existingRefreshToken,
    scope: OAUTH_SCOPE
  });

  const resp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: bodyParams.toString()
  });

  if (!resp.ok) {
    const txt = await resp.text();
    throw new Error(`Refresh token failed ${resp.status}: ${txt}`);
  }

  const newTokenResp = await resp.json();
  newTokenResp.expires_on = Date.now() + newTokenResp.expires_in * 1000;
  await cacheToken(newTokenResp);
  return newTokenResp;
}

// Esportiamo solo getAccessToken (il resto è "interno").
export {
  getAccessToken
};

