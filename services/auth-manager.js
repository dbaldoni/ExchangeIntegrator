// services/auth-manager.js

/**
 * AuthManager per Exchange/Office365 con il "public client" di Thunderbird.
 *  - Non richiede più clientId/secret dall’utente: usa sempre
 *    CLIENT_ID_TB = "9e5f94bc-e8a4-4e73-b8be-63364c29d753".
 *  - Implementa il flusso OAuth2 standard con PKCE (code_verifier/code_challenge).
 *  - Scambia l'authorization_code via token endpoint senza client_secret.
 *  - Salva i token (access_token, refresh_token, expires_on) in browser.storage.local.
 */

const CLIENT_ID_TB = "9e5f94bc-e8a4-4e73-b8be-63364c29d753";
const TENANT_ENDPOINT = "common"; // "common" per supportare account aziendali, personali e AAD
const OAUTH_STORAGE_KEY = "oauth2_token"; // dove salveremo l'access e refresh token
const OAUTH_SCOPE = [
  "openid",
  "profile",
  "offline_access",
  // Scope per Exchange IMAP/SMTP via Graph:
  "https://outlook.office.com/IMAP.AccessAsUser.All",
  "https://outlook.office.com/SMTP.Send"
].join(" ");


/**
 * Genera un codice "code_verifier" e calcola il corrispondente "code_challenge".
 * Ritorna un oggetto { codeVerifier, codeChallenge }.
 */
async function generatePKCEPair() {
  // 1) Genera un code_verifier (random 128 byte base64-url)
  const array = new Uint8Array(64);
  crypto.getRandomValues(array);
  const codeVerifier = btoa(String.fromCharCode(...array))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");

  // 2) Calcola code_challenge = base64urlEncode( SHA256(code_verifier) )
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
 * Recupera il token OAuth2 da storage (access_token, refresh_token, expires_on).
 * Ritorna l’oggetto salvato o null se non esiste.
 */
async function getCachedToken() {
  const result = await browser.storage.local.get(OAUTH_STORAGE_KEY);
  return result[OAUTH_STORAGE_KEY] || null;
}

/**
 * Salva il token OAuth2 in storage.
 * @param {object} tokenObj — { access_token, refresh_token, expires_in, token_type, scope, expires_on }
 */
async function cacheToken(tokenObj) {
  await browser.storage.local.set({ [OAUTH_STORAGE_KEY]: tokenObj });
}

/**
 * Restituisce un access_token valido:
 *   - Se esiste cachedToken con expires_on > now, ritorna access_token.
 *   - Altrimenti, se esiste refresh_token, chiama refresh() e ritorna il nuovo access_token.
 *   - Se non c’è nulla o refresh fallisce, lancia startAuthFlow() e scambia il code.
 */
async function getAccessToken() {
  let token = await getCachedToken();
  const now = Date.now();

  if (token && token.access_token && token.expires_on && token.expires_on > now) {
    // token ancora valido
    return token.access_token;
  }

  if (token && token.refresh_token) {
    try {
      const newToken = await refreshToken(token.refresh_token);
      return newToken.access_token;
    } catch (err) {
      console.warn("Refresh token failed, starting full auth flow:", err);
    }
  }

  // Se non c’è token valido o refreshToken manca/fallisce, facciamo il full OAuth2 flow
  const code = await startAuthFlow();
  const tokenResponse = await exchangeCodeForToken(code);
  return tokenResponse.access_token;
}

/**
 * Avvia il flusso OAuth2 autorizzativo:
 * 1) Genera code_verifier + code_challenge  
 * 2) Costruisce l’URL di authorize con tutti i parametri e code_challenge  
 * 3) Chiama browser.identity.launchWebAuthFlow per far aprire la finestra di login Microsoft  
 * 4) Intercetta il redirect con "code" nella URL  
 * Ritorna semplicemente l’authorization_code.
 */
async function startAuthFlow() {
  const { codeVerifier, codeChallenge } = await generatePKCEPair();

  // Salviamo temporaneamente il codeVerifier in session storage (per recuperarlo alla token request)
  // Non va in browser.storage.local perché serve solo per pochi secondi
  sessionStorage.setItem("oauth2_code_verifier", codeVerifier);

  // Redirect URI dinamico per estensioni Thunderbird
  const redirectUri = browser.identity.getRedirectURL("oauth-callback.html");

  // Costruzione URL di autorizzazione
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

  // Lancia la finestra di login
  const redirectResponse = await browser.identity.launchWebAuthFlow({
    interactive: true,
    url: authUrl
  });

  // Parsed redirect URL es.: "moz-extension://<UUID>/oauth-callback.html?code=XYZ&session_state=..."
  const urlObj = new URL(redirectResponse);
  const code = urlObj.searchParams.get("code");
  if (!code) {
    throw new Error("Authorization code not returned from redirect.");
  }
  return code;
}

/**
 * Scambia l’authorization_code per un access_token + refresh_token:
 *  POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
 *  Content-Type: application/x-www-form-urlencoded
 *  Body: client_id, grant_type=authorization_code, code, redirect_uri, code_verifier, scope
 */
async function exchangeCodeForToken(code) {
  const redirectUri = browser.identity.getRedirectURL("oauth-callback.html");
  const tokenEndpoint = `https://login.microsoftonline.com/${TENANT_ENDPOINT}/oauth2/v2.0/token`;

  // Riprendiamo il code_verifier dalla sessionStorage
  const codeVerifier = sessionStorage.getItem("oauth2_code_verifier");
  if (!codeVerifier) {
    throw new Error("PKCE code_verifier not found in sessionStorage.");
  }

  // Preparo il payload urlencoded
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
  // Calcoliamo expires_on = ora corrente + expires_in*1000
  tokenResp.expires_on = Date.now() + tokenResp.expires_in * 1000;

  // Salviamo in cache
  await cacheToken(tokenResp);

  // Rimuoviamo il code_verifier da sessionStorage per sicurezza
  sessionStorage.removeItem("oauth2_code_verifier");

  return tokenResp;
}

/**
 * Scambia un refresh_token per un nuovo access_token (grazie a grant_type=refresh_token).
 * Endpoint: POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
 * Body (x-www-form-urlencoded): client_id, grant_type=refresh_token, refresh_token, scope
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

// Esportiamo solo la funzione getAccessToken (tutto il resto è "interno")
export {
  getAccessToken
};


