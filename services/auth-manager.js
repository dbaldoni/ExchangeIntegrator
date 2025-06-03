// services/auth-manager.js

/**
 * AuthManager
 * 
 * Gestisce l’OAuth2 per Office365/Exchange via browser.identity.
 * Legge i parametri clientId, tenantId e clientSecret da browser.storage.local → "oauthConfig".
 * Espone:
 *   - getOAuthConfig(): Promise<{clientId, tenantId, clientSecret}>
 *   - getAccessToken(): Promise<{ access_token, expires_in, token_type, refresh_token }>
 *   - startAuthFlow(): Promise<string>  (ritorna auth code)
 *   - refreshAccessToken(refreshToken): Promise<{ access_token, expires_in, token_type, refresh_token }>
 */

const STORAGE_KEY = "oauthConfig";
const TOKEN_STORAGE_KEY = "oauthToken"; // dove salviamo access/refresh token

/**
 * Recupera la configurazione OAuth2 (clientId, tenantId, clientSecret) da storage.
 * Se non trovata, ritorna un oggetto vuoto.
 */
async function getOAuthConfig() {
  let result = await browser.storage.local.get(STORAGE_KEY);
  if (result[STORAGE_KEY]) {
    return result[STORAGE_KEY];
  }
  return {}; // se non presente, ritorna {}
}

/**
 * Recupera dal storage il token salvato (con refresh_token, expires_at, ecc.)
 * Ritorna un oggetto token o null se non esiste.
 */
async function getSavedToken() {
  let result = await browser.storage.local.get(TOKEN_STORAGE_KEY);
  return result[TOKEN_STORAGE_KEY] || null;
}

/**
 * Salva in storage i token OAuth (access_token, refresh_token, expires_at, ecc.)
 * @param {object} tokenObj — { access_token, refresh_token, expires_on, expires_in, token_type }
 */
async function saveToken(tokenObj) {
  await browser.storage.local.set({ [TOKEN_STORAGE_KEY]: tokenObj });
}

/**
 * Costruisce l’URL di autorizzazione per Microsoft OAuth2
 *   per ottenere il codice authorization_code.
 */
function buildAuthUrl({ clientId, tenantId, redirectUri, scope }) {
  const base = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;
  const params = new URLSearchParams({
    client_id: clientId,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: scope, // es. "openid profile offline_access https://graph.microsoft.com/.default"
  });
  return `${base}?${params.toString()}`;
}

/**
 * Avvia il flusso OAuth interattivo per ottenere il code di autorizzazione.
 * Usa browser.identity.launchWebAuthFlow per aprire la pagina di login Microsoft.
 * Ritorna il codice (authorization_code) dalla query string del redirectUri.
 */
async function startAuthFlow() {
  // 1) Recuperiamo configurazione
  let { clientId, tenantId, clientSecret } = await getOAuthConfig();
  if (!clientId || !tenantId || !clientSecret) {
    throw new Error("OAuth2 not configured: missing clientId, tenantId or clientSecret.");
  }

  // 2) Determiniamo il redirectUri che abbiamo registrato su Azure
  //    Per estensioni Thunderbird, di solito è di forma:
  //      moz-extension://<UUID>/oauth-callback.html
  //    Qui usiamo browser.identity.getRedirectURL() per ottenerlo dinamicamente.
  let redirectUri = browser.identity.getRedirectURL("oauth-callback.html");

  // 3) Scope consigliato per Graph API oppure OWA:
  //    - "openid profile offline_access"
  //    - Aggiungi "https://graph.microsoft.com/.default" per ottenere il token per Graph.
  let scope = "openid profile offline_access https://graph.microsoft.com/.default";

  let authUrl = buildAuthUrl({ clientId, tenantId, redirectUri, scope });

  // 4) Launch Web Auth Flow
  let redirectResponse;
  try {
    redirectResponse = await browser.identity.launchWebAuthFlow({
      interactive: true,
      url: authUrl
    });
  } catch (err) {
    throw new Error("User cancelled login or error launching auth flow: " + err.message);
  }

  // 5) Estrarre il parametro code dalla redirectResponse (es. "https://<extension>/oauth-callback.html?code=abcd&state=xyz")
  let urlObj = new URL(redirectResponse);
  let code = urlObj.searchParams.get("code");
  if (!code) {
    throw new Error("Authorization code not found in redirect response.");
  }

  return code;
}

/**
 * Scambia l'authorization code per un access token + refresh token.
 * Chiama l’endpoint token di Microsoft:
 *   POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
 * Con body x-www-form-urlencoded:
 *   client_id, scope, code, redirect_uri, grant_type=authorization_code, client_secret
 */
async function requestTokenWithCode(code) {
  let { clientId, tenantId, clientSecret } = await getOAuthConfig();
  let redirectUri = browser.identity.getRedirectURL("oauth-callback.html");
  let tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  let body = new URLSearchParams({
    client_id: clientId,
    scope: "openid profile offline_access https://graph.microsoft.com/.default",
    code: code,
    redirect_uri: redirectUri,
    grant_type: "authorization_code",
    client_secret: clientSecret
  });

  let resp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body
  });

  if (!resp.ok) {
    let text = await resp.text();
    throw new Error(`Token request failed: ${resp.status} ${text}`);
  }

  let tokenResponse = await resp.json();
  /**
   * tokenResponse contiene almeno:
   *   - access_token
   *   - refresh_token
   *   - expires_in (secondi)
   *   - ext_expires_in
   *   - token_type
   *   - scope
   */
  // Calcoliamo expires_on come ora corrente + expires_in*1000 (millisecondi)
  let now = Date.now();
  tokenResponse.expires_on = now + tokenResponse.expires_in * 1000;

  // Salviamo il token completo
  await saveToken(tokenResponse);

  return tokenResponse;
}

/**
 * Scambia un refresh token per un nuovo access token + refresh token.
 * Endpoint token di Microsoft con grant_type=refresh_token.
 */
async function refreshAccessToken(refreshToken) {
  let { clientId, tenantId, clientSecret } = await getOAuthConfig();
  let tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  let body = new URLSearchParams({
    client_id: clientId,
    scope: "openid profile offline_access https://graph.microsoft.com/.default",
    refresh_token: refreshToken,
    grant_type: "refresh_token",
    client_secret: clientSecret
  });

  let resp = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body
  });

  if (!resp.ok) {
    let text = await resp.text();
    throw new Error(`Refresh token request failed: ${resp.status} ${text}`);
  }

  let tokenResponse = await resp.json();
  let now = Date.now();
  tokenResponse.expires_on = now + tokenResponse.expires_in * 1000;
  await saveToken(tokenResponse);
  return tokenResponse;
}

/**
 * Restituisce un access token valido:
 *  1) Se in storage c'è un token con expires_on > Date.now(), ritorna subito quell'access_token.  
 *  2) Se in storage esiste un refresh_token senza scadenza, chiama `refreshAccessToken(...)` e ritorna il nuovo access_token.  
 *  3) Se non esiste nulla o non è valido, lancia `startAuthFlow()` per ottenere un nuovo code, poi chiama `requestTokenWithCode(...)`.  
 */
async function getAccessToken() {
  let token = await getSavedToken();
  let now = Date.now();

  if (token && token.access_token && token.expires_on && token.expires_on > now) {
    // Il token è ancora valido
    return token.access_token;
  }

  if (token && token.refresh_token) {
    // Proviamo a usare il refresh token
    try {
      let newTokenResp = await refreshAccessToken(token.refresh_token);
      return newTokenResp.access_token;
    } catch (err) {
      console.warn("Failed to refresh token, will start auth flow:", err);
    }
  }

  // Se non c'è token oppure non si riesce a refreshare, iniziamo il flow interattivo
  let code = await startAuthFlow();
  let tokenResp = await requestTokenWithCode(code);
  return tokenResp.access_token;
}

// Esportiamo le funzioni principali
export {
  getOAuthConfig,
  getAccessToken,
  startAuthFlow,
  requestTokenWithCode,
  refreshAccessToken
};

