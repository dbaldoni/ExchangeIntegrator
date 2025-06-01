/**
 * Authentication Manager
 * Handles OAuth2 authentication flows and token management for Exchange/Office365
 */

class AuthManager {
  constructor() {
    this.oauthFlow = new OAuthFlow();
    this.tokenCache = new Map(); // Cache for access tokens
    this.refreshTokens = new Map(); // Store refresh tokens
    this.tokenExpiryTimes = new Map(); // Track token expiry times
    this.settingsManager = new SettingsManager();
    
    // OAuth2 endpoints for different Exchange environments
    this.authEndpoints = {
      office365: {
        authority: 'https://login.microsoftonline.com/common',
        authorizationEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
        tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        scope: 'https://outlook.office365.com/EWS.AccessAsUser.All offline_access'
      },
      exchange: {
        // For on-premises Exchange with ADFS
        authority: null, // Will be determined during autodiscovery
        authorizationEndpoint: null,
        tokenEndpoint: null,
        scope: 'https://outlook.office365.com/EWS.AccessAsUser.All'
      }
    };
  }

  /**
   * Get OAuth configuration from settings
   */
  async getOAuthCredentials() {
    const oauthConfig = await this.settingsManager.getOAuthConfig();
    return {
      clientId: oauthConfig.clientId,
      tenantId: oauthConfig.tenantId || 'common',
      redirectUri: oauthConfig.redirectUri || this.getRedirectUri()
    };
  }

  /**
   * Get redirect URI for OAuth flow
   */
  getRedirectUri() {
    try {
      if (typeof browser !== 'undefined' && browser.identity && browser.identity.getRedirectURL) {
        return browser.identity.getRedirectURL();
      }
      // Fallback for testing environment
      return 'urn:ietf:wg:oauth:2.0:oob';
    } catch (error) {
      console.warn('Could not get redirect URI:', error);
      return 'urn:ietf:wg:oauth:2.0:oob';
    }
  }

  /**
   * Check if OAuth is properly configured
   */
  async isOAuthConfigured() {
    return await this.settingsManager.isOAuthConfigured();
  }

  /**
   * Authenticate user with Exchange/Office365
   */
  async authenticate(email, password, serverSettings) {
    console.log('Starting authentication for:', email);

    try {
      // Determine authentication method based on server settings
      if (serverSettings.authMethod === 'Modern' || serverSettings.authMethod === 'OAuth2') {
        return await this.authenticateOAuth2(email, serverSettings);
      } else if (serverSettings.authMethod === 'Basic') {
        return await this.authenticateBasic(email, password, serverSettings);
      } else {
        // Try OAuth2 first, fallback to basic auth
        try {
          return await this.authenticateOAuth2(email, serverSettings);
        } catch (error) {
          console.warn('OAuth2 authentication failed, trying basic auth:', error);
          return await this.authenticateBasic(email, password, serverSettings);
        }
      }

    } catch (error) {
      console.error('Authentication failed:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Authenticate using OAuth2/Modern authentication
   */
  async authenticateOAuth2(email, serverSettings) {
    console.log('Performing OAuth2 authentication for:', email);

    try {
      // Check if OAuth is configured
      const isConfigured = await this.isOAuthConfigured();
      if (!isConfigured) {
        throw new Error('OAuth2 is not configured. Please configure Microsoft application credentials in settings.');
      }

      // Get OAuth credentials from settings
      const credentials = await this.getOAuthCredentials();
      
      // Determine OAuth2 endpoints
      const endpoints = await this.getOAuth2Endpoints(serverSettings);
      
      if (!endpoints) {
        throw new Error('Could not determine OAuth2 endpoints');
      }

      // Start OAuth2 flow
      const authResult = await this.oauthFlow.startFlow({
        clientId: credentials.clientId,
        redirectUri: credentials.redirectUri,
        authorizationEndpoint: endpoints.authorizationEndpoint,
        tokenEndpoint: endpoints.tokenEndpoint,
        scope: endpoints.scope,
        responseType: 'code',
        loginHint: email
      });

      if (!authResult.success) {
        throw new Error('OAuth2 flow failed: ' + authResult.error);
      }

      // Store tokens
      this.storeTokens(email, authResult.tokens);

      console.log('OAuth2 authentication successful for:', email);

      return {
        success: true,
        token: authResult.tokens.accessToken,
        refreshToken: authResult.tokens.refreshToken,
        expiresIn: authResult.tokens.expiresIn,
        authMethod: 'OAuth2'
      };

    } catch (error) {
      console.error('OAuth2 authentication failed:', error);
      throw error;
    }
  }

  /**
   * Authenticate using basic authentication
   */
  async authenticateBasic(email, password, serverSettings) {
    console.log('Performing basic authentication for:', email);

    try {
      // Test basic authentication by making a simple EWS request
      const testResult = await this.testBasicAuth(email, password, serverSettings);
      
      if (!testResult.success) {
        throw new Error('Basic authentication failed: Invalid credentials');
      }

      console.log('Basic authentication successful for:', email);

      return {
        success: true,
        token: null, // No token for basic auth
        refreshToken: null,
        authMethod: 'Basic',
        credentials: {
          username: email,
          password: password
        }
      };

    } catch (error) {
      console.error('Basic authentication failed:', error);
      throw error;
    }
  }

  /**
   * Get OAuth2 endpoints based on server settings
   */
  async getOAuth2Endpoints(serverSettings) {
    try {
      // Check if it's Office365
      if (this.isOffice365(serverSettings)) {
        return this.authEndpoints.office365;
      }

      // For on-premises Exchange, try to discover OAuth2 endpoints
      const discoveredEndpoints = await this.discoverOAuth2Endpoints(serverSettings);
      if (discoveredEndpoints) {
        return discoveredEndpoints;
      }

      // Default to Office365 endpoints as fallback
      return this.authEndpoints.office365;

    } catch (error) {
      console.error('Failed to get OAuth2 endpoints:', error);
      return this.authEndpoints.office365;
    }
  }

  /**
   * Check if server is Office365
   */
  isOffice365(serverSettings) {
    const office365Indicators = [
      'outlook.office365.com',
      'outlook.office.com',
      '.microsoftonline.com',
      '.onmicrosoft.com'
    ];

    const serverUrl = serverSettings.serverUrl || serverSettings.ewsUrl || '';
    
    return office365Indicators.some(indicator => 
      serverUrl.toLowerCase().includes(indicator)
    );
  }

  /**
   * Discover OAuth2 endpoints for on-premises Exchange
   */
  async discoverOAuth2Endpoints(serverSettings) {
    try {
      // Try to discover OAuth2 endpoints from server metadata
      const metadataUrl = `${serverSettings.serverUrl}/.well-known/openid_configuration`;
      
      const response = await fetch(metadataUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json'
        },
        signal: AbortSignal.timeout(10000) // 10 second timeout
      });

      if (response.ok) {
        const metadata = await response.json();
        
        return {
          authority: metadata.issuer,
          authorizationEndpoint: metadata.authorization_endpoint,
          tokenEndpoint: metadata.token_endpoint,
          scope: 'https://outlook.office365.com/EWS.AccessAsUser.All'
        };
      }

      return null;

    } catch (error) {
      console.warn('Failed to discover OAuth2 endpoints:', error);
      return null;
    }
  }

  /**
   * Test basic authentication with a simple EWS request
   */
  async testBasicAuth(email, password, serverSettings) {
    try {
      const credentials = btoa(`${email}:${password}`);
      
      // Make a simple GetFolder request to test authentication
      const soapBody = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2016"/>
  </soap:Header>
  <soap:Body xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="inbox"/>
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>`;

      const response = await fetch(serverSettings.ewsUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/xml; charset=utf-8',
          'SOAPAction': '',
          'Authorization': `Basic ${credentials}`,
          'User-Agent': 'ExchangeThunderbirdExtension/1.0'
        },
        body: soapBody,
        signal: AbortSignal.timeout(30000) // 30 second timeout
      });

      return { success: response.ok };

    } catch (error) {
      console.error('Basic auth test failed:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Refresh access token if needed
   */
  async refreshTokenIfNeeded(account) {
    try {
      if (!account.refreshToken) {
        return { success: true }; // No refresh token available
      }

      // Check if token is expired or will expire soon (within 5 minutes)
      const expiryTime = this.tokenExpiryTimes.get(account.email);
      const now = Date.now();
      const fiveMinutes = 5 * 60 * 1000;

      if (!expiryTime || (expiryTime - now) > fiveMinutes) {
        return { success: true }; // Token is still valid
      }

      console.log('Refreshing access token for:', account.email);

      // Get OAuth2 endpoints
      const endpoints = await this.getOAuth2Endpoints(account.serverSettings);
      
      if (!endpoints) {
        throw new Error('Could not determine OAuth2 endpoints for token refresh');
      }

      // Refresh the token
      const refreshResult = await this.oauthFlow.refreshToken({
        clientId: this.clientId,
        tokenEndpoint: endpoints.tokenEndpoint,
        refreshToken: account.refreshToken
      });

      if (!refreshResult.success) {
        throw new Error('Token refresh failed: ' + refreshResult.error);
      }

      // Update account with new tokens
      account.authToken = refreshResult.tokens.accessToken;
      if (refreshResult.tokens.refreshToken) {
        account.refreshToken = refreshResult.tokens.refreshToken;
      }

      // Update cached tokens
      this.storeTokens(account.email, refreshResult.tokens);

      console.log('Token refresh successful for:', account.email);

      return { success: true, newToken: refreshResult.tokens.accessToken };

    } catch (error) {
      console.error('Token refresh failed for:', account.email, error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Store tokens in cache with expiry tracking
   */
  storeTokens(email, tokens) {
    this.tokenCache.set(email, tokens.accessToken);
    
    if (tokens.refreshToken) {
      this.refreshTokens.set(email, tokens.refreshToken);
    }

    // Calculate and store expiry time
    if (tokens.expiresIn) {
      const expiryTime = Date.now() + (tokens.expiresIn * 1000);
      this.tokenExpiryTimes.set(email, expiryTime);
    }
  }

  /**
   * Get cached access token
   */
  getCachedToken(email) {
    return this.tokenCache.get(email);
  }

  /**
   * Get cached refresh token
   */
  getCachedRefreshToken(email) {
    return this.refreshTokens.get(email);
  }

  /**
   * Check if token is expired
   */
  isTokenExpired(email) {
    const expiryTime = this.tokenExpiryTimes.get(email);
    if (!expiryTime) {
      return true; // Assume expired if no expiry time
    }

    return Date.now() >= expiryTime;
  }

  /**
   * Revoke authentication for an account
   */
  async revokeAuthentication(account) {
    try {
      console.log('Revoking authentication for:', account.email);

      // If we have a refresh token, try to revoke it
      if (account.refreshToken) {
        await this.revokeRefreshToken(account);
      }

      // Clear cached tokens
      this.clearCachedTokens(account.email);

      return { success: true };

    } catch (error) {
      console.error('Failed to revoke authentication:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Revoke refresh token
   */
  async revokeRefreshToken(account) {
    try {
      const endpoints = await this.getOAuth2Endpoints(account.serverSettings);
      
      if (!endpoints || !endpoints.tokenEndpoint) {
        return; // Can't revoke without endpoint
      }

      // Microsoft doesn't have a standard revoke endpoint, but we can try
      const revokeEndpoint = endpoints.tokenEndpoint.replace('/token', '/revoke');

      await fetch(revokeEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams({
          token: account.refreshToken,
          client_id: this.clientId
        }),
        signal: AbortSignal.timeout(10000)
      });

      // Don't throw on failure as revoke endpoint might not exist
      console.log('Refresh token revoke request sent');

    } catch (error) {
      console.warn('Failed to revoke refresh token:', error);
    }
  }

  /**
   * Clear cached tokens for an account
   */
  clearCachedTokens(email) {
    this.tokenCache.delete(email);
    this.refreshTokens.delete(email);
    this.tokenExpiryTimes.delete(email);
  }

  /**
   * Get authorization header for API requests
   */
  async getAuthorizationHeader(account) {
    try {
      // Ensure token is fresh
      await this.refreshTokenIfNeeded(account);

      if (account.authToken) {
        return `Bearer ${account.authToken}`;
      } else if (account.authMethod === 'Basic' && account.credentials) {
        const credentials = btoa(`${account.credentials.username}:${account.credentials.password}`);
        return `Basic ${credentials}`;
      } else {
        throw new Error('No valid authentication credentials available');
      }

    } catch (error) {
      console.error('Failed to get authorization header:', error);
      throw error;
    }
  }

  /**
   * Handle authentication errors and retry logic
   */
  async handleAuthenticationError(account, error) {
    console.log('Handling authentication error for:', account.email);

    try {
      // Check if it's a token expiry error
      if (this.isTokenExpiryError(error)) {
        console.log('Token appears to be expired, attempting refresh');
        
        const refreshResult = await this.refreshTokenIfNeeded(account);
        
        if (refreshResult.success) {
          return { success: true, action: 'token_refreshed' };
        } else {
          return { success: false, action: 'reauthentication_required' };
        }
      }

      // Check if it's an invalid credentials error
      if (this.isInvalidCredentialsError(error)) {
        return { success: false, action: 'reauthentication_required' };
      }

      // For other errors, return as-is
      return { success: false, action: 'unknown_error', error: error.message };

    } catch (handlingError) {
      console.error('Failed to handle authentication error:', handlingError);
      return { success: false, action: 'error_handling_failed', error: handlingError.message };
    }
  }

  /**
   * Check if error indicates token expiry
   */
  isTokenExpiryError(error) {
    const expiryIndicators = [
      'token_expired',
      'invalid_token',
      'unauthorized',
      '401',
      'ErrorTokenExpired'
    ];

    const errorMessage = error.message || error.toString();
    
    return expiryIndicators.some(indicator => 
      errorMessage.toLowerCase().includes(indicator.toLowerCase())
    );
  }

  /**
   * Check if error indicates invalid credentials
   */
  isInvalidCredentialsError(error) {
    const credentialIndicators = [
      'invalid_credentials',
      'authentication_failed',
      'invalid_grant',
      'bad_request',
      'ErrorLogonFailure'
    ];

    const errorMessage = error.message || error.toString();
    
    return credentialIndicators.some(indicator => 
      errorMessage.toLowerCase().includes(indicator.toLowerCase())
    );
  }

  /**
   * Validate authentication configuration
   */
  validateAuthConfig() {
    const errors = [];

    if (!this.clientId || this.clientId === 'your-client-id-here') {
      errors.push('Microsoft Client ID is not configured');
    }

    if (!this.redirectUri) {
      errors.push('OAuth2 redirect URI is not available');
    }

    return {
      isValid: errors.length === 0,
      errors: errors
    };
  }

  /**
   * Get authentication status for an account
   */
  getAuthenticationStatus(account) {
    const status = {
      isAuthenticated: false,
      authMethod: account.authMethod || 'Unknown',
      tokenExpired: false,
      needsReauthentication: false
    };

    if (account.authMethod === 'OAuth2') {
      status.isAuthenticated = !!account.authToken;
      status.tokenExpired = this.isTokenExpired(account.email);
      status.needsReauthentication = status.tokenExpired && !account.refreshToken;
    } else if (account.authMethod === 'Basic') {
      status.isAuthenticated = !!(account.credentials && account.credentials.username && account.credentials.password);
    }

    return status;
  }

  /**
   * Clean up resources
   */
  cleanup() {
    this.tokenCache.clear();
    this.refreshTokens.clear();
    this.tokenExpiryTimes.clear();
  }
}
