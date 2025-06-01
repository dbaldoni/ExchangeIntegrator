/**
 * OAuth2 Flow Handler for Syncbird
 * Handles OAuth2 authentication flows for Exchange/Office365
 */

class OAuthFlow {
  constructor() {
    this.authWindows = new Map();
    this.pendingRequests = new Map();
  }

  /**
   * Start OAuth2 authentication flow
   */
  async startFlow(config) {
    try {
      console.log('Starting OAuth2 flow for:', config.loginHint);

      // Build authorization URL
      const authUrl = this.buildAuthorizationUrl(config);
      
      // For browser testing, we'll simulate the OAuth flow
      // In a real Thunderbird extension, this would use browser.identity.launchWebAuthFlow
      console.log('Authorization URL:', authUrl);
      
      // Simulate successful OAuth flow for testing
      const tokens = await this.simulateOAuthSuccess(config);
      
      return {
        success: true,
        tokens: tokens
      };

    } catch (error) {
      console.error('OAuth2 flow failed:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Build authorization URL
   */
  buildAuthorizationUrl(config) {
    const params = new URLSearchParams({
      client_id: config.clientId,
      response_type: config.responseType || 'code',
      redirect_uri: config.redirectUri,
      scope: config.scope,
      state: this.generateState(),
      response_mode: 'query'
    });

    if (config.loginHint) {
      params.append('login_hint', config.loginHint);
    }

    return `${config.authorizationEndpoint}?${params.toString()}`;
  }

  /**
   * Simulate OAuth success for testing
   */
  async simulateOAuthSuccess(config) {
    // In a real implementation, this would exchange the authorization code for tokens
    return {
      accessToken: `mock_access_token_${Date.now()}`,
      refreshToken: `mock_refresh_token_${Date.now()}`,
      expiresIn: 3600,
      tokenType: 'Bearer'
    };
  }

  /**
   * Refresh OAuth2 token
   */
  async refreshToken(config) {
    try {
      console.log('Refreshing OAuth2 token');

      // For testing, simulate successful token refresh
      const tokens = await this.simulateTokenRefresh(config);

      return {
        success: true,
        tokens: tokens
      };

    } catch (error) {
      console.error('Token refresh failed:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Simulate token refresh for testing
   */
  async simulateTokenRefresh(config) {
    return {
      accessToken: `refreshed_access_token_${Date.now()}`,
      refreshToken: config.refreshToken, // Keep the same refresh token
      expiresIn: 3600,
      tokenType: 'Bearer'
    };
  }

  /**
   * Generate random state parameter
   */
  generateState() {
    return Math.random().toString(36).substring(2, 15) + 
           Math.random().toString(36).substring(2, 15);
  }

  /**
   * Exchange authorization code for tokens
   */
  async exchangeCodeForTokens(config, authorizationCode) {
    try {
      const response = await fetch(config.tokenEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Accept': 'application/json'
        },
        body: new URLSearchParams({
          grant_type: 'authorization_code',
          client_id: config.clientId,
          code: authorizationCode,
          redirect_uri: config.redirectUri,
          scope: config.scope
        })
      });

      if (!response.ok) {
        throw new Error(`Token exchange failed: ${response.status} ${response.statusText}`);
      }

      const tokens = await response.json();
      
      if (tokens.error) {
        throw new Error(`Token exchange error: ${tokens.error_description || tokens.error}`);
      }

      return {
        accessToken: tokens.access_token,
        refreshToken: tokens.refresh_token,
        expiresIn: tokens.expires_in,
        tokenType: tokens.token_type || 'Bearer'
      };

    } catch (error) {
      console.error('Failed to exchange code for tokens:', error);
      throw error;
    }
  }
}