/**
 * Exchange Autodiscovery Service
 * Implements the Exchange Autodiscovery protocol to automatically discover server settings
 */

class Autodiscovery {
  constructor() {
    this.xmlParser = new XMLParser();
    this.timeoutMs = 30000; // 30 seconds timeout
  }

  /**
   * Discover Exchange server settings for an email address
   */
  async discover(email, password) {
    console.log('Starting autodiscovery for:', email);
    
    try {
      const domain = this.extractDomain(email);
      
      // Try multiple autodiscovery methods in order of preference
      const methods = [
        () => this.tryAutodiscoverUrl(email, password, `https://autodiscover.${domain}/autodiscover/autodiscover.xml`),
        () => this.tryAutodiscoverUrl(email, password, `https://${domain}/autodiscover/autodiscover.xml`),
        () => this.tryAutodiscoverUrl(email, password, `https://autodiscover.${domain}/Autodiscover/Autodiscover.xml`),
        () => this.tryAutodiscoverUrl(email, password, `https://${domain}/Autodiscover/Autodiscover.xml`),
        () => this.tryDnsRedirect(email, password, domain),
        () => this.tryOffice365Autodiscover(email, password),
        () => this.tryLegacyUrls(email, password, domain)
      ];

      for (const method of methods) {
        try {
          const result = await method();
          if (result && result.success) {
            console.log('Autodiscovery successful using method');
            return result;
          }
        } catch (error) {
          console.log('Autodiscovery method failed:', error.message);
          // Continue to next method
        }
      }

      throw new Error('All autodiscovery methods failed');
    } catch (error) {
      console.error('Autodiscovery failed:', error);
      throw error;
    }
  }

  /**
   * Extract domain from email address
   */
  extractDomain(email) {
    const atIndex = email.indexOf('@');
    if (atIndex === -1) {
      throw new Error('Invalid email address');
    }
    return email.substring(atIndex + 1);
  }

  /**
   * Try autodiscovery with a specific URL
   */
  async tryAutodiscoverUrl(email, password, url) {
    console.log('Trying autodiscovery URL:', url);

    const autodiscoverRequest = this.buildAutodiscoverRequest(email);
    
    try {
      const response = await this.makeAutodiscoverRequest(url, autodiscoverRequest, email, password);
      
      if (response.ok) {
        const responseText = await response.text();
        return this.parseAutodiscoverResponse(responseText);
      } else {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Autodiscovery failed for URL ${url}:`, error);
      throw error;
    }
  }

  /**
   * Try DNS-based autodiscovery redirect
   */
  async tryDnsRedirect(email, password, domain) {
    console.log('Trying DNS redirect autodiscovery for domain:', domain);
    
    try {
      // Attempt to resolve autodiscover CNAME record
      // This is a simplified approach since we can't do actual DNS lookups in a browser extension
      const redirectDomain = `autodiscover.${domain}`;
      const url = `https://${redirectDomain}/autodiscover/autodiscover.xml`;
      
      return await this.tryAutodiscoverUrl(email, password, url);
    } catch (error) {
      throw new Error(`DNS redirect autodiscovery failed: ${error.message}`);
    }
  }

  /**
   * Try Office365 specific autodiscovery
   */
  async tryOffice365Autodiscover(email, password) {
    console.log('Trying Office365 autodiscovery');
    
    const domain = this.extractDomain(email);
    
    // Check if it's a known Office365 domain
    const office365Domains = [
      'outlook.com', 'hotmail.com', 'live.com', 'msn.com',
      'office365.com', 'onmicrosoft.com'
    ];

    const isO365Domain = office365Domains.some(o365Domain => 
      domain.toLowerCase().includes(o365Domain.toLowerCase())
    );

    if (isO365Domain || await this.checkIfOffice365(domain)) {
      return {
        success: true,
        serverSettings: {
          serverUrl: 'outlook.office365.com',
          ewsUrl: 'https://outlook.office365.com/EWS/Exchange.asmx',
          authMethod: 'Modern', // OAuth2
          serverType: 'Exchange2016', // Office365 uses Exchange 2016 API
          autodiscoverUrl: 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml'
        }
      };
    }

    throw new Error('Not an Office365 domain');
  }

  /**
   * Check if domain is hosted on Office365
   */
  async checkIfOffice365(domain) {
    try {
      // Try to detect Office365 by making a test request
      const testUrl = `https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml`;
      const autodiscoverRequest = this.buildAutodiscoverRequest(`test@${domain}`);
      
      const response = await fetch(testUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/xml; charset=utf-8',
          'User-Agent': 'ExchangeThunderbirdExtension/1.0'
        },
        body: autodiscoverRequest,
        signal: AbortSignal.timeout(5000) // 5 second timeout
      });

      // If we get any response (not necessarily successful), it might be Office365
      return response.status !== 404;
    } catch (error) {
      return false;
    }
  }

  /**
   * Try legacy/fallback URLs
   */
  async tryLegacyUrls(email, password, domain) {
    console.log('Trying legacy autodiscovery URLs');
    
    const legacyUrls = [
      `http://autodiscover.${domain}/autodiscover/autodiscover.xml`,
      `http://${domain}/autodiscover/autodiscover.xml`,
      `https://mail.${domain}/autodiscover/autodiscover.xml`,
      `https://webmail.${domain}/autodiscover/autodiscover.xml`
    ];

    for (const url of legacyUrls) {
      try {
        return await this.tryAutodiscoverUrl(email, password, url);
      } catch (error) {
        // Continue to next URL
      }
    }

    throw new Error('All legacy autodiscovery URLs failed');
  }

  /**
   * Build autodiscovery XML request
   */
  buildAutodiscoverRequest(email) {
    return `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
  <Request>
    <EMailAddress>${this.escapeXml(email)}</EMailAddress>
    <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
  </Request>
</Autodiscover>`;
  }

  /**
   * Make autodiscovery HTTP request
   */
  async makeAutodiscoverRequest(url, requestBody, email, password) {
    const credentials = btoa(`${email}:${password}`);
    
    return await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/xml; charset=utf-8',
        'User-Agent': 'ExchangeThunderbirdExtension/1.0',
        'Authorization': `Basic ${credentials}`,
        'SOAPAction': ''
      },
      body: requestBody,
      signal: AbortSignal.timeout(this.timeoutMs)
    });
  }

  /**
   * Parse autodiscovery XML response
   */
  parseAutodiscoverResponse(responseXml) {
    try {
      console.log('Parsing autodiscovery response');
      
      const doc = this.xmlParser.parseXML(responseXml);
      
      // Check for errors in response
      const errorElements = doc.getElementsByTagName('Error');
      if (errorElements.length > 0) {
        const errorCode = errorElements[0].getAttribute('Code') || 'Unknown';
        const errorMessage = errorElements[0].textContent || 'Unknown error';
        throw new Error(`Autodiscovery error ${errorCode}: ${errorMessage}`);
      }

      // Look for Account/Protocol elements
      const protocolElements = doc.getElementsByTagName('Protocol');
      let ewsUrl = null;
      let serverUrl = null;
      let authMethod = 'Basic';
      let serverType = 'Exchange';

      for (let i = 0; i < protocolElements.length; i++) {
        const protocol = protocolElements[i];
        const typeElement = protocol.getElementsByTagName('Type')[0];
        
        if (typeElement) {
          const protocolType = typeElement.textContent;
          
          if (protocolType === 'EXPR' || protocolType === 'EXCH') {
            // Exchange protocol
            const ewsElement = protocol.getElementsByTagName('EwsUrl')[0];
            const serverElement = protocol.getElementsByTagName('Server')[0];
            const authElement = protocol.getElementsByTagName('AuthPackage')[0];
            
            if (ewsElement) {
              ewsUrl = ewsElement.textContent;
            }
            
            if (serverElement) {
              serverUrl = serverElement.textContent;
            }
            
            if (authElement) {
              authMethod = authElement.textContent;
            }

            // Determine server type from URL or other indicators
            if (ewsUrl && ewsUrl.includes('office365.com')) {
              serverType = 'Exchange2016'; // Office365
              authMethod = 'Modern'; // OAuth2
            }
            
            break;
          }
        }
      }

      if (!ewsUrl && !serverUrl) {
        throw new Error('No EWS URL found in autodiscovery response');
      }

      // If we have server but no EWS URL, construct it
      if (!ewsUrl && serverUrl) {
        ewsUrl = `https://${serverUrl}/EWS/Exchange.asmx`;
      }

      // If we have EWS URL but no server, extract it
      if (ewsUrl && !serverUrl) {
        const urlParts = ewsUrl.split('/');
        if (urlParts.length >= 3) {
          serverUrl = urlParts[2]; // Extract hostname
        }
      }

      const result = {
        success: true,
        serverSettings: {
          serverUrl: serverUrl,
          ewsUrl: ewsUrl,
          authMethod: authMethod,
          serverType: serverType,
          autodiscoverUrl: null // Will be set by the calling method
        }
      };

      console.log('Autodiscovery successful:', result);
      return result;

    } catch (error) {
      console.error('Failed to parse autodiscovery response:', error);
      throw new Error(`Autodiscovery response parsing failed: ${error.message}`);
    }
  }

  /**
   * Validate discovered settings
   */
  validateSettings(settings) {
    if (!settings.ewsUrl) {
      throw new Error('EWS URL is required');
    }

    if (!settings.serverUrl) {
      throw new Error('Server URL is required');
    }

    // Validate URL format
    try {
      new URL(settings.ewsUrl);
      if (settings.serverUrl.includes('://')) {
        new URL(settings.serverUrl);
      }
    } catch (error) {
      throw new Error('Invalid URL format in server settings');
    }

    return true;
  }

  /**
   * Escape XML special characters
   */
  escapeXml(text) {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /**
   * Get default settings for common providers
   */
  getDefaultSettings(domain) {
    const defaults = {
      'outlook.office365.com': {
        serverUrl: 'outlook.office365.com',
        ewsUrl: 'https://outlook.office365.com/EWS/Exchange.asmx',
        authMethod: 'Modern',
        serverType: 'Exchange2016'
      },
      'outlook.com': {
        serverUrl: 'outlook.office365.com',
        ewsUrl: 'https://outlook.office365.com/EWS/Exchange.asmx',
        authMethod: 'Modern',
        serverType: 'Exchange2016'
      }
    };

    return defaults[domain.toLowerCase()] || null;
  }
}
