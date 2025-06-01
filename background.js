/**
 * Syncbird - Background script for Exchange/Office365 Thunderbird extension
 * Handles account management, synchronization coordination, and API communication
 */

class ExchangeExtension {
  constructor() {
    this.accounts = new Map();
    this.syncIntervals = new Map();
    this.settingsManager = new SettingsManager();
    this.authManager = new AuthManager();
    this.exchangeClient = new ExchangeClient();
    this.emailSync = new EmailSync();
    this.contactSync = new ContactSync();
    this.calendarSync = new CalendarSync();
    
    this.init();
  }

  async init() {
    console.log('Exchange Extension: Initializing background script');
    
    // Load existing accounts from storage
    await this.loadAccounts();
    
    // Set up message listeners
    this.setupMessageListeners();
    
    // Start periodic sync for active accounts
    this.startPeriodicSync();
    
    console.log('Exchange Extension: Background script initialized');
  }

  async loadAccounts() {
    try {
      const result = await browser.storage.local.get('exchangeAccounts');
      if (result.exchangeAccounts) {
        for (const account of result.exchangeAccounts) {
          this.accounts.set(account.id, account);
          console.log(`Loaded Exchange account: ${account.email}`);
        }
      }
    } catch (error) {
      console.error('Failed to load accounts:', error);
    }
  }

  async saveAccounts() {
    try {
      const accounts = Array.from(this.accounts.values());
      await browser.storage.local.set({ exchangeAccounts: accounts });
    } catch (error) {
      console.error('Failed to save accounts:', error);
    }
  }

  setupMessageListeners() {
    // Listen for messages from content scripts and other parts of the extension
    browser.runtime.onMessage.addListener(async (message, sender, sendResponse) => {
      try {
        switch (message.action) {
          case 'setupAccount':
            return await this.setupAccount(message.data);
          
          case 'testConnection':
            return await this.testConnection(message.data);
          
          case 'syncNow':
            return await this.syncAccount(message.accountId);
          
          case 'removeAccount':
            return await this.removeAccount(message.accountId);
          
          case 'getAccounts':
            return Array.from(this.accounts.values());
          
          case 'updateSyncSettings':
            return await this.updateSyncSettings(message.accountId, message.settings);
          
          case 'getOAuthConfig':
            return await this.getOAuthConfig();
          
          case 'saveOAuthConfig':
            return await this.saveOAuthConfig(message.data);
          
          default:
            console.warn('Unknown message action:', message.action);
            return { success: false, error: 'Unknown action' };
        }
      } catch (error) {
        console.error('Error handling message:', error);
        return { success: false, error: error.message };
      }
    });
  }

  async setupAccount(accountData) {
    console.log('Setting up Exchange account for:', accountData.email);
    
    try {
      // Step 1: Autodiscovery
      const autodiscovery = new Autodiscovery();
      const serverSettings = await autodiscovery.discover(accountData.email, accountData.password);
      
      if (!serverSettings) {
        throw new Error('Could not discover Exchange server settings');
      }

      // Step 2: Authentication
      const authResult = await this.authManager.authenticate(
        accountData.email,
        accountData.password,
        serverSettings
      );

      if (!authResult.success) {
        throw new Error('Authentication failed: ' + authResult.error);
      }

      // Step 3: Create account object
      const account = {
        id: this.generateAccountId(),
        email: accountData.email,
        displayName: accountData.displayName || accountData.email,
        serverSettings: serverSettings,
        authToken: authResult.token,
        refreshToken: authResult.refreshToken,
        syncSettings: {
          email: true,
          contacts: true,
          calendar: true,
          syncInterval: 300000 // 5 minutes
        },
        lastSync: {
          email: null,
          contacts: null,
          calendar: null
        },
        createdAt: new Date().toISOString()
      };

      // Step 4: Test connection with EWS
      await this.exchangeClient.testConnection(account);

      // Step 5: Create Thunderbird account
      await this.createThunderbirdAccount(account);

      // Step 6: Store account
      this.accounts.set(account.id, account);
      await this.saveAccounts();

      // Step 7: Start initial sync
      await this.syncAccount(account.id);

      console.log('Exchange account setup completed:', account.email);
      return { success: true, accountId: account.id };

    } catch (error) {
      console.error('Account setup failed:', error);
      return { success: false, error: error.message };
    }
  }

  async testConnection(accountData) {
    try {
      const autodiscovery = new Autodiscovery();
      const serverSettings = await autodiscovery.discover(accountData.email, accountData.password);
      
      if (!serverSettings) {
        throw new Error('Could not discover Exchange server settings');
      }

      const authResult = await this.authManager.authenticate(
        accountData.email,
        accountData.password,
        serverSettings
      );

      return { 
        success: true, 
        serverSettings: serverSettings,
        authenticationWorking: authResult.success
      };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  async createThunderbirdAccount(account) {
    try {
      // Create a new mail account in Thunderbird
      const mailAccount = await browser.accounts.create({
        type: 'imap', // We'll simulate IMAP behavior via EWS
        name: account.displayName,
        identities: [{
          name: account.displayName,
          email: account.email,
          replyTo: account.email,
          signature: ''
        }]
      });

      account.thunderbirdAccountId = mailAccount.id;
      console.log('Created Thunderbird account:', mailAccount.id);
      
    } catch (error) {
      console.error('Failed to create Thunderbird account:', error);
      throw error;
    }
  }

  async syncAccount(accountId) {
    const account = this.accounts.get(accountId);
    if (!account) {
      throw new Error('Account not found');
    }

    console.log('Starting sync for account:', account.email);

    try {
      // Refresh authentication token if needed
      await this.authManager.refreshTokenIfNeeded(account);

      const syncPromises = [];

      // Email sync
      if (account.syncSettings.email) {
        syncPromises.push(
          this.emailSync.sync(account).then(() => {
            account.lastSync.email = new Date().toISOString();
          })
        );
      }

      // Contact sync
      if (account.syncSettings.contacts) {
        syncPromises.push(
          this.contactSync.sync(account).then(() => {
            account.lastSync.contacts = new Date().toISOString();
          })
        );
      }

      // Calendar sync
      if (account.syncSettings.calendar) {
        syncPromises.push(
          this.calendarSync.sync(account).then(() => {
            account.lastSync.calendar = new Date().toISOString();
          })
        );
      }

      // Wait for all sync operations to complete
      await Promise.all(syncPromises);

      // Update account in storage
      this.accounts.set(accountId, account);
      await this.saveAccounts();

      console.log('Sync completed for account:', account.email);
      return { success: true };

    } catch (error) {
      console.error('Sync failed for account:', account.email, error);
      return { success: false, error: error.message };
    }
  }

  async removeAccount(accountId) {
    const account = this.accounts.get(accountId);
    if (!account) {
      return { success: false, error: 'Account not found' };
    }

    try {
      // Stop sync interval
      if (this.syncIntervals.has(accountId)) {
        clearInterval(this.syncIntervals.get(accountId));
        this.syncIntervals.delete(accountId);
      }

      // Remove Thunderbird account
      if (account.thunderbirdAccountId) {
        await browser.accounts.delete(account.thunderbirdAccountId);
      }

      // Remove from storage
      this.accounts.delete(accountId);
      await this.saveAccounts();

      console.log('Removed Exchange account:', account.email);
      return { success: true };

    } catch (error) {
      console.error('Failed to remove account:', error);
      return { success: false, error: error.message };
    }
  }

  async updateSyncSettings(accountId, settings) {
    const account = this.accounts.get(accountId);
    if (!account) {
      return { success: false, error: 'Account not found' };
    }

    account.syncSettings = { ...account.syncSettings, ...settings };
    this.accounts.set(accountId, account);
    await this.saveAccounts();

    // Restart sync interval if interval changed
    if (settings.syncInterval) {
      this.stopSyncInterval(accountId);
      this.startSyncInterval(accountId);
    }

    return { success: true };
  }

  startPeriodicSync() {
    for (const [accountId, account] of this.accounts) {
      this.startSyncInterval(accountId);
    }
  }

  startSyncInterval(accountId) {
    const account = this.accounts.get(accountId);
    if (!account) return;

    if (this.syncIntervals.has(accountId)) {
      clearInterval(this.syncIntervals.get(accountId));
    }

    const interval = setInterval(() => {
      this.syncAccount(accountId).catch(error => {
        console.error('Periodic sync failed:', error);
      });
    }, account.syncSettings.syncInterval);

    this.syncIntervals.set(accountId, interval);
    console.log(`Started sync interval for ${account.email}: ${account.syncSettings.syncInterval}ms`);
  }

  stopSyncInterval(accountId) {
    if (this.syncIntervals.has(accountId)) {
      clearInterval(this.syncIntervals.get(accountId));
      this.syncIntervals.delete(accountId);
    }
  }

  async getOAuthConfig() {
    try {
      const config = await this.settingsManager.getOAuthConfig();
      return { success: true, config: config };
    } catch (error) {
      console.error('Failed to get OAuth config:', error);
      return { success: false, error: error.message };
    }
  }

  async saveOAuthConfig(data) {
    try {
      const result = await this.settingsManager.updateOAuthCredentials(data.clientId, data.tenantId);
      return result;
    } catch (error) {
      console.error('Failed to save OAuth config:', error);
      return { success: false, error: error.message };
    }
  }

  generateAccountId() {
    return 'exchange_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }
}

// Initialize the extension when the background script loads
const exchangeExtension = new ExchangeExtension();

// Handle extension startup
browser.runtime.onStartup.addListener(() => {
  console.log('Exchange Extension: Extension started');
});

// Handle extension installation
browser.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    console.log('Exchange Extension: Extension installed');
    // Open setup page on first install
    browser.tabs.create({ url: browser.runtime.getURL('content/account-setup.html') });
  }
});
