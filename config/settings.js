/**
 * Syncbird Configuration Manager
 * Handles secure storage and retrieval of API credentials and settings
 */

class SettingsManager {
  constructor() {
    this.storageKey = 'syncbird_settings';
    this.encryptionKey = null; // Would be generated securely in production
  }

  /**
   * Get stored settings
   */
  async getSettings() {
    try {
      const result = await browser.storage.local.get(this.storageKey);
      return result[this.storageKey] || this.getDefaultSettings();
    } catch (error) {
      console.error('Failed to get settings:', error);
      return this.getDefaultSettings();
    }
  }

  /**
   * Save settings
   */
  async saveSettings(settings) {
    try {
      await browser.storage.local.set({
        [this.storageKey]: settings
      });
      return { success: true };
    } catch (error) {
      console.error('Failed to save settings:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Get default settings structure
   */
  getDefaultSettings() {
    return {
      oauth: {
        clientId: '',
        tenantId: 'common',
        redirectUri: ''
      },
      sync: {
        interval: 300000, // 5 minutes
        enableEmail: true,
        enableContacts: true,
        enableCalendar: true,
        batchSize: 50
      },
      ui: {
        theme: 'light',
        language: 'en',
        showNotifications: true
      },
      debug: {
        enableLogging: false,
        logLevel: 'info'
      }
    };
  }

  /**
   * Update OAuth credentials
   */
  async updateOAuthCredentials(clientId, tenantId = 'common') {
    const settings = await this.getSettings();
    settings.oauth.clientId = clientId;
    settings.oauth.tenantId = tenantId;
    
    // Set redirect URI based on extension context
    try {
      if (browser.identity && browser.identity.getRedirectURL) {
        settings.oauth.redirectUri = browser.identity.getRedirectURL();
      }
    } catch (error) {
      console.warn('Could not set redirect URI:', error);
    }

    return await this.saveSettings(settings);
  }

  /**
   * Update sync settings
   */
  async updateSyncSettings(syncSettings) {
    const settings = await this.getSettings();
    settings.sync = { ...settings.sync, ...syncSettings };
    return await this.saveSettings(settings);
  }

  /**
   * Check if OAuth is configured
   */
  async isOAuthConfigured() {
    const settings = await this.getSettings();
    return settings.oauth.clientId && settings.oauth.clientId !== '';
  }

  /**
   * Get OAuth configuration
   */
  async getOAuthConfig() {
    const settings = await this.getSettings();
    return settings.oauth;
  }

  /**
   * Clear all settings (for logout/reset)
   */
  async clearSettings() {
    try {
      await browser.storage.local.remove(this.storageKey);
      return { success: true };
    } catch (error) {
      console.error('Failed to clear settings:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Export settings for backup
   */
  async exportSettings() {
    const settings = await this.getSettings();
    // Remove sensitive data for export
    const exportData = { ...settings };
    delete exportData.oauth.clientId;
    return exportData;
  }

  /**
   * Import settings from backup
   */
  async importSettings(importData) {
    try {
      const currentSettings = await this.getSettings();
      const mergedSettings = { ...currentSettings, ...importData };
      return await this.saveSettings(mergedSettings);
    } catch (error) {
      console.error('Failed to import settings:', error);
      return { success: false, error: error.message };
    }
  }
}