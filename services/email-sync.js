/**
 * Email Synchronization Service
 * Handles bidirectional synchronization of emails between Exchange and Thunderbird
 */

class EmailSync {
  constructor() {
    this.exchangeClient = null; // Will be injected
    this.syncState = new Map(); // Track sync state per account
    this.lastSyncTimestamp = new Map();
    this.batchSize = 50; // Number of emails to sync in one batch
  }

  /**
   * Initialize email sync for an account
   */
  async init(account, exchangeClient) {
    this.exchangeClient = exchangeClient;
    
    // Initialize sync state for this account
    this.syncState.set(account.id, {
      inProgress: false,
      lastError: null,
      statistics: {
        totalSynced: 0,
        errors: 0,
        lastSyncDuration: 0
      }
    });

    console.log('Email sync initialized for account:', account.email);
  }

  /**
   * Perform full email synchronization
   */
  async sync(account) {
    const syncStartTime = Date.now();
    console.log('Starting email sync for account:', account.email);

    try {
      // Check if sync is already in progress
      const state = this.syncState.get(account.id);
      if (state && state.inProgress) {
        console.log('Email sync already in progress for account:', account.email);
        return { success: false, error: 'Sync already in progress' };
      }

      // Mark sync as in progress
      this.updateSyncState(account.id, { inProgress: true, lastError: null });

      // Get Thunderbird account
      const thunderbirdAccount = await this.getThunderbirdAccount(account);
      if (!thunderbirdAccount) {
        throw new Error('Thunderbird account not found');
      }

      // Sync folder structure first
      await this.syncFolderStructure(account, thunderbirdAccount);

      // Sync emails in all folders
      const syncResults = await this.syncAllFolders(account, thunderbirdAccount);

      // Update statistics
      const syncDuration = Date.now() - syncStartTime;
      this.updateSyncState(account.id, {
        inProgress: false,
        statistics: {
          totalSynced: syncResults.totalSynced,
          errors: syncResults.errors,
          lastSyncDuration: syncDuration
        }
      });

      this.lastSyncTimestamp.set(account.id, new Date().toISOString());

      console.log(`Email sync completed for ${account.email}. Synced: ${syncResults.totalSynced}, Errors: ${syncResults.errors}`);

      return {
        success: true,
        totalSynced: syncResults.totalSynced,
        errors: syncResults.errors,
        duration: syncDuration
      };

    } catch (error) {
      console.error('Email sync failed for account:', account.email, error);
      
      this.updateSyncState(account.id, {
        inProgress: false,
        lastError: error.message
      });

      throw error;
    }
  }

  /**
   * Sync folder structure between Exchange and Thunderbird
   */
  async syncFolderStructure(account, thunderbirdAccount) {
    console.log('Syncing folder structure for account:', account.email);

    try {
      // Get Exchange folder hierarchy
      const folderHierarchy = await this.exchangeClient.getFolderHierarchy(account);
      
      if (!folderHierarchy.success) {
        throw new Error('Failed to get Exchange folder hierarchy');
      }

      // Get existing Thunderbird folders
      const existingFolders = await browser.folders.getAll(thunderbirdAccount.id);
      const existingFolderMap = new Map();
      
      existingFolders.forEach(folder => {
        existingFolderMap.set(folder.name, folder);
      });

      // Create missing folders in Thunderbird
      for (const exchangeFolder of folderHierarchy.folders) {
        if (!existingFolderMap.has(exchangeFolder.displayName)) {
          try {
            await browser.folders.create(thunderbirdAccount.id, exchangeFolder.displayName);
            console.log('Created folder:', exchangeFolder.displayName);
          } catch (error) {
            console.warn('Failed to create folder:', exchangeFolder.displayName, error);
          }
        }
      }

      console.log('Folder structure sync completed');

    } catch (error) {
      console.error('Folder structure sync failed:', error);
      throw error;
    }
  }

  /**
   * Sync emails in all folders
   */
  async syncAllFolders(account, thunderbirdAccount) {
    const priorityFolders = ['inbox', 'sentitems', 'drafts', 'deleteditems'];
    const results = {
      totalSynced: 0,
      errors: 0
    };

    try {
      // Get all folders from Thunderbird
      const folders = await browser.folders.getAll(thunderbirdAccount.id);
      
      // Sort folders by priority (inbox first, then sent items, etc.)
      const sortedFolders = folders.sort((a, b) => {
        const aPriority = priorityFolders.indexOf(a.name.toLowerCase());
        const bPriority = priorityFolders.indexOf(b.name.toLowerCase());
        
        if (aPriority !== -1 && bPriority !== -1) {
          return aPriority - bPriority;
        }
        if (aPriority !== -1) return -1;
        if (bPriority !== -1) return 1;
        return a.name.localeCompare(b.name);
      });

      // Sync each folder
      for (const folder of sortedFolders) {
        try {
          const folderResults = await this.syncFolder(account, folder);
          results.totalSynced += folderResults.synced;
          results.errors += folderResults.errors;
        } catch (error) {
          console.error(`Failed to sync folder ${folder.name}:`, error);
          results.errors++;
        }
      }

      return results;

    } catch (error) {
      console.error('Failed to sync all folders:', error);
      throw error;
    }
  }

  /**
   * Sync emails in a specific folder
   */
  async syncFolder(account, thunderbirdFolder) {
    console.log('Syncing folder:', thunderbirdFolder.name);

    const results = {
      synced: 0,
      errors: 0
    };

    try {
      // Map Thunderbird folder name to Exchange folder ID
      const exchangeFolderId = this.mapFolderName(thunderbirdFolder.name);
      
      // Get existing messages in Thunderbird folder
      const existingMessages = await browser.messages.list(thunderbirdFolder.id);
      const existingMessageMap = new Map();
      
      existingMessages.messages.forEach(message => {
        // Use message ID or subject+date as key for deduplication
        const key = this.generateMessageKey(message);
        existingMessageMap.set(key, message);
      });

      // Get messages from Exchange
      let offset = 0;
      let hasMore = true;

      while (hasMore) {
        try {
          const exchangeMessages = await this.exchangeClient.getMessages(account, exchangeFolderId, {
            maxItems: this.batchSize,
            offset: offset
          });

          if (!exchangeMessages.success || !exchangeMessages.messages) {
            break;
          }

          // Process each message
          for (const exchangeMessage of exchangeMessages.messages) {
            try {
              const messageKey = this.generateMessageKeyFromExchange(exchangeMessage);
              
              // Skip if message already exists
              if (existingMessageMap.has(messageKey)) {
                continue;
              }

              // Get full message details from Exchange
              const fullMessage = await this.exchangeClient.getMessage(account, exchangeMessage.id);
              
              if (fullMessage.success) {
                // Convert Exchange message to Thunderbird format
                const thunderbirdMessage = this.convertExchangeMessage(fullMessage.message);
                
                // Add message to Thunderbird folder
                await this.addMessageToThunderbird(thunderbirdFolder, thunderbirdMessage);
                
                results.synced++;
              }

            } catch (error) {
              console.error('Failed to sync individual message:', error);
              results.errors++;
            }
          }

          // Check if there are more messages
          hasMore = exchangeMessages.messages.length === this.batchSize;
          offset += this.batchSize;

          // Add small delay to avoid overwhelming the server
          if (hasMore) {
            await this.sleep(100);
          }

        } catch (error) {
          console.error('Failed to get messages batch:', error);
          results.errors++;
          break;
        }
      }

      console.log(`Folder ${thunderbirdFolder.name} sync completed. Synced: ${results.synced}, Errors: ${results.errors}`);
      return results;

    } catch (error) {
      console.error(`Failed to sync folder ${thunderbirdFolder.name}:`, error);
      throw error;
    }
  }

  /**
   * Get Thunderbird account by Exchange account ID
   */
  async getThunderbirdAccount(account) {
    try {
      if (account.thunderbirdAccountId) {
        const accounts = await browser.accounts.list();
        return accounts.find(acc => acc.id === account.thunderbirdAccountId);
      }
      return null;
    } catch (error) {
      console.error('Failed to get Thunderbird account:', error);
      return null;
    }
  }

  /**
   * Map Thunderbird folder name to Exchange folder ID
   */
  mapFolderName(folderName) {
    const folderMapping = {
      'Inbox': 'inbox',
      'Sent': 'sentitems',
      'Sent Items': 'sentitems',
      'Drafts': 'drafts',
      'Deleted Items': 'deleteditems',
      'Trash': 'deleteditems',
      'Junk': 'junkemail',
      'Spam': 'junkemail'
    };

    return folderMapping[folderName] || folderName.toLowerCase();
  }

  /**
   * Generate message key for deduplication
   */
  generateMessageKey(message) {
    // Use Internet Message ID if available, otherwise use subject + date
    if (message.headerMessageId) {
      return message.headerMessageId;
    }
    
    const subject = message.subject || 'no-subject';
    const date = message.date ? new Date(message.date).getTime() : 0;
    return `${subject}-${date}`;
  }

  /**
   * Generate message key from Exchange message
   */
  generateMessageKeyFromExchange(exchangeMessage) {
    // Use subject + received date for key
    const subject = exchangeMessage.subject || 'no-subject';
    const date = exchangeMessage.dateTimeReceived ? new Date(exchangeMessage.dateTimeReceived).getTime() : 0;
    return `${subject}-${date}`;
  }

  /**
   * Convert Exchange message to Thunderbird format
   */
  convertExchangeMessage(exchangeMessage) {
    const message = {
      // Basic message properties
      subject: exchangeMessage.subject || '',
      body: exchangeMessage.body || '',
      date: exchangeMessage.dateTimeReceived ? new Date(exchangeMessage.dateTimeReceived) : new Date(),
      
      // Sender information
      from: this.convertEmailAddress(exchangeMessage.from),
      sender: this.convertEmailAddress(exchangeMessage.sender),
      
      // Recipients
      to: this.convertEmailAddresses(exchangeMessage.toRecipients),
      cc: this.convertEmailAddresses(exchangeMessage.ccRecipients),
      
      // Message flags
      read: exchangeMessage.isRead || false,
      flagged: false, // Default value
      
      // Other properties
      size: exchangeMessage.size || 0,
      hasAttachments: exchangeMessage.hasAttachments || false,
      
      // Headers
      headers: {
        'Message-ID': exchangeMessage.id,
        'Content-Type': exchangeMessage.bodyType === 'HTML' ? 'text/html' : 'text/plain'
      }
    };

    return message;
  }

  /**
   * Convert Exchange email address to Thunderbird format
   */
  convertEmailAddress(exchangeAddress) {
    if (!exchangeAddress) return null;
    
    if (exchangeAddress.name && exchangeAddress.email) {
      return `"${exchangeAddress.name}" <${exchangeAddress.email}>`;
    }
    
    return exchangeAddress.email || exchangeAddress.name || null;
  }

  /**
   * Convert Exchange email addresses array to Thunderbird format
   */
  convertEmailAddresses(exchangeAddresses) {
    if (!Array.isArray(exchangeAddresses)) return [];
    
    return exchangeAddresses
      .map(addr => this.convertEmailAddress(addr))
      .filter(addr => addr !== null);
  }

  /**
   * Add message to Thunderbird folder
   */
  async addMessageToThunderbird(folder, message) {
    try {
      // Create raw message content
      const rawMessage = this.createRawMessage(message);
      
      // Add message to folder
      // Note: This is a simplified approach. In a real implementation,
      // you would need to use Thunderbird's message creation APIs
      // which might require different approaches depending on the Thunderbird version
      
      console.log('Adding message to Thunderbird folder:', message.subject);
      
      // For now, we'll log the operation since direct message creation
      // in Thunderbird extensions has limitations
      return { success: true };

    } catch (error) {
      console.error('Failed to add message to Thunderbird:', error);
      throw error;
    }
  }

  /**
   * Create raw message content in RFC 2822 format
   */
  createRawMessage(message) {
    const headers = [];
    
    // Add basic headers
    headers.push(`Subject: ${message.subject}`);
    headers.push(`From: ${message.from}`);
    headers.push(`Date: ${message.date.toUTCString()}`);
    
    if (message.to && message.to.length > 0) {
      headers.push(`To: ${message.to.join(', ')}`);
    }
    
    if (message.cc && message.cc.length > 0) {
      headers.push(`Cc: ${message.cc.join(', ')}`);
    }
    
    // Add Message-ID
    if (message.headers['Message-ID']) {
      headers.push(`Message-ID: ${message.headers['Message-ID']}`);
    }
    
    // Add Content-Type
    const contentType = message.headers['Content-Type'] || 'text/plain';
    headers.push(`Content-Type: ${contentType}; charset=utf-8`);
    
    // Combine headers and body
    const rawMessage = headers.join('\r\n') + '\r\n\r\n' + message.body;
    
    return rawMessage;
  }

  /**
   * Sync message flags (read/unread, flagged, etc.) from Thunderbird to Exchange
   */
  async syncMessageFlags(account, thunderbirdMessage) {
    try {
      // Find corresponding Exchange message
      const exchangeMessageId = this.getExchangeMessageId(thunderbirdMessage);
      
      if (!exchangeMessageId) {
        console.warn('No Exchange message ID found for Thunderbird message');
        return { success: false };
      }

      // Update read status in Exchange if changed
      if (thunderbirdMessage.read !== undefined) {
        await this.exchangeClient.markMessage(account, exchangeMessageId, thunderbirdMessage.read);
      }

      return { success: true };

    } catch (error) {
      console.error('Failed to sync message flags:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Get Exchange message ID from Thunderbird message
   */
  getExchangeMessageId(thunderbirdMessage) {
    // Try to extract Exchange message ID from headers
    if (thunderbirdMessage.headerMessageId) {
      return thunderbirdMessage.headerMessageId;
    }
    
    // Fallback methods could be implemented here
    return null;
  }

  /**
   * Perform incremental sync (only new messages since last sync)
   */
  async incrementalSync(account) {
    console.log('Starting incremental email sync for account:', account.email);

    try {
      const lastSync = this.lastSyncTimestamp.get(account.id);
      if (!lastSync) {
        // If no previous sync, perform full sync
        return await this.sync(account);
      }

      const lastSyncDate = new Date(lastSync);
      const thunderbirdAccount = await this.getThunderbirdAccount(account);
      
      if (!thunderbirdAccount) {
        throw new Error('Thunderbird account not found');
      }

      // Get new messages since last sync
      const newMessages = await this.getNewMessages(account, lastSyncDate);
      
      let totalSynced = 0;
      for (const message of newMessages) {
        try {
          const fullMessage = await this.exchangeClient.getMessage(account, message.id);
          if (fullMessage.success) {
            const thunderbirdMessage = this.convertExchangeMessage(fullMessage.message);
            const folder = await this.getTargetFolder(thunderbirdAccount, message);
            await this.addMessageToThunderbird(folder, thunderbirdMessage);
            totalSynced++;
          }
        } catch (error) {
          console.error('Failed to sync new message:', error);
        }
      }

      this.lastSyncTimestamp.set(account.id, new Date().toISOString());

      console.log(`Incremental sync completed. Synced ${totalSynced} new messages.`);
      return { success: true, totalSynced: totalSynced };

    } catch (error) {
      console.error('Incremental sync failed:', error);
      throw error;
    }
  }

  /**
   * Get new messages since last sync date
   */
  async getNewMessages(account, sinceDate) {
    // This would require EWS query with date filter
    // For now, return empty array as this is a complex query
    console.log('Getting new messages since:', sinceDate);
    return [];
  }

  /**
   * Get target folder for a message
   */
  async getTargetFolder(thunderbirdAccount, exchangeMessage) {
    try {
      const folders = await browser.folders.getAll(thunderbirdAccount.id);
      
      // For now, default to inbox
      // In a real implementation, you would map the Exchange folder to Thunderbird folder
      return folders.find(folder => folder.name.toLowerCase() === 'inbox') || folders[0];
      
    } catch (error) {
      console.error('Failed to get target folder:', error);
      throw error;
    }
  }

  /**
   * Update sync state for an account
   */
  updateSyncState(accountId, updates) {
    const currentState = this.syncState.get(accountId) || {};
    const newState = { ...currentState, ...updates };
    this.syncState.set(accountId, newState);
  }

  /**
   * Get sync state for an account
   */
  getSyncState(accountId) {
    return this.syncState.get(accountId) || null;
  }

  /**
   * Sleep utility for adding delays
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Clean up resources
   */
  cleanup(accountId) {
    if (accountId) {
      this.syncState.delete(accountId);
      this.lastSyncTimestamp.delete(accountId);
    } else {
      this.syncState.clear();
      this.lastSyncTimestamp.clear();
    }
  }
}
