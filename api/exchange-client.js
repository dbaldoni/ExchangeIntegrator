/**
 * Exchange Client - Main API interface for Exchange/Office365 communication
 * Handles high-level operations and coordinates between different services
 */

class ExchangeClient {
  constructor() {
    this.ewsClient = new EWSClient();
    this.baseRetryDelay = 1000; // 1 second
    this.maxRetries = 3;
  }

  /**
   * Test connection to Exchange server
   */
  async testConnection(account) {
    try {
      console.log('Testing Exchange connection for:', account.email);
      
      // Test basic connectivity with GetFolder operation
      const result = await this.ewsClient.getFolder(account, 'inbox');
      
      if (result && result.success) {
        console.log('Exchange connection test successful');
        return { success: true };
      } else {
        throw new Error('Connection test failed');
      }
    } catch (error) {
      console.error('Exchange connection test failed:', error);
      throw new Error(`Connection test failed: ${error.message}`);
    }
  }

  /**
   * Get folder information
   */
  async getFolder(account, folderId) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.getFolder(account, folderId);
    });
  }

  /**
   * Get folder hierarchy
   */
  async getFolderHierarchy(account) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.getFolderHierarchy(account);
    });
  }

  /**
   * Get messages from a folder
   */
  async getMessages(account, folderId, options = {}) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.getMessages(account, folderId, options);
    });
  }

  /**
   * Get message details
   */
  async getMessage(account, messageId) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.getMessage(account, messageId);
    });
  }

  /**
   * Send a message
   */
  async sendMessage(account, message) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.sendMessage(account, message);
    });
  }

  /**
   * Move message to folder
   */
  async moveMessage(account, messageId, targetFolderId) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.moveMessage(account, messageId, targetFolderId);
    });
  }

  /**
   * Mark message as read/unread
   */
  async markMessage(account, messageId, isRead) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.markMessage(account, messageId, isRead);
    });
  }

  /**
   * Delete message
   */
  async deleteMessage(account, messageId) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.deleteMessage(account, messageId);
    });
  }

  /**
   * Get contacts
   */
  async getContacts(account, options = {}) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.getContacts(account, options);
    });
  }

  /**
   * Create contact
   */
  async createContact(account, contact) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.createContact(account, contact);
    });
  }

  /**
   * Update contact
   */
  async updateContact(account, contactId, updates) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.updateContact(account, contactId, updates);
    });
  }

  /**
   * Delete contact
   */
  async deleteContact(account, contactId) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.deleteContact(account, contactId);
    });
  }

  /**
   * Get calendar items
   */
  async getCalendarItems(account, startDate, endDate, options = {}) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.getCalendarItems(account, startDate, endDate, options);
    });
  }

  /**
   * Create calendar item
   */
  async createCalendarItem(account, item) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.createCalendarItem(account, item);
    });
  }

  /**
   * Update calendar item
   */
  async updateCalendarItem(account, itemId, updates) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.updateCalendarItem(account, itemId, updates);
    });
  }

  /**
   * Delete calendar item
   */
  async deleteCalendarItem(account, itemId) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.deleteCalendarItem(account, itemId);
    });
  }

  /**
   * Get user settings and configuration
   */
  async getUserSettings(account) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.getUserSettings(account);
    });
  }

  /**
   * Subscribe to notifications
   */
  async subscribeToNotifications(account, folders = ['inbox']) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.subscribeToNotifications(account, folders);
    });
  }

  /**
   * Get notification events
   */
  async getNotificationEvents(account, subscriptionId) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.getNotificationEvents(account, subscriptionId);
    });
  }

  /**
   * Unsubscribe from notifications
   */
  async unsubscribeFromNotifications(account, subscriptionId) {
    return await this.executeWithRetry(async () => {
      return await this.ewsClient.unsubscribeFromNotifications(account, subscriptionId);
    });
  }

  /**
   * Execute operation with retry logic
   */
  async executeWithRetry(operation, retryCount = 0) {
    try {
      return await operation();
    } catch (error) {
      console.error(`Operation failed (attempt ${retryCount + 1}):`, error);

      // Check if error is retryable
      if (this.isRetryableError(error) && retryCount < this.maxRetries) {
        const delay = this.baseRetryDelay * Math.pow(2, retryCount); // Exponential backoff
        console.log(`Retrying in ${delay}ms...`);
        
        await this.sleep(delay);
        return await this.executeWithRetry(operation, retryCount + 1);
      }

      // If not retryable or max retries reached, throw the error
      throw error;
    }
  }

  /**
   * Check if an error is retryable
   */
  isRetryableError(error) {
    // Network errors
    if (error.message.includes('NetworkError') || 
        error.message.includes('ECONNRESET') ||
        error.message.includes('ETIMEDOUT')) {
      return true;
    }

    // Server errors (5xx)
    if (error.status >= 500 && error.status < 600) {
      return true;
    }

    // Rate limiting
    if (error.status === 429) {
      return true;
    }

    // EWS specific throttling errors
    if (error.message.includes('ErrorServerBusy') ||
        error.message.includes('ErrorTimeoutExpired') ||
        error.message.includes('ErrorConnectionFailed')) {
      return true;
    }

    return false;
  }

  /**
   * Sleep utility for retry delays
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Batch operations for better performance
   */
  async batchOperation(account, operations) {
    try {
      console.log(`Executing batch operation with ${operations.length} operations`);
      
      // Group operations by type for optimal batching
      const grouped = this.groupOperationsByType(operations);
      const results = [];

      for (const [type, ops] of Object.entries(grouped)) {
        const batchResults = await this.executeBatchByType(account, type, ops);
        results.push(...batchResults);
      }

      return {
        success: true,
        results: results
      };
    } catch (error) {
      console.error('Batch operation failed:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Group operations by type for efficient batching
   */
  groupOperationsByType(operations) {
    const grouped = {};
    
    operations.forEach(op => {
      if (!grouped[op.type]) {
        grouped[op.type] = [];
      }
      grouped[op.type].push(op);
    });

    return grouped;
  }

  /**
   * Execute batch operations by type
   */
  async executeBatchByType(account, type, operations) {
    const batchSize = 50; // EWS recommended batch size
    const results = [];

    for (let i = 0; i < operations.length; i += batchSize) {
      const batch = operations.slice(i, i + batchSize);
      
      try {
        let batchResult;
        
        switch (type) {
          case 'getMessage':
            batchResult = await this.ewsClient.getMessagesBatch(account, batch);
            break;
          case 'markMessage':
            batchResult = await this.ewsClient.markMessagesBatch(account, batch);
            break;
          case 'moveMessage':
            batchResult = await this.ewsClient.moveMessagesBatch(account, batch);
            break;
          case 'deleteMessage':
            batchResult = await this.ewsClient.deleteMessagesBatch(account, batch);
            break;
          default:
            // Execute operations individually if no batch method available
            batchResult = await this.executeIndividualOperations(account, batch);
        }

        results.push(...batchResult);
      } catch (error) {
        console.error(`Batch operation failed for type ${type}:`, error);
        // Continue with individual operations as fallback
        const individualResults = await this.executeIndividualOperations(account, batch);
        results.push(...individualResults);
      }
    }

    return results;
  }

  /**
   * Execute operations individually as fallback
   */
  async executeIndividualOperations(account, operations) {
    const results = [];
    
    for (const operation of operations) {
      try {
        let result;
        
        switch (operation.type) {
          case 'getMessage':
            result = await this.getMessage(account, operation.messageId);
            break;
          case 'markMessage':
            result = await this.markMessage(account, operation.messageId, operation.isRead);
            break;
          case 'moveMessage':
            result = await this.moveMessage(account, operation.messageId, operation.targetFolderId);
            break;
          case 'deleteMessage':
            result = await this.deleteMessage(account, operation.messageId);
            break;
          default:
            result = { success: false, error: 'Unknown operation type' };
        }
        
        results.push({ operation, result });
      } catch (error) {
        results.push({ 
          operation, 
          result: { success: false, error: error.message } 
        });
      }
    }

    return results;
  }

  /**
   * Health check for the Exchange connection
   */
  async healthCheck(account) {
    try {
      const start = Date.now();
      
      // Test basic operations
      const folderTest = await this.getFolder(account, 'inbox');
      const latency = Date.now() - start;

      return {
        success: true,
        latency: latency,
        server: account.serverSettings.ewsUrl,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      return {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }
}
