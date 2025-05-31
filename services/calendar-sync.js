/**
 * Calendar Synchronization Service
 * Handles bidirectional synchronization of calendar events between Exchange and Thunderbird
 */

class CalendarSync {
  constructor() {
    this.exchangeClient = null; // Will be injected
    this.syncState = new Map(); // Track sync state per account
    this.lastSyncTimestamp = new Map();
    this.batchSize = 50; // Number of calendar items to sync in one batch
    this.syncWindowDays = 90; // Sync events within 90 days window (past 30 + future 60)
  }

  /**
   * Initialize calendar sync for an account
   */
  async init(account, exchangeClient) {
    this.exchangeClient = exchangeClient;
    
    // Initialize sync state for this account
    this.syncState.set(account.id, {
      inProgress: false,
      lastError: null,
      statistics: {
        totalSynced: 0,
        created: 0,
        updated: 0,
        deleted: 0,
        errors: 0,
        lastSyncDuration: 0
      }
    });

    console.log('Calendar sync initialized for account:', account.email);
  }

  /**
   * Perform full calendar synchronization
   */
  async sync(account) {
    const syncStartTime = Date.now();
    console.log('Starting calendar sync for account:', account.email);

    try {
      // Check if sync is already in progress
      const state = this.syncState.get(account.id);
      if (state && state.inProgress) {
        console.log('Calendar sync already in progress for account:', account.email);
        return { success: false, error: 'Sync already in progress' };
      }

      // Mark sync as in progress
      this.updateSyncState(account.id, { inProgress: true, lastError: null });

      // Get or create Thunderbird calendar
      const calendar = await this.getOrCreateCalendar(account);
      
      if (!calendar) {
        throw new Error('Failed to get or create calendar');
      }

      // Define sync window
      const syncWindow = this.getSyncWindow();

      // Perform bidirectional sync
      const syncResults = await this.performBidirectionalSync(account, calendar, syncWindow);

      // Update statistics
      const syncDuration = Date.now() - syncStartTime;
      this.updateSyncState(account.id, {
        inProgress: false,
        statistics: {
          totalSynced: syncResults.totalSynced,
          created: syncResults.created,
          updated: syncResults.updated,
          deleted: syncResults.deleted,
          errors: syncResults.errors,
          lastSyncDuration: syncDuration
        }
      });

      this.lastSyncTimestamp.set(account.id, new Date().toISOString());

      console.log(`Calendar sync completed for ${account.email}. Total: ${syncResults.totalSynced}, Created: ${syncResults.created}, Updated: ${syncResults.updated}, Deleted: ${syncResults.deleted}, Errors: ${syncResults.errors}`);

      return {
        success: true,
        totalSynced: syncResults.totalSynced,
        created: syncResults.created,
        updated: syncResults.updated,
        deleted: syncResults.deleted,
        errors: syncResults.errors,
        duration: syncDuration
      };

    } catch (error) {
      console.error('Calendar sync failed for account:', account.email, error);
      
      this.updateSyncState(account.id, {
        inProgress: false,
        lastError: error.message
      });

      throw error;
    }
  }

  /**
   * Get or create calendar for Exchange events
   */
  async getOrCreateCalendar(account) {
    try {
      const calendarName = `Exchange - ${account.displayName}`;
      
      // Note: Thunderbird calendar APIs might be limited in WebExtensions
      // This is a simplified approach - real implementation would depend on
      // available calendar APIs or Lightning extension integration
      
      console.log('Getting or creating calendar:', calendarName);
      
      // For now, we'll create a virtual calendar object
      // In a real implementation, this would interact with Thunderbird's calendar system
      return {
        id: `exchange-calendar-${account.id}`,
        name: calendarName,
        type: 'exchange',
        accountId: account.id
      };

    } catch (error) {
      console.error('Failed to get or create calendar:', error);
      throw error;
    }
  }

  /**
   * Get sync window (date range for synchronization)
   */
  getSyncWindow() {
    const now = new Date();
    const startDate = new Date(now);
    startDate.setDate(now.getDate() - 30); // 30 days in the past
    
    const endDate = new Date(now);
    endDate.setDate(now.getDate() + 60); // 60 days in the future

    return { startDate, endDate };
  }

  /**
   * Perform bidirectional synchronization
   */
  async performBidirectionalSync(account, calendar, syncWindow) {
    const results = {
      totalSynced: 0,
      created: 0,
      updated: 0,
      deleted: 0,
      errors: 0
    };

    try {
      // Step 1: Sync from Exchange to Thunderbird
      const exchangeToThunderbirdResults = await this.syncExchangeToThunderbird(
        account, 
        calendar, 
        syncWindow
      );
      
      results.totalSynced += exchangeToThunderbirdResults.totalSynced;
      results.created += exchangeToThunderbirdResults.created;
      results.updated += exchangeToThunderbirdResults.updated;
      results.deleted += exchangeToThunderbirdResults.deleted;
      results.errors += exchangeToThunderbirdResults.errors;

      // Step 2: Sync from Thunderbird to Exchange
      const thunderbirdToExchangeResults = await this.syncThunderbirdToExchange(
        account, 
        calendar, 
        syncWindow
      );
      
      results.totalSynced += thunderbirdToExchangeResults.totalSynced;
      results.created += thunderbirdToExchangeResults.created;
      results.updated += thunderbirdToExchangeResults.updated;
      results.deleted += thunderbirdToExchangeResults.deleted;
      results.errors += thunderbirdToExchangeResults.errors;

      return results;

    } catch (error) {
      console.error('Bidirectional sync failed:', error);
      throw error;
    }
  }

  /**
   * Sync calendar items from Exchange to Thunderbird
   */
  async syncExchangeToThunderbird(account, calendar, syncWindow) {
    console.log('Syncing calendar items from Exchange to Thunderbird');

    const results = {
      totalSynced: 0,
      created: 0,
      updated: 0,
      deleted: 0,
      errors: 0
    };

    try {
      // Get existing calendar items from Thunderbird
      const thunderbirdItems = await this.getThunderbirdCalendarItems(calendar, syncWindow);
      const thunderbirdItemMap = new Map();
      
      thunderbirdItems.forEach(item => {
        const key = this.generateCalendarItemKey(item);
        thunderbirdItemMap.set(key, item);
      });

      // Get calendar items from Exchange
      const exchangeItems = await this.exchangeClient.getCalendarItems(
        account, 
        syncWindow.startDate, 
        syncWindow.endDate
      );

      if (!exchangeItems.success || !exchangeItems.items) {
        throw new Error('Failed to get Exchange calendar items');
      }

      // Process each Exchange calendar item
      for (const exchangeItem of exchangeItems.items) {
        try {
          const itemKey = this.generateCalendarItemKeyFromExchange(exchangeItem);
          const thunderbirdItem = thunderbirdItemMap.get(itemKey);

          if (thunderbirdItem) {
            // Update existing item if needed
            if (this.needsUpdate(thunderbirdItem, exchangeItem)) {
              const updatedItem = this.convertExchangeCalendarItem(exchangeItem);
              await this.updateThunderbirdCalendarItem(calendar, thunderbirdItem.id, updatedItem);
              results.updated++;
              results.totalSynced++;
            }
          } else {
            // Create new item
            const newItem = this.convertExchangeCalendarItem(exchangeItem);
            await this.createThunderbirdCalendarItem(calendar, newItem);
            results.created++;
            results.totalSynced++;
          }

        } catch (error) {
          console.error('Failed to sync individual calendar item:', error);
          results.errors++;
        }
      }

      console.log(`Exchange to Thunderbird calendar sync completed. Created: ${results.created}, Updated: ${results.updated}, Errors: ${results.errors}`);
      return results;

    } catch (error) {
      console.error('Failed to sync from Exchange to Thunderbird:', error);
      throw error;
    }
  }

  /**
   * Sync calendar items from Thunderbird to Exchange
   */
  async syncThunderbirdToExchange(account, calendar, syncWindow) {
    console.log('Syncing calendar items from Thunderbird to Exchange');

    const results = {
      totalSynced: 0,
      created: 0,
      updated: 0,
      deleted: 0,
      errors: 0
    };

    try {
      // Get calendar items from Thunderbird
      const thunderbirdItems = await this.getThunderbirdCalendarItems(calendar, syncWindow);

      // Get existing items from Exchange for comparison
      const exchangeItems = await this.exchangeClient.getCalendarItems(
        account, 
        syncWindow.startDate, 
        syncWindow.endDate
      );

      const exchangeItemMap = new Map();
      if (exchangeItems.success && exchangeItems.items) {
        exchangeItems.items.forEach(item => {
          const key = this.generateCalendarItemKeyFromExchange(item);
          exchangeItemMap.set(key, item);
        });
      }

      // Process each Thunderbird calendar item
      for (const thunderbirdItem of thunderbirdItems) {
        try {
          const itemKey = this.generateCalendarItemKey(thunderbirdItem);
          const exchangeItem = exchangeItemMap.get(itemKey);

          if (exchangeItem) {
            // Update existing item in Exchange if needed
            if (this.thunderbirdItemNeedsUpdate(thunderbirdItem, exchangeItem)) {
              const updates = this.convertThunderbirdCalendarItem(thunderbirdItem);
              await this.exchangeClient.updateCalendarItem(account, exchangeItem.id, updates);
              results.updated++;
              results.totalSynced++;
            }
          } else {
            // Create new item in Exchange
            const newItem = this.convertThunderbirdCalendarItem(thunderbirdItem);
            await this.exchangeClient.createCalendarItem(account, newItem);
            results.created++;
            results.totalSynced++;
          }

        } catch (error) {
          console.error('Failed to sync Thunderbird calendar item to Exchange:', error);
          results.errors++;
        }
      }

      console.log(`Thunderbird to Exchange calendar sync completed. Created: ${results.created}, Updated: ${results.updated}, Errors: ${results.errors}`);
      return results;

    } catch (error) {
      console.error('Failed to sync from Thunderbird to Exchange:', error);
      throw error;
    }
  }

  /**
   * Get calendar items from Thunderbird
   */
  async getThunderbirdCalendarItems(calendar, syncWindow) {
    try {
      // Note: This is a simplified implementation
      // Real implementation would depend on available Thunderbird calendar APIs
      console.log('Getting Thunderbird calendar items for calendar:', calendar.name);
      
      // For now, return empty array as Thunderbird calendar APIs in WebExtensions are limited
      // In a real implementation, this would query the Lightning calendar database
      return [];

    } catch (error) {
      console.error('Failed to get Thunderbird calendar items:', error);
      return [];
    }
  }

  /**
   * Create calendar item in Thunderbird
   */
  async createThunderbirdCalendarItem(calendar, item) {
    try {
      console.log('Creating Thunderbird calendar item:', item.subject);
      
      // Note: This is a simplified implementation
      // Real implementation would use Thunderbird calendar APIs
      // For now, we'll log the operation
      
      return { success: true, id: `tb-item-${Date.now()}` };

    } catch (error) {
      console.error('Failed to create Thunderbird calendar item:', error);
      throw error;
    }
  }

  /**
   * Update calendar item in Thunderbird
   */
  async updateThunderbirdCalendarItem(calendar, itemId, updates) {
    try {
      console.log('Updating Thunderbird calendar item:', itemId);
      
      // Note: This is a simplified implementation
      // Real implementation would use Thunderbird calendar APIs
      
      return { success: true };

    } catch (error) {
      console.error('Failed to update Thunderbird calendar item:', error);
      throw error;
    }
  }

  /**
   * Generate calendar item key for deduplication
   */
  generateCalendarItemKey(thunderbirdItem) {
    // Use subject + start time as key
    const subject = thunderbirdItem.subject || thunderbirdItem.title || 'no-subject';
    const startTime = thunderbirdItem.startDate ? new Date(thunderbirdItem.startDate).getTime() : 0;
    
    return `${subject}-${startTime}`;
  }

  /**
   * Generate calendar item key from Exchange item
   */
  generateCalendarItemKeyFromExchange(exchangeItem) {
    // Use subject + start time as key
    const subject = exchangeItem.subject || 'no-subject';
    const startTime = exchangeItem.start ? new Date(exchangeItem.start).getTime() : 0;
    
    return `${subject}-${startTime}`;
  }

  /**
   * Convert Exchange calendar item to Thunderbird format
   */
  convertExchangeCalendarItem(exchangeItem) {
    const item = {
      subject: exchangeItem.subject || '',
      title: exchangeItem.subject || '', // Alternative property name
      body: exchangeItem.body || '',
      description: exchangeItem.body || '', // Alternative property name
      startDate: exchangeItem.start ? new Date(exchangeItem.start) : new Date(),
      endDate: exchangeItem.end ? new Date(exchangeItem.end) : new Date(),
      location: exchangeItem.location || '',
      
      // Status and priority
      status: this.convertFreeBusyStatus(exchangeItem.freeBusyStatus),
      priority: 'normal', // Default priority
      
      // Organizer and attendees
      organizer: this.convertEmailAddress(exchangeItem.organizer),
      attendees: this.convertEmailAddresses(exchangeItem.attendees),
      
      // Recurrence (simplified)
      isRecurring: false, // Would need more complex logic for recurrence
      
      // Metadata
      exchangeId: exchangeItem.id,
      lastModified: new Date(),
      created: new Date()
    };

    return item;
  }

  /**
   * Convert Thunderbird calendar item to Exchange format
   */
  convertThunderbirdCalendarItem(thunderbirdItem) {
    const item = {
      subject: thunderbirdItem.subject || thunderbirdItem.title || '',
      body: thunderbirdItem.body || thunderbirdItem.description || '',
      start: thunderbirdItem.startDate || new Date(),
      end: thunderbirdItem.endDate || new Date(),
      location: thunderbirdItem.location || '',
      
      // Status
      freeBusyStatus: this.convertStatusToFreeBusy(thunderbirdItem.status),
      
      // Body type
      bodyType: 'Text' // Default to text, could be HTML if rich content is detected
    };

    return item;
  }

  /**
   * Convert Exchange FreeBusy status to Thunderbird status
   */
  convertFreeBusyStatus(freeBusyStatus) {
    const statusMap = {
      'Free': 'available',
      'Tentative': 'tentative',
      'Busy': 'busy',
      'OOF': 'out-of-office', // Out of Office
      'WorkingElsewhere': 'working-elsewhere'
    };

    return statusMap[freeBusyStatus] || 'busy';
  }

  /**
   * Convert Thunderbird status to Exchange FreeBusy status
   */
  convertStatusToFreeBusy(status) {
    const statusMap = {
      'available': 'Free',
      'tentative': 'Tentative',
      'busy': 'Busy',
      'out-of-office': 'OOF',
      'working-elsewhere': 'WorkingElsewhere'
    };

    return statusMap[status] || 'Busy';
  }

  /**
   * Convert Exchange email address to Thunderbird format
   */
  convertEmailAddress(exchangeAddress) {
    if (!exchangeAddress) return null;
    
    return {
      name: exchangeAddress.name || '',
      email: exchangeAddress.email || ''
    };
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
   * Check if Thunderbird item needs update based on Exchange item
   */
  needsUpdate(thunderbirdItem, exchangeItem) {
    // Compare key fields
    const fieldsToCompare = [
      { tb: 'subject', ex: 'subject' },
      { tb: 'body', ex: 'body' },
      { tb: 'location', ex: 'location' }
    ];

    for (const field of fieldsToCompare) {
      const tbValue = thunderbirdItem[field.tb] || '';
      const exValue = exchangeItem[field.ex] || '';
      
      if (tbValue !== exValue) {
        return true;
      }
    }

    // Compare dates
    const tbStart = thunderbirdItem.startDate ? new Date(thunderbirdItem.startDate).getTime() : 0;
    const exStart = exchangeItem.start ? new Date(exchangeItem.start).getTime() : 0;
    
    const tbEnd = thunderbirdItem.endDate ? new Date(thunderbirdItem.endDate).getTime() : 0;
    const exEnd = exchangeItem.end ? new Date(exchangeItem.end).getTime() : 0;

    if (tbStart !== exStart || tbEnd !== exEnd) {
      return true;
    }

    return false;
  }

  /**
   * Check if Thunderbird item needs to be updated in Exchange
   */
  thunderbirdItemNeedsUpdate(thunderbirdItem, exchangeItem) {
    // Similar to needsUpdate but checking if Thunderbird version is newer
    // For simplicity, we'll use the same logic
    return this.needsUpdate(thunderbirdItem, exchangeItem);
  }

  /**
   * Perform incremental calendar sync
   */
  async incrementalSync(account) {
    console.log('Starting incremental calendar sync for account:', account.email);

    try {
      const lastSync = this.lastSyncTimestamp.get(account.id);
      if (!lastSync) {
        // If no previous sync, perform full sync
        return await this.sync(account);
      }

      // For incremental sync, we'll use a shorter time window
      const incrementalWindow = this.getIncrementalSyncWindow(new Date(lastSync));
      const calendar = await this.getOrCreateCalendar(account);
      
      if (!calendar) {
        throw new Error('Failed to get calendar');
      }

      // Perform sync for the incremental window
      const results = await this.performBidirectionalSync(account, calendar, incrementalWindow);

      this.lastSyncTimestamp.set(account.id, new Date().toISOString());

      console.log(`Incremental calendar sync completed. Synced: ${results.totalSynced}`);
      return { success: true, totalSynced: results.totalSynced };

    } catch (error) {
      console.error('Incremental calendar sync failed:', error);
      throw error;
    }
  }

  /**
   * Get incremental sync window (focused on recent changes)
   */
  getIncrementalSyncWindow(lastSyncDate) {
    const now = new Date();
    
    // Start from last sync date minus some buffer for safety
    const startDate = new Date(lastSyncDate);
    startDate.setHours(startDate.getHours() - 1); // 1 hour buffer
    
    // End date is future events (next 30 days)
    const endDate = new Date(now);
    endDate.setDate(now.getDate() + 30);

    return { startDate, endDate };
  }

  /**
   * Handle calendar item deletion
   */
  async deleteCalendarItem(account, itemId, isThunderbirdItem = false) {
    try {
      if (isThunderbirdItem) {
        // Delete from Exchange
        const exchangeItemId = await this.findExchangeItemId(account, itemId);
        if (exchangeItemId) {
          await this.exchangeClient.deleteCalendarItem(account, exchangeItemId);
        }
      } else {
        // Delete from Thunderbird
        const thunderbirdItemId = await this.findThunderbirdItemId(account, itemId);
        if (thunderbirdItemId) {
          await this.deleteThunderbirdCalendarItem(thunderbirdItemId);
        }
      }

      return { success: true };

    } catch (error) {
      console.error('Failed to delete calendar item:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Delete calendar item from Thunderbird
   */
  async deleteThunderbirdCalendarItem(itemId) {
    try {
      console.log('Deleting Thunderbird calendar item:', itemId);
      
      // Note: This is a simplified implementation
      // Real implementation would use Thunderbird calendar APIs
      
      return { success: true };

    } catch (error) {
      console.error('Failed to delete Thunderbird calendar item:', error);
      throw error;
    }
  }

  /**
   * Find Exchange item ID for Thunderbird item
   */
  async findExchangeItemId(account, thunderbirdItemId) {
    // This would require maintaining a mapping between Thunderbird and Exchange item IDs
    // For now, return null as this is complex to implement without proper change tracking
    return null;
  }

  /**
   * Find Thunderbird item ID for Exchange item
   */
  async findThunderbirdItemId(account, exchangeItemId) {
    // This would require maintaining a mapping between Exchange and Thunderbird item IDs
    // For now, return null as this is complex to implement without proper change tracking
    return null;
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
