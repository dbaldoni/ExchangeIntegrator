/**
 * Contact Synchronization Service
 * Handles bidirectional synchronization of contacts between Exchange and Thunderbird
 */

class ContactSync {
  constructor() {
    this.exchangeClient = null; // Will be injected
    this.syncState = new Map(); // Track sync state per account
    this.lastSyncTimestamp = new Map();
    this.batchSize = 100; // Number of contacts to sync in one batch
  }

  /**
   * Initialize contact sync for an account
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
        errors: 0,
        lastSyncDuration: 0
      }
    });

    console.log('Contact sync initialized for account:', account.email);
  }

  /**
   * Perform full contact synchronization
   */
  async sync(account) {
    const syncStartTime = Date.now();
    console.log('Starting contact sync for account:', account.email);

    try {
      // Check if sync is already in progress
      const state = this.syncState.get(account.id);
      if (state && state.inProgress) {
        console.log('Contact sync already in progress for account:', account.email);
        return { success: false, error: 'Sync already in progress' };
      }

      // Mark sync as in progress
      this.updateSyncState(account.id, { inProgress: true, lastError: null });

      // Get or create Thunderbird address book
      const addressBook = await this.getOrCreateAddressBook(account);
      
      if (!addressBook) {
        throw new Error('Failed to get or create address book');
      }

      // Perform bidirectional sync
      const syncResults = await this.performBidirectionalSync(account, addressBook);

      // Update statistics
      const syncDuration = Date.now() - syncStartTime;
      this.updateSyncState(account.id, {
        inProgress: false,
        statistics: {
          totalSynced: syncResults.totalSynced,
          created: syncResults.created,
          updated: syncResults.updated,
          errors: syncResults.errors,
          lastSyncDuration: syncDuration
        }
      });

      this.lastSyncTimestamp.set(account.id, new Date().toISOString());

      console.log(`Contact sync completed for ${account.email}. Total: ${syncResults.totalSynced}, Created: ${syncResults.created}, Updated: ${syncResults.updated}, Errors: ${syncResults.errors}`);

      return {
        success: true,
        totalSynced: syncResults.totalSynced,
        created: syncResults.created,
        updated: syncResults.updated,
        errors: syncResults.errors,
        duration: syncDuration
      };

    } catch (error) {
      console.error('Contact sync failed for account:', account.email, error);
      
      this.updateSyncState(account.id, {
        inProgress: false,
        lastError: error.message
      });

      throw error;
    }
  }

  /**
   * Get or create address book for Exchange contacts
   */
  async getOrCreateAddressBook(account) {
    try {
      const addressBookName = `Exchange - ${account.displayName}`;
      
      // Get all address books
      const addressBooks = await browser.addressBooks.list();
      
      // Look for existing Exchange address book
      let exchangeAddressBook = addressBooks.find(ab => ab.name === addressBookName);
      
      if (!exchangeAddressBook) {
        // Create new address book
        exchangeAddressBook = await browser.addressBooks.create({
          name: addressBookName,
          type: 'jsaddressbook'
        });
        
        console.log('Created new address book:', addressBookName);
      }

      return exchangeAddressBook;

    } catch (error) {
      console.error('Failed to get or create address book:', error);
      throw error;
    }
  }

  /**
   * Perform bidirectional synchronization
   */
  async performBidirectionalSync(account, addressBook) {
    const results = {
      totalSynced: 0,
      created: 0,
      updated: 0,
      errors: 0
    };

    try {
      // Step 1: Sync from Exchange to Thunderbird
      const exchangeToThunderbirdResults = await this.syncExchangeToThunderbird(account, addressBook);
      
      results.totalSynced += exchangeToThunderbirdResults.totalSynced;
      results.created += exchangeToThunderbirdResults.created;
      results.updated += exchangeToThunderbirdResults.updated;
      results.errors += exchangeToThunderbirdResults.errors;

      // Step 2: Sync from Thunderbird to Exchange
      const thunderbirdToExchangeResults = await this.syncThunderbirdToExchange(account, addressBook);
      
      results.totalSynced += thunderbirdToExchangeResults.totalSynced;
      results.created += thunderbirdToExchangeResults.created;
      results.updated += thunderbirdToExchangeResults.updated;
      results.errors += thunderbirdToExchangeResults.errors;

      return results;

    } catch (error) {
      console.error('Bidirectional sync failed:', error);
      throw error;
    }
  }

  /**
   * Sync contacts from Exchange to Thunderbird
   */
  async syncExchangeToThunderbird(account, addressBook) {
    console.log('Syncing contacts from Exchange to Thunderbird');

    const results = {
      totalSynced: 0,
      created: 0,
      updated: 0,
      errors: 0
    };

    try {
      // Get existing contacts from Thunderbird
      const thunderbirdContacts = await browser.contacts.list(addressBook.id);
      const thunderbirdContactMap = new Map();
      
      thunderbirdContacts.forEach(contact => {
        const key = this.generateContactKey(contact);
        thunderbirdContactMap.set(key, contact);
      });

      // Get contacts from Exchange
      let offset = 0;
      let hasMore = true;

      while (hasMore) {
        try {
          const exchangeContacts = await this.exchangeClient.getContacts(account, {
            maxItems: this.batchSize,
            offset: offset
          });

          if (!exchangeContacts.success || !exchangeContacts.contacts) {
            break;
          }

          // Process each contact
          for (const exchangeContact of exchangeContacts.contacts) {
            try {
              const contactKey = this.generateContactKeyFromExchange(exchangeContact);
              const thunderbirdContact = thunderbirdContactMap.get(contactKey);

              if (thunderbirdContact) {
                // Update existing contact if needed
                if (this.needsUpdate(thunderbirdContact, exchangeContact)) {
                  const updatedContact = this.convertExchangeContact(exchangeContact);
                  await browser.contacts.update(thunderbirdContact.id, updatedContact);
                  results.updated++;
                  results.totalSynced++;
                }
              } else {
                // Create new contact
                const newContact = this.convertExchangeContact(exchangeContact);
                await browser.contacts.create(addressBook.id, newContact);
                results.created++;
                results.totalSynced++;
              }

            } catch (error) {
              console.error('Failed to sync individual contact:', error);
              results.errors++;
            }
          }

          // Check if there are more contacts
          hasMore = exchangeContacts.contacts.length === this.batchSize;
          offset += this.batchSize;

          // Add small delay to avoid overwhelming the server
          if (hasMore) {
            await this.sleep(100);
          }

        } catch (error) {
          console.error('Failed to get contacts batch:', error);
          results.errors++;
          break;
        }
      }

      console.log(`Exchange to Thunderbird sync completed. Created: ${results.created}, Updated: ${results.updated}, Errors: ${results.errors}`);
      return results;

    } catch (error) {
      console.error('Failed to sync from Exchange to Thunderbird:', error);
      throw error;
    }
  }

  /**
   * Sync contacts from Thunderbird to Exchange
   */
  async syncThunderbirdToExchange(account, addressBook) {
    console.log('Syncing contacts from Thunderbird to Exchange');

    const results = {
      totalSynced: 0,
      created: 0,
      updated: 0,
      errors: 0
    };

    try {
      // Get contacts from Thunderbird
      const thunderbirdContacts = await browser.contacts.list(addressBook.id);

      // Get existing contacts from Exchange for comparison
      const exchangeContacts = await this.getAllExchangeContacts(account);
      const exchangeContactMap = new Map();
      
      exchangeContacts.forEach(contact => {
        const key = this.generateContactKeyFromExchange(contact);
        exchangeContactMap.set(key, contact);
      });

      // Process each Thunderbird contact
      for (const thunderbirdContact of thunderbirdContacts) {
        try {
          const contactKey = this.generateContactKey(thunderbirdContact);
          const exchangeContact = exchangeContactMap.get(contactKey);

          if (exchangeContact) {
            // Update existing contact in Exchange if needed
            if (this.thunderbirdContactNeedsUpdate(thunderbirdContact, exchangeContact)) {
              const updates = this.convertThunderbirdContact(thunderbirdContact);
              await this.exchangeClient.updateContact(account, exchangeContact.id, updates);
              results.updated++;
              results.totalSynced++;
            }
          } else {
            // Create new contact in Exchange
            const newContact = this.convertThunderbirdContact(thunderbirdContact);
            await this.exchangeClient.createContact(account, newContact);
            results.created++;
            results.totalSynced++;
          }

        } catch (error) {
          console.error('Failed to sync Thunderbird contact to Exchange:', error);
          results.errors++;
        }
      }

      console.log(`Thunderbird to Exchange sync completed. Created: ${results.created}, Updated: ${results.updated}, Errors: ${results.errors}`);
      return results;

    } catch (error) {
      console.error('Failed to sync from Thunderbird to Exchange:', error);
      throw error;
    }
  }

  /**
   * Get all contacts from Exchange
   */
  async getAllExchangeContacts(account) {
    const allContacts = [];
    let offset = 0;
    let hasMore = true;

    while (hasMore) {
      try {
        const exchangeContacts = await this.exchangeClient.getContacts(account, {
          maxItems: this.batchSize,
          offset: offset
        });

        if (!exchangeContacts.success || !exchangeContacts.contacts) {
          break;
        }

        allContacts.push(...exchangeContacts.contacts);
        hasMore = exchangeContacts.contacts.length === this.batchSize;
        offset += this.batchSize;

      } catch (error) {
        console.error('Failed to get Exchange contacts batch:', error);
        break;
      }
    }

    return allContacts;
  }

  /**
   * Generate contact key for deduplication
   */
  generateContactKey(thunderbirdContact) {
    // Use email as primary key, fallback to name
    if (thunderbirdContact.properties && thunderbirdContact.properties.PrimaryEmail) {
      return thunderbirdContact.properties.PrimaryEmail.toLowerCase();
    }
    
    const firstName = thunderbirdContact.properties?.FirstName || '';
    const lastName = thunderbirdContact.properties?.LastName || '';
    const displayName = thunderbirdContact.properties?.DisplayName || '';
    
    return `${firstName} ${lastName} ${displayName}`.trim().toLowerCase();
  }

  /**
   * Generate contact key from Exchange contact
   */
  generateContactKeyFromExchange(exchangeContact) {
    // Use email as primary key, fallback to name
    if (exchangeContact.email) {
      return exchangeContact.email.toLowerCase();
    }
    
    const displayName = exchangeContact.displayName || '';
    const firstName = exchangeContact.firstName || '';
    const lastName = exchangeContact.lastName || '';
    
    return `${firstName} ${lastName} ${displayName}`.trim().toLowerCase();
  }

  /**
   * Convert Exchange contact to Thunderbird format
   */
  convertExchangeContact(exchangeContact) {
    const contact = {
      properties: {}
    };

    // Map Exchange properties to Thunderbird properties
    if (exchangeContact.displayName) {
      contact.properties.DisplayName = exchangeContact.displayName;
    }

    if (exchangeContact.firstName) {
      contact.properties.FirstName = exchangeContact.firstName;
    }

    if (exchangeContact.lastName) {
      contact.properties.LastName = exchangeContact.lastName;
    }

    if (exchangeContact.email) {
      contact.properties.PrimaryEmail = exchangeContact.email;
    }

    if (exchangeContact.phone) {
      contact.properties.WorkPhone = exchangeContact.phone;
    }

    if (exchangeContact.company) {
      contact.properties.Company = exchangeContact.company;
    }

    // Additional properties that might be available
    if (exchangeContact.mobilePhone) {
      contact.properties.CellularNumber = exchangeContact.mobilePhone;
    }

    if (exchangeContact.homePhone) {
      contact.properties.HomePhone = exchangeContact.homePhone;
    }

    if (exchangeContact.workAddress) {
      contact.properties.WorkAddress = exchangeContact.workAddress;
    }

    if (exchangeContact.homeAddress) {
      contact.properties.HomeAddress = exchangeContact.homeAddress;
    }

    if (exchangeContact.notes) {
      contact.properties.Notes = exchangeContact.notes;
    }

    return contact;
  }

  /**
   * Convert Thunderbird contact to Exchange format
   */
  convertThunderbirdContact(thunderbirdContact) {
    const props = thunderbirdContact.properties || {};
    
    const contact = {};

    // Map Thunderbird properties to Exchange properties
    if (props.DisplayName) {
      contact.displayName = props.DisplayName;
    }

    if (props.FirstName) {
      contact.firstName = props.FirstName;
    }

    if (props.LastName) {
      contact.lastName = props.LastName;
    }

    if (props.PrimaryEmail) {
      contact.email = props.PrimaryEmail;
    }

    if (props.WorkPhone) {
      contact.phone = props.WorkPhone;
    }

    if (props.Company) {
      contact.company = props.Company;
    }

    // Additional properties
    if (props.CellularNumber) {
      contact.mobilePhone = props.CellularNumber;
    }

    if (props.HomePhone) {
      contact.homePhone = props.HomePhone;
    }

    if (props.WorkAddress) {
      contact.workAddress = props.WorkAddress;
    }

    if (props.HomeAddress) {
      contact.homeAddress = props.HomeAddress;
    }

    if (props.Notes) {
      contact.notes = props.Notes;
    }

    return contact;
  }

  /**
   * Check if Thunderbird contact needs update based on Exchange contact
   */
  needsUpdate(thunderbirdContact, exchangeContact) {
    const props = thunderbirdContact.properties || {};
    
    // Compare key fields
    const fieldsToCompare = [
      { tb: 'DisplayName', ex: 'displayName' },
      { tb: 'FirstName', ex: 'firstName' },
      { tb: 'LastName', ex: 'lastName' },
      { tb: 'PrimaryEmail', ex: 'email' },
      { tb: 'WorkPhone', ex: 'phone' },
      { tb: 'Company', ex: 'company' }
    ];

    for (const field of fieldsToCompare) {
      const tbValue = props[field.tb] || '';
      const exValue = exchangeContact[field.ex] || '';
      
      if (tbValue !== exValue) {
        return true;
      }
    }

    return false;
  }

  /**
   * Check if Thunderbird contact needs to be updated in Exchange
   */
  thunderbirdContactNeedsUpdate(thunderbirdContact, exchangeContact) {
    const props = thunderbirdContact.properties || {};
    
    // Compare key fields (opposite direction)
    const fieldsToCompare = [
      { tb: 'DisplayName', ex: 'displayName' },
      { tb: 'FirstName', ex: 'firstName' },
      { tb: 'LastName', ex: 'lastName' },
      { tb: 'PrimaryEmail', ex: 'email' },
      { tb: 'WorkPhone', ex: 'phone' },
      { tb: 'Company', ex: 'company' }
    ];

    // Check if Thunderbird contact was modified more recently
    // For now, we'll assume Thunderbird takes precedence if there are differences
    for (const field of fieldsToCompare) {
      const tbValue = props[field.tb] || '';
      const exValue = exchangeContact[field.ex] || '';
      
      if (tbValue !== exValue && tbValue.trim() !== '') {
        return true;
      }
    }

    return false;
  }

  /**
   * Perform incremental contact sync
   */
  async incrementalSync(account) {
    console.log('Starting incremental contact sync for account:', account.email);

    try {
      const lastSync = this.lastSyncTimestamp.get(account.id);
      if (!lastSync) {
        // If no previous sync, perform full sync
        return await this.sync(account);
      }

      // For contacts, we'll perform a lightweight check for changes
      // since most contact systems don't provide reliable change tracking
      const addressBook = await this.getOrCreateAddressBook(account);
      
      if (!addressBook) {
        throw new Error('Failed to get address book');
      }

      // Perform a simplified sync focusing on new contacts
      const results = await this.syncNewContacts(account, addressBook, new Date(lastSync));

      this.lastSyncTimestamp.set(account.id, new Date().toISOString());

      console.log(`Incremental contact sync completed. Synced: ${results.totalSynced}`);
      return { success: true, totalSynced: results.totalSynced };

    } catch (error) {
      console.error('Incremental contact sync failed:', error);
      throw error;
    }
  }

  /**
   * Sync only new contacts since last sync date
   */
  async syncNewContacts(account, addressBook, sinceDate) {
    // This is a simplified implementation
    // In a real scenario, you'd need to track contact modification dates
    console.log('Syncing new contacts since:', sinceDate);
    
    return {
      totalSynced: 0,
      created: 0,
      updated: 0,
      errors: 0
    };
  }

  /**
   * Delete contact synchronization
   */
  async deleteContact(account, contactId, isThunderbirdContact = false) {
    try {
      if (isThunderbirdContact) {
        // Delete from Exchange
        const exchangeContactId = await this.findExchangeContactId(account, contactId);
        if (exchangeContactId) {
          await this.exchangeClient.deleteContact(account, exchangeContactId);
        }
      } else {
        // Delete from Thunderbird
        const thunderbirdContactId = await this.findThunderbirdContactId(account, contactId);
        if (thunderbirdContactId) {
          await browser.contacts.delete(thunderbirdContactId);
        }
      }

      return { success: true };

    } catch (error) {
      console.error('Failed to delete contact:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Find Exchange contact ID for Thunderbird contact
   */
  async findExchangeContactId(account, thunderbirdContactId) {
    // This would require maintaining a mapping between Thunderbird and Exchange contact IDs
    // For now, return null as this is complex to implement without proper change tracking
    return null;
  }

  /**
   * Find Thunderbird contact ID for Exchange contact
   */
  async findThunderbirdContactId(account, exchangeContactId) {
    // This would require maintaining a mapping between Exchange and Thunderbird contact IDs
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
