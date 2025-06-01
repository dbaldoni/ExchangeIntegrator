/**
 * Account setup page script for Exchange/Office365 integration
 * Handles the user interface for adding and managing Exchange accounts
 */

class AccountSetupUI {
  constructor() {
    this.currentStep = 1;
    this.accountData = {};
    this.accounts = [];
    
    this.init();
  }

  async init() {
    console.log('Initializing account setup UI');
    
    // Load existing accounts
    await this.loadAccounts();
    
    // Setup event listeners
    this.setupEventListeners();
    
    // Render initial state
    this.render();
    
    console.log('Account setup UI initialized');
  }

  async loadAccounts() {
    try {
      const response = await browser.runtime.sendMessage({ action: 'getAccounts' });
      this.accounts = response || [];
      console.log('Loaded accounts:', this.accounts.length);
    } catch (error) {
      console.error('Failed to load accounts:', error);
      this.showError('Failed to load existing accounts');
    }
  }

  setupEventListeners() {
    // Setup form
    const setupForm = document.getElementById('setup-form');
    if (setupForm) {
      setupForm.addEventListener('submit', (e) => this.handleSetupSubmit(e));
    }

    // Test connection button
    const testButton = document.getElementById('test-connection');
    if (testButton) {
      testButton.addEventListener('click', () => this.testConnection());
    }

    // Navigation buttons
    const nextButton = document.getElementById('next-step');
    const prevButton = document.getElementById('prev-step');
    const finishButton = document.getElementById('finish-setup');

    if (nextButton) {
      nextButton.addEventListener('click', () => this.nextStep());
    }
    if (prevButton) {
      prevButton.addEventListener('click', () => this.prevStep());
    }
    if (finishButton) {
      finishButton.addEventListener('click', () => this.finishSetup());
    }

    // Account management
    document.addEventListener('click', (e) => {
      if (e.target.classList.contains('sync-account')) {
        const accountId = e.target.dataset.accountId;
        this.syncAccount(accountId);
      }
      if (e.target.classList.contains('remove-account')) {
        const accountId = e.target.dataset.accountId;
        this.removeAccount(accountId);
      }
      if (e.target.classList.contains('add-new-account')) {
        this.showSetupForm();
      }
    });

    // Form validation
    const emailInput = document.getElementById('email');
    const passwordInput = document.getElementById('password');

    if (emailInput) {
      emailInput.addEventListener('input', () => this.validateForm());
    }
    if (passwordInput) {
      passwordInput.addEventListener('input', () => this.validateForm());
    }
  }

  render() {
    if (this.accounts.length === 0) {
      this.showSetupForm();
    } else {
      this.showAccountList();
    }
  }

  showSetupForm() {
    const container = document.getElementById('main-container');
    container.innerHTML = `
      <div class="setup-container">
        <h1>Syncbird - Add Exchange/Office365 Account</h1>
        
        <div class="step-indicator">
          <div class="step ${this.currentStep >= 1 ? 'active' : ''} ${this.currentStep > 1 ? 'completed' : ''}">
            <span class="step-number">1</span>
            <span class="step-label">Account Details</span>
          </div>
          <div class="step ${this.currentStep >= 2 ? 'active' : ''} ${this.currentStep > 2 ? 'completed' : ''}">
            <span class="step-number">2</span>
            <span class="step-label">Server Discovery</span>
          </div>
          <div class="step ${this.currentStep >= 3 ? 'active' : ''} ${this.currentStep > 3 ? 'completed' : ''}">
            <span class="step-number">3</span>
            <span class="step-label">Sync Settings</span>
          </div>
        </div>

        <form id="setup-form" class="setup-form">
          ${this.renderStep()}
        </form>

        <div class="form-actions">
          ${this.renderActions()}
        </div>

        <div id="status-message" class="status-message"></div>
      </div>
    `;

    this.validateForm();
  }

  showAccountList() {
    const container = document.getElementById('main-container');
    container.innerHTML = `
      <div class="account-list-container">
        <h1>Syncbird - Exchange/Office365 Accounts</h1>
        
        <div class="account-list">
          ${this.accounts.map(account => this.renderAccountCard(account)).join('')}
        </div>
        
        <button class="btn btn-primary add-new-account">
          <i class="icon-plus"></i>
          Add New Account
        </button>
      </div>
    `;
  }

  renderStep() {
    switch (this.currentStep) {
      case 1:
        return `
          <div class="form-group">
            <label for="email">Email Address</label>
            <input type="email" id="email" name="email" required 
                   value="${this.accountData.email || ''}"
                   placeholder="user@company.com">
          </div>
          
          <div class="form-group">
            <label for="password">Password</label>
            <input type="password" id="password" name="password" required
                   value="${this.accountData.password || ''}"
                   placeholder="Enter your password">
          </div>
          
          <div class="form-group">
            <label for="display-name">Display Name (Optional)</label>
            <input type="text" id="display-name" name="displayName"
                   value="${this.accountData.displayName || ''}"
                   placeholder="John Doe">
          </div>
          
          <div class="form-group">
            <label>
              <input type="checkbox" id="use-autodiscovery" name="useAutodiscovery" checked>
              Use server autodiscovery (recommended)
            </label>
          </div>

          <div class="form-group">
            <p><strong>Note:</strong> For Office365 accounts, OAuth2 configuration may be required.</p>
            <a href="oauth-setup.html" target="_blank" class="btn btn-outline">Configure OAuth2</a>
          </div>
        `;

      case 2:
        return `
          <div class="discovery-status">
            <h3>Server Discovery</h3>
            <div id="discovery-progress" class="progress-container">
              <div class="progress-bar"></div>
              <div class="progress-text">Discovering server settings...</div>
            </div>
            
            <div id="discovery-results" class="discovery-results" style="display: none;">
              <h4>Discovered Settings:</h4>
              <div class="settings-display">
                <div class="setting">
                  <label>Server URL:</label>
                  <span id="server-url"></span>
                </div>
                <div class="setting">
                  <label>EWS URL:</label>
                  <span id="ews-url"></span>
                </div>
                <div class="setting">
                  <label>Authentication:</label>
                  <span id="auth-method"></span>
                </div>
              </div>
            </div>
          </div>
        `;

      case 3:
        return `
          <div class="sync-settings">
            <h3>Synchronization Settings</h3>
            
            <div class="form-group">
              <label>
                <input type="checkbox" id="sync-email" checked>
                Synchronize Email
              </label>
            </div>
            
            <div class="form-group">
              <label>
                <input type="checkbox" id="sync-contacts" checked>
                Synchronize Contacts
              </label>
            </div>
            
            <div class="form-group">
              <label>
                <input type="checkbox" id="sync-calendar" checked>
                Synchronize Calendar
              </label>
            </div>
            
            <div class="form-group">
              <label for="sync-interval">Sync Interval</label>
              <select id="sync-interval">
                <option value="60000">1 minute</option>
                <option value="300000" selected>5 minutes</option>
                <option value="600000">10 minutes</option>
                <option value="1800000">30 minutes</option>
                <option value="3600000">1 hour</option>
              </select>
            </div>
          </div>
        `;

      default:
        return '';
    }
  }

  renderActions() {
    const actions = [];

    if (this.currentStep > 1) {
      actions.push(`<button type="button" id="prev-step" class="btn btn-secondary">Previous</button>`);
    }

    if (this.currentStep < 3) {
      if (this.currentStep === 1) {
        actions.push(`<button type="button" id="test-connection" class="btn btn-outline">Test Connection</button>`);
      }
      actions.push(`<button type="button" id="next-step" class="btn btn-primary" disabled>Next</button>`);
    } else {
      actions.push(`<button type="button" id="finish-setup" class="btn btn-success">Finish Setup</button>`);
    }

    return actions.join('');
  }

  renderAccountCard(account) {
    const lastSyncEmail = account.lastSync.email ? new Date(account.lastSync.email).toLocaleString() : 'Never';
    const lastSyncContacts = account.lastSync.contacts ? new Date(account.lastSync.contacts).toLocaleString() : 'Never';
    const lastSyncCalendar = account.lastSync.calendar ? new Date(account.lastSync.calendar).toLocaleString() : 'Never';

    return `
      <div class="account-card">
        <div class="account-header">
          <h3>${account.displayName}</h3>
          <p class="account-email">${account.email}</p>
        </div>
        
        <div class="account-details">
          <div class="sync-info">
            <h4>Last Synchronization</h4>
            <div class="sync-item">
              <span class="sync-type">Email:</span>
              <span class="sync-time">${lastSyncEmail}</span>
            </div>
            <div class="sync-item">
              <span class="sync-type">Contacts:</span>
              <span class="sync-time">${lastSyncContacts}</span>
            </div>
            <div class="sync-item">
              <span class="sync-type">Calendar:</span>
              <span class="sync-time">${lastSyncCalendar}</span>
            </div>
          </div>
        </div>
        
        <div class="account-actions">
          <button class="btn btn-primary sync-account" data-account-id="${account.id}">
            Sync Now
          </button>
          <button class="btn btn-danger remove-account" data-account-id="${account.id}">
            Remove
          </button>
        </div>
      </div>
    `;
  }

  validateForm() {
    const emailInput = document.getElementById('email');
    const passwordInput = document.getElementById('password');
    const nextButton = document.getElementById('next-step');

    if (!emailInput || !passwordInput || !nextButton) return;

    const isValid = emailInput.value.trim() !== '' && 
                   passwordInput.value.trim() !== '' &&
                   emailInput.checkValidity();

    nextButton.disabled = !isValid;
  }

  async testConnection() {
    const email = document.getElementById('email').value;
    const password = document.getElementById('password').value;

    if (!email || !password) {
      this.showError('Please enter email and password');
      return;
    }

    this.showStatus('Testing connection...', 'info');

    try {
      const response = await browser.runtime.sendMessage({
        action: 'testConnection',
        data: { email, password }
      });

      if (response.success) {
        this.showStatus('Connection successful!', 'success');
      } else {
        this.showError('Connection failed: ' + response.error);
      }
    } catch (error) {
      this.showError('Connection test failed: ' + error.message);
    }
  }

  async nextStep() {
    if (this.currentStep === 1) {
      // Collect account data
      this.accountData.email = document.getElementById('email').value;
      this.accountData.password = document.getElementById('password').value;
      this.accountData.displayName = document.getElementById('display-name').value;
      
      // Start autodiscovery
      this.currentStep = 2;
      this.showSetupForm();
      await this.performAutodiscovery();
    } else if (this.currentStep === 2) {
      this.currentStep = 3;
      this.showSetupForm();
    }
  }

  prevStep() {
    if (this.currentStep > 1) {
      this.currentStep--;
      this.showSetupForm();
    }
  }

  async performAutodiscovery() {
    this.showStatus('Discovering server settings...', 'info');

    try {
      const response = await browser.runtime.sendMessage({
        action: 'testConnection',
        data: {
          email: this.accountData.email,
          password: this.accountData.password
        }
      });

      if (response.success) {
        // Display discovered settings
        const resultsDiv = document.getElementById('discovery-results');
        document.getElementById('server-url').textContent = response.serverSettings.serverUrl || 'Auto-detected';
        document.getElementById('ews-url').textContent = response.serverSettings.ewsUrl || 'Auto-detected';
        document.getElementById('auth-method').textContent = response.serverSettings.authMethod || 'OAuth2';
        
        resultsDiv.style.display = 'block';
        document.getElementById('discovery-progress').style.display = 'none';
        
        // Enable next button
        const nextButton = document.getElementById('next-step');
        if (nextButton) {
          nextButton.disabled = false;
        }
        
        this.showStatus('Server discovery completed successfully!', 'success');
      } else {
        this.showError('Server discovery failed: ' + response.error);
      }
    } catch (error) {
      this.showError('Server discovery failed: ' + error.message);
    }
  }

  async finishSetup() {
    // Collect sync settings
    const syncSettings = {
      email: document.getElementById('sync-email').checked,
      contacts: document.getElementById('sync-contacts').checked,
      calendar: document.getElementById('sync-calendar').checked,
      syncInterval: parseInt(document.getElementById('sync-interval').value)
    };

    this.accountData.syncSettings = syncSettings;

    this.showStatus('Setting up account...', 'info');

    try {
      const response = await browser.runtime.sendMessage({
        action: 'setupAccount',
        data: this.accountData
      });

      if (response.success) {
        this.showStatus('Account setup completed successfully!', 'success');
        
        // Reload accounts and show account list
        await this.loadAccounts();
        setTimeout(() => {
          this.showAccountList();
        }, 2000);
      } else {
        this.showError('Account setup failed: ' + response.error);
      }
    } catch (error) {
      this.showError('Account setup failed: ' + error.message);
    }
  }

  async syncAccount(accountId) {
    this.showStatus('Starting synchronization...', 'info');

    try {
      const response = await browser.runtime.sendMessage({
        action: 'syncNow',
        accountId: accountId
      });

      if (response.success) {
        this.showStatus('Synchronization completed successfully!', 'success');
        await this.loadAccounts();
        this.showAccountList();
      } else {
        this.showError('Synchronization failed: ' + response.error);
      }
    } catch (error) {
      this.showError('Synchronization failed: ' + error.message);
    }
  }

  async removeAccount(accountId) {
    if (!confirm('Are you sure you want to remove this account?')) {
      return;
    }

    this.showStatus('Removing account...', 'info');

    try {
      const response = await browser.runtime.sendMessage({
        action: 'removeAccount',
        accountId: accountId
      });

      if (response.success) {
        this.showStatus('Account removed successfully!', 'success');
        await this.loadAccounts();
        this.render();
      } else {
        this.showError('Failed to remove account: ' + response.error);
      }
    } catch (error) {
      this.showError('Failed to remove account: ' + error.message);
    }
  }

  handleSetupSubmit(e) {
    e.preventDefault();
    // Form submission is handled by step navigation
  }

  showStatus(message, type = 'info') {
    const statusDiv = document.getElementById('status-message');
    if (statusDiv) {
      statusDiv.textContent = message;
      statusDiv.className = `status-message ${type}`;
      statusDiv.style.display = 'block';
    }
  }

  showError(message) {
    this.showStatus(message, 'error');
  }
}

// Initialize the UI when the page loads
document.addEventListener('DOMContentLoaded', () => {
  new AccountSetupUI();
});
