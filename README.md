# Syncbird - Exchange/Office365 Integration for Thunderbird

Syncbird is a Thunderbird extension that enables seamless integration with Exchange and Office365 email accounts using Outlook Web Access (OWA) protocols. It provides automatic server discovery and full synchronization of emails, contacts, and calendar events.

## Features

- **Server Autodiscovery**: Automatically discovers Exchange server settings
- **OAuth2 Authentication**: Modern authentication support for Office365
- **Basic Authentication**: Support for on-premises Exchange servers
- **Email Synchronization**: Full bidirectional email sync
- **Contact Management**: Sync contacts between Exchange and Thunderbird
- **Calendar Integration**: Calendar event synchronization
- **Real-time Updates**: Background synchronization with configurable intervals

## Installation

### From Source (Development)

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/syncbird.git
   cd syncbird
   ```

2. In Thunderbird:
   - Open **Tools** → **Add-ons and Themes**
   - Click the gear icon → **Debug Add-ons**
   - Click **Load Temporary Add-on**
   - Select the `manifest.json` file from the cloned directory

### Production Installation

1. Download the latest release from the [releases page](https://github.com/your-username/syncbird/releases)
2. In Thunderbird, go to **Tools** → **Add-ons and Themes**
3. Click the gear icon → **Install Add-on From File**
4. Select the downloaded `.xpi` file

## Configuration

### For Office365 Accounts

1. Open Syncbird from **Tools** → **Add-ons** → **Syncbird**
2. Enter your Office365 email address
3. The extension will automatically detect Office365 settings
4. Complete OAuth2 authentication when prompted
5. Configure synchronization preferences

### For Exchange On-Premises

1. Enter your Exchange email address and password
2. The extension will attempt autodiscovery
3. If autodiscovery fails, you may need to manually configure:
   - EWS URL (typically: `https://your-server/EWS/Exchange.asmx`)
   - Server URL
   - Authentication method

## Configuration for Developers

### OAuth2 Setup

For OAuth2 authentication to work with Office365, you need to register an application with Microsoft:

1. Go to [Azure App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps)
2. Create a new application registration
3. Add the redirect URI: `https://your-extension-id.extensions.allizom.org/`
4. Grant necessary permissions:
   - `https://outlook.office365.com/EWS.AccessAsUser.All`
   - `offline_access`
5. Set the client ID in your environment or configuration

### Environment Variables

The extension supports the following environment variables:

- `MICROSOFT_CLIENT_ID`: Your registered Microsoft application client ID

## File Structure

```
syncbird/
├── manifest.json              # Extension manifest
├── background.js             # Main background script
├── content/
│   ├── account-setup.html    # Account setup page
│   ├── account-setup.js      # Setup page logic
│   └── account-setup.css     # Setup page styles
├── api/
│   ├── exchange-client.js    # High-level Exchange API client
│   ├── ews-soap.js          # EWS SOAP protocol implementation
│   └── autodiscovery.js     # Exchange autodiscovery service
├── services/
│   ├── auth-manager.js      # Authentication management
│   ├── email-sync.js        # Email synchronization service
│   ├── contact-sync.js      # Contact synchronization service
│   └── calendar-sync.js     # Calendar synchronization service
└── utils/
    ├── xml-parser.js        # XML parsing utilities
    └── oauth-flow.js        # OAuth2 flow handler
```

## API Overview

### Main Classes

- **ExchangeClient**: High-level API for Exchange operations
- **EWSClient**: Low-level EWS SOAP protocol implementation
- **Autodiscovery**: Exchange server autodiscovery service
- **AuthManager**: Authentication and token management
- **EmailSync**: Email synchronization service
- **ContactSync**: Contact synchronization service
- **CalendarSync**: Calendar synchronization service

### Key Features

- **Retry Logic**: Automatic retry for transient failures
- **Batch Operations**: Efficient bulk operations where supported
- **Error Handling**: Comprehensive error handling and user feedback
- **Token Management**: Automatic token refresh for OAuth2
- **Incremental Sync**: Smart synchronization to minimize bandwidth

## Troubleshooting

### Common Issues

1. **Authentication Fails**
   - Verify credentials are correct
   - Check if MFA is enabled (may require app-specific password)
   - Ensure OAuth2 client ID is properly configured

2. **Autodiscovery Fails**
   - Check network connectivity
   - Verify domain is correctly configured
   - Try manual server configuration

3. **Sync Issues**
   - Check Thunderbird console for error messages
   - Verify account permissions in Exchange
   - Check firewall and proxy settings

### Debug Mode

Enable debug logging by opening the Browser Console (**Tools** → **Developer Tools** → **Browser Console**) to see detailed logs.

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Make your changes and test thoroughly
4. Commit your changes: `git commit -am 'Add feature'`
5. Push to the branch: `git push origin feature-name`
6. Submit a pull request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- **Issues**: [GitHub Issues](https://github.com/your-username/syncbird/issues)
- **Discussions**: [GitHub Discussions](https://github.com/your-username/syncbird/discussions)
- **Email**: syncbird-support@yourdomain.com

## Acknowledgments

- Microsoft Exchange Web Services (EWS) documentation
- Thunderbird WebExtension API documentation
- OAuth2 and OpenID Connect specifications