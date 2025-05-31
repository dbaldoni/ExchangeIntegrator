/**
 * Exchange Web Services (EWS) SOAP Client
 * Handles low-level SOAP communication with Exchange servers
 */

class EWSClient {
  constructor() {
    this.xmlParser = new XMLParser();
    this.timeoutMs = 60000; // 60 seconds timeout for EWS requests
    this.schemaNamespaces = {
      's': 'http://schemas.xmlsoap.org/soap/envelope/',
      't': 'http://schemas.microsoft.com/exchange/services/2006/types',
      'm': 'http://schemas.microsoft.com/exchange/services/2006/messages'
    };
  }

  /**
   * Get folder information
   */
  async getFolder(account, folderId) {
    const soapBody = `
      <m:GetFolder>
        <m:FolderShape>
          <t:BaseShape>AllProperties</t:BaseShape>
        </m:FolderShape>
        <m:FolderIds>
          <t:DistinguishedFolderId Id="${this.escapeXml(folderId)}" />
        </m:FolderIds>
      </m:GetFolder>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseFolderResponse(response);
    } catch (error) {
      console.error('GetFolder failed:', error);
      throw error;
    }
  }

  /**
   * Get folder hierarchy
   */
  async getFolderHierarchy(account) {
    const soapBody = `
      <m:FindFolder Traversal="Deep">
        <m:FolderShape>
          <t:BaseShape>Default</t:BaseShape>
          <t:AdditionalProperties>
            <t:FieldURI FieldURI="folder:DisplayName"/>
            <t:FieldURI FieldURI="folder:TotalCount"/>
            <t:FieldURI FieldURI="folder:UnreadCount"/>
            <t:FieldURI FieldURI="folder:FolderClass"/>
          </t:AdditionalProperties>
        </m:FolderShape>
        <m:ParentFolderIds>
          <t:DistinguishedFolderId Id="msgfolderroot"/>
        </m:ParentFolderIds>
      </m:FindFolder>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseFolderHierarchyResponse(response);
    } catch (error) {
      console.error('GetFolderHierarchy failed:', error);
      throw error;
    }
  }

  /**
   * Get messages from folder
   */
  async getMessages(account, folderId, options = {}) {
    const maxItems = options.maxItems || 50;
    const offset = options.offset || 0;
    const sortOrder = options.sortOrder || 'Descending';

    const soapBody = `
      <m:FindItem Traversal="Shallow">
        <m:ItemShape>
          <t:BaseShape>Default</t:BaseShape>
          <t:AdditionalProperties>
            <t:FieldURI FieldURI="message:Subject"/>
            <t:FieldURI FieldURI="message:DateTimeReceived"/>
            <t:FieldURI FieldURI="message:DateTimeSent"/>
            <t:FieldURI FieldURI="message:From"/>
            <t:FieldURI FieldURI="message:Sender"/>
            <t:FieldURI FieldURI="message:ToRecipients"/>
            <t:FieldURI FieldURI="message:CcRecipients"/>
            <t:FieldURI FieldURI="message:IsRead"/>
            <t:FieldURI FieldURI="message:Importance"/>
            <t:FieldURI FieldURI="message:Size"/>
            <t:FieldURI FieldURI="message:HasAttachments"/>
          </t:AdditionalProperties>
        </m:ItemShape>
        <m:IndexedPageItemView MaxEntriesReturned="${maxItems}" Offset="${offset}" BasePoint="Beginning"/>
        <m:SortOrder>
          <t:FieldOrder Order="${sortOrder}">
            <t:FieldURI FieldURI="message:DateTimeReceived"/>
          </t:FieldOrder>
        </m:SortOrder>
        <m:ParentFolderIds>
          <t:DistinguishedFolderId Id="${this.escapeXml(folderId)}"/>
        </m:ParentFolderIds>
      </m:FindItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseMessagesResponse(response);
    } catch (error) {
      console.error('GetMessages failed:', error);
      throw error;
    }
  }

  /**
   * Get specific message details
   */
  async getMessage(account, messageId) {
    const soapBody = `
      <m:GetItem>
        <m:ItemShape>
          <t:BaseShape>AllProperties</t:BaseShape>
          <t:IncludeMimeContent>true</t:IncludeMimeContent>
        </m:ItemShape>
        <m:ItemIds>
          <t:ItemId Id="${this.escapeXml(messageId)}"/>
        </m:ItemIds>
      </m:GetItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseMessageResponse(response);
    } catch (error) {
      console.error('GetMessage failed:', error);
      throw error;
    }
  }

  /**
   * Send a message
   */
  async sendMessage(account, message) {
    const toRecipients = message.to.map(addr => 
      `<t:Mailbox><t:EmailAddress>${this.escapeXml(addr)}</t:EmailAddress></t:Mailbox>`
    ).join('');

    const ccRecipients = (message.cc || []).map(addr => 
      `<t:Mailbox><t:EmailAddress>${this.escapeXml(addr)}</t:EmailAddress></t:Mailbox>`
    ).join('');

    const soapBody = `
      <m:CreateItem MessageDisposition="SendAndSaveCopy">
        <m:Items>
          <t:Message>
            <t:Subject>${this.escapeXml(message.subject)}</t:Subject>
            <t:Body BodyType="${message.bodyType || 'HTML'}">${this.escapeXml(message.body)}</t:Body>
            <t:ToRecipients>${toRecipients}</t:ToRecipients>
            ${ccRecipients ? `<t:CcRecipients>${ccRecipients}</t:CcRecipients>` : ''}
            <t:IsRead>true</t:IsRead>
          </t:Message>
        </m:Items>
      </m:CreateItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseCreateItemResponse(response);
    } catch (error) {
      console.error('SendMessage failed:', error);
      throw error;
    }
  }

  /**
   * Move message to folder
   */
  async moveMessage(account, messageId, targetFolderId) {
    const soapBody = `
      <m:MoveItem>
        <m:ToFolderId>
          <t:DistinguishedFolderId Id="${this.escapeXml(targetFolderId)}"/>
        </m:ToFolderId>
        <m:ItemIds>
          <t:ItemId Id="${this.escapeXml(messageId)}"/>
        </m:ItemIds>
      </m:MoveItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseMoveItemResponse(response);
    } catch (error) {
      console.error('MoveMessage failed:', error);
      throw error;
    }
  }

  /**
   * Mark message as read/unread
   */
  async markMessage(account, messageId, isRead) {
    const soapBody = `
      <m:UpdateItem ConflictResolution="AutoResolve" MessageDisposition="SaveOnly">
        <m:ItemChanges>
          <t:ItemChange>
            <t:ItemId Id="${this.escapeXml(messageId)}"/>
            <t:Updates>
              <t:SetItemField>
                <t:FieldURI FieldURI="message:IsRead"/>
                <t:Message>
                  <t:IsRead>${isRead}</t:IsRead>
                </t:Message>
              </t:SetItemField>
            </t:Updates>
          </t:ItemChange>
        </m:ItemChanges>
      </m:UpdateItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseUpdateItemResponse(response);
    } catch (error) {
      console.error('MarkMessage failed:', error);
      throw error;
    }
  }

  /**
   * Delete message
   */
  async deleteMessage(account, messageId) {
    const soapBody = `
      <m:DeleteItem DeleteType="MoveToDeletedItems">
        <m:ItemIds>
          <t:ItemId Id="${this.escapeXml(messageId)}"/>
        </m:ItemIds>
      </m:DeleteItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseDeleteItemResponse(response);
    } catch (error) {
      console.error('DeleteMessage failed:', error);
      throw error;
    }
  }

  /**
   * Get contacts
   */
  async getContacts(account, options = {}) {
    const maxItems = options.maxItems || 100;
    const offset = options.offset || 0;

    const soapBody = `
      <m:FindItem Traversal="Shallow">
        <m:ItemShape>
          <t:BaseShape>AllProperties</t:BaseShape>
        </m:ItemShape>
        <m:IndexedPageItemView MaxEntriesReturned="${maxItems}" Offset="${offset}" BasePoint="Beginning"/>
        <m:ParentFolderIds>
          <t:DistinguishedFolderId Id="contacts"/>
        </m:ParentFolderIds>
      </m:FindItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseContactsResponse(response);
    } catch (error) {
      console.error('GetContacts failed:', error);
      throw error;
    }
  }

  /**
   * Create contact
   */
  async createContact(account, contact) {
    const soapBody = `
      <m:CreateItem>
        <m:SavedItemFolderId>
          <t:DistinguishedFolderId Id="contacts"/>
        </m:SavedItemFolderId>
        <m:Items>
          <t:Contact>
            <t:DisplayName>${this.escapeXml(contact.displayName)}</t:DisplayName>
            <t:GivenName>${this.escapeXml(contact.firstName || '')}</t:GivenName>
            <t:Surname>${this.escapeXml(contact.lastName || '')}</t:Surname>
            <t:EmailAddresses>
              <t:Entry Key="EmailAddress1">${this.escapeXml(contact.email)}</t:Entry>
            </t:EmailAddresses>
            ${contact.phone ? `<t:PhoneNumbers><t:Entry Key="BusinessPhone">${this.escapeXml(contact.phone)}</t:Entry></t:PhoneNumbers>` : ''}
            ${contact.company ? `<t:CompanyName>${this.escapeXml(contact.company)}</t:CompanyName>` : ''}
          </t:Contact>
        </m:Items>
      </m:CreateItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseCreateItemResponse(response);
    } catch (error) {
      console.error('CreateContact failed:', error);
      throw error;
    }
  }

  /**
   * Update contact
   */
  async updateContact(account, contactId, updates) {
    const updateFields = Object.keys(updates).map(field => {
      const fieldMapping = {
        displayName: 'contacts:DisplayName',
        firstName: 'contacts:GivenName',
        lastName: 'contacts:Surname',
        email: 'contacts:EmailAddress1',
        phone: 'contacts:BusinessPhone',
        company: 'contacts:CompanyName'
      };

      const ewsField = fieldMapping[field];
      if (!ewsField) return '';

      return `
        <t:SetItemField>
          <t:FieldURI FieldURI="${ewsField}"/>
          <t:Contact>
            <t:${ewsField.split(':')[1]}>${this.escapeXml(updates[field])}</t:${ewsField.split(':')[1]}>
          </t:Contact>
        </t:SetItemField>
      `;
    }).join('');

    const soapBody = `
      <m:UpdateItem ConflictResolution="AutoResolve">
        <m:ItemChanges>
          <t:ItemChange>
            <t:ItemId Id="${this.escapeXml(contactId)}"/>
            <t:Updates>
              ${updateFields}
            </t:Updates>
          </t:ItemChange>
        </m:ItemChanges>
      </m:UpdateItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseUpdateItemResponse(response);
    } catch (error) {
      console.error('UpdateContact failed:', error);
      throw error;
    }
  }

  /**
   * Delete contact
   */
  async deleteContact(account, contactId) {
    const soapBody = `
      <m:DeleteItem DeleteType="MoveToDeletedItems">
        <m:ItemIds>
          <t:ItemId Id="${this.escapeXml(contactId)}"/>
        </m:ItemIds>
      </m:DeleteItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseDeleteItemResponse(response);
    } catch (error) {
      console.error('DeleteContact failed:', error);
      throw error;
    }
  }

  /**
   * Get calendar items
   */
  async getCalendarItems(account, startDate, endDate, options = {}) {
    const start = startDate.toISOString();
    const end = endDate.toISOString();

    const soapBody = `
      <m:FindItem Traversal="Shallow">
        <m:ItemShape>
          <t:BaseShape>AllProperties</t:BaseShape>
        </m:ItemShape>
        <m:CalendarView StartDate="${start}" EndDate="${end}"/>
        <m:ParentFolderIds>
          <t:DistinguishedFolderId Id="calendar"/>
        </m:ParentFolderIds>
      </m:FindItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseCalendarItemsResponse(response);
    } catch (error) {
      console.error('GetCalendarItems failed:', error);
      throw error;
    }
  }

  /**
   * Create calendar item
   */
  async createCalendarItem(account, item) {
    const start = new Date(item.start).toISOString();
    const end = new Date(item.end).toISOString();

    const soapBody = `
      <m:CreateItem SendMeetingInvitations="SendToNone">
        <m:SavedItemFolderId>
          <t:DistinguishedFolderId Id="calendar"/>
        </m:SavedItemFolderId>
        <m:Items>
          <t:CalendarItem>
            <t:Subject>${this.escapeXml(item.subject)}</t:Subject>
            <t:Body BodyType="${item.bodyType || 'Text'}">${this.escapeXml(item.body || '')}</t:Body>
            <t:Start>${start}</t:Start>
            <t:End>${end}</t:End>
            <t:Location>${this.escapeXml(item.location || '')}</t:Location>
            <t:LegacyFreeBusyStatus>${item.freeBusyStatus || 'Busy'}</t:LegacyFreeBusyStatus>
          </t:CalendarItem>
        </m:Items>
      </m:CreateItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseCreateItemResponse(response);
    } catch (error) {
      console.error('CreateCalendarItem failed:', error);
      throw error;
    }
  }

  /**
   * Update calendar item
   */
  async updateCalendarItem(account, itemId, updates) {
    const updateFields = Object.keys(updates).map(field => {
      let ewsField, value;
      
      switch (field) {
        case 'subject':
          ewsField = 'item:Subject';
          value = this.escapeXml(updates[field]);
          break;
        case 'body':
          ewsField = 'item:Body';
          value = this.escapeXml(updates[field]);
          break;
        case 'start':
          ewsField = 'calendar:Start';
          value = new Date(updates[field]).toISOString();
          break;
        case 'end':
          ewsField = 'calendar:End';
          value = new Date(updates[field]).toISOString();
          break;
        case 'location':
          ewsField = 'calendar:Location';
          value = this.escapeXml(updates[field]);
          break;
        default:
          return '';
      }

      return `
        <t:SetItemField>
          <t:FieldURI FieldURI="${ewsField}"/>
          <t:CalendarItem>
            <t:${ewsField.split(':')[1]}>${value}</t:${ewsField.split(':')[1]}>
          </t:CalendarItem>
        </t:SetItemField>
      `;
    }).join('');

    const soapBody = `
      <m:UpdateItem ConflictResolution="AutoResolve" SendMeetingInvitationsOrCancellations="SendToNone">
        <m:ItemChanges>
          <t:ItemChange>
            <t:ItemId Id="${this.escapeXml(itemId)}"/>
            <t:Updates>
              ${updateFields}
            </t:Updates>
          </t:ItemChange>
        </m:ItemChanges>
      </m:UpdateItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseUpdateItemResponse(response);
    } catch (error) {
      console.error('UpdateCalendarItem failed:', error);
      throw error;
    }
  }

  /**
   * Delete calendar item
   */
  async deleteCalendarItem(account, itemId) {
    const soapBody = `
      <m:DeleteItem DeleteType="MoveToDeletedItems" SendMeetingCancellations="SendToNone">
        <m:ItemIds>
          <t:ItemId Id="${this.escapeXml(itemId)}"/>
        </m:ItemIds>
      </m:DeleteItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseDeleteItemResponse(response);
    } catch (error) {
      console.error('DeleteCalendarItem failed:', error);
      throw error;
    }
  }

  /**
   * Subscribe to notifications
   */
  async subscribeToNotifications(account, folders = ['inbox']) {
    const folderIds = folders.map(folderId => 
      `<t:DistinguishedFolderId Id="${this.escapeXml(folderId)}"/>`
    ).join('');

    const soapBody = `
      <m:Subscribe>
        <m:PushSubscriptionRequest>
          <t:FolderIds>
            ${folderIds}
          </t:FolderIds>
          <t:EventTypes>
            <t:EventType>NewMailEvent</t:EventType>
            <t:EventType>ModifiedEvent</t:EventType>
            <t:EventType>DeletedEvent</t:EventType>
            <t:EventType>MovedEvent</t:EventType>
          </t:EventTypes>
          <t:StatusFrequency>1</t:StatusFrequency>
          <t:URL>https://localhost/notifications</t:URL>
        </m:PushSubscriptionRequest>
      </m:Subscribe>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseSubscribeResponse(response);
    } catch (error) {
      console.error('SubscribeToNotifications failed:', error);
      throw error;
    }
  }

  /**
   * Get notification events
   */
  async getNotificationEvents(account, subscriptionId) {
    const soapBody = `
      <m:GetEvents>
        <m:SubscriptionId>${this.escapeXml(subscriptionId)}</m:SubscriptionId>
        <m:Watermark>AQAAAA==</m:Watermark>
      </m:GetEvents>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseEventsResponse(response);
    } catch (error) {
      console.error('GetNotificationEvents failed:', error);
      throw error;
    }
  }

  /**
   * Unsubscribe from notifications
   */
  async unsubscribeFromNotifications(account, subscriptionId) {
    const soapBody = `
      <m:Unsubscribe>
        <m:SubscriptionId>${this.escapeXml(subscriptionId)}</m:SubscriptionId>
      </m:Unsubscribe>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseUnsubscribeResponse(response);
    } catch (error) {
      console.error('UnsubscribeFromNotifications failed:', error);
      throw error;
    }
  }

  /**
   * Get user settings
   */
  async getUserSettings(account) {
    const soapBody = `
      <m:GetUserConfiguration>
        <m:UserConfigurationName Name="UserOptions">
          <t:DistinguishedFolderId Id="inbox"/>
        </m:UserConfigurationName>
        <m:UserConfigurationProperties>All</m:UserConfigurationProperties>
      </m:GetUserConfiguration>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseUserConfigurationResponse(response);
    } catch (error) {
      console.error('GetUserSettings failed:', error);
      throw error;
    }
  }

  /**
   * Make EWS SOAP request
   */
  async makeEWSRequest(account, soapBody) {
    const soapEnvelope = this.buildSoapEnvelope(soapBody);
    
    try {
      const response = await fetch(account.serverSettings.ewsUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/xml; charset=utf-8',
          'SOAPAction': '',
          'User-Agent': 'ExchangeThunderbirdExtension/1.0',
          'Authorization': await this.getAuthorizationHeader(account)
        },
        body: soapEnvelope,
        signal: AbortSignal.timeout(this.timeoutMs)
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const responseText = await response.text();
      return responseText;
    } catch (error) {
      console.error('EWS request failed:', error);
      throw error;
    }
  }

  /**
   * Build SOAP envelope
   */
  buildSoapEnvelope(body) {
    return `<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <t:RequestServerVersion Version="Exchange2016" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"/>
  </s:Header>
  <s:Body xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    ${body}
  </s:Body>
</s:Envelope>`;
  }

  /**
   * Get authorization header
   */
  async getAuthorizationHeader(account) {
    if (account.authToken) {
      return `Bearer ${account.authToken}`;
    } else {
      // Fallback to basic auth (for older Exchange servers)
      const credentials = btoa(`${account.email}:${account.password}`);
      return `Basic ${credentials}`;
    }
  }

  /**
   * Parse folder response
   */
  parseFolderResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    const folderElements = doc.getElementsByTagName('t:Folder');
    
    if (folderElements.length === 0) {
      throw new Error('No folder found in response');
    }

    const folder = folderElements[0];
    return {
      success: true,
      folder: this.extractFolderInfo(folder)
    };
  }

  /**
   * Parse folder hierarchy response
   */
  parseFolderHierarchyResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    const folderElements = doc.getElementsByTagName('t:Folder');
    
    const folders = [];
    for (let i = 0; i < folderElements.length; i++) {
      folders.push(this.extractFolderInfo(folderElements[i]));
    }

    return {
      success: true,
      folders: folders
    };
  }

  /**
   * Parse messages response
   */
  parseMessagesResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    const messageElements = doc.getElementsByTagName('t:Message');
    
    const messages = [];
    for (let i = 0; i < messageElements.length; i++) {
      messages.push(this.extractMessageInfo(messageElements[i]));
    }

    return {
      success: true,
      messages: messages
    };
  }

  /**
   * Parse single message response
   */
  parseMessageResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    const messageElements = doc.getElementsByTagName('t:Message');
    
    if (messageElements.length === 0) {
      throw new Error('No message found in response');
    }

    return {
      success: true,
      message: this.extractMessageInfo(messageElements[0], true)
    };
  }

  /**
   * Parse contacts response
   */
  parseContactsResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    const contactElements = doc.getElementsByTagName('t:Contact');
    
    const contacts = [];
    for (let i = 0; i < contactElements.length; i++) {
      contacts.push(this.extractContactInfo(contactElements[i]));
    }

    return {
      success: true,
      contacts: contacts
    };
  }

  /**
   * Parse calendar items response
   */
  parseCalendarItemsResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    const calendarElements = doc.getElementsByTagName('t:CalendarItem');
    
    const items = [];
    for (let i = 0; i < calendarElements.length; i++) {
      items.push(this.extractCalendarItemInfo(calendarElements[i]));
    }

    return {
      success: true,
      items: items
    };
  }

  /**
   * Extract folder information from XML element
   */
  extractFolderInfo(folderElement) {
    return {
      id: this.getElementText(folderElement, 't:FolderId'),
      displayName: this.getElementText(folderElement, 't:DisplayName'),
      totalCount: parseInt(this.getElementText(folderElement, 't:TotalCount')) || 0,
      unreadCount: parseInt(this.getElementText(folderElement, 't:UnreadCount')) || 0,
      folderClass: this.getElementText(folderElement, 't:FolderClass')
    };
  }

  /**
   * Extract message information from XML element
   */
  extractMessageInfo(messageElement, includeBody = false) {
    const message = {
      id: this.getElementAttribute(messageElement, 't:ItemId', 'Id'),
      subject: this.getElementText(messageElement, 't:Subject'),
      dateTimeReceived: this.getElementText(messageElement, 't:DateTimeReceived'),
      dateTimeSent: this.getElementText(messageElement, 't:DateTimeSent'),
      from: this.extractEmailAddress(messageElement, 't:From'),
      sender: this.extractEmailAddress(messageElement, 't:Sender'),
      toRecipients: this.extractEmailAddresses(messageElement, 't:ToRecipients'),
      ccRecipients: this.extractEmailAddresses(messageElement, 't:CcRecipients'),
      isRead: this.getElementText(messageElement, 't:IsRead') === 'true',
      importance: this.getElementText(messageElement, 't:Importance'),
      size: parseInt(this.getElementText(messageElement, 't:Size')) || 0,
      hasAttachments: this.getElementText(messageElement, 't:HasAttachments') === 'true'
    };

    if (includeBody) {
      message.body = this.getElementText(messageElement, 't:Body');
      message.bodyType = this.getElementAttribute(messageElement, 't:Body', 'BodyType');
    }

    return message;
  }

  /**
   * Extract contact information from XML element
   */
  extractContactInfo(contactElement) {
    return {
      id: this.getElementAttribute(contactElement, 't:ItemId', 'Id'),
      displayName: this.getElementText(contactElement, 't:DisplayName'),
      firstName: this.getElementText(contactElement, 't:GivenName'),
      lastName: this.getElementText(contactElement, 't:Surname'),
      email: this.extractEmailFromCollection(contactElement, 't:EmailAddresses', 'EmailAddress1'),
      phone: this.extractPhoneFromCollection(contactElement, 't:PhoneNumbers', 'BusinessPhone'),
      company: this.getElementText(contactElement, 't:CompanyName')
    };
  }

  /**
   * Extract calendar item information from XML element
   */
  extractCalendarItemInfo(calendarElement) {
    return {
      id: this.getElementAttribute(calendarElement, 't:ItemId', 'Id'),
      subject: this.getElementText(calendarElement, 't:Subject'),
      body: this.getElementText(calendarElement, 't:Body'),
      start: this.getElementText(calendarElement, 't:Start'),
      end: this.getElementText(calendarElement, 't:End'),
      location: this.getElementText(calendarElement, 't:Location'),
      freeBusyStatus: this.getElementText(calendarElement, 't:LegacyFreeBusyStatus'),
      organizer: this.extractEmailAddress(calendarElement, 't:Organizer'),
      attendees: this.extractEmailAddresses(calendarElement, 't:RequiredAttendees')
    };
  }

  /**
   * Extract email address from XML element
   */
  extractEmailAddress(parentElement, tagName) {
    const element = parentElement.getElementsByTagName(tagName)[0];
    if (!element) return null;

    const mailbox = element.getElementsByTagName('t:Mailbox')[0];
    if (!mailbox) return null;

    return {
      name: this.getElementText(mailbox, 't:Name'),
      email: this.getElementText(mailbox, 't:EmailAddress')
    };
  }

  /**
   * Extract multiple email addresses from XML element
   */
  extractEmailAddresses(parentElement, tagName) {
    const container = parentElement.getElementsByTagName(tagName)[0];
    if (!container) return [];

    const mailboxes = container.getElementsByTagName('t:Mailbox');
    const addresses = [];

    for (let i = 0; i < mailboxes.length; i++) {
      addresses.push({
        name: this.getElementText(mailboxes[i], 't:Name'),
        email: this.getElementText(mailboxes[i], 't:EmailAddress')
      });
    }

    return addresses;
  }

  /**
   * Extract email from collection
   */
  extractEmailFromCollection(parentElement, collectionTagName, key) {
    const collection = parentElement.getElementsByTagName(collectionTagName)[0];
    if (!collection) return null;

    const entries = collection.getElementsByTagName('t:Entry');
    for (let i = 0; i < entries.length; i++) {
      if (entries[i].getAttribute('Key') === key) {
        return entries[i].textContent;
      }
    }

    return null;
  }

  /**
   * Extract phone from collection
   */
  extractPhoneFromCollection(parentElement, collectionTagName, key) {
    return this.extractEmailFromCollection(parentElement, collectionTagName, key);
  }

  /**
   * Parse generic create item response
   */
  parseCreateItemResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    const itemElements = doc.getElementsByTagName('t:ItemId');
    
    if (itemElements.length === 0) {
      throw new Error('No item ID found in create response');
    }

    return {
      success: true,
      itemId: itemElements[0].getAttribute('Id')
    };
  }

  /**
   * Parse generic update item response
   */
  parseUpdateItemResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    const responseCodeElements = doc.getElementsByTagName('m:ResponseCode');
    
    if (responseCodeElements.length === 0) {
      throw new Error('No response code found in update response');
    }

    const responseCode = responseCodeElements[0].textContent;
    return {
      success: responseCode === 'NoError',
      responseCode: responseCode
    };
  }

  /**
   * Parse generic delete item response
   */
  parseDeleteItemResponse(responseXml) {
    return this.parseUpdateItemResponse(responseXml);
  }

  /**
   * Parse move item response
   */
  parseMoveItemResponse(responseXml) {
    return this.parseUpdateItemResponse(responseXml);
  }

  /**
   * Parse subscribe response
   */
  parseSubscribeResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    const subscriptionIdElements = doc.getElementsByTagName('t:SubscriptionId');
    
    if (subscriptionIdElements.length === 0) {
      throw new Error('No subscription ID found in subscribe response');
    }

    return {
      success: true,
      subscriptionId: subscriptionIdElements[0].textContent
    };
  }

  /**
   * Parse events response
   */
  parseEventsResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    // Parse notification events - implementation would depend on specific event types
    return {
      success: true,
      events: [] // Simplified for this example
    };
  }

  /**
   * Parse unsubscribe response
   */
  parseUnsubscribeResponse(responseXml) {
    return this.parseUpdateItemResponse(responseXml);
  }

  /**
   * Parse user configuration response
   */
  parseUserConfigurationResponse(responseXml) {
    const doc = this.xmlParser.parseXML(responseXml);
    return {
      success: true,
      configuration: {} // Simplified for this example
    };
  }

  /**
   * Get text content of first matching element
   */
  getElementText(parentElement, tagName) {
    const elements = parentElement.getElementsByTagName(tagName);
    return elements.length > 0 ? elements[0].textContent : null;
  }

  /**
   * Get attribute value of first matching element
   */
  getElementAttribute(parentElement, tagName, attributeName) {
    const elements = parentElement.getElementsByTagName(tagName);
    return elements.length > 0 ? elements[0].getAttribute(attributeName) : null;
  }

  /**
   * Escape XML special characters
   */
  escapeXml(text) {
    if (!text) return '';
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /**
   * Batch operations for better performance
   */
  async getMessagesBatch(account, messageIds) {
    const itemIds = messageIds.map(id => `<t:ItemId Id="${this.escapeXml(id)}"/>`).join('');
    
    const soapBody = `
      <m:GetItem>
        <m:ItemShape>
          <t:BaseShape>Default</t:BaseShape>
        </m:ItemShape>
        <m:ItemIds>
          ${itemIds}
        </m:ItemIds>
      </m:GetItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseMessagesResponse(response);
    } catch (error) {
      console.error('GetMessagesBatch failed:', error);
      throw error;
    }
  }

  async markMessagesBatch(account, operations) {
    const itemChanges = operations.map(op => `
      <t:ItemChange>
        <t:ItemId Id="${this.escapeXml(op.messageId)}"/>
        <t:Updates>
          <t:SetItemField>
            <t:FieldURI FieldURI="message:IsRead"/>
            <t:Message>
              <t:IsRead>${op.isRead}</t:IsRead>
            </t:Message>
          </t:SetItemField>
        </t:Updates>
      </t:ItemChange>
    `).join('');

    const soapBody = `
      <m:UpdateItem ConflictResolution="AutoResolve" MessageDisposition="SaveOnly">
        <m:ItemChanges>
          ${itemChanges}
        </m:ItemChanges>
      </m:UpdateItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseUpdateItemResponse(response);
    } catch (error) {
      console.error('MarkMessagesBatch failed:', error);
      throw error;
    }
  }

  async moveMessagesBatch(account, operations) {
    // Group by target folder for efficient batch operations
    const groupedOperations = {};
    operations.forEach(op => {
      if (!groupedOperations[op.targetFolderId]) {
        groupedOperations[op.targetFolderId] = [];
      }
      groupedOperations[op.targetFolderId].push(op.messageId);
    });

    const results = [];
    for (const [targetFolderId, messageIds] of Object.entries(groupedOperations)) {
      const itemIds = messageIds.map(id => `<t:ItemId Id="${this.escapeXml(id)}"/>`).join('');
      
      const soapBody = `
        <m:MoveItem>
          <m:ToFolderId>
            <t:DistinguishedFolderId Id="${this.escapeXml(targetFolderId)}"/>
          </m:ToFolderId>
          <m:ItemIds>
            ${itemIds}
          </m:ItemIds>
        </m:MoveItem>
      `;

      try {
        const response = await this.makeEWSRequest(account, soapBody);
        const result = this.parseMoveItemResponse(response);
        results.push(result);
      } catch (error) {
        console.error('MoveMessagesBatch failed:', error);
        results.push({ success: false, error: error.message });
      }
    }

    return results;
  }

  async deleteMessagesBatch(account, operations) {
    const itemIds = operations.map(op => `<t:ItemId Id="${this.escapeXml(op.messageId)}"/>`).join('');
    
    const soapBody = `
      <m:DeleteItem DeleteType="MoveToDeletedItems">
        <m:ItemIds>
          ${itemIds}
        </m:ItemIds>
      </m:DeleteItem>
    `;

    try {
      const response = await this.makeEWSRequest(account, soapBody);
      return this.parseDeleteItemResponse(response);
    } catch (error) {
      console.error('DeleteMessagesBatch failed:', error);
      throw error;
    }
  }
}
