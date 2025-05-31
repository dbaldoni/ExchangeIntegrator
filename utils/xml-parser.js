/**
 * XML Parser Utility
 * Provides XML parsing functionality for EWS SOAP responses and autodiscovery
 */

class XMLParser {
  constructor() {
    this.namespaces = {
      's': 'http://schemas.xmlsoap.org/soap/envelope/',
      't': 'http://schemas.microsoft.com/exchange/services/2006/types',
      'm': 'http://schemas.microsoft.com/exchange/services/2006/messages',
      'autodiscover': 'http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a'
    };
  }

  /**
   * Parse XML string into DOM document
   */
  parseXML(xmlString) {
    try {
      // Remove BOM if present
      const cleanXml = xmlString.replace(/^\uFEFF/, '');
      
      // Create DOM parser
      const parser = new DOMParser();
      const doc = parser.parseFromString(cleanXml, 'text/xml');
      
      // Check for parsing errors
      const parseError = doc.querySelector('parsererror');
      if (parseError) {
        throw new Error('XML parsing error: ' + parseError.textContent);
      }

      return doc;

    } catch (error) {
      console.error('Failed to parse XML:', error);
      throw new Error('XML parsing failed: ' + error.message);
    }
  }

  /**
   * Get text content of first matching element by tag name
   */
  getElementText(parentElement, tagName, defaultValue = null) {
    try {
      const elements = parentElement.getElementsByTagName(tagName);
      if (elements.length > 0) {
        return elements[0].textContent || defaultValue;
      }
      return defaultValue;
    } catch (error) {
      console.warn('Failed to get element text for:', tagName, error);
      return defaultValue;
    }
  }

  /**
   * Get text content of all matching elements by tag name
   */
  getElementsText(parentElement, tagName) {
    try {
      const elements = parentElement.getElementsByTagName(tagName);
      const textValues = [];
      
      for (let i = 0; i < elements.length; i++) {
        const text = elements[i].textContent;
        if (text) {
          textValues.push(text);
        }
      }
      
      return textValues;
    } catch (error) {
      console.warn('Failed to get elements text for:', tagName, error);
      return [];
    }
  }

  /**
   * Get attribute value of first matching element
   */
  getElementAttribute(parentElement, tagName, attributeName, defaultValue = null) {
    try {
      const elements = parentElement.getElementsByTagName(tagName);
      if (elements.length > 0) {
        return elements[0].getAttribute(attributeName) || defaultValue;
      }
      return defaultValue;
    } catch (error) {
      console.warn('Failed to get element attribute for:', tagName, attributeName, error);
      return defaultValue;
    }
  }

  /**
   * Get all attributes of first matching element
   */
  getElementAttributes(parentElement, tagName) {
    try {
      const elements = parentElement.getElementsByTagName(tagName);
      if (elements.length > 0) {
        const element = elements[0];
        const attributes = {};
        
        for (let i = 0; i < element.attributes.length; i++) {
          const attr = element.attributes[i];
          attributes[attr.name] = attr.value;
        }
        
        return attributes;
      }
      return {};
    } catch (error) {
      console.warn('Failed to get element attributes for:', tagName, error);
      return {};
    }
  }

  /**
   * Check if element exists
   */
  elementExists(parentElement, tagName) {
    try {
      const elements = parentElement.getElementsByTagName(tagName);
      return elements.length > 0;
    } catch (error) {
      return false;
    }
  }

  /**
   * Get child elements by tag name
   */
  getChildElements(parentElement, tagName) {
    try {
      const children = [];
      const elements = parentElement.getElementsByTagName(tagName);
      
      for (let i = 0; i < elements.length; i++) {
        children.push(elements[i]);
      }
      
      return children;
    } catch (error) {
      console.warn('Failed to get child elements for:', tagName, error);
      return [];
    }
  }

  /**
   * Parse SOAP fault from response
   */
  parseSOAPFault(doc) {
    try {
      const faultElements = doc.getElementsByTagName('soap:Fault');
      if (faultElements.length === 0) {
        // Try without namespace prefix
        const altFaultElements = doc.getElementsByTagName('Fault');
        if (altFaultElements.length === 0) {
          return null;
        }
        return this.extractFaultInfo(altFaultElements[0]);
      }
      
      return this.extractFaultInfo(faultElements[0]);
    } catch (error) {
      console.error('Failed to parse SOAP fault:', error);
      return null;
    }
  }

  /**
   * Extract fault information from fault element
   */
  extractFaultInfo(faultElement) {
    try {
      const fault = {
        code: this.getElementText(faultElement, 'faultcode') || 
              this.getElementText(faultElement, 'soap:Code'),
        string: this.getElementText(faultElement, 'faultstring') || 
                this.getElementText(faultElement, 'soap:Reason'),
        detail: this.getElementText(faultElement, 'detail') || 
                this.getElementText(faultElement, 'soap:Detail')
      };

      return fault;
    } catch (error) {
      console.error('Failed to extract fault info:', error);
      return null;
    }
  }

  /**
   * Parse EWS response code and message
   */
  parseEWSResponseCode(doc) {
    try {
      const responseElements = doc.getElementsByTagName('m:ResponseCode');
      if (responseElements.length === 0) {
        return null;
      }

      const responseCode = responseElements[0].textContent;
      const messageElements = doc.getElementsByTagName('m:MessageText');
      const messageText = messageElements.length > 0 ? messageElements[0].textContent : '';

      return {
        code: responseCode,
        message: messageText,
        isSuccess: responseCode === 'NoError'
      };
    } catch (error) {
      console.error('Failed to parse EWS response code:', error);
      return null;
    }
  }

  /**
   * Extract EWS items from response
   */
  extractEWSItems(doc, itemType = 'Item') {
    try {
      const items = [];
      const itemElements = doc.getElementsByTagName(`t:${itemType}`);

      for (let i = 0; i < itemElements.length; i++) {
        const item = this.extractItemData(itemElements[i]);
        if (item) {
          items.push(item);
        }
      }

      return items;
    } catch (error) {
      console.error('Failed to extract EWS items:', error);
      return [];
    }
  }

  /**
   * Extract item data from XML element
   */
  extractItemData(itemElement) {
    try {
      const item = {};

      // Common item properties
      item.id = this.getElementAttribute(itemElement, 't:ItemId', 'Id');
      item.changeKey = this.getElementAttribute(itemElement, 't:ItemId', 'ChangeKey');
      item.subject = this.getElementText(itemElement, 't:Subject');
      item.body = this.getElementText(itemElement, 't:Body');
      item.importance = this.getElementText(itemElement, 't:Importance');
      item.sensitivity = this.getElementText(itemElement, 't:Sensitivity');
      item.size = parseInt(this.getElementText(itemElement, 't:Size')) || 0;
      item.dateTimeCreated = this.getElementText(itemElement, 't:DateTimeCreated');
      item.lastModifiedTime = this.getElementText(itemElement, 't:LastModifiedTime');

      // Item-specific properties based on type
      const itemClass = this.getElementText(itemElement, 't:ItemClass');
      
      if (itemClass && itemClass.startsWith('IPM.Note')) {
        // Email message
        this.extractEmailProperties(itemElement, item);
      } else if (itemClass && itemClass.startsWith('IPM.Contact')) {
        // Contact
        this.extractContactProperties(itemElement, item);
      } else if (itemClass && itemClass.startsWith('IPM.Appointment')) {
        // Calendar appointment
        this.extractCalendarProperties(itemElement, item);
      }

      return item;
    } catch (error) {
      console.error('Failed to extract item data:', error);
      return null;
    }
  }

  /**
   * Extract email-specific properties
   */
  extractEmailProperties(itemElement, item) {
    try {
      item.from = this.extractEmailAddress(itemElement, 't:From');
      item.sender = this.extractEmailAddress(itemElement, 't:Sender');
      item.toRecipients = this.extractEmailAddresses(itemElement, 't:ToRecipients');
      item.ccRecipients = this.extractEmailAddresses(itemElement, 't:CcRecipients');
      item.bccRecipients = this.extractEmailAddresses(itemElement, 't:BccRecipients');
      item.replyTo = this.extractEmailAddresses(itemElement, 't:ReplyTo');
      
      item.isRead = this.getElementText(itemElement, 't:IsRead') === 'true';
      item.isDeliveryReceiptRequested = this.getElementText(itemElement, 't:IsDeliveryReceiptRequested') === 'true';
      item.isReadReceiptRequested = this.getElementText(itemElement, 't:IsReadReceiptRequested') === 'true';
      item.hasAttachments = this.getElementText(itemElement, 't:HasAttachments') === 'true';
      
      item.dateTimeReceived = this.getElementText(itemElement, 't:DateTimeReceived');
      item.dateTimeSent = this.getElementText(itemElement, 't:DateTimeSent');
      
      item.conversationId = this.getElementAttribute(itemElement, 't:ConversationId', 'Id');
      item.messageId = this.getElementText(itemElement, 't:InternetMessageId');
    } catch (error) {
      console.warn('Failed to extract email properties:', error);
    }
  }

  /**
   * Extract contact-specific properties
   */
  extractContactProperties(itemElement, item) {
    try {
      item.displayName = this.getElementText(itemElement, 't:DisplayName');
      item.givenName = this.getElementText(itemElement, 't:GivenName');
      item.surname = this.getElementText(itemElement, 't:Surname');
      item.middleName = this.getElementText(itemElement, 't:MiddleName');
      item.nickname = this.getElementText(itemElement, 't:Nickname');
      item.companyName = this.getElementText(itemElement, 't:CompanyName');
      item.jobTitle = this.getElementText(itemElement, 't:JobTitle');
      item.department = this.getElementText(itemElement, 't:Department');
      
      // Extract email addresses
      item.emailAddresses = this.extractContactEmailAddresses(itemElement);
      
      // Extract phone numbers
      item.phoneNumbers = this.extractContactPhoneNumbers(itemElement);
      
      // Extract physical addresses
      item.physicalAddresses = this.extractContactPhysicalAddresses(itemElement);
    } catch (error) {
      console.warn('Failed to extract contact properties:', error);
    }
  }

  /**
   * Extract calendar-specific properties
   */
  extractCalendarProperties(itemElement, item) {
    try {
      item.start = this.getElementText(itemElement, 't:Start');
      item.end = this.getElementText(itemElement, 't:End');
      item.location = this.getElementText(itemElement, 't:Location');
      item.organizer = this.extractEmailAddress(itemElement, 't:Organizer');
      item.requiredAttendees = this.extractEmailAddresses(itemElement, 't:RequiredAttendees');
      item.optionalAttendees = this.extractEmailAddresses(itemElement, 't:OptionalAttendees');
      item.resources = this.extractEmailAddresses(itemElement, 't:Resources');
      
      item.isAllDayEvent = this.getElementText(itemElement, 't:IsAllDayEvent') === 'true';
      item.legacyFreeBusyStatus = this.getElementText(itemElement, 't:LegacyFreeBusyStatus');
      item.myResponseType = this.getElementText(itemElement, 't:MyResponseType');
      
      item.recurrence = this.extractRecurrenceInfo(itemElement);
    } catch (error) {
      console.warn('Failed to extract calendar properties:', error);
    }
  }

  /**
   * Extract email address from XML element
   */
  extractEmailAddress(parentElement, tagName) {
    try {
      const addressElement = parentElement.getElementsByTagName(tagName)[0];
      if (!addressElement) return null;

      const mailboxElement = addressElement.getElementsByTagName('t:Mailbox')[0];
      if (!mailboxElement) return null;

      return {
        name: this.getElementText(mailboxElement, 't:Name'),
        email: this.getElementText(mailboxElement, 't:EmailAddress'),
        routingType: this.getElementText(mailboxElement, 't:RoutingType')
      };
    } catch (error) {
      console.warn('Failed to extract email address:', error);
      return null;
    }
  }

  /**
   * Extract multiple email addresses from XML element
   */
  extractEmailAddresses(parentElement, tagName) {
    try {
      const addresses = [];
      const containerElement = parentElement.getElementsByTagName(tagName)[0];
      
      if (!containerElement) return addresses;

      const mailboxElements = containerElement.getElementsByTagName('t:Mailbox');
      
      for (let i = 0; i < mailboxElements.length; i++) {
        const address = {
          name: this.getElementText(mailboxElements[i], 't:Name'),
          email: this.getElementText(mailboxElements[i], 't:EmailAddress'),
          routingType: this.getElementText(mailboxElements[i], 't:RoutingType')
        };
        
        if (address.email) {
          addresses.push(address);
        }
      }

      return addresses;
    } catch (error) {
      console.warn('Failed to extract email addresses:', error);
      return [];
    }
  }

  /**
   * Extract contact email addresses
   */
  extractContactEmailAddresses(itemElement) {
    try {
      const emailAddresses = {};
      const emailsElement = itemElement.getElementsByTagName('t:EmailAddresses')[0];
      
      if (emailsElement) {
        const entryElements = emailsElement.getElementsByTagName('t:Entry');
        
        for (let i = 0; i < entryElements.length; i++) {
          const entry = entryElements[i];
          const key = entry.getAttribute('Key');
          const value = entry.textContent;
          
          if (key && value) {
            emailAddresses[key] = value;
          }
        }
      }

      return emailAddresses;
    } catch (error) {
      console.warn('Failed to extract contact email addresses:', error);
      return {};
    }
  }

  /**
   * Extract contact phone numbers
   */
  extractContactPhoneNumbers(itemElement) {
    try {
      const phoneNumbers = {};
      const phonesElement = itemElement.getElementsByTagName('t:PhoneNumbers')[0];
      
      if (phonesElement) {
        const entryElements = phonesElement.getElementsByTagName('t:Entry');
        
        for (let i = 0; i < entryElements.length; i++) {
          const entry = entryElements[i];
          const key = entry.getAttribute('Key');
          const value = entry.textContent;
          
          if (key && value) {
            phoneNumbers[key] = value;
          }
        }
      }

      return phoneNumbers;
    } catch (error) {
      console.warn('Failed to extract contact phone numbers:', error);
      return {};
    }
  }

  /**
   * Extract contact physical addresses
   */
  extractContactPhysicalAddresses(itemElement) {
    try {
      const addresses = {};
      const addressesElement = itemElement.getElementsByTagName('t:PhysicalAddresses')[0];
      
      if (addressesElement) {
        const entryElements = addressesElement.getElementsByTagName('t:Entry');
        
        for (let i = 0; i < entryElements.length; i++) {
          const entry = entryElements[i];
          const key = entry.getAttribute('Key');
          
          if (key) {
            addresses[key] = {
              street: this.getElementText(entry, 't:Street'),
              city: this.getElementText(entry, 't:City'),
              state: this.getElementText(entry, 't:State'),
              countryOrRegion: this.getElementText(entry, 't:CountryOrRegion'),
              postalCode: this.getElementText(entry, 't:PostalCode')
            };
          }
        }
      }

      return addresses;
    } catch (error) {
      console.warn('Failed to extract contact physical addresses:', error);
      return {};
    }
  }

  /**
   * Extract recurrence information
   */
  extractRecurrenceInfo(itemElement) {
    try {
      const recurrenceElement = itemElement.getElementsByTagName('t:Recurrence')[0];
      if (!recurrenceElement) return null;

      const recurrence = {};

      // Extract pattern
      const patternElements = recurrenceElement.children;
      for (let i = 0; i < patternElements.length; i++) {
        const pattern = patternElements[i];
        
        if (pattern.tagName.includes('Pattern')) {
          recurrence.pattern = {
            type: pattern.tagName.replace('t:', '').replace('Pattern', ''),
            interval: parseInt(this.getElementText(pattern, 't:Interval')) || 1
          };

          // Add pattern-specific properties
          if (pattern.tagName.includes('Weekly')) {
            recurrence.pattern.daysOfWeek = this.getElementText(pattern, 't:DaysOfWeek');
          } else if (pattern.tagName.includes('Monthly')) {
            recurrence.pattern.dayOfMonth = parseInt(this.getElementText(pattern, 't:DayOfMonth'));
          }
        }
        
        if (pattern.tagName.includes('Range')) {
          recurrence.range = {
            type: pattern.tagName.replace('t:', '').replace('Range', ''),
            startDate: this.getElementText(pattern, 't:StartDate'),
            endDate: this.getElementText(pattern, 't:EndDate'),
            numberOfOccurrences: parseInt(this.getElementText(pattern, 't:NumberOfOccurrences'))
          };
        }
      }

      return recurrence;
    } catch (error) {
      console.warn('Failed to extract recurrence info:', error);
      return null;
    }
  }

  /**
   * Convert XML document to JSON object
   */
  xmlToJson(xmlDoc) {
    try {
      const result = {};
      
      if (xmlDoc.nodeType === Node.ELEMENT_NODE) {
        // Handle attributes
        if (xmlDoc.attributes.length > 0) {
          result['@attributes'] = {};
          for (let i = 0; i < xmlDoc.attributes.length; i++) {
            const attr = xmlDoc.attributes.item(i);
            result['@attributes'][attr.nodeName] = attr.nodeValue;
          }
        }
        
        // Handle child nodes
        if (xmlDoc.hasChildNodes()) {
          for (let i = 0; i < xmlDoc.childNodes.length; i++) {
            const child = xmlDoc.childNodes.item(i);
            const nodeName = child.nodeName;
            
            if (child.nodeType === Node.TEXT_NODE) {
              const text = child.nodeValue.trim();
              if (text) {
                result['#text'] = text;
              }
            } else if (child.nodeType === Node.ELEMENT_NODE) {
              if (result[nodeName]) {
                if (!Array.isArray(result[nodeName])) {
                  result[nodeName] = [result[nodeName]];
                }
                result[nodeName].push(this.xmlToJson(child));
              } else {
                result[nodeName] = this.xmlToJson(child);
              }
            }
          }
        }
      }
      
      return result;
    } catch (error) {
      console.error('Failed to convert XML to JSON:', error);
      return {};
    }
  }

  /**
   * Validate XML structure
   */
  validateXML(xmlString, expectedRootElement = null) {
    try {
      const doc = this.parseXML(xmlString);
      
      if (expectedRootElement) {
        const rootElement = doc.documentElement;
        if (rootElement.tagName !== expectedRootElement) {
          return {
            isValid: false,
            error: `Expected root element '${expectedRootElement}', found '${rootElement.tagName}'`
          };
        }
      }

      return { isValid: true };
    } catch (error) {
      return {
        isValid: false,
        error: error.message
      };
    }
  }

  /**
   * Pretty print XML
   */
  prettyPrintXML(xmlString) {
    try {
      const doc = this.parseXML(xmlString);
      const serializer = new XMLSerializer();
      const formatted = serializer.serializeToString(doc);
      
      // Basic formatting (this could be enhanced with proper indentation)
      return formatted
        .replace(/></g, '>\n<')
        .replace(/^\s*\n/gm, '');
    } catch (error) {
      console.error('Failed to pretty print XML:', error);
      return xmlString;
    }
  }
}
