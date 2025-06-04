// background.js

import { fetchInboxMessages } from "./services/ews-client.js";

function log(...args) {
  console.log("Syncbird:", ...args);
}

/**
 * All’avvio del background script, logghiamo quali account Thunderbird sono presenti.
 */
(async () => {
  log("Initializing background script");
  try {
    let accounts = await browser.accounts.list();
    log("Loaded Thunderbird accounts:", accounts);
  } catch (err) {
    log("Error listing accounts:", err);
  }
  log("Background script initialized");
})();

/**
 * Se un account nuovo viene aggiunto (opzionale)
 */
browser.accounts.onCreated.addListener((account) => {
  log("New account created:", account);
});

/**
 * Listener per i nuovi messaggi in arrivo.
 * Quando Thunderbird riceve una mail, chiamiamo fetchInboxMessages() per verificare che le credenziali Basic Auth funzionino.
 */
browser.messages.onNewMailReceived.addListener(async (folder, messages) => {
  try {
    log("New mail received in folder:", folder.name, "Message IDs:", messages);
    let items = await fetchInboxMessages();
    log("Fetched items from Inbox (via EWS):", items);
  } catch (err) {
    log("Error in onNewMailReceived (EWS Basic Auth):", err);
  }
});

