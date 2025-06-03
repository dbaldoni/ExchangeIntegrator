// content/add-account.js

document.getElementById("createBtn").addEventListener("click", async () => {
  const statusEl = document.getElementById("status");
  statusEl.textContent = "";
  statusEl.className = "";

  const name = document.getElementById("name").value.trim();
  const email = document.getElementById("email").value.trim();
  const username = document.getElementById("username").value.trim();
  const password = document.getElementById("password").value;
  const protocol = document.getElementById("protocol").value;
  const webmail = document.getElementById("webmail").value.trim();

  // Validazioni base
  if (!name || !email || !username || !password || !protocol || !webmail) {
    statusEl.textContent = "All fields are required.";
    statusEl.className = "error";
    return;
  }

  // Costruzione dell’oggetto di configurazione per browser.accounts.create
  let newAccountConfig = {
    name: name,
    // L’indirizzo email è obbligatorio
    identities: [{
      id: "id1",
      name: name,
      email: email
    }]
  };

  if (protocol === "EWS") {
    // Configurazione EWS (Exchange Web Services)
    // Thunderbird usa la chiave `incoming` con `type: "ews"`
    newAccountConfig.incoming = {
      type: "ews",
      hostname: webmail.replace(/\\/g, ""), // es. "outlook.office365.com"
      username: username,
      auth: "password",
      password: password,
      port: 443,
      socketType: 3 // 3 = SSL/TLS, 2 = STARTTLS, 1 = None
    };
    // Outgoing (SMTP) si può generare in automatico usando lo stesso host
    newAccountConfig.outgoing = {
      type: "smtp",
      hostname: "smtp.office365.com",
      username: username,
      auth: "password",
      password: password,
      port: 587,
      socketType: 2 // STARTTLS
    };
  } else if (protocol === "OWA") {
    // Configurazione “Outlook Web Access” – in Thunderbird si traduce in protocollo “owaservice”
    newAccountConfig.incoming = {
      type: "owaservice",
      hostname: webmail.replace(/\\/g, ""),
      username: username,
      auth: "password",
      password: password,
      port: 443,
      socketType: 3
    };
    newAccountConfig.outgoing = {
      type: "smtp",
      hostname: "smtp.office365.com",
      username: username,
      auth: "password",
      password: password,
      port: 587,
      socketType: 2
    };
  } else if (protocol === "EAS") {
    // ActiveSync (sperimentale)
    newAccountConfig.incoming = {
      type: "ews", // Thunderbird non ha tipo “eas” WebExtension, quindi usiamo EWS
      hostname: webmail.replace(/\\/g, ""),
      username: username,
      auth: "password",
      password: password,
      port: 443,
      socketType: 3
    };
    newAccountConfig.outgoing = {
      type: "smtp",
      hostname: "smtp.office365.com",
      username: username,
      auth: "password",
      password: password,
      port: 587,
      socketType: 2
    };
  }

  // Proviamo a creare l’account
  try {
    await browser.accounts.create(newAccountConfig);
    statusEl.textContent = "Account created successfully!";
    statusEl.className = "success";
  } catch (err) {
    statusEl.textContent = "Error creating account: " + err.message;
    statusEl.className = "error";
  }
});

