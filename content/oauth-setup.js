// oauth-setup.js

document.addEventListener("DOMContentLoaded", async () => {
  // 1) Carichiamo eventuali credenziali già salvate da browser.storage.local
  try {
    let result = await browser.storage.local.get("oauthConfig");
    if (result.oauthConfig) {
      document.getElementById("client-id").value = result.oauthConfig.clientId || "";
      document.getElementById("tenant-id").value = result.oauthConfig.tenantId || "";
      document.getElementById("client-secret").value = result.oauthConfig.clientSecret || "";
    }
  } catch (err) {
    console.error("Errore nel leggere oauthConfig da storage:", err);
  }

  // 2) Gestiamo il submit del form
  document.getElementById("oauth-form").addEventListener("submit", async (e) => {
    e.preventDefault();

    let clientId = document.getElementById("client-id").value.trim();
    let tenantId = document.getElementById("tenant-id").value.trim();
    let clientSecret = document.getElementById("client-secret").value.trim();
    let messageEl = document.getElementById("message");

    if (!clientId || !tenantId || !clientSecret) {
      messageEl.style.color = "red";
      messageEl.textContent = "Tutti i campi sono obbligatori.";
      return;
    }

    let cfg = {
      clientId,
      tenantId,
      clientSecret
    };

    try {
      await browser.storage.local.set({ oauthConfig: cfg });
      messageEl.style.color = "green";
      messageEl.textContent = "Credenziali salvate correttamente!";
      console.log("Credenziali OAuth2 salvate:", cfg);
    } catch (err) {
      console.error("Errore nel salvare oauthConfig su storage:", err);
      messageEl.style.color = "red";
      messageEl.textContent = "Errore nel salvataggio delle credenziali.";
    }
  });
});

