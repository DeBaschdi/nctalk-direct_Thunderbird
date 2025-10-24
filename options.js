'use strict';
/**
 * Settings page for Nextcloud Talk Direct.
 */
const OPTIONS_FALLBACKS = {
  options_title: "Nextcloud Talk - Einstellungen",
  options_heading: "Nextcloud Talk - Einstellungen",
  options_base_url_label: "Nextcloud URL",
  options_base_url_placeholder: "https://cloud.example.com",
  options_user_label: "Benutzername",
  options_app_pass_label: "App-Passwort",
  options_debug_label: "Debug-Logging",
  options_debug_hint: "Konsole f\u00fcr Diagnosemeldungen verwenden",
  options_support_label: "Unterst\u00fctzen",
  options_support_link: "PayPal: paypal.me/debaschdi",
  options_save_button: "Speichern",
  options_status_saved: "Gespeichert.",
  options_status_save_failed: "Speichern fehlgeschlagen.",
  options_status_load_failed: "Einstellungen konnten nicht geladen werden.",
  options_test_button: "Verbindung testen",
  options_test_running: "Pr\u00fcfe Verbindung...",
  options_test_success: "Verbindung erfolgreich.",
  options_test_failed: "Verbindung fehlgeschlagen.",
  options_test_failed_auth: "Benutzername/App-Passwort ung\u00fcltig.",
  options_test_missing: "Bitte zuerst URL, Benutzername und App-Passwort eingeben."
};

function applySubstitutions(template, substitutions){
  if (!template) return "";
  if (!Array.isArray(substitutions) || substitutions.length === 0){
    return template;
  }
  return template.replace(/\$(\d+)/g, (match, index) => {
    const idx = Number(index) - 1;
    return idx >= 0 && idx < substitutions.length && substitutions[idx] != null
      ? String(substitutions[idx])
      : "";
  });
}

function i18n(key, substitutions = []){
  try{
    if (browser && browser.i18n){
      const message = browser.i18n.getMessage(key, substitutions);
      if (message){
        return message;
      }
    }
  }catch(_){}
  const fallback = OPTIONS_FALLBACKS[key];
  if (fallback){
    return applySubstitutions(fallback, Array.isArray(substitutions) ? substitutions : [substitutions]);
  }
  if (Array.isArray(substitutions) && substitutions.length){
    return String(substitutions[0]);
  }
  return "";
}

function translatePage(){
  const textNodes = document.querySelectorAll("[data-i18n]");
  textNodes.forEach((el) => {
    const key = el.dataset.i18n;
    if (!key) return;
    const message = i18n(key);
    if (message) el.textContent = message;
  });

  const placeholderNodes = document.querySelectorAll("[data-i18n-placeholder]");
  placeholderNodes.forEach((el) => {
    const key = el.dataset.i18nPlaceholder;
    if (!key) return;
    const message = i18n(key);
    if (message) el.setAttribute("placeholder", message);
  });

  const titleMessage = i18n("options_title");
  if (titleMessage) document.title = titleMessage;
}

translatePage();

const statusEl = document.getElementById("status");
let statusTimer = null;

function showStatus(message, isError = false){
  if (statusTimer){
    clearTimeout(statusTimer);
    statusTimer = null;
  }
  statusEl.textContent = message || "";
  statusEl.style.color = isError ? "#b00020" : "";
  if (message && !isError){
    statusTimer = setTimeout(() => {
      statusEl.textContent = "";
      statusTimer = null;
    }, 2000);
  }
}

async function load(){
  const stored = await browser.storage.local.get(["baseUrl","user","appPass","debugEnabled"]);
  if (stored.baseUrl) document.getElementById("baseUrl").value = stored.baseUrl;
  if (stored.user) document.getElementById("user").value = stored.user;
  if (stored.appPass) document.getElementById("appPass").value = stored.appPass;
  document.getElementById("debugEnabled").checked = !!stored.debugEnabled;
}

async function save(){
  const baseUrl = document.getElementById("baseUrl").value.trim();
  const user = document.getElementById("user").value.trim();
  const appPass = document.getElementById("appPass").value;
  const debugEnabled = document.getElementById("debugEnabled").checked;
  await browser.storage.local.set({ baseUrl, user, appPass, debugEnabled });
  showStatus(i18n("options_status_saved"));
}

document.getElementById("save").addEventListener("click", async () => {
  try{
    await save();
  }catch(e){
    console.error(e);
    showStatus(e?.message || i18n("options_status_save_failed"), true);
  }
});

const testButton = document.getElementById("testConnection");
if (testButton){
  testButton.addEventListener("click", async () => {
    const button = testButton;
    if (button.disabled) return;
    const baseUrl = document.getElementById("baseUrl").value.trim();
    const user = document.getElementById("user").value.trim();
    const appPass = document.getElementById("appPass").value;
    if (!baseUrl || !user || !appPass){
      showStatus(i18n("options_test_missing"), true);
      return;
    }
    button.disabled = true;
    const originalLabel = button.textContent;
    button.textContent = i18n("options_test_running");
    try{
      const response = await browser.runtime.sendMessage({
        type: "options:testConnection",
        payload: { baseUrl, user, appPass }
      });
      if (response?.ok){
        const message = response?.message ? String(response.message) : i18n("options_test_success");
        showStatus(message, false);
      }else{
        const code = response?.code;
        const fallbackKey = code === "auth" ? "options_test_failed_auth" : "options_test_failed";
        const message = response?.error || i18n(fallbackKey);
        showStatus(message, true);
      }
    }catch(err){
      console.error(err);
      showStatus(err?.message || i18n("options_test_failed"), true);
    }finally{
      button.disabled = false;
      button.textContent = originalLabel || i18n("options_test_button");
    }
  });
}

load().catch((e) => {
  console.error(e);
  showStatus(e?.message || i18n("options_status_load_failed"), true);
});
