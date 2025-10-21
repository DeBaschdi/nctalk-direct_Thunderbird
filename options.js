'use strict';
/**
 * Einstellungsseite – verwaltet Nextcloud-Zugangsdaten.
 */
const statusEl = document.getElementById("status");
let statusTimer = null;

/**
 * Zeigt eine Statusmeldung über dem Speichern-Button an.
 * @param {string} message
 * @param {boolean} [isError=false]
 */
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

/**
 * Lädt gespeicherte Anmeldedaten und füllt die Felder.
 */
async function load(){
  const stored = await browser.storage.local.get(["baseUrl","user","appPass"]);
  if (stored.baseUrl) document.getElementById("baseUrl").value = stored.baseUrl;
  if (stored.user) document.getElementById("user").value = stored.user;
  if (stored.appPass) document.getElementById("appPass").value = stored.appPass;
}

/**
 * Persistiert die Eingaben und meldet Erfolg im UI.
 */
async function save(){
  const baseUrl = document.getElementById("baseUrl").value.trim();
  const user = document.getElementById("user").value.trim();
  const appPass = document.getElementById("appPass").value;
  await browser.storage.local.set({ baseUrl, user, appPass });
  showStatus("Gespeichert.");
}

document.getElementById("save").addEventListener("click", async () => {
  try{
    await save();
  }catch(e){
    console.error(e);
    showStatus(e?.message || "Speichern fehlgeschlagen.", true);
  }
});

load().catch((e) => {
  console.error(e);
  showStatus(e?.message || "Einstellungen konnten nicht geladen werden.", true);
});
