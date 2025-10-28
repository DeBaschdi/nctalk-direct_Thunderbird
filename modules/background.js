/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
'use strict';
/**
 * Hintergrundskript fuer Nextcloud Talk Direkt.
 * Verantwortlich fuer API-Aufrufe (Talk + CardDAV), Caching und
 * Utilities, die der Frontend-Teil per Messaging abruft.
 */
let DEBUG_ENABLED = false;
(async () => {
  try{
    const stored = await browser.storage.local.get("debugEnabled");
    DEBUG_ENABLED = !!stored.debugEnabled;
  }catch(_){ }
})();
browser.storage.onChanged.addListener((changes, area) => {
  if (area !== "local") return;
  if (Object.prototype.hasOwnProperty.call(changes, "debugEnabled")){
    DEBUG_ENABLED = !!changes.debugEnabled.newValue;
  }
});
function L(...a){
  if (!DEBUG_ENABLED) return;
  try{
    console.log("[NCBG]", ...a);
  }catch(_){ }
}

function shortToken(token, { keepStart = 4, keepEnd = 3 } = {}){
  if (!token) return "";
  const str = String(token);
  if (str.length <= keepStart + keepEnd + 3){
    return str;
  }
  return str.slice(0, keepStart) + "..." + str.slice(str.length - keepEnd);
}

function shortId(value, max = 12){
  if (value == null) return "";
  const str = String(value);
  if (str.length <= max){
    return str;
  }
  return str.slice(0, max) + "...";
}

function normalizeBaseUrl(input){
  if (!input) return "";
  return String(input).trim().replace(/\/+$/, "");
}

const BG_I18N_FALLBACKS = {
  error_canvas_support_missing: "Canvas-Unterstuetzung fehlt.",
  error_avatar_data_missing: "Avatar-Daten fehlen.",
  error_canvas_context_missing: "Canvas-Kontext nicht verfuegbar.",
  error_image_decode_failed: "Bild konnte nicht dekodiert werden.",
  error_image_size_unknown: "Bildgroesse konnte nicht ermittelt werden.",
  error_image_load_failed: "Bild konnte nicht geladen werden.",
  error_pixel_data_missing: "Pixel-Daten nicht verfuegbar.",
  error_credentials_missing: "Nextcloud Zugangsdaten fehlen (Add-on-Optionen).",
  error_system_addressbook_failed: "System-Adressbuch konnte nicht geladen werden: $1",
  error_lobby_set_failed: "Lobby konnte nicht gesetzt werden: $1",
  error_lobby_set_failed_short: "Lobby konnte nicht gesetzt werden.",
  error_room_create_failed: "Raum konnte nicht erstellt werden.",
  error_token_missing_in_response: "Kein Token im OCS-Response.",
  error_room_token_missing: "Raum-Token fehlt.",
  error_lobby_no_permission: "Keine Berechtigung zum Aendern der Lobby (HTTP 403).",
  error_lobby_update_failed: "Lobby-Update fehlgeschlagen (HTTP $1).",
  error_ocs: "OCS-Fehler: $1",
  error_room_delete_failed: "Talk-Raum konnte nicht geloescht werden: $1",
  error_token_or_actor_missing: "Raum-Token oder Teilnehmer-ID fehlt.",
  error_participant_add_failed: "Teilnehmer konnte nicht hinzugefuegt werden: $1",
  error_moderator_id_missing: "Moderator-ID fehlt.",
  error_moderator_set_failed: "Moderator konnte nicht gesetzt werden: $1",
  error_leave_failed: "Verlassen des Raums fehlgeschlagen: $1",
  error_delegation_data_missing: "Delegationsdaten fehlen.",
  error_moderator_target_missing: "Moderator-ID ist leer.",
  error_participant_not_found: "Teilnehmer wurde nicht gefunden, bitte Schreibweise pruefen.",
  error_users_load_failed: "Benutzer konnten nicht geladen werden.",
  error_unknown_utility_request: "Unbekannte Utility-Anfrage."
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

function bgI18n(key, substitutions = []){
  try{
    if (browser && browser.i18n){
      const message = browser.i18n.getMessage(key, substitutions);
      if (message){
        return message;
      }
    }
  }catch(_){ }
  const fallback = BG_I18N_FALLBACKS[key];
  if (fallback){
    return applySubstitutions(fallback, Array.isArray(substitutions) ? substitutions : [substitutions]);
  }
  if (Array.isArray(substitutions) && substitutions.length){
    return String(substitutions[0]);
  }
  return "";
}

/**
 * Liefert einen lokalisierten Error anhand der README-Fallback-Tabelle.
 * @param {string} key - Übersetzungsschlüssel
 * @param {Array} substitutions - Platzhalterwerte
 */
function localizedError(key, substitutions = []){
  const message = bgI18n(key, substitutions);
  return new Error(message || key);
}

/**
 * Dekodiert Basis64-Avatar-Daten in ein Plain-Array für das Frontend.
 * Wird sowohl für Adressbuch-Avatare als auch Moderatoren genutzt.
 */
async function decodeAvatarPixels({ base64, mime } = {}){
  const clean = String(base64 || "").replace(/\s+/g, "");
  if (!clean){
    throw localizedError("error_avatar_data_missing");
  }
  const binary = atob(clean);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++){
    bytes[i] = binary.charCodeAt(i) & 0xff;
  }
  const blob = new Blob([bytes], { type: mime || "image/png" });
  const hidden = getHiddenWindow();
  let canvas = null;
  let ctx = null;
  if (typeof createImageBitmap === "function"){
    try{
      const bitmap = await createImageBitmap(blob);
      canvas = createScratchCanvas(bitmap.width || 1, bitmap.height || 1);
      ctx = canvas.getContext("2d");
      if (!ctx){
        throw localizedError("error_canvas_context_missing");
      }
      ctx.drawImage(bitmap, 0, 0);
      if (typeof bitmap.close === "function"){
        try { bitmap.close(); } catch (_){ }
      }
    }catch(e){
      canvas = null;
      ctx = null;
    }
  }
  if (!canvas){
    if (!hidden || !hidden.document || typeof hidden.document.createElement !== "function"){
      throw localizedError("error_image_decode_failed");
    }
    const img = hidden.document.createElement("img");
    img.src = "data:" + (mime || "image/png") + ";base64," + clean;
    await new Promise((resolve, reject) => {
      img.onload = () => resolve();
      img.onerror = (event) => reject(event || localizedError("error_image_load_failed"));
    });
    const width = img.naturalWidth || img.width || 0;
    const height = img.naturalHeight || img.height || 0;
    if (!width || !height){
      throw localizedError("error_image_size_unknown");
    }
    canvas = createScratchCanvas(width, height);
    ctx = canvas.getContext("2d");
    if (!ctx){
      throw localizedError("error_canvas_context_missing");
    }
    ctx.drawImage(img, 0, 0);
  }
  const finalCtx = ctx || canvas.getContext("2d");
  if (!finalCtx){
    throw localizedError("error_canvas_context_missing");
  }
  const width = canvas.width || 0;
  const height = canvas.height || 0;
  const imageData = finalCtx.getImageData(0, 0, width, height);
  const sourcePixels = imageData && imageData.data ? imageData.data : null;
  if (!sourcePixels || typeof sourcePixels.length !== "number"){
    throw localizedError("error_pixel_data_missing");
  }
  const plain = Array.from(sourcePixels);
  const byteLength = sourcePixels.byteLength || plain.length;
  return {
    width,
    height,
    pixels: plain,
    byteLength
  };
}

/**
 * Zentraler Background-Listener für experimentelle Schnittstellen und Optionen.
 * Bedient Create-/Lobby-/Utility-Requests, die aus dem Experiment kommen.
 */
browser.runtime.onMessage.addListener(async (msg) => {
  if (!msg || !msg.type) return;
  L("msg", msg.type, { hasPayload: !!msg.payload });
  if (msg.type === "talkMenu:newPublicSubmit"){
    try {
      const out = await createTalkPublicRoom(msg.payload);
      return { ok:true, url: out.url, token: out.token, fallback: !!out.fallback, reason: out.reason };
    } catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:getConfig"){
    try{
      const config = await getOpts();
      return { ok:true, config };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:decodeAvatar"){
    try{
      const payload = msg.payload || msg;
      const result = await decodeAvatarPixels({ base64: payload?.base64, mime: payload?.mime });
      return {
        ok:true,
        width: result.width,
        height: result.height,
        byteLength: result.byteLength,
        pixels: result.pixels
      };
    }catch(e){
      L("decodeAvatar runtime error", e?.message || String(e));
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:updateLobby"){
    try{
      await updateTalkLobby(msg.payload);
      return { ok:true };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:delegateModerator"){
    try{
      const result = await delegateRoomModerator(msg.payload || {});
      return { ok:true, result };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:searchUsers"){
    try{
      const users = await searchSystemAddressbook(msg.payload || {});
      return { ok:true, users };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "options:testConnection"){
    try{
      const result = await probeNextcloudCredentials(msg.payload || {});
      if (result.ok){
        return { ok:true, message: result.message || "", version: result.version || "" };
      }
      return { ok:false, error: result.message || bgI18n("error_credentials_missing"), code: result.code || "" };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
});

browser.calToolbar.onCreateRequest.addListener(async (payload = {}) => {
  L("calToolbar.onCreateRequest", {
    title: payload?.title || "",
    hasPassword: !!payload?.password,
    enableLobby: !!payload?.enableLobby,
    enableListable: !!payload?.enableListable,
    eventConversation: !!payload?.eventConversation
  });
  try {
    const out = await createTalkPublicRoom(payload);
    return { ok:true, url: out.url, token: out.token, fallback: !!out.fallback, reason: out.reason };
  } catch(e){
    return { ok:false, error: e?.message || String(e) };
  }
});

browser.calToolbar.onLobbyUpdate.addListener(async (payload = {}) => {
  L("calToolbar.onLobbyUpdate", {
    token: payload?.token ? shortToken(payload.token) : "",
    enableLobby: !!payload?.enableLobby,
    startTimestamp: payload?.startTimestamp || null
  });
  try{
    await updateTalkLobby(payload);
    return { ok:true };
  }catch(e){
    return { ok:false, error: e?.message || String(e) };
  }
});

browser.calToolbar.onUtilityRequest.addListener(async (payload = {}) => {
  L("calToolbar.onUtilityRequest", {
    type: payload?.type || "",
    token: payload?.token ? shortToken(payload.token) : "",
    searchTerm: payload?.searchTerm || "",
    delegate: payload?.newModerator || ""
  });
  try{
    const type = payload?.type;
    if (type === "getConfig"){
      const config = await getOpts();
      return { ok:true, config };
    }
    if (type === "decodeAvatar"){
      try{
        const result = await decodeAvatarPixels({ base64: payload?.base64, mime: payload?.mime });
        return {
          ok:true,
          width: result.width,
          height: result.height,
          byteLength: result.byteLength,
          pixels: result.pixels
        };
      }catch(err){
        L("decodeAvatar utility error", err?.message || String(err));
        throw err;
      }
    }
    if (type === "supportsEventConversation"){
      const info = await getEventConversationSupport();
      return { ok:true, supported: info.supported, reason: info.reason || "" };
    }
    if (type === "delegateModerator"){
      const result = await delegateRoomModerator(payload || {});
      return { ok:true, result };
    }
    if (type === "deleteRoom"){
      await deleteTalkRoom(payload || {});
      return { ok:true };
    }
    if (type === "searchUsers"){
      const users = await searchSystemAddressbook(payload || {});
      return { ok:true, users };
    }
    return { ok:false, error: bgI18n("error_unknown_utility_request") || "Unbekannte Utility-Anfrage." };
  }catch(e){
    return { ok:false, error: e?.message || String(e) };
  }
});

// *** IMPORTANT: initialize experiment on startup ***
(async () => {
  try{
    const defaultLabel = browser?.i18n?.getMessage("ui_insert_button_label") || "Talk-Link einfuegen";
    const defaultTooltip = browser?.i18n?.getMessage("ui_toolbar_tooltip") || "Nextcloud Talk";
    const ok = await browser.calToolbar.init({ label: defaultLabel, tooltip: defaultTooltip });
    if (!ok){
      console.warn("[NCBG] experiment init returned falsy value");
    }
  }catch(e){
    console.error("[NCBG] init error", e);
  }
})();























