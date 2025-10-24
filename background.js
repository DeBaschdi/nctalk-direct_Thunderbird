'use strict';
/**
 * Hintergrundskript für Nextcloud Talk Direkt.
 * Verantwortlich für API-Aufrufe (Talk + CardDAV), Caching und
 * Utilities, die der Frontend-Teil per Messaging abruft.
 */
let DEBUG_ENABLED = false;
(async () => {
  try{
    const stored = await browser.storage.local.get("debugEnabled");
    DEBUG_ENABLED = !!stored.debugEnabled;
  }catch(_){}
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
  }catch(_){}
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
  return String(input).trim().replace(/\/+$/,"");
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
  error_room_details_missing: "Raumdetails fehlen im OCS-Response.",
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
  }catch(_){}
  const fallback = BG_I18N_FALLBACKS[key];
  if (fallback){
    return applySubstitutions(fallback, Array.isArray(substitutions) ? substitutions : [substitutions]);
  }
  if (Array.isArray(substitutions) && substitutions.length){
    return String(substitutions[0]);
  }
  return "";
}

function localizedError(key, substitutions = []){
  const message = bgI18n(key, substitutions);
  return new Error(message || key);
}

/**
 * Lädt die in storage abgelegten Verbindungsdaten und bereitet sie auf.
 * @returns {Promise<{baseUrl:string,user:string,appPass:string}>}
 */
async function getOpts(){
  const stored = await browser.storage.local.get([
    "baseUrl",
    "user",
    "appPass"
  ]);
  return {
    baseUrl: stored.baseUrl ? String(stored.baseUrl).replace(/\/+$/,"") : "",
    user: stored.user || "",
    appPass: stored.appPass || ""
  };
}

const EVENT_SUPPORT_CACHE = {
  value: null,
  reason: "",
  expires: 0
};
const EVENT_SUPPORT_TTL = 5 * 60 * 1000;

function noteEventSupport(value, reason){
  EVENT_SUPPORT_CACHE.value = value;
  EVENT_SUPPORT_CACHE.reason = reason || "";
  EVENT_SUPPORT_CACHE.expires = Date.now() + EVENT_SUPPORT_TTL;
}

function markEventSupportUnsupported(reason){
  noteEventSupport(false, reason || "");
}

function parseEventSupportFlag(data){
  if (!data) return { status:null, hint:"" };
  const featureSources = [];
  if (Array.isArray(data.features)) featureSources.push(...data.features);
  if (Array.isArray(data.optionalFeatures)) featureSources.push(...data.optionalFeatures);
  if (Array.isArray(data.localFeatures)) featureSources.push(...data.localFeatures);
  const normalizedFeatures = featureSources.map((feature) => String(feature || "").toLowerCase());
  const eventFeatureTokens = [
    "event-conversation",
    "event-conversations",
    "conversation-object",
    "conversation-objects",
    "conversation-object-bind",
    "dashboard-event-rooms",
    "mutual-calendar-events",
    "unbind-conversation"
  ];
  for (const token of normalizedFeatures){
    const match = eventFeatureTokens.find((needle) => token.includes(needle));
    if (match) return { status:true, hint:"Feature '" + match + "'" };
    if (token.includes("event") && token.includes("conversation")){
      return { status:true, hint:"Feature '" + token + "'" };
    }
  }
  const flagCandidates = [
    ["eventConversation", data.eventConversation],
    ["eventConversations", data.eventConversations],
    ["supportsEventConversation", data.supportsEventConversation],
    ["supportsEventConversations", data.supportsEventConversations],
    ["conversationObject", data.conversationObject],
    ["supportsConversationObjects", data.supportsConversationObjects]
  ];
  for (const [name, entry] of flagCandidates){
    if (entry === true) return { status:true, hint:"Flag '" + name + "'" };
    if (entry === false) return { status:false, hint:"Flag '" + name + "'" };
  }
  const convoConfig = data.config || data.conversations || data.configurations || {};
  const retention = (convoConfig.conversations && (convoConfig.conversations["retention-event"] ?? convoConfig.conversations.retentionEvent))
    ?? convoConfig["retention-event"]
    ?? convoConfig.retentionEvent;
  if (retention !== undefined) return { status:true, hint:"Config 'retention-event'" };
  return { status:null, hint:"" };
}

function parseMajorVersion(value){
  if (value == null) return null;
  if (typeof value === "number"){
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === "string"){
    const match = /^(\d+)/.exec(value.trim());
    if (match && match[1]) return parseInt(match[1], 10);
  }
  if (typeof value === "object"){
    if (typeof value.major === "number") return Number.isFinite(value.major) ? value.major : null;
    if (typeof value.major === "string"){
      const parsed = parseInt(value.major, 10);
      if (!Number.isNaN(parsed)) return parsed;
    }
    if (typeof value.string === "string"){
      const match = /^(\d+)/.exec(value.string.trim());
      if (match && match[1]) return parseInt(match[1], 10);
    }
  }
  return null;
}

async function requestTalkCapabilities(url, headers){
  try{
    L("request talk capabilities", { url });
    const res = await fetch(url, { method:"GET", headers });
    L("talk capabilities status", { status: res.status, ok: res.ok });
    const raw = await res.text();
    let data = null;
    try{ data = raw ? JSON.parse(raw) : null; }catch(_){}
    if (res.status === 404){
      return { supported:null, reason:"Talk Capabilities-Endpunkt liefert HTTP 404." };
    }
    if (!res.ok){
      const meta = data?.ocs?.meta || {};
      const detailParts = [];
      if (meta.message && meta.message !== meta.status) detailParts.push(meta.message);
      if (meta.status && meta.status !== meta.statuscode) detailParts.push(meta.status);
      if (meta.statuscode) detailParts.push("HTTP " + meta.statuscode);
      if (res.status) detailParts.push("HTTP " + res.status + " " + res.statusText);
      const detail = detailParts.filter(Boolean).join(" / ") || raw || ("HTTP " + res.status + " " + res.statusText);
      return { supported:null, reason:"Talk Capabilities fehlgeschlagen: " + detail };
    }
    const spreedCaps = data?.ocs?.data?.spreed ?? data?.ocs?.data ?? data?.spreed ?? null;
    const parsed = parseEventSupportFlag(spreedCaps);
    if (parsed.status === true){
      return { supported:true, reason:"Talk Capabilities: " + parsed.hint };
    }
    if (parsed.status === false){
      return { supported:false, reason:"Talk Capabilities: " + parsed.hint + " => Event nicht verfuegbar." };
    }
    return { supported:null, reason: parsed.hint ? "Talk Capabilities: " + parsed.hint : "Talk Capabilities ohne Event-Flag." };
  }catch(e){
    return { supported:null, reason: e?.message || "Talk Capabilities nicht erreichbar." };
  }
}

async function requestCoreCapabilities(baseUrl, headers){
  const coreUrl = baseUrl + "/ocs/v2.php/cloud/capabilities";
  try{
    L("request core capabilities", { url: coreUrl });
    const res = await fetch(coreUrl, { method:"GET", headers });
    L("core capabilities status", { status: res.status, ok: res.ok });
    const raw = await res.text();
    let data = null;
    try{ data = raw ? JSON.parse(raw) : null; }catch(_){}
    if (!res.ok){
      const meta = data?.ocs?.meta || {};
      const detailParts = [];
      if (meta.message && meta.message !== meta.status) detailParts.push(meta.message);
      if (meta.status && meta.status !== meta.statuscode) detailParts.push(meta.status);
      if (meta.statuscode) detailParts.push("HTTP " + meta.statuscode);
      if (res.status) detailParts.push("HTTP " + res.status + " " + res.statusText);
      const detail = detailParts.filter(Boolean).join(" / ") || raw || ("HTTP " + res.status + " " + res.statusText);
      return { supported:null, reason:"Cloud Capabilities fehlgeschlagen: " + detail };
    }
    const capabilities = data?.ocs?.data?.capabilities || {};
    const spreedCaps = capabilities.spreed ?? data?.ocs?.data?.spreed ?? null;
    const parsed = parseEventSupportFlag(spreedCaps);
    if (parsed.status === true){
      return { supported:true, reason:"Cloud Capabilities: " + parsed.hint };
    }
    if (parsed.status === false){
      return { supported:false, reason:"Cloud Capabilities: " + parsed.hint + " => Event nicht verfuegbar." };
    }
    const versionMajor =
      parseMajorVersion(spreedCaps?.version) ??
      parseMajorVersion(capabilities?.spreed?.version) ??
      parseMajorVersion(data?.ocs?.data?.version) ??
      parseMajorVersion(data?.ocs?.data?.installed?.version) ??
      parseMajorVersion(data?.ocs?.data?.system?.version);
    if (versionMajor !== null && versionMajor < 32){
      return { supported:false, reason:"Cloud Capabilities: Nextcloud-Version " + versionMajor + " (<32) => Event deaktiviert." };
    }
    if (versionMajor !== null && versionMajor >= 32){
      return { supported:null, reason:"Cloud Capabilities: Nextcloud-Version " + versionMajor + " meldet kein Event-Flag." };
    }
    return { supported:null, reason:"Cloud Capabilities ohne Event-Angaben." };
  }catch(e){
    return { supported:null, reason: e?.message || "Cloud Capabilities nicht erreichbar." };
  }
}

async function getEventConversationSupport(){
  const now = Date.now();
  if (EVENT_SUPPORT_CACHE.expires > now && EVENT_SUPPORT_CACHE.value !== null){
    L("event support cache hit", {
      supported: EVENT_SUPPORT_CACHE.value,
      reason: EVENT_SUPPORT_CACHE.reason || "",
      expiresInMs: Math.max(0, EVENT_SUPPORT_CACHE.expires - now)
    });
    return { supported: EVENT_SUPPORT_CACHE.value, reason: EVENT_SUPPORT_CACHE.reason };
  }
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass){
    L("event support aborted", "credentials missing");
    noteEventSupport(false, "Zugangsdaten fehlen");
    return { supported:false, reason:"Zugangsdaten fehlen" };
  }
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": "Basic " + btoa(user + ":" + appPass),
    "Accept": "application/json"
  };
  const talkUrl = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/capabilities";
  const talkResult = await requestTalkCapabilities(talkUrl, headers);
  if (talkResult.supported === true){
    if (talkResult.reason){
      L("event capability (talk)", talkResult.reason);
    } else {
      L("event capability (talk)", "Event-Unterstuetzung bestaetigt (Talk Capabilities).");
    }
    noteEventSupport(true, talkResult.reason || "");
    return { supported:true, reason: talkResult.reason || "" };
  }
  if (talkResult.reason){
    L("event capability (talk)", talkResult.reason);
  }
  const reasons = [];
  if (talkResult.reason) reasons.push(talkResult.reason);
  if (talkResult.supported === false){
    const reason = reasons.filter(Boolean).join(" | ") || "";
    noteEventSupport(false, reason);
    return { supported:false, reason };
  }
  const coreResult = await requestCoreCapabilities(baseUrl, headers);
  if (coreResult.supported === true){
    const reason = coreResult.reason || "";
    if (reason){
      L("event capability (core)", reason);
    } else {
      L("event capability (core)", "Event-Unterstuetzung bestaetigt (Cloud Capabilities).");
    }
    noteEventSupport(true, reason);
    return { supported:true, reason };
  }
  if (coreResult.reason){
    L("event capability (core)", coreResult.reason);
  }
  if (coreResult.reason) reasons.push(coreResult.reason);
  if (coreResult.supported === false){
    const reason = reasons.filter(Boolean).join(" | ") || coreResult.reason || "";
    noteEventSupport(false, reason);
    return { supported:false, reason };
  }
  const aggregatedReason = reasons.filter(Boolean).join(" | ") || "Capabilities nicht auswertbar.";
  noteEventSupport(null, aggregatedReason);
  L("event support indeterminate", { reason: aggregatedReason });
  return { supported:null, reason: aggregatedReason };
}

/**
 * Cache für System-Adressbuch-Einträge, um CardDAV-Traffic zu begrenzen.
 * Structure:
 * {
 *   contacts: Array<Contact>,
 *   fetchedAt: number,
 *   baseUrl: string,
 *   user: string
 * }
 */
const SYSTEM_ADDRESSBOOK_CACHE = {
  contacts: [],
  fetchedAt: 0,
  baseUrl: "",
  user: ""
};
const SYSTEM_ADDRESSBOOK_TTL = 5 * 60 * 1000;

/**
 * Entfernt vCard-Escaping (z.B. \n, \\, \;) aus einem Feld.
 * @param {string} value
 * @returns {string}
 */
function decodeVCardValue(value){
  return String(value ?? "")
    .replace(/\\\\/g, "\\")
    .replace(/\\n/gi, "\n")
    .replace(/\\,/g, ",")
    .replace(/\\;/g, ";")
    .replace(/\\:/g, ":");
}

/**
 * Klappt gefaltete vCard-Zeilen (RFC 6350) wieder auf.
 * @param {string} data
 * @returns {string[]}
 */
function unfoldVCardLines(data){
  const normalized = String(data ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  const unfolded = [];
  for (const line of lines){
    if (!line.length && !unfolded.length){
      continue;
    }
    if ((line.startsWith(" ") || line.startsWith("\t")) && unfolded.length){
      unfolded[unfolded.length - 1] += line.slice(1);
    } else {
      unfolded.push(line);
    }
  }
  return unfolded;
}

function normalizePhotoMime(value){
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const lower = raw.toLowerCase();
  if (lower.startsWith("image/")) return lower;
  if (lower === "jpeg" || lower === "jpg") return "image/jpeg";
  if (lower === "png") return "image/png";
  if (lower === "gif") return "image/gif";
  if (lower === "bmp") return "image/bmp";
  if (lower === "webp") return "image/webp";
  return lower;
}

function extractMimeFromDataUrl(dataUrl){
  const match = /^data:([^;,]+)[;,]?/i.exec(dataUrl);
  return match && match[1] ? match[1].toLowerCase() : "";
}

function createPhotoDataUrl(photo){
  if (!photo) return null;
  const raw = String(photo.raw || "").trim();
  if (!raw) return null;
  if (raw.startsWith("data:")){
    const mime = extractMimeFromDataUrl(raw) || photo.mime || "";
    return { dataUrl: raw, mime };
  }
  if ((photo.valueType || "") === "uri"){
    if (raw.startsWith("data:")){
      const mime = extractMimeFromDataUrl(raw) || photo.mime || "";
      return { dataUrl: raw, mime };
    }
    return null;
  }
  const encoding = (photo.encoding || "").toLowerCase();
  if (encoding === "b" || encoding === "base64" || !encoding){
    const cleaned = raw.replace(/\s+/g, "");
    if (!cleaned) return null;
    const mime = photo.mime || "image/png";
    const dataUrl = "data:" + (mime || "image/png") + ";base64," + cleaned;
    return { dataUrl, mime, base64: cleaned };
  }
  return null;
}

function getHiddenWindow(){
  try{
    if (typeof Services !== "undefined" && Services && Services.appShell?.hiddenDOMWindow){
      return Services.appShell.hiddenDOMWindow;
    }
  }catch(_){}
  if (typeof window !== "undefined" && window){
    return window;
  }
  if (typeof globalThis !== "undefined"){
    if (globalThis.window){
      return globalThis.window;
    }
    return globalThis;
  }
  return null;
}
function createScratchCanvas(width, height){
  const w = Math.max(1, Number(width) || 1);
  const h = Math.max(1, Number(height) || 1);
  if (typeof OffscreenCanvas === "function"){
    return new OffscreenCanvas(w, h);
  }
  const hidden = getHiddenWindow();
  if (hidden && hidden.document && typeof hidden.document.createElement === "function"){
    const canvas = hidden.document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    return canvas;
  }
  throw localizedError("error_canvas_support_missing");
}

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
 * Wandelt das exportierte System-Adressbuch in ein internes Format um.
 * @param {string} data - Rohes vCard-Dokument.
 * @returns {Array<{id:string,label:string,email:string,idLower:string,labelLower:string,emailLower:string,avatarDataUrl:string|null}>}
 */
function parseSystemAddressbook(data){
  const unfolded = unfoldVCardLines(data);
  const contacts = [];
  let card = null;
  for (const rawLine of unfolded){
    const line = rawLine.trimEnd();
    if (!line) continue;
    if (line.toUpperCase() === "BEGIN:VCARD"){
      card = { emails: [], photo: null };
      continue;
    }
    if (!card) continue;
    if (line.toUpperCase() === "END:VCARD"){
      if (card.uid && card.emails.length){
        const preferred = card.emails.find((item) => (item.params["X-NC-SCOPE"] || "").toLowerCase() === "v2-federated") ||
          card.emails[0];
        if (preferred && preferred.value){
          const label = card.fn || card.nickname || card.displayName || card.uid;
          const idLower = card.uid.toLowerCase();
          const labelLower = label ? label.toLowerCase() : "";
          const emailLower = preferred.value.toLowerCase();
          const avatar = createPhotoDataUrl(card.photo);
          contacts.push({
            id: card.uid,
            label,
            email: preferred.value,
            idLower,
            labelLower,
            emailLower,
            avatarDataUrl: avatar ? avatar.dataUrl : null
          });
        }
      }
      card = null;
      continue;
    }
    const colonIdx = line.indexOf(":");
    if (colonIdx === -1) continue;
    const lhs = line.slice(0, colonIdx);
    const rhs = line.slice(colonIdx + 1);
    const segments = lhs.split(";");
    const tag = (segments.shift() || "").toUpperCase();
    const params = {};
    for (const segment of segments){
      if (!segment) continue;
      const eqIdx = segment.indexOf("=");
      if (eqIdx === -1){
        params[segment.toUpperCase()] = true;
      } else {
        const key = segment.slice(0, eqIdx).toUpperCase();
        const value = decodeVCardValue(segment.slice(eqIdx + 1));
        params[key] = value;
      }
    }
    const decoded = decodeVCardValue(rhs);
    if (tag === "UID"){
      card.uid = decoded.trim();
    } else if (tag === "FN"){
      card.fn = decoded.trim();
    } else if (tag === "N" && !card.fn){
      const parts = decoded.split(";").filter(Boolean);
      if (parts.length){
        card.fn = parts.join(" ").trim();
      }
    } else if (tag === "NICKNAME" && !card.nickname){
      card.nickname = decoded.trim();
    } else if (tag === "EMAIL"){
      const email = decoded.trim();
      if (email){
        card.emails.push({ value: email, params });
      }
    } else if (tag === "X-ABSHOWAS"){
      card.displayName = decoded.trim();
    } else if (tag === "PHOTO"){
      const encoding = String(params.ENCODING || params.ENC || "").toLowerCase();
      const valueType = String(params.VALUE || "").toLowerCase();
      const mime = normalizePhotoMime(params.TYPE || params.MEDIATYPE);
      card.photo = {
        raw: decoded,
        encoding,
        valueType,
        mime,
        params
      };
    }
  }
  contacts.sort((a, b) => {
    const labelA = a.labelLower || a.emailLower || a.idLower;
    const labelB = b.labelLower || b.emailLower || b.idLower;
    const cmp = labelA.localeCompare(labelB, undefined, { sensitivity: "base" });
    if (cmp !== 0) return cmp;
    return a.idLower.localeCompare(b.idLower);
  });
  return contacts;
}

/**
 * Holt das System-Adressbuch aus Nextcloud (oder den Cache).
 * @param {boolean} [force=false] - erzwingt frischen Abruf.
 * @returns {Promise<Array<{id:string,label:string,email:string,idLower:string,labelLower:string,emailLower:string,avatarDataUrl:string|null}>>}
 */
async function getSystemAddressbookContacts(force = false){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  const now = Date.now();
  if (!force &&
      SYSTEM_ADDRESSBOOK_CACHE.contacts.length &&
      SYSTEM_ADDRESSBOOK_CACHE.user === user &&
      SYSTEM_ADDRESSBOOK_CACHE.baseUrl === baseUrl &&
      now - SYSTEM_ADDRESSBOOK_CACHE.fetchedAt < SYSTEM_ADDRESSBOOK_TTL){
    L("system addressbook cache hit", {
      entries: SYSTEM_ADDRESSBOOK_CACHE.contacts.length,
      ageMs: now - SYSTEM_ADDRESSBOOK_CACHE.fetchedAt
    });
    return SYSTEM_ADDRESSBOOK_CACHE.contacts;
  }
  const auth = "Basic " + btoa(user + ":" + appPass);
  const base = baseUrl.replace(/\/$/,"");
  L("system addressbook fetch", { base, user });
  // Zugriff auf das serverseitige System-Adressbuch (CardDAV) – erfordert remote.php-Berechtigung.
  const addressUrl = base + "/remote.php/dav/addressbooks/users/" + encodeURIComponent(user) + "/z-server-generated--system/?export";
  const res = await fetch(addressUrl, {
    method: "GET",
    headers: {
      "Authorization": auth,
      "Accept": "text/directory",
      "Cache-Control": "no-cache"
    }
  });
  if (!res.ok){
    const text = await res.text().catch(() => "");
    throw localizedError("error_system_addressbook_failed", [text || (res.status + " " + res.statusText)]);
  }
  const raw = await res.text();
  const contacts = parseSystemAddressbook(raw);
  L("system addressbook fetched", { count: contacts.length });
  SYSTEM_ADDRESSBOOK_CACHE.contacts = contacts;
  SYSTEM_ADDRESSBOOK_CACHE.fetchedAt = now;
  SYSTEM_ADDRESSBOOK_CACHE.user = user;
  SYSTEM_ADDRESSBOOK_CACHE.baseUrl = baseUrl;
  return contacts;
}

/**
 * Filtert Kontakte aus dem System-Adressbuch anhand eines Suchbegriffs.
 * @param {{searchTerm?:string, limit?:number, forceRefresh?:boolean}} [param0]
 * @returns {Promise<Array<{id:string,label:string,email:string,avatarDataUrl:string|null}>>}
 */
async function searchSystemAddressbook({ searchTerm = "", limit = 200, forceRefresh = false } = {}){
  const contacts = await getSystemAddressbookContacts(forceRefresh);
  const term = String(searchTerm || "").trim().toLowerCase();
  let filtered = contacts;
  if (term){
    filtered = contacts.filter((entry) => {
      return entry.idLower.includes(term) ||
        (entry.labelLower && entry.labelLower.includes(term)) ||
        entry.emailLower.includes(term);
    });
  }
  const limited = typeof limit === "number" && limit > 0 ? filtered.slice(0, limit) : filtered;
  L("search system addressbook", {
    term,
    limit,
    total: contacts.length,
    matches: limited.length
  });
  return limited.map(({ id, label, email, avatarDataUrl }) => ({
    id,
    label,
    email,
    avatarDataUrl: avatarDataUrl || null
  }));
}

function randToken(len=10){ const a="abcdefghijklmnopqrstuvwxyz0123456789"; let s=""; for(let i=0;i<len;i++) s+=a[Math.floor(Math.random()*a.length)]; return s; }

function sanitizeDescription(desc){
  if (!desc) return "";
  return String(desc).trim();
}

function buildRoomDescription(baseDescription, url, password){
  const parts = [];
  if (baseDescription && String(baseDescription).trim()) parts.push(String(baseDescription).trim());
  if (url) parts.push("Talk-Link: " + url);
  if (password) parts.push("Passwort: " + password);
  return parts.join("\n\n").trim();
}

async function createTalkPublicRoom({
  title,
  password,
  enableLobby,
  enableListable,
  description,
  startTimestamp,
  objectType,
  objectId,
  eventConversation
} = {}){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");

  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json", "Content-Type":"application/json" };

  const base = baseUrl.replace(/\/$/,"");
  const createUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room";
  const ROOM_TYPE_PUBLIC = 3;
  const LISTABLE_NONE = 0;
  const LISTABLE_USERS = 1;
  const listableScope = enableListable ? LISTABLE_USERS : LISTABLE_NONE;
  const cleanedDescription = sanitizeDescription(description);
  const attemptEvent = !!(eventConversation && objectType === "event" && objectId && String(objectId).trim().length);
  L("create talk room request", {
    title: title || "",
    hasPassword: !!password,
    enableLobby: !!enableLobby,
    enableListable: !!enableListable,
    descriptionLength: (cleanedDescription || "").length,
    attemptEvent,
    objectType: attemptEvent ? objectType : null,
    objectId: attemptEvent ? shortId(objectId) : null,
    startTimestamp: typeof startTimestamp === "number" ? startTimestamp : null
  });
  let supportInfo = { supported:null, reason:"" };
  if (attemptEvent){
    supportInfo = await getEventConversationSupport();
    L("event support info (create)", {
      supported: supportInfo.supported,
      reason: supportInfo.reason || ""
    });
  }
  const attempts = [];
  if (attemptEvent && supportInfo.supported !== false){
    attempts.push({ includeEvent:true });
  }
  attempts.push({ includeEvent:false });

  let lastError = null;
  for (const attempt of attempts){
    L("create attempt start", { includeEvent: attempt.includeEvent });
    const body = {
      roomType: ROOM_TYPE_PUBLIC,
      type: ROOM_TYPE_PUBLIC,
      roomName: title || "Besprechung",
      listable: listableScope,
      participants: {}
    };
    if (password) body.password = password;
    if (cleanedDescription) body.description = cleanedDescription;
    if (attempt.includeEvent){
      body.objectType = "event";
      body.objectId = String(objectId).trim();
    }
    const res = await fetch(createUrl, { method:"POST", headers, body: JSON.stringify(body) });
    L("create attempt status", { includeEvent: attempt.includeEvent, status: res.status, ok: res.ok });
    const raw = await res.text();
    let data = null;
    try { data = raw ? JSON.parse(raw) : null; } catch(_){ }
    if (!res.ok){
      const meta = data?.ocs?.meta || {};
      const payload = data?.ocs?.data || {};
      L("create attempt failure", {
        includeEvent: attempt.includeEvent,
        status: res.status,
        meta: meta?.message || null,
        error: payload?.error || null
      });
      if (attempt.includeEvent && isEventConversationError(meta, payload, raw)){
        markEventSupportUnsupported(payload?.error || meta?.message || "");
        L("event conversation rejected by server, falling back");
        continue;
      }
      const parts = [];
      if (meta.message && meta.message !== meta.status) parts.push(meta.message);
      if (payload.error) parts.push(payload.error);
      if (Array.isArray(payload.errors)) parts.push(...payload.errors);
      if (meta.statuscode) parts.push("Statuscode " + meta.statuscode);
      if (res.status) parts.push("HTTP " + res.status + " " + res.statusText);
      const detail = parts.filter(Boolean).join(" / ") || raw || (res.status + " " + res.statusText);
      const err = localizedError("error_ocs", [detail]);
      err.fatal = true;
      err.status = res.status;
      err.response = raw;
      err.meta = meta;
      err.payload = payload;
      lastError = err;
      break;
    }
    if (attempt.includeEvent){
      noteEventSupport(true, "");
    }
    let token = data?.ocs?.data?.token || data?.ocs?.data?.roomToken || data?.token || data?.data?.token;
    if (!token){
      lastError = localizedError("error_token_missing_in_response");
      break;
    }
    L("create attempt success", {
      includeEvent: attempt.includeEvent,
      token: shortToken(token)
    });
    const url = base + "/call/" + token;
    if (enableLobby){
      try{
        const lobbyUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + token + "/webinar/lobby";
        const lobbyPayload = { state: 1 };
        if (typeof startTimestamp === "number" && Number.isFinite(startTimestamp) && startTimestamp > 0){
          let timerVal = startTimestamp;
          if (timerVal > 1e12) timerVal = Math.floor(timerVal / 1000);
          lobbyPayload.timer = Math.floor(timerVal);
        }
        L("set lobby payload", lobbyPayload);
        const lobbyRes = await fetch(lobbyUrl, { method:"PUT", headers, body: JSON.stringify(lobbyPayload) });
        if (!lobbyRes.ok){
          const lobbyText = await lobbyRes.text().catch(() => "");
          L("lobby set failed", lobbyRes.status, lobbyRes.statusText, lobbyText);
          throw localizedError("error_lobby_set_failed", [lobbyText || (lobbyRes.status + " " + lobbyRes.statusText)]);
        }
        L("lobby set success", {
          token: shortToken(token),
          timer: lobbyPayload.timer ?? null
        });
      }catch(e){
        L("lobby update error", e?.message || String(e));
        return { url, token, fallback:true, reason: e?.message || bgI18n("error_lobby_set_failed_short") };
      }
    }
    if (enableListable){
      try{
        const listableUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + token + "/listable";
        const listableRes = await fetch(listableUrl, { method:"PUT", headers, body: JSON.stringify({ scope: listableScope }) });
        if (!listableRes.ok){
          L("listable set failed status", listableRes.status, listableRes.statusText);
        } else {
          L("listable set success", {
            token: shortToken(token),
            scope: listableScope
          });
        }
      }catch(e){
        L("listable update error", e?.message || String(e));
      }
    }
    const finalDescription = buildRoomDescription(description, url, password);
    const allowDescriptionUpdate = !(attempt.includeEvent && eventConversation);
    if (allowDescriptionUpdate && finalDescription && finalDescription !== cleanedDescription){
      try{
        const descUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + token + "/description";
        const descRes = await fetch(descUrl, { method:"PUT", headers, body: JSON.stringify({ description: finalDescription }) });
        if (!descRes.ok){
          L("description set failed status", descRes.status, descRes.statusText);
        } else {
          L("description update success", { token: shortToken(token) });
        }
      }catch(e){
        L("description update error", e?.message || String(e));
      }
    }
    const fallbackFlag = attemptEvent && !attempt.includeEvent;
    const fallbackReason = fallbackFlag ? supportInfo.reason || "Event-Konversation nicht verfuegbar." : null;
    L("create attempt complete", {
      includeEvent: attempt.includeEvent,
      token: shortToken(token),
      fallback: fallbackFlag,
      reason: fallbackReason
    });
    return {
      url,
      token,
      fallback: fallbackFlag,
      reason: fallbackReason,
      description: finalDescription || cleanedDescription || ""
    };
  }
  if (lastError){
    if (lastError.fatal){
      L("create attempt fatal", {
        message: lastError?.message || "",
        status: lastError?.status || null
      });
      throw lastError;
    }
    L("create via OCS failed, fallback to pseudo url:", lastError?.message);
    const fallbackToken = randToken(10);
    L("create fallback token", { token: shortToken(fallbackToken) });
    return {
      url: base + "/call/" + fallbackToken,
      token: fallbackToken,
      fallback: true,
      reason: lastError?.message || String(lastError)
    };
  }
  L("create talk room failed", "unknown error");
  throw localizedError("error_room_create_failed");
}
async function updateTalkLobby({ token, enableLobby, startTimestamp } = {}){
  if (!token) throw localizedError("error_room_token_missing");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json", "Content-Type":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const lobbyUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + token + "/webinar/lobby";
  const payload = { state: enableLobby ? 1 : 0 };
  if (enableLobby && typeof startTimestamp === "number" && Number.isFinite(startTimestamp) && startTimestamp > 0){
    let timerVal = startTimestamp;
    if (timerVal > 1e12) timerVal = Math.floor(timerVal / 1000);
    payload.timer = Math.floor(timerVal);
  }
  if (!enableLobby) delete payload.timer;
  L("update lobby payload", payload);
  const res = await fetch(lobbyUrl, { method:"PUT", headers, body: JSON.stringify(payload) });
  if (!res.ok) {
    if (res.status === 403) {
      throw localizedError("error_lobby_no_permission");
    }
    throw localizedError("error_lobby_update_failed", [res.status]);
  }
  L("update lobby success", {
    token: shortToken(token),
    enableLobby: !!enableLobby,
    timer: payload.timer ?? null
  });
  return true;
}

function isTruthy(value){
  if (value === true || value === 1) return true;
  if (typeof value === "string"){
    const lower = value.toLowerCase();
    return ["true","1","yes","open","public","all","guests","guest","world"].includes(lower);
  }
  return false;
}

function isFalsy(value){
  if (value === false || value === 0) return true;
  if (typeof value === "string"){
    const lower = value.toLowerCase();
    return ["false","0","no","closed","private","none","off","invite","invite_only","restricted"].includes(lower);
  }
  return false;
}

async function listTalkUserRooms({ searchTerm = "" } = {}){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  L("list user rooms request", { searchTerm: searchTerm || "" });
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const url = base + "/ocs/v2.php/apps/spreed/api/v4/room?includeStatus=true&includeLastMessage=false";
  const res = await fetch(url, { method:"GET", headers });
  const raw = await res.text();
  let data = null;
  try { data = raw ? JSON.parse(raw) : null; } catch(_){}
  if (!res.ok){
    const meta = data?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw localizedError("error_ocs", [detail]);
  }
  const rooms = Array.isArray(data?.ocs?.data) ? data.ocs.data : [];
  const term = searchTerm ? searchTerm.toLowerCase() : null;
  const out = [];
  for (const room of rooms){
    const normalized = normalizeRoomEntry(room, base, "own");
    if (!normalized) continue;
    if (term){
      const hay = [normalized.displayName, normalized.name, normalized.description].filter(Boolean).join(" ").toLowerCase();
      if (!hay.includes(term)) continue;
    }
    normalized.source = normalized.source || "own";
    out.push(normalized);
  }
  L("list user rooms result", { total: rooms.length, filtered: out.length });
  return out;
}

/**
 * Normalisiert Talk-Raumdaten (egal ob öffentlich gelistet oder Eigenbestand).
 * Filtert gleichzeitig Räume ohne Gastzugang heraus.
 * @param {object} room - API-Response-Objekt.
 * @param {string} [base=""] - Basis-URL für Avatar-Link.
 * @param {string} [source=""] - Quelle (listed|own).
 * @returns {object|null}
 */
function normalizeRoomEntry(room, base = "", source = ""){
  if (!room || typeof room !== "object") return null;
  const token = room.token || room.roomToken || room.id || room.identifier;
  if (!token) return null;
  const type = room.type ?? room.roomType ?? room.kind ?? null;
  const config = room.config || {};
  const listable = room.listable === 1 || room.listable === true;
  const accessRaw = config.access ?? config.accessLevel ?? room.access ?? room.accessLevel ?? null;
  const access = typeof accessRaw === "string" ? accessRaw.toLowerCase() : accessRaw;
  const guestsFlag =
    config.allowGuests ??
    config.allow_guests ??
    config.guestAccess ??
    config.guest_access ??
    room.allowGuests ??
    room.allow_guests ??
    room.guestAccess ??
    room.guest_access ??
    null;
  let guestsAllowed = null;
  if (guestsFlag != null){
    if (isTruthy(guestsFlag)) guestsAllowed = true;
    else if (isFalsy(guestsFlag)) guestsAllowed = false;
  }
  if (guestsAllowed == null && access){
    if (isTruthy(access)) guestsAllowed = true;
    else if (isFalsy(access)) guestsAllowed = false;
  }
  if (guestsAllowed == null){
    if (room.guestURL || room.link || room.publicUrl || room.guestUrl) guestsAllowed = true;
  }
  if (guestsAllowed == null && typeof type === "number" && type === 3){
    guestsAllowed = true;
  }
  if (guestsAllowed == null && source){
    if (source === "own" || source === "listed-own") guestsAllowed = true;
  }
  if (guestsAllowed == null && listable) guestsAllowed = true;
  if (guestsAllowed == null) guestsAllowed = false;
  const allowedType = typeof type === "number" ? (type === 3 || type === 2) : true;
  if (!allowedType && !listable && !guestsAllowed) return null;
  const normalized = Object.assign({}, room);
  normalized.token = token;
  if (!normalized.displayName && normalized.name) normalized.displayName = normalized.name;
  normalized.listable = listable ? 1 : 0;
  normalized.access = access || normalized.access || null;
  normalized.guestsAllowed = !!guestsAllowed;
  normalized.source = source || normalized.source || null;
  normalized.isParticipant =
    room.isParticipant === true ||
    typeof room.participantType === "number" ||
    (Array.isArray(room.participants) && room.participants.length > 0) ||
    normalized.source === "own";
  return normalized;
}

/**
 * Liefert eine kombinierte Liste aller öffentlichen bzw. freigegebenen Talk-Räume.
 * Verbindet serverweite Listen mit eigenen Räumen und filtert doppelte Einträge.
 * @param {{searchTerm?:string}} [param0]
 * @returns {Promise<object[]>}
 */
async function listTalkPublicRooms({ searchTerm = "" } = {}){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  L("list public rooms request", { searchTerm: searchTerm || "" });
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const listedUrl = base + "/ocs/v2.php/apps/spreed/api/v4/listed-room?searchTerm=" + encodeURIComponent(searchTerm || "");
  const [listedRes, personalRooms] = await Promise.all([
    fetch(listedUrl, { method:"GET", headers }),
    listTalkUserRooms({ searchTerm })
  ]);
  const rawListed = await listedRes.text();
  let listedData = null;
  try { listedData = rawListed ? JSON.parse(rawListed) : null; } catch(_){}
  if (!listedRes.ok){
    const meta = listedData?.ocs?.meta || {};
    const detail = meta.message || rawListed || (listedRes.status + " " + listedRes.statusText);
    throw localizedError("error_ocs", [detail]);
  }
  const listedRooms = Array.isArray(listedData?.ocs?.data) ? listedData.ocs.data : [];
  const map = new Map();
  const push = (room, source) => {
    const normalized = normalizeRoomEntry(room, base, source);
    if (!normalized) return;
    const token = normalized.token || normalized.roomToken;
    if (!token) return;
    if (searchTerm){
      const term = searchTerm.toLowerCase();
      const hay = [normalized.displayName, normalized.name, normalized.description].filter(Boolean).join(" ").toLowerCase();
      if (!hay.includes(term)) return;
    }
    normalized.source = normalized.source || source || "unknown";
    map.set(token, normalized);
  };
  for (const room of listedRooms){
    push(room, "listed");
  }
  for (const room of personalRooms){
    push(room, "own");
  }
  const filtered = Array.from(map.values()).filter((room) => {
    const joinable = room.listable || room.source === "own" || room.isParticipant;
    return joinable && room.guestsAllowed !== false;
  });
  L("list public rooms result", {
    listedCount: listedRooms.length,
    personalCount: personalRooms.length,
    total: filtered.length
  });
  return filtered;
}

async function getTalkRoomInfo({ token } = {}){
  if (!token) throw localizedError("error_room_token_missing");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  L("get room info request", { token: shortToken(token) });
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const infoUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token);
  const res = await fetch(infoUrl, { method:"GET", headers });
  const raw = await res.text();
  let data = null;
  try { data = raw ? JSON.parse(raw) : null; } catch(_){}
  if (!res.ok){
    const meta = data?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw localizedError("error_ocs", [detail]);
  }
  const room = data?.ocs?.data;
  if (!room || typeof room !== "object"){
    throw localizedError("error_room_details_missing");
  }
  L("get room info success", { token: shortToken(token) });
  return room;
}

/**
 * Ruft Teilnehmer eines Talk-Raums ab. 404 wird als "keine Daten" interpretiert.
 * @param {{token:string}} param0
 * @returns {Promise<object[]>}
 */
async function getTalkRoomParticipants({ token } = {}){
  if (!token) throw localizedError("error_room_token_missing");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  L("get room participants request", { token: shortToken(token) });
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const infoUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/participants?includeStatus=true";
  const res = await fetch(infoUrl, { method:"GET", headers });
  const raw = await res.text();
  let data = null;
  try { data = raw ? JSON.parse(raw) : null; } catch(_){}
  if (res.status === 404){
    return [];
  }
  if (!res.ok){
    const meta = data?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw localizedError("error_ocs", [detail]);
  }
  const participants = data?.ocs?.data;
  const list = Array.isArray(participants) ? participants : [];
  L("get room participants result", { token: shortToken(token), count: list.length });
  return list;
}

async function addTalkParticipant({ token, actorId, source = "users" } = {}){
  if (!token || !actorId) throw localizedError("error_token_or_actor_missing");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  L("add participant request", {
    token: shortToken(token),
    actor: String(actorId).trim(),
    source: source || "users"
  });
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": auth,
    "Accept": "application/json",
    "Content-Type": "application/json"
  };
  const base = baseUrl.replace(/\/$/,"");
  const url = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/participants";
  const body = { newParticipant: actorId, source: source || "users" };
  const res = await fetch(url, { method:"POST", headers, body: JSON.stringify(body) });
  const raw = await res.text().catch(() => "");
  let json = null;
  try { json = raw ? JSON.parse(raw) : null; } catch(_){}
  if (!res.ok && res.status !== 409){
    const meta = json?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw localizedError("error_participant_add_failed", [detail]);
  }
  L("add participant result", {
    token: shortToken(token),
    status: res.status,
    conflict: res.status === 409
  });
  const added = json?.ocs?.data;
  return added || null;
}

async function promoteTalkModerator({ token, attendeeId } = {}){
  if (!token || typeof attendeeId !== "number") throw localizedError("error_moderator_id_missing");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  L("promote moderator request", {
    token: shortToken(token),
    attendeeId
  });
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": auth,
    "Accept": "application/json",
    "Content-Type": "application/json"
  };
  const base = baseUrl.replace(/\/$/,"");
  const url = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/moderators";
  const body = { attendeeId };
  const res = await fetch(url, { method:"POST", headers, body: JSON.stringify(body) });
  if (!res.ok && res.status !== 409){
    const raw = await res.text().catch(() => "");
    let json = null;
    try { json = raw ? JSON.parse(raw) : null; } catch(_){}
    const meta = json?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw localizedError("error_moderator_set_failed", [detail]);
  }
  L("promote moderator success", {
    token: shortToken(token),
    status: res.status,
    conflict: res.status === 409
  });
  return true;
}

async function leaveTalkRoom({ token } = {}){
  if (!token) throw localizedError("error_room_token_missing");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  L("leave room request", { token: shortToken(token) });
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": auth,
    "Accept": "application/json"
  };
  const base = baseUrl.replace(/\/$/,"");
  const url = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/participants/self";
  const res = await fetch(url, { method:"DELETE", headers });
  if (!res.ok && res.status !== 404){
    const raw = await res.text().catch(() => "");
    let json = null;
    try { json = raw ? JSON.parse(raw) : null; } catch(_){}
    const meta = json?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw localizedError("error_leave_failed", [detail]);
  }
  L("leave room success", { token: shortToken(token), status: res.status });
  return true;
}

async function delegateRoomModerator({ token, newModerator } = {}){
  if (!token || !newModerator) throw localizedError("error_delegation_data_missing");
  const { user } = await getOpts();
  const targetId = String(newModerator).trim();
  if (!targetId) throw localizedError("error_moderator_target_missing");
  const currentUser = (user || "").trim();
  L("delegate moderator request", {
    token: shortToken(token),
    target: targetId,
    currentUser
  });
  await addTalkParticipant({ token, actorId: targetId, source: "users" });
  const participants = await getTalkRoomParticipants({ token });
  const match = participants.find((p) => {
    if (!p) return false;
    const actor = (p.actorId || "").trim().toLowerCase();
    return actor === targetId.toLowerCase();
  });
  if (!match){
    throw localizedError("error_participant_not_found");
  }
  await promoteTalkModerator({ token, attendeeId: match.attendeeId });
  const loweredUser = currentUser.toLowerCase();
  if (loweredUser && loweredUser !== targetId.toLowerCase()){
    await leaveTalkRoom({ token });
    L("delegate moderator completed", { token: shortToken(token), delegate: targetId, leftSelf: true });
    return { leftSelf: true, delegate: targetId };
  }
  L("delegate moderator completed", { token: shortToken(token), delegate: targetId, leftSelf: false });
  return { leftSelf: false, delegate: targetId };
}

async function searchSharees({ searchTerm = "", limit = 50 } = {}){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  L("search sharees request", { searchTerm, limit });
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": auth,
    "Accept": "application/json",
    "X-Requested-With": "XMLHttpRequest"
  };
  const base = baseUrl.replace(/\/$/,"");
  const params = new URLSearchParams({
    format: "json",
    search: searchTerm,
    perPage: String(Math.max(1, Math.min(limit, 200))),
    itemType: "call",
    lookup: "false"
  });
  const url = base + "/ocs/v2.php/apps/files_sharing/api/v1/sharees?" + params.toString();
  const res = await fetch(url, { method:"GET", headers });
  const raw = await res.text().catch(() => "");
  let json = null;
  try { json = raw ? JSON.parse(raw) : null; } catch(_){}
  if (!res.ok){
    const meta = json?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw (detail ? new Error(detail) : localizedError("error_users_load_failed"));
  }
  const exactUsers = Array.isArray(json?.ocs?.data?.exact?.users) ? json.ocs.data.exact.users : [];
  const regularUsers = Array.isArray(json?.ocs?.data?.users) ? json.ocs.data.users : [];
  const combined = [...exactUsers, ...regularUsers];
  const emailPattern = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/;
  const seen = new Set();
  const suggestions = [];
  for (const entry of combined){
    const shareWith = entry?.value?.shareWith || entry?.value?.share_with;
    if (!shareWith) continue;
    const id = String(shareWith).trim();
    if (!id || seen.has(id)) continue;
    seen.add(id);
    const additionalInfo =
      entry?.value?.shareWithAdditionalInfo ??
      entry?.value?.share_with_additional_info ??
      entry?.value?.additionalInfo ??
      entry?.value?.additional_info ??
      null;
    const candidates = [
      entry?.label,
      additionalInfo,
      entry?.value?.email,
      entry?.value?.mail,
      entry?.value?.mailAddress,
      entry?.value?.mail_address,
      entry?.value?.emailAddress,
      entry?.value?.email_address,
      entry?.value?.EMail
    ];
    let email = null;
    for (const candidate of candidates){
      if (typeof candidate !== "string") continue;
      const match = candidate.match(emailPattern);
      if (match && match[0]){
        email = match[0].trim().toLowerCase();
        break;
      }
    }
    if (!email) continue;
    suggestions.push({
      id,
      label: entry?.label || id,
      email,
      shareType: entry?.value?.shareType ?? entry?.value?.share_type ?? null
    });
  }
  L("search sharees result", { matches: suggestions.length });
  return suggestions;
}

async function probeNextcloudCredentials({ baseUrl, user, appPass } = {}){
  const normalizedBase = normalizeBaseUrl(baseUrl);
  const trimmedUser = typeof user === "string" ? user.trim() : "";
  const password = typeof appPass === "string" ? appPass : "";
  if (!normalizedBase || !trimmedUser || !password){
    return { ok:false, code:"missing", message: bgI18n("error_credentials_missing") };
  }
  const basicHeader = "Basic " + btoa(trimmedUser + ":" + password);
  try{
    L("options test connection", { base: normalizedBase, user: shortId(trimmedUser) });
    const headers = {
      "OCS-APIRequest": "true",
      "Authorization": basicHeader,
      "Accept": "application/json"
    };
    const url = normalizedBase + "/ocs/v2.php/cloud/capabilities";
    const res = await fetch(url, { method:"GET", headers });
    const raw = await res.text().catch(() => "");
    let data = null;
    try{ data = raw ? JSON.parse(raw) : null; }catch(_){}
    if (res.status === 401 || res.status === 403){
      const detail = data?.ocs?.meta?.message || "HTTP " + res.status;
      return { ok:false, code:"auth", message: detail };
    }
    if (!res.ok){
      const detail = data?.ocs?.meta?.message || raw || (res.status + " " + res.statusText);
      return { ok:false, code:"http", message: detail };
    }
    const versionRaw = data?.ocs?.meta?.version || data?.ocs?.data?.version || "";
    let versionStr = "";
    if (typeof versionRaw === "string"){
      versionStr = versionRaw;
    } else if (versionRaw && typeof versionRaw === "object"){
      if (typeof versionRaw.string === "string" && versionRaw.string.trim()){
        versionStr = versionRaw.string.trim();
      } else {
        const parts = [];
        if (versionRaw.major != null) parts.push(String(versionRaw.major));
        if (versionRaw.minor != null) parts.push(String(versionRaw.minor));
        if (versionRaw.micro != null) parts.push(String(versionRaw.micro));
        if (parts.length){
          versionStr = parts.join(".");
        }
      }
    }
    const message = versionStr ? "Nextcloud " + versionStr : "";

    // Zusätzlicher Login-Check gegen /ocs/v2.php/cloud/user
    try{
      const userUrl = normalizedBase + "/ocs/v2.php/cloud/user";
      const userRes = await fetch(userUrl, {
        method: "GET",
        headers: {
          "OCS-APIRequest": "true",
          "Authorization": basicHeader,
          "Accept": "application/json"
        }
      });
      if (userRes.status === 401 || userRes.status === 403){
        return { ok:false, code:"auth", message: "Benutzername oder App-Passwort ung\u00fcltig." };
      }
      if (!userRes.ok){
        const userRaw = await userRes.text().catch(() => "");
        const userData = (() => { try{ return userRaw ? JSON.parse(userRaw) : null; }catch(_){ return null; }})();
        const detail = userData?.ocs?.meta?.message || userRaw || (userRes.status + " " + userRes.statusText);
        return { ok:false, code:"http", message: detail };
      }
    }catch(userErr){
      return { ok:false, code:"network", message: userErr?.message || String(userErr) };
    }

    return { ok:true, version: versionStr, message };
  }catch(e){
    return { ok:false, code:"network", message: e?.message || String(e) };
  }
}

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
  if (msg.type === "talkMenu:chooseExisting"){
    return { ok:true };
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
  if (msg.type === "talkMenu:listPublicRooms"){
    try{
      const rooms = await listTalkPublicRooms(msg.payload || {});
      return { ok:true, rooms };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:getRoomInfo"){
    try{
      const room = await getTalkRoomInfo(msg.payload || {});
      return { ok:true, room };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:getRoomParticipants"){
    try{
      const participants = await getTalkRoomParticipants(msg.payload || {});
      return { ok:true, participants };
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
    if (type === "listPublicRooms"){
      const rooms = await listTalkPublicRooms({ searchTerm: payload?.searchTerm || "" });
      return { ok:true, rooms };
    }
    if (type === "getRoomInfo"){
      const room = await getTalkRoomInfo({ token: payload?.token });
      return { ok:true, room };
    }
    if (type === "getRoomParticipants"){
      const participants = await getTalkRoomParticipants({ token: payload?.token });
      return { ok:true, participants };
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














