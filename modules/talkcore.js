/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
/**
 * Ladet die in den Add-on-Optionen hinterlegten Nextcloud-Credentials
 * und bereinigt dabei z.B. überflüssige Slashes in der Basis-URL.
 * @returns {Promise<{baseUrl:string,user:string,appPass:string}>}
 */


/**
 * Liefert Add-on-Credentials. Bevorzugt NCCore.getOpts(),
 * fällt aber auf lokale Speicherabfrage zurück, wenn NCCore hier nicht bekannt ist.
 * @returns {Promise<{baseUrl:string,user:string,appPass:string,debugEnabled:boolean,authMode:string}>}
 */
async function getOpts(){
  return NCCore.getOpts();
}

const EVENT_SUPPORT_CACHE = {
  value: null,
  reason: "",
  expires: 0
};
const EVENT_SUPPORT_TTL = 5 * 60 * 1000;

/**
 * Schreibt die ermittelte Event-Unterstützung in den lokalen Cache.
 * @param {boolean|null} value - Ergebnis der Event-Fähigkeit (true/false/null)
 * @param {string} reason - Erläuterung für Logs und Debugging
 */
function noteEventSupport(value, reason){
  EVENT_SUPPORT_CACHE.value = value;
  EVENT_SUPPORT_CACHE.reason = reason || "";
  EVENT_SUPPORT_CACHE.expires = Date.now() + EVENT_SUPPORT_TTL;
}

/**
 * Convenience-Funktion, um Event-Unterstützung negativ zu cachen.
 * @param {string} reason - Erläuterung, weshalb Event-Konversationen nicht verfügbar sind
 */
function markEventSupportUnsupported(reason){
  noteEventSupport(false, reason || "");
}

/**
 * Analysiert den Capabilities-Response von Talk/Cloud
 * und versucht ein Boolean für Event-Konversationen abzuleiten.
 * @param {object} data - Capability-Ausschnitt des Servers
 * @returns {{status:boolean|null, hint:string}}
 */
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

/**
 * Extrahiert eine grobe Major-Version aus verschiedenen Feldformaten.
 * @param {* } value - beliebiges Versionsfeld
 * @returns {number|null} - Major-Version oder null
 */
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

/**
 * Fragt die Talk-spezifischen Capabilities ab und liefert ein einheitliches Ergebnis.
 * @param {string} url - vollständiger Endpunkt
 * @param {object} headers - vorbereitete OCS-Header
 * @returns {Promise<{supported:boolean|null, reason:string}>}
 */
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

/**
 * Fragt die generischen Nextcloud Capabilities ab und interpretiert sie bzgl. Event-Unterstützung.
 * @param {string} baseUrl - bereinigte Nextcloud Basis-URL
 * @param {object} headers - vorbereitete OCS-Header
 * @returns {Promise<{supported:boolean|null, reason:string}>}
 */
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

/**
 * Ermittelt (mit Cache) ob Event-Konversationen unterstützt werden.
 * @returns {Promise<{supported:boolean|null, reason:string}>}
 */
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
 * Cache fÃ¼r System-Adressbuch-EintrÃ¤ge, um CardDAV-Traffic zu begrenzen.
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
/**
 * Dekodiert Escapes innerhalb einer vCard-Zeile.
 * @param {string} value - roher vCard-Inhalt nach dem Doppelpunkt
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
/**
 * Entfernt RFC-6350-Zeilenumbrüche innerhalb einer vCard.
 * @param {string} data - vollständiger vCard-Export
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
/**
 * Erstellt aus dem vCard-Export des Systemadressbuchs eine strukturierte Kontaktliste.
 * @param {string} data - Rohdaten aus dem CardDAV-Export
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
        const preferred = card.emails.find((item) => {
          const scope = (item.params["X-NC-SCOPE"] || "").toLowerCase();
          return scope === "v2-federated";
        }) || card.emails[0];
        if (preferred && preferred.value){
          const email = preferred.value.trim();
          if (email){
            const label = (card.fn || card.nickname || card.displayName || email || card.uid).trim() || email;
            const avatar = card.photo ? createPhotoDataUrl(card.photo) : null;
            contacts.push({
              id: card.uid,
              label,
              email,
              idLower: card.uid.toLowerCase(),
              labelLower: label.toLowerCase(),
              emailLower: email.toLowerCase(),
              avatarDataUrl: avatar ? avatar.dataUrl : null
            });
          }
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
    switch (tag){
      case "UID":
        if (decoded.trim()) card.uid = decoded.trim();
        break;
      case "FN":
        if (decoded.trim()) card.fn = decoded.trim();
        break;
      case "N":
        if (!card.fn){
          const parts = decoded.split(";");
          const given = (parts[1] || "").trim();
          const additional = (parts[2] || "").trim();
          const family = (parts[0] || "").trim();
          const labelParts = [given, additional, family].filter(Boolean);
          if (labelParts.length){
            card.fn = labelParts.join(" ");
          }
        }
        break;
      case "NICKNAME":
        if (decoded.trim()) card.nickname = decoded.trim();
        break;
      case "EMAIL": {
        const email = decoded.trim();
        if (email){
          card.emails.push({ value: email, params });
        }
        break;
      }
      case "PHOTO":
        card.photo = {
          raw: decoded,
          encoding: params.ENCODING || params["ENCODING"],
          valueType: params.VALUE || params["VALUE"],
          mime: params.TYPE || params["TYPE"] || params.MEDIATYPE || params["MEDIATYPE"] || ""
        };
        break;
      case "X-NC-SHARE-WITH-NAME":
      case "X-NC-SHARE-WITH-DISPLAYNAME":
        if (!card.displayName && decoded.trim()){
          card.displayName = decoded.trim();
        }
        break;
      case "ORG":
        if (!card.displayName && decoded.trim()){
          card.displayName = decoded.trim();
        }
        break;
      default:
        break;
    }
  }
  contacts.sort((a, b) => {
    const byLabel = a.labelLower.localeCompare(b.labelLower);
    if (byLabel !== 0) return byLabel;
    return a.idLower.localeCompare(b.idLower);
  });
  return contacts;
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

/**
 * Liefert ein Fensterobjekt, das DOM-APIs (Canvas etc.) bereitstellt.
 * @returns {Window|Global|null}
 */
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
/**
 * Erstellt einen Canvas für Imaging-Operationen – bevorzugt OffscreenCanvas.
 * @param {number} width
 * @param {number} height
 * @returns {HTMLCanvasElement|OffscreenCanvas}
 */
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


/**
 * Lädt und cached das Nextcloud-Systemadressbuch als strukturierte Liste.
 * @param {boolean} force - true erzwingt einen frischen Abruf
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
  // Zugriff auf das serverseitige System-Adressbuch (CardDAV) â€“ erfordert remote.php-Berechtigung.
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
/**
 * Filtert Kontakte aus dem Systemadressbuch nach Suchbegriff und Limit.
 * @param {{searchTerm?:string, limit?:number, forceRefresh?:boolean}} param0
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

/** Erzeugt einen einfachen Zufallstoken für Pseudo-Fallbacks. */
function randToken(len=10){ const a="abcdefghijklmnopqrstuvwxyz0123456789"; let s=""; for(let i=0;i<len;i++) s+=a[Math.floor(Math.random()*a.length)]; return s; }
/** Trimmt und normalisiert Raumbeschreibungen. */
function sanitizeDescription(desc){
  if (!desc) return "";
  return String(desc).trim();
}

/**
 * Fügt optionale Informationen (Talk-Link & Passwort) zu einer Beschreibung zusammen.
 */
function buildRoomDescription(baseDescription, url, password){
  const parts = [];
  if (baseDescription && String(baseDescription).trim()){
    parts.push(String(baseDescription).trim());
  }
  const talkBlock = buildStandardTalkDescription(url, password);
  if (talkBlock){
    parts.push(talkBlock);
  }
  return parts.join("\n\n").trim();
}

function descriptionI18n(key, substitutions = []){
  try{
    if (typeof bgI18n === "function"){
      const msg = bgI18n(key, substitutions);
      if (msg) return msg;
    }
  }catch(_){}
  try{
    if (typeof NCI18n !== "undefined" && typeof NCI18n.translate === "function"){
      const msg = NCI18n.translate(key, substitutions);
      if (msg) return msg;
    }
  }catch(_){}
  if (substitutions.length){
    return String(substitutions[0]);
  }
  return "";
}

function buildStandardTalkDescription(url, password){
  const heading = descriptionI18n("ui_description_heading") || "Nextcloud Talk";
  const joinLabel = descriptionI18n("ui_description_join_label") || "Jetzt an der Besprechung teilnehmen :";
  const passwordLine = password ? (descriptionI18n("ui_description_password_line", [password]) || `Passwort: ${password}`) : "";
  const helpLabel = descriptionI18n("ui_description_help_label") || "Benötigen Sie Hilfe?";
  const helpUrl = descriptionI18n("ui_description_help_url") || "https://docs.nextcloud.com/server/latest/user_manual/de/talk/join_a_call_or_chat_as_guest.html";
  const lines = [
    heading,
    "",
    joinLabel,
    url || "",
    ""
  ];
  if (passwordLine){
    lines.push(passwordLine, "");
  }
  lines.push(helpLabel, "", helpUrl);
  return lines.join("\n").trim();
}

/**
 * Erstellt (inkl. Fallbacks) einen Talk-Raum mit optionaler Event-Bindung.
 * @returns {Promise<{url:string,token:string,fallback:boolean,reason:string|null,description:string}>}
 */
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
/**
 * Aktualisiert den Lobby-Zustand (inkl. Startzeit) eines bestehenden Talk-Raums.
 */
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
/**
 * Löscht einen Talk-Raum via OCS-API. 404 wird als Erfolg interpretiert.
 */
async function deleteTalkRoom({ token } = {}){
  if (!token) throw localizedError("error_room_token_missing");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const url = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token);
  L("delete talk room request", { token: shortToken(token) });
  const res = await fetch(url, { method:"DELETE", headers });
  const raw = await res.text().catch(() => "");
  let data = null;
  try{ data = raw ? JSON.parse(raw) : null; }catch(_){}
  L("delete talk room status", { token: shortToken(token), status: res.status, ok: res.ok });
  if (res.status === 404){
    L("delete talk room already removed", { token: shortToken(token) });
    return true;
  }
  if (!res.ok){
    const meta = data?.ocs?.meta || {};
    const payload = data?.ocs?.data || {};
    const parts = [];
    if (meta.message && meta.message !== meta.status) parts.push(meta.message);
    if (payload.error) parts.push(payload.error);
    if (meta.statuscode) parts.push("Statuscode " + meta.statuscode);
    if (res.status) parts.push("HTTP " + res.status + " " + res.statusText);
    const detail = parts.filter(Boolean).join(" / ") || raw || (res.status + " " + res.statusText);
    throw localizedError("error_room_delete_failed", [detail]);
  }
  L("delete talk room success", { token: shortToken(token) });
  return true;
}

/**
 * Ruft Teilnehmer eines Talk-Raums ab. 404 wird als "keine Daten" interpretiert.
 * @param {{token:string}} param0
 * @returns {Promise<object[]>}
 */
/**
 * Ruft Teilnehmerliste eines Raums ab (inkl. Moderatorinformation).
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
/**
 * Fügt einen Benutzer über die OCS-API zum Talk-Raum hinzu.
 */
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
/**
 * Ernennt einen vorhandenen Raumteilnehmer zum Moderator.
 */
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
/**
 * Entfernt den authentifizierten Benutzer aus dem Talk-Raum (self-leave).
 */
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
/**
 * Überträgt die Moderation an einen anderen Benutzer (inkl. optionalem Selbst-Verlassen).
 */
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


