'use strict';
/**
 * Frontend-Teil der Erweiterung: injiziert den Toolbar-Button,
 * öffnet den Erstell-Dialog und bedient die Hilfsfunktionen im Terminfenster.
 */
var { ExtensionCommon } = ChromeUtils.importESModule("resource://gre/modules/ExtensionCommon.sys.mjs");
let ExtensionSupport = null, Services = null;
try { ExtensionSupport = ChromeUtils.importESModule("resource:///modules/ExtensionSupport.sys.mjs").ExtensionSupport; } catch(e) {}
try { Services = ChromeUtils.importESModule("resource://gre/modules/Services.sys.mjs").Services; } catch(e) {}
try {
  if (typeof Components !== "undefined" && Components.utils?.importGlobalProperties){
    Components.utils.importGlobalProperties(["atob","btoa","Image","Blob","FileReader","fetch"]);
  }
}catch(_){}

/**
 * Liefert das versteckte DOM-Fenster von Thunderbird, falls verfügbar.
 * Einige APIs (Image/FileReader) benötigen dieses Fenster, weil es Browser-Kontext bereitstellt.
 * @returns {Window|null}
 */
function getHiddenDOMWindow(){
  try{
    if (!Services){
      Services = ChromeUtils.importESModule("resource://gre/modules/Services.sys.mjs").Services;
    }
  }catch(e){
    return null;
  }
  try{
    return Services?.appShell?.hiddenDOMWindow || null;
  }catch(e){
    return null;
  }
}

/**
 * Stellt sicher, dass globale Browser-APIs (Image, Blob, fetch, …) im Kontext verfügbar sind.
 * Thunderbird deaktiviert diese teilweise im Add-on-Prozess.
 */
function ensureBrowserGlobals(){
  const hidden = getHiddenDOMWindow();
  try{
    if (hidden){
      if (typeof globalThis.Image === "undefined" && hidden.Image){
        globalThis.Image = hidden.Image;
      }
      if (typeof globalThis.Blob === "undefined" && hidden.Blob){
        globalThis.Blob = hidden.Blob;
      }
      if (typeof globalThis.FileReader === "undefined" && hidden.FileReader){
        globalThis.FileReader = hidden.FileReader;
      }
      if (typeof globalThis.atob !== "function" && typeof hidden.atob === "function"){
        globalThis.atob = hidden.atob.bind(hidden);
      }
      if (typeof globalThis.btoa !== "function" && typeof hidden.btoa === "function"){
        globalThis.btoa = hidden.btoa.bind(hidden);
      }
      if (typeof globalThis.fetch !== "function" && typeof hidden.fetch === "function"){
        globalThis.fetch = hidden.fetch.bind(hidden);
      }
      if (typeof globalThis.URL === "undefined" && hidden.URL){
        globalThis.URL = hidden.URL;
      }
      if (typeof globalThis.createImageBitmap !== "function" && typeof hidden.createImageBitmap === "function"){
        globalThis.createImageBitmap = hidden.createImageBitmap.bind(hidden);
      }
    }
  }catch(e){
    err(e);
  }
}

ensureBrowserGlobals();

const ALERT_TITLE_FALLBACK = "Nextcloud Talk";
let LAST_CONTEXT = null;

const AVATAR_BITMAP_CACHE = new Map();

function shortToken(token, keepStart = 4, keepEnd = 3){
  if (!token) return "";
  const str = String(token);
  if (str.length <= keepStart + keepEnd + 3){
    return str;
  }
  return str.slice(0, keepStart) + "..." + str.slice(str.length - keepEnd);
}

function shortString(value, max = 20){
  if (value == null) return "";
  const str = String(value);
  if (str.length <= max){
    return str;
  }
  return str.slice(0, max) + "...";
}

function describeCreatePayload(payload){
  if (!payload || typeof payload !== "object") return {};
  return {
    title: payload.title || "",
    enableLobby: !!payload.enableLobby,
    enableListable: !!payload.enableListable,
    hasPassword: !!payload.password,
    descriptionLength: payload.description ? String(payload.description).length : 0,
    startTimestamp: typeof payload.startTimestamp === "number" ? payload.startTimestamp : null,
    eventConversation: !!payload.eventConversation,
    objectType: payload.objectType || null,
    objectId: payload.objectId ? shortString(payload.objectId, 12) : null
  };
}

function summarizeUtilityPayload(payload){
  if (!payload || typeof payload !== "object") return {};
  return {
    type: payload.type || "",
    token: payload.token ? shortToken(payload.token) : "",
    searchTerm: payload.searchTerm || "",
    limit: typeof payload.limit === "number" ? payload.limit : undefined,
    delegate: payload.newModerator || ""
  };
}

const I18N_FALLBACKS = {
  ui_button_ok: "OK",
  ui_button_cancel: "Abbrechen",
  ui_button_apply: "\u00dcbernehmen",
  ui_button_apply_progress: "\u00dcbernehme...",
  ui_button_create_progress: "Erstelle...",
  ui_button_clear: "Leeren",
  ui_insert_button_label: "Talk-Link einf\u00fcgen",
  ui_create_heading: "\u00d6ffentliche Unterhaltung erstellen",
  ui_create_title_label: "Titel",
  ui_create_password_label: "Passwort (optional)",
  ui_create_password_placeholder: "Optional",
  ui_create_lobby_label: "Lobby bis Startzeit",
  ui_create_listable_label: "In Suche anzeigen",
  ui_create_roomtype_label: "Raumtyp",
  ui_create_mode_event: "Event-Konversation (Talk Raum an Termin binden)",
  ui_create_mode_standard: "Standard-Raum (eigenstaendig)",
  ui_create_mode_unsupported: "Event-Konversation wird vom Server nicht unterstuetzt.",
  ui_create_moderator_label: "Moderator (optional)",
  ui_create_moderator_placeholder: "Benutzername eingeben",
  ui_create_moderator_hint: "Bei Angabe wird die Moderation nach Erstellung an diesen Benutzer uebertragen und Sie verlassen den Raum.",
  ui_delegate_selected_title: "Ausgewaehlt",
  ui_delegate_status_searching: "Suche...",
  ui_delegate_status_loading: "Lade Benutzer...",
  ui_delegate_status_none_with_email: "Keine Treffer mit E-Mail.",
  ui_delegate_status_none_found: "Keine Benutzer mit E-Mail gefunden.",
  ui_delegate_status_single: "1 Treffer mit E-Mail.",
  ui_delegate_status_many: "$1 Treffer mit E-Mail.",
  ui_delegate_status_error: "Benutzersuche fehlgeschlagen.",
  ui_create_password_short: "Das Passwort muss mindestens 5 Zeichen lang sein.",
  ui_create_send_failed: "Senden an Hintergrundskript fehlgeschlagen:\n$1",
  ui_create_no_handler: "Kein Nextcloud Talk Handler registriert. Bitte Add-on neu laden.",
  ui_create_unknown_error: "Unbekannter Fehler beim Erstellen der Unterhaltung.",
  ui_create_failed: "Nextcloud Talk konnte nicht erstellt werden:\n$1\nBitte Optionen pruefen.",
  ui_moderator_transfer_failed: "Moderator konnte nicht uebertragen werden.",
  ui_moderator_transfer_failed_with_reason: "Moderator konnte nicht uebertragen werden:\n$1",
  ui_alert_title: "Nextcloud Talk",
  ui_alert_link_inserted: "Talk-Link eingefuegt:\n$1",
  ui_alert_location_missing: "(Hinweis: Feld 'Ort' wurde nicht automatisch gefunden.)",
  ui_alert_event_fallback: "Hinweis: Server unterstuetzt Event-Konversationen nicht. Es wurde ein Standard-Raum erstellt.",
  ui_alert_generic_fallback: "(Hinweis: Es wurde ein Fallback-Link ohne API erzeugt.)",
  ui_alert_reason: "Grund: $1",
  ui_alert_pending_delegation: "Moderator wird beim Speichern/Senden uebertragen an: $1.",
  ui_alert_delegation_done: "Moderator uebertragen an: $1.",
  ui_alert_delegation_removed: "Sie wurden aus der Unterhaltung entfernt.",
  ui_alert_password_protected: "Hinweis: Diese Unterhaltung ist passwortgeschuetzt.",
  ui_alert_lobby_state_active: "Lobby ist aktiv.",
  ui_alert_lobby_state_inactive: "Lobby ist deaktiviert.",
  ui_alert_lobby_no_rights: "(Hinweis: Keine Lobby-Rechte, Zustand unveraendert.)",
  ui_select_heading: "\u00d6ffentliche Unterhaltung auswaehlen",
  ui_select_search_placeholder: "Suche",
  ui_select_no_selection: "Keine Unterhaltung ausgewaehlt",
  ui_select_lobby_toggle: "Lobby einschalten (Startzeit aus Termin)",
  ui_select_missing_base_url: "Bitte hinterlegen Sie die Nextcloud URL in den Add-on-Optionen.",
  ui_select_status_loading: "Lade...",
  ui_select_status_none: "Keine Unterhaltungen gefunden.",
  ui_select_status_single: "1 Unterhaltung gefunden.",
  ui_select_status_many: "$1 Unterhaltungen gefunden.",
  ui_select_status_error: "Fehler: $1",
  ui_badge_password_required: "Passwort erforderlich",
  ui_badge_listed: "Oeffentlich gelistet",
  ui_badge_owned: "Eigene Teilnahme",
  ui_badge_guests_allowed: "Gaeste erlaubt",
  ui_badge_guests_forbidden: "Keine Gaeste erlaubt",
  ui_password_info_yes: "Passwortschutz aktiv.",
  ui_password_info_no: "Kein Passwort gesetzt.",
  ui_lobby_loading_details: "Lade Raumdetails...",
  ui_lobby_disabled: "Lobby deaktiviert.",
  ui_lobby_active_with_time: "Lobby aktiv (Start: $1)",
  ui_lobby_active_without_time: "Lobby aktiv (keine Startzeit gesetzt).",
  ui_lobby_fetching_status: "Lobby-Status wird ermittelt...",
  ui_details_error: "Details konnten nicht geladen werden.",
  ui_moderator_not_participant: "Keine Teilnahme an dieser Unterhaltung (nur Anzeige).",
  ui_moderator_can_manage: "Sie koennen die Lobby verwalten.",
  ui_moderator_cannot_manage: "Sie sind kein Moderator dieser Unterhaltung.",
  ui_lobby_no_permission: "Keine Berechtigung zum Aendern der Lobby.",
  ui_select_apply_permission_denied: "Lobby konnte nicht angepasst werden: Keine Moderatorrechte.",
  ui_select_apply_failed: "\u00dcbernahme fehlgeschlagen:\n$1",
  ui_select_rooms_load_failed: "\u00d6ffentliche Unterhaltungen konnten nicht geladen werden:\n$1",
  ui_prompt_base_url: "Nextcloud Basis-URL (z.B. https://cloud.example.com):",
  ui_prompt_title: "Titel fuer neue oeffentliche Unterhaltung:",
  ui_default_title: "Besprechung",
  ui_description_line_link: "\nTalk: $1",
  ui_description_line_password: "Passwort: $1",
  ui_toolbar_tooltip: "Nextcloud Talk",
  error_invalid_base64_length: "Ungueltige Base64-Laenge.",
  error_invalid_base64_data: "Ungueltige Base64-Daten.",
  error_avatar_load_failed: "Avatar konnte nicht geladen werden.",
  error_credentials_missing: "Nextcloud Zugang fehlt (URL/Nutzer/App-Pass).",
  error_ocs: "OCS-Fehler: $1",
  error_room_token_missing: "Raum-Token fehlt.",
  error_room_details_missing: "Raumdetails fehlen im Response.",
  error_lobby_update_failed: "Lobby-Update fehlgeschlagen (HTTP $1)."
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
  const resolvedSubs = Array.isArray(substitutions) ? substitutions : [substitutions];
  const ctxMessage = getLocalizedMessage(key, resolvedSubs);
  if (ctxMessage){
    return ctxMessage;
  }
  try{
    const api = getI18nApi();
    if (api){
      const message = api.getMessage(key, resolvedSubs);
      if (message){
        return message;
      }
    }
  }catch(_){}
  const fallback = I18N_FALLBACKS[key];
  if (fallback){
    return applySubstitutions(fallback, resolvedSubs);
  }
  if (resolvedSubs.length){
    return String(resolvedSubs[0] ?? "");
  }
  return "";
}

function localizedError(key, substitutions = []){
  const message = i18n(key, substitutions);
  return new Error(message || key);
}

let BROWSER = null;
let DEBUG = false;
let debugListenerRegistered = false;
let debugInitScheduled = false;

function getLocalizedMessage(key, substitutions){
  const ctx = LAST_CONTEXT;
  if (!ctx) return "";
  const options = ctx.cloneScope ? { cloneScope: ctx.cloneScope } : undefined;
  try{
    const ext = ctx.extension;
    if (ext?.localizeMessage){
      const message = ext.localizeMessage(key, substitutions, options || {});
      if (message){
        return message;
      }
    }
    const localeData = ext?.localeData;
    if (localeData?.localizeMessage){
      const message = localeData.localizeMessage(key, substitutions, options || {});
      if (message){
        return message;
      }
    }
  }catch(_){}
  try{
    if (typeof ctx.localizeMessage === "function"){
      const message = ctx.localizeMessage(key, substitutions, options || {});
      if (message){
        return message;
      }
    }
  }catch(_){}
  return "";
}

function getI18nApi(){
  try{
    if (typeof browser !== "undefined" && browser?.i18n){
      return browser.i18n;
    }
  }catch(_){}
  if (BROWSER?.i18n){
    return BROWSER.i18n;
  }
  if (LAST_CONTEXT){
    try{
      const resolved = resolveBrowser(LAST_CONTEXT);
      if (resolved?.i18n){
        return resolved.i18n;
      }
    }catch(_){}
  }
  return null;
}

function resolveBrowser(context){
  if (context){
    LAST_CONTEXT = context;
  }
  if (BROWSER) return BROWSER;
  if (context?.extension?.browser){
    BROWSER = context.extension.browser;
    return BROWSER;
  }
  if (context?.browser){
    BROWSER = context.browser;
    return BROWSER;
  }
  if (LAST_CONTEXT?.extension?.browser){
    BROWSER = LAST_CONTEXT.extension.browser;
    return BROWSER;
  }
  if (typeof browser !== "undefined"){
    BROWSER = browser;
    return BROWSER;
  }
  if (typeof globalThis !== "undefined" && globalThis && globalThis.browser){
    BROWSER = globalThis.browser;
    return BROWSER;
  }
  return null;
}

function ensureDebugState(context){
  const api = resolveBrowser(context);
  if (!api || !api.storage) return;
  if (!debugListenerRegistered && api.storage.onChanged){
    try{
      api.storage.onChanged.addListener((changes, area) => {
        if (area !== "local") return;
        if (Object.prototype.hasOwnProperty.call(changes, "debugEnabled")){
          DEBUG = !!changes.debugEnabled.newValue;
        }
      });
      debugListenerRegistered = true;
    }catch(_){}
  }
  if (debugInitScheduled) return;
  debugInitScheduled = true;
  try{
    const result = api.storage.local.get("debugEnabled");
    if (result && typeof result.then === "function"){
      result.then((stored) => {
        if (stored && Object.prototype.hasOwnProperty.call(stored, "debugEnabled")){
          DEBUG = !!stored.debugEnabled;
        }else{
          DEBUG = false;
        }
      }).catch(() => {});
    }
  }catch(_){}
}

const BASE64_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
const OBJECT_TO_STRING = Object.prototype.toString;

function isArrayBufferLike(value){
  return value && typeof value === "object" && OBJECT_TO_STRING.call(value) === "[object ArrayBuffer]";
}

function isUint8ClampedArrayLike(value){
  return value && typeof value === "object" && OBJECT_TO_STRING.call(value) === "[object Uint8ClampedArray]";
}

function toUint8ClampedArray(value){
  if (!value) return null;
  try{
    if (typeof Uint8ClampedArray !== "undefined" && value instanceof Uint8ClampedArray){
      return typeof value.slice === "function" ? value.slice() : new Uint8ClampedArray(value);
    }
  }catch(_){}
  if (isUint8ClampedArrayLike(value)){
    if (typeof value.slice === "function"){
      try{
        return value.slice();
      }catch(_){}
    }
    try{
      return new Uint8ClampedArray(value);
    }catch(_){}
  }
  if (typeof ArrayBuffer !== "undefined" && typeof ArrayBuffer.isView === "function" && ArrayBuffer.isView(value)){
    try{
      return new Uint8ClampedArray(value);
    }catch(_){}
  }
  if (isArrayBufferLike(value)){
    try{
      if (typeof value.slice === "function"){
        return new Uint8ClampedArray(value.slice(0));
      }
      return new Uint8ClampedArray(value);
    }catch(_){}
  }
  if (Array.isArray(value)){
    try{
      return new Uint8ClampedArray(value);
    }catch(_){}
  }
  if (typeof value.length === "number"){
    try{
      return new Uint8ClampedArray(Array.from(value));
    }catch(_){}
  }
  return null;
}

function manualAtob(input){
  const cleanBase = String(input ?? "").replace(/[\r\n\s]/g, "");
  if (!cleanBase){
    return "";
  }
  const sanitized = cleanBase.replace(/-/g, "+").replace(/_/g, "/");
  if (sanitized.length % 4 === 1){
    throw localizedError("error_invalid_base64_length");
  }
  let output = "";
  for (let i = 0; i < sanitized.length; i += 4){
    const char1 = sanitized.charAt(i);
    const char2 = sanitized.charAt(i + 1);
    const char3 = sanitized.charAt(i + 2) || "=";
    const char4 = sanitized.charAt(i + 3) || "=";
    const enc1 = BASE64_ALPHABET.indexOf(char1);
    const enc2 = BASE64_ALPHABET.indexOf(char2);
    const enc3 = char3 === "=" ? 64 : BASE64_ALPHABET.indexOf(char3);
    const enc4 = char4 === "=" ? 64 : BASE64_ALPHABET.indexOf(char4);
    if (enc1 < 0 || enc2 < 0 || (enc3 < 0 && enc3 !== 64) || (enc4 < 0 && enc4 !== 64)){
      throw localizedError("error_invalid_base64_data");
    }
    const chr1 = (enc1 << 2) | (enc2 >> 4);
    output += String.fromCharCode(chr1 & 0xff);
    if (enc3 !== 64){
      const chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
      output += String.fromCharCode(chr2 & 0xff);
    }
    if (enc4 !== 64){
      const chr3 = ((enc3 & 3) << 6) | enc4;
      output += String.fromCharCode(chr3 & 0xff);
    }
  }
  return output;
}

function safeAtob(input){
  ensureBrowserGlobals();
  if (typeof globalThis.atob === "function"){
    return globalThis.atob(String(input ?? ""));
  }
  return manualAtob(input);
}

function parseDataUrl(value){
  if (!value || typeof value !== "string") return null;
  const match = /^data:([^;,]+)?(?:;charset=[^;,]+)?(?:;base64)?,(.*)$/i.exec(value);
  if (!match) return null;
  return {
    mime: (match[1] || "").toLowerCase(),
    base64: match[2] || ""
  };
}

function getUrlFactory(win){
  if (win?.URL && typeof win.URL.createObjectURL === "function"){
    return win.URL;
  }
  if (typeof URL !== "undefined" && typeof URL.createObjectURL === "function"){
    return URL;
  }
  const hidden = getHiddenDOMWindow();
  if (hidden?.URL && typeof hidden.URL.createObjectURL === "function"){
    return hidden.URL;
  }
  return null;
}

function getBlobConstructor(){
  ensureBrowserGlobals();
  if (typeof Blob === "function"){
    return Blob;
  }
  if (typeof File === "function"){
    const Wrapper = function(parts, options){
      const opts = options && typeof options === "object" ? options : {};
      const name = typeof opts.name === "string" && opts.name.length ? opts.name : "blob";
      return new File(parts, name, opts);
    };
    return Wrapper;
  }
  const hidden = getHiddenDOMWindow();
  if (hidden && typeof hidden.Blob === "function"){
    return hidden.Blob;
  }
  if (hidden && typeof hidden.File === "function"){
    const Wrapper = function(parts, options){
      const opts = options && typeof options === "object" ? options : {};
      const name = typeof opts.name === "string" && opts.name.length ? opts.name : "blob";
      return new hidden.File(parts, name, opts);
    };
    return Wrapper;
  }
  return null;
}

async function decodeAvatarViaBackground(base64, mime){
  const clean = typeof base64 === "string" ? base64.replace(/\s+/g, "") : "";
  if (!clean){
    return null;
  }
  const payload = { base64: clean, mime: mime || "image/png" };
  if (LAST_CONTEXT){
    try{
      const runtimeResponse = await sendRuntimeMessage(LAST_CONTEXT, "talkMenu:decodeAvatar", payload);
      if (runtimeResponse && runtimeResponse.ok === false){
        log("background decode runtime error", runtimeResponse.error);
      }
      if (runtimeResponse && runtimeResponse.ok){
        return runtimeResponse;
      }
    }catch(e){
      err(e);
      err("background decode via runtime failed");
    }
    try{
      const utilityResponse = await requestUtility(LAST_CONTEXT, Object.assign({ type: "decodeAvatar" }, payload));
      if (utilityResponse && utilityResponse.ok === false){
        log("background decode utility error", utilityResponse.error);
      }
      if (utilityResponse && utilityResponse.ok){
        return utilityResponse;
      }
    }catch(e){
      err(e);
      err("background decode via utility failed");
    }
  }
  try{
    const runtime = (typeof browser !== "undefined" && browser?.runtime) ||
      (typeof globalThis !== "undefined" && globalThis.browser?.runtime) ||
      null;
    if (runtime?.sendMessage){
      const directResponse = await runtime.sendMessage(Object.assign({ type: "talkMenu:decodeAvatar" }, payload));
      if (directResponse && directResponse.ok === false){
        log("background decode direct error", directResponse.error);
      }
      if (directResponse && directResponse.ok){
        return directResponse;
      }
    }
  }catch(e){
    err(e);
    err("background decode direct messaging failed");
  }
  return null;
}

async function loadAvatarImage(dataUrl, win){
  if (!dataUrl) return null;
  if (AVATAR_BITMAP_CACHE.has(dataUrl)){
    return AVATAR_BITMAP_CACHE.get(dataUrl);
  }
  const task = (async () => {
    const parsed = parseDataUrl(dataUrl);
    if (!parsed || !parsed.base64){
      return null;
    }
    let cachedBackgroundEntry = null;
    const tryBackgroundDecode = async () => {
      if (cachedBackgroundEntry) return cachedBackgroundEntry;
      try{
        const decoded = await decodeAvatarViaBackground(parsed.base64, parsed.mime || "image/png");
        if (!decoded || decoded.ok === false){
          err(decoded?.error || "background decode returned empty");
          return null;
        }
        const width = decoded.width || decoded.imageWidth || 0;
        const height = decoded.height || decoded.imageHeight || 0;
        if (!width || !height){
          return null;
        }
        const pixelCandidates = [
          decoded.pixelData,
          decoded.data,
          decoded.array,
          decoded.pixels,
          decoded.bytes
        ];
        let pixels = null;
        for (const candidate of pixelCandidates){
          pixels = toUint8ClampedArray(candidate);
          if (pixels){
            break;
          }
        }
        if (!pixels){
          err("background decode produced no pixels");
          return null;
        }
        cachedBackgroundEntry = {
          width,
          height,
          pixelData: {
            width,
            height,
            data: pixels
          }
        };
        return cachedBackgroundEntry;
      }catch(e){
        err(e);
        return null;
      }
    };
    try{
      const binary = safeAtob(parsed.base64);
      const len = binary.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++){
        bytes[i] = binary.charCodeAt(i) & 0xff;
      }
      const mime = parsed.mime || "image/png";
      const BlobCtor = getBlobConstructor();
      let blob = null;
      if (BlobCtor){
        blob = new BlobCtor([bytes], { type: mime });
      }

      const globalBitmapFactory = (typeof globalThis !== "undefined" && typeof globalThis.createImageBitmap === "function")
        ? globalThis.createImageBitmap.bind(globalThis)
        : null;
      const hiddenWin = getHiddenDOMWindow();
      const hiddenBitmapFactory = hiddenWin && typeof hiddenWin.createImageBitmap === "function"
        ? hiddenWin.createImageBitmap.bind(hiddenWin)
        : null;
      const bitmapFactory =
        (win && typeof win.createImageBitmap === "function" && win.createImageBitmap.bind(win)) ||
        globalBitmapFactory ||
        hiddenBitmapFactory;
      if (bitmapFactory && blob){
        try{
          const bitmap = await bitmapFactory(blob);
          if (bitmap){
            return {
              image: bitmap,
              blob,
              width: bitmap.width || 0,
              height: bitmap.height || 0
            };
          }
        }catch(bitmapErr){
          err(bitmapErr);
        }
      }
      if (!blob){
        const bgEntry = await tryBackgroundDecode();
        if (bgEntry){
          return bgEntry;
        }
        return null;
      }
      const makeImage = () => {
        if (win && typeof win.Image === "function"){
          return new win.Image();
        }
        if (typeof Image === "function"){
          return new Image();
        }
        const hidden = getHiddenDOMWindow();
        if (hidden && typeof hidden.Image === "function"){
          return new hidden.Image();
        }
        return null;
      };
      const img = makeImage();
      if (!img){
        err("Image constructor unavailable");
        const bgEntry = await tryBackgroundDecode();
        if (bgEntry){
          return bgEntry;
        }
        return null;
      }
      const urlFactory = blob ? getUrlFactory(win) : null;
      const url = blob && urlFactory ? urlFactory.createObjectURL(blob) : null;
      if (!url){
        const bgEntry = await tryBackgroundDecode();
        if (bgEntry){
          return bgEntry;
        }
        return null;
      }
      const loadPromise = new Promise((resolve, reject) => {
        img.onload = () => resolve({
          image: img,
          blob,
          width: img.naturalWidth || img.width || 0,
          height: img.naturalHeight || img.height || 0,
          objectUrl: url,
          urlFactory
        });
        img.onerror = async (event) => {
          try{
            if (url && urlFactory?.revokeObjectURL){
              try { urlFactory.revokeObjectURL(url); }catch(_){}
            }
            const bgEntry = await tryBackgroundDecode();
            if (bgEntry){
              resolve(bgEntry);
              return;
            }
          }catch(errFallback){
            err(errFallback);
          }
          reject(event || localizedError("error_avatar_load_failed"));
        };
      });
      img.src = url;
      const result = await loadPromise;
      if (result.objectUrl && result.urlFactory?.revokeObjectURL){
        try{ result.urlFactory.revokeObjectURL(result.objectUrl); }catch(_){}
      }
      delete result.objectUrl;
      delete result.urlFactory;
      return result;
    }catch(e){
      err(e);
      const bgEntry = await tryBackgroundDecode();
      if (bgEntry){
        return bgEntry;
      }
      return null;
    }
  })();
  AVATAR_BITMAP_CACHE.set(dataUrl, task);
  return task;
}

async function drawAvatarOnCanvas(canvas, dataUrl){
  if (!canvas || !dataUrl) return false;
  try{
    const win = canvas.ownerDocument?.defaultView || (typeof window !== "undefined" ? window : null);
    const entry = await loadAvatarImage(dataUrl, win);
    if (!entry){
      err("avatar entry missing");
      return false;
    }
    const ctx = canvas.getContext("2d");
    if (!ctx){
      err("canvas context missing");
      return false;
    }
    let source = entry.image || null;
    let width = entry.width || 0;
    let height = entry.height || 0;
    if (entry.pixelData){
      width = entry.pixelData.width || width;
      height = entry.pixelData.height || height;
      const doc = canvas.ownerDocument || (typeof document !== "undefined" ? document : null);
      if (!doc || typeof doc.createElement !== "function"){
        err("no document to create temp canvas");
        return false;
      }
      const tempCanvas = doc.createElement("canvas");
      tempCanvas.width = width;
      tempCanvas.height = height;
      const tempCtx = tempCanvas.getContext("2d");
      if (!tempCtx){
        err("temp canvas context missing");
        return false;
      }
      const imageData = tempCtx.createImageData(width, height);
      imageData.data.set(entry.pixelData.data);
      tempCtx.putImageData(imageData, 0, 0);
      source = tempCanvas;
    }
    if (!source){
      source = entry.image;
    }
    if (!source){
      err("avatar source missing");
      return false;
    }
    if (!width || !height){
      width = entry.image?.naturalWidth || entry.image?.width || 0;
      height = entry.image?.naturalHeight || entry.image?.height || 0;
      if (!width || !height){
        err("avatar size missing");
        return false;
      }
    }
    const destW = canvas.width || 28;
    const destH = canvas.height || 28;
    ctx.clearRect(0, 0, destW, destH);
    let sx = 0;
    let sy = 0;
    let sWidth = width;
    let sHeight = height;
    if (width > height){
      sx = (width - height) / 2;
      sWidth = height;
    } else if (height > width){
      sy = (height - width) / 2;
      sHeight = width;
    }
    ctx.drawImage(source, sx, sy, sWidth, sHeight, 0, 0, destW, destH);
    return true;
  }catch(e){
      err(e);
      return false;
  }
}

function log(...a){
  if (!DEBUG) return;
  try { console.log("[NCExp]", ...a); } catch(_) {}
}
function err(e){ log("ERROR:", e && e.message ? e.message : String(e)); }

function getDocFromTarget(target){
  try{
    if (target && target.document) return target.document;
  }catch(_){}
  try{
    if (target && target.ownerDocument) return target.ownerDocument;
  }catch(_){}
  try{
    if (target && target.defaultView && target.defaultView.document) return target.defaultView.document;
  }catch(_){}
  try{
    if (target && target.parent && target.parent.document) return target.parent.document;
  }catch(_){}
  try{
    if (target && target.top && target.top.document) return target.top.document;
  }catch(_){}
  return (typeof document !== "undefined") ? document : null;
}

function showAlert(target, message, title){
  const resolvedTitle = title || i18n("ui_alert_title") || ALERT_TITLE_FALLBACK;
  const doc = getDocFromTarget(target);
  if (!doc || !doc.body){
    try{
      if (Services?.prompt){
        Services.prompt.alert(null, resolvedTitle, String(message));
        return;
      }
    }catch(_){}
    try{
      (target || window).alert(String(message));
    }catch(_){}
    return;
  }
  try{
    const existing = doc.getElementById("nctalk-alert");
    if (existing) existing.remove();
  }catch(_){}
  const overlay = doc.createElement("div");
  overlay.id = "nctalk-alert";
  Object.assign(overlay.style,{
    position:"fixed",
    inset:"0",
    background:"rgba(0,0,0,.35)",
    zIndex:"2147483647",
    display:"flex",
    alignItems:"center",
    justifyContent:"center",
    padding:"20px"
  });
  const panel = doc.createElement("div");
  Object.assign(panel.style,{
    background:"var(--arrowpanel-background,#fff)",
    border:"1px solid var(--arrowpanel-border-color,#c0c0c0)",
    borderRadius:"10px",
    boxShadow:"0 12px 36px rgba(0,0,0,.35)",
    maxWidth:"420px",
    width:"100%",
    padding:"18px 20px",
    display:"grid",
    gridTemplateColumns:"auto 1fr",
    gap:"14px 16px",
    font:"13px system-ui"
  });
  overlay.appendChild(panel);

  let iconUrl = null;
  try{
    if (LAST_CONTEXT) iconUrl = talkIconURL(LAST_CONTEXT, 48);
  }catch(_){}

  const iconBox = doc.createElement("div");
  Object.assign(iconBox.style,{display:"flex",alignItems:"flex-start"});
  if (iconUrl){
    const img = doc.createElement("img");
    img.src = iconUrl;
    img.alt = "";
    img.width = 40;
    img.height = 40;
    Object.assign(img.style,{borderRadius:"8px"});
    iconBox.appendChild(img);
  } else {
    const fallback = doc.createElement("div");
    Object.assign(fallback.style,{
      width:"40px",
      height:"40px",
      borderRadius:"8px",
      background:"linear-gradient(135deg,#0082c9,#00b6e8)"
    });
    iconBox.appendChild(fallback);
  }
  panel.appendChild(iconBox);

  const textBox = doc.createElement("div");
  Object.assign(textBox.style,{display:"flex",flexDirection:"column",gap:"6px"});
  const titleEl = doc.createElement("div");
  titleEl.textContent = resolvedTitle;
  Object.assign(titleEl.style,{fontWeight:"600",fontSize:"15px"});
  const msgEl = doc.createElement("div");
  msgEl.textContent = String(message ?? "");
  Object.assign(msgEl.style,{whiteSpace:"pre-wrap",lineHeight:"1.4"});
  textBox.appendChild(titleEl);
  textBox.appendChild(msgEl);
  panel.appendChild(textBox);

  const spacer = doc.createElement("div");
  spacer.style.gridColumn = "1 / -1";
  spacer.style.height = "4px";
  panel.appendChild(spacer);

  const buttonRow = doc.createElement("div");
  buttonRow.style.gridColumn = "1 / -1";
  Object.assign(buttonRow.style,{display:"flex",justifyContent:"flex-end"});
  const okBtn = doc.createElement("button");
  okBtn.textContent = i18n("ui_button_ok");
  Object.assign(okBtn.style,{padding:"6px 16px",borderRadius:"6px",border:"1px solid rgba(0,0,0,0.2)"});
  buttonRow.appendChild(okBtn);
  panel.appendChild(buttonRow);

  const close = () => {
    try { overlay.remove(); }catch(_){}
  };
  okBtn.addEventListener("click", close, { once:true });
  overlay.addEventListener("click", (ev) => { if (ev.target === overlay) close(); }, { once:false });

  const host = doc.body || doc.documentElement;
  host.appendChild(overlay);
}

function randToken(len=10){ const a="abcdefghijklmnopqrstuvwxyz0123456789"; let s=""; for(let i=0;i<len;i++) s+=a[Math.floor(Math.random()*a.length)]; return s; }

async function getCredentials(context){
  const keys = ["baseUrl","user","appPass"];
  const clean = (value) => (value == null ? "" : String(value).trim());
  const merge = (target, source) => {
    if (!source) return target;
    for (const key of keys){
      if (source[key] != null && source[key] !== "" && !target[key]){
        target[key] = clean(source[key]);
      }
    }
    return target;
  };
  let out = {};
  try{
    const storage = context?.extension?.browser?.storage?.local;
    if (storage?.get){
      out = merge(out, await storage.get(keys));
    }
  }catch(e){
      err(e);
      }
  try{
    const globalBrowser =
      (typeof browser !== "undefined" && browser?.storage?.local?.get && browser) ||
      (typeof globalThis !== "undefined" && globalThis.browser?.storage?.local?.get && globalThis.browser) ||
      null;
    if (globalBrowser){
      out = merge(out, await globalBrowser.storage.local.get(keys));
    }
  }catch(e){
      err(e);
      }
  return {
    baseUrl: out.baseUrl ? out.baseUrl.replace(/\/+$/,"") : "",
    user: out.user || "",
    appPass: out.appPass || ""
  };
}

async function getBaseUrl(context){
  try{
    const creds = await getCredentials(context);
    if (creds.baseUrl) return creds.baseUrl;
  }catch(e){
      err(e);
      }
  try{
    const res = await requestUtility(context, { type: "getConfig" });
    const base = res?.ok ? res?.config?.baseUrl : null;
    if (base) return String(base).replace(/\/$/,"");
  }catch(e){
      err(e);
      }
  try{
    const res = await sendRuntimeMessage(context, "talkMenu:getConfig", {});
    const base = res?.ok ? res?.config?.baseUrl : null;
    if (base) return String(base).replace(/\/$/,"");
  }catch(e){
      err(e);
      }
  return null;
}

function isEventDialog(win) {
  try { return win?.document?.documentElement?.getAttribute("windowtype") === "Calendar:EventDialog"; } catch(e){ return false; }
}

function findBar(doc){ return doc.querySelector(".calendar-dialog-toolbar, .dialog-buttons, toolbar") || doc.body; }

function talkIconURL(context, size=20){
  try { return context.extension.rootURI.spec + "icons/talk-" + size + ".png"; } catch(_) { return null; }
}

function getCalendarItem(doc){
  try{
    const win = doc?.defaultView || window;
    if (!win) return null;
    if (win.calendarItem) return win.calendarItem;
    if (win.gEvent && win.gEvent.event) return win.gEvent.event;
    if (win.arguments && win.arguments[0]){
      const arg = win.arguments[0];
      if (arg.calendarItem) return arg.calendarItem;
      if (arg.calendarEvent) return arg.calendarEvent;
    }
  }catch(_){}
  return null;
}

function hashStringToHex(value){
  const input = String(value ?? "");
  let hash = 0;
  for (let i = 0; i < input.length; i++){
    hash = ((hash << 5) - hash) + input.charCodeAt(i);
    hash |= 0;
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function buildEventObjectMetadata(doc, { title, startTimestamp } = {}){
  const candidates = [];
  const pushCandidate = (value) => {
    if (!value) return;
    const str = String(value).trim();
    if (str) candidates.push(str);
  };
  const item = getCalendarItem(doc);
  pushCandidate(item?.id);
  pushCandidate(item?.uid);
  pushCandidate(item?.hashId);
  if (item && typeof item.getProperty === "function"){
    try{ pushCandidate(item.getProperty("uid")); }catch(_){}
  }
  if (item?.icalComponent){
    try{ pushCandidate(item.icalComponent.uid); }catch(_){}
  }
  let objectId = candidates.find(Boolean);
  let fallback = false;
  if (!objectId){
    fallback = true;
    const seedParts = [
      startTimestamp != null ? String(startTimestamp) : "",
      title || "",
      Date.now().toString()
    ];
    objectId = "tb-" + hashStringToHex(seedParts.join("|"));
  }
  return {
    objectType: "event",
    objectId,
    fallback
  };
}

const PENDING_DELEGATION_KEY = "_nctalkPendingModerator";

function queuePendingModerator(win, data){
  if (!win) return;
  if (!data || !data.token || !data.delegateId){
    if (win[PENDING_DELEGATION_KEY]){
      log("pending delegation cleared", { reason: "invalid data" });
    }
    delete win[PENDING_DELEGATION_KEY];
    return;
  }
  win[PENDING_DELEGATION_KEY] = {
    token: data.token,
    delegateId: data.delegateId,
    displayName: data.displayName || data.delegateId,
    processed: false
  };
  log("pending delegation stored", {
    token: shortToken(data.token),
    delegateId: data.delegateId,
    displayName: data.displayName || data.delegateId
  });
}

function getPendingModerator(win){
  if (!win) return null;
  const pending = win[PENDING_DELEGATION_KEY];
  if (!pending || !pending.token || !pending.delegateId) return null;
  return pending;
}

function clearPendingModerator(win){
  if (win && win[PENDING_DELEGATION_KEY]){
    log("pending delegation cleared", {
      token: shortToken(win[PENDING_DELEGATION_KEY].token),
      delegateId: win[PENDING_DELEGATION_KEY].delegateId
    });
    delete win[PENDING_DELEGATION_KEY];
  }
}

/**
 * Beobachtet Terminänderungen (Startzeit, Buttons) und synchronisiert die Lobby.
 * @param {*} context - Add-on Kontext.
 * @param {Document} rootDoc - Hauptdokument des Terminfensters.
 * @param {Document} innerDoc - Optionales iframe-Dokument.
 * @param {string} token - Talk-Raum-Token.
 * @param {boolean} enableLobby - Soll die Lobby aktiv bleiben?
 */
function setupLobbyWatcher(context, rootDoc, innerDoc, token, enableLobby){
  try{
    const win = rootDoc?.defaultView || window;
    if (!win) return;
    const key = "_nctalkLobbyWatcher";
    const state = win[key] || (win[key] = { token:null, lastTs:null, debounce:null, cleanup:[] });
    const cleanupAll = () => {
      try{
        if (state.cleanup && state.cleanup.length){
          for (const fn of state.cleanup){
            try { fn(); } catch(_) {}
          }
        }
      }catch(_){}
      state.cleanup = [];
      if (state.debounce){
        try { win.clearTimeout(state.debounce); }catch(_){}
      }
      state.debounce = null;
      state.token = null;
      state.lastTs = null;
    };
    if (!enableLobby || !token){
      cleanupAll();
      return;
    }
    cleanupAll();
    state.token = token;
    const targetDoc = innerDoc || rootDoc;
    try {
      const current = extractStartTimestamp(targetDoc);
      if (current) state.lastTs = current;
    } catch(_) {}
    const runUpdate = async (forced = false) => {
      try{
        const ts = extractStartTimestamp(targetDoc);
        if (!ts) return;
        if (!forced && ts === state.lastTs) return;
        state.lastTs = ts;
        const payload = {
          token,
          enableLobby: true,
          startTimestamp: ts
        };
        await performLobbyUpdate(context, payload);
        log("lobby watcher update", { token: shortToken(token), startTimestamp: ts, forced });
        if (!forced) return;
        const pending = getPendingModerator(win);
        if (!pending || pending.processed || pending.token !== token) return;
        try{
          const delegateResponse = await requestUtility(context, { type: "delegateModerator", token, newModerator: pending.delegateId });
          if (!delegateResponse || !delegateResponse.ok){
            throw new Error(delegateResponse?.error || i18n("ui_moderator_transfer_failed"));
          }
          const delegationResult = delegateResponse.result || { delegate: pending.delegateId };
          pending.processed = true;
          log("pending delegation applied", {
            token: shortToken(token),
            delegate: delegationResult.delegate || pending.delegateId,
            leftSelf: !!delegationResult.leftSelf
          });
          const delegateName = pending.displayName || delegationResult.delegate || pending.delegateId;
          const delegationLines = [i18n("ui_alert_delegation_done", [delegateName])];
          if (delegationResult.leftSelf){
            delegationLines.push(i18n("ui_alert_delegation_removed"));
          }
          showAlert(win, delegationLines.join("\n"));
          clearPendingModerator(win);
          if (delegationResult.leftSelf){
            cleanupAll();
          }
        }catch(delegationErr){
          pending.processed = true;
          log("pending delegation failed", {
            token: shortToken(token),
            delegateId: pending.delegateId,
            error: delegationErr?.message || String(delegationErr)
          });
          showAlert(win, i18n("ui_moderator_transfer_failed_with_reason", [delegationErr?.message || delegationErr]));
          clearPendingModerator(win);
        }
      }catch(e){
        err(e);
        err("lobby watcher update failed");
      }
    };
    const handler = (forced = false) => {
      if (state.debounce){
        try { win.clearTimeout(state.debounce); }catch(_){}
        state.debounce = null;
      }
      if (forced){
        runUpdate(true);
        return;
      }
      state.debounce = win.setTimeout(() => {
        runUpdate(false).catch((e) => {
          err(e);
          err("lobby watcher update failed");
        });
      }, 400);
    };
    const trigger = (ev) => {
      const type = ev?.type || "mutation";
      const forced = type === "dialogaccept" || type === "dialogextra1";
      if (forced) state.lastTs = null;
      handler(forced);
    };
  const docs = collectEventDocs(targetDoc);
  const selectors = [
    "datetimepicker#event-starttime", "datetimepicker#item-starttime",
      'input#event-starttime', 'input#item-starttime', 'input[data-type="start-time"]',
      'input.start-time', 'timepicker#event-starttime', 'timepicker#item-starttime',
      'input#event-startdate', 'input#item-startdate', 'input[type="date"]', 'input.start-date', 'datepicker#event-startdate'
    ];
    for (const d of docs){
      for (const sel of selectors){
        const el = d.querySelector && d.querySelector(sel);
        if (!el) continue;
        try{
          el.addEventListener("change", trigger, true);
          state.cleanup.push(() => el.removeEventListener("change", trigger, true));
        }catch(_){}
        try{
          el.addEventListener("input", trigger, true);
          state.cleanup.push(() => el.removeEventListener("input", trigger, true));
        }catch(_){}
        try {
          const obs = new MutationObserver(() => trigger({ type: "mutation" }));
          obs.observe(el, { attributes: true, attributeFilter: ["value"] });
          state.cleanup.push(() => obs.disconnect());
        }catch(_){}
      }
    }
    const winEvents = ["dialogaccept","dialogextra1","DOMContentLoaded","unload"];
    for (const evt of winEvents){
      const listener = trigger.bind(null,{type:evt});
      try {
        win.addEventListener(evt, listener, true);
        state.cleanup.push(() => win.removeEventListener(evt, listener, true));
      }catch(_){}
    }
    const pollInterval = win.setInterval(() => trigger({ type: "poll" }), 4000);
    state.cleanup.push(() => win.clearInterval(pollInterval));
    handler(false);
  }catch(e){
      err(e);
      }
}

/**
 * Aktualisiert den Lobby-Status eines Raums. Behandelt Messaging-Fallbacks.
 */
async function performLobbyUpdate(context, payload){
  log("performLobbyUpdate", summarizeUtilityPayload(payload));
  const isPermissionError = (error) => {
    if (!error) return false;
    const msg = String(error.message || error).toLowerCase();
    return msg.includes("berechtigung") || msg.includes("403");
  };
  const handleError = (error) => {
    if (isPermissionError(error)) throw error;
    err(error);
  };
  try{
    const response = await requestLobbyUpdate(context, payload);
    if (response && response.ok === false){
      const errMsg = response.error || "Lobby-Update fehlgeschlagen.";
      if (isPermissionError({ message: errMsg })) throw new Error(errMsg);
      throw new Error(errMsg);
    }
    if (response) return;
  }catch(updateErr){
    handleError(updateErr);
  }
  try{
    const runtimeResponse = await sendRuntimeMessage(context, "talkMenu:updateLobby", payload);
    if (runtimeResponse && runtimeResponse.ok === false){
      const errMsg = runtimeResponse.error || "Lobby-Update fehlgeschlagen.";
      if (isPermissionError({ message: errMsg })) throw new Error(errMsg);
      throw new Error(errMsg);
    }
    if (runtimeResponse) return;
  }catch(runtimeErr){
    handleError(runtimeErr);
  }
  try{
    await updateLobbyDirect(context, payload);
  }catch(directErr){
    if (isPermissionError(directErr)){
      throw directErr;
    }
    err(directErr);
    throw directErr;
  }
}

async function sendRuntimeMessage(context, type, payload){
  let runtime = context?.extension?.browser?.runtime;
  if (!runtime && context?.extension?.runtime) runtime = context.extension.runtime;
  if (!runtime && typeof browser !== "undefined" && browser?.runtime) runtime = browser.runtime;
  if (!runtime && typeof globalThis !== "undefined" && globalThis.browser?.runtime) runtime = globalThis.browser.runtime;
  if (!runtime?.sendMessage){
    return null;
  }
  try{
    return await runtime.sendMessage({ type, payload });
  }catch(e){
      err(e);
      return null;
  }
}

async function listPublicRoomsDirect(context, searchTerm = ""){
  const { baseUrl, user, appPass } = await getCredentials(context);
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const url = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/listed-room?searchTerm=" + encodeURIComponent(searchTerm || "");
  const res = await fetch(url, { method:"GET", headers });
  const raw = await res.text();
  let data = null;
  try { data = raw ? JSON.parse(raw) : null; } catch(_){}
  if (!res.ok){
    const meta = data?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw localizedError("error_ocs", [detail]);
  }
  const rooms = data?.ocs?.data;
  return Array.isArray(rooms) ? rooms : [];
}

async function getRoomInfoDirect(context, token){
  if (!token) throw localizedError("error_room_token_missing");
  const { baseUrl, user, appPass } = await getCredentials(context);
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const url = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token);
  const res = await fetch(url, { method:"GET", headers });
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
  return room;
}

async function updateLobbyDirect(context, { token, enableLobby, startTimestamp } = {}){
  if (!token) throw localizedError("error_room_token_missing");
  const { baseUrl, user, appPass } = await getCredentials(context);
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json", "Content-Type":"application/json" };
  const url = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + token + "/webinar/lobby";
  const payload = { state: enableLobby ? 1 : 0 };
  if (enableLobby && typeof startTimestamp === "number" && Number.isFinite(startTimestamp) && startTimestamp > 0){
    let timerVal = startTimestamp;
    if (timerVal > 1e12) timerVal = Math.floor(timerVal / 1000);
    payload.timer = Math.floor(timerVal);
  }
  if (!enableLobby) delete payload.timer;
  const res = await fetch(url, { method:"PUT", headers, body: JSON.stringify(payload) });
  if (!res.ok){
    throw localizedError("error_lobby_update_failed", [res.status]);
  }
}

async function fetchListedRooms(context, searchTerm = ""){
  log("fetchListedRooms", { searchTerm });
  try{
    const res = await requestUtility(context, { type: "listPublicRooms", searchTerm });
    if (res){
      if (res.ok){
        const rooms = Array.isArray(res.rooms) ? res.rooms : [];
        log("fetchListedRooms response", { source: "utility", count: rooms.length });
        return rooms;
      }
      throw new Error(res.error || "Liste der Unterhaltungen nicht verf\\u00fcgbar.");
    }
  }catch(e){
      err(e);
      }
  try{
    const res = await sendRuntimeMessage(context, "talkMenu:listPublicRooms", { searchTerm });
    if (res){
      if (res.ok){
        const rooms = Array.isArray(res.rooms) ? res.rooms : [];
        log("fetchListedRooms response", { source: "runtime", count: rooms.length });
        return rooms;
      }
      throw new Error(res.error || "Liste der Unterhaltungen nicht verfuegbar.");
    }
  }catch(e){
      err(e);
      }
  const directRooms = await listPublicRoomsDirect(context, searchTerm);
  log("fetchListedRooms response", { source: "direct", count: Array.isArray(directRooms) ? directRooms.length : 0 });
  return directRooms;
}

async function fetchRoomDetails(context, token){
  log("fetchRoomDetails", { token: shortToken(token) });
  if (!token) throw localizedError("error_room_token_missing");
  try{
    const res = await requestUtility(context, { type: "getRoomInfo", token });
    if (res){
      if (res.ok && res.room){
        log("fetchRoomDetails response", { source: "utility", success: true });
        return res.room;
      }
      throw new Error(res.error || "Raumdetails konnten nicht geladen werden.");
    }
  }catch(e){
      err(e);
      }
  try{
    const res = await sendRuntimeMessage(context, "talkMenu:getRoomInfo", { token });
    if (res){
      if (res.ok && res.room){
        log("fetchRoomDetails response", { source: "runtime", success: true });
        return res.room;
      }
      throw new Error(res.error || "Raumdetails konnten nicht geladen werden.");
    }
  }catch(e){
      err(e);
      }
  const directRoom = await getRoomInfoDirect(context, token);
  log("fetchRoomDetails response", { source: "direct", success: !!directRoom });
  return directRoom;
}

function collectEventDocs(doc){
  const docs = [];
  try { if (doc) docs.push(doc); } catch(_){}
  try {
    const iframeDoc = doc.getElementById && doc.getElementById("calendar-item-panel-iframe")?.contentDocument;
    if (iframeDoc && iframeDoc !== doc) docs.push(iframeDoc);
  } catch(_){}
  return docs;
}

function extractDescriptionText(doc){
  try{
    const item = getCalendarItem(doc);
    if (item){
      let value = null;
      try { value = item.getProperty?.("description"); } catch(_){}
      if (value == null && "description" in item) value = item.description;
      if (value != null) return String(value);
    }
  }catch(_){}
  const docs = collectEventDocs(doc);
  for (const d of docs){
    try{
      const host = d.querySelector && d.querySelector("editor#item-description");
      let target = null;
      if (host) target = host.inputField || host.contentDocument?.body || host;
      if (!target){
        const fallbacks = [ "textarea#item-description", "textarea", "[contenteditable='true']", "div[role='textbox']" ];
        for (const sel of fallbacks){
          const el = d.querySelector && d.querySelector(sel);
          if (el){ target = el; break; }
        }
      }
      if (!target) continue;
      let text = "";
      if ("value" in target) text = target.value;
      else if (target.innerText) text = target.innerText;
      else text = target.textContent;
      if (text) return String(text).trim();
    }catch(_){}
  }
  return "";
}

function extractStartTimestamp(doc){
  try{
    const docs = collectEventDocs(doc);
    for (const d of docs){
      const picker = d.querySelector && (d.querySelector("datetimepicker#event-starttime") || d.querySelector("datetimepicker#item-starttime"));
      if (picker){
        try{
          let value = picker.value || (picker.getAttribute && picker.getAttribute("value"));
          if (!value && "valueAsDate" in picker) value = picker.valueAsDate;
          if (value){
            if (value instanceof Date || (typeof value === "object" && typeof value.getTime === "function")){
              const ts = value.getTime();
              if (!Number.isNaN(ts)) return Math.floor(ts / 1000);
            } else if (typeof value === "string"){
              const parsed = Date.parse(value);
              if (!Number.isNaN(parsed)) return Math.floor(parsed / 1000);
            }
          }
        }catch(_){}
      }
    }
  }catch(_){}
  try{
    const docs = collectEventDocs(doc);
    const selectors = ['input#event-starttime', 'input#item-starttime', 'input[data-type="start-time"]', 'input.start-time'];
    for (const d of docs){
      for (const sel of selectors){
        const el = d.querySelector && d.querySelector(sel);
        if (el && el.value){
          let dateInput = d.querySelector('input[type="date"]') || d.querySelector('input.start-date');
          let dateStr = dateInput && dateInput.value;
          const timeStr = el.value;
          if (!dateStr){
            const dateField = d.querySelector('input#event-startdate') || d.querySelector('input#item-startdate');
            if (dateField) dateStr = dateField.value;
          }
          if (dateStr && timeStr){
            const combined = new Date(dateStr + "T" + timeStr);
            if (!isNaN(combined.getTime())) return Math.floor(combined.getTime() / 1000);
          }
        }
      }
    }
  }catch(_){}
  try{
    const item = getCalendarItem(doc);
    const start = item?.startDate || item?.untilDate || item?.endDate;
    if (start){
      const jsDate = start.jsDate || start.getInTimezone?.(start.timezone || start.timezoneProvider?.tzid || "UTC")?.jsDate;
      if (jsDate && jsDate.getTime){
        const t = jsDate.getTime();
        if (!Number.isNaN(t)) return Math.floor(t / 1000);
      }
      if ("nativeTime" in start && typeof start.nativeTime === "number"){
        const native = start.nativeTime;
        if (native > 1e12) return Math.floor(native / 1e6);
        if (native > 1e9) return Math.floor(native / 1000);
        return Math.floor(native);
      }
    }
  }catch(_){}
  try{
    const win = doc.defaultView || window;
    const gStart = win?.gStartDateTime;
    if (gStart && typeof gStart === "object"){
      const jsDate = gStart.jsDate || gStart.getInTimezone?.("UTC")?.jsDate;
      if (jsDate && jsDate.getTime){
        const t = jsDate.getTime();
        if (!Number.isNaN(t)) return Math.floor(t / 1000);
      }
    }
  }catch(_){}
  return null;
}
const CREATE_HANDLERS = new WeakMap();
const LOBBY_HANDLERS = new WeakMap();
const UTILITY_HANDLERS = new WeakMap();

function getCreateHandlers(context){
  let set = CREATE_HANDLERS.get(context);
  if (!set){
    set = new Set();
    CREATE_HANDLERS.set(context, set);
  }
  return set;
}

function getLobbyHandlers(context){
  let set = LOBBY_HANDLERS.get(context);
  if (!set){
    set = new Set();
    LOBBY_HANDLERS.set(context, set);
  }
  return set;
}

function getUtilityHandlers(context){
  let set = UTILITY_HANDLERS.get(context);
  if (!set){
    set = new Set();
    UTILITY_HANDLERS.set(context, set);
  }
  return set;
}

async function requestCreateFromExtension(context, payload){
  const handlers = CREATE_HANDLERS.get(context);
  if (!handlers || handlers.size === 0) return null;
  for (const fire of handlers){
    try{
      const result = await fire.async(payload);
      if (result !== undefined) return result;
    }catch(e){
      err(e);
      }
  }
  return null;
}

async function fetchRoomParticipants(context, token){
  log("fetchRoomParticipants", { token: shortToken(token) });
  try{
    const res = await requestUtility(context, { type: "getRoomParticipants", token });
    if (res){
      if (res.ok) return Array.isArray(res.participants) ? res.participants : [];
      throw new Error(res.error || "Teilnehmer konnten nicht geladen werden.");
    }
  }catch(e){
      err(e);
      }
  return await getRoomParticipantsDirect(context, token);
}

async function requestLobbyUpdate(context, payload){
  log("requestLobbyUpdate", summarizeUtilityPayload(payload));
  const handlers = LOBBY_HANDLERS.get(context);
  if (!handlers || handlers.size === 0) return null;
  for (const fire of handlers){
    try{
      const result = await fire.async(payload);
      if (result !== undefined) return result;
    }catch(e){
      err(e);
      }
  }
  return null;
}
/**
 * Registriert den Direkt-Handler für den Toolbar-Button.
 * Da die Direkt-Version kein Menü mehr anbietet, genügt ein einfacher Click-Listener.
 */
function ensureMenu(doc, context, anchor){
  if (anchor.dataset.nctalkDirect === "1") return;
  anchor.dataset.nctalkDirect = "1";
  anchor.addEventListener("click", async (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (DEBUG) log("toolbar button clicked", { url: doc?.URL || "" });
    try{
      await openCreateDialog(doc, context);
    }catch(e){
      err(e);
      }
  });
}
/**
 * Erzeugt den Button, der im Termin-Dialog angezeigt wird.
 */
function buildButton(doc, context, label, tooltip){
  const btn = doc.createElement("button");
  btn.id = "nctalk-min-btn";
  btn.type = "button";
  btn.title = tooltip || "Nextcloud Talk";
  Object.assign(btn.style, {display:"inline-flex",alignItems:"center",gap:"6px",padding:"3px 10px",
    marginInlineStart:"8px",marginInlineEnd:"0"});
  const img = doc.createElement("img"); img.alt=""; img.width=20; img.height=20;
  const src = talkIconURL(context, 20); if (src) img.src = src;
  const span = doc.createElement("span"); span.textContent = label || i18n("ui_insert_button_label");
  btn.appendChild(img); btn.appendChild(span);
  return btn;
}

/**
 * Öffnet den Dialog zum Erstellen einer öffentlichen Talk-Unterhaltung.
 * Verarbeitet Eingaben, führt API-Aufrufe aus und übernimmt den Ergebnis-Link in den Termin.
 */
async function openCreateDialog(doc, context){
  try{
    const docWin = doc.defaultView || window;
    const defaultTitle = doc.querySelector('input[type="text"]')?.value || i18n("ui_default_title");
    log("openCreateDialog", {
      defaultTitle,
      url: doc?.URL || ""
    });
    const overlay = doc.createElement("div");
    Object.assign(overlay.style,{
      position:"fixed",
      inset:"0",
      background:"rgba(0,0,0,.25)",
      zIndex:"2147483646",
      display:"flex",
      alignItems:"center",
      justifyContent:"center",
      padding:"40px"
    });
    const panel = doc.createElement("div");
    Object.assign(panel.style,{
      position:"relative",
      background:"var(--arrowpanel-background,#fff)",
      border:"1px solid var(--arrowpanel-border-color,#c0c0c0)",
      borderRadius:"8px",
      boxShadow:"0 10px 30px rgba(0,0,0,.25)",
      minWidth:"480px",
      maxWidth:"640px",
      minHeight:"560px",
      maxHeight:"calc(100vh - 40px)",
      overflowY:"auto",
      padding:"20px",
      zIndex:"2147483647"
    });

    const heading = doc.createElement("h2");
    heading.textContent = i18n("ui_create_heading");
    Object.assign(heading.style,{margin:"0 0 10px",font:"600 16px system-ui"});

    const grid = doc.createElement("div");
    Object.assign(grid.style,{display:"grid",gridTemplateColumns:"160px 1fr",gap:"10px",alignItems:"center"});

    const titleLabel = doc.createElement("label");
    titleLabel.textContent = i18n("ui_create_title_label");
    const titleInput = doc.createElement("input");
    Object.assign(titleInput,{id:"nc_title",type:"text",value:defaultTitle});

    const passLabel = doc.createElement("label");
    passLabel.textContent = i18n("ui_create_password_label");
    const passInput = doc.createElement("input");
    Object.assign(passInput,{id:"nc_pass",type:"password",placeholder: i18n("ui_create_password_placeholder")});

    const lobbyLabel = doc.createElement("label");
    const lobbyCheckbox = doc.createElement("input");
    Object.assign(lobbyCheckbox,{id:"nc_lobby",type:"checkbox"});
    lobbyCheckbox.checked = true;
    lobbyLabel.appendChild(lobbyCheckbox);
    lobbyLabel.appendChild(doc.createTextNode(" " + i18n("ui_create_lobby_label")));
    const lobbySpacer = doc.createElement("div");

    const listableLabel = doc.createElement("label");
    const listableCheckbox = doc.createElement("input");
    Object.assign(listableCheckbox,{id:"nc_listable",type:"checkbox"});
    listableCheckbox.checked = true;
    listableLabel.appendChild(listableCheckbox);
    listableLabel.appendChild(doc.createTextNode(" " + i18n("ui_create_listable_label")));
    const listableSpacer = doc.createElement("div");

    const modeLabel = doc.createElement("label");
    modeLabel.textContent = i18n("ui_create_roomtype_label");
    modeLabel.style.alignSelf = "start";
    const modeField = doc.createElement("div");
    Object.assign(modeField.style,{display:"flex",flexDirection:"column",gap:"4px"});
    const roomModeName = "nc_room_mode";
    const modeEventOption = doc.createElement("label");
    Object.assign(modeEventOption.style,{display:"flex",alignItems:"center",gap:"6px"});
    const modeEventRadio = doc.createElement("input");
    Object.assign(modeEventRadio,{type:"radio",name:roomModeName,value:"event"});
    modeEventRadio.checked = true;
    const modeEventText = doc.createElement("span");
    modeEventText.textContent = i18n("ui_create_mode_event");
    modeEventOption.appendChild(modeEventRadio);
    modeEventOption.appendChild(modeEventText);
    const modeStandardOption = doc.createElement("label");
    Object.assign(modeStandardOption.style,{display:"flex",alignItems:"center",gap:"6px"});
    const modeStandardRadio = doc.createElement("input");
    Object.assign(modeStandardRadio,{type:"radio",name:roomModeName,value:"standard"});
    const modeStandardText = doc.createElement("span");
    modeStandardText.textContent = i18n("ui_create_mode_standard");
    modeStandardOption.appendChild(modeStandardRadio);
    modeStandardOption.appendChild(modeStandardText);
    const modeStatus = doc.createElement("div");
    Object.assign(modeStatus.style,{fontSize:"11px",opacity:"0.7"});
    modeStatus.textContent = "";
    modeField.appendChild(modeEventOption);
    modeField.appendChild(modeStandardOption);
    modeField.appendChild(modeStatus);

    const applyEventModeSupport = ({ supported, reason } = {}) => {
      const note = reason ? String(reason) : "";
      if (supported === false){
        modeEventRadio.disabled = true;
        modeEventRadio.checked = false;
        modeStandardRadio.checked = true;
        modeStatus.textContent = note || i18n("ui_create_mode_unsupported");
      } else {
        modeEventRadio.disabled = false;
        if (!modeStandardRadio.checked && !modeEventRadio.checked){
          modeEventRadio.checked = true;
        }
        modeStatus.textContent = "";
      }
    };
    applyEventModeSupport({ supported:null, reason:"" });

    const delegateLabel = doc.createElement("label");
    delegateLabel.textContent = i18n("ui_create_moderator_label");
    delegateLabel.style.alignSelf = "start";
    const delegateField = doc.createElement("div");
    Object.assign(delegateField.style,{display:"flex",flexDirection:"column",gap:"6px",position:"relative"});
    const delegateInput = doc.createElement("input");
    Object.assign(delegateInput,{id:"nc_delegate",type:"text",placeholder: i18n("ui_create_moderator_placeholder")});
    delegateInput.dataset.selectionLabel = "";
    delegateInput.dataset.selectionAvatar = "";
    delegateInput.dataset.selectionInitials = "";
    delegateInput.autocomplete = "off";
    delegateInput.spellcheck = false;
    const delegateInputRow = doc.createElement("div");
    Object.assign(delegateInputRow.style,{display:"flex",alignItems:"center",gap:"6px"});
    delegateInputRow.appendChild(delegateInput);
    const delegateClearBtn = doc.createElement("button");
    delegateClearBtn.type = "button";
    delegateClearBtn.textContent = i18n("ui_button_clear");
    Object.assign(delegateClearBtn.style,{padding:"4px 10px",fontSize:"12px"});
    delegateInputRow.appendChild(delegateClearBtn);
    delegateField.appendChild(delegateInputRow);
    const delegateStatus = doc.createElement("div");
    Object.assign(delegateStatus.style,{fontSize:"11px",opacity:"0.7",minHeight:"14px"});
    delegateField.appendChild(delegateStatus);
    const delegateSelectedInfo = doc.createElement("div");
    Object.assign(delegateSelectedInfo.style,{
      display:"none",
      alignItems:"center",
      gap:"10px",
      marginTop:"6px",
      fontSize:"11px",
      color:"var(--arrowpanel-color,#1b1b1b)",
      minHeight:"42px",
      padding:"8px 12px",
      borderRadius:"10px",
      background:"var(--arrowpanel-dimmed, rgba(0,0,0,0.04))"
    });
    const delegateSelectedAvatar = doc.createElement("div");
    Object.assign(delegateSelectedAvatar.style,{
      width:"36px",
      height:"36px",
      borderRadius:"50%",
      background:"var(--toolbarbutton-hover-background, rgba(0,0,0,0.08))",
      display:"none",
      alignItems:"center",
      justifyContent:"center",
      overflow:"hidden",
      flex:"0 0 36px",
      position:"relative",
      boxShadow:"0 0 0 1px rgba(0,0,0,0.06)"
    });
    const delegateSelectedAvatarCanvas = doc.createElement("canvas");
    delegateSelectedAvatarCanvas.width = 36;
    delegateSelectedAvatarCanvas.height = 36;
    Object.assign(delegateSelectedAvatarCanvas.style,{
      width:"100%",
      height:"100%",
      display:"none",
      borderRadius:"50%"
    });
    const delegateSelectedAvatarInitials = doc.createElement("span");
    Object.assign(delegateSelectedAvatarInitials.style,{
      width:"100%",
      height:"100%",
      alignItems:"center",
      justifyContent:"center",
      fontWeight:"600",
      fontSize:"13px",
      letterSpacing:"0.02em",
      color:"var(--arrowpanel-color,#1b1b1b)",
      display:"none"
    });
    delegateSelectedAvatar.appendChild(delegateSelectedAvatarCanvas);
    delegateSelectedAvatar.appendChild(delegateSelectedAvatarInitials);
    const delegateSelectedText = doc.createElement("div");
    Object.assign(delegateSelectedText.style,{
      flex:"1",
      lineHeight:"1.4",
      display:"flex",
      flexDirection:"column",
      gap:"2px"
    });
    const delegateSelectedTitle = doc.createElement("div");
    Object.assign(delegateSelectedTitle.style,{
      fontSize:"10px",
      textTransform:"uppercase",
      letterSpacing:"0.08em",
      opacity:"0.65"
    });
    const delegateSelectedDetails = doc.createElement("div");
    Object.assign(delegateSelectedDetails.style,{
      fontSize:"12px",
      fontWeight:"500",
      color:"var(--arrowpanel-color,#1b1b1b)",
      wordBreak:"break-word"
    });
    delegateSelectedText.appendChild(delegateSelectedTitle);
    delegateSelectedText.appendChild(delegateSelectedDetails);
    delegateSelectedInfo.appendChild(delegateSelectedAvatar);
    delegateSelectedInfo.appendChild(delegateSelectedText);
    delegateField.appendChild(delegateSelectedInfo);
    const delegateHint = doc.createElement("div");
    Object.assign(delegateHint.style,{fontSize:"11px",opacity:"0.75"});
    delegateHint.textContent = i18n("ui_create_moderator_hint");
    delegateField.appendChild(delegateHint);
    const DROPDOWN_ROW_HEIGHT = 48;
    const DROPDOWN_MIN_VISIBLE_ROWS = 5;
    const DROPDOWN_DESIRED_HEIGHT = DROPDOWN_ROW_HEIGHT * DROPDOWN_MIN_VISIBLE_ROWS;

    const delegateDropdown = doc.createElement("div");
    Object.assign(delegateDropdown.style,{
      position:"absolute",
      left:"0",
      right:"0",
      top:"calc(100% + 4px)",
      background:"var(--arrowpanel-background,#fff)",
      border:"1px solid var(--arrowpanel-border-color,#c0c0c0)",
      borderRadius:"6px",
      boxShadow:"0 12px 26px rgba(0,0,0,.25)",
      maxHeight: DROPDOWN_DESIRED_HEIGHT + "px",
      overflowY:"auto",
      display:"none",
      zIndex:"2147483647"
    });
    delegateField.appendChild(delegateDropdown);

    const adjustDelegateDropdownPosition = () => {
      delegateDropdown.style.maxHeight = DROPDOWN_DESIRED_HEIGHT + "px";
      delegateDropdown.style.top = "calc(100% + 4px)";
      delegateDropdown.style.bottom = "";
      try{
        const panelRect = panel.getBoundingClientRect();
        const fieldRect = delegateField.getBoundingClientRect();
        const dropdownRect = delegateDropdown.getBoundingClientRect();
        const margin = 12;
        const availableBelow = panelRect.bottom - fieldRect.bottom - margin;
        const availableAbove = fieldRect.top - panelRect.top - margin;
        const desired = DROPDOWN_DESIRED_HEIGHT;
        const minVisible = DROPDOWN_ROW_HEIGHT * 3;
        const computeHeight = (available) => {
          if (!Number.isFinite(available) || available <= 0){
            return desired;
          }
          const cappedDesired = Math.min(available, desired);
          const cappedMin = Math.min(available, minVisible);
          return Math.max(cappedDesired, cappedMin);
        };
        if (dropdownRect.bottom > panelRect.bottom - margin){
          if (availableAbove > availableBelow){
            delegateDropdown.style.top = "";
            delegateDropdown.style.bottom = "calc(100% + 4px)";
            delegateDropdown.style.maxHeight = computeHeight(availableAbove) + "px";
          } else {
            delegateDropdown.style.top = "calc(100% + 4px)";
            delegateDropdown.style.bottom = "";
            delegateDropdown.style.maxHeight = computeHeight(availableBelow) + "px";
          }
        } else if (availableBelow > 0){
          delegateDropdown.style.maxHeight = computeHeight(availableBelow) + "px";
        }
      }catch(_){}
    };

    const delegateSuggestionsState = {
      items: [],
      activeIndex: -1,
      visible: false
    };
    let delegateSearchTimeout = null;
    let delegateSearchSeq = 0;

    const updateDelegateStatus = (text = "", isError = false) => {
      delegateStatus.textContent = text || "";
      delegateStatus.style.color = isError ? "#b00020" : "";
    };

    const hideDelegateDropdown = () => {
      delegateDropdown.style.display = "none";
      delegateDropdown.style.maxHeight = DROPDOWN_DESIRED_HEIGHT + "px";
      delegateDropdown.style.top = "calc(100% + 4px)";
      delegateDropdown.style.bottom = "";
      delegateSuggestionsState.visible = false;
      delegateSuggestionsState.activeIndex = -1;
    };

    const updateDelegateRowHighlight = () => {
      const rows = delegateDropdown.children;
      for (const row of rows){
        if (!(row instanceof doc.defaultView.HTMLElement)) continue;
        const idx = Number(row.dataset.index);
        row.style.background = idx === delegateSuggestionsState.activeIndex
          ? "var(--arrowpanel-dimmed, rgba(0,0,0,0.08))"
          : "transparent";
      }
    };

    const ensureDelegateActiveVisible = () => {
      if (delegateSuggestionsState.activeIndex < 0) return;
      const selector = `[data-index='${delegateSuggestionsState.activeIndex}']`;
      const activeRow = delegateDropdown.querySelector(selector);
      if (!activeRow) return;
      const dropdownRect = delegateDropdown.getBoundingClientRect();
      const rowRect = activeRow.getBoundingClientRect();
      if (rowRect.top < dropdownRect.top){
        delegateDropdown.scrollTop -= (dropdownRect.top - rowRect.top);
      } else if (rowRect.bottom > dropdownRect.bottom){
        delegateDropdown.scrollTop += (rowRect.bottom - dropdownRect.bottom);
      }
    };

    const renderDelegateDropdown = () => {
      const savedScrollTop = delegateDropdown.scrollTop;
      delegateDropdown.textContent = "";
      if (!delegateSuggestionsState.items.length || doc.activeElement !== delegateInput){
        hideDelegateDropdown();
        return;
      }
    delegateSuggestionsState.visible = true;
    delegateDropdown.style.display = "block";
    adjustDelegateDropdownPosition();
    delegateSuggestionsState.items.forEach((item, index) => {
      const row = doc.createElement("div");
      Object.assign(row.style,{
        padding:"6px 10px",
        cursor:"pointer",
        display:"flex",
          alignItems:"flex-start",
          gap:"8px",
          background: index === delegateSuggestionsState.activeIndex ? "var(--arrowpanel-dimmed, rgba(0,0,0,0.08))" : "transparent"
        });
        row.dataset.index = String(index);
        if (item.avatarDataUrl){
          const avatarWrapper = doc.createElement("div");
          Object.assign(avatarWrapper.style,{
            width:"28px",
            height:"28px",
            borderRadius:"50%",
            overflow:"hidden",
            flex:"0 0 28px",
            background:"var(--toolbarbutton-hover-background, rgba(0,0,0,0.05))",
            display:"flex",
            alignItems:"center",
            justifyContent:"center",
            position:"relative",
            color:"var(--arrowpanel-color, #1b1b1b)",
            fontSize:"12px",
            fontWeight:"600"
          });
          const avatarCanvas = doc.createElement("canvas");
          avatarCanvas.width = 28;
          avatarCanvas.height = 28;
          Object.assign(avatarCanvas.style,{
            width:"100%",
            height:"100%",
            display:"none"
          });
          const placeholder = doc.createElement("span");
          const placeholderSource = (item.label && item.label.trim()) || (item.email && item.email.trim()) || "";
          const placeholderChar = placeholderSource ? placeholderSource.charAt(0).toUpperCase() : "";
          placeholder.textContent = placeholderChar;
          Object.assign(placeholder.style,{
            pointerEvents:"none",
            display:"flex",
            alignItems:"center",
            justifyContent:"center",
            width:"100%",
            height:"100%"
          });
          avatarWrapper.appendChild(placeholder);
          avatarWrapper.appendChild(avatarCanvas);
          drawAvatarOnCanvas(avatarCanvas, item.avatarDataUrl).then((ok) => {
            if (!avatarCanvas.isConnected){
              return;
            }
            if (ok){
              avatarCanvas.style.display = "block";
              if (placeholder.isConnected){
                avatarWrapper.removeChild(placeholder);
              }
            } else {
              avatarCanvas.remove();
              if (!placeholderChar){
                avatarWrapper.style.display = "none";
              }
            }
          }).catch((e) => {
            err(e);
            if (!avatarCanvas.isConnected){
              return;
            }
            avatarCanvas.remove();
            if (!placeholderChar){
              avatarWrapper.style.display = "none";
            }
          });
          row.appendChild(avatarWrapper);
        }
        const textBox = doc.createElement("div");
        Object.assign(textBox.style,{
          display:"flex",
          flexDirection:"column",
          gap:"2px",
          minWidth:"0"
        });
        const primary = doc.createElement("div");
        primary.textContent = item.label || item.id || item.email || "";
        primary.style.fontSize = "12px";
        const emailLine = doc.createElement("div");
        emailLine.textContent = item.email || "";
        Object.assign(emailLine.style,{fontSize:"11px",opacity:"0.75"});
        textBox.appendChild(primary);
        if (item.email){
          textBox.appendChild(emailLine);
        }
        if (item.id && item.id !== item.email){
          const idLine = doc.createElement("div");
          idLine.textContent = item.id;
          Object.assign(idLine.style,{fontSize:"10px",opacity:"0.6"});
          textBox.appendChild(idLine);
        }
        row.appendChild(textBox);
        row.addEventListener("mouseenter", () => {
          delegateSuggestionsState.activeIndex = index;
          updateDelegateRowHighlight();
        });
        const handleRowActivate = (event) => {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          selectDelegateSuggestion(index);
        };
        row.addEventListener("mousedown", handleRowActivate, true);
        row.addEventListener("mouseup", handleRowActivate, true);
        row.addEventListener("click", handleRowActivate, true);
        delegateDropdown.appendChild(row);
      });
      delegateDropdown.scrollTop = savedScrollTop;
      updateDelegateRowHighlight();
      ensureDelegateActiveVisible();
    };

    const formatDelegateDisplay = (item) => {
      if (!item) return "";
      const base = item.label || item.id || "";
      const email = typeof item.email === "string" ? item.email.trim() : "";
      if (email && base && email !== base){
        return base + " <" + email + ">";
      }
      return email || base;
    };

    const computeDelegateInitials = (item) => {
      if (!item) return "";
      const source = (item.label || item.id || item.email || "").trim();
      if (!source) return "";
      const parts = source.split(/\s+/).filter(Boolean);
      const letters = ((parts[0] || "").charAt(0) + (parts[1] || "").charAt(0)).toUpperCase();
      const fallback = letters || source.charAt(0);
      return fallback.toUpperCase().slice(0, 2);
    };

    let selectedAvatarDrawSource = "";
    const updateDelegateSelectedDisplay = () => {
      const label = delegateInput.dataset.selectionLabel || "";
      const avatarData = delegateInput.dataset.selectionAvatar || "";
      const initials = delegateInput.dataset.selectionInitials || "";
      if (!label){
        delegateSelectedInfo.style.display = "none";
        delegateSelectedAvatar.style.display = "none";
        delegateSelectedAvatarCanvas.style.display = "none";
        delegateSelectedAvatarInitials.style.display = "none";
        delegateSelectedTitle.textContent = "";
        delegateSelectedDetails.textContent = "";
        selectedAvatarDrawSource = "";
        return;
      }
      delegateSelectedInfo.style.display = "flex";
      delegateSelectedTitle.textContent = i18n("ui_delegate_selected_title");
      delegateSelectedDetails.textContent = label;
      const effectiveInitials = initials || label.trim().charAt(0).toUpperCase();
      delegateSelectedAvatarInitials.textContent = effectiveInitials;
      delegateSelectedAvatarInitials.style.display = effectiveInitials ? "flex" : "none";
      delegateSelectedAvatar.style.display = (avatarData || effectiveInitials) ? "flex" : "none";
      if (!avatarData){
        delegateSelectedAvatarCanvas.style.display = "none";
        selectedAvatarDrawSource = "";
        if (!effectiveInitials){
          delegateSelectedAvatar.style.display = "none";
        }
        return;
      }
      if (avatarData === selectedAvatarDrawSource && delegateSelectedAvatarCanvas.style.display === "block"){
        return;
      }
      selectedAvatarDrawSource = avatarData;
      delegateSelectedAvatarCanvas.style.display = "none";
      drawAvatarOnCanvas(delegateSelectedAvatarCanvas, avatarData).then((ok) => {
        if (delegateInput.dataset.selectionAvatar !== avatarData){
          return;
        }
        if (ok){
          delegateSelectedAvatarCanvas.style.display = "block";
          delegateSelectedAvatarInitials.style.display = "none";
        } else {
          delegateSelectedAvatarCanvas.style.display = "none";
          delegateSelectedAvatarInitials.style.display = effectiveInitials ? "flex" : "none";
          if (!effectiveInitials){
            delegateSelectedAvatar.style.display = "none";
          }
        }
      }).catch(() => {
        if (delegateInput.dataset.selectionAvatar !== avatarData){
          return;
        }
        delegateSelectedAvatarCanvas.style.display = "none";
        delegateSelectedAvatarInitials.style.display = effectiveInitials ? "flex" : "none";
        if (!effectiveInitials){
          delegateSelectedAvatar.style.display = "none";
        }
      });
    };

    updateDelegateSelectedDisplay();

    const selectDelegateSuggestion = (index) => {
      const suggestion = delegateSuggestionsState.items[index];
      if (!suggestion) return;
      delegateInput.value = suggestion.id;
      const selectionLabel = formatDelegateDisplay(suggestion);
      delegateInput.dataset.selectionLabel = selectionLabel;
      delegateInput.dataset.selectionAvatar = suggestion.avatarDataUrl || "";
      delegateInput.dataset.selectionInitials = computeDelegateInitials(suggestion);
      updateDelegateSelectedDisplay();
      updateDelegateStatus("");
      hideDelegateDropdown();
      try{
        const len = delegateInput.value.length;
        delegateInput.setSelectionRange(len, len);
      }catch(_){}
    };

    const scheduleDelegateSearch = (term) => {
      if (delegateSearchTimeout){
        docWin.clearTimeout(delegateSearchTimeout);
      }
      delegateSearchTimeout = docWin.setTimeout(async () => {
        const seq = ++delegateSearchSeq;
        delegateSuggestionsState.items = [];
        delegateSuggestionsState.activeIndex = -1;
        hideDelegateDropdown();
        updateDelegateStatus(term ? i18n("ui_delegate_status_searching") : i18n("ui_delegate_status_loading"));
        try{
          const response = await requestUtility(context, { type: "searchUsers", searchTerm: term, limit: 200 });
          if (seq !== delegateSearchSeq) return;
          let items = [];
          if (response){
            if (response.ok && Array.isArray(response.users)){
              items = response.users;
            } else if (Array.isArray(response.result)){
              items = response.result;
            } else if (!response.ok && response.error){
              throw new Error(response.error);
            }
          }
          const emailPattern = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
          items = items
            .filter((item) => {
              if (!item) return false;
              const email = typeof item.email === "string" ? item.email.trim() : "";
              return email && emailPattern.test(email);
            })
            .map((item) => {
              const email = typeof item.email === "string" ? item.email.trim() : "";
              return Object.assign({}, item, { email });
            });
          delegateSuggestionsState.items = items;
          delegateSuggestionsState.activeIndex = items.length ? 0 : -1;
          if (!items.length){
            updateDelegateStatus(term ? i18n("ui_delegate_status_none_with_email") : i18n("ui_delegate_status_none_found"));
            hideDelegateDropdown();
          }else{
            const countText = items.length === 1
              ? i18n("ui_delegate_status_single")
              : i18n("ui_delegate_status_many", [items.length]);
            updateDelegateStatus(countText);
            renderDelegateDropdown();
          }
        }catch(e){
          if (seq !== delegateSearchSeq) return;
          err(e);
          updateDelegateStatus(e?.message || i18n("ui_delegate_status_error"), true);
          delegateSuggestionsState.items = [];
          hideDelegateDropdown();
        }
      }, 250);
    };

    const handleDelegateDocumentClick = (event) => {
      if (!delegateField.contains(event.target)){
        hideDelegateDropdown();
      }
    };

    const handleDelegateKeyDown = (event) => {
      if (!delegateSuggestionsState.visible || !delegateSuggestionsState.items.length) return;
      if (event.key === "ArrowDown"){
        event.preventDefault();
        const count = delegateSuggestionsState.items.length;
        delegateSuggestionsState.activeIndex = (delegateSuggestionsState.activeIndex + 1 + count) % count;
        updateDelegateRowHighlight();
        ensureDelegateActiveVisible();
      } else if (event.key === "ArrowUp"){
        event.preventDefault();
        const count = delegateSuggestionsState.items.length;
        delegateSuggestionsState.activeIndex = (delegateSuggestionsState.activeIndex - 1 + count) % count;
        updateDelegateRowHighlight();
        ensureDelegateActiveVisible();
      } else if (event.key === "Enter"){
        if (delegateSuggestionsState.activeIndex >= 0){
          event.preventDefault();
          selectDelegateSuggestion(delegateSuggestionsState.activeIndex);
        }
      } else if (event.key === "Escape"){
        hideDelegateDropdown();
      }
    };

    delegateInput.addEventListener("focus", () => {
      delegateSuggestionsState.items = [];
      delegateSuggestionsState.activeIndex = -1;
      updateDelegateSelectedDisplay();
      scheduleDelegateSearch(delegateInput.value.trim());
    });

    delegateInput.addEventListener("input", () => {
      delegateInput.dataset.selectionLabel = "";
      delegateInput.dataset.selectionAvatar = "";
      delegateInput.dataset.selectionInitials = "";
      updateDelegateSelectedDisplay();
      scheduleDelegateSearch(delegateInput.value.trim());
    });

    delegateInput.addEventListener("keydown", handleDelegateKeyDown);
    delegateInput.addEventListener("blur", () => {
      docWin.setTimeout(() => hideDelegateDropdown(), 80);
    });

    delegateClearBtn.addEventListener("click", () => {
      delegateInput.value = "";
      delegateInput.dataset.selectionLabel = "";
      delegateInput.dataset.selectionAvatar = "";
      delegateInput.dataset.selectionInitials = "";
      updateDelegateSelectedDisplay();
      updateDelegateStatus("");
      scheduleDelegateSearch("");
      delegateInput.focus();
    });

    doc.addEventListener("mousedown", handleDelegateDocumentClick, true);
    scheduleDelegateSearch("");

    const cleanupDialogState = () => {
      updateDelegateStatus("");
      if (delegateSearchTimeout){
        docWin.clearTimeout(delegateSearchTimeout);
        delegateSearchTimeout = null;
      }
      delegateSuggestionsState.items = [];
      delegateSuggestionsState.activeIndex = -1;
      hideDelegateDropdown();
      delegateInput.dataset.selectionLabel = "";
      delegateInput.dataset.selectionAvatar = "";
      delegateInput.dataset.selectionInitials = "";
      updateDelegateSelectedDisplay();
    };

    const closeOverlay = () => {
      cleanupDialogState();
      doc.removeEventListener("mousedown", handleDelegateDocumentClick, true);
      overlay.remove();
    };

    grid.appendChild(titleLabel);
    grid.appendChild(titleInput);
    grid.appendChild(passLabel);
    grid.appendChild(passInput);
    grid.appendChild(lobbyLabel);
    grid.appendChild(lobbySpacer);
    grid.appendChild(listableLabel);
    grid.appendChild(listableSpacer);
    grid.appendChild(modeLabel);
    grid.appendChild(modeField);
    grid.appendChild(delegateLabel);
    grid.appendChild(delegateField);

    const buttons = doc.createElement("div");
    Object.assign(buttons.style,{display:"flex",justifyContent:"flex-end",gap:"8px",marginTop:"16px"});

    const cancelBtn = doc.createElement("button");
    cancelBtn.id = "nc_cancel";
    cancelBtn.textContent = i18n("ui_button_cancel");

    const okBtn = doc.createElement("button");
    okBtn.id = "nc_ok";
    okBtn.textContent = i18n("ui_button_ok");
    okBtn.style.appearance = "auto";

    buttons.appendChild(cancelBtn);
    buttons.appendChild(okBtn);

    panel.appendChild(heading);
    panel.appendChild(grid);
    panel.appendChild(buttons);

    requestUtility(context, { type: "supportsEventConversation" }).then((res) => {
      log("supportsEventConversation response", {
        ok: !!(res && res.ok),
        supported: res ? res.supported : undefined,
        reason: res ? (res.reason || res.error || "") : ""
      });
      if (res && res.ok){
        applyEventModeSupport({ supported: res.supported, reason: res.reason });
      } else if (res){
        applyEventModeSupport({ supported: null, reason: res.error || "" });
      }
    }).catch((e) => {
      err(e);
      applyEventModeSupport({ supported:null, reason: e?.message || "" });
    });

    overlay.addEventListener("click",(e)=>{ if(e.target===overlay) closeOverlay(); });
    cancelBtn.addEventListener("click", () => closeOverlay());
    okBtn.addEventListener("click", async ()=>{
      if (okBtn.disabled) return;
      const title = titleInput.value.trim() || defaultTitle;
      const password = passInput.value.trim();
      const enableLobby = !!lobbyCheckbox.checked;
      const enableListable = !!listableCheckbox.checked;
      const originalLabel = okBtn.textContent;
      if (password && password.length < 5){
        showAlert(docWin, i18n("ui_create_password_short"));
        return;
      }
      okBtn.disabled = true;
      okBtn.textContent = i18n("ui_button_create_progress");

      const restore = () => {
        okBtn.disabled = false;
        okBtn.textContent = originalLabel;
      };

      const innerDoc = (doc.getElementById && doc.getElementById("calendar-item-panel-iframe") && doc.getElementById("calendar-item-panel-iframe").contentDocument) || doc;
      const descriptionText = extractDescriptionText(innerDoc);

      const startTs = extractStartTimestamp(innerDoc);
      const delegateId = delegateInput.value.trim();
      const useEventConversation = modeEventRadio.checked && !modeEventRadio.disabled;
      const eventMetadata = useEventConversation ? buildEventObjectMetadata(innerDoc, { title, startTimestamp: startTs }) : null;

      const createPayload = {
        title,
        password: password || undefined,
        enableLobby,
        enableListable,
        description: descriptionText || "",
        startTimestamp: startTs,
        eventConversation: useEventConversation,
        objectType: eventMetadata?.objectType,
        objectId: eventMetadata?.objectId
      };
      log("create dialog payload", describeCreatePayload(createPayload));

      let response = null;
      try {
        response = await requestCreateFromExtension(context, createPayload);
      } catch(sendErr) {
        err(sendErr);
        restore();
        showAlert(docWin, i18n("ui_create_send_failed", [sendErr?.message || String(sendErr)]));
        return;
      }

      if (!response) {
        restore();
        showAlert(docWin, i18n("ui_create_no_handler"));
        return;
      }

      if (!response.ok || !response.url) {
        restore();
        const msg = response.error ? response.error : i18n("ui_create_unknown_error");
        showAlert(docWin, i18n("ui_create_failed", [msg]));
        return;
      }

      log("create dialog success", {
        token: response.token ? shortToken(response.token) : "",
        fallback: !!response.fallback,
        reason: response.reason || ""
      });

      clearPendingModerator(docWin);

      let delegationResult = null;
      let pendingDelegation = null;
      const delegateDisplayName = delegateInput.dataset.selectionLabel || delegateId;
      if (delegateId){
        log("create dialog delegate request", {
          token: response.token ? shortToken(response.token) : "",
          delegateId,
          fromLobbyQueue: !!enableLobby
        });
        if (enableLobby){
          pendingDelegation = {
            token: response.token,
            delegateId,
            displayName: delegateDisplayName || delegateId
          };
          queuePendingModerator(docWin, pendingDelegation);
        } else {
          try{
            const delegateResponse = await requestUtility(context, { type: "delegateModerator", token: response.token, newModerator: delegateId });
            if (!delegateResponse || !delegateResponse.ok){
              const msg = delegateResponse?.error || i18n("ui_moderator_transfer_failed");
              showAlert(docWin, msg);
            }else{
              delegationResult = delegateResponse.result || { delegate: delegateId };
              log("create dialog delegate success", {
                token: response.token ? shortToken(response.token) : "",
                delegate: delegationResult.delegate || delegateId,
                leftSelf: !!delegationResult.leftSelf
              });
            }
          }catch(delegateErr){
            err(delegateErr);
            showAlert(docWin, i18n("ui_moderator_transfer_failed_with_reason", [delegateErr?.message || delegateErr]));
          }
        }
      }

      closeOverlay();
      const placed = fillIntoEvent(innerDoc, response.url, password || null, title);
      if (useEventConversation && response.fallback){
        applyEventModeSupport({ supported:false, reason: response.reason || "" });
      }
      const alertLines = [i18n("ui_alert_link_inserted", [response.url])];
      if (!placed) alertLines.push(i18n("ui_alert_location_missing"));
      if (response.fallback) {
        if (useEventConversation){
          alertLines.push(i18n("ui_alert_event_fallback"));
        } else {
          alertLines.push(i18n("ui_alert_generic_fallback"));
        }
        if (response.reason) alertLines.push(i18n("ui_alert_reason", [response.reason]));
      }
      if (pendingDelegation){
        alertLines.push(i18n("ui_alert_pending_delegation", [pendingDelegation.displayName || pendingDelegation.delegateId]));
      } else if (delegationResult && delegationResult.delegate){
        alertLines.push(i18n("ui_alert_delegation_done", [delegateDisplayName || delegationResult.delegate]));
        if (delegationResult.leftSelf){
          alertLines.push(i18n("ui_alert_delegation_removed"));
        }
      }
      showAlert(docWin, alertLines.join("\n"));
      const watcherAllowed = enableLobby && response.token;
      if (watcherAllowed){
        setupLobbyWatcher(context, doc, innerDoc, response.token, true);
      }
    });

    overlay.appendChild(panel);
    (doc.body||doc.documentElement).appendChild(overlay);
    okBtn.focus();
  }catch(e){
    err(e);
    err("create dialog failed");
    const docWin = doc.defaultView || window;
    const base = await getBaseUrl(context) || docWin.prompt(i18n("ui_prompt_base_url"), "https://cloud.example.com");
    if (!base) return;
    const title = docWin.prompt(i18n("ui_prompt_title"), i18n("ui_default_title"));
    if (title === null) return;
    const token = randToken(10);
    const url = String(base).replace(/\/$/,"") + "/call/" + token;
    const ok = fillIntoEvent(doc, url, null);
    const fallbackLines = [i18n("ui_alert_link_inserted", [url])];
    if (!ok) fallbackLines.push(i18n("ui_alert_location_missing"));
    showAlert(docWin, fallbackLines.join("\n"));
  }
}
async function openSelectDialog(doc, context){
  let overlay = null;
  try{
    const docWin = doc.defaultView || window;
    log("openSelectDialog", { url: doc?.URL || "" });
    overlay = doc.createElement("div");
    Object.assign(overlay.style,{position:"fixed",inset:"0",background:"rgba(0,0,0,.25)",zIndex:"2147483646"});
    const panel = doc.createElement("div");
    Object.assign(panel.style,{position:"fixed",top:"15%",left:"50%",transform:"translateX(-50%)",
      background:"var(--arrowpanel-background,#fff)",border:"1px solid var(--arrowpanel-border-color,#c0c0c0)",
      borderRadius:"8px",boxShadow:"0 10px 30px rgba(0,0,0,.25)",minWidth:"520px",maxWidth:"640px",
      maxHeight:"80vh",padding:"16px",zIndex:"2147483647",display:"flex",flexDirection:"column",overflowY:"auto",overflowX:"hidden"});
    overlay.appendChild(panel);
    (doc.body || doc.documentElement).appendChild(overlay);

    const heading = doc.createElement("h2");
    heading.textContent = i18n("ui_select_heading");
    Object.assign(heading.style,{margin:"0 0 10px",font:"600 16px system-ui"});
    panel.appendChild(heading);

    const searchBox = doc.createElement("div");
    Object.assign(searchBox.style,{marginBottom:"8px"});
    const searchInput = doc.createElement("input");
    Object.assign(searchInput,{type:"search",placeholder: i18n("ui_select_search_placeholder")});
    Object.assign(searchInput.style,{width:"100%",padding:"8px 10px",border:"1px solid var(--arrowpanel-border-color,#c0c0c0)",borderRadius:"6px"});
    searchBox.appendChild(searchInput);
    panel.appendChild(searchBox);

    const statusLabel = doc.createElement("div");
    Object.assign(statusLabel.style,{fontSize:"12px",opacity:"0.8",minHeight:"16px"});
    panel.appendChild(statusLabel);

    const listContainer = doc.createElement("div");
    Object.assign(listContainer.style,{flex:"1 1 auto",overflowY:"auto",border:"1px solid var(--arrowpanel-border-color,#c0c0c0)",
      borderRadius:"6px",padding:"4px",background:"rgba(0,0,0,0.02)"});
    panel.appendChild(listContainer);

    const detailBox = doc.createElement("div");
    Object.assign(detailBox.style,{marginTop:"12px",padding:"12px",border:"1px solid var(--arrowpanel-border-color,#c0c0c0)",
      borderRadius:"6px",background:"rgba(0,0,0,0.03)",display:"flex",flexDirection:"column",gap:"6px"});
    panel.appendChild(detailBox);

    const detailHeader = doc.createElement("div");
    Object.assign(detailHeader.style,{display:"flex",flexDirection:"column",gap:"4px"});
    detailBox.appendChild(detailHeader);

    const detailTitle = doc.createElement("h3");
    detailTitle.textContent = i18n("ui_select_no_selection");
    Object.assign(detailTitle.style,{margin:"0",font:"600 15px system-ui"});
    detailHeader.appendChild(detailTitle);

    const detailMeta = doc.createElement("div");
    Object.assign(detailMeta.style,{fontSize:"11px",opacity:"0.75"});
    detailHeader.appendChild(detailMeta);

    const linkRow = doc.createElement("div");
    Object.assign(linkRow.style,{fontSize:"12px",wordBreak:"break-all"});
    detailBox.appendChild(linkRow);

    const descLabel = doc.createElement("div");
    Object.assign(descLabel.style,{fontSize:"12px",whiteSpace:"pre-wrap"});
    detailBox.appendChild(descLabel);

    const passwordInfo = doc.createElement("div");
    Object.assign(passwordInfo.style,{fontSize:"12px",opacity:"0.8"});
    detailBox.appendChild(passwordInfo);

    const lobbyRow = doc.createElement("label");
    Object.assign(lobbyRow.style,{display:"flex",alignItems:"center",gap:"6px",fontSize:"12px",userSelect:"none"});
    const lobbyCheckbox = doc.createElement("input");
    lobbyCheckbox.type = "checkbox";
    lobbyCheckbox.disabled = true;
    lobbyRow.appendChild(lobbyCheckbox);
    lobbyRow.appendChild(doc.createTextNode(" " + i18n("ui_select_lobby_toggle")));
    detailBox.appendChild(lobbyRow);

    const lobbyInfo = doc.createElement("div");
    Object.assign(lobbyInfo.style,{fontSize:"12px",opacity:"0.8"});
    detailBox.appendChild(lobbyInfo);

    const moderatorInfo = doc.createElement("div");
    Object.assign(moderatorInfo.style,{fontSize:"11px",opacity:"0.7"});
    detailBox.appendChild(moderatorInfo);

    const creds = await getCredentials(context);
    const currentUser = (creds.user || "").trim();
    const baseUrl = creds.baseUrl || await getBaseUrl(context);
    if (!baseUrl){
      overlay.remove();
      showAlert(docWin, i18n("ui_select_missing_base_url"));
      return;
    }
    const normalizedBaseUrl = String(baseUrl).replace(/\/$/,"");

    const buttons = doc.createElement("div");
    Object.assign(buttons.style,{display:"flex",justifyContent:"flex-end",gap:"8px",marginTop:"14px"});
    panel.appendChild(buttons);

    const cancelBtn = doc.createElement("button");
    cancelBtn.textContent = i18n("ui_button_cancel");
    const okBtn = doc.createElement("button");
    okBtn.textContent = i18n("ui_button_apply");
    okBtn.disabled = true;

    buttons.appendChild(cancelBtn);
    buttons.appendChild(okBtn);

    const innerDoc = (doc.getElementById && doc.getElementById("calendar-item-panel-iframe") && doc.getElementById("calendar-item-panel-iframe").contentDocument) || doc;
    const winObj = doc.defaultView || window;

    let rooms = [];
    let activeItem = null;
    let selectedDetails = null;
    let desiredLobbyState = false;
    let detailRequestToken = null;
    let searchSeq = 0;
    let currentToken = null;
    let isModerator = false;

    const selectCss = {
      display:"flex",
      flexDirection:"column",
      alignItems:"flex-start",
      gap:"4px",
      padding:"8px 10px",
      borderRadius:"6px",
      margin:"4px 0",
      background:"transparent",
      cursor:"pointer",
      transition:"background-color 120ms ease"
    };

    function setStatus(text){
      statusLabel.textContent = text || "";
    }

    function updateLobbyInfo(timer, enabled){
      if (!enabled){
        lobbyInfo.textContent = i18n("ui_lobby_disabled");
        return;
      }
      if (typeof timer === "number" && Number.isFinite(timer) && timer > 0){
        const date = new Date(timer * 1000);
        lobbyInfo.textContent = i18n("ui_lobby_active_with_time", [date.toLocaleString()]);
      } else {
        lobbyInfo.textContent = i18n("ui_lobby_active_without_time");
      }
    }

    function clearDetails(){
      detailTitle.textContent = i18n("ui_select_no_selection");
      detailMeta.textContent = "";
      linkRow.textContent = "";
      descLabel.textContent = "";
      passwordInfo.textContent = "";
      lobbyCheckbox.checked = false;
      lobbyCheckbox.disabled = true;
      lobbyCheckbox.title = "";
      lobbyInfo.textContent = "";
      moderatorInfo.textContent = "";
      selectedDetails = null;
      desiredLobbyState = false;
      currentToken = null;
      isModerator = false;
      okBtn.disabled = true;
    }

    function renderList(){
      listContainer.textContent = "";
      activeItem = null;
      if (!rooms.length){
        setStatus(i18n("ui_select_status_none"));
        clearDetails();
        return;
      }
      const countText = rooms.length === 1
        ? i18n("ui_select_status_single")
        : i18n("ui_select_status_many", [rooms.length]);
      setStatus(countText);
      for (const room of rooms){
        const item = doc.createElement("div");
        Object.assign(item.style, selectCss);
        item.dataset.token = room.token;
        item.addEventListener("mouseenter", () => { if (item !== activeItem) item.style.background = "rgba(0,0,0,0.06)"; });
        item.addEventListener("mouseleave", () => { if (item !== activeItem) item.style.background = "transparent"; });
        const title = doc.createElement("div");
        title.textContent = room.displayName || room.name || room.token;
        title.style.fontWeight = "500";
        item.appendChild(title);
        const linkLine = doc.createElement("div");
        linkLine.textContent = normalizedBaseUrl + "/call/" + room.token;
        Object.assign(linkLine.style,{fontSize:"11px",opacity:"0.75",wordBreak:"break-all"});
        item.appendChild(linkLine);
        const firstLine = (room.description || "")
          .split(/\r?\n/)
          .map((line) => line.trim())
          .find((line) => line.length > 0) || "";
        if (firstLine){
          const desc = doc.createElement("div");
          desc.textContent = firstLine;
          Object.assign(desc.style,{fontSize:"11px",opacity:"0.75"});
          item.appendChild(desc);
        }
        const badges = [];
        if (room.hasPassword) badges.push(i18n("ui_badge_password_required"));
        if (room.listable){
          badges.push(i18n("ui_badge_listed"));
        } else if (room.source === "own" || room.isParticipant){
          badges.push(i18n("ui_badge_owned"));
        }
        if (badges.length){
          const meta = doc.createElement("div");
          meta.textContent = badges.join(" | ");
          Object.assign(meta.style,{fontSize:"10px",opacity:"0.6"});
          item.appendChild(meta);
        }
        item.addEventListener("click", () => selectRoom(room, item));
        listContainer.appendChild(item);
      }
    }

    async function selectRoom(room, item){
      if (!room) return;
      if (activeItem) activeItem.style.background = "transparent";
      activeItem = item;
      if (activeItem) activeItem.style.background = "var(--arrowpanel-dimmed, rgba(0,0,0,0.12))";
      clearDetails();
      currentToken = room.token;
      detailTitle.textContent = room.displayName || room.name || room.token;
      const metaParts = [];
      if (room.listable) metaParts.push(i18n("ui_badge_listed"));
      else if (room.source === "own" || room.isParticipant) metaParts.push(i18n("ui_badge_owned"));
      if (room.hasPassword) metaParts.push(i18n("ui_badge_password_required"));
      if (room.guestsAllowed !== false) metaParts.push(i18n("ui_badge_guests_allowed"));
      detailMeta.textContent = metaParts.join(" | ");
      linkRow.textContent = normalizedBaseUrl + "/call/" + currentToken;
      descLabel.textContent = (room.description || "").trim();
      passwordInfo.textContent = room.hasPassword ? i18n("ui_password_info_yes") : i18n("ui_password_info_no");
      lobbyInfo.textContent = i18n("ui_lobby_loading_details");
      okBtn.disabled = true;
      const requestToken = ++detailRequestToken;

      try{
        const rawDetails = await fetchRoomDetails(context, currentToken);
        if (detailRequestToken !== requestToken) return;
        const normalized = Object.assign({}, room, rawDetails || {});
        normalized.token = normalized.token || currentToken;
        selectedDetails = normalized;
        detailTitle.textContent = normalized.displayName || normalized.name || normalized.token;
        const normalizedMeta = [];
        if (normalized.listable) normalizedMeta.push(i18n("ui_badge_listed"));
        else if (normalized.source === "own" || normalized.isParticipant) normalizedMeta.push(i18n("ui_badge_owned"));
        if (normalized.hasPassword) normalizedMeta.push(i18n("ui_badge_password_required"));
        normalizedMeta.push(normalized.guestsAllowed === false ? i18n("ui_badge_guests_forbidden") : i18n("ui_badge_guests_allowed"));
        detailMeta.textContent = normalizedMeta.join(" | ");
        linkRow.textContent = normalizedBaseUrl + "/call/" + normalized.token;
        descLabel.textContent = (normalized.description || "").trim();
        passwordInfo.textContent = normalized.hasPassword ? i18n("ui_password_info_yes") : i18n("ui_password_info_no");
        desiredLobbyState = normalized.lobbyState === 1;
        lobbyCheckbox.checked = desiredLobbyState;
        lobbyInfo.textContent = i18n("ui_lobby_fetching_status");
        okBtn.disabled = true;

        const roomPermissions = typeof normalized.permissions === "number" ? normalized.permissions : null;
        isModerator = !!(roomPermissions != null && (roomPermissions & 64) !== 0);
        let participants = [];
        const allowParticipantLookup = isModerator || normalized.isParticipant || normalized.source === "own";
        if (allowParticipantLookup){
          try{
            participants = await fetchRoomParticipants(context, normalized.token);
          }catch(partErr){
            err(partErr);
          }
        }
        if (!isModerator && participants.length){
          const loweredUser = currentUser ? currentUser.toLowerCase() : "";
          if (loweredUser){
            isModerator = participants.some((p) => {
              if (!p) return false;
              const actor = (p.actorId || "").trim().toLowerCase();
              if (!actor || actor !== loweredUser) return false;
              if (p.participantType === 1) return true;
              if (typeof p.permissions === "number" && (p.permissions & 64) !== 0) return true;
              return false;
            });
          }
        }
        if (!allowParticipantLookup && !isModerator){
          moderatorInfo.textContent = i18n("ui_moderator_not_participant");
        }
        lobbyCheckbox.disabled = !isModerator;
        lobbyCheckbox.title = isModerator ? "" : i18n("ui_lobby_no_permission");
        moderatorInfo.textContent = isModerator ? i18n("ui_moderator_can_manage") : i18n("ui_moderator_cannot_manage");
        desiredLobbyState = lobbyCheckbox.checked;
        updateLobbyInfo(normalized.lobbyTimer, desiredLobbyState);
        okBtn.disabled = false;
        const entry = rooms.find(r => r.token === normalized.token);
        if (entry){
          entry.description = normalized.description || entry.description;
          entry.hasPassword = !!normalized.hasPassword;
        }
      }catch(fetchErr){
        err(fetchErr);
        if (detailRequestToken !== requestToken) return;
        lobbyInfo.textContent = i18n("ui_details_error");
        moderatorInfo.textContent = "";
        okBtn.disabled = true;
      }
    }

    lobbyCheckbox.addEventListener("change", () => {
      if (!selectedDetails){
        lobbyCheckbox.checked = false;
        desiredLobbyState = false;
        updateLobbyInfo(0, false);
        return;
      }
      desiredLobbyState = lobbyCheckbox.checked;
      updateLobbyInfo(selectedDetails?.lobbyTimer || 0, desiredLobbyState);
    });

    cancelBtn.addEventListener("click", () => overlay.remove());
    overlay.addEventListener("click", (ev) => { if (ev.target === overlay) overlay.remove(); });

    let searchTimeout = null;
    searchInput.addEventListener("input", () => {
      if (searchTimeout) winObj.clearTimeout(searchTimeout);
      searchTimeout = winObj.setTimeout(() => {
        loadRooms(searchInput.value.trim());
      }, 300);
    });

    async function loadRooms(term){
      const seq = ++searchSeq;
      setStatus(i18n("ui_select_status_loading"));
      listContainer.textContent = "";
      clearDetails();
      try{
        const fetched = await fetchListedRooms(context, term || "");
        if (seq !== searchSeq) return;
        rooms = (Array.isArray(fetched) ? fetched : []).map((room) => {
          const token = room && (room.token || room.roomToken);
          if (!token) return null;
          const copy = Object.assign({}, room, { token });
          copy.displayName = copy.displayName || copy.name || "";
          copy.description = copy.description || "";
          copy.hasPassword = !!copy.hasPassword;
          copy.listable = copy.listable === 1 || copy.listable === true;
          copy.guestsAllowed = copy.guestsAllowed !== false;
          copy.isParticipant = copy.isParticipant === true || typeof copy.participantType === "number";
          return copy;
        }).filter(Boolean);
        rooms = rooms.filter((room) => {
          if (!room.guestsAllowed) return false;
          const joinable = room.listable || room.source === "own" || room.isParticipant;
          return joinable;
        });
        rooms.sort((a,b) => {
          const left = (a.displayName || a.name || "").toLowerCase();
          const right = (b.displayName || b.name || "").toLowerCase();
          return left.localeCompare(right);
        });
        log("loadRooms result", { term: term || "", count: rooms.length });
        renderList();
      }catch(loadErr){
        err(loadErr);
        if (seq !== searchSeq) return;
        setStatus(i18n("ui_select_status_error", [loadErr?.message || loadErr]));
      }
    }

    await loadRooms("");

    okBtn.addEventListener("click", async () => {
      if (!selectedDetails) return;
      const originalLabel = okBtn.textContent;
      okBtn.disabled = true;
      okBtn.textContent = i18n("ui_button_apply_progress");
      try{
        const title = selectedDetails.displayName || selectedDetails.name || i18n("ui_default_title");
        const roomUrl = normalizedBaseUrl + "/call/" + selectedDetails.token;
        log("select dialog apply", {
          token: shortToken(selectedDetails.token),
          desiredLobbyState,
          isModerator
        });
        const initialLobbyState = selectedDetails.lobbyState === 1;
        let lobbyStateForWatcher = initialLobbyState;

        if (isModerator){
          if (desiredLobbyState !== initialLobbyState){
            const startTs = desiredLobbyState ? extractStartTimestamp(innerDoc) : undefined;
            await performLobbyUpdate(context, {
              token: selectedDetails.token,
              enableLobby: desiredLobbyState,
              startTimestamp: desiredLobbyState ? startTs : undefined
            });
            lobbyStateForWatcher = !!desiredLobbyState;
          }
        } else {
          lobbyStateForWatcher = initialLobbyState;
        }

        selectedDetails.lobbyState = lobbyStateForWatcher ? 1 : 0;

        const placed = fillIntoEvent(innerDoc, roomUrl, null, title);
        appendEventDescription(innerDoc, selectedDetails.description || "");

        overlay.remove();
        setupLobbyWatcher(context, doc, innerDoc, selectedDetails.token, lobbyStateForWatcher);

        const messageLines = [i18n("ui_alert_link_inserted", [roomUrl])];
        if (!placed) messageLines.push(i18n("ui_alert_location_missing"));
        if (selectedDetails.hasPassword) messageLines.push(i18n("ui_alert_password_protected"));
        messageLines.push(lobbyStateForWatcher ? i18n("ui_alert_lobby_state_active") : i18n("ui_alert_lobby_state_inactive"));
        if (!isModerator){
          messageLines.push(i18n("ui_alert_lobby_no_rights"));
        }
        showAlert(docWin, messageLines.join("\n"));
      }catch(applyErr){
        const applyMsg = applyErr?.message || String(applyErr);
        const permissionProblem = applyMsg && (applyMsg.toLowerCase().includes("berechtigung") || applyMsg.includes("403"));
        if (!permissionProblem){
          err(applyErr);
        }
        okBtn.disabled = false;
        okBtn.textContent = originalLabel;
        const alertMsg = permissionProblem
          ? i18n("ui_select_apply_permission_denied")
          : i18n("ui_select_apply_failed", [applyMsg]);
        if (permissionProblem && selectedDetails){
          const currentLobby = selectedDetails.lobbyState === 1;
          desiredLobbyState = currentLobby;
          lobbyCheckbox.checked = currentLobby;
          lobbyCheckbox.disabled = true;
          lobbyCheckbox.title = i18n("ui_lobby_no_permission");
          moderatorInfo.textContent = i18n("ui_moderator_cannot_manage");
          updateLobbyInfo(selectedDetails.lobbyTimer || 0, currentLobby);
        }
        showAlert(docWin, alertMsg);
      }
    });
  }catch(e){
    if (overlay) {
      try { overlay.remove(); }catch(_){}
    }
    err(e);
    const docWin = doc?.defaultView || window;
    showAlert(docWin, i18n("ui_select_rooms_load_failed", [e?.message || e]));
  }
}
function findDescriptionField(candidates){
  for (const d of candidates){
    try{
      const host = d.querySelector && d.querySelector("editor#item-description");
      let target = null;
      if (host) target = host.inputField || host.contentDocument?.body || host;
      if (!target){
        const fallbacks = [ "textarea", "[contenteditable='true']", "div[role='textbox']" ];
        for (const sel of fallbacks){
          const el = d.querySelector && d.querySelector(sel);
          if (el){ target = el; break; }
        }
      }
      if (target) return target;
    }catch(_){}
  }
  return null;
}

async function requestUtility(context, payload){
  log("requestUtility", summarizeUtilityPayload(payload));
  const handlers = UTILITY_HANDLERS.get(context);
  if (!handlers || handlers.size === 0) return null;
  for (const fire of handlers){
    try{
      const result = await fire.async(payload);
      if (result !== undefined) return result;
    }catch(e){
      err(e);
      }
  }
  return null;
}

function appendEventDescription(doc, text){
  if (text == null) return;
  const value = String(text).trim();
  if (!value) return;
  const candidates = collectEventDocs(doc);
  const target = findDescriptionField(candidates);
  if (!target) return;
  try{
    let current = "";
    if ("value" in target) current = target.value || "";
    else if (target.innerText) current = target.innerText;
    else current = target.textContent || "";
    if (current && current.includes(value)) return;
    const needsSpacing = current && !/\n$/.test(current);
    const addition = (current ? (needsSpacing ? "\n\n" : "\n") : "") + value;
    if ("value" in target){
      target.value = (target.value || "") + addition;
      const descDoc = target.ownerDocument || doc;
      try{
        target.dispatchEvent(new (descDoc?.defaultView || window).Event("input",{ bubbles:true }));
      }catch(_){}
    } else if (target.ownerDocument && target.ownerDocument.execCommand){
      target.ownerDocument.execCommand("insertText", false, addition);
    } else {
      target.textContent = (target.textContent || "") + addition;
    }
  }catch(_){}
}

function fillIntoEvent(doc, url, password, title){
  const candidates = collectEventDocs(doc);
  let placed = false;

  try {
    for (const d of candidates){
      const t = d.querySelector("#item-title");
      if (t && title){
        t.focus?.();
        if ("value" in t) t.value = title; else t.textContent = title;
        t.dispatchEvent?.(new d.defaultView.Event("input",{bubbles:true}));
        break;
      }
    }
  } catch(_){ }

  const placeSelectors = [
    'input[aria-label="Ort"]', 'input[placeholder="Ort"]',
    'input#item-location', 'input[name="location"]',
    'textbox[id*="location"]', 'input[type="text"]'
  ];
  let placeInput = null;
  for (const d of candidates){
    for (const sel of placeSelectors){
      const el = d.querySelector(sel);
      if (el) { placeInput = el; break; }
    }
    if (placeInput) break;
  }
  if (placeInput) {
    try {
      placeInput.focus();
      placeInput.value = url;
      const placeDoc = placeInput.ownerDocument || doc;
      placeInput.dispatchEvent(new placeDoc.defaultView.Event("input",{bubbles:true}));
      placed = true;
    } catch(_){ }
  }

  const desc = findDescriptionField(candidates);
  try {
    const lines = [i18n("ui_description_line_link", [url])];
    if (password) lines.push(i18n("ui_description_line_password", [password]));
    const txt = lines.join("\n");
    if (desc){
      if (desc.tagName && desc.tagName.toLowerCase() === "textarea"){
        desc.value = (desc.value||"") + txt;
        const descDoc = desc.ownerDocument || doc;
        desc.dispatchEvent(new descDoc.defaultView.Event("input",{bubbles:true}));
      } else {
        if (desc.ownerDocument && desc.ownerDocument.execCommand){
          desc.ownerDocument.execCommand("insertText", false, txt);
        } else {
          desc.textContent = (desc.textContent||"") + txt;
        }
      }
    }
  } catch(_){ }
  return placed;
}
function inject(doc, context, label, tooltip) {
  if (!doc) return false;
  const bar = findBar(doc);
  if (!bar) {
    err("toolbar injection skipped: no bar element");
    return false;
  }
  const btn = buildButton(doc, context, label, tooltip);
  ensureMenu(doc, context, btn);
  bar.appendChild(btn);
  return true;
}

function handle(win, context, label, tooltip) {
  if (!isEventDialog(win)) return;
  try {
    inject(win.document, context, label, tooltip);
  } catch(e){
    err(e);
    err("toolbar injection failed");
  }
  const iframe = win.document.getElementById("calendar-item-panel-iframe");
  if (iframe) {
    const run = () => {
      try {
        inject(iframe.contentDocument, context, label, tooltip);
      } catch(e){
        err(e);
        err("toolbar injection failed");
      }
    };
    if (iframe.contentDocument?.readyState === "complete") run();
    iframe.addEventListener("load", run, { once: true });
  }
}

const LISTENER_NAME = "nctalk-caltoolbar";

this.calToolbar = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    LAST_CONTEXT = context;
    resolveBrowser(context);
    ensureDebugState(context);
    const createEvent = new ExtensionCommon.EventManager({
      context,
      name: "calToolbar.onCreateRequest",
      register: (fire) => {
        getCreateHandlers(context).add(fire);
        return () => getCreateHandlers(context).delete(fire);
      }
    });
    const lobbyEvent = new ExtensionCommon.EventManager({
      context,
      name: "calToolbar.onLobbyUpdate",
      register: (fire) => {
        getLobbyHandlers(context).add(fire);
        return () => getLobbyHandlers(context).delete(fire);
      }
    });
    const utilityEvent = new ExtensionCommon.EventManager({
      context,
      name: "calToolbar.onUtilityRequest",
      register: (fire) => {
        getUtilityHandlers(context).add(fire);
        return () => getUtilityHandlers(context).delete(fire);
      }
    });
    return {
      calToolbar: {
        onCreateRequest: createEvent.api(),
        onLobbyUpdate: lobbyEvent.api(),
        onUtilityRequest: utilityEvent.api(),
        async init(opts) {
          const label = (opts && opts.label) || i18n("ui_insert_button_label");
          const tooltip = (opts && opts.tooltip) || i18n("ui_toolbar_tooltip");
          try {
            if (ExtensionSupport && ExtensionSupport.registerWindowListener) {
              ExtensionSupport.registerWindowListener(LISTENER_NAME, {
                chromeURLs: [
                  "chrome://calendar/content/calendar-event-dialog.xhtml",
                  "chrome://calendar/content/calendar-event-dialog.xul",
                  "chrome://calendar/content/calendar-event-dialog"
                ],
                onLoadWindow(win) { try { handle(win, context, label, tooltip); } catch(e) { err(e); } }
              });
              return true;
            }
            return false;
          } catch(e) {
            err(e);
            return false;
          }
        }
      }
    };
  }
}


async function getRoomParticipantsDirect(context, token){
  if (!token) throw localizedError("error_room_token_missing");
  const { baseUrl, user, appPass } = await getCredentials(context);
  if (!baseUrl || !user || !appPass) throw localizedError("error_credentials_missing");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const url = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/participants?includeStatus=true";
  const res = await fetch(url, { method:"GET", headers });
  const raw = await res.text();
  let data = null;
  try { data = raw ? JSON.parse(raw) : null; } catch(_){ }
  if (res.status === 404){
    return [];
  }
  if (!res.ok){
    const meta = data?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw localizedError("error_ocs", [detail]);
  }
  const participants = data?.ocs?.data;
  return Array.isArray(participants) ? participants : [];
}

































