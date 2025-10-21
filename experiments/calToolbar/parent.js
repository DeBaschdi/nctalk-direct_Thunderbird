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
    }
  }catch(e){
    err(e);
  }
  try{
    if (typeof globalThis.atob !== "function" && typeof ChromeUtils?.base64URLDecode === "function"){
      globalThis.atob = function(input){
        const clean = String(input ?? "").replace(/[\r\n\s]/g, "");
        const normalized = clean.replace(/\+/g, "-").replace(/\//g, "_");
        const bytes = ChromeUtils.base64URLDecode(normalized, { padding: "ignore" });
        let binary = "";
        for (const b of bytes){
          binary += String.fromCharCode(b);
        }
        return binary;
      };
    }
    if (typeof globalThis.btoa !== "function" && typeof ChromeUtils?.base64URLEncode === "function"){
      globalThis.btoa = function(input){
        const str = String(input ?? "");
        const len = str.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++){
          bytes[i] = str.charCodeAt(i) & 0xff;
        }
        const encoded = ChromeUtils.base64URLEncode(bytes, { pad: true });
        return encoded.replace(/-/g, "+").replace(/_/g, "/");
      };
    }
  }catch(e){
    err(e);
  }
}

ensureBrowserGlobals();

const ADDON_TITLE = "Nextcloud Talk";
let LAST_CONTEXT = null;

function log(...a){ try { console.log("[NCExp]", ...a); } catch(_) {} }
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

function showAlert(target, message, title = ADDON_TITLE){
  const doc = getDocFromTarget(target);
  if (!doc || !doc.body){
    try{
      if (Services?.prompt){
        Services.prompt.alert(null, title, String(message));
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
  titleEl.textContent = title || ADDON_TITLE;
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
  okBtn.textContent = "OK";
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

async function getAvatarPreferences(context){
  const keys = ["defaultAvatarMode","defaultAvatarData","defaultAvatarMime"];
  const merge = (target, source) => {
    if (!source) return target;
    for (const key of keys){
      if (source[key] != null && target[key] == null){
        target[key] = source[key];
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
  let mode = typeof out.defaultAvatarMode === "string" ? out.defaultAvatarMode.trim() : "";
  if (!mode || !["addon","custom","none"].includes(mode)){
    mode = "addon";
  }
  const data = (typeof out.defaultAvatarData === "string" && out.defaultAvatarData.length) ? out.defaultAvatarData : null;
  const mime = (typeof out.defaultAvatarMime === "string" && out.defaultAvatarMime.length) ? out.defaultAvatarMime : null;
  return { mode, data, mime };
}

async function resolveDefaultAvatar(context, prefs){
  let avatarPrefs = prefs;
  if (!avatarPrefs){
    avatarPrefs = await getAvatarPreferences(context);
  }
  let mode = avatarPrefs?.mode || "addon";
  if (!["addon","custom","none"].includes(mode)){
    mode = "addon";
  }
  const fallbackPreview = defaultRoomIconURL(context, 48) || talkIconURL(context, 48) || "";
  if (mode === "custom"){
    const data = avatarPrefs?.data;
    const mime = avatarPrefs?.mime || "image/png";
    if (data){
      const previewUrl = `data:${mime};base64,${data}`;
      return {
        mode: "custom",
        hasImage: true,
        previewUrl,
        async getBlob(){
          return base64ToBlob(data, mime);
        }
      };
    }
    mode = "none";
  }
  if (mode === "addon"){
    let cleanupPreview = null;
    try{
      const addonBlob = await fetchAddonAvatarBlob(context);
      if (addonBlob && addonBlob.size > 0){
        const mime = addonBlob.type && addonBlob.type !== "application/octet-stream" ? addonBlob.type : "image/png";
        const arrayBuffer = await addonBlob.arrayBuffer();
        const makeBlob = () => new Blob([arrayBuffer.slice(0)], { type: mime });
        let previewUrl = fallbackPreview;
        try{
          const objectUrl = URL.createObjectURL(makeBlob());
          previewUrl = objectUrl;
          cleanupPreview = () => {
            try { URL.revokeObjectURL(objectUrl); }catch(_){}
          };
        }catch(e){
          err(e);
          previewUrl = fallbackPreview;
        }
        return {
          mode: "addon",
          hasImage: true,
          previewUrl,
          cleanup: cleanupPreview,
          async getBlob(){
            return makeBlob();
          }
        };
      }
    }catch(e){
      err(e);
    }
    return {
      mode: "addon",
      hasImage: false,
      previewUrl: fallbackPreview,
      cleanup: cleanupPreview,
      async getBlob(){
        return null;
      }
    };
  }
  return {
    mode: "none",
    hasImage: false,
    previewUrl: "",
    async getBlob(){
      return null;
    }
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

function defaultRoomIconURL(context, size=48){
  const src = talkIconURL(context, size);
  if (src) return src;
  try { return context.extension.rootURI.spec + "icons/talk-48.png"; } catch(_) { return null; }
}

function buildRoomAvatarUrl(baseUrl, token, version, theme="dark"){
  if (!baseUrl || !token) return null;
  const v = version ? "?v=" + encodeURIComponent(version) : "";
  return baseUrl.replace(/\/$/,"") + "/ocs/v2.php/apps/spreed/api/v1/room/" + encodeURIComponent(token) + "/avatar/" + theme + v;
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

const PENDING_DELEGATION_KEY = "_nctalkPendingModerator";

function queuePendingModerator(win, data){
  if (!win) return;
  if (!data || !data.token || !data.delegateId){
    delete win[PENDING_DELEGATION_KEY];
    return;
  }
  win[PENDING_DELEGATION_KEY] = {
    token: data.token,
    delegateId: data.delegateId,
    displayName: data.displayName || data.delegateId,
    processed: false
  };
}

function getPendingModerator(win){
  if (!win) return null;
  const pending = win[PENDING_DELEGATION_KEY];
  if (!pending || !pending.token || !pending.delegateId) return null;
  return pending;
}

function clearPendingModerator(win){
  if (win && win[PENDING_DELEGATION_KEY]){
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
    const runUpdate = async (forced=false) => {
      try{
        const ts = extractStartTimestamp(targetDoc);
        log("lobby watcher check", ts, state.lastTs, forced);
        if (!ts) return;
        if (!forced && ts === state.lastTs) return;
        state.lastTs = ts;
        const payload = {
          token,
          enableLobby: true,
          startTimestamp: ts
        };
        await performLobbyUpdate(context, payload);
        if (forced){
          const pending = getPendingModerator(win);
          if (pending && !pending.processed && pending.token === token){
            try{
              const delegateResponse = await requestUtility(context, { type: "delegateModerator", token, newModerator: pending.delegateId });
              if (!delegateResponse || !delegateResponse.ok){
                throw new Error(delegateResponse?.error || "Moderator konnte nicht \u00fcbertragen werden.");
              }
              const delegationResult = delegateResponse.result || { delegate: pending.delegateId };
              pending.processed = true;
              const delegateName = pending.displayName || delegationResult.delegate || pending.delegateId;
              let message = "Moderator \u00fcbertragen an: " + delegateName + ".";
              if (delegationResult.leftSelf){
                message += "\nSie wurden aus der Unterhaltung entfernt.";
              }
              showAlert(win, message);
              clearPendingModerator(win);
              if (delegationResult.leftSelf){
                cleanupAll();
              }
            }catch(delegationErr){
              pending.processed = true;
              showAlert(win, "Moderator konnte nicht \u00fcbertragen werden:\n" + (delegationErr?.message || delegationErr));
              clearPendingModerator(win);
            }
          }
        }
      }catch(e){ err(e); }
    };
    const handler = (forced=false) => {
      if (state.debounce){
        try { win.clearTimeout(state.debounce); }catch(_){}
        state.debounce = null;
      }
      if (forced){
        runUpdate(true);
        return;
      }
      state.debounce = win.setTimeout(() => {
        try{
          runUpdate(false);
        }catch(e){ err(e); }
      }, 400);
    };
    const trigger = (ev) => {
      const type = ev?.type || "mutation";
      const forced = type === "dialogaccept" || type === "dialogextra1";
      if (forced) state.lastTs = null;
      log("lobby watcher event", type, forced);
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
        log("lobby watcher attached", sel);
      }
    }
    const winEvents = ["dialogaccept","dialogextra1","DOMContentLoaded","unload"];
    for (const evt of winEvents){
      const listener = trigger.bind(null,{type:evt});
      try {
        win.addEventListener(evt, listener, true);
        state.cleanup.push(() => win.removeEventListener(evt, listener, true));
        log("lobby watcher listening", evt);
      }catch(_){}
    }
    const pollInterval = win.setInterval(() => trigger({ type: "poll" }), 4000);
    state.cleanup.push(() => win.clearInterval(pollInterval));
    handler(false);
  }catch(e){ err(e); }
}

/**
 * Aktualisiert den Lobby-Status eines Raums. Behandelt Messaging-Fallbacks.
 */
async function performLobbyUpdate(context, payload){
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
    throw new Error("runtime messaging unavailable");
  }
  return runtime.sendMessage({ type, payload });
}

async function listPublicRoomsDirect(context, searchTerm = ""){
  const { baseUrl, user, appPass } = await getCredentials(context);
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugang fehlt (URL/Nutzer/App-Pass).");
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
    throw new Error("OCS-Fehler: " + detail);
  }
  const rooms = data?.ocs?.data;
  return Array.isArray(rooms) ? rooms : [];
}

async function getRoomInfoDirect(context, token){
  if (!token) throw new Error("Raum-Token fehlt.");
  const { baseUrl, user, appPass } = await getCredentials(context);
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugang fehlt (URL/Nutzer/App-Pass).");
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
    throw new Error("OCS-Fehler: " + detail);
  }
  const room = data?.ocs?.data;
  if (!room || typeof room !== "object"){
    throw new Error("Raumdetails fehlen im Response.");
  }
  return room;
}

async function updateLobbyDirect(context, { token, enableLobby, startTimestamp } = {}){
  if (!token) throw new Error("Raum-Token fehlt.");
  const { baseUrl, user, appPass } = await getCredentials(context);
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugang fehlt (URL/Nutzer/App-Pass).");
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
    throw new Error("Lobby-Update fehlgeschlagen: " + res.status);
  }
}

async function fetchListedRooms(context, searchTerm = ""){
  try{
    const res = await requestUtility(context, { type: "listPublicRooms", searchTerm });
    if (res){
      if (res.ok) return Array.isArray(res.rooms) ? res.rooms : [];
      throw new Error(res.error || "Liste der Unterhaltungen nicht verf\u00fcgbar.");
    }
  }catch(e){
    err(e);
  }
  try{
    const res = await sendRuntimeMessage(context, "talkMenu:listPublicRooms", { searchTerm });
    if (res){
      if (res.ok) return Array.isArray(res.rooms) ? res.rooms : [];
      throw new Error(res.error || "Liste der Unterhaltungen nicht verfuegbar.");
    }
  }catch(e){
    err(e);
  }
  return await listPublicRoomsDirect(context, searchTerm);
}

async function fetchRoomDetails(context, token){
  if (!token) throw new Error("Raum-Token fehlt.");
  try{
    const res = await requestUtility(context, { type: "getRoomInfo", token });
    if (res){
      if (res.ok && res.room) return res.room;
      throw new Error(res.error || "Raumdetails konnten nicht geladen werden.");
    }
  }catch(e){
    err(e);
  }
  try{
    const res = await sendRuntimeMessage(context, "talkMenu:getRoomInfo", { token });
    if (res){
      if (res.ok && res.room) return res.room;
      throw new Error(res.error || "Raumdetails konnten nicht geladen werden.");
    }
  }catch(e){
    err(e);
  }
  return await getRoomInfoDirect(context, token);
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
          log("extract start from input", sel, el.value, el.getAttribute && el.getAttribute("value"));
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

async function uploadRoomAvatar(context, token, blob){
  try{
    const base64 = await blobToBase64(blob);
    const res = await requestUtility(context, { type: "setRoomAvatar", token, data: base64, mime: blob.type });
    if (res){
      if (res.ok){
        let outBase64 = base64;
        let outMime = blob.type || "image/png";
        if (res.dataUrl){
          const parts = res.dataUrl.split(",");
          if (parts.length === 2){
            outBase64 = parts[1];
            const meta = parts[0].match(/^data:(.*?);base64$/);
            if (meta && meta[1]) outMime = meta[1];
          }
        }
        return {
          avatarVersion: res.avatarVersion || res.result?.avatarVersion || null,
          base64: outBase64,
          mime: outMime
        };
      }
      throw new Error(res.error || "Avatar konnte nicht gesetzt werden.");
    }
    return await uploadRoomAvatarDirect(context, token, blob, base64);
  }catch(e){
    err(e);
    const fallbackBase64 = await blobToBase64(blob);
    return await uploadRoomAvatarDirect(context, token, blob, fallbackBase64);
  }
}

function loadImageElement(file){
  return new Promise((resolve, reject) => {
    ensureBrowserGlobals();
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error || new Error("Datei konnte nicht gelesen werden."));
    reader.onload = () => {
      let img = null;
      try{
        const hidden = getHiddenDOMWindow();
        if (hidden && typeof hidden.Image === "function"){
          img = new hidden.Image();
        } else if (typeof Image === "function"){
          img = new Image();
        }
      }catch(e){
        err(e);
      }
      if (!img){
        reject(new Error("Bildunterstuetzung nicht verfuegbar."));
        return;
      }
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error("Bild konnte nicht geladen werden."));
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

async function prepareAvatarBlob(file, size = 512){
  if (!file) throw new Error("Keine Bilddatei ausgewaehlt.");
  ensureBrowserGlobals();
  const img = await loadImageElement(file);
  const hidden = getHiddenDOMWindow();
  const doc = (hidden && hidden.document) ? hidden.document : (typeof document !== "undefined" ? document : null);
  if (!doc || typeof doc.createElement !== "function"){
    throw new Error("Canvas-Unterstuetzung nicht verfuegbar.");
  }
  const canvas = doc.createElement("canvas");
  canvas.width = size;
  canvas.height = size;
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0,0,size,size);
  const scale = Math.min(size / img.width, size / img.height);
  const w = Math.round(img.width * scale);
  const h = Math.round(img.height * scale);
  const dx = Math.round((size - w) / 2);
  const dy = Math.round((size - h) / 2);
  ctx.drawImage(img, dx, dy, w, h);
  return await new Promise((resolve, reject) => {
    if (typeof canvas.toBlob === "function"){
      canvas.toBlob((blob) => {
        if (blob) resolve(blob); else reject(new Error("Avatar konnte nicht vorbereitet werden."));
      }, "image/png", 0.92);
      return;
    }
    try{
      const dataUrl = canvas.toDataURL("image/png", 0.92);
      const base64 = (dataUrl.split(",")[1] || "").trim();
      const blob = base64ToBlob(base64, "image/png");
      if (blob) resolve(blob); else reject(new Error("Avatar konnte nicht vorbereitet werden."));
    }catch(e){
      reject(e);
    }
  });
}

async function blobToBase64(blob){
  const buffer = await blob.arrayBuffer();
  let binary = "";
  const bytes = new Uint8Array(buffer);
  for (const b of bytes){
    binary += String.fromCharCode(b);
  }
  return btoa(binary);
}

function base64ToBlob(base64, mime = "image/png"){
  try{
    ensureBrowserGlobals();
    if (!base64) return null;
    const clean = String(base64).replace(/[\r\n\s]/g, "");
    let bytes = null;
    try{
      if (typeof ChromeUtils?.base64URLDecode === "function"){
        const normalized = clean.replace(/\+/g, "-").replace(/\//g, "_");
        bytes = ChromeUtils.base64URLDecode(normalized, { padding: "ignore" });
      }
    }catch(e){
      err(e);
      bytes = null;
    }
    if (!bytes){
      let binary = null;
      if (typeof globalThis.atob === "function"){
        binary = globalThis.atob(clean);
      } else if (typeof Components !== "undefined" && Components.utils && typeof Components.utils.atob === "function"){
        binary = Components.utils.atob(clean);
      }
      if (binary != null){
        const len = binary.length;
        bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++){
          bytes[i] = binary.charCodeAt(i);
        }
      }
    }
    if (!bytes) return null;
    return new Blob([bytes], { type: mime || "image/png" });
  }catch(e){
    err(e);
    return null;
  }
}

async function fetchAddonAvatarBlob(context){
  const sizes = [128, 96, 64, 48, 32, 24];
  for (const size of sizes){
    const url = talkIconURL(context, size);
    if (!url) continue;
    try{
      const res = await fetch(url);
      if (!res.ok) continue;
      const blob = await res.blob();
      if (!blob || blob.size === 0) continue;
      if (!blob.type || blob.type === "application/octet-stream"){
        try{
          const buffer = await blob.arrayBuffer();
          return new Blob([buffer], { type: "image/png" });
        }catch(e){
          err(e);
          return blob;
        }
      }
      return blob;
    }catch(e){
      err(e);
    }
  }
  return null;
}

async function requestLobbyUpdate(context, payload){
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
  const span = doc.createElement("span"); span.textContent = label || "Talk-Link einf\u00fcgen";
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
    const defaultTitle = (doc.querySelector('input[type="text"]')?.value || "Besprechung");
    const overlay = doc.createElement("div");
    Object.assign(overlay.style,{position:"fixed",inset:"0",background:"rgba(0,0,0,.25)",zIndex:"2147483646"});
    const panel = doc.createElement("div");
    Object.assign(panel.style,{position:"fixed",top:"20%",left:"50%",transform:"translateX(-50%)",
      background:"var(--arrowpanel-background,#fff)",border:"1px solid var(--arrowpanel-border-color,#c0c0c0)",
      borderRadius:"8px",boxShadow:"0 10px 30px rgba(0,0,0,.25)",minWidth:"460px",padding:"16px",zIndex:"2147483647"});

    const heading = doc.createElement("h2");
    heading.textContent = "\u00d6ffentliche Unterhaltung erstellen";
    Object.assign(heading.style,{margin:"0 0 10px",font:"600 16px system-ui"});

    const grid = doc.createElement("div");
    Object.assign(grid.style,{display:"grid",gridTemplateColumns:"160px 1fr",gap:"10px",alignItems:"center"});

    const titleLabel = doc.createElement("label");
    titleLabel.textContent = "Titel";
    const titleInput = doc.createElement("input");
    Object.assign(titleInput,{id:"nc_title",type:"text",value:defaultTitle});

    const passLabel = doc.createElement("label");
    passLabel.textContent = "Passwort (optional)";
    const passInput = doc.createElement("input");
    Object.assign(passInput,{id:"nc_pass",type:"password",placeholder:"Optional"});

    const lobbyLabel = doc.createElement("label");
    const lobbyCheckbox = doc.createElement("input");
    Object.assign(lobbyCheckbox,{id:"nc_lobby",type:"checkbox"});
    lobbyCheckbox.checked = true;
    lobbyLabel.appendChild(lobbyCheckbox);
    lobbyLabel.appendChild(doc.createTextNode(" Lobby bis Startzeit"));
    const lobbySpacer = doc.createElement("div");

    const listableLabel = doc.createElement("label");
    const listableCheckbox = doc.createElement("input");
    Object.assign(listableCheckbox,{id:"nc_listable",type:"checkbox"});
    listableCheckbox.checked = true;
    listableLabel.appendChild(listableCheckbox);
    listableLabel.appendChild(doc.createTextNode(" In Suche anzeigen"));
    const listableSpacer = doc.createElement("div");

    const delegateLabel = doc.createElement("label");
    delegateLabel.textContent = "Moderator (optional)";
    delegateLabel.style.alignSelf = "start";
    const delegateField = doc.createElement("div");
    Object.assign(delegateField.style,{display:"flex",flexDirection:"column",gap:"6px",position:"relative"});
    const delegateInput = doc.createElement("input");
    Object.assign(delegateInput,{id:"nc_delegate",type:"text",placeholder:"Benutzername eingeben"});
    delegateInput.dataset.selectionLabel = "";
    delegateInput.autocomplete = "off";
    delegateInput.spellcheck = false;
    const delegateInputRow = doc.createElement("div");
    Object.assign(delegateInputRow.style,{display:"flex",alignItems:"center",gap:"6px"});
    delegateInputRow.appendChild(delegateInput);
    const delegateClearBtn = doc.createElement("button");
    delegateClearBtn.type = "button";
    delegateClearBtn.textContent = "Leeren";
    Object.assign(delegateClearBtn.style,{padding:"4px 10px",fontSize:"12px"});
    delegateInputRow.appendChild(delegateClearBtn);
    delegateField.appendChild(delegateInputRow);
    const delegateStatus = doc.createElement("div");
    Object.assign(delegateStatus.style,{fontSize:"11px",opacity:"0.7",minHeight:"14px"});
    delegateField.appendChild(delegateStatus);
    const delegateSelectedInfo = doc.createElement("div");
    Object.assign(delegateSelectedInfo.style,{fontSize:"11px",opacity:"0.75",minHeight:"14px"});
    delegateField.appendChild(delegateSelectedInfo);
    const delegateHint = doc.createElement("div");
    Object.assign(delegateHint.style,{fontSize:"11px",opacity:"0.75"});
    delegateHint.textContent = "Bei Angabe wird die Moderation nach Erstellung an diesen Benutzer \u00fcbertragen und Sie verlassen den Raum.";
    delegateField.appendChild(delegateHint);
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
      maxHeight:"220px",
      overflowY:"auto",
      display:"none",
      zIndex:"2147483647"
    });
    delegateField.appendChild(delegateDropdown);

    const adjustDelegateDropdownPosition = () => {
      delegateDropdown.style.maxHeight = "220px";
      delegateDropdown.style.top = "calc(100% + 4px)";
      delegateDropdown.style.bottom = "";
      try{
        const panelRect = panel.getBoundingClientRect();
        const fieldRect = delegateField.getBoundingClientRect();
        const dropdownRect = delegateDropdown.getBoundingClientRect();
        const margin = 12;
        const availableBelow = panelRect.bottom - fieldRect.bottom - margin;
        const availableAbove = fieldRect.top - panelRect.top - margin;
        if (dropdownRect.bottom > panelRect.bottom - margin){
          if (availableAbove > availableBelow){
            delegateDropdown.style.top = "";
            delegateDropdown.style.bottom = "calc(100% + 4px)";
            const maxHeight = Math.max(120, availableAbove);
            delegateDropdown.style.maxHeight = maxHeight + "px";
          } else {
            delegateDropdown.style.top = "calc(100% + 4px)";
            delegateDropdown.style.bottom = "";
            const maxHeight = Math.max(120, availableBelow);
            delegateDropdown.style.maxHeight = maxHeight + "px";
          }
        } else if (availableBelow > 0){
          delegateDropdown.style.maxHeight = Math.max(120, availableBelow) + "px";
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
      delegateDropdown.style.maxHeight = "220px";
      delegateDropdown.style.top = "calc(100% + 4px)";
      delegateDropdown.style.bottom = "";
      delegateSuggestionsState.visible = false;
      delegateSuggestionsState.activeIndex = -1;
    };

    const renderDelegateDropdown = () => {
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
          flexDirection:"column",
          gap:"2px",
          background: index === delegateSuggestionsState.activeIndex ? "var(--arrowpanel-dimmed, rgba(0,0,0,0.08))" : "transparent"
        });
        row.dataset.index = String(index);
        const primary = doc.createElement("div");
        primary.textContent = item.label || item.id || item.email || "";
        primary.style.fontSize = "12px";
        const emailLine = doc.createElement("div");
        emailLine.textContent = item.email || "";
        Object.assign(emailLine.style,{fontSize:"11px",opacity:"0.75"});
        row.appendChild(primary);
        if (item.email){
          row.appendChild(emailLine);
        }
        if (item.id && item.id !== item.email){
          const idLine = doc.createElement("div");
          idLine.textContent = item.id;
          Object.assign(idLine.style,{fontSize:"10px",opacity:"0.6"});
          row.appendChild(idLine);
        }
        row.addEventListener("mouseenter", () => {
          delegateSuggestionsState.activeIndex = index;
          renderDelegateDropdown();
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

    const selectDelegateSuggestion = (index) => {
      const suggestion = delegateSuggestionsState.items[index];
      if (!suggestion) return;
      delegateInput.value = suggestion.id;
      const selectionLabel = formatDelegateDisplay(suggestion);
      delegateInput.dataset.selectionLabel = selectionLabel;
      delegateSelectedInfo.textContent = selectionLabel ? "Ausgew\u00e4hlt: " + selectionLabel : "";
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
        updateDelegateStatus(term ? "Suche..." : "Lade Benutzer...");
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
            updateDelegateStatus(term ? "Keine Treffer mit E-Mail." : "Keine Benutzer mit E-Mail gefunden.");
            hideDelegateDropdown();
          }else{
            const countText = items.length === 1 ? "1 Treffer mit E-Mail." : items.length + " Treffer mit E-Mail.";
            updateDelegateStatus(countText);
            renderDelegateDropdown();
          }
        }catch(e){
          if (seq !== delegateSearchSeq) return;
          err(e);
          updateDelegateStatus(e?.message || "Benutzersuche fehlgeschlagen.", true);
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
        renderDelegateDropdown();
      } else if (event.key === "ArrowUp"){
        event.preventDefault();
        const count = delegateSuggestionsState.items.length;
        delegateSuggestionsState.activeIndex = (delegateSuggestionsState.activeIndex - 1 + count) % count;
        renderDelegateDropdown();
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
      delegateSelectedInfo.textContent = delegateInput.dataset.selectionLabel ? "Ausgew\u00e4hlt: " + delegateInput.dataset.selectionLabel : "";
      scheduleDelegateSearch(delegateInput.value.trim());
    });

    delegateInput.addEventListener("input", () => {
      delegateInput.dataset.selectionLabel = "";
      delegateSelectedInfo.textContent = "";
      scheduleDelegateSearch(delegateInput.value.trim());
    });

    delegateInput.addEventListener("keydown", handleDelegateKeyDown);
    delegateInput.addEventListener("blur", () => {
      docWin.setTimeout(() => hideDelegateDropdown(), 80);
    });

    delegateClearBtn.addEventListener("click", () => {
      delegateInput.value = "";
      delegateInput.dataset.selectionLabel = "";
      delegateSelectedInfo.textContent = "";
      updateDelegateStatus("");
      scheduleDelegateSearch("");
      delegateInput.focus();
    });

    doc.addEventListener("mousedown", handleDelegateDocumentClick, true);
    scheduleDelegateSearch("");

    const cleanupDialogState = () => {
      updateDelegateStatus("");
      delegateSelectedInfo.textContent = "";
      if (delegateSearchTimeout){
        docWin.clearTimeout(delegateSearchTimeout);
        delegateSearchTimeout = null;
      }
      delegateSuggestionsState.items = [];
      delegateSuggestionsState.activeIndex = -1;
      hideDelegateDropdown();
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
    grid.appendChild(delegateLabel);
    grid.appendChild(delegateField);

    const buttons = doc.createElement("div");
    Object.assign(buttons.style,{display:"flex",justifyContent:"flex-end",gap:"8px",marginTop:"16px"});

    const cancelBtn = doc.createElement("button");
    cancelBtn.id = "nc_cancel";
    cancelBtn.textContent = "Abbrechen";

    const okBtn = doc.createElement("button");
    okBtn.id = "nc_ok";
    okBtn.textContent = "OK";
    okBtn.style.appearance = "auto";

    buttons.appendChild(cancelBtn);
    buttons.appendChild(okBtn);

    panel.appendChild(heading);
    panel.appendChild(grid);
    panel.appendChild(buttons);

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
        showAlert(docWin, "Das Passwort muss mindestens 5 Zeichen lang sein.");
        return;
      }
      okBtn.disabled = true;
      okBtn.textContent = "Erstelle...";

      const restore = () => {
        okBtn.disabled = false;
        okBtn.textContent = originalLabel;
      };

      const innerDoc = (doc.getElementById && doc.getElementById("calendar-item-panel-iframe") && doc.getElementById("calendar-item-panel-iframe").contentDocument) || doc;
      const descriptionText = extractDescriptionText(innerDoc);

      const startTs = extractStartTimestamp(innerDoc);
      const delegateId = delegateInput.value.trim();

      let response = null;
      try {
        response = await requestCreateFromExtension(context, {
          title,
          password: password || undefined,
          enableLobby,
          enableListable,
          description: descriptionText || "",
          startTimestamp: startTs
        });
      } catch(sendErr) {
        err(sendErr);
        restore();
        showAlert(docWin, "Senden an Hintergrundskript fehlgeschlagen:\n" + (sendErr?.message || String(sendErr)));
        return;
      }

      if (!response) {
        restore();
        showAlert(docWin, "Kein Nextcloud Talk Handler registriert. Bitte Add-on neu laden.");
        return;
      }

      if (!response.ok || !response.url) {
        restore();
        const msg = response.error ? response.error : "Unbekannter Fehler beim Erstellen der Unterhaltung.";
        showAlert(docWin, "Nextcloud Talk konnte nicht erstellt werden:\n" + msg + "\nBitte Optionen pruefen.");
        return;
      }

      clearPendingModerator(docWin);

      let delegationResult = null;
      let pendingDelegation = null;
      const delegateDisplayName = delegateInput.dataset.selectionLabel || delegateId;
      if (delegateId){
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
              const msg = delegateResponse?.error || "Moderator konnte nicht \u00fcbertragen werden.";
              showAlert(docWin, msg);
            }else{
              delegationResult = delegateResponse.result || { delegate: delegateId };
            }
          }catch(delegateErr){
            err(delegateErr);
            showAlert(docWin, "Moderator konnte nicht \u00fcbertragen werden:\n" + (delegateErr?.message || delegateErr));
          }
        }
      }

      closeOverlay();
      const placed = fillIntoEvent(innerDoc, response.url, password || null, title);
      let alertMsg = "Talk-Link eingefuegt:\n" + response.url;
      if (!placed) alertMsg += "\n(Hinweis: Feld 'Ort' wurde nicht automatisch gefunden.)";
      if (response.fallback) {
        alertMsg += "\n(Hinweis: Es wurde ein Fallback-Link ohne API erzeugt.)";
        if (response.reason) alertMsg += "\nGrund: " + response.reason;
      }
      if (pendingDelegation){
        alertMsg += "\nModerator wird beim Speichern/Senden \u00fcbertragen an: " + (pendingDelegation.displayName || pendingDelegation.delegateId) + ".";
      } else if (delegationResult && delegationResult.delegate){
        alertMsg += "\nModerator \u00fcbertragen an: " + (delegateDisplayName || delegationResult.delegate) + ".";
        if (delegationResult.leftSelf){
          alertMsg += "\nSie wurden aus der Unterhaltung entfernt.";
        }
      }
      showAlert(docWin, alertMsg);
      const watcherAllowed = enableLobby && response.token;
      if (watcherAllowed){
        setupLobbyWatcher(context, doc, innerDoc, response.token, true);
      }
    });

    overlay.appendChild(panel);
    (doc.body||doc.documentElement).appendChild(overlay);
    okBtn.focus();
    log("create dialog shown");
  }catch(e){
    err(e);
    const docWin = doc.defaultView || window;
    const base = await getBaseUrl(context) || docWin.prompt("Nextcloud Basis-URL (z.B. https://cloud.example.com):","https://cloud.example.com");
    if (!base) return;
    const title = docWin.prompt("Titel fuer neue oeffentliche Unterhaltung:","Besprechung");
    if (title === null) return;
    const token = randToken(10);
    const url = String(base).replace(/\/$/,"") + "/call/" + token;
    const ok = fillIntoEvent(doc, url, null);
    showAlert(docWin, "Talk-Link eingef\u00fcgt:\n" + url + (ok ? "" : "\n(Hinweis: Feld 'Ort' wurde nicht automatisch gefunden.)"));
  }
}
async function openSelectDialog(doc, context){
  let overlay = null;
  try{
    const docWin = doc.defaultView || window;
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
    heading.textContent = "\u00d6ffentliche Unterhaltung ausw\u00e4hlen";
    Object.assign(heading.style,{margin:"0 0 10px",font:"600 16px system-ui"});
    panel.appendChild(heading);

    const searchBox = doc.createElement("div");
    Object.assign(searchBox.style,{marginBottom:"8px"});
    const searchInput = doc.createElement("input");
    Object.assign(searchInput,{type:"search",placeholder:"Suche"});
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
    detailTitle.textContent = "Keine Unterhaltung ausgewaehlt";
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
    lobbyRow.appendChild(doc.createTextNode(" Lobby einschalten (Startzeit aus Termin)"));
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
      showAlert(docWin, "Bitte hinterlegen Sie die Nextcloud URL in den Add-on-Optionen.");
      return;
    }
    const normalizedBaseUrl = String(baseUrl).replace(/\/$/,"");

    const buttons = doc.createElement("div");
    Object.assign(buttons.style,{display:"flex",justifyContent:"flex-end",gap:"8px",marginTop:"14px"});
    panel.appendChild(buttons);

    const cancelBtn = doc.createElement("button");
    cancelBtn.textContent = "Abbrechen";
    const okBtn = doc.createElement("button");
    okBtn.textContent = "\u00dcbernehmen";
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
        lobbyInfo.textContent = "Lobby deaktiviert.";
        return;
      }
      if (typeof timer === "number" && Number.isFinite(timer) && timer > 0){
        const date = new Date(timer * 1000);
        lobbyInfo.textContent = "Lobby aktiv (Start: " + date.toLocaleString() + ")";
      } else {
        lobbyInfo.textContent = "Lobby aktiv (keine Startzeit gesetzt).";
      }
    }

    function clearDetails(){
      detailTitle.textContent = "Keine Unterhaltung ausgewaehlt";
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
        setStatus("Keine Unterhaltungen gefunden.");
        clearDetails();
        return;
      }
      const countText = rooms.length === 1 ? "1 Unterhaltung gefunden." : rooms.length + " Unterhaltungen gefunden.";
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
        if (room.hasPassword) badges.push("Passwort erforderlich");
        if (room.listable){
          badges.push("Oeffentlich gelistet");
        } else if (room.source === "own" || room.isParticipant){
          badges.push("Eigene Teilnahme");
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
      if (room.listable) metaParts.push("Oeffentlich gelistet");
      else if (room.source === "own" || room.isParticipant) metaParts.push("Eigene Teilnahme");
      if (room.hasPassword) metaParts.push("Passwort erforderlich");
      if (room.guestsAllowed !== false) metaParts.push("Gaeste erlaubt");
      detailMeta.textContent = metaParts.join(" | ");
      linkRow.textContent = normalizedBaseUrl + "/call/" + currentToken;
      descLabel.textContent = (room.description || "").trim();
      passwordInfo.textContent = room.hasPassword ? "Passwortschutz aktiv." : "Kein Passwort gesetzt.";
      lobbyInfo.textContent = "Lade Raumdetails...";
      okBtn.disabled = true;
      const requestToken = ++detailRequestToken;

      try{
        const rawDetails = await fetchRoomDetails(context, currentToken);
        if (detailRequestToken !== requestToken) return;
        const normalized = Object.assign({}, room, rawDetails || {});
        normalized.token = normalized.token || currentToken;
        if (room._avatarObjectUrl && !normalized._avatarObjectUrl){
          normalized._avatarObjectUrl = room._avatarObjectUrl;
        }
        if (room._avatarBase64 && !normalized._avatarBase64){
          normalized._avatarBase64 = room._avatarBase64;
          normalized._avatarMime = room._avatarMime;
        }
        if (!normalized.avatarVersion && room.avatarVersion){
          normalized.avatarVersion = room.avatarVersion;
        }
        selectedDetails = normalized;
        detailTitle.textContent = normalized.displayName || normalized.name || normalized.token;
        const normalizedMeta = [];
        if (normalized.listable) normalizedMeta.push("Oeffentlich gelistet");
        else if (normalized.source === "own" || normalized.isParticipant) normalizedMeta.push("Eigene Teilnahme");
        if (normalized.hasPassword) normalizedMeta.push("Passwort erforderlich");
        normalizedMeta.push(normalized.guestsAllowed === false ? "Keine Gaeste erlaubt" : "Gaeste erlaubt");
        detailMeta.textContent = normalizedMeta.join(" | ");
        linkRow.textContent = normalizedBaseUrl + "/call/" + normalized.token;
        descLabel.textContent = (normalized.description || "").trim();
        passwordInfo.textContent = normalized.hasPassword ? "Passwortschutz aktiv." : "Kein Passwort gesetzt.";
        desiredLobbyState = normalized.lobbyState === 1;
        lobbyCheckbox.checked = desiredLobbyState;
        lobbyInfo.textContent = "Lobby-Status wird ermittelt...";
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
          moderatorInfo.textContent = "Keine Teilnahme an dieser Unterhaltung (nur Anzeige).";
        }
        lobbyCheckbox.disabled = !isModerator;
        lobbyCheckbox.title = isModerator ? "" : "Keine Berechtigung zum Aendern der Lobby.";
        moderatorInfo.textContent = isModerator ? "Sie koennen die Lobby verwalten." : "Sie sind kein Moderator dieser Unterhaltung.";
        desiredLobbyState = lobbyCheckbox.checked;
        updateLobbyInfo(normalized.lobbyTimer, desiredLobbyState);
        okBtn.disabled = false;
        const entry = rooms.find(r => r.token === normalized.token);
        if (entry){
          entry.description = normalized.description || entry.description;
          entry.hasPassword = !!normalized.hasPassword;
          entry.avatarVersion = normalized.avatarVersion || entry.avatarVersion;
          if (normalized._avatarObjectUrl){
            entry._avatarObjectUrl = normalized._avatarObjectUrl;
          }
          if (normalized._avatarBase64){
            entry._avatarBase64 = normalized._avatarBase64;
            entry._avatarMime = normalized._avatarMime;
          }
        }
      }catch(fetchErr){
        err(fetchErr);
        if (detailRequestToken !== requestToken) return;
        lobbyInfo.textContent = "Details konnten nicht geladen werden.";
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
      setStatus("Lade...");
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
        renderList();
      }catch(loadErr){
        err(loadErr);
        if (seq !== searchSeq) return;
        setStatus("Fehler: " + (loadErr?.message || loadErr));
      }
    }

    await loadRooms("");

    okBtn.addEventListener("click", async () => {
      if (!selectedDetails) return;
      const originalLabel = okBtn.textContent;
      okBtn.disabled = true;
      okBtn.textContent = "\u00dcbernehme...";
      try{
        const title = selectedDetails.displayName || selectedDetails.name || "Besprechung";
        const roomUrl = normalizedBaseUrl + "/call/" + selectedDetails.token;
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

        let message = "Talk-Link eingef\u00fcgt:\n" + roomUrl;
        if (!placed) message += "\n(Hinweis: Feld 'Ort' wurde nicht automatisch gefunden.)";
        if (selectedDetails.hasPassword) message += "\nHinweis: Diese Unterhaltung ist passwortgeschuetzt.";
        message += "\nLobby ist " + (lobbyStateForWatcher ? "aktiv." : "deaktiviert.");
        if (!isModerator){
          message += "\n(Hinweis: Keine Lobby-Rechte, Zustand unveraendert.)";
        }
        showAlert(docWin, message);
      }catch(applyErr){
        const applyMsg = applyErr?.message || String(applyErr);
        const permissionProblem = applyMsg && (applyMsg.toLowerCase().includes("berechtigung") || applyMsg.includes("403"));
        if (!permissionProblem){
          err(applyErr);
        }
        okBtn.disabled = false;
        okBtn.textContent = originalLabel;
        const alertMsg = permissionProblem
          ? "Lobby konnte nicht angepasst werden: Keine Moderatorrechte."
          : "\u00dcbernahme fehlgeschlagen:\n" + applyMsg;
        if (permissionProblem && selectedDetails){
          const currentLobby = selectedDetails.lobbyState === 1;
          desiredLobbyState = currentLobby;
          lobbyCheckbox.checked = currentLobby;
          lobbyCheckbox.disabled = true;
          lobbyCheckbox.title = "Keine Berechtigung zum Aendern der Lobby.";
          moderatorInfo.textContent = "Sie sind kein Moderator dieser Unterhaltung.";
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
    showAlert(docWin, "\u00d6ffentliche Unterhaltungen konnten nicht geladen werden:\n" + (e?.message || e));
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
    const lines = ["\nTalk: " + url];
    if (password) lines.push("Passwort: " + password);
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
  if (!bar) { log("no bar"); return false; }
  const btn = buildButton(doc, context, label, tooltip);
  ensureMenu(doc, context, btn);
  bar.appendChild(btn);
  log("button injected (direct create)");
  return true;
}

function handle(win, context, label, tooltip) {
  if (!isEventDialog(win)) return;
  try { inject(win.document, context, label, tooltip); } catch(e){ err(e); }
  const iframe = win.document.getElementById("calendar-item-panel-iframe");
  if (iframe) {
    const run = () => { try { inject(iframe.contentDocument, context, label, tooltip); } catch(e){ err(e); } };
    if (iframe.contentDocument?.readyState === "complete") run();
    iframe.addEventListener("load", run, { once: true });
  }
}

const LISTENER_NAME = "nctalk-caltoolbar";

this.calToolbar = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    LAST_CONTEXT = context;
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
          const label = (opts && opts.label) || "Talk-Link einf\u00fcgen";
          const tooltip = (opts && opts.tooltip) || "Nextcloud Talk";
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
  if (!token) throw new Error("Raum-Token fehlt.");
  const { baseUrl, user, appPass } = await getCredentials(context);
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugang fehlt (URL/Nutzer/App-Pass).");
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
    throw new Error("OCS-Fehler: " + detail);
  }
  const participants = data?.ocs?.data;
  return Array.isArray(participants) ? participants : [];
}

async function uploadRoomAvatarDirect(context, token, blob, base64Override){
  if (!token || !blob) throw new Error("Avatar-Daten fehlen.");
  const { baseUrl, user, appPass } = await getCredentials(context);
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugang fehlt (URL/Nutzer/App-Pass).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const finalMime = blob.type || "image/png";
  const base = baseUrl.replace(/\/$/,"");
  const base64 = base64Override || await blobToBase64(blob);
  const dataUrl = "data:" + finalMime + ";base64," + base64;

  const uploadWithFormData = async () => {
    const headers = {
      "OCS-APIRequest": "true",
      "Authorization": auth,
      "Accept": "application/json",
      "X-Requested-With": "XMLHttpRequest"
    };
    const form = new FormData();
    const payloadBlob = base64ToBlob(base64, finalMime);
    if (!payloadBlob) throw new Error("Avatar konnte nicht vorbereitet werden.");
    const filename = finalMime === "image/jpeg" ? "avatar.jpg" : "avatar.png";
    form.append("file", payloadBlob, filename);
    const url = base + "/ocs/v2.php/apps/spreed/api/v1/room/" + encodeURIComponent(token) + "/avatar";
    const res = await fetch(url, { method:"POST", headers, body: form });
    const raw = await res.text().catch(() => "");
    let data = null;
    try { data = raw ? JSON.parse(raw) : null; } catch(_){ }
    if (!res.ok){
      const meta = data?.ocs?.meta || {};
      const detail = meta.message || raw || (res.status + " " + res.statusText);
      throw new Error(detail || "Avatar-Upload fehlgeschlagen.");
    }
    return {
      avatarVersion: data?.ocs?.data?.avatarVersion || data?.ocs?.data?.version || null,
      base64,
      mime: finalMime
    };
  };

  const uploadWithJson = async () => {
    const headers = {
      "OCS-APIRequest": "true",
      "Authorization": auth,
      "Accept": "application/json",
      "Content-Type": "application/json",
      "X-Requested-With": "XMLHttpRequest"
    };
    const url = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/avatar";
    const res = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify({ image: base64, mimetype: finalMime })
    });
    const raw = await res.text().catch(() => "");
    let data = null;
    try { data = raw ? JSON.parse(raw) : null; } catch(_){ }
    if (!res.ok){
      const meta = data?.ocs?.meta || {};
      const detail = meta.message || raw || (res.status + " " + res.statusText);
      throw new Error(detail || "Avatar-Upload fehlgeschlagen.");
    }
    return {
      avatarVersion: data?.ocs?.data?.avatarVersion || data?.ocs?.data?.version || null,
      base64,
      mime: finalMime
    };
  };

  try{
    return await uploadWithFormData();
  }catch(e){
    err(e);
    return await uploadWithJson();
  }
}
