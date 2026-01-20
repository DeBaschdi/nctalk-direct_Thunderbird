/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
(() => {
  "use strict";

  const API = window.NCTalkBridge;
  const LOG_CHANNEL = "[NCUI][Talk]";
  if (!API){
    console.error(`${LOG_CHANNEL} Bridge nicht verfügbar.`);
    return;
  }

  const CalUtils = window.NCTalkCalUtils;
  if (!CalUtils){
    console.error(`${LOG_CHANNEL} Shared utils missing.`);
    return;
  }

  const EVENT_DIALOG_URLS = [
    "chrome://calendar/content/calendar-event-dialog.xhtml",
    "chrome://calendar/content/calendar-event-dialog.xul"
  ];

  const STATE = {
    label: null,
    tooltip: null
  };

  const ROOM_CLEANUP_KEY = "_nctalkRoomCleanup";

  function log(...args){
    try{
      console.log(LOG_CHANNEL, ...args);
    }catch(_){}
  }
  function err(...args){
    try{
      console.error(LOG_CHANNEL, ...args);
    }catch(_){}
  }

  function i18n(key, substitutions = []){
    try{
      const message = API.i18n(key, substitutions);
      if (message){
        return message;
      }
    }catch(_){}
    if (Array.isArray(substitutions) && substitutions.length){
      return String(substitutions[0] ?? "");
    }
    return key;
  }

  async function syncConfigState(context){
    try{
      await requestUtility(context, { type: "getConfig" });
    }catch(_){}
  }

  function getTalkMetadataFromWindow(win){
    const metadata = CalUtils.readTalkMetadataFromDocument(win.document);
    return { ok:true, metadata };
  }

  function setTalkMetadataOnWindow(win, payload = {}){
    return CalUtils.setTalkMetadataOnWindow(win, payload);
  }

  function getEventSnapshotFromWindow(win){
    return CalUtils.getEventSnapshotFromWindow(win);
  }

  function applyEventFieldsOnWindow(win, payload = {}){
    return CalUtils.applyEventFieldsOnWindow(win, payload);
  }

  function extractStartTimestamp(doc){
    try{
      const docs = CalUtils.collectEventDocs(doc);
      for (const d of docs){
        const picker = d.querySelector && (d.querySelector("datetimepicker#event-starttime") || d.querySelector("datetimepicker#item-starttime"));
        if (picker){
          try{
            let value = picker.value || (picker.getAttribute && picker.getAttribute("value"));
            if (!value && "valueAsDate" in picker) value = picker.valueAsDate;
            if (value instanceof Date || (typeof value === "object" && typeof value.getTime === "function")){
              const ts = value.getTime();
              if (!Number.isNaN(ts)) return Math.floor(ts / 1000);
            }else if (typeof value === "string"){
              const parsed = Date.parse(value);
              if (!Number.isNaN(parsed)) return Math.floor(parsed / 1000);
            }
          }catch(_){}
        }
      }
    }catch(_){}
    try{
      const item = CalUtils.getCalendarItemFromDocument(doc);
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
    return null;
  }

  function getRoomCleanupState(win){
    if (!win) return null;
    let state = win[ROOM_CLEANUP_KEY];
    if (!state){
      state = { token:null, context:null, info:null, cleanup:[], deleting:false };
      win[ROOM_CLEANUP_KEY] = state;
    }
    return state;
  }

  function cleanupRoomCleanupState(state){
    if (!state || !Array.isArray(state.cleanup)) return;
    while (state.cleanup.length){
      const fn = state.cleanup.pop();
      try{ fn(); }catch(_){}
    }
  }

  function resetRoomCleanupState(state){
    if (!state) return;
    cleanupRoomCleanupState(state);
    state.token = null;
    state.context = null;
    state.info = null;
    state.deleting = false;
  }

  function triggerRoomCleanupDelete(state, reason){
    if (!state || state.deleting) return;
    const token = state.token;
    const context = state.context;
    if (!token || !context) return;
    state.deleting = true;
    cleanupRoomCleanupState(state);
    state.token = null;
    const info = state.info || {};
    state.info = null;
    log("room cleanup delete request", { token, reason: reason || "", fallback: !!info.fallback });
    (async () => {
      try{
        const result = await requestUtility(context, { type: "deleteRoom", token });
        if (!result || !result.ok){
          log("room cleanup delete failed", { token, reason: reason || "", error: result?.error || "" });
        }else{
          log("room cleanup delete success", { token });
        }
      }catch(e){
        err(e);
      }finally{
        state.deleting = false;
        state.context = null;
      }
    })();
  }

  function registerRoomCleanup(win, context, token, info = {}){
    if (!win || !context || !token) return;
    const state = getRoomCleanupState(win);
    if (!state) return;
    if (state.token && state.token !== token){
      triggerRoomCleanupDelete(state, "superseded");
    }else{
      cleanupRoomCleanupState(state);
    }
    state.context = context;
    state.token = token;
    state.info = info || {};
    state.deleting = false;

    const addListener = (target, type, handler, options) => {
      if (!target || typeof target.addEventListener !== "function") return;
      target.addEventListener(type, handler, options);
      state.cleanup.push(() => {
        if (typeof target?.removeEventListener === "function"){
          target.removeEventListener(type, handler, options);
        }
      });
    };

    const markPersisted = () => {
      try{
        const meta = CalUtils.readTalkMetadataFromDocument(win.document);
        const startTs = extractStartTimestamp(win.document);
        log("room cleanup metadata snapshot", {
          token: meta?.token || "",
          delegateId: meta?.delegateId || "",
          delegateName: meta?.delegateName || "",
          delegated: meta?.delegated === true,
          delegateReady: meta?.delegateReady ?? null,
          startTimestamp: startTs ?? null
        });
        const pendingDelegate = (meta?.delegateId || "").trim();
        if (meta?.token && pendingDelegate && meta.delegated !== true){
          CalUtils.writeTalkMetadataToDocument(win.document, { delegateReady: true });
          log("delegate armed for calendar flow", {
            token: meta.token,
            delegate: pendingDelegate
          });
        }
        if (meta?.token && typeof startTs === "number"){
          CalUtils.writeTalkMetadataToDocument(win.document, { startTimestamp: startTs });
          (async () => {
            try{
              await requestUtility(context, {
                type: "calendarUpdateLobby",
                token: meta.token,
                startTimestamp: startTs,
                delegateId: meta.delegateId || "",
                delegated: meta.delegated === true,
                windowId: API.windowId || null
              });
            }catch(_){}
            try{
              await requestUtility(context, {
                type: "trackRoom",
                token: meta.token,
                lobbyEnabled: meta.lobbyEnabled !== false,
                eventConversation: meta.eventConversation === true,
                startTimestamp: startTs
              });
            }catch(_){}
          })();
        }
      }catch(_){}
      if (!state.token) return;
      log("room cleanup persisted", { token: state.token });
      resetRoomCleanupState(state);
    };

    const drop = (reason) => triggerRoomCleanupDelete(state, reason);

    addListener(win, "dialogaccept", markPersisted, true);
    addListener(win, "dialogextra1", markPersisted, true);
    addListener(win, "dialogextra2", () => drop("dialogextra2"), true);
    addListener(win, "unload", () => drop("unload"), true);
  }

  function requestUtility(context, payload){
    try{
      return context.requestUtility(payload || {});
    }catch(e){
      err(e);
      return Promise.resolve(null);
    }
  }

  function ensureTrackedFromMetadata(context, doc, innerDoc){
    try{
      const meta = CalUtils.readTalkMetadataFromDocument(innerDoc || doc);
      if (!meta?.token) return;
      requestUtility(context, {
        type: "trackRoom",
        token: meta.token,
        lobbyEnabled: meta.lobbyEnabled !== false,
        eventConversation: meta.eventConversation === true,
        startTimestamp: meta.startTimestamp ?? extractStartTimestamp(innerDoc || doc)
      }).catch(() => {});
    }catch(_){}
  }

  function isEventDialogWindow(win){
    const href = win?.location?.href || win?.document?.location?.href;
    if (!href) return false;
    return EVENT_DIALOG_URLS.some((url) => href.startsWith(url));
  }

  function talkIconURL(size = 20){
    try{
      return API.getURL(`icons/talk-${size}.png`);
    }catch(_){
      return null;
    }
  }

  function findBar(doc){
    try{
      const candidates = doc.querySelector(".calendar-dialog-toolbar, .dialog-buttons, toolbar");
      if (candidates) return candidates;
    }catch(_){}
    return doc.body || doc.documentElement || null;
  }

  function buildButton(doc, context, label, tooltip){
    const btn = doc.createElement("button");
    btn.id = "nctalk-mini-btn";
    btn.type = "button";
    btn.title = tooltip || i18n("ui_toolbar_tooltip");
    Object.assign(btn.style, {
      display: "inline-flex",
      alignItems: "center",
      gap: "6px",
      padding: "3px 10px",
      marginInlineStart: "8px",
      marginInlineEnd: "0"
    });
    const img = doc.createElement("img");
    img.alt = "";
    img.width = 20;
    img.height = 20;
    const src = talkIconURL(20);
    if (src) img.src = src;
    const span = doc.createElement("span");
    span.textContent = label || i18n("ui_insert_button_label");
    btn.appendChild(img);
    btn.appendChild(span);
    ensureMenu(doc, context, btn);
    return btn;
  }

  function ensureMenu(doc, context, anchor){
    if (anchor.dataset.nctalkBridge === "1") return;
    anchor.dataset.nctalkBridge = "1";
    anchor.addEventListener("click", async (event) => {
      event.preventDefault();
      event.stopPropagation();
      try{
        await launchTalkPopup(context);
      }catch(e){
        err("Failed to launch Talk popup", e?.message || e);
      }
    });
  }

  async function launchTalkPopup(context){
    const response = await requestUtility(context, {
      type: "openDialog",
      windowId: API.windowId || null
    });
    if (!response?.ok){
      throw new Error(response?.error || "talk popup failed");
    }
  }

  function inject(doc, context, label, tooltip){
    if (!doc) return false;
    if (doc.getElementById("nctalk-mini-btn")) return true;
    const bar = findBar(doc);
    if (!bar){
      err("toolbar injection skipped: no bar element");
      return false;
    }
    const btn = buildButton(doc, context, label, tooltip);
    bar.appendChild(btn);
    return true;
  }

  function handle(win, context, label, tooltip){
    if (!isEventDialogWindow(win)) return;
    try{
      inject(win.document, context, label, tooltip);
    }catch(e){
      err(e);
    }
    const iframe = win.document.getElementById("calendar-item-panel-iframe");
    if (iframe){
      const run = () => {
        try{
          inject(iframe.contentDocument, context, label, tooltip);
        }catch(e){
          err(e);
        }
      };
      if (iframe.contentDocument?.readyState === "complete"){
        run();
      }else{
        iframe.addEventListener("load", run, { once:true });
      }
    }
  }

  async function requestCleanup(payload = {}){
    const token = payload?.token;
    if (!token){
      return { ok:false, error:"token_missing" };
    }
    const info = payload?.info || {};
    registerRoomCleanup(window, API, token, info);
    return { ok:true };
  }

  function updateBridgeState(){
    STATE.label = API.label || i18n("ui_insert_button_label");
    STATE.tooltip = API.tooltip || i18n("ui_toolbar_tooltip");
  }

  async function bootstrap(){
    try{
      await syncConfigState(API);
      updateBridgeState();
      handle(window, API, STATE.label, STATE.tooltip);
      const doc = window.document;
      const innerDoc = doc.getElementById && doc.getElementById("calendar-item-panel-iframe")?.contentDocument;
      ensureTrackedFromMetadata(API, doc, innerDoc || doc);
    }catch(e){
      err(e);
    }
  }

  if (document.readyState === "complete"){
    bootstrap();
  }else{
    window.addEventListener("load", bootstrap, { once:true });
  }
  window.addEventListener("nctalk-bridge-refresh", bootstrap);
  window.NCTalkLog = (text, details) => {
    try{
      log(text, details);
    }catch(_){}
  };
  window.NCTalkRegisterCleanup = requestCleanup;
})();
