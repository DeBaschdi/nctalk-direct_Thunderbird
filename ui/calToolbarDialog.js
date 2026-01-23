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

  /**
   * Log helper for the event dialog bridge.
   * @param {...any} args
   */
  function log(...args){
    try{
      console.log(LOG_CHANNEL, ...args);
    }catch(_){}
  }
  /**
   * Error log helper for the event dialog bridge.
   * @param {...any} args
   */
  function err(...args){
    try{
      console.error(LOG_CHANNEL, ...args);
    }catch(_){}
  }

  /**
   * Translate a key via the experiment bridge.
   * @param {string} key
   * @param {string[]|string} substitutions
   * @returns {string}
   */
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

  /**
   * Refresh cached config in background via the bridge.
   * @param {object} context
   * @returns {Promise<void>}
   */
  async function syncConfigState(context){
    try{
      await requestUtility(context, { type: "getConfig" });
    }catch(_){}
  }

  /**
   * Read Talk metadata from the event dialog window.
   * @param {Window} win
   * @returns {{ok:boolean,metadata:object}}
   */
  function getTalkMetadataFromWindow(win){
    const metadata = CalUtils.readTalkMetadataFromDocument(win.document);
    return { ok:true, metadata };
  }

  /**
   * Write Talk metadata into the event dialog window.
   * @param {Window} win
   * @param {object} payload
   * @returns {object}
   */
  function setTalkMetadataOnWindow(win, payload = {}){
    return CalUtils.setTalkMetadataOnWindow(win, payload);
  }

  /**
   * Collect title/location/description from the event dialog.
   * @param {Window} win
   * @returns {object}
   */
  function getEventSnapshotFromWindow(win){
    return CalUtils.getEventSnapshotFromWindow(win);
  }

  /**
   * Apply title/location/description into the event dialog.
   * @param {Window} win
   * @param {object} payload
   * @returns {object}
   */
  function applyEventFieldsOnWindow(win, payload = {}){
    return CalUtils.applyEventFieldsOnWindow(win, payload);
  }

  /**
   * Derive a start timestamp from the event dialog UI or item.
   * @param {Document} doc
   * @returns {number|null}
   */
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

  /**
   * Get or initialize room cleanup state stored on a window.
   * @param {Window} win
   * @returns {object|null}
   */
  function getRoomCleanupState(win){
    if (!win) return null;
    let state = win[ROOM_CLEANUP_KEY];
    if (!state){
      state = { token:null, context:null, info:null, cleanup:[], deleting:false };
      win[ROOM_CLEANUP_KEY] = state;
    }
    return state;
  }

  /**
   * Run and clear all cleanup callbacks for a room state.
   * @param {object} state
   */
  function cleanupRoomCleanupState(state){
    if (!state || !Array.isArray(state.cleanup)) return;
    while (state.cleanup.length){
      const fn = state.cleanup.pop();
      try{ fn(); }catch(_){}
    }
  }

  /**
   * Reset the room cleanup state to its defaults.
   * @param {object} state
   */
  function resetRoomCleanupState(state){
    if (!state) return;
    cleanupRoomCleanupState(state);
    state.token = null;
    state.context = null;
    state.info = null;
    state.deleting = false;
  }

  /**
   * Request Talk room deletion when the dialog is canceled or superseded.
   * @param {object} state
   * @param {string} reason
   */
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

  /**
   * Register cleanup hooks to delete the Talk room if the event is not saved.
   * @param {Window} win
   * @param {object} context
   * @param {string} token
   * @param {object} info
   */
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

    /**
     * Track a listener and ensure it is removed on cleanup.
     * @param {EventTarget} target
     * @param {string} type
     * @param {Function} handler
     * @param {any} options
     */
    const addListener = (target, type, handler, options) => {
      if (!target || typeof target.addEventListener !== "function") return;
      target.addEventListener(type, handler, options);
      state.cleanup.push(() => {
        if (typeof target?.removeEventListener === "function"){
          target.removeEventListener(type, handler, options);
        }
      });
    };

    /**
     * Persist updated metadata when the dialog is saved.
     */
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

    /**
     * Trigger cleanup with a reason string.
     * @param {string} reason
     */
    const drop = (reason) => triggerRoomCleanupDelete(state, reason);

    addListener(win, "dialogaccept", markPersisted, true);
    addListener(win, "dialogextra1", markPersisted, true);
    addListener(win, "dialogextra2", () => drop("dialogextra2"), true);
    addListener(win, "unload", () => drop("unload"), true);
  }

  /**
   * Send a utility request through the experiment bridge.
   * @param {object} context
   * @param {object} payload
   * @returns {Promise<any>}
   */
  function requestUtility(context, payload){
    try{
      return context.requestUtility(payload || {});
    }catch(e){
      err(e);
      return Promise.resolve(null);
    }
  }

  /**
   * Ensure the current event is tracked from stored metadata.
   * @param {object} context
   * @param {Document} doc
   * @param {Document} innerDoc
   */
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

  /**
   * Check if a window hosts the calendar event dialog.
   * @param {Window} win
   * @returns {boolean}
   */
  function isEventDialogWindow(win){
    const href = win?.location?.href || win?.document?.location?.href;
    if (!href) return false;
    return EVENT_DIALOG_URLS.some((url) => href.startsWith(url));
  }

  /**
   * Resolve the Talk icon URL for the given size.
   * @param {number} size
   * @returns {string|null}
   */
  function talkIconURL(size = 20){
    try{
      return API.getURL(`icons/talk-${size}.png`);
    }catch(_){
      return null;
    }
  }

  /**
   * Find a suitable toolbar container for the injected button.
   * @param {Document} doc
   * @returns {Element|null}
   */
  function findBar(doc){
    try{
      const candidates = doc.querySelector(".calendar-dialog-toolbar, .dialog-buttons, toolbar");
      if (candidates) return candidates;
    }catch(_){}
    return doc.body || doc.documentElement || null;
  }

  /**
   * Build the Talk toolbar button element.
   * @param {Document} doc
   * @param {object} context
   * @param {string} label
   * @param {string} tooltip
   * @returns {HTMLButtonElement}
   */
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

  /**
   * Bind click handler to open the Talk popup.
   * @param {Document} doc
   * @param {object} context
   * @param {HTMLElement} anchor
   */
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

  /**
   * Request the background to open the Talk dialog.
   * @param {object} context
   * @returns {Promise<void>}
   */
  async function launchTalkPopup(context){
    const response = await requestUtility(context, {
      type: "openDialog",
      windowId: API.windowId || null
    });
    if (!response?.ok){
      throw new Error(response?.error || "talk popup failed");
    }
  }

  /**
   * Inject the Talk toolbar button into the document.
   * @param {Document} doc
   * @param {object} context
   * @param {string} label
   * @param {string} tooltip
   * @returns {boolean}
   */
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

  /**
   * Inject the Talk button into the window and its iframe (if present).
   * @param {Window} win
   * @param {object} context
   * @param {string} label
   * @param {string} tooltip
   */
  function handle(win, context, label, tooltip){
    if (!isEventDialogWindow(win)) return;
    try{
      inject(win.document, context, label, tooltip);
    }catch(e){
      err(e);
    }
    const iframe = win.document.getElementById("calendar-item-panel-iframe");
    if (iframe){
      /**
       * Inject after the inner iframe has loaded.
       */
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

  /**
   * Register a cleanup hook for the current dialog.
   * @param {object} payload
   * @returns {Promise<{ok:boolean,error?:string}>}
   */
  async function requestCleanup(payload = {}){
    const token = payload?.token;
    if (!token){
      return { ok:false, error:"token_missing" };
    }
    const info = payload?.info || {};
    registerRoomCleanup(window, API, token, info);
    return { ok:true };
  }

  /**
   * Update local label and tooltip from the bridge state.
   */
  function updateBridgeState(){
    STATE.label = API.label || i18n("ui_insert_button_label");
    STATE.tooltip = API.tooltip || i18n("ui_toolbar_tooltip");
  }

  /**
   * Bootstrap the dialog bridge and inject the button.
   */
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
