(() => {
  "use strict";

  const API = window.NCTalkBridge;
  const LOG_CHANNEL = "[NCUI][Talk]";
  if (!API){
    console.error(`${LOG_CHANNEL} Bridge nicht verfügbar.`);
    return;
  }

  const EVENT_DIALOG_URLS = [
    "chrome://calendar/content/calendar-event-dialog.xhtml",
    "chrome://calendar/content/calendar-event-dialog.xul"
  ];

  const TALK_PROP_TOKEN = "X-NCTALK-TOKEN";
  const TALK_PROP_URL = "X-NCTALK-URL";
  const TALK_PROP_LOBBY = "X-NCTALK-LOBBY";
  const TALK_PROP_START = "X-NCTALK-START";
  const TALK_PROP_EVENT = "X-NCTALK-EVENT";
  const TALK_PROP_OBJECT_ID = "X-NCTALK-OBJECTID";
  const TALK_PROP_DELEGATE = "X-NCTALK-DELEGATE";
  const TALK_PROP_DELEGATE_NAME = "X-NCTALK-DELEGATE-NAME";
  const TALK_PROP_DELEGATED = "X-NCTALK-DELEGATED";
  const TALK_PROP_DELEGATE_READY = "X-NCTALK-DELEGATE-READY";

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

  function safeString(value){
    return typeof value === "string" && value.length ? value : null;
  }

  function parseBooleanProp(value){
    if (typeof value === "boolean") return value;
    if (typeof value === "string"){
      const norm = value.trim().toLowerCase();
      if (norm === "true" || norm === "1" || norm === "yes") return true;
      if (norm === "false" || norm === "0" || norm === "no") return false;
    }
    return value ? true : false;
  }

  function parseNumberProp(value){
    const str = safeString(value);
    if (!str) return null;
    const parsed = parseInt(str, 10);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function boolToProp(value){
    if (typeof value === "string"){
      const norm = value.trim().toLowerCase();
      if (norm === "true" || norm === "1" || norm === "yes") return "TRUE";
      if (norm === "false" || norm === "0" || norm === "no") return "FALSE";
    }
    return value ? "TRUE" : "FALSE";
  }

  function getCalendarItem(doc){
    try{
      const win = doc?.defaultView || window;
      if (win?.calendarItem) return win.calendarItem;
      if (win?.gEvent?.event) return win.gEvent.event;
      if (Array.isArray(win?.arguments) && win.arguments[0]){
        const arg = win.arguments[0];
        if (arg.calendarItem) return arg.calendarItem;
        if (arg.calendarEvent) return arg.calendarEvent;
      }
    }catch(_){}
    return null;
  }

  function collectEventDocs(doc){
    const docs = [];
    try{ if (doc) docs.push(doc); }catch(_){}
    try{
      const iframe = doc?.getElementById?.("calendar-item-panel-iframe");
      if (iframe?.contentDocument && iframe.contentDocument !== doc){
        docs.push(iframe.contentDocument);
      }
    }catch(_){}
    return docs;
  }

  function findField(docs, selectors){
    for (const doc of docs){
      if (!doc || typeof doc.querySelector !== "function") continue;
      for (const sel of selectors){
        try{
          const el = doc.querySelector(sel);
          if (el) return el;
        }catch(_){}
      }
    }
    return null;
  }

  function findDescriptionFieldInDocs(docs){
    for (const doc of docs){
      try{
        const host = doc.querySelector && doc.querySelector("editor#item-description");
        let target = null;
        if (host){
          target = host.inputField || host.contentDocument?.body || host;
        }
        if (!target){
          const fallbacks = [
            "textarea#item-description",
            "textarea",
            "[contenteditable='true']",
            "div[role='textbox']"
          ];
          for (const sel of fallbacks){
            const el = doc.querySelector && doc.querySelector(sel);
            if (el){
              target = el;
              break;
            }
          }
        }
        if (target) return target;
      }catch(_){}
    }
    return null;
  }

  function getFieldValue(field){
    if (!field) return "";
    if ("value" in field) return field.value || "";
    if ("textContent" in field) return field.textContent || "";
    return "";
  }

  function setFieldValue(field, value){
    if (!field) return;
    if ("value" in field){
      field.focus?.();
      field.value = value;
      const doc = field.ownerDocument || document;
      field.dispatchEvent?.(new doc.defaultView.Event("input", { bubbles:true }));
      return;
    }
    field.textContent = value;
  }

  function readTalkMetadata(doc){
    try{
      const item = getCalendarItem(doc);
      if (!item || typeof item.getProperty !== "function"){
        return {};
      }
      const get = (name) => {
        try{
          return safeString(item.getProperty(name));
        }catch(_){
          return null;
        }
      };
      return {
        token: get(TALK_PROP_TOKEN),
        url: get(TALK_PROP_URL),
        lobbyEnabled: (() => {
          const raw = get(TALK_PROP_LOBBY);
          return raw == null ? null : parseBooleanProp(raw);
        })(),
        startTimestamp: parseNumberProp(get(TALK_PROP_START)),
        eventConversation: (() => {
          const raw = get(TALK_PROP_EVENT);
          if (!raw) return null;
          return raw.trim().toLowerCase() === "event";
        })(),
        objectId: get(TALK_PROP_OBJECT_ID),
        delegateId: get(TALK_PROP_DELEGATE),
        delegateName: get(TALK_PROP_DELEGATE_NAME),
        delegated: (() => {
          const raw = get(TALK_PROP_DELEGATED);
          if (raw == null) return false;
          return parseBooleanProp(raw);
        })(),
        delegateReady: (() => {
          const raw = get(TALK_PROP_DELEGATE_READY);
          if (raw == null) return null;
          return parseBooleanProp(raw);
        })()
      };
    }catch(_){
      return {};
    }
  }

  function writeTalkMetadata(doc, meta = {}){
    const item = getCalendarItem(doc);
    if (!item || typeof item.setProperty !== "function"){
      return { ok:false, error:"no_calendar_item" };
    }
    const setProp = (name, value) => {
      try{
        if (value == null || value === ""){
          if (typeof item.deleteProperty === "function"){
            item.deleteProperty(name);
          }else{
            item.setProperty(name, "");
          }
        }else{
          item.setProperty(name, String(value));
        }
      }catch(_){}
    };
    if ("token" in meta) setProp(TALK_PROP_TOKEN, meta.token);
    if ("url" in meta) setProp(TALK_PROP_URL, meta.url);
    if ("lobbyEnabled" in meta) setProp(TALK_PROP_LOBBY, boolToProp(meta.lobbyEnabled));
    if ("startTimestamp" in meta && meta.startTimestamp != null){
      const ts = Number(meta.startTimestamp);
      if (Number.isFinite(ts)){
        setProp(TALK_PROP_START, String(Math.floor(ts)));
      }
    }
    if ("eventConversation" in meta){
      setProp(TALK_PROP_EVENT, meta.eventConversation ? "event" : "standard");
    }
    if ("objectId" in meta) setProp(TALK_PROP_OBJECT_ID, meta.objectId);
    if ("delegateId" in meta) setProp(TALK_PROP_DELEGATE, meta.delegateId);
    if ("delegateName" in meta) setProp(TALK_PROP_DELEGATE_NAME, meta.delegateName);
    if ("delegated" in meta) setProp(TALK_PROP_DELEGATED, boolToProp(!!meta.delegated));
    if ("delegateReady" in meta){
      const ready = meta.delegateReady;
      if (ready == null){
        setProp(TALK_PROP_DELEGATE_READY, "");
      }else{
        setProp(TALK_PROP_DELEGATE_READY, boolToProp(!!ready));
      }
    }
    return { ok:true };
  }

  function getTalkMetadataFromWindow(win){
    const metadata = readTalkMetadata(win.document);
    return { ok:true, metadata };
  }

  function setTalkMetadataOnWindow(win, payload = {}){
    return writeTalkMetadata(win.document, payload);
  }

  function getEventSnapshotFromWindow(win){
    const metadata = readTalkMetadata(win.document);
    const docs = collectEventDocs(win.document);
    const titleField = findField(docs, [
      "#item-title",
      'input[id^="event-grid-title"]',
      'input[type="text"]'
    ]);
    const locationField = findField(docs, [
      'input[aria-label="Ort"]',
      'input[placeholder="Ort"]',
      'input#item-location',
      'input[name="location"]',
      'textbox[id*="location"]'
    ]);
    const descField = findDescriptionFieldInDocs(docs);
    const event = {
      title: getFieldValue(titleField) || metadata.title || "",
      location: getFieldValue(locationField) || "",
      description: getFieldValue(descField) || "",
      startTimestamp: metadata.startTimestamp || null,
      endTimestamp: metadata.endTimestamp || null
    };
    return { ok:true, event, metadata };
  }

  function applyEventFieldsOnWindow(win, payload = {}){
    const docs = collectEventDocs(win.document);
    const titleField = findField(docs, [
      "#item-title",
      'input[id^="event-grid-title"]',
      'input[type="text"]'
    ]);
    const locationField = findField(docs, [
      'input[aria-label="Ort"]',
      'input[placeholder="Ort"]',
      'input#item-location',
      'input[name="location"]',
      'textbox[id*="location"]'
    ]);
    const descField = findDescriptionFieldInDocs(docs);
    if (typeof payload.title === "string" && titleField){
      setFieldValue(titleField, payload.title);
    }
    if (typeof payload.location === "string" && locationField){
      setFieldValue(locationField, payload.location);
    }
    if (typeof payload.description === "string" && descField){
      setFieldValue(descField, payload.description);
    }
    return { ok:true };
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
        const meta = readTalkMetadata(win.document);
        const startTs = extractStartTimestamp(win.document);
        log("room cleanup metadata snapshot", {
          token: meta?.token || "",
          delegateId: meta?.delegateId || "",
          delegateName: meta?.delegateName || "",
          delegated: meta?.delegated === true,
          delegateReady: meta?.delegateReady ?? null,
          startTimestamp: startTs ?? null
        });
        if (meta?.token && typeof startTs === "number"){
          writeTalkMetadata(win.document, { startTimestamp: startTs });
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
            const pendingDelegate = (meta.delegateId || "").trim();
            if (pendingDelegate && meta.delegated !== true){
              writeTalkMetadata(win.document, { delegateReady: true });
              log("delegate armed for calendar flow", {
                token: meta.token,
                delegate: pendingDelegate
              });
            }
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
      const meta = readTalkMetadata(innerDoc || doc);
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
