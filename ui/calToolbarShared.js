/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
/**
 * Shared helpers for the calendar event dialog integration.
 */
(() => {
  "use strict";

  const globalScope = typeof window !== "undefined"
    ? window
    : (typeof globalThis !== "undefined" ? globalThis : this);

  if (globalScope.NCTalkCalUtils){
    return;
  }

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
    return null;
  }

  function parseNumberProp(value){
    if (typeof value === "number" && Number.isFinite(value)){
      return value;
    }
    if (typeof value === "string"){
      const trimmed = value.trim();
      if (!trimmed) return null;
      const parsed = parseInt(trimmed, 10);
      return Number.isNaN(parsed) ? null : parsed;
    }
    return null;
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
      const win = doc?.defaultView || globalScope;
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

  function resolveDocument(target){
    if (!target) return null;
    if (target.document) return target.document;
    if (target.nodeType === 9) return target;
    if (target.defaultView && target.defaultView.document) return target.defaultView.document;
    return null;
  }

  function collectEventDocs(target){
    const docs = [];
    const doc = resolveDocument(target);
    if (!doc) return docs;
    const pushDoc = (entry) => {
      if (entry && docs.indexOf(entry) === -1){
        docs.push(entry);
      }
    };
    try{
      pushDoc(doc);
    }catch(_){}
    try{
      const iframe = doc.getElementById && doc.getElementById("calendar-item-panel-iframe");
      if (iframe?.contentDocument){
        pushDoc(iframe.contentDocument);
      }
    }catch(_){}
    return docs;
  }

  function findField(docs, selectors){
    for (const doc of docs){
      if (!doc || typeof doc.querySelector !== "function") continue;
      for (const sel of selectors){
        try{
          const element = doc.querySelector(sel);
          if (element) return element;
        }catch(_){}
      }
    }
    return null;
  }

  /**
   * Locate the event description field across dialog variants and editors.
   * @param {Document[]} docs
   * @returns {Element|null}
   */
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

  function dispatchInputEvent(field){
    if (!field) return;
    try{
      const doc = field.ownerDocument || field.document;
      const win = doc?.defaultView;
      if (win){
        const evt = new win.Event("input", { bubbles:true });
        field.dispatchEvent(evt);
      }
    }catch(_){}
  }

  function setFieldValue(field, value, opts = {}){
    if (!field) return;
    const doc = field.ownerDocument || field.document || field.contentDocument || null;
    const execPreferred = opts.preferExec === true;

    const tryExecCommand = () => {
      if (!doc || typeof doc.execCommand !== "function"){
        return false;
      }
      try{
        field.focus?.();
        doc.execCommand("selectAll", false, null);
        doc.execCommand("insertText", false, value);
        return true;
      }catch(_){
        return false;
      }
    };

    if (execPreferred && tryExecCommand()){
      dispatchInputEvent(field);
      return;
    }

    if ("value" in field){
      try{ field.focus?.(); }catch(_){}
      field.value = value;
      dispatchInputEvent(field);
      return;
    }

    if ((field.isContentEditable || field.tagName?.toLowerCase() === "body") && tryExecCommand()){
      dispatchInputEvent(field);
      return;
    }

    if (field.textContent !== undefined){
      field.textContent = value;
      dispatchInputEvent(field);
    }
  }

  function getFieldValue(field){
    if (!field) return "";
    if ("value" in field){
      return field.value || "";
    }
    if (field.textContent != null){
      return field.textContent;
    }
    return "";
  }

  /**
   * Read Talk metadata properties from the calendar item in a document.
   * @param {Document} doc
   * @returns {object}
   */
  function readTalkMetadataFromDocument(doc){
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

  /**
   * Write Talk metadata properties to the calendar item in a document.
   * @param {Document} doc
   * @param {object} meta
   * @returns {{ok:boolean, error?:string}}
   */
  function writeTalkMetadataToDocument(doc, meta = {}){
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

  function setTalkMetadataOnWindow(target, payload = {}){
    const doc = resolveDocument(target);
    if (!doc){
      return { ok:false, error:"no_document" };
    }
    return writeTalkMetadataToDocument(doc, payload);
  }

  /**
   * Collect event fields and Talk metadata from the event dialog window.
   * @param {Window|Document} target
   * @param {object} options
   * @returns {object}
   */
  function getEventSnapshotFromWindow(target, options = {}){
    const doc = resolveDocument(target);
    if (!doc){
      return { ok:false, error:"no_document" };
    }
    const metadata = options.metadata
      || (typeof options.readMetadata === "function" ? options.readMetadata(doc) : readTalkMetadataFromDocument(doc));
    const docs = collectEventDocs(doc);
    const titleField = findField(docs, [
      "#item-title",
      'input[id^="event-grid-title"]',
      'input[type="text"]'
    ]);
    const locationField = findField(docs, [
      'input[aria-label="Ort"]',
      'input[placeholder="Ort"]',
      "input#item-location",
      'input[name="location"]',
      'textbox[id*="location"]'
    ]);
    const descField = findDescriptionFieldInDocs(docs);
    const event = {
      title: getFieldValue(titleField) || metadata?.title || "",
      location: getFieldValue(locationField) || "",
      description: getFieldValue(descField) || "",
      startTimestamp: metadata?.startTimestamp || null,
      endTimestamp: metadata?.endTimestamp || null
    };
    return { ok:true, event, metadata };
  }

  /**
   * Apply title, location, and description with editor-aware fallbacks.
   * @param {Window|Document} target
   * @param {object} payload
   * @param {object} options
   * @returns {object}
   */
  function applyEventFieldsOnWindow(target, payload = {}, options = {}){
    const doc = resolveDocument(target);
    if (!doc){
      return { ok:false, error:"no_document" };
    }
    const hasFieldOptions = options
      && !options.titleOptions
      && !options.locationOptions
      && !options.descriptionOptions
      && !options.preferExecForDescription;
    const baseOptions = hasFieldOptions ? options : null;
    const titleOptions = options.titleOptions || baseOptions || {};
    const locationOptions = options.locationOptions || baseOptions || {};
    let descriptionOptions = options.descriptionOptions || baseOptions || {};
    if (options.preferExecForDescription === true && !options.descriptionOptions){
      descriptionOptions = Object.assign({}, descriptionOptions, { preferExec:true });
    }
    const docs = collectEventDocs(doc);
    const titleField = findField(docs, [
      "#item-title",
      'input[id^="event-grid-title"]',
      'input[type="text"]'
    ]);
    const locationField = findField(docs, [
      'input[aria-label="Ort"]',
      'input[placeholder="Ort"]',
      "input#item-location",
      'input[name="location"]',
      'textbox[id*="location"]'
    ]);
    const descField = findDescriptionFieldInDocs(docs);
    if (typeof payload.title === "string" && titleField){
      setFieldValue(titleField, payload.title, titleOptions);
    }
    if (typeof payload.location === "string" && locationField){
      setFieldValue(locationField, payload.location, locationOptions);
    }
    if (typeof payload.description === "string" && descField){
      setFieldValue(descField, payload.description, descriptionOptions);
    }
    return {
      ok:true,
      applied: {
        title: !!titleField,
        location: !!locationField,
        description: !!descField
      }
    };
  }

  globalScope.NCTalkCalUtils = {
    getCalendarItemFromDocument: getCalendarItem,
    safeString,
    parseBooleanProp,
    parseNumberProp,
    boolToProp,
    collectEventDocs,
    findField,
    findDescriptionFieldInDocs,
    getFieldValue,
    setFieldValue,
    readTalkMetadataFromDocument,
    writeTalkMetadataToDocument,
    setTalkMetadataOnWindow,
    getEventSnapshotFromWindow,
    applyEventFieldsOnWindow
  };
})();
