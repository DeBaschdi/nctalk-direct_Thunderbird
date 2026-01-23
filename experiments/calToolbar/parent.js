/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
/**
 * Minimal experiment bridge for the calendar dialog.
 * Registers window listeners and exposes a WebExtension bridge.
 */
'use strict';

var { ExtensionCommon } = ChromeUtils.importESModule("resource://gre/modules/ExtensionCommon.sys.mjs");
var { ExtensionSupport } = ChromeUtils.importESModule("resource:///modules/ExtensionSupport.sys.mjs");

const BRIDGE_SCRIPT_PATH = "ui/calToolbarDialog.js";
const CAL_SHARED_PATH = "ui/calToolbarShared.js";
const EVENT_DIALOG_URLS = [
  "chrome://calendar/content/calendar-event-dialog.xhtml",
  "chrome://calendar/content/calendar-event-dialog.xul"
];
const WINDOW_LISTENER_ID = "ext-nextcloud-enterprise-thunderbird";

class CalToolbarExperiment {
  constructor(context) {
    this.context = context;
    this.extension = context.extension;
    this.calSharedUrl = "";
    this.calUtils = null;
    this.windowRegistry = new Map();
    this.windowIdCounter = 1;
    this.windowCleanups = new Map();
    this.stopListening = null;
    this.createHandlers = new Set();
    this.utilityHandlers = new Set();
  }

  close() {
    if (this.stopListening) {
      this.stopListening();
      this.stopListening = null;
    }
    this.cleanupWindows();
    this.createHandlers.clear();
    this.utilityHandlers.clear();
  }

  getAPI() {
    const createEvent = new ExtensionCommon.EventManager({
      context: this.context,
      name: "calToolbar.onCreateRequest",
      register: (fire) => {
        this.createHandlers.add(fire);
        return () => this.createHandlers.delete(fire);
      }
    });
    const utilityEvent = new ExtensionCommon.EventManager({
      context: this.context,
      name: "calToolbar.onUtilityRequest",
      register: (fire) => {
        this.utilityHandlers.add(fire);
        return () => this.utilityHandlers.delete(fire);
      }
    });

    return {
      calToolbar: {
        onCreateRequest: createEvent.api(),
        onUtilityRequest: utilityEvent.api(),
        init: (opts = {}) => this.init(opts),
        invokeWindow: (options = {}) => this.invokeWindow(options)
      }
    };
  }

  i18n(key, substitutions = []) {
    const localeData = this.extension?.localeData;
    if (localeData && typeof localeData.localizeMessage === "function") {
      const message = localeData.localizeMessage(key, substitutions);
      if (message) {
        return message;
      }
    }
    if (Array.isArray(substitutions) && substitutions.length) {
      return String(substitutions[0] ?? "");
    }
    return key;
  }

  ensureCalUtils() {
    if (this.calUtils) return this.calUtils;
    if (typeof Services === "undefined" || !Services?.scriptloader) return null;
    const url = this.calSharedUrl || (this.extension?.getURL ? this.extension.getURL(CAL_SHARED_PATH) : "");
    if (!url) return null;
    this.calSharedUrl = url;
    try {
      const globalScope = typeof globalThis !== "undefined" ? globalThis : this;
      Services.scriptloader.loadSubScript(url, globalScope, "UTF-8");
    } catch (e) {
      console.error("[NCExp] shared utils load failed", e);
      return null;
    }
    this.calUtils = (typeof globalThis !== "undefined" ? globalThis.NCTalkCalUtils : this.NCTalkCalUtils) || null;
    return this.calUtils;
  }

  ensureCalUtilsInWindow(win) {
    if (!win || !this.calSharedUrl) return null;
    if (win.NCTalkCalUtils) return win.NCTalkCalUtils;
    if (typeof Services === "undefined" || !Services?.scriptloader) return null;
    try {
      Services.scriptloader.loadSubScript(this.calSharedUrl, win, "UTF-8");
    } catch (e) {
      console.error("[NCExp] shared utils window load failed", e);
      return null;
    }
    return win.NCTalkCalUtils || null;
  }

  async dispatchHandlers(set, payload) {
    if (!set || set.size === 0) return null;
    for (const fire of set) {
      try {
        const result = await fire.async(payload);
        if (result !== undefined) {
          return result;
        }
      } catch (e) {
        console.error("[NCExp] handler error", e);
      }
    }
    return null;
  }

  isEventDialogWindow(win) {
    if (!win) return false;
    try {
      const doc = win.document;
      const windowType = doc?.documentElement?.getAttribute("windowtype") || "";
      if (windowType === "Calendar:EventDialog" || windowType === "Calendar:EventSummaryDialog") {
        return true;
      }
      const href = doc?.location?.href || win.location?.href || "";
      if (!href) return false;
      return EVENT_DIALOG_URLS.some((url) => href.startsWith(url));
    } catch (_) {
      return false;
    }
  }

  ensureWindowId(win) {
    if (!win || typeof win !== "object") {
      return null;
    }
    if (win.NCTalkWindowId && this.windowRegistry.has(win.NCTalkWindowId)) {
      return win.NCTalkWindowId;
    }
    const id = this.windowIdCounter++;
    win.NCTalkWindowId = id;
    this.windowRegistry.set(id, win);
    return id;
  }

  releaseWindowId(win) {
    if (!win || !win.NCTalkWindowId) {
      return;
    }
    const id = win.NCTalkWindowId;
    this.windowRegistry.delete(id);
    delete win.NCTalkWindowId;
  }

  activateEventDialogWindow(win) {
    if (!win || this.windowCleanups.has(win)) {
      return;
    }
    /**
     * Release stored window data when the dialog unloads.
     */
    const cleanup = () => {
      win.removeEventListener("unload", cleanup);
      this.windowCleanups.delete(win);
      this.releaseWindowId(win);
    };
    win.addEventListener("unload", cleanup);
    this.windowCleanups.set(win, cleanup);
  }

  cleanupWindows() {
    for (const cleanup of Array.from(this.windowCleanups.values())) {
      try {
        cleanup();
      } catch (e) {
        console.error("[NCExp] window cleanup failed", e);
      }
    }
    for (const win of Array.from(this.windowRegistry.values())) {
      this.releaseWindowId(win);
    }
  }

  getRegisteredWindow(windowId) {
    if (typeof windowId !== "number") {
      return null;
    }
    const win = this.windowRegistry.get(windowId);
    if (!win || win.closed || !win.document) {
      this.windowRegistry.delete(windowId);
      return null;
    }
    return win;
  }

  createBridgeAPI(init = {}) {
    const state = {
      label: init.label || this.i18n("ui_insert_button_label"),
      tooltip: init.tooltip || this.i18n("ui_toolbar_tooltip"),
      windowId: init.windowId || null
    };
    return {
      get label() { return state.label; },
      get tooltip() { return state.tooltip; },
      get windowId() { return state.windowId; },
      updateInit: (opts = {}) => {
        if (typeof opts.label === "string") state.label = opts.label;
        if (typeof opts.tooltip === "string") state.tooltip = opts.tooltip;
        if (typeof opts.windowId === "number") state.windowId = opts.windowId;
      },
      i18n: (key, subs) => this.i18n(key, subs),
      getURL: (path) => this.extension?.getURL?.(path || "") || "",
      requestCreate: (payload) => this.dispatchHandlers(this.createHandlers, payload || {}),
      requestUtility: (payload = {}) => this.dispatchHandlers(this.utilityHandlers, payload),
      openDialog: async () => {
        const windowId = state.windowId || null;
        return this.dispatchHandlers(this.utilityHandlers, {
          type: "openDialog",
          windowId
        });
      }
    };
  }

  dispatchBridgeRefresh(win) {
    const eventCtor = win?.CustomEvent || win?.Event;
    if (typeof win?.dispatchEvent === "function" && typeof eventCtor === "function") {
      win.dispatchEvent(new eventCtor("nctalk-bridge-refresh"));
    }
  }

  installBridge(win, init) {
    if (!win || !win.document) return false;
    const windowId = this.ensureWindowId(win);
    const initWithId = Object.assign({}, init, { windowId });
    const api = this.createBridgeAPI(initWithId);
    this.ensureCalUtils();
    this.ensureCalUtilsInWindow(win);
    if (win.NCTalkBridge && typeof win.NCTalkBridge.updateInit === "function") {
      win.NCTalkBridge.updateInit(initWithId);
      this.dispatchBridgeRefresh(win);
    } else {
      Object.defineProperty(win, "NCTalkBridge", {
        value: api,
        configurable: true
      });
    }
    if (win.NCTalkBridgeLoaded) {
      this.dispatchBridgeRefresh(win);
      return true;
    }
    try {
      Services.scriptloader.loadSubScript(
        this.extension.getURL(BRIDGE_SCRIPT_PATH),
        win,
        "UTF-8"
      );
      win.NCTalkBridgeLoaded = true;
      this.dispatchBridgeRefresh(win);
      return true;
    } catch (e) {
      console.error("[NCExp] bridge script load failed", e);
      return false;
    }
  }

  registerWindowListener(init) {
    /**
     * Inject the bridge into a newly opened event dialog window.
     * @param {Window} win
     */
    const onLoadWindow = (win) => {
      if (!win) return;
      /**
       * Inject the dialog bridge after verifying window type.
       */
      const inject = () => {
        if (!this.isEventDialogWindow(win)) {
          return;
        }
        if (this.installBridge(win, init)) {
          this.activateEventDialogWindow(win);
          console.log("[NCExp] bridge injected window");
        }
      };
      if (win.document?.readyState === "complete") {
        inject();
      } else {
        /**
         * Wait for the window to finish loading before injecting.
         */
        const onLoad = () => {
          win.removeEventListener("load", onLoad);
          inject();
        };
        win.addEventListener("load", onLoad, { once: true });
      }
    };

    ExtensionSupport.registerWindowListener(WINDOW_LISTENER_ID, {
      chromeURLs: EVENT_DIALOG_URLS,
      onLoadWindow
    });

    let active = true;
    return () => {
      if (!active) return;
      active = false;
      try {
        ExtensionSupport.unregisterWindowListener(WINDOW_LISTENER_ID);
      } catch (_) {}
      this.cleanupWindows();
    };
  }

  async init(opts = {}) {
    const label = opts.label || this.i18n("ui_insert_button_label");
    const tooltip = opts.tooltip || this.i18n("ui_toolbar_tooltip");
    try {
      if (this.stopListening) {
        this.stopListening();
        this.stopListening = null;
      }
      this.stopListening = this.registerWindowListener({ label, tooltip });
      return true;
    } catch (e) {
      console.error("[NCExp] init error", e);
      if (this.stopListening) {
        this.stopListening();
        this.stopListening = null;
      }
      return false;
    }
  }

  async invokeWindow(options = {}) {
    const windowId = options.windowId;
    const action = options.action;
    const payload = options.payload || {};
    if (typeof windowId !== "number" || !action) {
      throw new Error("windowId and action are required");
    }
    const win = this.getRegisteredWindow(windowId);
    if (!win) {
      throw new Error("target window not available");
    }
    return await this.executeWindowAction(win, action, payload);
  }

  async executeWindowAction(win, action, payload) {
    this.ensureCalUtils();
    switch (action) {
      case "ping":
        return { ok: true };
      case "getTalkMetadata":
        return this.getTalkMetadataFromWindow(win);
      case "setTalkMetadata":
        return this.setTalkMetadataOnWindow(win, payload || {});
      case "getEventSnapshot":
        return this.getEventSnapshotFromWindow(win);
      case "applyEventFields":
        return this.applyEventFieldsOnWindow(win, payload || {});
      case "registerCleanup":
        if (typeof win.NCTalkRegisterCleanup === "function") {
          return await win.NCTalkRegisterCleanup(payload || {});
        }
        return { ok: false, error: "register_cleanup_unavailable" };
      default:
        return { ok: false, error: "unknown_action" };
    }
  }

  getTalkMetadataFromWindow(win) {
    const item = getCalendarItemFromWindow(win);
    const utils = this.ensureCalUtils();
    if (!item || !utils) {
      return { ok: false, error: "no_calendar_item" };
    }
    const metadata = extractTalkMetadataFromItem(item, utils) || {};
    return { ok: true, metadata };
  }

  setTalkMetadataOnWindow(win, meta = {}) {
    return this.calUtils?.setTalkMetadataOnWindow
      ? this.calUtils.setTalkMetadataOnWindow(win, meta)
      : { ok: false, error: "cal_utils_unavailable" };
  }

  getEventSnapshotFromWindow(win) {
    const metadata = this.getTalkMetadataFromWindow(win);
    const meta = metadata.ok ? (metadata.metadata || {}) : {};
    return this.calUtils?.getEventSnapshotFromWindow
      ? this.calUtils.getEventSnapshotFromWindow(win, { metadata: meta })
      : { ok: false, error: "cal_utils_unavailable" };
  }

  applyEventFieldsOnWindow(win, payload = {}) {
    const result = this.calUtils?.applyEventFieldsOnWindow
      ? this.calUtils.applyEventFieldsOnWindow(win, payload, { preferExecForDescription: true })
      : null;
    return result || { ok: false, error: "apply_failed" };
  }
}

/**
 * Read a calendar item property safely.
 */
function safeItemProperty(item, prop, calUtils) {
  if (!item || typeof item.getProperty !== "function") return null;
  try {
    return calUtils?.safeString ? calUtils.safeString(item.getProperty(prop)) : item.getProperty(prop);
  } catch (_) {
    return null;
  }
}

/**
 * Resolve the calendar item from an event dialog window context.
 */
function getCalendarItemFromWindow(win) {
  try {
    if (!win) return null;
    if (win.calendarItem) return win.calendarItem;
    if (win.gEvent && win.gEvent.event) return win.gEvent.event;
    if (Array.isArray(win.arguments) && win.arguments[0]) {
      const arg = win.arguments[0];
      if (arg.calendarItem) return arg.calendarItem;
      if (arg.calendarEvent) return arg.calendarEvent;
    }
  } catch (_) {}
  return null;
}

/**
 * Extract Talk-specific metadata from a calendar item.
 */
function extractTalkMetadataFromItem(item, calUtils) {
  if (!item) return {};
  /**
   * Convert raw property values to finite numbers.
   * @param {any} raw
   * @returns {number|null}
   */
  const toNumber = (raw) => {
    if (raw == null) return null;
    const num = Number(raw);
    return Number.isFinite(num) ? num : null;
  };
  const title = calUtils?.safeString?.(item?.title) || (() => {
    try {
      if (typeof item.getProperty === "function") {
        return calUtils?.safeString ? calUtils.safeString(item.getProperty("SUMMARY")) : item.getProperty("SUMMARY");
      }
    } catch (_) {}
    return null;
  })();
  const startTs = toNumber(safeItemProperty(item, "X-NCTALK-START", calUtils)) ?? extractStartTimestampFromItem(item);
  const endTs = extractEndTimestampFromItem(item);
  return {
    title,
    token: safeItemProperty(item, "X-NCTALK-TOKEN", calUtils),
    lobbyEnabled: calUtils?.parseBooleanProp ? calUtils.parseBooleanProp(safeItemProperty(item, "X-NCTALK-LOBBY", calUtils)) : null,
    startTimestamp: startTs,
    endTimestamp: endTs,
    eventConversation: (() => {
      const raw = safeItemProperty(item, "X-NCTALK-EVENT", calUtils);
      if (!raw) return null;
      return raw.trim().toLowerCase() === "event";
    })(),
    objectId: safeItemProperty(item, "X-NCTALK-OBJECTID", calUtils),
    delegateId: safeItemProperty(item, "X-NCTALK-DELEGATE", calUtils),
    delegateName: safeItemProperty(item, "X-NCTALK-DELEGATE-NAME", calUtils),
    delegated: calUtils?.parseBooleanProp ? calUtils.parseBooleanProp(safeItemProperty(item, "X-NCTALK-DELEGATED", calUtils)) : null,
    delegateReady: calUtils?.parseBooleanProp ? calUtils.parseBooleanProp(safeItemProperty(item, "X-NCTALK-DELEGATE-READY", calUtils)) : null
  };
}

/**
 * Convert date-like objects to unix seconds.
 */
function getUnixSecondsFromDate(date) {
  if (!date) return null;
  try {
    const jsDate = date.jsDate || date.getInTimezone?.(date.timezone || "UTC")?.jsDate;
    if (jsDate && jsDate.getTime) {
      const ms = jsDate.getTime();
      if (!Number.isNaN(ms)) {
        return Math.floor(ms / 1000);
      }
    }
  } catch (_) {}
  const native = date.nativeTime;
  if (typeof native === "number" && Number.isFinite(native)) {
    if (native > 1e12) return Math.floor(native / 1e6);
    if (native > 1e9) return Math.floor(native / 1000);
    return Math.floor(native);
  }
  return null;
}

/**
 * Resolve the event start timestamp from calendar item fields.
 */
function extractStartTimestampFromItem(item) {
  if (!item) return null;
  const sources = [item.startDate, item.entryDate, item.beginDate, item.endDate, item.untilDate];
  for (const source of sources) {
    const value = getUnixSecondsFromDate(source);
    if (value != null) return value;
  }
  return null;
}

/**
 * Resolve the event end timestamp from calendar item fields.
 * @param {object} item
 * @returns {number|null}
 */
function extractEndTimestampFromItem(item) {
  if (!item) return null;
  const sources = [item.endDate, item.dueDate, item.untilDate];
  for (const source of sources) {
    const value = getUnixSecondsFromDate(source);
    if (value != null) return value;
  }
  return null;
}

this.calToolbar = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const impl = new CalToolbarExperiment(context);
    context.callOnClose(impl);
    return impl.getAPI();
  }
};
