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

const { classes: Cc, interfaces: Ci } = Components;
var { ExtensionCommon } = ChromeUtils.importESModule("resource://gre/modules/ExtensionCommon.sys.mjs");
var { ExtensionParent } = ChromeUtils.importESModule("resource://gre/modules/ExtensionParent.sys.mjs");
let CAL_MODULE = null;
try{
  CAL_MODULE = ChromeUtils.importESModule("resource:///modules/calendar/calUtils.sys.mjs");
}catch(_){
  try{
    CAL_MODULE = ChromeUtils.import("resource://calendar/modules/calUtils.jsm");
  }catch(__){
    CAL_MODULE = null;
  }
}
const CAL = CAL_MODULE?.cal || null;

if (typeof Services === "undefined" || !Services || !Services.wm){
  throw new Error("Services.wm ist im Experiment-Kontext nicht verfuegbar.");
}

const EXTENSION_ID = "nctalk-direct-esr140@example.com";
const EXTENSION = ExtensionParent.GlobalManager.getExtension(EXTENSION_ID);
if (!EXTENSION){
  throw new Error("Extension " + EXTENSION_ID + " nicht gefunden.");
}

const BRIDGE_SCRIPT_PATH = "ui/calToolbarDialog.js";
const CAL_SHARED_PATH = "ui/calToolbarShared.js";
let CAL_SHARED_URL = "";
let CalUtils = null;
const EVENT_DIALOG_URLS = [
  "chrome://calendar/content/calendar-event-dialog.xhtml",
  "chrome://calendar/content/calendar-event-dialog.xul"
];

const CREATE_HANDLERS = new WeakMap();
const LOBBY_HANDLERS = new WeakMap();
const UTILITY_HANDLERS = new WeakMap();
const ACTIVE_CONTEXTS = new Set();
const TALK_LINK_REGEX = /(https?:\/\/[^\s"'<>]+\/call\/([A-Za-z0-9_-]+))/i;
let stopCalendarObservers = null;
let pendingEventDialogReleases = 0;
let calendarObserverRefs = 0;
let activeEventDialogRefs = 0;
const TAB_WINDOW_CLEANUPS = new Map();
const STANDALONE_WINDOW_CLEANUPS = new Map();
const EVENT_WINDOW_CLEANUPS = new Map();
const WINDOW_REGISTRY = new Map();
let WINDOW_ID_COUNTER = 1;

function ensureCalUtils(context){
  if (CalUtils) return CalUtils;
  const url = CAL_SHARED_URL || (context?.extension?.getURL ? context.extension.getURL(CAL_SHARED_PATH) : "");
  if (!url) return null;
  CAL_SHARED_URL = url;
  try{
    const globalScope = typeof globalThis !== "undefined" ? globalThis : this;
    Services.scriptloader.loadSubScript(url, globalScope, "UTF-8");
  }catch(e){
    console.error("[NCExp] shared utils load failed", e);
    return null;
  }
  CalUtils = (typeof globalThis !== "undefined" ? globalThis.NCTalkCalUtils : this.NCTalkCalUtils) || null;
  return CalUtils;
}

function ensureCalUtilsInWindow(win){
  if (!win) return null;
  if (win.NCTalkCalUtils) return win.NCTalkCalUtils;
  if (!CAL_SHARED_URL) return null;
  try{
    Services.scriptloader.loadSubScript(CAL_SHARED_URL, win, "UTF-8");
  }catch(e){
    console.error("[NCExp] shared utils window load failed", e);
    return null;
  }
  return win.NCTalkCalUtils || null;
}

/**
 * Return localized strings for the experiment.
 */
function i18n(key, substitutions = []){
  const localeData = EXTENSION?.localeData;
  if (localeData && typeof localeData.localizeMessage === "function"){
    const message = localeData.localizeMessage(key, substitutions);
    if (message){
      return message;
    }
  }
  if (Array.isArray(substitutions) && substitutions.length){
    return String(substitutions[0] ?? "");
  }
  return key;
}

/**
 * Return (and create) the handler set for a given context.
 */
function getHandlerSet(store, context){
  let set = store.get(context);
  if (!set){
    set = new Set();
    store.set(context, set);
  }
  return set;
}

/**
 * Invoke all handlers in a set with the same payload.
 */
async function dispatchHandlerSet(set, payload){
  if (!set || set.size === 0) return null;
  for (const fire of set){
    try{
      const result = await fire.async(payload);
      if (result !== undefined){
        return result;
      }
    }catch(e){
      console.error("[NCExp] handler error", e);
    }
  }
  return null;
}

/**
 * Broadcast payloads to all registered window contexts.
 */
async function broadcastUtilityPayload(payload){
  if (!payload) return [];
  const tasks = [];
  for (const context of ACTIVE_CONTEXTS){
    const set = UTILITY_HANDLERS.get(context);
    if (set && set.size){
      tasks.push(dispatchHandlerSet(set, payload));
    }
  }
  if (!tasks.length){
    return [];
  }
  try{
    return await Promise.all(tasks);
  }catch(e){
    console.error("[NCExp] broadcast utility error", e);
    return [];
  }
}

function retainCalendarObservers(reason){
  if (reason === "event-dialog" && activeEventDialogRefs <= 0){
    console.warn("[NCExp] blocked unexpected event-dialog retain", new Error().stack);
    return;
  }
  calendarObserverRefs++;
  if (calendarObserverRefs === 1){
    ensureCalendarObservers();
    if (stopCalendarObservers){
      console.log("[NCExp] calendar observers enabled", { reason: reason || "auto" });
    }
  }
}

function releaseCalendarObservers(reason){
  if (calendarObserverRefs === 0){
    return;
  }
  calendarObserverRefs--;
  if (calendarObserverRefs === 0 && stopCalendarObservers){
    try{
      stopCalendarObservers();
    }catch(_){}
    stopCalendarObservers = null;
    console.log("[NCExp] calendar observers disabled", { reason: reason || "auto" });
  }
}

function flushPendingEventDialogReleases(){
  if (!pendingEventDialogReleases) return;
  if (activeEventDialogRefs > 0) return;
  while (pendingEventDialogReleases > 0){
    pendingEventDialogReleases--;
    releaseCalendarObservers("event-dialog");
  }
}

function shouldDelayObserverRelease(win){
  try{
    const item = getCalendarItemFromWindow(win);
    if (!item) return false;
    const meta = extractTalkMetadataFromItem(item);
    if (!meta || !meta.delegateId) return false;
    if (meta.delegated === true) return false;
    return true;
  }catch(_){
    return false;
  }
}

function ensureWindowId(win){
  if (!win || typeof win !== "object"){
    return null;
  }
  if (win.NCTalkWindowId && WINDOW_REGISTRY.has(win.NCTalkWindowId)){
    return win.NCTalkWindowId;
  }
  const id = WINDOW_ID_COUNTER++;
  win.NCTalkWindowId = id;
  WINDOW_REGISTRY.set(id, win);
  const cleanup = () => releaseWindowId(win);
  win.addEventListener("unload", cleanup, { once: true });
  return id;
}

function releaseWindowId(win){
  if (!win || !win.NCTalkWindowId){
    return;
  }
  const id = win.NCTalkWindowId;
  WINDOW_REGISTRY.delete(id);
  delete win.NCTalkWindowId;
}

function watch3PaneCalendarTabs(win){
  try{
    const doc = win?.document;
    if (!doc) return;
    if (doc.documentElement?.getAttribute("windowtype") !== "mail:3pane"){
      return;
    }
    const startMonitor = () => {
      if (TAB_WINDOW_CLEANUPS.has(win)){
        return;
      }
      const tabmail = doc.getElementById("tabmail");
      const container = tabmail?.tabContainer;
      if (!tabmail || !container){
        return;
      }
      const isCalendarMode = (info) => {
        const modeName = info?.mode?.name ?? info?.mode?.typeName ?? info?.modeName ?? info?.mode ?? "";
        const normalized = String(modeName || "").toLowerCase();
        return normalized.includes("calendar") || normalized.includes("task");
      };
      const evaluate = () => {
        const infos = Array.isArray(tabmail.tabInfo) ? tabmail.tabInfo : [];
        const hasCalendar = infos.some((info) => isCalendarMode(info));
        if (hasCalendar && !win.__NCTalkCalendarTabActive){
          win.__NCTalkCalendarTabActive = true;
          retainCalendarObservers("calendar-tab");
        }else if (!hasCalendar && win.__NCTalkCalendarTabActive){
          win.__NCTalkCalendarTabActive = false;
          releaseCalendarObservers("calendar-tab");
        }
      };
      const onTabEvent = () => win.setTimeout(evaluate, 0);
      const cleanup = () => {
        container.removeEventListener("TabSelect", onTabEvent);
        container.removeEventListener("TabClose", onTabEvent);
        container.removeEventListener("TabOpen", onTabEvent);
        win.removeEventListener("unload", cleanup);
        TAB_WINDOW_CLEANUPS.delete(win);
        if (win.__NCTalkCalendarTabActive){
          win.__NCTalkCalendarTabActive = false;
          releaseCalendarObservers("calendar-tab");
        }
      };
      container.addEventListener("TabSelect", onTabEvent);
      container.addEventListener("TabClose", onTabEvent);
      container.addEventListener("TabOpen", onTabEvent);
      win.addEventListener("unload", cleanup);
      TAB_WINDOW_CLEANUPS.set(win, cleanup);
      evaluate();
    };
    if (doc.readyState === "complete"){
      startMonitor();
    }else{
      const onReady = () => {
        doc.removeEventListener("DOMContentLoaded", onReady);
        startMonitor();
      };
      doc.addEventListener("DOMContentLoaded", onReady);
    }
  }catch(e){
    console.error("[NCExp] tab window monitor failed", e);
  }
}
function watchStandaloneCalendarWindow(win){
  try{
    const doc = win?.document;
    if (!doc) return;
    if (isEventDialogWindow(win)){
      return;
    }
    const windowType = doc.documentElement?.getAttribute("windowtype") || "";
    const href = String(doc.location?.href || "");
    const isStandalone = windowType === "calendarMainWindow" || windowType === "calendar:calendar" || href.startsWith("chrome://calendar/content/");
    if (!isStandalone){
      return;
    }
    if (STANDALONE_WINDOW_CLEANUPS.has(win)){
      return;
    }
    const cleanup = () => {
      win.removeEventListener("unload", cleanup);
      STANDALONE_WINDOW_CLEANUPS.delete(win);
      if (win.__NCTalkStandaloneActive){
        win.__NCTalkStandaloneActive = false;
        releaseCalendarObservers("calendar-window");
      }
    };
    win.addEventListener("unload", cleanup);
    STANDALONE_WINDOW_CLEANUPS.set(win, cleanup);
    if (!win.__NCTalkStandaloneActive){
      win.__NCTalkStandaloneActive = true;
      retainCalendarObservers("calendar-window");
    }
  }catch(e){
    console.error("[NCExp] standalone calendar monitor failed", e);
  }
}

function trackCalendarWindowPresence(win){
  if (isEventDialogWindow(win)){
    activateEventDialogWindow(win);
    return;
  }
  watch3PaneCalendarTabs(win);
  watchStandaloneCalendarWindow(win);
}

function activateEventDialogWindow(win){
  if (!win || EVENT_WINDOW_CLEANUPS.has(win)){
    return;
  }
  const cleanup = () => {
    win.removeEventListener("unload", cleanup);
    EVENT_WINDOW_CLEANUPS.delete(win);
    if (win.__NCTalkEventActive){
      releaseEventDialogWindow(win, "event-dialog");
    }
  };
  win.addEventListener("unload", cleanup);
  EVENT_WINDOW_CLEANUPS.set(win, cleanup);
  if (win.__NCTalkEventActive !== true){
    win.__NCTalkEventActive = false;
  }
}

function engageEventDialogWindow(win, reason = "event-dialog", source = "unknown"){
  console.log("[NCExp] engageEventDialogWindow", { source, reason });
  if (!win) return false;
  activateEventDialogWindow(win);
  if (win.__NCTalkEventActive){
    return false;
  }
  win.__NCTalkEventActive = true;
  activeEventDialogRefs++;
  retainCalendarObservers(reason || "event-dialog");
  return true;
}

function releaseEventDialogWindow(win, reason = "event-dialog"){
  if (!win || !win.__NCTalkEventActive){
    return false;
  }
  win.__NCTalkEventActive = false;
  if (activeEventDialogRefs > 0){
    activeEventDialogRefs--;
  }
  const releaseReason = reason || "event-dialog";
  if (releaseReason === "event-dialog" && shouldDelayObserverRelease(win)){
    pendingEventDialogReleases++;
    console.log("[NCExp] calendar observer release pending", { pending: pendingEventDialogReleases });
    return true;
  }
  releaseCalendarObservers(releaseReason);
  return true;
}

function stopAllCalendarActivity(){
  for (const cleanup of Array.from(TAB_WINDOW_CLEANUPS.values())){
    try{
      cleanup();
    }catch(e){
      console.error("[NCExp] tab window cleanup failed", e);
    }
  }
  for (const cleanup of Array.from(STANDALONE_WINDOW_CLEANUPS.values())){
    try{
      cleanup();
    }catch(e){
      console.error("[NCExp] standalone window cleanup failed", e);
    }
  }
  for (const cleanup of Array.from(EVENT_WINDOW_CLEANUPS.values())){
    try{
      cleanup();
    }catch(e){
      console.error("[NCExp] event window cleanup failed", e);
    }
  }
  for (const win of Array.from(WINDOW_REGISTRY.values())){
    releaseWindowId(win);
  }
}

function getRegisteredWindow(windowId){
  if (typeof windowId !== "number"){
    return null;
  }
  const win = WINDOW_REGISTRY.get(windowId);
  if (!win || win.closed || !win.document){
    WINDOW_REGISTRY.delete(windowId);
    return null;
  }
  return win;
}

function engageEventObserversById(windowId, reason = "event-dialog", source = "unknown"){
  const win = getRegisteredWindow(windowId);
  if (!win) return false;
  return engageEventDialogWindow(win, reason, source);
}

function releaseEventObserversById(windowId, reason = "event-dialog"){
  const win = getRegisteredWindow(windowId);
  if (!win) return false;
  return releaseEventDialogWindow(win, reason);
}

async function invokeWindowAction(options = {}){
  const windowId = options.windowId;
  const action = options.action;
  const payload = options.payload || {};
  if (typeof windowId !== "number" || !action){
    throw new Error("windowId and action are required");
  }
  const win = getRegisteredWindow(windowId);
  if (!win){
    throw new Error("target window not available");
  }
  return await executeWindowAction(win, action, payload);
}

async function executeWindowAction(win, action, payload){
switch (action){
  case "ping":
    return { ok:true };
  case "getTalkMetadata":
    return getTalkMetadataFromWindow(win);
  case "setTalkMetadata":
    return setTalkMetadataOnWindow(win, payload || {});
  case "getEventSnapshot":
    return getEventSnapshotFromWindow(win);
  case "applyEventFields":
    return applyEventFieldsOnWindow(win, payload || {});
  case "releaseEventObservers":
    releaseEventDialogWindow(win, payload?.reason || "event-dialog");
    return { ok:true };
  case "registerCleanup":
    if (typeof win.NCTalkRegisterCleanup === "function"){
      return await win.NCTalkRegisterCleanup(payload || {});
    }
    return { ok:false, error:"register_cleanup_unavailable" };
  default:
    return { ok:false, error:"unknown_action" };
}
}
/**
 * Read a calendar item property safely.
 */
function safeItemProperty(item, prop){
  if (!item || typeof item.getProperty !== "function") return null;
  try{
    return CalUtils.safeString(item.getProperty(prop));
  }catch(_){
    return null;
  }
}


/**
 * Resolve the calendar item from an event dialog window context.
 */
function getCalendarItemFromWindow(win){
  try{
    if (!win) return null;
    if (win.calendarItem) return win.calendarItem;
    if (win.gEvent && win.gEvent.event) return win.gEvent.event;
    if (Array.isArray(win.arguments) && win.arguments[0]){
      const arg = win.arguments[0];
      if (arg.calendarItem) return arg.calendarItem;
      if (arg.calendarEvent) return arg.calendarEvent;
    }
  }catch(_){}
  return null;
}

function getTalkMetadataFromWindow(win){
  const item = getCalendarItemFromWindow(win);
  if (!item){
    return { ok:false, error:"no_calendar_item" };
  }
  const metadata = extractTalkMetadataFromItem(item) || {};
  return { ok:true, metadata };
}

function setTalkMetadataOnWindow(win, meta = {}){
  return CalUtils.setTalkMetadataOnWindow(win, meta);
}

function getEventSnapshotFromWindow(win){
  const metaResult = getTalkMetadataFromWindow(win);
  const metadata = metaResult.ok ? (metaResult.metadata || {}) : {};
  return CalUtils.getEventSnapshotFromWindow(win, { metadata });
}

function applyEventFieldsOnWindow(win, payload = {}){
  const entry = Array.from(WINDOW_REGISTRY.entries()).find(([, candidate]) => candidate === win);
  console.log("[NCExp] applyEventFields", {
    windowId: entry ? entry[0] : null,
    title: payload?.title || "",
    hasLocation: !!payload?.location,
    hasDescription: typeof payload?.description === "string"
  });
  const result = CalUtils.applyEventFieldsOnWindow(win, payload, { preferExecForDescription: true });
  const applied = result?.applied || {};
  try{
    win?.NCTalkLog?.("applyEventFields", {
      titleApplied: !!applied.title,
      locationApplied: !!applied.location,
      descriptionApplied: !!applied.description
    });
  }catch(_){}
  return result || { ok:false, error:"apply_failed" };
}

/**
 * Extract Talk-specific metadata from a calendar item.
 */
function extractTalkMetadataFromItem(item){
  if (!item) return {};
  const toNumber = (raw) => {
    if (raw == null) return null;
    const num = Number(raw);
    return Number.isFinite(num) ? num : null;
  };
  const title = CalUtils.safeString(item?.title) || (() => {
    try{
      if (typeof item.getProperty === "function"){
        return CalUtils.safeString(item.getProperty("SUMMARY"));
      }
    }catch(_){}
    return null;
  })();
  const startTs = toNumber(safeItemProperty(item, "X-NCTALK-START")) ?? extractStartTimestampFromItem(item);
  const endTs = extractEndTimestampFromItem(item);
  return {
    title,
    token: safeItemProperty(item, "X-NCTALK-TOKEN"),
    lobbyEnabled: CalUtils.parseBooleanProp(safeItemProperty(item, "X-NCTALK-LOBBY")),
    startTimestamp: startTs,
    endTimestamp: endTs,
    eventConversation: (() => {
      const raw = safeItemProperty(item, "X-NCTALK-EVENT");
      if (!raw) return null;
      return raw.trim().toLowerCase() === "event";
    })(),
    objectId: safeItemProperty(item, "X-NCTALK-OBJECTID"),
    delegateId: safeItemProperty(item, "X-NCTALK-DELEGATE"),
    delegateName: safeItemProperty(item, "X-NCTALK-DELEGATE-NAME"),
    delegated: CalUtils.parseBooleanProp(safeItemProperty(item, "X-NCTALK-DELEGATED")),
    delegateReady: CalUtils.parseBooleanProp(safeItemProperty(item, "X-NCTALK-DELEGATE-READY"))
  };
}

/**
 * Mark an item as having its delegation applied.
 */
function markDelegationApplied(item){
  try{
    if (!item || typeof item.clone !== "function") return false;
    if (CalUtils.parseBooleanProp(safeItemProperty(item, "X-NCTALK-DELEGATED")) === true) return false;
    const calendar = item.calendar;
    if (!calendar || typeof calendar.modifyItem !== "function") return false;
    const clone = item.clone();
    clone.setProperty("X-NCTALK-DELEGATED", "TRUE");
    try{
      if (typeof clone.deleteProperty === "function"){
        clone.deleteProperty("X-NCTALK-DELEGATE-READY");
      }else{
        clone.setProperty("X-NCTALK-DELEGATE-READY", "");
      }
    }catch(_){}
    calendar.modifyItem(clone, item, null);
    console.log("[NCExp] delegation flag persisted", { token: safeItemProperty(item, "X-NCTALK-TOKEN") });
    return true;
  }catch(e){
    console.error("[NCExp] failed to persist delegation flag", e);
    return false;
  }
}

async function triggerDelegationIfReady(item, meta, context = {}){
  if (!item || !meta) return false;
  const delegateId = meta.delegateId && meta.delegateId.trim();
  if (!delegateId) return false;
  if (meta.delegated === true){
    console.log("[NCExp] calendar delegation already applied", {
      token: context.token || meta.token || null,
      delegate: delegateId
    });
    return false;
  }
  const readyState = meta.delegateReady;
  const legacyMode = safeItemProperty(item, "X-NCTALK-DELEGATE-READY") == null;
  console.log("[NCExp] calendar delegation state", {
    token: context.token || meta.token || null,
    delegate: delegateId,
    readyState,
    legacyMode,
    delegated: meta.delegated === true,
    reason: context.reason || "unknown"
  });
  if (readyState !== true && !legacyMode){
    console.log("[NCExp] calendar delegation skipped (not armed)", {
      token: context.token || meta.token || null,
      delegate: delegateId,
      reason: context.reason || "not-ready"
    });
    return false;
  }
  console.log("[NCExp] calendar delegation trigger", {
    token: context.token || meta.token || null,
    delegate: delegateId,
    reason: context.reason || "post-lobby"
  });
  const payload = {
    type: "calendarDelegateModerator",
    token: context.token || meta.token || null,
    delegateId,
    delegateName: meta.delegateName || delegateId,
    objectId: meta.objectId || null,
    startTimestamp: context.startTimestamp ?? null,
    uid: safeItemUid(item)
  };
  const results = await broadcastUtilityPayload(payload);
  const success = Array.isArray(results) && results.some((res) => res && res.ok);
  if (success){
    markDelegationApplied(item);
  }
  return success;
}
/**
 * Extract the Talk link from item properties or text fields.
 */
function extractTalkLinkFromItem(item){
  if (!item) return null;
  const propToken = safeItemProperty(item, "X-NCTALK-TOKEN");
  if (propToken){
    const propUrl = safeItemProperty(item, "X-NCTALK-URL");
    return {
      token: propToken,
      url: propUrl || ("https://nextcloud.local/call/" + propToken)
    };
  }
  const candidates = [];
  const location = safeItemProperty(item, "LOCATION") || CalUtils.safeString(item.location);
  if (location) candidates.push(location);
  const description = safeItemProperty(item, "DESCRIPTION") || CalUtils.safeString(item.description);
  if (description) candidates.push(description);
  const url = safeItemProperty(item, "URL");
  if (url) candidates.push(url);
  const summary = safeItemProperty(item, "SUMMARY") || CalUtils.safeString(item.title);
  if (summary && summary.includes("/call/")){
    candidates.push(summary);
  }
  for (const text of candidates){
    const match = TALK_LINK_REGEX.exec(text);
    if (match){
      return { url: match[1], token: match[2] };
    }
  }
  return null;
}

/**
 * Return the item UID when available.
 */
function safeItemUid(item){
  if (!item) return null;
  const inputs = [
    item.id,
    item.uid,
    item.hashId,
    item?.parentItem?.id,
    item?.parentItem?.uid
  ];
  for (const value of inputs){
    if (CalUtils.safeString(value)){
      return String(value);
    }
  }
  return null;
}

/**
 * Convert date-like objects to unix seconds.
 */
function getUnixSecondsFromDate(date){
  if (!date) return null;
  try{
    const jsDate = date.jsDate || date.getInTimezone?.(date.timezone || "UTC")?.jsDate;
    if (jsDate && jsDate.getTime){
      const ms = jsDate.getTime();
      if (!Number.isNaN(ms)){
        return Math.floor(ms / 1000);
      }
    }
  }catch(_){}
  const native = date.nativeTime;
  if (typeof native === "number" && Number.isFinite(native)){
    if (native > 1e12) return Math.floor(native / 1e6);
    if (native > 1e9) return Math.floor(native / 1000);
    return Math.floor(native);
  }
  return null;
}

/**
 * Resolve the event start timestamp from calendar item fields.
 */
function extractStartTimestampFromItem(item){
  if (!item) return null;
  const sources = [item.startDate, item.entryDate, item.beginDate, item.endDate, item.untilDate];
  for (const source of sources){
    const value = getUnixSecondsFromDate(source);
    if (value != null) return value;
  }
  return null;
}

function extractEndTimestampFromItem(item){
  if (!item) return null;
  const sources = [item.endDate, item.dueDate, item.untilDate];
  for (const source of sources){
    const value = getUnixSecondsFromDate(source);
    if (value != null) return value;
  }
  return null;
}

/**
 * Handle deleted items and clean up metadata.
 */
async function handleCalendarDelete(item){
  try{
    if (!item) return;
    const link = extractTalkLinkFromItem(item);
    if (!link || !link.token) return;
    await broadcastUtilityPayload({
      type: "deleteRoom",
      token: link.token,
      url: link.url,
      reason: "calendarDelete",
      uid: safeItemUid(item)
    });
  }finally{
    flushPendingEventDialogReleases();
  }
}

/**
 * Handle modified items and update lobby metadata.
 */
async function handleCalendarModify(newItem, oldItem){
  try{
    if (!newItem) return;
    const link = extractTalkLinkFromItem(newItem);
    if (!link || !link.token) return;
    const meta = extractTalkMetadataFromItem(newItem);
    const oldMeta = extractTalkMetadataFromItem(oldItem);
    const calculatedNewStart = extractStartTimestampFromItem(newItem);
    const calculatedOldStart = extractStartTimestampFromItem(oldItem);
    const newStart = calculatedNewStart ?? meta.startTimestamp ?? calculatedOldStart;
    const oldStart = oldMeta.startTimestamp ?? calculatedOldStart ?? meta.startTimestamp;
    const lobbyEnabled = meta.lobbyEnabled !== false;

    let lobbyUpdated = false;
    if (lobbyEnabled && newStart != null){
      const startChanged = oldStart == null || Math.abs(oldStart - newStart) >= 1;
      if (startChanged){
        console.log("[NCExp] calendar lobby update", {
          token: link.token,
          startTimestamp: newStart,
          delegate: meta.delegateId || ""
        });
        await broadcastUtilityPayload({
          type: "calendarUpdateLobby",
          token: link.token,
          startTimestamp: newStart,
          delegateId: meta.delegateId || "",
          delegated: meta.delegated === true,
          uid: safeItemUid(newItem),
          url: link.url
        });
        persistTalkStart(newItem, newStart);
        lobbyUpdated = true;
      }
    }

    await triggerDelegationIfReady(newItem, meta, {
      token: link.token,
      startTimestamp: newStart,
      reason: lobbyUpdated ? "post-lobby-update" : "pending-lobby"
    });
  }finally{
    flushPendingEventDialogReleases();
  }
}
/**
 * Install observers for all available calendars.
 */
function installCalendarObservers(){
  if (!CAL || !CAL.manager) return null;
  console.log("[NCExp] calendar observers init");
  const manager = CAL.manager;
  const tracked = new Set();
  const calendarObserver = {
    QueryInterface: ChromeUtils.generateQI
      ? ChromeUtils.generateQI(["calIObserver"])
      : undefined,
    onStartBatch() {},
    onEndBatch() {},
    onLoad() {},
    onAddItem(newItem) {
      handleCalendarModify(newItem, null).catch((err) => console.error("[NCExp] calendar add", err));
    },
    onModifyItem(newItem, oldItem) {
      handleCalendarModify(newItem, oldItem).catch((err) => console.error("[NCExp] calendar modify", err));
    },
    onDeleteItem(item) {
      handleCalendarDelete(item).catch((err) => console.error("[NCExp] calendar delete", err));
    },
    onError() {},
    onPropertyChanged() {},
    onPropertyDeleting() {}
  };
  const managerObserver = {
    QueryInterface: ChromeUtils.generateQI
      ? ChromeUtils.generateQI(["calICalendarManagerObserver"])
      : undefined,
    onCalendarRegistered(calendar) {
      addCalendar(calendar);
    },
    onCalendarUnregistering(calendar) {
      removeCalendar(calendar);
    },
    onCalendarDeleting(calendar) {
      removeCalendar(calendar);
    },
    onCalendarPrefChanged() {},
    onCalendarSettingChanged() {},
    onCalendarPropertyChanged() {}
  };

  /**
   * Register a calendar observer when a calendar is added.
   */
  function addCalendar(calendar){
    if (!calendar || tracked.has(calendar)) return;
    try{
      calendar.addObserver(calendarObserver);
      tracked.add(calendar);
    }catch(e){
      console.error("[NCExp] calendar observer add failed", e);
    }
  }

  /**
   * Remove the calendar observer when a calendar is removed.
   */
  function removeCalendar(calendar){
    if (!calendar || !tracked.has(calendar)) return;
    try{
      calendar.removeObserver(calendarObserver);
    }catch(_){}
    tracked.delete(calendar);
  }

  try{
    const calendars = manager.getCalendars({});
    if (Array.isArray(calendars)){
      for (const calendar of calendars){
        addCalendar(calendar);
      }
    }
    manager.addObserver(managerObserver);
  }catch(e){
    console.error("[NCExp] calendar manager observer failed", e);
    for (const calendar of Array.from(tracked)){
      removeCalendar(calendar);
    }
    return null;
  }

  return () => {
    try{
      manager.removeObserver(managerObserver);
    }catch(_){}
    for (const calendar of Array.from(tracked)){
      removeCalendar(calendar);
    }
    console.log("[NCExp] calendar observers stopped");
  };
}

/**
 * Ensure calendar observers are installed.
 */
function ensureCalendarObservers(){
  if (stopCalendarObservers || !CAL || !CAL.manager) return;
  stopCalendarObservers = installCalendarObservers();
}

/**
 * Check whether a window is a calendar event dialog.
 */
function isEventDialogWindow(win){
  if (!win) return false;
  try{
    const doc = win.document;
    const windowType = doc?.documentElement?.getAttribute("windowtype") || "";
    if (windowType === "Calendar:EventDialog" || windowType === "Calendar:EventSummaryDialog"){
      return true;
    }
    const href = doc?.location?.href || win.location?.href || "";
    if (!href) return false;
    return EVENT_DIALOG_URLS.some((url) => href.startsWith(url));
  }catch(_){
    return false;
  }
}

/**
 * Create the bridge API exposed to the dialog window.
 */
function createBridgeAPI(context, init = {}){
  const state = {
    label: init.label || i18n("ui_insert_button_label"),
    tooltip: init.tooltip || i18n("ui_toolbar_tooltip"),
    windowId: init.windowId || null
  };
  return {
    get label(){ return state.label; },
    get tooltip(){ return state.tooltip; },
    get windowId(){ return state.windowId; },
    updateInit(opts = {}){
      if (typeof opts.label === "string") state.label = opts.label;
      if (typeof opts.tooltip === "string") state.tooltip = opts.tooltip;
      if (typeof opts.windowId === "number") state.windowId = opts.windowId;
    },
    i18n: (key, subs) => i18n(key, subs),
    getURL: (path) => context.extension?.getURL(path || "") || "",
    requestCreate: (payload) => dispatchHandlerSet(getHandlerSet(CREATE_HANDLERS, context), payload || {}),
    requestLobby: (payload) => dispatchHandlerSet(getHandlerSet(LOBBY_HANDLERS, context), payload || {}),
    requestUtility: (payload = {}) => {
      if (payload && payload.type === "openDialog" && typeof state.windowId === "number"){
        engageEventObserversById(
          state.windowId,
          payload?.reason || "event-dialog",
          "utility/openDialog"
        );
      }
      return dispatchHandlerSet(getHandlerSet(UTILITY_HANDLERS, context), payload);
    },
    engageEventObservers: (reason = "event-dialog") => {
      if (typeof state.windowId === "number"){
        engageEventObserversById(state.windowId, reason, "bridge.manual");
      }
    },
    openDialog: async () => {
      const windowId = state.windowId || null;
      if (!context?.extension?.browser?.runtime?.sendMessage){
        throw new Error("runtime messaging unavailable");
      }
      return await context.extension.browser.runtime.sendMessage({
        type: "talk:openDialog",
        windowId
      });
    }
  };
}

/**
 * Dispatch a bridge refresh event into the target window.
 */
function dispatchBridgeRefresh(win){
  const eventCtor = win?.CustomEvent || win?.Event;
  if (typeof win?.dispatchEvent === "function" && typeof eventCtor === "function"){
    win.dispatchEvent(new eventCtor("nctalk-bridge-refresh"));
  }
}

/**
 * Inject the bridge into a newly opened event window.
 */
function installBridge(win, context, init){
  if (!win || !win.document) return false;
  const windowId = ensureWindowId(win);
  const initWithId = Object.assign({}, init, { windowId });
  const api = createBridgeAPI(context, initWithId);
  ensureCalUtils(context);
  ensureCalUtilsInWindow(win);
  if (win.NCTalkBridge && typeof win.NCTalkBridge.updateInit === "function"){
    win.NCTalkBridge.updateInit(initWithId);
    dispatchBridgeRefresh(win);
  }else{
    Object.defineProperty(win, "NCTalkBridge", {
      value: api,
      configurable: true
    });
  }
  if (win.NCTalkBridgeLoaded){
    dispatchBridgeRefresh(win);
    return true;
  }
  try{
    Services.scriptloader.loadSubScript(
      context.extension.getURL(BRIDGE_SCRIPT_PATH),
      win,
      "UTF-8"
    );
    win.NCTalkBridgeLoaded = true;
    dispatchBridgeRefresh(win);
    return true;
  }catch(e){
    console.error("[NCExp] bridge script load failed", e);
    return false;
  }
}

/**
 * Register listeners for new and closing windows.
 */
function registerWindowListener(context, init){
  const listener = {
    onOpenWindow(xulWindow){
      const docShell = xulWindow?.docShell || null;
      const win = docShell?.DOMWindow || docShell?.domWindow || null;
      if (!win) return;
      const inject = () => {
        trackCalendarWindowPresence(win);
        if (isEventDialogWindow(win)){
          if (installBridge(win, context, init)){
            console.log("[NCExp] bridge injected window");
            activateEventDialogWindow(win);
          }
        }
      };
      if (win.document?.readyState === "complete"){
        inject();
      }else{
        const onLoad = () => {
          win.removeEventListener("load", onLoad);
          inject();
        };
        win.addEventListener("load", onLoad, { once: true });
      }
    },
    onCloseWindow(){},
    onWindowTitleChange(){}
  };

  Services.wm.addListener(listener);
  const enumerator = Services.wm.getEnumerator(null);
  if (enumerator){
    while (enumerator.hasMoreElements()){
      const win = enumerator.getNext();
      const domWindow = win?.docShell?.DOMWindow || win?.docShell?.domWindow || win;
      if (!domWindow) continue;
      trackCalendarWindowPresence(domWindow);
      if (isEventDialogWindow(domWindow)){
        if (installBridge(domWindow, context, init)){
          console.log("[NCExp] bridge injected existing window");
          activateEventDialogWindow(domWindow);
        }
      }
    }
  }

  let active = true;
  return () => {
    if (!active) return;
    active = false;
    Services.wm.removeListener(listener);
    stopAllCalendarActivity();
  };
}

this.calToolbar = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    ACTIVE_CONTEXTS.add(context);
    context.callOnClose(() => {
      ACTIVE_CONTEXTS.delete(context);
    });
    ensureCalUtils(context);
    const createEvent = new ExtensionCommon.EventManager({
      context,
      name: "calToolbar.onCreateRequest",
      register: (fire) => {
        getHandlerSet(CREATE_HANDLERS, context).add(fire);
        return () => getHandlerSet(CREATE_HANDLERS, context).delete(fire);
      }
    });
    const lobbyEvent = new ExtensionCommon.EventManager({
      context,
      name: "calToolbar.onLobbyUpdate",
      register: (fire) => {
        getHandlerSet(LOBBY_HANDLERS, context).add(fire);
        return () => getHandlerSet(LOBBY_HANDLERS, context).delete(fire);
      }
    });
    const utilityEvent = new ExtensionCommon.EventManager({
      context,
      name: "calToolbar.onUtilityRequest",
      register: (fire) => {
        getHandlerSet(UTILITY_HANDLERS, context).add(fire);
        return () => getHandlerSet(UTILITY_HANDLERS, context).delete(fire);
      }
    });

    let stopListening = null;

    return {
      calToolbar: {
        onCreateRequest: createEvent.api(),
        onLobbyUpdate: lobbyEvent.api(),
        onUtilityRequest: utilityEvent.api(),
        async init(opts = {}) {
    const label = opts.label || i18n("ui_insert_button_label");
    const tooltip = opts.tooltip || i18n("ui_toolbar_tooltip");
    try{
      if (stopListening){
        stopListening();
        stopListening = null;
      }
      stopListening = registerWindowListener(context, { label, tooltip });
      context.callOnClose(() => {
        if (stopListening){
          stopListening();
          stopListening = null;
        }
      });
      return true;
    }catch(e){
      console.error("[NCExp] init error", e);
      if (stopListening){
        stopListening();
        stopListening = null;
      }
      return false;
    }
        },
        invokeWindow(options = {}){
          return invokeWindowAction(options);
        }
      }
    };
  }
};

function persistTalkStart(item, startTimestamp){
  try{
    if (!item || typeof item.clone !== "function") return false;
    if (typeof startTimestamp !== "number" || !Number.isFinite(startTimestamp)) return false;
    const calendar = item.calendar;
    if (!calendar || typeof calendar.modifyItem !== "function") return false;
    const current = CalUtils.parseNumberProp(safeItemProperty(item, "X-NCTALK-START"));
    if (current === Math.floor(startTimestamp)){
      return false;
    }
    const clone = item.clone();
    clone.setProperty("X-NCTALK-START", String(Math.floor(startTimestamp)));
    calendar.modifyItem(clone, item, null);
    console.log("[NCExp] start timestamp persisted", {
      token: safeItemProperty(item, "X-NCTALK-TOKEN"),
      startTimestamp: Math.floor(startTimestamp)
    });
    return true;
  }catch(e){
    console.error("[NCExp] failed to persist start timestamp", e);
    return false;
  }
}
