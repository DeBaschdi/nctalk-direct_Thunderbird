/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
'use strict';
/**
 * Background script for Nextcloud Talk Direct.
 * Handles API calls (Talk + CardDAV), caching, and messaging utilities.
 */
const ROOM_META_KEY = "nctalkRoomMeta";
let DEBUG_ENABLED = false;
let ROOM_META = {};
const TALK_POPUP_WIDTH = 540;
const TALK_POPUP_HEIGHT = 860;
const FILELINK_POPUP_WIDTH = 660;
const FILELINK_POPUP_HEIGHT = 760;

(async () => {
  try{
    const stored = await browser.storage.local.get(["debugEnabled", ROOM_META_KEY]);
    DEBUG_ENABLED = !!stored.debugEnabled;
    ROOM_META = stored[ROOM_META_KEY] || {};
  }catch(_){ }
})();
browser.storage.onChanged.addListener((changes, area) => {
  if (area !== "local") return;
  if (Object.prototype.hasOwnProperty.call(changes, "debugEnabled")){
    DEBUG_ENABLED = !!changes.debugEnabled.newValue;
  }
  if (Object.prototype.hasOwnProperty.call(changes, ROOM_META_KEY)){
    ROOM_META = changes[ROOM_META_KEY].newValue || {};
  }
});
function L(...a){
  if (!DEBUG_ENABLED) return;
  try{
    console.log("[NCBG]", ...a);
  }catch(_){ }
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


/**
 * Create a localized Error using the i18n catalog.
 */
function localizedError(key, substitutions = []){
  const message = bgI18n(key, substitutions);
  return new Error(message || key);
}

async function openLoginUrl(url){
  if (!url) return;
  if (browser?.windows?.openDefaultBrowser){
    try{
      await browser.windows.openDefaultBrowser(url);
      return;
    }catch(_){}
  }
  try{
    await browser.tabs.create({ url, active: true });
  }catch(_){}
}

browser.composeAction.onClicked.addListener(async (tab) => {
  try{
    const popupUrl = browser.runtime.getURL(`ui/filelinkWizard.html?tabId=${tab.id}`);
    await browser.windows.create({
      url: popupUrl,
      type: "popup",
      width: FILELINK_POPUP_WIDTH,
      height: FILELINK_POPUP_HEIGHT
    });
  }catch(e){
    console.error("[NCBG] composeAction.onClicked", e);
  }
});

async function setRoomMeta(token, data = {}){
  if (!token) return;
  const next = Object.assign({}, ROOM_META[token], data, { updated: Date.now() });
  ROOM_META[token] = next;
  try{
    await browser.storage.local.set({ [ROOM_META_KEY]: ROOM_META });
  }catch(e){
    console.error("[NCBG] setRoomMeta", e);
  }
}

async function deleteRoomMeta(token){
  if (!token || !ROOM_META[token]) return;
  delete ROOM_META[token];
  try{
    await browser.storage.local.set({ [ROOM_META_KEY]: ROOM_META });
  }catch(e){
    console.error("[NCBG] deleteRoomMeta", e);
  }
}

function getRoomMeta(token){
  if (!token) return null;
  return ROOM_META[token] || null;
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
 * Open the Talk dialog popup for a calendar window.
 */
async function openTalkDialogWindow(windowId){
  const url = new URL(browser.runtime.getURL("ui/talkDialog.html"));
  if (typeof windowId === "number"){
    url.searchParams.set("windowId", String(windowId));
  }
  await browser.windows.create({
    url: url.toString(),
    type: "popup",
    width: TALK_POPUP_WIDTH,
    height: TALK_POPUP_HEIGHT
  });
}

browser.runtime.onMessage.addListener(async (msg) => {
  if (!msg || !msg.type) return;
  L("msg", msg.type, { hasPayload: !!msg.payload });
  if (msg.type === "debug:log"){
    const source = msg.payload?.source ? String(msg.payload.source) : "frontend";
    const text = msg.payload?.text ? String(msg.payload.text) : "";
    const extras = Array.isArray(msg.payload?.details)
      ? msg.payload.details
      : (msg.payload?.details != null ? [msg.payload.details] : []);
    const channelRaw = msg.payload?.channel ? String(msg.payload.channel) : "NCDBG";
    const channel = channelRaw.toUpperCase();
    const label = msg.payload?.label ? String(msg.payload.label) : source;
    const prefix = label ? `[${channel}][${label}]` : `[${channel}]`;
    if (DEBUG_ENABLED || channel === "NCUI"){
      try{
        console.log(prefix, text, ...extras);
      }catch(_){ }
    }
    return { ok:true };
  }
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
      const config = await NCCore.getOpts();
      return { ok:true, config };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
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
  if (msg.type === "talkMenu:delegateModerator"){
    try{
      const result = await delegateRoomModerator(msg.payload || {});
      return { ok:true, result };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talk:delegateModerator"){
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
  if (msg.type === "talk:searchUsers"){
    try{
      const users = await searchSystemAddressbook(msg.payload || {});
      return { ok:true, users };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talk:openDialog"){
    const windowId = msg.windowId ?? msg?.payload?.windowId;
    try{
      await openTalkDialogWindow(windowId);
      return { ok:true };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talk:initDialog"){
    const windowId = msg.windowId ?? msg?.payload?.windowId;
    if (typeof windowId !== "number"){
      return { ok:false, error: "windowId required" };
    }
    try{
      const response = await browser.calToolbar.invokeWindow({
        windowId,
        action: "ping",
        payload: {}
      });
      return { ok:true, response };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talk:getEventSnapshot"){
    const windowId = msg.windowId ?? msg?.payload?.windowId;
    if (typeof windowId !== "number"){
      return { ok:false, error: "windowId required" };
    }
    try{
      const response = await browser.calToolbar.invokeWindow({
        windowId,
        action: "getEventSnapshot",
        payload: {}
      });
      return response;
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talk:applyEventFields"){
    const windowId = msg.windowId ?? msg?.payload?.windowId;
    if (typeof windowId !== "number"){
      return { ok:false, error: "windowId required" };
    }
    const fields = msg.fields ?? msg?.payload?.fields ?? {};
    try{
      L("talk:applyEventFields dispatch", {
        windowId,
        hasTitle: !!fields.title,
        hasLocation: !!fields.location,
        hasDescription: typeof fields.description === "string"
      });
      const response = await browser.calToolbar.invokeWindow({
        windowId,
        action: "applyEventFields",
        payload: fields
      });
      L("talk:applyEventFields response", {
        windowId,
        ok: response?.ok,
        error: response?.error || ""
      });
      return response;
    }catch(e){
      L("talk:applyEventFields error", e?.message || String(e));
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talk:createRoom"){
    try{
      const result = await createTalkPublicRoom(msg.payload);
      return { ok:true, result };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talk:trackRoom"){
    try{
      const payload = msg.payload || {};
      const token = msg.token ?? payload.token;
      if (!token){
        return { ok:false, error: "token required" };
      }
      const updates = {};
      if (Object.prototype.hasOwnProperty.call(msg, "lobbyEnabled") || Object.prototype.hasOwnProperty.call(payload, "lobbyEnabled")){
        updates.lobbyEnabled = !!(msg.lobbyEnabled ?? payload.lobbyEnabled);
      }
      if (Object.prototype.hasOwnProperty.call(msg, "eventConversation") || Object.prototype.hasOwnProperty.call(payload, "eventConversation")){
        updates.eventConversation = !!(msg.eventConversation ?? payload.eventConversation);
      }
      const startRaw = msg.startTimestamp ?? payload.startTimestamp;
      if (typeof startRaw === "number" && Number.isFinite(startRaw)){
        updates.startTimestamp = startRaw;
      }
      await setRoomMeta(token, updates);
      return { ok:true };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talk:applyMetadata"){
    const windowId = msg.windowId ?? msg?.payload?.windowId;
    if (typeof windowId !== "number"){
      return { ok:false, error: "windowId required" };
    }
    const meta = msg.metadata ?? msg?.payload?.metadata ?? {};
    try{
      const response = await browser.calToolbar.invokeWindow({
        windowId,
        action: "setTalkMetadata",
        payload: meta
      });
      if (meta?.token){
        const updates = {};
        if (Object.prototype.hasOwnProperty.call(meta, "lobbyEnabled")){
          updates.lobbyEnabled = !!meta.lobbyEnabled;
        }
        if (Object.prototype.hasOwnProperty.call(meta, "eventConversation")){
          updates.eventConversation = !!meta.eventConversation;
        }
        if (typeof meta.startTimestamp === "number" && Number.isFinite(meta.startTimestamp)){
          updates.startTimestamp = meta.startTimestamp;
        }
        if (Object.keys(updates).length){
          await setRoomMeta(meta.token, updates);
        }
      }
      return response;
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talk:registerCleanup"){
    const windowId = msg.windowId ?? msg?.payload?.windowId;
    if (typeof windowId !== "number"){
      return { ok:false, error: "windowId required" };
    }
    const token = msg.token ?? msg?.payload?.token;
    if (!token){
      return { ok:false, error: "token required" };
    }
    const info = msg.info ?? msg?.payload?.info ?? {};
    try{
      const response = await browser.calToolbar.invokeWindow({
        windowId,
        action: "registerCleanup",
        payload: { token, info }
      });
      return response;
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "options:testConnection"){
    try{
      const result = await NCCore.testCredentials(msg.payload || {});
      if (result.ok){
        return { ok:true, message: result.message || "", version: result.version || "" };
      }
      return { ok:false, error: result.message || bgI18n("error_credentials_missing"), code: result.code || "" };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "options:loginFlow"){
    try{
      const baseUrl = NCCore.normalizeBaseUrl(msg.payload?.baseUrl || "");
      if (!baseUrl){
        return { ok:false, error: bgI18n("options_loginflow_missing") };
      }
      const start = await NCCore.startLoginFlow(baseUrl);
      await openLoginUrl(start.loginUrl);
      const creds = await NCCore.completeLoginFlow({
        pollEndpoint: start.pollEndpoint,
        pollToken: start.pollToken
      });
      return { ok:true, user: creds.loginName, appPass: creds.appPassword };
    }catch(e){
      return { ok:false, error: e?.message || bgI18n("options_loginflow_failed") };
    }
  }
  if (msg.type === "filelink:insertHtml"){
    try{
      const tabId = msg.payload?.tabId;
      const html = msg.payload?.html || "";
      if (!tabId || !html){
        return { ok:false, error: "tab/html missing" };
      }
      const details = await browser.compose.getComposeDetails(tabId);
      const currentBody = details.body || "";
      const blockSegment = `<br>${html}<br><br>`;
      const bodyMatch = currentBody.match(/<body[^>]*>/i);
      let newBody = "";
      if (bodyMatch){
        const insertIndex = bodyMatch.index + bodyMatch[0].length;
        newBody = currentBody.slice(0, insertIndex) + blockSegment + currentBody.slice(insertIndex);
      }else{
        newBody = blockSegment + currentBody;
      }
      await browser.compose.setComposeDetails(tabId, { body: newBody, isPlainText: false });
      return { ok:true };
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
  const type = payload?.type || "";
  L("calToolbar.onUtilityRequest", {
    type,
    token: payload?.token ? shortToken(payload.token) : "",
    searchTerm: payload?.searchTerm || "",
    delegate: payload?.newModerator || ""
  });

  const handlers = {
    debugLog: async () => {
      if (DEBUG_ENABLED){
        const channel = String(payload?.channel || "NCDBG").toUpperCase();
        const label = payload?.label ? String(payload.label) : "";
        const prefix = label ? `[${channel}][${label}]` : `[${channel}]`;
        const textLog = payload?.text ? String(payload.text) : "";
        const extras = Array.isArray(payload?.details)
          ? payload.details
          : (payload?.details != null ? [payload.details] : []);
        console.log(prefix, textLog, ...extras);
      }
      return { ok:true };
    },
    trackRoom: async () => {
      if (!payload?.token){
        return { ok:false, error: bgI18n("error_room_token_missing") };
      }
      const updates = {
        lobbyEnabled: !!payload.lobbyEnabled,
        eventConversation: !!payload.eventConversation
      };
      if (typeof payload?.startTimestamp === "number"){
        updates.startTimestamp = payload.startTimestamp;
      }
      await setRoomMeta(payload.token, updates);
      return { ok:true };
    },
    untrackRoom: async () => {
      await deleteRoomMeta(payload?.token);
      return { ok:true };
    },
    calendarUpdateLobby: async () => {
      const token = payload?.token;
      if (!token){
        return { ok:false, error: bgI18n("error_room_token_missing") };
      }
      const meta = getRoomMeta(token) || {};
      const { user: currentUserRaw } = await NCCore.getOpts();
      const delegateIdRaw = (payload?.delegateId ?? meta.delegateId ?? "").trim();
      const delegateTarget = delegateIdRaw.toLowerCase();
      const currentUser = (currentUserRaw || "").trim().toLowerCase();
      const delegated = payload?.delegated === true || meta.delegated === true;
      const incomingStart = typeof payload?.startTimestamp === "number" ? payload.startTimestamp : null;
      const metaStart = typeof meta.startTimestamp === "number" ? meta.startTimestamp : null;

      if (DEBUG_ENABLED){
        L("calendar lobby update payload", {
          token: shortToken(token),
          delegate: delegateTarget ? shortId(delegateTarget, 20) : "",
          delegated,
          startTimestamp: incomingStart,
          metaStart
        });
      }

      if (delegateTarget && currentUser && delegateTarget !== currentUser){
        if (delegated){
          L("calendar lobby update skipped (delegate mismatch)", {
            token: shortToken(token),
            delegate: delegateIdRaw || meta.delegateId || "",
            currentUser: currentUserRaw || ""
          });
          return { ok:false, skipped:true, reason:"delegateMismatch" };
        }
        L("calendar lobby update by owner before delegation", {
          token: shortToken(token),
          delegate: delegateIdRaw || meta.delegateId || "",
          currentUser: currentUserRaw || ""
        });
      }

      if (meta.lobbyEnabled === false){
        return { ok:false, skipped:true, reason:"lobbyDisabled" };
      }

      const startTs = incomingStart ?? metaStart;
      if (typeof startTs !== "number"){
        return { ok:false, error: bgI18n("error_unknown_utility_request") };
      }
      if (metaStart === startTs){
        L("calendar lobby update skipped (unchanged start)", {
          token: shortToken(token),
          startTimestamp: startTs
        });
        return { ok:true, skipped:true, reason:"startUnchanged" };
      }

      L("calendar lobby update apply", { token: shortToken(token), startTimestamp: startTs });
      await updateTalkLobby({
        token,
        enableLobby: true,
        startTimestamp: startTs
      });
      await setRoomMeta(token, {
        lobbyEnabled: true,
        startTimestamp: startTs
      });
      L("calendar lobby update success", { token: shortToken(token), startTimestamp: startTs });
      return { ok:true };
    },
    calendarDelegateModerator: async () => {
      const token = payload?.token;
      const delegateId = payload?.delegateId;
      if (!token || !delegateId){
        return { ok:false, error: bgI18n("error_delegation_data_missing") };
      }
      const { user } = await NCCore.getOpts();
      const targetNorm = String(delegateId).trim().toLowerCase();
      const currentNorm = (user || "").trim().toLowerCase();
      if (targetNorm === currentNorm){
        L("calendar delegation skipped (same user)", { token: shortToken(token), delegate: delegateId });
        return { ok:false, skipped:true, reason:"sameUser" };
      }
      const result = await delegateRoomModerator({ token, newModerator: delegateId });
      await setRoomMeta(token, {
        delegated: true,
        delegateId,
        delegateName: payload?.delegateName || delegateId
      });
      return { ok:true, result };
    },
    getConfig: async () => {
      const config = await NCCore.getOpts();
      return { ok:true, config };
    },
    decodeAvatar: async () => {
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
    },
    supportsEventConversation: async () => {
      const info = await getEventConversationSupport();
      return { ok:true, supported: info.supported, reason: info.reason || "" };
    },
    delegateModerator: async () => {
      const result = await delegateRoomModerator(payload || {});
      return { ok:true, result };
    },
    deleteRoom: async () => {
      await deleteTalkRoom(payload || {});
      await deleteRoomMeta(payload?.token);
      return { ok:true };
    },
    searchUsers: async () => {
      const users = await searchSystemAddressbook(payload || {});
      return { ok:true, users };
    },
    openDialog: async () => {
      try{
        const windowId = payload?.windowId ?? null;
        await openTalkDialogWindow(windowId);
        return { ok:true };
      }catch(e){
        return { ok:false, error: e?.message || String(e) };
      }
    }
  };

  try{
    const handler = handlers[type];
    if (!handler){
      return { ok:false, error: bgI18n("error_unknown_utility_request") };
    }
    return await handler();
  }catch(e){
    return { ok:false, error: e?.message || String(e) };
  }
});

// *** IMPORTANT: initialize experiment on startup ***
(async () => {
  try{
    const defaultLabel = bgI18n("ui_insert_button_label");
    const defaultTooltip = bgI18n("ui_toolbar_tooltip");
    const ok = await browser.calToolbar.init({ label: defaultLabel, tooltip: defaultTooltip });
    if (!ok){
      console.warn("[NCBG] experiment init returned falsy value");
    }
  }catch(e){
    console.error("[NCBG] init error", e);
  }
})();




























  
