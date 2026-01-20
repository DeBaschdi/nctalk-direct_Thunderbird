/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
'use strict';
const i18n = NCI18n.translate;
const DEFAULT_FILELINK_EXPIRE_DAYS = 7;
const DEFAULT_FILELINK_SHARE_NAME = i18n("filelink_share_default") || "Freigabename";
const DEFAULT_TALK_TITLE = i18n("ui_default_title") || "Besprechung";

NCTalkDomI18n.translatePage(i18n, { titleKey: "options_title" });
initTabs();
initAbout();

const statusEl = document.getElementById("status");
const baseUrlInput = document.getElementById("baseUrl");
const userInput = document.getElementById("user");
const appPassInput = document.getElementById("appPass");
const fileLinkBaseInput = document.getElementById("fileLinkBase");
const filelinkDefaultShareNameInput = document.getElementById("filelinkDefaultShareName");
const filelinkDefaultPermCreateInput = document.getElementById("filelinkDefaultPermCreate");
const filelinkDefaultPermWriteInput = document.getElementById("filelinkDefaultPermWrite");
const filelinkDefaultPermDeleteInput = document.getElementById("filelinkDefaultPermDelete");
const filelinkDefaultPasswordInput = document.getElementById("filelinkDefaultPassword");
const filelinkDefaultExpireDaysInput = document.getElementById("filelinkDefaultExpireDays");
const talkDefaultTitleInput = document.getElementById("talkDefaultTitle");
const talkDefaultLobbyInput = document.getElementById("talkDefaultLobby");
const talkDefaultListableInput = document.getElementById("talkDefaultListable");
const talkDefaultRoomTypeRadios = Array.from(document.querySelectorAll("input[name='talkDefaultRoomType']"));
const DEFAULT_FILELINK_BASE = (typeof NCFileLink !== "undefined" ? NCFileLink.DEFAULT_BASE_PATH : "90 Freigaben - extern");
let statusTimer = null;

function showStatus(message, isError = false, sticky = false, isSuccess = false){
  if (statusTimer){
    clearTimeout(statusTimer);
    statusTimer = null;
  }
  statusEl.textContent = message || "";
  statusEl.style.color = isError ? "#b00020" : (isSuccess ? "#11883a" : "");
  if (message && !isError && !sticky){
    statusTimer = setTimeout(() => {
      statusEl.textContent = "";
      statusTimer = null;
    }, 2000);
  }
}

async function load(){
  const stored = await browser.storage.local.get([
    "baseUrl",
    "user",
    "appPass",
    "debugEnabled",
    "authMode",
    "fileLinkBasePath",
    "filelinkDefaultShareName",
    "filelinkDefaultPermCreate",
    "filelinkDefaultPermWrite",
    "filelinkDefaultPermDelete",
    "filelinkDefaultPassword",
    "filelinkDefaultExpireDays",
    "talkDefaultTitle",
    "talkDefaultLobby",
    "talkDefaultListable",
    "talkDefaultRoomType"
  ]);
  if (stored.baseUrl) baseUrlInput.value = stored.baseUrl;
  if (stored.user) userInput.value = stored.user;
  if (stored.appPass) appPassInput.value = stored.appPass;
  document.getElementById("debugEnabled").checked = !!stored.debugEnabled;
  if (fileLinkBaseInput){
    fileLinkBaseInput.value = stored.fileLinkBasePath || DEFAULT_FILELINK_BASE;
  }
  if (filelinkDefaultShareNameInput){
    filelinkDefaultShareNameInput.value = stored.filelinkDefaultShareName || DEFAULT_FILELINK_SHARE_NAME;
  }
  if (filelinkDefaultPermCreateInput){
    filelinkDefaultPermCreateInput.checked = !!stored.filelinkDefaultPermCreate;
  }
  if (filelinkDefaultPermWriteInput){
    filelinkDefaultPermWriteInput.checked = !!stored.filelinkDefaultPermWrite;
  }
  if (filelinkDefaultPermDeleteInput){
    filelinkDefaultPermDeleteInput.checked = !!stored.filelinkDefaultPermDelete;
  }
  if (filelinkDefaultPasswordInput){
    filelinkDefaultPasswordInput.checked = stored.filelinkDefaultPassword !== undefined
      ? !!stored.filelinkDefaultPassword
      : true;
  }
  if (filelinkDefaultExpireDaysInput){
  filelinkDefaultExpireDaysInput.value = String(NCTalkTextUtils.normalizeExpireDays(stored.filelinkDefaultExpireDays, DEFAULT_FILELINK_EXPIRE_DAYS));
  }
  if (talkDefaultTitleInput){
    talkDefaultTitleInput.value = stored.talkDefaultTitle || DEFAULT_TALK_TITLE;
  }
  if (talkDefaultLobbyInput){
    talkDefaultLobbyInput.checked = stored.talkDefaultLobby !== undefined
      ? !!stored.talkDefaultLobby
      : true;
  }
  if (talkDefaultListableInput){
    talkDefaultListableInput.checked = stored.talkDefaultListable !== undefined
      ? !!stored.talkDefaultListable
      : true;
  }
  setTalkDefaultRoomType(stored.talkDefaultRoomType);
  setAuthMode(stored.authMode || "manual");
  updateAuthModeUI();
}

async function save(){
  const baseUrl = baseUrlInput.value.trim();
  const user = userInput.value.trim();
  const appPass = appPassInput.value;
  const debugEnabled = document.getElementById("debugEnabled").checked;
  const authMode = getSelectedAuthMode();
  const fileLinkBasePath = (fileLinkBaseInput?.value?.trim()) || DEFAULT_FILELINK_BASE;
  const filelinkDefaultShareName = (filelinkDefaultShareNameInput?.value || "").trim() || DEFAULT_FILELINK_SHARE_NAME;
  const filelinkDefaultPermCreate = !!filelinkDefaultPermCreateInput?.checked;
  const filelinkDefaultPermWrite = !!filelinkDefaultPermWriteInput?.checked;
  const filelinkDefaultPermDelete = !!filelinkDefaultPermDeleteInput?.checked;
  const filelinkDefaultPassword = filelinkDefaultPasswordInput
    ? !!filelinkDefaultPasswordInput.checked
    : true;
  const filelinkDefaultExpireDays = NCTalkTextUtils.normalizeExpireDays(filelinkDefaultExpireDaysInput?.value, DEFAULT_FILELINK_EXPIRE_DAYS);
  const talkDefaultTitle = (talkDefaultTitleInput?.value || "").trim() || DEFAULT_TALK_TITLE;
  const talkDefaultLobby = talkDefaultLobbyInput ? !!talkDefaultLobbyInput.checked : true;
  const talkDefaultListable = talkDefaultListableInput ? !!talkDefaultListableInput.checked : true;
  const talkDefaultRoomType = getSelectedTalkDefaultRoomType();
  await browser.storage.local.set({
    baseUrl,
    user,
    appPass,
    debugEnabled,
    authMode,
    fileLinkBasePath,
    filelinkDefaultShareName,
    filelinkDefaultPermCreate,
    filelinkDefaultPermWrite,
    filelinkDefaultPermDelete,
    filelinkDefaultPassword,
    filelinkDefaultExpireDays,
    talkDefaultTitle,
    talkDefaultLobby,
    talkDefaultListable,
    talkDefaultRoomType
  });
  showStatus(i18n("options_status_saved"));
}

document.getElementById("save").addEventListener("click", async () => {
  try{
    await save();
  }catch(e){
    console.error(e);
    showStatus(e?.message || i18n("options_status_save_failed"), true);
  }
});

const testButton = document.getElementById("testConnection");
if (testButton){
  testButton.addEventListener("click", async () => {
    const button = testButton;
    if (button.disabled) return;
    button.disabled = true;
    const originalLabel = button.textContent;
    button.textContent = i18n("options_test_running");
    try{
      await runConnectionTest({ showMissing: true });
    }catch(err){
      console.error(err);
      showStatus(err?.message || i18n("options_test_failed"), true);
    }finally{
      button.disabled = false;
      button.textContent = originalLabel || i18n("options_test_button");
    }
  });
}

load().catch((e) => {
  console.error(e);
  showStatus(e?.message || i18n("options_status_load_failed"), true);
});

function initTabs(){
  const buttons = Array.from(document.querySelectorAll(".tab-btn"));
  const panels = Array.from(document.querySelectorAll(".tab-panel"));
  const activate = (id) => {
    buttons.forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.tab === id);
    });
    panels.forEach((panel) => {
      panel.classList.toggle("active", panel.id === `tab-${id}`);
    });
  };
  buttons.forEach((btn) => {
    btn.addEventListener("click", () => activate(btn.dataset.tab));
  });
}

function initAbout(){
  const versionEl = document.getElementById("aboutVersion");
  try{
    const manifest = browser?.runtime?.getManifest?.();
    if (manifest?.version && versionEl){
      versionEl.textContent = manifest.version;
    }
  }catch(_){}
  const licenseLink = document.getElementById("licenseLink");
  if (licenseLink && browser?.runtime?.getURL){
    licenseLink.href = browser.runtime.getURL("LICENSE.txt");
  }
}

const authRadios = Array.from(document.querySelectorAll("input[name='authMode']"));
const loginFlowButton = document.getElementById("loginFlowButton");
let loginFlowInProgress = false;

authRadios.forEach((radio) => {
  radio.addEventListener("change", () => {
    updateAuthModeUI();
  });
});

function getSelectedAuthMode(){
  const checked = document.querySelector("input[name='authMode']:checked");
  return checked ? checked.value : "manual";
}

function setAuthMode(mode){
  const target = authRadios.find((radio) => radio.value === mode);
  if (target){
    target.checked = true;
  } else if (authRadios.length){
    authRadios[0].checked = true;
  }
}

function updateAuthModeUI(){
  const mode = getSelectedAuthMode();
  const manual = mode === "manual";
  if (userInput) userInput.disabled = !manual;
  if (appPassInput) appPassInput.disabled = !manual;
  if (loginFlowButton){
    loginFlowButton.disabled = loginFlowInProgress || mode !== "loginFlow";
  }
}

function getSelectedTalkDefaultRoomType(){
  const checked = talkDefaultRoomTypeRadios.find((radio) => radio.checked);
  return checked?.value === "normal" ? "normal" : "event";
}

function setTalkDefaultRoomType(value){
  const normalized = value === "normal" ? "normal" : "event";
  talkDefaultRoomTypeRadios.forEach((radio) => {
    radio.checked = radio.value === normalized;
  });
}

if (loginFlowButton){
  loginFlowButton.addEventListener("click", async () => {
    if (loginFlowButton.disabled || loginFlowInProgress) return;
    const baseUrl = baseUrlInput.value.trim();
    if (!baseUrl){
      showStatus(i18n("options_loginflow_missing"), true);
      return;
    }
    loginFlowInProgress = true;
    updateAuthModeUI();
    try{
      const promise = browser.runtime.sendMessage({
        type: "options:loginFlow",
        payload: { baseUrl }
      });
      showStatus(i18n("options_loginflow_starting"), false, true);
      showStatus(i18n("options_loginflow_browser"), false, true);
      const response = await promise;
      if (response?.ok){
        if (response.user) userInput.value = response.user;
        if (response.appPass) appPassInput.value = response.appPass;
        showStatus(i18n("options_loginflow_success"), false, false, true);
        await runConnectionTest({ showMissing: false });
      }else{
        showStatus(response?.error || i18n("options_loginflow_failed"), true);
      }
    }catch(err){
      console.error(err);
      showStatus(err?.message || i18n("options_loginflow_failed"), true);
    }finally{
      loginFlowInProgress = false;
      updateAuthModeUI();
    }
  });
}

async function runConnectionTest({ showMissing = true } = {}){
  const baseUrl = baseUrlInput.value.trim();
  const user = userInput.value.trim();
  const appPass = appPassInput.value;
  if (!baseUrl || !user || !appPass){
    if (showMissing){
      showStatus(i18n("options_test_missing"), true);
    }
    return { ok:false, skipped:true, reason:"missing" };
  }
  try{
    const response = await browser.runtime.sendMessage({
      type: "options:testConnection",
      payload: { baseUrl, user, appPass }
    });
    if (response?.ok){
      const message = response?.message ? String(response.message) : i18n("options_test_success");
      showStatus(message, false, false, true);
    }else{
      const code = response?.code;
      const fallbackKey = code === "auth" ? "options_test_failed_auth" : "options_test_failed";
      const message = response?.error || i18n(fallbackKey);
      showStatus(message, true);
    }
    return response;
  }catch(err){
    console.error(err);
    showStatus(err?.message || i18n("options_test_failed"), true);
    return { ok:false, error: err?.message || String(err) };
  }
}



