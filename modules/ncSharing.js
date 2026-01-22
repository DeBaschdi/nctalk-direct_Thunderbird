/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
'use strict';
(function(__context){
  const DEFAULT_BASE_PATH = "90 Freigaben - extern";
  const NEXTCLOUD_DEVICE_NAME = "Nextcloud Enterprise for Thunderbird";
  const PERMISSION_FLAGS = {
    read: 1,
    write: 2,
    create: 4,
    delete: 8
  };
  const INVALID_PATH_CHARS = /[\\/:*?"<>|]/g;
  let cachedLogoBase64 = null;

  function logDebug(opts, ...args){
    if (!opts?.debugEnabled){
      return;
    }
    try{
      console.log("[NCSHARE]", ...args);
    }catch(_){}
  }

  const sharedTranslator = (typeof NCI18n !== "undefined" && typeof NCI18n.translate === "function")
    ? NCI18n.translate
    : null;
  const escapeHtml = NCTalkTextUtils.escapeHtml;

  function i18n(key, substitutions = []){
    if (sharedTranslator){
      try{
        const translated = sharedTranslator(key, substitutions);
        if (translated){
          return translated;
        }
      }catch(_){}
    }
    try{
      if (typeof browser !== "undefined" && browser?.i18n?.getMessage){
        const fallback = browser.i18n.getMessage(key, substitutions);
        if (fallback){
          return fallback;
        }
      }
    }catch(_){}
    if (Array.isArray(substitutions) && substitutions.length){
      return String(substitutions[0] ?? "");
    }
    return key || "";
  }

  function hostPermissionError(){
    return new Error(i18n("error_host_permission_missing"));
  }

  async function ensureHostPermission(baseUrl){
    if (typeof NCHostPermissions === "undefined" || !NCHostPermissions?.hasOriginPermission){
      return true;
    }
    const ok = await NCHostPermissions.hasOriginPermission(baseUrl);
    if (!ok){
      throw hostPermissionError();
    }
    return true;
  }

  async function getShareBlockLang(){
    if (typeof browser === "undefined" || !browser?.storage?.local){
      return "default";
    }
    const stored = await browser.storage.local.get(["shareBlockLang"]);
    return stored.shareBlockLang || "default";
  }

  async function tShare(lang, key, substitutions = []){
    if (typeof NCI18nOverride !== "undefined" && typeof NCI18nOverride.tInLang === "function"){
      const translated = await NCI18nOverride.tInLang(lang, key, substitutions);
      if (translated){
        return translated;
      }
    }
    return i18n(key, substitutions);
  }

  function sanitizeShareName(value){
    const fallback = i18n("sharing_share_default") || "Freigabe";
    if (!value) return fallback;
    const normalized = String(value).normalize("NFKC").replace(INVALID_PATH_CHARS, "_").trim();
    return normalized || fallback;
  }

  function sanitizeFileName(value, fallback = "Datei"){
    if (!value && value !== 0) return fallback;
    const normalized = String(value).normalize("NFKC").replace(INVALID_PATH_CHARS, "_").trim();
    return normalized || fallback;
  }

  function formatDateForFolder(date){
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}${month}${day}`;
  }

  function normalizeRelativePath(path){
    if (!path) return "";
    return String(path).replace(/\\/g, "/").replace(/^\/+/, "").replace(/\/+$/, "");
  }

  function joinRelativePath(base, child){
    const normalizedBase = normalizeRelativePath(base);
    const normalizedChild = normalizeRelativePath(child);
    if (!normalizedBase) return normalizedChild;
    if (!normalizedChild) return normalizedBase;
    return normalizedBase + "/" + normalizedChild;
  }

  function buildShareFolderInfo(basePath, shareName, referenceDate){
    const dateObj = referenceDate instanceof Date ? referenceDate : new Date();
    const folderName = `${formatDateForFolder(dateObj)}_${sanitizeShareName(shareName)}`;
    const relativeBase = normalizeRelativePath(basePath || DEFAULT_BASE_PATH);
    const relativeFolder = joinRelativePath(relativeBase, folderName);
    return {
      date: dateObj,
      folderName,
      relativeBase,
      relativeFolder
    };
  }

  function sanitizeRelativeDir(dir){
    if (!dir) return "";
    return String(dir)
      .split(/[\\/]+/)
      .filter(Boolean)
      .map((segment) => sanitizeFileName(segment, "Ordner"))
      .join("/");
  }

  async function pathExists({ davRoot, relativePath, authHeader }){
    const cleanPath = normalizeRelativePath(relativePath || "");
    if (!cleanPath){
      return false;
    }
    const url = davRoot + "/" + encodePath(cleanPath);
    const res = await fetch(url, {
      method: "PROPFIND",
      headers: {
        "Authorization": authHeader,
        "Depth": "0"
      }
    });
    if (res.status === 404){
      return false;
    }
    if (res.status === 207 || res.status === 200){
      return true;
    }
    if (!res.ok){
      const text = await res.text().catch(() => "");
      throw new Error(text || `Path check failed (${res.status})`);
    }
    return true;
  }

  async function ensureFolderExists(davRoot, relativePath, authHeader){
    const segments = normalizeRelativePath(relativePath).split("/").filter(Boolean);
    let current = "";
    for (const segment of segments){
      current = current ? current + "/" + segment : segment;
      const url = davRoot + "/" + encodePath(current);
      const res = await fetch(url, {
        method: "MKCOL",
        headers: {
          "Authorization": authHeader
        }
      });
      if (res.status === 201 || res.status === 405){
        continue;
      }
      if (!res.ok){
        const text = await res.text().catch(() => "");
        throw new Error(text || `MKCOL failed (${res.status})`);
      }
    }
  }

  async function deleteRemotePath(davRoot, relativePath, authHeader){
    const clean = normalizeRelativePath(relativePath);
    if (!clean){
      return false;
    }
    const url = davRoot + "/" + encodePath(clean);
    const res = await fetch(url, {
      method: "DELETE",
      headers: {
        "Authorization": authHeader
      }
    });
    if (res.status === 404){
      return false;
    }
    if (!res.ok){
      const text = await res.text().catch(() => "");
      throw new Error(text || `DELETE failed (${res.status})`);
    }
    return true;
  }

  function encodePath(path){
    return path.split("/").map((segment) => encodeURIComponent(segment)).join("/");
  }

  async function uploadFile({ davRoot, relativeFolder, fileName, file, authHeader, progressCb, statusCb, displayPath, itemId }){
    const relativePath = joinRelativePath(relativeFolder, fileName);
    const url = davRoot + "/" + encodePath(relativePath);
    if (typeof statusCb === "function"){
      statusCb({ phase: "start", fileName, displayPath, itemId });
    }
    await new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open("PUT", url, true);
      xhr.setRequestHeader("Authorization", authHeader);
      xhr.setRequestHeader("Content-Type", file.type || "application/octet-stream");
      xhr.upload.onprogress = (event) => {
        if (event.lengthComputable){
          const percent = Math.round((event.loaded / event.total) * 100);
          if (typeof statusCb === "function"){
            statusCb({
              phase: "progress",
              fileName,
              displayPath,
              itemId,
              loaded: event.loaded,
              total: event.total,
              percent
            });
          }
        }
      };
      xhr.onerror = () => {
        if (typeof statusCb === "function"){
          statusCb({ phase: "error", fileName, displayPath, itemId, error: "Network error" });
        }
        reject(new Error("Upload failed (network error)"));
      };
      xhr.onload = () => {
        if (xhr.status >= 200 && xhr.status < 300){
          if (typeof progressCb === "function"){
            progressCb(file.name);
          }
          if (typeof statusCb === "function"){
            statusCb({ phase: "done", fileName, displayPath, itemId });
          }
          resolve();
        }else{
          if (typeof statusCb === "function"){
            statusCb({ phase: "error", fileName, displayPath, itemId, error: `Upload failed (${xhr.status})` });
          }
          reject(new Error(`Upload failed (${xhr.status})`));
        }
      };
      xhr.send(file);
    });
  }

  function buildPermissionMask(perms){
    let mask = 0;
    if (perms?.read) mask |= PERMISSION_FLAGS.read;
    if (perms?.write) mask |= PERMISSION_FLAGS.write;
    if (perms?.create) mask |= PERMISSION_FLAGS.create;
    if (perms?.delete) mask |= PERMISSION_FLAGS.delete;
    if (!mask){
      mask = PERMISSION_FLAGS.read;
    }
    return mask;
  }

  async function requestShare(baseUrl, relativeFolder, authHeader, perms, password, expireDate, publicUpload){
    const url = baseUrl.replace(/\/+$/, "") + "/ocs/v2.php/apps/files_sharing/api/v1/shares";
    const params = new URLSearchParams();
    params.append("path", "/" + normalizeRelativePath(relativeFolder));
    params.append("shareType", "3");
    params.append("permissions", String(buildPermissionMask(perms)));
    if (password){
      params.append("password", password);
    }
    if (expireDate){
      params.append("expireDate", expireDate);
    }
    if (publicUpload){
      params.append("publicUpload", "true");
    }
    const response = await NCOcs.ocsRequest({
      url,
      method: "POST",
      headers: {
        "Authorization": authHeader,
        "OCS-APIREQUEST": "true",
        "Accept": "application/json",
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: params
    });
    const raw = response.raw || "";
    const data = response.data;
    if (!response.ok){
      const detail = data?.ocs?.meta?.message || raw || `HTTP ${response.status}`;
      throw new Error(detail);
    }
    return {
      url: data?.ocs?.data?.url || "",
      token: data?.ocs?.data?.token || "",
      id: data?.ocs?.data?.id || ""
    };
  }

  async function updateShareMetadata({ baseUrl, shareId, authHeader, note, permissions, expireDate, password, label }){
    if (!shareId){
      return;
    }
    const url = baseUrl.replace(/\/+$/, "") + `/ocs/v2.php/apps/files_sharing/api/v1/shares/${shareId}`;
    const payload = {
      permissions: String(buildPermissionMask(permissions || {})),
      attributes: "[]",
      note: typeof note === "string" ? note : "",
      expireDate: expireDate || "",
      label: label || "",
      password: password || "",
      hideDownload: "false"
    };
    const response = await NCOcs.ocsRequest({
      url,
      method: "PUT",
      headers: {
        "Authorization": authHeader,
        "OCS-APIREQUEST": "true",
        "Accept": "application/json",
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });
    const raw = response.raw || "";
    const data = response.data;
    if (!response.ok){
      const detail = data?.ocs?.meta?.message || raw || `HTTP ${response.status}`;
      throw new Error(detail);
    }
  }

  async function getLogoBase64(){
    if (cachedLogoBase64){
      return cachedLogoBase64;
    }
    if (typeof browser === "undefined" || !browser?.runtime?.getURL){
      return "";
    }
    const response = await fetch(browser.runtime.getURL("logo-nextcloud-filelink.png"));
    const buffer = await response.arrayBuffer();
    let binary = "";
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++){
      binary += String.fromCharCode(bytes[i]);
    }
    cachedLogoBase64 = btoa(binary);
    return cachedLogoBase64;
  }

  async function checkShareFolderAvailability({ shareName, basePath, shareDate } = {}){
    const opts = await NCCore.getOpts();
    if (!opts.baseUrl || !opts.user || !opts.appPass){
      throw new Error(i18n("error_credentials_missing"));
    }
    await ensureHostPermission(opts.baseUrl);
    const info = buildShareFolderInfo(basePath || await getFileLinkBasePath(), shareName, shareDate ? new Date(shareDate) : new Date());
    const authHeader = NCOcs.buildAuthHeader(opts.user, opts.appPass);
    const davRoot = `${opts.baseUrl.replace(/\/+$/, "")}/remote.php/dav/files/${encodeURIComponent(opts.user)}`;
    logDebug(opts, "availability:check", {
      shareName,
      basePath: basePath || "",
      relativeFolder: info.relativeFolder
    });
    const exists = await pathExists({
      davRoot,
      relativePath: info.relativeFolder,
      authHeader
    });
    logDebug(opts, "availability:result", {
      relativeFolder: info.relativeFolder,
      exists
    });
    return {
      exists,
      folderInfo: info
    };
  }

  async function checkRemotePathExists(input){
    const opts = await NCCore.getOpts();
    if (!opts.baseUrl || !opts.user || !opts.appPass){
      throw new Error(i18n("error_credentials_missing"));
    }
    await ensureHostPermission(opts.baseUrl);
    const relativePath = typeof input === "string" ? input : input?.relativePath || "";
    const authHeader = NCOcs.buildAuthHeader(opts.user, opts.appPass);
    const davRoot = `${opts.baseUrl.replace(/\/+$/, "")}/remote.php/dav/files/${encodeURIComponent(opts.user)}`;
    logDebug(opts, "remotePath:check", { relativePath });
    const exists = await pathExists({ davRoot, relativePath, authHeader });
    logDebug(opts, "remotePath:result", { relativePath, exists });
    return exists;
  }

  async function buildHtmlBlock(result, request){
    const shareLang = await getShareBlockLang();
    const logo = await getLogoBase64();
    const paragraphs = [];
    if (request?.noteEnabled && request?.note){
      paragraphs.push(`<p style="margin:0 0 14px 0;line-height:1.4;">${escapeHtml(request.note)}</p>`);
    }
    const introLine = await tShare(shareLang, "sharing_html_intro_line");
    if (introLine){
      paragraphs.push(`<p style="margin:0 0 14px 0;line-height:1.4;">${escapeHtml(introLine)}<br /></p>`);
    }
    const downloadLink = `<a href="${escapeHtml(result.shareUrl)}" style="color:#0067c0;text-decoration:none;">${escapeHtml(result.shareUrl)}</a>`;
    const permissionLabels = {
      read: await tShare(shareLang, "sharing_permission_read"),
      create: await tShare(shareLang, "sharing_permission_create"),
      write: await tShare(shareLang, "sharing_permission_write"),
      delete: await tShare(shareLang, "sharing_permission_delete")
    };
    const rows = [];
    rows.push(buildTableRow(await tShare(shareLang, "sharing_html_download_label"), downloadLink));
    if (result.password){
      const badge = `<span style="display:inline-block;font-family:'Consolas','Courier New',monospace;padding:2px 6px;border:1px solid #c7c7c7;border-radius:3px;background-color:#f4f4f4;-ms-user-select:all;user-select:all;" ondblclick="try{window.getSelection().selectAllChildren(this);}catch(e){}" onclick="try{window.getSelection().selectAllChildren(this);}catch(e){}">${escapeHtml(result.password)}</span>`;
      rows.push(buildTableRow(await tShare(shareLang, "sharing_html_password_label"), badge));
    }
    if (result.expireDate){
      rows.push(buildTableRow(await tShare(shareLang, "sharing_html_expire_label"), escapeHtml(result.expireDate)));
    }
    rows.push(buildTableRow(await tShare(shareLang, "sharing_html_permissions_label"), buildPermissionsBadges(result.permissions, permissionLabels)));
    const nextcloudAnchor = `<a href="https://nextcloud.com/" style="color:#0067c0;text-decoration:none;">Nextcloud</a>`;
    const footer = (await tShare(shareLang, "sharing_html_footer", [nextcloudAnchor])) || "";
    return `
<div style="font-family:Calibri,'Segoe UI',Arial,sans-serif;font-size:11pt;color:#1f1f1f;margin:16px 0;">
  <table role="presentation" width="640" style="border-collapse:collapse;width:640px;margin:0;background-color:transparent;">
    <tr>
      <td style="padding:0;">
        <table role="presentation" width="640" style="border-collapse:collapse;width:640px;background-color:#0078d4;height:54px;">
          <tr>
            <td style="text-align:center;vertical-align:middle;padding:0;background-color:#0078d4;">
              <img alt="Nextcloud" style="height:30px;width:auto;display:inline-block;" height="30" src="data:image/png;base64,${logo}" />
            </td>
          </tr>
        </table>
        <div style="padding:18px 0 12px 0;">
          ${paragraphs.join("\n")}
          <table style="width:100%;border-collapse:collapse;margin-bottom:10px;">
            ${rows.join("\n")}
          </table>
        </div>
        <div style="padding:10px 0 16px 0;font-size:9pt;font-style:italic;color:#555555;">
          ${footer}
        </div>
      </td>
    </tr>
  </table>
</div>`;
  }

  function buildTableRow(label, valueHtml){
    if (!valueHtml){
      return "";
    }
    return `<tr>
      <th style="text-align:left;width:12ch;vertical-align:top;padding:6px 10px 6px 0;color:#333333;">${escapeHtml(label)}</th>
      <td style="padding:6px 0;max-width:50ch;word-break:break-word;">${valueHtml}</td>
    </tr>`;
  }

  function buildPermissionsBadges(perms, labels = {}){
    const safePerms = perms || {};
    const entries = [
      { label: labels.read || i18n("sharing_permission_read"), enabled: !!safePerms.read },
      { label: labels.create || i18n("sharing_permission_create"), enabled: !!safePerms.create },
      { label: labels.write || i18n("sharing_permission_write"), enabled: !!safePerms.write },
      { label: labels.delete || i18n("sharing_permission_delete"), enabled: !!safePerms.delete }
    ];
    const cells = entries.map((entry) => {
      const color = entry.enabled ? "#0078d4" : "#c62828";
      return `<td style="padding:0 18px 6px 0;">
        <span style="display:inline-flex;align-items:center;">
          <span style="display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border:1px solid ${color};color:${color};font-size:13px;font-weight:700;">
            ${entry.enabled ? "&#10003;" : "&#10007;"}
          </span>
          <span style="padding-left:6px;font-weight:600;">${escapeHtml(entry.label)}</span>
        </span>
      </td>`;
    }).join("");
    return `<table style="border-collapse:collapse;"><tr>${cells}</tr></table>`;
  }

  /**
   * Create a Nextcloud share, upload files, and return HTML output.
   * @param {object} request
   * @returns {Promise<{html:string, shareUrl:string, shareInfo:object}>}
   */
  async function createFileLink(request){
    const opts = await NCCore.getOpts();
    logDebug(opts, "createFileLink:start", {
      shareName: request?.shareName || "",
      files: Array.isArray(request?.files) ? request.files.length : 0
    });
    if (!opts.baseUrl || !opts.user || !opts.appPass){
      throw new Error(i18n("error_credentials_missing"));
    }
    await ensureHostPermission(opts.baseUrl);
    const basePathSetting = request?.basePath && request.basePath.trim()
      ? request.basePath.trim()
      : (await getFileLinkBasePath());
    const shareDate = request?.shareDate ? new Date(request.shareDate) : new Date();
    const folderInfo = request?.folderInfo
      ? request.folderInfo
      : buildShareFolderInfo(basePathSetting, request?.shareName, shareDate);
    const relativeBase = folderInfo.relativeBase;
    const relativeFolder = folderInfo.relativeFolder;
    const authHeader = NCOcs.buildAuthHeader(opts.user, opts.appPass);
    const davRoot = `${opts.baseUrl.replace(/\/+$/, "")}/remote.php/dav/files/${encodeURIComponent(opts.user)}`;
    logDebug(opts, "folders:ensure", { relativeBase, relativeFolder });
    await ensureFolderExists(davRoot, relativeBase, authHeader);
    await ensureFolderExists(davRoot, relativeFolder, authHeader);
    const noteEnabled = !!request?.noteEnabled;
    const noteValue = noteEnabled ? String(request?.note || "").trim() : "";
    request.note = noteValue;
    request.noteEnabled = noteEnabled;
    const normalizedShareName = sanitizeShareName(request?.shareName) || folderInfo.folderName;

    const onProgress = typeof request?.onProgress === "function" ? request.onProgress : null;
    const statusCallback = typeof request?.onUploadStatus === "function" ? request.onUploadStatus : null;
    const files = Array.isArray(request?.files) ? request.files : [];
    if (files.length){
      let uploaded = 0;
      for (const item of files){
        const displayPath = item.displayPath || item.file?.name || "";
        const sanitizedFileName = sanitizeFileName(item.renamedName || item.file?.name || "Datei");
        const relativeDir = sanitizeRelativeDir(item.relativeDir || "");
        const targetFolder = relativeDir ? joinRelativePath(relativeFolder, relativeDir) : relativeFolder;
        if (relativeDir){
          await ensureFolderExists(davRoot, targetFolder, authHeader);
        }
        logDebug(opts, "upload:start", { file: sanitizedFileName, folder: targetFolder });
        await uploadFile({
          davRoot,
          relativeFolder: targetFolder,
          fileName: sanitizedFileName,
          file: item.file,
          authHeader,
          displayPath,
          itemId: item.id,
          statusCb: statusCallback,
          progressCb: () => {
            uploaded++;
            if (onProgress){
              onProgress({ type: "upload", current: uploaded, total: files.length, fileName: displayPath || sanitizedFileName });
            }
          }
        });
        logDebug(opts, "upload:done", { file: sanitizedFileName });
      }
    }

    const share = await requestShare(
      opts.baseUrl,
      relativeFolder,
      authHeader,
      request.permissions,
      request.passwordEnabled ? (request.password || "") : "",
      request.expireEnabled ? (request.expireDate || "") : "",
      !!request.permissions?.create);
    logDebug(opts, "share:created", { url: share.url });
    if (share.id){
      logDebug(opts, "share:updateRequest", {
        shareId: share.id,
        hasNote: !!(noteEnabled && noteValue),
        label: normalizedShareName
      });
      await updateShareMetadata({
        baseUrl: opts.baseUrl,
        shareId: share.id,
        authHeader,
        note: noteEnabled ? noteValue : "",
        permissions: request.permissions,
        expireDate: request.expireEnabled ? (request.expireDate || "") : "",
        password: request.passwordEnabled ? (request.password || "") : "",
        label: normalizedShareName
      });
      logDebug(opts, "share:metadataUpdated", {
        shareId: share.id,
        hasNote: !!(noteEnabled && noteValue),
        label: normalizedShareName
      });
    }

    const resultPayload = {
      shareUrl: share.url,
      password: request.passwordEnabled ? (request.password || "") : "",
      expireDate: request.expireEnabled ? (request.expireDate || "") : "",
      permissions: request.permissions,
      folderInfo,
      note: noteValue,
      noteEnabled,
      shareId: share.id || "",
      label: normalizedShareName
    };
    const html = await buildHtmlBlock(resultPayload, request);
    logDebug(opts, "createFileLink:done", { shareUrl: share.url });

    return {
      html,
      shareUrl: share.url,
      shareInfo: resultPayload
    };
  }

  async function getFileLinkBasePath(){
    if (typeof browser === "undefined" || !browser?.storage?.local){
      return DEFAULT_BASE_PATH;
    }
    const stored = await browser.storage.local.get(["sharingBasePath"]);
    return stored.sharingBasePath || DEFAULT_BASE_PATH;
  }

  /**
   * Update note/label metadata for an existing share (for example after wizard step 4).
   * @param {{shareInfo:Object,noteEnabled:boolean,note:string}} options
   */
  async function updateShareDetails({ shareInfo, noteEnabled, note } = {}){
    if (!shareInfo?.shareId){
      throw new Error(i18n("sharing_error_upload_required"));
    }
    const opts = await NCCore.getOpts();
    if (!opts.baseUrl || !opts.user || !opts.appPass){
      throw new Error(i18n("error_credentials_missing"));
    }
    await ensureHostPermission(opts.baseUrl);
    const authHeader = NCOcs.buildAuthHeader(opts.user, opts.appPass);
    const normalizedLabel = shareInfo.label || sanitizeShareName(shareInfo.folderInfo?.folderName || shareInfo.shareUrl);
    logDebug(opts, "share:updateMeta", {
      shareId: shareInfo.shareId,
      label: normalizedLabel,
      noteEnabled: !!noteEnabled
    });
    await updateShareMetadata({
      baseUrl: opts.baseUrl,
      shareId: shareInfo.shareId,
      authHeader,
      note: noteEnabled ? (note || "") : "",
      permissions: shareInfo.permissions,
      expireDate: shareInfo.expireDate || "",
      password: shareInfo.password || "",
      label: normalizedLabel
    });
    logDebug(opts, "share:updateMeta:done", { shareId: shareInfo.shareId });
  }

  async function deleteShareFolder({ folderInfo } = {}){
    if (!folderInfo?.relativeFolder){
      return false;
    }
    const opts = await NCCore.getOpts();
    if (!opts.baseUrl || !opts.user || !opts.appPass){
      throw new Error(i18n("error_credentials_missing"));
    }
    await ensureHostPermission(opts.baseUrl);
    const authHeader = NCOcs.buildAuthHeader(opts.user, opts.appPass);
    const davRoot = `${opts.baseUrl.replace(/\/+$/, "")}/remote.php/dav/files/${encodeURIComponent(opts.user)}`;
    logDebug(opts, "folders:delete", { relativeFolder: folderInfo.relativeFolder });
    await deleteRemotePath(davRoot, folderInfo.relativeFolder, authHeader);
    return true;
  }

  const api = {
    DEFAULT_BASE_PATH,
    createFileLink,
    buildHtmlBlock,
    getFileLinkBasePath,
    buildShareFolderInfo,
    checkShareFolderAvailability,
    checkRemotePathExists,
    sanitizeShareName,
    sanitizeFileName,
    sanitizeRelativeDir,
    updateShareDetails,
    deleteShareFolder
  };

  if (__context){
    __context.NCSharing = api;
  }
})(typeof window !== "undefined" ? window : (typeof globalThis !== "undefined" ? globalThis : this));

