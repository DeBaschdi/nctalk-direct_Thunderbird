/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
'use strict';

/**
 * Kernfunktionen rund um Login-Flow V2 und Credential-Checks.
 * Läuft im gleichen Kontext wie talkcore/background.
 */
const NCCore = (() => {
  const DEVICE_NAME = "Nextcloud Enterprise for Thunderbird";

  /**
   * Normalisiert eine Basis-URL (Trim + Trailing Slashes entfernen).
   * @param {string} input - Benutzerangabe aus den Einstellungen.
   * @returns {string} - Bereinigte Basis-URL oder leerer String.
   */
  function normalizeBaseUrl(input){
    if (!input) return "";
    return String(input).trim().replace(/\/+$/, "");
  }

  /**
   * Prüft Basis-URL, Benutzername und App-Passwort per OCS-API.
   * Ruft zuerst /cloud/capabilities und danach /cloud/user auf.
   * @param {{baseUrl:string,user:string,appPass:string}} params
   * @returns {Promise<{ok:boolean, code:string, message?:string, version?:string}>}
   */
  async function testCredentials({ baseUrl, user, appPass } = {}){
    const normalizedBase = normalizeBaseUrl(baseUrl);
    const trimmedUser = typeof user === "string" ? user.trim() : "";
    const password = typeof appPass === "string" ? appPass : "";
    if (!normalizedBase || !trimmedUser || !password){
      return { ok:false, code:"missing", message: bgI18n("error_credentials_missing") };
    }
    const basicHeader = "Basic " + btoa(trimmedUser + ":" + password);
    try{
      L("options test connection", { base: normalizedBase, user: shortId(trimmedUser) });
      const headers = {
        "OCS-APIRequest": "true",
        "Authorization": basicHeader,
        "Accept": "application/json"
      };
      const url = normalizedBase + "/ocs/v2.php/cloud/capabilities";
      const res = await fetch(url, { method:"GET", headers });
      const raw = await res.text().catch(() => "");
      let data = null;
      try{ data = raw ? JSON.parse(raw) : null; }catch(_){}
      if (res.status === 401 || res.status === 403){
        const detail = data?.ocs?.meta?.message || "HTTP " + res.status;
        return { ok:false, code:"auth", message: detail };
      }
      if (!res.ok){
        const detail = data?.ocs?.meta?.message || raw || (res.status + " " + res.statusText);
        return { ok:false, code:"http", message: detail };
      }
      const versionRaw = data?.ocs?.meta?.version || data?.ocs?.data?.version || "";
      let versionStr = "";
      if (typeof versionRaw === "string"){
        versionStr = versionRaw;
      } else if (versionRaw && typeof versionRaw === "object"){
        if (typeof versionRaw.string === "string" && versionRaw.string.trim()){
          versionStr = versionRaw.string.trim();
        } else {
          const parts = [];
          if (versionRaw.major != null) parts.push(String(versionRaw.major));
          if (versionRaw.minor != null) parts.push(String(versionRaw.minor));
          if (versionRaw.micro != null) parts.push(String(versionRaw.micro));
          if (parts.length){
            versionStr = parts.join(".");
          }
        }
      }
      const message = versionStr ? "Nextcloud " + versionStr : "";
      try{
        const userUrl = normalizedBase + "/ocs/v2.php/cloud/user";
        const userRes = await fetch(userUrl, {
          method: "GET",
          headers: {
            "OCS-APIRequest": "true",
            "Authorization": basicHeader,
            "Accept": "application/json"
          }
        });
        if (userRes.status === 401 || userRes.status === 403){
          return { ok:false, code:"auth", message: bgI18n("options_test_failed_auth") };
        }
        if (!userRes.ok){
          const userRaw = await userRes.text().catch(() => "");
          const userData = (() => { try{ return userRaw ? JSON.parse(userRaw) : null; }catch(_){ return null; }})();
          const detail = userData?.ocs?.meta?.message || userRaw || (userRes.status + " " + userRes.statusText);
          return { ok:false, code:"http", message: detail };
        }
      }catch(userErr){
        return { ok:false, code:"network", message: userErr?.message || String(userErr) };
      }
      return { ok:true, version: versionStr, message };
    }catch(e){
      return { ok:false, code:"network", message: e?.message || String(e) };
    }
  }

  /**
   * Startet den Login Flow V2 und liefert Browser-URL + Poll-Endpunkt.
   * @param {string} baseUrl - Nextcloud-Instanz (bereits geprüft).
   * @returns {Promise<{loginUrl:string,pollEndpoint:string,pollToken:string}>}
   */
  async function startLoginFlow(baseUrl){
    const normalized = normalizeBaseUrl(baseUrl);
    if (!normalized){
      throw new Error(bgI18n("error_credentials_missing"));
    }
    const url = normalized + "/index.php/login/v2";
    const headers = {
      "OCS-APIRequest": "true",
      "Accept": "application/json",
      "Content-Type": "application/json"
    };
    const body = JSON.stringify({ name: DEVICE_NAME });
    const res = await fetch(url, { method:"POST", headers, body });
    const raw = await res.text().catch(() => "");
    let data = null;
    try{ data = raw ? JSON.parse(raw) : null; }catch(_){}
    if (!res.ok){
      const detail = data?.ocs?.meta?.message || raw || (res.status + " " + res.statusText);
      throw new Error(detail || bgI18n("options_loginflow_failed"));
    }
    const loginUrl = data?.login;
    const poll = data?.poll || {};
    let pollEndpoint = poll.endpoint || "";
    const pollToken = poll.token || "";
    if (!loginUrl || !pollEndpoint || !pollToken){
      throw new Error("Login-Flow Antwort unvollständig.");
    }
    if (!/^https?:/i.test(pollEndpoint)){
      pollEndpoint = normalized + pollEndpoint;
    }
    return {
      loginUrl,
      pollEndpoint,
      pollToken
    };
  }

  /**
   * Pollt den Login Flow Endpunkt bis ein App-Passwort geliefert wird.
   * @param {{pollEndpoint:string,pollToken:string,timeoutMs?:number,intervalMs?:number}} options
   * @returns {Promise<{loginName:string,appPassword:string}>}
   */
  async function completeLoginFlow({ pollEndpoint, pollToken, timeoutMs = 120000, intervalMs = 2000 } = {}){
    if (!pollEndpoint || !pollToken){
      throw new Error("Login-Flow Daten fehlen.");
    }
    const headers = {
      "OCS-APIRequest": "true",
      "Accept": "application/json",
      "Content-Type": "application/json"
    };
    const deadline = Date.now() + (timeoutMs > 0 ? timeoutMs : 120000);
    while (Date.now() < deadline){
      try{
        const payload = JSON.stringify({ token: pollToken, deviceName: DEVICE_NAME });
        const res = await fetch(pollEndpoint, { method:"POST", headers, body: payload });
        if (res.status === 404){
          await delay(intervalMs);
          continue;
        }
        const raw = await res.text().catch(() => "");
        let data = null;
        try{ data = raw ? JSON.parse(raw) : null; }catch(_){}
        if (!res.ok){
          const detail = data?.ocs?.meta?.message || raw || (res.status + " " + res.statusText);
          throw new Error(detail || bgI18n("options_loginflow_failed"));
        }
        const appPassword = data?.appPassword || data?.token || data?.ocs?.data?.appPassword || data?.ocs?.data?.token;
        const loginName = data?.loginName || data?.ocs?.data?.loginName;
        if (!appPassword || !loginName){
          throw new Error("Login-Flow liefert kein App-Passwort.");
        }
        return {
          loginName,
          appPassword
        };
      }catch(err){
        if (err?.message && /login flow/i.test(err.message)){
          throw err;
        }
        if (err && err.statusCode === 404){
          await delay(intervalMs);
          continue;
        }
        throw err;
      }
    }
    throw new Error("Login-Flow Timeout.");
  }

  /**
   * Liefert ein Promise, das nach ms Millisekunden erfüllt wird.
   * @param {number} ms - Wartezeit.
   * @returns {Promise<void>}
   */
  function delay(ms){
    return new Promise((resolve) => setTimeout(resolve, Math.max(ms || 0, 50)));
  }

  /**
   * Gibt die gespeicherten Zugangsdaten inkl. normalisierter Basis-URL zurück.
   * @returns {Promise<{baseUrl:string,user:string,appPass:string,debugEnabled:boolean,authMode:string}>}
   */
  async function getOpts(){
    if (typeof browser === "undefined" || !browser?.storage?.local){
      return {
        baseUrl: "",
        user: "",
        appPass: "",
        debugEnabled: false,
        authMode: "manual"
      };
    }
    const stored = await browser.storage.local.get([
      "baseUrl",
      "user",
      "appPass",
      "debugEnabled",
      "authMode"
    ]);
    return {
      baseUrl: normalizeBaseUrl(stored.baseUrl || ""),
      user: typeof stored.user === "string" ? stored.user.trim() : "",
      appPass: typeof stored.appPass === "string" ? stored.appPass : "",
      debugEnabled: !!stored.debugEnabled,
      authMode: stored.authMode || "manual"
    };
  }

  return {
    normalizeBaseUrl,
    testCredentials,
    startLoginFlow,
    completeLoginFlow,
    getOpts
  };
})();
