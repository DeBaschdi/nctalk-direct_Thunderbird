/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
"use strict";
(function(global){
  const api = {
    normalizeOriginPattern,
    hasOriginPermission,
    ensureOriginPermissionInteractive
  };

  /**
   * Normalize a user-provided base URL into an origin pattern for optional permissions.
   * @param {string} baseUrl
   * @returns {string} Origin pattern like "https://cloud.example.com/*" or empty string.
   */
  function normalizeOriginPattern(baseUrl){
    if (!baseUrl) return "";
    try{
      const url = new URL(String(baseUrl));
      if (url.protocol !== "https:" && url.protocol !== "http:"){
        return "";
      }
      return url.origin + "/*";
    }catch(_){
      return "";
    }
  }

  /**
   * Check if the extension already has host permission for the given base URL.
   * @param {string} baseUrl
   * @returns {Promise<boolean>}
   */
  async function hasOriginPermission(baseUrl){
    const pattern = normalizeOriginPattern(baseUrl);
    if (!pattern){
      return false;
    }
    if (!global?.browser?.permissions?.contains){
      return true;
    }
    try{
      return await global.browser.permissions.contains({ origins: [pattern] });
    }catch(_){
      return false;
    }
  }

  /**
   * Request optional host permission for the given base URL (interactive by default).
   * @param {string} baseUrl
   * @param {{prompt?:boolean}} options
   * @returns {Promise<boolean>}
   */
  async function ensureOriginPermissionInteractive(baseUrl, options = {}){
    const pattern = normalizeOriginPattern(baseUrl);
    if (!pattern){
      return false;
    }
    if (!global?.browser?.permissions?.request){
      return true;
    }
    const allowPrompt = options?.prompt !== false;
    if (!allowPrompt){
      return await hasOriginPermission(baseUrl);
    }
    try{
      // Request immediately to keep user activation for the prompt.
      const granted = await global.browser.permissions.request({ origins: [pattern] });
      if (granted){
        return true;
      }
      if (global?.browser?.permissions?.contains){
        return await global.browser.permissions.contains({ origins: [pattern] });
      }
      return false;
    }catch(_){
      try{
        if (global?.browser?.permissions?.contains){
          return await global.browser.permissions.contains({ origins: [pattern] });
        }
      }catch(__){}
      return false;
    }
  }

  if (typeof module !== "undefined" && module.exports){
    module.exports = api;
  }
  global.NCHostPermissions = api;
})(typeof window !== "undefined" ? window : (typeof globalThis !== "undefined" ? globalThis : this));
