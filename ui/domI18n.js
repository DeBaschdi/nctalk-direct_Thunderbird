/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
/**
 * Apply i18n strings to DOM attributes and the document title.
 */
(function(global){
  "use strict";

  function translatePage(i18nFn, options){
    if (typeof document === "undefined"){
      return;
    }
    const translate = typeof i18nFn === "function" ? i18nFn : null;
    if (!translate){
      return;
    }
    const opts = options || {};
    const titleKey = opts.titleKey || "";
    const titleFallback = opts.titleFallback || "";

    const textNodes = document.querySelectorAll("[data-i18n]");
    textNodes.forEach((el) => {
      const key = el.dataset.i18n;
      if (!key) return;
      const message = translate(key);
      if (message) el.textContent = message;
    });

    const placeholderNodes = document.querySelectorAll("[data-i18n-placeholder]");
    placeholderNodes.forEach((el) => {
      const key = el.dataset.i18nPlaceholder;
      if (!key) return;
      const message = translate(key);
      if (message) el.setAttribute("placeholder", message);
    });

    if (titleKey){
      const titleMessage = translate(titleKey);
      if (titleMessage){
        document.title = titleMessage;
      } else if (titleFallback){
        document.title = titleFallback;
      }
    } else if (titleFallback){
      document.title = titleFallback;
    }
  }

  global.NCTalkDomI18n = { translatePage };
})(typeof window !== "undefined" ? window : (typeof globalThis !== "undefined" ? globalThis : this));
