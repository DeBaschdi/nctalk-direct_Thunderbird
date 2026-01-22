/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
"use strict";
(function(global){
  const cache = new Map();
  const pending = new Map();

  function normalizeLang(input){
    const raw = String(input || "").toLowerCase();
    if (raw.startsWith("de")) return "de";
    if (raw.startsWith("fr")) return "fr";
    return "en";
  }

  function getEffectiveLang(requested){
    if (!requested || requested === "default"){
      const ui = global?.browser?.i18n?.getUILanguage
        ? global.browser.i18n.getUILanguage()
        : "en";
      return normalizeLang(ui);
    }
    return normalizeLang(requested);
  }

  async function loadLocale(lang){
    const normalized = normalizeLang(lang);
    if (cache.has(normalized)){
      return cache.get(normalized);
    }
    if (pending.has(normalized)){
      return pending.get(normalized);
    }
    if (!global?.browser?.runtime?.getURL){
      return {};
    }
    const url = global.browser.runtime.getURL(`_locales/${normalized}/messages.json`);
    const promise = fetch(url)
      .then((res) => (res.ok ? res.json() : {}))
      .catch(() => ({}))
      .then((data) => {
        const value = data && typeof data === "object" ? data : {};
        cache.set(normalized, value);
        pending.delete(normalized);
        return value;
      });
    pending.set(normalized, promise);
    return promise;
  }

  function applySubstitutions(message, substitutions){
    const text = String(message || "");
    if (!substitutions || (Array.isArray(substitutions) && substitutions.length === 0)){
      return text.replace(/\$\$/g, "$");
    }
    const values = Array.isArray(substitutions) ? substitutions : [substitutions];
    let out = text.replace(/\$\$/g, "$");
    values.forEach((value, index) => {
      const token = "$" + (index + 1);
      out = out.split(token).join(String(value ?? ""));
    });
    return out;
  }

  async function tInLang(lang, key, substitutions){
    const effective = getEffectiveLang(lang);
    const data = await loadLocale(effective);
    let message = "";
    if (data && data[key] && typeof data[key].message === "string"){
      message = data[key].message;
    }
    if (!message && global?.browser?.i18n?.getMessage){
      try{
        message = global.browser.i18n.getMessage(key, substitutions);
      }catch(_){
        message = "";
      }
    }
    return applySubstitutions(message, substitutions);
  }

  const api = {
    loadLocale,
    getEffectiveLang,
    tInLang
  };

  if (typeof module !== "undefined" && module.exports){
    module.exports = api;
  }
  global.NCI18nOverride = api;
})(typeof window !== "undefined" ? window : (typeof globalThis !== "undefined" ? globalThis : this));
