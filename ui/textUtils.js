/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
/**
 * Text helpers for HTML escaping and expiry normalization.
 */
(function(global){
  "use strict";

  function escapeHtml(value){
    if (value == null) return "";
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function normalizeExpireDays(value, fallbackDays){
    const parsed = parseInt(value, 10);
    if (Number.isFinite(parsed) && parsed > 0){
      return parsed;
    }
    const fallback = parseInt(fallbackDays, 10);
    if (Number.isFinite(fallback) && fallback > 0){
      return fallback;
    }
    return 0;
  }

  global.NCTalkTextUtils = { escapeHtml, normalizeExpireDays };
})(typeof window !== "undefined" ? window : (typeof globalThis !== "undefined" ? globalThis : this));
