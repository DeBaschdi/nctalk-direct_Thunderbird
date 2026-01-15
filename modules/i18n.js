/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
'use strict';
(function(global){
  function translate(key, substitutions){
    try{
      if (global.browser?.i18n?.getMessage){
        const msg = global.browser.i18n.getMessage(key, substitutions);
        if (msg){
          return msg;
        }
      }
    }catch(err){
      try{ console.error('[NC-I18N]', err); }catch(_){ }
    }
    if (Array.isArray(substitutions) && substitutions.length){
      return String(substitutions[0] ?? '');
    }
    return key || '';
  }

  const api = { translate };

  if (typeof module !== 'undefined' && module.exports){
    module.exports = api;
  }
  global.NCI18n = api;
  if (!global.bgI18n){
    global.bgI18n = api.translate;
  }
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
