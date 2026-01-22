/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */
(() => {
  "use strict";

  const i18n = (key, subs) => {
    try{
      return browser.i18n.getMessage(key, subs);
    }catch(_){
      return "";
    }
  };

  NCTalkDomI18n.translatePage(i18n, { titleKey: "open_url_title" });

  const params = new URLSearchParams(window.location.search);
  const url = params.get("url") || "";
  const urlInput = document.getElementById("urlInput");
  const copyBtn = document.getElementById("copyBtn");
  const closeBtn = document.getElementById("closeBtn");
  const statusEl = document.getElementById("status");

  if (urlInput){
    urlInput.value = url;
  }

  function setStatus(key, isError = false){
    if (!statusEl){
      return;
    }
    statusEl.textContent = i18n(key) || "";
    statusEl.style.color = isError ? "#b00020" : "#1f1f1f";
  }

  async function copyUrl(){
    if (!url){
      return;
    }
    try{
      if (navigator.clipboard && navigator.clipboard.writeText){
        await navigator.clipboard.writeText(url);
        setStatus("open_url_copied");
        return;
      }
    }catch(_){}
    try{
      if (urlInput){
        urlInput.focus();
        urlInput.select();
      }
      const ok = document.execCommand("copy");
      if (ok){
        setStatus("open_url_copied");
      }else{
        setStatus("open_url_copy_failed", true);
      }
    }catch(_){
      setStatus("open_url_copy_failed", true);
    }
  }

  copyBtn?.addEventListener("click", copyUrl);
  closeBtn?.addEventListener("click", () => window.close());
})();
