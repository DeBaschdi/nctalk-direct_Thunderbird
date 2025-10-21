'use strict';
/**
 * Hintergrundskript für Nextcloud Talk Direkt.
 * Verantwortlich für API-Aufrufe (Talk + CardDAV), Caching und
 * Utilities, die der Frontend-Teil per Messaging abruft.
 */
function L(...a){ console.log("[NCBG]", ...a); }

/**
 * Lädt die in storage abgelegten Verbindungsdaten und bereitet sie auf.
 * @returns {Promise<{baseUrl:string,user:string,appPass:string}>}
 */
async function getOpts(){
  const stored = await browser.storage.local.get([
    "baseUrl",
    "user",
    "appPass"
  ]);
  return {
    baseUrl: stored.baseUrl ? String(stored.baseUrl).replace(/\/+$/,"") : "",
    user: stored.user || "",
    appPass: stored.appPass || ""
  };
}

/**
 * Cache für System-Adressbuch-Einträge, um CardDAV-Traffic zu begrenzen.
 * Structure:
 * {
 *   contacts: Array<Contact>,
 *   fetchedAt: number,
 *   baseUrl: string,
 *   user: string
 * }
 */
const SYSTEM_ADDRESSBOOK_CACHE = {
  contacts: [],
  fetchedAt: 0,
  baseUrl: "",
  user: ""
};
const SYSTEM_ADDRESSBOOK_TTL = 5 * 60 * 1000;

/**
 * Entfernt vCard-Escaping (z.B. \n, \\, \;) aus einem Feld.
 * @param {string} value
 * @returns {string}
 */
function decodeVCardValue(value){
  return String(value ?? "")
    .replace(/\\\\/g, "\\")
    .replace(/\\n/gi, "\n")
    .replace(/\\,/g, ",")
    .replace(/\\;/g, ";")
    .replace(/\\:/g, ":");
}

/**
 * Klappt gefaltete vCard-Zeilen (RFC 6350) wieder auf.
 * @param {string} data
 * @returns {string[]}
 */
function unfoldVCardLines(data){
  const normalized = String(data ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  const unfolded = [];
  for (const line of lines){
    if (!line.length && !unfolded.length){
      continue;
    }
    if ((line.startsWith(" ") || line.startsWith("\t")) && unfolded.length){
      unfolded[unfolded.length - 1] += line.slice(1);
    } else {
      unfolded.push(line);
    }
  }
  return unfolded;
}

/**
 * Wandelt das exportierte System-Adressbuch in ein internes Format um.
 * @param {string} data - Rohes vCard-Dokument.
 * @returns {Array<{id:string,label:string,email:string,idLower:string,labelLower:string,emailLower:string}>}
 */
function parseSystemAddressbook(data){
  const unfolded = unfoldVCardLines(data);
  const contacts = [];
  let card = null;
  for (const rawLine of unfolded){
    const line = rawLine.trimEnd();
    if (!line) continue;
    if (line.toUpperCase() === "BEGIN:VCARD"){
      card = { emails: [] };
      continue;
    }
    if (!card) continue;
    if (line.toUpperCase() === "END:VCARD"){
      if (card.uid && card.emails.length){
        const preferred = card.emails.find((item) => (item.params["X-NC-SCOPE"] || "").toLowerCase() === "v2-federated") ||
          card.emails[0];
        if (preferred && preferred.value){
          const label = card.fn || card.nickname || card.displayName || card.uid;
          const idLower = card.uid.toLowerCase();
          const labelLower = label ? label.toLowerCase() : "";
          const emailLower = preferred.value.toLowerCase();
          contacts.push({
            id: card.uid,
            label,
            email: preferred.value,
            idLower,
            labelLower,
            emailLower
          });
        }
      }
      card = null;
      continue;
    }
    const colonIdx = line.indexOf(":");
    if (colonIdx === -1) continue;
    const lhs = line.slice(0, colonIdx);
    const rhs = line.slice(colonIdx + 1);
    const segments = lhs.split(";");
    const tag = (segments.shift() || "").toUpperCase();
    const params = {};
    for (const segment of segments){
      if (!segment) continue;
      const eqIdx = segment.indexOf("=");
      if (eqIdx === -1){
        params[segment.toUpperCase()] = true;
      } else {
        const key = segment.slice(0, eqIdx).toUpperCase();
        const value = decodeVCardValue(segment.slice(eqIdx + 1));
        params[key] = value;
      }
    }
    const decoded = decodeVCardValue(rhs);
    if (tag === "UID"){
      card.uid = decoded.trim();
    } else if (tag === "FN"){
      card.fn = decoded.trim();
    } else if (tag === "N" && !card.fn){
      const parts = decoded.split(";").filter(Boolean);
      if (parts.length){
        card.fn = parts.join(" ").trim();
      }
    } else if (tag === "NICKNAME" && !card.nickname){
      card.nickname = decoded.trim();
    } else if (tag === "EMAIL"){
      const email = decoded.trim();
      if (email){
        card.emails.push({ value: email, params });
      }
    } else if (tag === "X-ABSHOWAS"){
      card.displayName = decoded.trim();
    }
  }
  contacts.sort((a, b) => {
    const labelA = a.labelLower || a.emailLower || a.idLower;
    const labelB = b.labelLower || b.emailLower || b.idLower;
    const cmp = labelA.localeCompare(labelB, undefined, { sensitivity: "base" });
    if (cmp !== 0) return cmp;
    return a.idLower.localeCompare(b.idLower);
  });
  return contacts;
}

/**
 * Holt das System-Adressbuch aus Nextcloud (oder den Cache).
 * @param {boolean} [force=false] - erzwingt frischen Abruf.
 * @returns {Promise<Array<{id:string,label:string,email:string,idLower:string,labelLower:string,emailLower:string}>>}
 */
async function getSystemAddressbookContacts(force = false){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const now = Date.now();
  if (!force &&
      SYSTEM_ADDRESSBOOK_CACHE.contacts.length &&
      SYSTEM_ADDRESSBOOK_CACHE.user === user &&
      SYSTEM_ADDRESSBOOK_CACHE.baseUrl === baseUrl &&
      now - SYSTEM_ADDRESSBOOK_CACHE.fetchedAt < SYSTEM_ADDRESSBOOK_TTL){
    return SYSTEM_ADDRESSBOOK_CACHE.contacts;
  }
  const auth = "Basic " + btoa(user + ":" + appPass);
  const base = baseUrl.replace(/\/$/,"");
  // Zugriff auf das serverseitige System-Adressbuch (CardDAV) – erfordert remote.php-Berechtigung.
  const addressUrl = base + "/remote.php/dav/addressbooks/users/" + encodeURIComponent(user) + "/z-server-generated--system/?export";
  const res = await fetch(addressUrl, {
    method: "GET",
    headers: {
      "Authorization": auth,
      "Accept": "text/directory",
      "Cache-Control": "no-cache"
    }
  });
  if (!res.ok){
    const text = await res.text().catch(() => "");
    throw new Error("System-Adressbuch konnte nicht geladen werden: " + (text || (res.status + " " + res.statusText)));
  }
  const raw = await res.text();
  const contacts = parseSystemAddressbook(raw);
  SYSTEM_ADDRESSBOOK_CACHE.contacts = contacts;
  SYSTEM_ADDRESSBOOK_CACHE.fetchedAt = now;
  SYSTEM_ADDRESSBOOK_CACHE.user = user;
  SYSTEM_ADDRESSBOOK_CACHE.baseUrl = baseUrl;
  return contacts;
}

/**
 * Filtert Kontakte aus dem System-Adressbuch anhand eines Suchbegriffs.
 * @param {{searchTerm?:string, limit?:number, forceRefresh?:boolean}} [param0]
 * @returns {Promise<Array<{id:string,label:string,email:string}>>}
 */
async function searchSystemAddressbook({ searchTerm = "", limit = 200, forceRefresh = false } = {}){
  const contacts = await getSystemAddressbookContacts(forceRefresh);
  const term = String(searchTerm || "").trim().toLowerCase();
  let filtered = contacts;
  if (term){
    filtered = contacts.filter((entry) => {
      return entry.idLower.includes(term) ||
        (entry.labelLower && entry.labelLower.includes(term)) ||
        entry.emailLower.includes(term);
    });
  }
  const limited = typeof limit === "number" && limit > 0 ? filtered.slice(0, limit) : filtered;
  return limited.map(({ id, label, email }) => ({ id, label, email }));
}

function randToken(len=10){ const a="abcdefghijklmnopqrstuvwxyz0123456789"; let s=""; for(let i=0;i<len;i++) s+=a[Math.floor(Math.random()*a.length)]; return s; }

function sanitizeDescription(desc){
  if (!desc) return "";
  return String(desc).trim();
}

function buildRoomDescription(baseDescription, url, password){
  const parts = [];
  if (baseDescription && String(baseDescription).trim()) parts.push(String(baseDescription).trim());
  if (url) parts.push("Talk-Link: " + url);
  if (password) parts.push("Passwort: " + password);
  return parts.join("\n\n").trim();
}

async function createTalkPublicRoom({ title, password, enableLobby, enableListable, description, startTimestamp } = {}){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");

  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json", "Content-Type":"application/json" };

  const base = baseUrl.replace(/\/$/,"");
  const createUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room";
  const ROOM_TYPE_PUBLIC = 3;
  const LISTABLE_NONE = 0;
  const LISTABLE_USERS = 1;
  const listableScope = enableListable ? LISTABLE_USERS : LISTABLE_NONE;
  const body = {
    roomType: ROOM_TYPE_PUBLIC,
    type: ROOM_TYPE_PUBLIC,
    roomName: title || "Besprechung",
    listable: listableScope,
    participants: {}
  };
  if (password) body.password = password;
  const cleanedDescription = sanitizeDescription(description);
  if (cleanedDescription) body.description = cleanedDescription;
  try{
    const res = await fetch(createUrl, { method:"POST", headers, body: JSON.stringify(body) });
    const raw = await res.text();
    let data = null;
    try { data = raw ? JSON.parse(raw) : null; } catch(_){}
    if (!res.ok) {
      const meta = data?.ocs?.meta || {};
      const payload = data?.ocs?.data || {};
      const parts = [];
      if (meta.message && meta.message !== meta.status) parts.push(meta.message);
      if (payload.error) parts.push(payload.error);
      if (Array.isArray(payload.errors)) parts.push(...payload.errors);
      if (meta.statuscode) parts.push("Statuscode " + meta.statuscode);
      if (res.status) parts.push("HTTP " + res.status + " " + res.statusText);
      const detail = parts.filter(Boolean).join(" / ") || raw || (res.status + " " + res.statusText);
      const err = new Error("OCS-Fehler: " + detail);
      err.fatal = true;
      err.status = res.status;
      err.response = raw;
      err.meta = meta;
      err.payload = payload;
      throw err;
    }
    let token = data?.ocs?.data?.token || data?.ocs?.data?.roomToken || data?.token || data?.data?.token;
    if (!token) throw new Error("Kein Token im OCS-Response.");
    const url = base + "/call/" + token;
    if (enableLobby){
      try{
        const lobbyUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + token + "/webinar/lobby";
        const payload = { state: 1 };
        if (typeof startTimestamp === "number" && Number.isFinite(startTimestamp) && startTimestamp > 0){
          let timerVal = startTimestamp;
          if (timerVal > 1e12) timerVal = Math.floor(timerVal / 1000);
          payload.timer = Math.floor(timerVal);
        }
        L("set lobby payload", payload);
        const lobbyRes = await fetch(lobbyUrl, { method:"PUT", headers, body: JSON.stringify(payload) });
        if (!lobbyRes.ok) L("lobby set failed status", lobbyRes.status);
      }catch(e){ L("lobby set failed", e); }
    }
    if (enableListable){
      try{
        const listableUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + token + "/listable";
        const listableRes = await fetch(listableUrl, { method:"PUT", headers, body: JSON.stringify({ scope: LISTABLE_USERS }) });
        if (!listableRes.ok) L("listable set failed status", listableRes.status);
      }catch(e){ L("listable set failed", e); }
    }
    const finalDescription = buildRoomDescription(description, url, password);
    if (finalDescription && finalDescription !== cleanedDescription){
      try{
        const descUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + token + "/description";
        L("set description payload");
        const descRes = await fetch(descUrl, { method:"PUT", headers, body: JSON.stringify({ description: finalDescription }) });
        if (!descRes.ok) L("description set failed status", descRes.status);
      }catch(e){ L("description set failed", e); }
    }
    return { url, token, fallback: false, description: finalDescription || cleanedDescription || "" };
  }catch(e){
    if (e && e.fatal) {
      L("create via OCS fatal:", e.message, e.status);
      if (e.response) L("OCS response:", e.response);
      throw e;
    }
    L("create via OCS failed, fallback to pseudo url:", e?.message);
    const token = randToken(10);
    return { url: base + "/call/" + token, token, fallback: true, reason: e?.message || String(e) };
  }
}

async function updateTalkLobby({ token, enableLobby, startTimestamp } = {}){
  if (!token) throw new Error("Raum-Token fehlt.");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json", "Content-Type":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const lobbyUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + token + "/webinar/lobby";
  const payload = { state: enableLobby ? 1 : 0 };
  if (enableLobby && typeof startTimestamp === "number" && Number.isFinite(startTimestamp) && startTimestamp > 0){
    let timerVal = startTimestamp;
    if (timerVal > 1e12) timerVal = Math.floor(timerVal / 1000);
    payload.timer = Math.floor(timerVal);
  }
  if (!enableLobby) delete payload.timer;
  L("update lobby payload", payload);
  const res = await fetch(lobbyUrl, { method:"PUT", headers, body: JSON.stringify(payload) });
  if (!res.ok) {
    if (res.status === 403) {
      throw new Error("Keine Berechtigung zum Aendern der Lobby (HTTP 403).");
    }
    throw new Error("Lobby-Update fehlgeschlagen (HTTP " + res.status + ").");
  }
  return true;
}

function isTruthy(value){
  if (value === true || value === 1) return true;
  if (typeof value === "string"){
    const lower = value.toLowerCase();
    return ["true","1","yes","open","public","all","guests","guest","world"].includes(lower);
  }
  return false;
}

function isFalsy(value){
  if (value === false || value === 0) return true;
  if (typeof value === "string"){
    const lower = value.toLowerCase();
    return ["false","0","no","closed","private","none","off","invite","invite_only","restricted"].includes(lower);
  }
  return false;
}

async function listTalkUserRooms({ searchTerm = "" } = {}){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const url = base + "/ocs/v2.php/apps/spreed/api/v4/room?includeStatus=true&includeLastMessage=false";
  const res = await fetch(url, { method:"GET", headers });
  const raw = await res.text();
  let data = null;
  try { data = raw ? JSON.parse(raw) : null; } catch(_){}
  if (!res.ok){
    const meta = data?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw new Error("OCS-Fehler: " + detail);
  }
  const rooms = Array.isArray(data?.ocs?.data) ? data.ocs.data : [];
  const term = searchTerm ? searchTerm.toLowerCase() : null;
  const out = [];
  for (const room of rooms){
    const normalized = normalizeRoomEntry(room, base, "own");
    if (!normalized) continue;
    if (term){
      const hay = [normalized.displayName, normalized.name, normalized.description].filter(Boolean).join(" ").toLowerCase();
      if (!hay.includes(term)) continue;
    }
    normalized.source = normalized.source || "own";
    out.push(normalized);
  }
  return out;
}

function roomAvatarUrl(base, token, version, theme = "dark"){
  if (!base || !token) return null;
  const query = version ? "?v=" + encodeURIComponent(version) : "";
  return base.replace(/\/$/,"") + "/ocs/v2.php/apps/spreed/api/v1/room/" + encodeURIComponent(token) + "/avatar/" + theme + query;
}

/**
 * Normalisiert Talk-Raumdaten (egal ob öffentlich gelistet oder Eigenbestand).
 * Filtert gleichzeitig Räume ohne Gastzugang heraus.
 * @param {object} room - API-Response-Objekt.
 * @param {string} [base=""] - Basis-URL für Avatar-Link.
 * @param {string} [source=""] - Quelle (listed|own).
 * @returns {object|null}
 */
function normalizeRoomEntry(room, base = "", source = ""){
  if (!room || typeof room !== "object") return null;
  const token = room.token || room.roomToken || room.id || room.identifier;
  if (!token) return null;
  const type = room.type ?? room.roomType ?? room.kind ?? null;
  const config = room.config || {};
  const listable = room.listable === 1 || room.listable === true;
  const accessRaw = config.access ?? config.accessLevel ?? room.access ?? room.accessLevel ?? null;
  const access = typeof accessRaw === "string" ? accessRaw.toLowerCase() : accessRaw;
  const guestsFlag =
    config.allowGuests ??
    config.allow_guests ??
    config.guestAccess ??
    config.guest_access ??
    room.allowGuests ??
    room.allow_guests ??
    room.guestAccess ??
    room.guest_access ??
    null;
  let guestsAllowed = null;
  if (guestsFlag != null){
    if (isTruthy(guestsFlag)) guestsAllowed = true;
    else if (isFalsy(guestsFlag)) guestsAllowed = false;
  }
  if (guestsAllowed == null && access){
    if (isTruthy(access)) guestsAllowed = true;
    else if (isFalsy(access)) guestsAllowed = false;
  }
  if (guestsAllowed == null){
    if (room.guestURL || room.link || room.publicUrl || room.guestUrl) guestsAllowed = true;
  }
  if (guestsAllowed == null && typeof type === "number" && type === 3){
    guestsAllowed = true;
  }
  if (guestsAllowed == null && source){
    if (source === "own" || source === "listed-own") guestsAllowed = true;
  }
  if (guestsAllowed == null && listable) guestsAllowed = true;
  if (guestsAllowed == null) guestsAllowed = false;
  const allowedType = typeof type === "number" ? (type === 3 || type === 2) : true;
  if (!allowedType && !listable && !guestsAllowed) return null;
  const normalized = Object.assign({}, room);
  normalized.token = token;
  if (!normalized.displayName && normalized.name) normalized.displayName = normalized.name;
  normalized.listable = listable ? 1 : 0;
  normalized.access = access || normalized.access || null;
  normalized.guestsAllowed = !!guestsAllowed;
  normalized.source = source || normalized.source || null;
  normalized.isParticipant =
    room.isParticipant === true ||
    typeof room.participantType === "number" ||
    (Array.isArray(room.participants) && room.participants.length > 0) ||
    normalized.source === "own";
  if (normalized.avatar_version && !normalized.avatarVersion) normalized.avatarVersion = normalized.avatar_version;
  if (base){
    normalized._avatarUrl = roomAvatarUrl(base, token, normalized.avatarVersion);
  }
  return normalized;
}

/**
 * Liefert eine kombinierte Liste aller öffentlichen bzw. freigegebenen Talk-Räume.
 * Verbindet serverweite Listen mit eigenen Räumen und filtert doppelte Einträge.
 * @param {{searchTerm?:string}} [param0]
 * @returns {Promise<object[]>}
 */
async function listTalkPublicRooms({ searchTerm = "" } = {}){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const listedUrl = base + "/ocs/v2.php/apps/spreed/api/v4/listed-room?searchTerm=" + encodeURIComponent(searchTerm || "");
  const [listedRes, personalRooms] = await Promise.all([
    fetch(listedUrl, { method:"GET", headers }),
    listTalkUserRooms({ searchTerm })
  ]);
  const rawListed = await listedRes.text();
  let listedData = null;
  try { listedData = rawListed ? JSON.parse(rawListed) : null; } catch(_){}
  if (!listedRes.ok){
    const meta = listedData?.ocs?.meta || {};
    const detail = meta.message || rawListed || (listedRes.status + " " + listedRes.statusText);
    throw new Error("OCS-Fehler: " + detail);
  }
  const listedRooms = Array.isArray(listedData?.ocs?.data) ? listedData.ocs.data : [];
  const map = new Map();
  const push = (room, source) => {
    const normalized = normalizeRoomEntry(room, base, source);
    if (!normalized) return;
    const token = normalized.token || normalized.roomToken;
    if (!token) return;
    if (searchTerm){
      const term = searchTerm.toLowerCase();
      const hay = [normalized.displayName, normalized.name, normalized.description].filter(Boolean).join(" ").toLowerCase();
      if (!hay.includes(term)) return;
    }
    normalized.source = normalized.source || source || "unknown";
    map.set(token, normalized);
  };
  for (const room of listedRooms){
    push(room, "listed");
  }
  for (const room of personalRooms){
    push(room, "own");
  }
  const filtered = Array.from(map.values()).filter((room) => {
    const joinable = room.listable || room.source === "own" || room.isParticipant;
    return joinable && room.guestsAllowed !== false;
  });
  return filtered;
}

async function getTalkRoomInfo({ token } = {}){
  if (!token) throw new Error("Raum-Token fehlt.");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const infoUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token);
  const res = await fetch(infoUrl, { method:"GET", headers });
  const raw = await res.text();
  let data = null;
  try { data = raw ? JSON.parse(raw) : null; } catch(_){}
  if (!res.ok){
    const meta = data?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw new Error("OCS-Fehler: " + detail);
  }
  const room = data?.ocs?.data;
  if (!room || typeof room !== "object"){
    throw new Error("Raumdetails fehlen im OCS-Response.");
  }
  return room;
}

/**
 * Ruft Teilnehmer eines Talk-Raums ab. 404 wird als "keine Daten" interpretiert.
 * @param {{token:string}} param0
 * @returns {Promise<object[]>}
 */
async function getTalkRoomParticipants({ token } = {}){
  if (!token) throw new Error("Raum-Token fehlt.");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = { "OCS-APIRequest": "true", "Authorization": auth, "Accept":"application/json" };
  const base = baseUrl.replace(/\/$/,"");
  const infoUrl = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/participants?includeStatus=true";
  const res = await fetch(infoUrl, { method:"GET", headers });
  const raw = await res.text();
  let data = null;
  try { data = raw ? JSON.parse(raw) : null; } catch(_){}
  if (res.status === 404){
    return [];
  }
  if (!res.ok){
    const meta = data?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw new Error("OCS-Fehler: " + detail);
  }
  const participants = data?.ocs?.data;
  return Array.isArray(participants) ? participants : [];
}

async function setTalkRoomAvatar({ token, data, mime } = {}){
  if (!token || !data) throw new Error("Avatar-Daten fehlen.");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const finalMime = mime || "image/png";
  const base = baseUrl.replace(/\/$/,"");
  const dataUrl = "data:" + finalMime + ";base64," + data;

  const uploadWithFormData = async () => {
    const headers = {
      "OCS-APIRequest": "true",
      "Authorization": auth,
      "Accept": "application/json",
      "X-Requested-With": "XMLHttpRequest"
    };
    const form = new FormData();
    const binary = Uint8Array.from(atob(data), c => c.charCodeAt(0));
    const blob = new Blob([binary], { type: finalMime });
    const filename = finalMime === "image/jpeg" ? "avatar.jpg" : "avatar.png";
    form.append("file", blob, filename);
    const url = base + "/ocs/v2.php/apps/spreed/api/v1/room/" + encodeURIComponent(token) + "/avatar";
    const res = await fetch(url, { method:"POST", headers, body: form });
    const raw = await res.text().catch(() => "");
    let json = null;
    try { json = raw ? JSON.parse(raw) : null; } catch(_){}
    if (!res.ok){
      const meta = json?.ocs?.meta || {};
      const detail = meta.message || raw || (res.status + " " + res.statusText);
      throw new Error(detail || "Avatar-Upload fehlgeschlagen.");
    }
    return {
      avatarVersion: json?.ocs?.data?.avatarVersion || json?.ocs?.data?.version || null,
      base64: data,
      mime: finalMime
    };
  };

  const uploadWithJson = async () => {
    const headers = {
      "OCS-APIRequest": "true",
      "Authorization": auth,
      "Accept": "application/json",
      "Content-Type": "application/json",
      "X-Requested-With": "XMLHttpRequest"
    };
    const url = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/avatar";
    const body = { image: data, mimetype: finalMime };
    const res = await fetch(url, { method:"POST", headers, body: JSON.stringify(body) });
    const raw = await res.text().catch(() => "");
    let json = null;
    try { json = raw ? JSON.parse(raw) : null; } catch(_){}
    if (!res.ok){
      const meta = json?.ocs?.meta || {};
      const detail = meta.message || raw || (res.status + " " + res.statusText);
      throw new Error(detail || "Avatar-Upload fehlgeschlagen.");
    }
    return {
      avatarVersion: json?.ocs?.data?.avatarVersion || json?.ocs?.data?.version || null,
      base64: data,
      mime: finalMime
    };
  };

  try{
    return await uploadWithFormData();
  }catch(err){
    L("Avatar upload via v1 failed, fallback to v4", err?.message || err);
    return await uploadWithJson();
  }
}

async function getTalkRoomAvatar({ token, version, theme = "dark" } = {}){
  if (!token) throw new Error("Raum-Token fehlt.");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": auth,
    "Accept": "image/png,image/jpeg,*/*"
  };
  const base = baseUrl.replace(/\/$/,"");
  let url = base + "/ocs/v2.php/apps/spreed/api/v1/room/" + encodeURIComponent(token) + "/avatar/" + theme;
  if (version){
    url += "?v=" + encodeURIComponent(version);
  }
  const res = await fetch(url, { method:"GET", headers });
  if (!res.ok){
    const raw = await res.text().catch(() => "");
    const detail = raw || (res.status + " " + res.statusText);
    throw new Error("Avatar konnte nicht geladen werden: " + detail);
  }
  const blob = await res.blob();
  const buffer = await blob.arrayBuffer();
  let binary = "";
  const bytes = new Uint8Array(buffer);
  for (const b of bytes){
    binary += String.fromCharCode(b);
  }
  const base64 = btoa(binary);
  const mime = res.headers.get("Content-Type") || blob.type || "image/png";
  return { base64, mime };
}

async function addTalkParticipant({ token, actorId, source = "users" } = {}){
  if (!token || !actorId) throw new Error("Raum-Token oder Teilnehmer-ID fehlt.");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": auth,
    "Accept": "application/json",
    "Content-Type": "application/json"
  };
  const base = baseUrl.replace(/\/$/,"");
  const url = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/participants";
  const body = { newParticipant: actorId, source: source || "users" };
  const res = await fetch(url, { method:"POST", headers, body: JSON.stringify(body) });
  const raw = await res.text().catch(() => "");
  let json = null;
  try { json = raw ? JSON.parse(raw) : null; } catch(_){}
  if (!res.ok && res.status !== 409){
    const meta = json?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw new Error("Teilnehmer konnte nicht hinzugef\u00fcgt werden: " + detail);
  }
  const added = json?.ocs?.data;
  return added || null;
}

async function promoteTalkModerator({ token, attendeeId } = {}){
  if (!token || typeof attendeeId !== "number") throw new Error("Moderator-ID fehlt.");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": auth,
    "Accept": "application/json",
    "Content-Type": "application/json"
  };
  const base = baseUrl.replace(/\/$/,"");
  const url = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/moderators";
  const body = { attendeeId };
  const res = await fetch(url, { method:"POST", headers, body: JSON.stringify(body) });
  if (!res.ok && res.status !== 409){
    const raw = await res.text().catch(() => "");
    let json = null;
    try { json = raw ? JSON.parse(raw) : null; } catch(_){}
    const meta = json?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw new Error("Moderator konnte nicht gesetzt werden: " + detail);
  }
  return true;
}

async function leaveTalkRoom({ token } = {}){
  if (!token) throw new Error("Raum-Token fehlt.");
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": auth,
    "Accept": "application/json"
  };
  const base = baseUrl.replace(/\/$/,"");
  const url = base + "/ocs/v2.php/apps/spreed/api/v4/room/" + encodeURIComponent(token) + "/participants/self";
  const res = await fetch(url, { method:"DELETE", headers });
  if (!res.ok && res.status !== 404){
    const raw = await res.text().catch(() => "");
    let json = null;
    try { json = raw ? JSON.parse(raw) : null; } catch(_){}
    const meta = json?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw new Error("Verlassen des Raums fehlgeschlagen: " + detail);
  }
  return true;
}

async function delegateRoomModerator({ token, newModerator } = {}){
  if (!token || !newModerator) throw new Error("Delegationsdaten fehlen.");
  const { user } = await getOpts();
  const targetId = String(newModerator).trim();
  if (!targetId) throw new Error("Moderator-ID ist leer.");
  const currentUser = (user || "").trim();
  await addTalkParticipant({ token, actorId: targetId, source: "users" });
  const participants = await getTalkRoomParticipants({ token });
  const match = participants.find((p) => {
    if (!p) return false;
    const actor = (p.actorId || "").trim().toLowerCase();
    return actor === targetId.toLowerCase();
  });
  if (!match){
    throw new Error("Teilnehmer wurde nicht gefunden, bitte Schreibweise pr\u00fcfen.");
  }
  await promoteTalkModerator({ token, attendeeId: match.attendeeId });
  const loweredUser = currentUser.toLowerCase();
  if (loweredUser && loweredUser !== targetId.toLowerCase()){
    await leaveTalkRoom({ token });
    return { leftSelf: true, delegate: targetId };
  }
  return { leftSelf: false, delegate: targetId };
}

async function searchSharees({ searchTerm = "", limit = 50 } = {}){
  const { baseUrl, user, appPass } = await getOpts();
  if (!baseUrl || !user || !appPass) throw new Error("Nextcloud Zugangsdaten fehlen (Add-on-Optionen).");
  const auth = "Basic " + btoa(user + ":" + appPass);
  const headers = {
    "OCS-APIRequest": "true",
    "Authorization": auth,
    "Accept": "application/json",
    "X-Requested-With": "XMLHttpRequest"
  };
  const base = baseUrl.replace(/\/$/,"");
  const params = new URLSearchParams({
    format: "json",
    search: searchTerm,
    perPage: String(Math.max(1, Math.min(limit, 200))),
    itemType: "call",
    lookup: "false"
  });
  const url = base + "/ocs/v2.php/apps/files_sharing/api/v1/sharees?" + params.toString();
  const res = await fetch(url, { method:"GET", headers });
  const raw = await res.text().catch(() => "");
  let json = null;
  try { json = raw ? JSON.parse(raw) : null; } catch(_){}
  if (!res.ok){
    const meta = json?.ocs?.meta || {};
    const detail = meta.message || raw || (res.status + " " + res.statusText);
    throw new Error(detail || "Benutzer konnten nicht geladen werden.");
  }
  const exactUsers = Array.isArray(json?.ocs?.data?.exact?.users) ? json.ocs.data.exact.users : [];
  const regularUsers = Array.isArray(json?.ocs?.data?.users) ? json.ocs.data.users : [];
  const combined = [...exactUsers, ...regularUsers];
  const emailPattern = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/;
  const seen = new Set();
  const suggestions = [];
  for (const entry of combined){
    const shareWith = entry?.value?.shareWith || entry?.value?.share_with;
    if (!shareWith) continue;
    const id = String(shareWith).trim();
    if (!id || seen.has(id)) continue;
    seen.add(id);
    const additionalInfo =
      entry?.value?.shareWithAdditionalInfo ??
      entry?.value?.share_with_additional_info ??
      entry?.value?.additionalInfo ??
      entry?.value?.additional_info ??
      null;
    const candidates = [
      entry?.label,
      additionalInfo,
      entry?.value?.email,
      entry?.value?.mail,
      entry?.value?.mailAddress,
      entry?.value?.mail_address,
      entry?.value?.emailAddress,
      entry?.value?.email_address,
      entry?.value?.EMail
    ];
    let email = null;
    for (const candidate of candidates){
      if (typeof candidate !== "string") continue;
      const match = candidate.match(emailPattern);
      if (match && match[0]){
        email = match[0].trim().toLowerCase();
        break;
      }
    }
    if (!email) continue;
    suggestions.push({
      id,
      label: entry?.label || id,
      email,
      shareType: entry?.value?.shareType ?? entry?.value?.share_type ?? null
    });
  }
  return suggestions;
}

browser.runtime.onMessage.addListener(async (msg) => {
  if (!msg || !msg.type) return;
  L("msg", msg.type);
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
      const config = await getOpts();
      return { ok:true, config };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:chooseExisting"){
    return { ok:true };
  }
  if (msg.type === "talkMenu:updateLobby"){
    try{
      await updateTalkLobby(msg.payload);
      return { ok:true };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:listPublicRooms"){
    try{
      const rooms = await listTalkPublicRooms(msg.payload || {});
      return { ok:true, rooms };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:getRoomInfo"){
    try{
      const room = await getTalkRoomInfo(msg.payload || {});
      return { ok:true, room };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:getRoomParticipants"){
    try{
      const participants = await getTalkRoomParticipants(msg.payload || {});
      return { ok:true, participants };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:setRoomAvatar"){
    try{
      const result = await setTalkRoomAvatar(msg.payload || {});
      return { ok:true, avatarVersion: result?.avatarVersion || null, dataUrl: result?.dataUrl || null, result };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
  if (msg.type === "talkMenu:getRoomAvatar"){
    try{
      const result = await getTalkRoomAvatar(msg.payload || {});
      return { ok:true, base64: result.base64, mime: result.mime };
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
  if (msg.type === "talkMenu:searchUsers"){
    try{
      const users = await searchSystemAddressbook(msg.payload || {});
      return { ok:true, users };
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }
});

browser.calToolbar.onCreateRequest.addListener(async (payload = {}) => {
  try {
    const out = await createTalkPublicRoom(payload);
    return { ok:true, url: out.url, token: out.token, fallback: !!out.fallback, reason: out.reason };
  } catch(e){
    return { ok:false, error: e?.message || String(e) };
  }
});

browser.calToolbar.onLobbyUpdate.addListener(async (payload = {}) => {
  try{
    await updateTalkLobby(payload);
    return { ok:true };
  }catch(e){
    return { ok:false, error: e?.message || String(e) };
  }
});

browser.calToolbar.onUtilityRequest.addListener(async (payload = {}) => {
  try{
    const type = payload?.type;
    if (type === "getConfig"){
      const config = await getOpts();
      return { ok:true, config };
    }
    if (type === "listPublicRooms"){
      const rooms = await listTalkPublicRooms({ searchTerm: payload?.searchTerm || "" });
      return { ok:true, rooms };
    }
    if (type === "getRoomInfo"){
      const room = await getTalkRoomInfo({ token: payload?.token });
      return { ok:true, room };
    }
    if (type === "getRoomParticipants"){
      const participants = await getTalkRoomParticipants({ token: payload?.token });
      return { ok:true, participants };
    }
    if (type === "setRoomAvatar"){
      const result = await setTalkRoomAvatar(payload || {});
      return { ok:true, avatarVersion: result?.avatarVersion || null, dataUrl: result?.dataUrl || null, result };
    }
    if (type === "getRoomAvatar"){
      const result = await getTalkRoomAvatar(payload || {});
      return { ok:true, base64: result.base64, mime: result.mime };
    }
    if (type === "delegateModerator"){
      const result = await delegateRoomModerator(payload || {});
      return { ok:true, result };
    }
    if (type === "searchUsers"){
      const users = await searchSystemAddressbook(payload || {});
      return { ok:true, users };
    }
    return { ok:false, error: "Unbekannte Utility-Anfrage." };
  }catch(e){
    return { ok:false, error: e?.message || String(e) };
  }
});

// *** IMPORTANT: initialize experiment on startup ***
(async () => {
  try{
    const ok = await browser.calToolbar.init({ label: "Talk-Link einf\u00fcgen", tooltip: "Nextcloud Talk" });
    L("experiment init result:", ok);
  }catch(e){
    console.error("[NCBG] init error", e);
  }
})();

