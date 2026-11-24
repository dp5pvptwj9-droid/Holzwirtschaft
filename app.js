/* Holzwirtschaft iPhone 1.0.1 – Anmeldung, OneDrive-Zugriff und Oberfläche */
(function () {
  "use strict";
  const CFG = window.HW_CONFIG || {};
  const { Store, AppError, CREW, EXTERNAL, TYPE_TEXT, STATUS_TEXT, LABEL_RECEIVER } = HW;
  const { scaled, unscaled, money, fmtRm, fmtEur, fmtDate, personName, displayName } = HW;
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = "openid offline_access User.Read Files.ReadWrite";
  const LOCK_STALE_SECONDS = 600, HEARTBEAT_MS = 120000, DEVICE = "iPhone";
  const SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024 - 1024;
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  const ls = { get: k => localStorage.getItem("hw." + k), set: (k, v) => localStorage.setItem("hw." + k, v), del: k => localStorage.removeItem("hw." + k) };

  // =====================================================================
  // Anmeldung (OAuth 2.0 mit PKCE, ohne Zusatzbibliothek)
  // =====================================================================
  const Auth = {
    tenant: () => CFG.authority || "consumers",
    redirectUri: () => location.origin + location.pathname.replace(/index\.html$/, ""),
    b64url: buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, ""),
    async login() {
      const verifier = Auth.b64url(crypto.getRandomValues(new Uint8Array(48)));
      const challenge = Auth.b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
      const state = Auth.b64url(crypto.getRandomValues(new Uint8Array(16)));
      ls.set("pkce", JSON.stringify({ verifier, state }));
      const q = new URLSearchParams({ client_id: CFG.clientId, response_type: "code", redirect_uri: Auth.redirectUri(),
        scope: SCOPES, code_challenge: challenge, code_challenge_method: "S256", state, prompt: "select_account" });
      location.assign(`https://login.microsoftonline.com/${Auth.tenant()}/oauth2/v2.0/authorize?${q}`);
    },
    async token(body) {
      const r = await fetch(`https://login.microsoftonline.com/${Auth.tenant()}/oauth2/v2.0/token`, {
        method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({ client_id: CFG.clientId, scope: SCOPES, ...body }) });
      const j = await r.json();
      if (!r.ok) throw new AppError("Anmeldung fehlgeschlagen: " + (j.error_description || j.error || r.status).toString().split("\r\n")[0]);
      ls.set("tok", JSON.stringify({ access: j.access_token, refresh: j.refresh_token, exp: Date.now() + (j.expires_in - 120) * 1000 }));
      return j.access_token;
    },
    async handleRedirect() {
      const p = new URLSearchParams(location.search);
      if (!p.has("code") && !p.has("error")) return;
      history.replaceState(null, "", Auth.redirectUri());
      if (p.has("error")) throw new AppError("Anmeldung abgebrochen: " + (p.get("error_description") || p.get("error")));
      const saved = JSON.parse(ls.get("pkce") || "{}"); ls.del("pkce");
      if (!saved.state || saved.state !== p.get("state")) throw new AppError("Anmeldung ungültig. Bitte erneut anmelden.");
      await Auth.token({ grant_type: "authorization_code", code: p.get("code"), redirect_uri: Auth.redirectUri(), code_verifier: saved.verifier });
    },
    async access() {
      const t = JSON.parse(ls.get("tok") || "null");
      if (!t) return null;
      if (Date.now() < t.exp) return t.access;
      try { return await Auth.token({ grant_type: "refresh_token", refresh_token: t.refresh }); }
      catch (e) { ls.del("tok"); return null; }
    },
    logout() { ls.del("tok"); ls.del("pkce"); }
  };

  // =====================================================================
  // OneDrive (Microsoft Graph)
  // =====================================================================
  const Drive = {
    pathUrl: p => `${GRAPH}/me/drive/root:/${p.split("/").filter(Boolean).map(encodeURIComponent).join("/")}`,
    async call(url, opt = {}, raw = false) {
      const tok = await Auth.access();
      if (!tok) throw new AppError("NOLOGIN");
      let r;
      try { r = await fetch(url, { ...opt, headers: { Authorization: "Bearer " + tok, ...(opt.headers || {}) } }); }
      catch (e) { throw new AppError("Keine Internetverbindung zu OneDrive."); }
      if (raw) return r;
      if (r.status === 401) { Auth.logout(); throw new AppError("NOLOGIN"); }
      if (!r.ok) {
        let msg = ""; try { msg = (await r.json()).error.message; } catch (_) {}
        const e = new AppError(`OneDrive-Fehler ${r.status}${msg ? ": " + msg : ""}`); e.status = r.status; throw e;
      }
      return r.status === 204 ? null : r.json();
    },
    async item(path) {
      try { return await Drive.call(Drive.pathUrl(path)); }
      catch (e) { if (e.status === 404) return null; throw e; }
    },
    async download(item) {
      let r;
      try { r = await fetch(item["@microsoft.graph.downloadUrl"], { cache: "no-store" }); }
      catch (e) { throw new AppError("Keine Internetverbindung zu OneDrive."); }
      if (!r.ok) throw new AppError("Die Datei konnte nicht aus OneDrive geladen werden.");
      return new Uint8Array(await r.arrayBuffer());
    },
    /** Datei ersetzen – nur wenn sie seit dem Laden unverändert ist (eTag). Ergebnis: neues Item. */
    async replace(item, bytes, eTag) {
      const ifMatch = eTag ? { "If-Match": eTag } : {};
      if (bytes.length <= SIMPLE_UPLOAD_MAX) {
        const r = await Drive.call(`${GRAPH}/me/drive/items/${item.id}/content`,
          { method: "PUT", headers: { "Content-Type": "application/octet-stream", ...ifMatch }, body: bytes }, true);
        if (r.status === 412) throw Object.assign(new AppError("CONFLICT"), { conflict: true });
        if (!r.ok) throw new AppError(`Speichern in OneDrive fehlgeschlagen (${r.status}).`);
        return r.json();
      }
      const s = await Drive.call(`${GRAPH}/me/drive/items/${item.id}/createUploadSession`, { method: "POST",
        headers: { "Content-Type": "application/json", ...ifMatch },
        body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }) }, true);
      if (s.status === 412) throw Object.assign(new AppError("CONFLICT"), { conflict: true });
      if (!s.ok) throw new AppError(`Speichern in OneDrive fehlgeschlagen (${s.status}).`);
      const { uploadUrl } = await s.json();
      const CHUNK = 320 * 1024 * 10;
      let last;
      for (let i = 0; i < bytes.length; i += CHUNK) {
        const part = bytes.subarray(i, Math.min(i + CHUNK, bytes.length));
        last = await fetch(uploadUrl, { method: "PUT", body: part, headers: {
          "Content-Range": `bytes ${i}-${i + part.length - 1}/${bytes.length}` } });
        if (!last.ok && last.status !== 202) throw new AppError(`Speichern in OneDrive fehlgeschlagen (${last.status}).`);
      }
      return last.json();
    },
    async putSmall(path, body, type = "application/octet-stream") {
      return Drive.call(Drive.pathUrl(path) + ":/content", { method: "PUT", headers: { "Content-Type": type }, body });
    },
    async remove(itemId) {
      try { await Drive.call(`${GRAPH}/me/drive/items/${itemId}`, { method: "DELETE" }); } catch (_) {}
    }
  };

  // =====================================================================
  // Datenbank im OneDrive: Laden, Sperre, Speichern
  // =====================================================================
  let SQL = null;
  const Cloud = {
    path: () => ls.get("dbPath") || CFG.defaultDbPath || "Holzwirtschaft/holzwirtschaft.db",
    item: null, eTag: null, sqlDb: null, store: null, readOnly: false, lockedBy: "", lockItem: null,
    original: null, lastSync: null, timer: null,

    lockPath: () => Cloud.path() + ".lock",
    async readLock() {
      const it = await Drive.item(Cloud.lockPath());
      if (!it) return null;
      try { return { item: it, data: JSON.parse(new TextDecoder().decode(await Drive.download(it))) }; }
      catch (_) { return { item: it, data: {} }; }
    },
    foreignFresh(lock) {
      if (!lock || !lock.data) return false;
      const d = lock.data;
      return !(d.computer === DEVICE && d.user === DEVICE) && (Date.now() / 1000 - Number(d.heartbeat || 0)) < LOCK_STALE_SECONDS;
    },
    async writeLock() {
      const since = Cloud.lockSince || new Date().toISOString();
      Cloud.lockSince = since;
      Cloud.lockItem = await Drive.putSmall(Cloud.lockPath(), JSON.stringify({ computer: DEVICE, user: DEVICE, since,
        heartbeat: Date.now() / 1000, app: "Holzwirtschaft" }), "application/json");
    },
    async releaseLock() {
      if (Cloud.readOnly || !Cloud.lockItem) return;
      const lock = await Cloud.readLock();
      if (lock && !Cloud.foreignFresh(lock)) await Drive.remove(lock.item.id);
      Cloud.lockItem = null; Cloud.lockSince = null;
    },

    async open(askLocked) {
      if (!SQL) SQL = await initSqlJs({ locateFile: f => `https://cdn.jsdelivr.net/npm/sql.js@1.10.3/dist/${f}` });
      const item = await Drive.item(Cloud.path());
      if (!item || !item.file) throw new AppError(`Die Datenbank wurde im OneDrive nicht gefunden:\n${Cloud.path()}\n\n` +
        "Bitte unter ⚙︎ den Pfad prüfen (relativ zum OneDrive-Hauptordner).");
      const bytes = await Drive.download(item);
      const lock = await Cloud.readLock();
      let readOnly = false;
      if (Cloud.foreignFresh(lock)) {
        const choice = await askLocked(lock.data);
        if (choice === "cancel") throw new AppError("Das Öffnen wurde abgebrochen.");
        readOnly = choice === "readonly";
      }
      Cloud.load(item, bytes);
      Cloud.released = false;
      Cloud.readOnly = readOnly;
      Cloud.lockedBy = readOnly ? (lock.data.computer || "?") : "";
      if (!readOnly) await Cloud.writeLock();
      Cloud.startTimer();
    },
    load(item, bytes) {
      const db = new SQL.Database(bytes);
      db.run("PRAGMA foreign_keys = ON");
      const adapter = {
        all(sql, p = []) {
          const st = db.prepare(sql); const rows = [];
          try { st.bind(p); while (st.step()) rows.push(st.getAsObject()); } finally { st.free(); }
          return rows;
        },
        run(sql, p = []) { db.run(sql, p); return db.getRowsModified(); }
      };
      const store = new Store(adapter);
      store.checkSchema();
      if (Cloud.sqlDb) Cloud.sqlDb.close();
      Object.assign(Cloud, { item, eTag: item.eTag, sqlDb: db, store, original: bytes, lastSync: new Date() });
    },
    async reload() {
      const item = await Drive.item(Cloud.path());
      if (!item) throw new AppError("Die Datenbank ist im OneDrive nicht mehr vorhanden.");
      Cloud.load(item, await Drive.download(item));
    },
    async backupOncePerDay() {
      const key = "backup." + new Date().toISOString().slice(0, 10);
      if (ls.get(key) || !Cloud.original || Cloud.original.length > SIMPLE_UPLOAD_MAX) return;
      const p = Cloud.path().split("/"); const name = p.pop();
      const stem = name.replace(/\.db$/i, "");
      try { await Drive.putSmall([...p, "Sicherungen", `${stem}_iPhone_${new Date().toISOString().slice(0, 10)}.db`].join("/"), Cloud.original); ls.set(key, "1"); }
      catch (_) { /* Sicherung ist optional */ }
    },
    /** Änderung ausführen und sofort im OneDrive speichern. */
    async change(fn) {
      if (Cloud.readOnly) throw new AppError(`Nur ansehen: Die Datenbank wird gerade auf „${Cloud.lockedBy}“ bearbeitet.`);
      const lock = await Cloud.readLock();
      if (Cloud.foreignFresh(lock)) {
        Cloud.readOnly = true; Cloud.lockedBy = lock.data.computer || "?";
        throw new AppError(`Die Datenbank wird inzwischen auf „${Cloud.lockedBy}“ bearbeitet. Das iPhone zeigt sie nur noch an.`);
      }
      await Cloud.backupOncePerDay();
      const snapshot = Cloud.sqlDb.export();           // Stand vor der Änderung (für Rücknahme)
      Cloud.sqlDb.run("PRAGMA foreign_keys = ON");
      const result = fn(Cloud.store);                   // Fachprüfung + Änderung (AppError -> nichts geändert)
      const bytes = Cloud.sqlDb.export();
      Cloud.sqlDb.run("PRAGMA foreign_keys = ON");
      try {
        const it = await Drive.replace(Cloud.item, bytes, Cloud.eTag);
        Object.assign(Cloud, { item: { ...Cloud.item, ...it }, eTag: it.eTag, original: bytes, lastSync: new Date() });
        await Cloud.writeLock();
      } catch (e) {
        Cloud.load(Cloud.item, snapshot);               // Änderung zurücknehmen
        if (e.conflict) {
          await Cloud.reload();
          throw new AppError("Die Datenbank wurde zwischenzeitlich auf einem anderen Gerät geändert. " +
            "Die Daten wurden neu geladen – bitte die Eingabe erneut vornehmen.");
        }
        throw new AppError("Die Änderung konnte nicht im OneDrive gespeichert werden und wurde verworfen.\n\n" + e.message);
      }
      return result;
    },
    /** Regelmäßig: Lebenszeichen der Sperre, fremde Änderungen erkennen. */
    async tick() {
      if (!Cloud.store || document.hidden) return;
      try {
        const lock = await Cloud.readLock();
        if (!Cloud.readOnly && Cloud.foreignFresh(lock)) {
          Cloud.readOnly = true; Cloud.lockedBy = lock.data.computer || "?";
          UI.toast(`Die Datenbank wird jetzt auf „${Cloud.lockedBy}“ bearbeitet – das iPhone zeigt sie nur noch an.`);
        } else if (!Cloud.readOnly) await Cloud.writeLock();
        else if (!Cloud.released && !Cloud.foreignFresh(lock)) { /* PC fertig: Bearbeitung wieder möglich */
          Cloud.readOnly = false; Cloud.lockedBy = ""; await Cloud.writeLock(); UI.toast("Die Datenbank ist wieder frei – Bearbeiten ist möglich.");
        }
        const it = await Drive.item(Cloud.path());
        if (it && it.eTag !== Cloud.eTag) { Cloud.load(it, await Drive.download(it)); UI.toast("Neue Daten aus OneDrive geladen."); }
        Cloud.lastSync = new Date();
      } catch (e) { if (e.message === "NOLOGIN") UI.showLogin(); }
      UI.render();
    },
    startTimer() { clearInterval(Cloud.timer); Cloud.timer = setInterval(Cloud.tick, HEARTBEAT_MS); }
  };

  // =====================================================================
  // Oberfläche
  // =====================================================================
  const UI = {
    tab: ls.get("tab") || "orders", pid: ls.get("period") || null,

    toast(text) { const t = $("toast"); t.textContent = text; t.style.display = "block"; clearTimeout(UI._tt); UI._tt = setTimeout(() => t.style.display = "none", 3500); },
    error(e) { const m = e && e.message || String(e); if (m === "NOLOGIN") return UI.showLogin(); alert(m); },
    busy(on) { document.body.style.opacity = on ? ".6" : ""; document.body.style.pointerEvents = on ? "none" : ""; },

    openSheet(html) { $("sheet").innerHTML = html; $("sheetBg").classList.add("open"); $("sheet").scrollTop = 0; },
    closeSheet() { $("sheetBg").classList.remove("open"); },
    /** Startbild ausblenden und in die App wechseln. */
    hideSplash() {
      const sp = $("splash"); if (!sp || sp.dataset.done) return;
      sp.dataset.done = "1"; sp.classList.add("hide"); setTimeout(() => sp.remove(), 350);
    },

    showLogin(msg = "") {
      $("fab").style.display = "none"; $("period").innerHTML = ""; $("sync").textContent = "Nicht angemeldet";
      const noId = !CFG.clientId;
      $("main").innerHTML = `<div id="start"><img src="icon-180.png" alt=""><h2>Holzwirtschaft</h2>
        ${msg ? `<div class="banner err">${esc(msg)}</div>` : ""}
        ${noId ? `<div class="banner warn">Die App ist noch nicht eingerichtet: In <b>config.js</b> fehlt die Client-ID. ` +
                 `Bitte am PC <b>Holzwirtschaft_iPhone_starten.bat</b> ausführen (siehe Anleitung).</div>` :
          `<p>Mit dem Microsoft-Konto anmelden, in dessen OneDrive die Datenbank liegt.</p>
           <button class="btn blue" id="btnLogin">Mit Microsoft anmelden</button>`}</div>`;
      if (!noId) $("btnLogin").onclick = () => Auth.login().catch(UI.error);
    },

    async start() {
      if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
      try {
        try { await Auth.handleRedirect(); } catch (e) { return UI.showLogin(e.message); }
        if (!CFG.clientId || !(await Auth.access())) return UI.showLogin();
        await UI.connect();
      } finally { UI.hideSplash(); }          // erst wechseln, wenn alles geladen ist (oder eine Meldung kommt)
    },
    async connect() {
      $("main").innerHTML = `<div class="empty">Datenbank wird aus OneDrive geladen …</div>`;
      try {
        await Cloud.open(UI.askLocked);
        UI.render();
      } catch (e) {
        if (e.message === "NOLOGIN") return UI.showLogin();
        $("main").innerHTML = `<div class="banner err">${esc(e.message).replace(/\n/g, "<br>")}</div>
          <button class="btn blue" id="btnRetry">Erneut versuchen</button>`;
        $("btnRetry").onclick = UI.connect;
        $("sync").textContent = "Nicht verbunden";
      }
    },
    askLocked(lock) {
      UI.hideSplash();                        // Rückfrage muss sichtbar sein
      return new Promise(resolve => {
        let since = lock.since || ""; try { since = new Date(since).toLocaleString("de-DE"); } catch (_) {}
        UI.openSheet(`<h2>Datenbank wird bearbeitet</h2>
          <div class="banner warn">Die Datenbank ist gerade auf <b>${esc(lock.computer || "?")}</b>
          (${esc(lock.user || "?")}) geöffnet, seit ${esc(since)}.</div>
          <p class="hint">„Nur ansehen“ ist ohne Risiko. „Trotzdem bearbeiten“ nur wählen, wenn das Programm am PC
          sicher beendet ist – sonst können Änderungen verloren gehen.</p>
          <div class="btns"><button class="btn blue" id="lkRo">Nur ansehen</button></div>
          <div class="btns"><button class="btn edit" id="lkTake">Trotzdem bearbeiten</button>
          <button class="btn sec" id="lkCancel">Abbrechen</button></div>`);
        const done = v => { UI.closeSheet(); resolve(v); };
        $("lkRo").onclick = () => done("readonly"); $("lkTake").onclick = () => done("takeover"); $("lkCancel").onclick = () => done("cancel");
      });
    },

    // ------------------------------------------------------------ Anzeige
    render() {
      const st = Cloud.store; if (!st) return;
      const periods = st.periods();
      if (!periods.find(p => p.id === UI.pid)) UI.pid = periods[0] ? periods[0].id : null;
      $("period").innerHTML = periods.map(p => `<option value="${p.id}" ${p.id === UI.pid ? "selected" : ""}>` +
        `${esc(p.name)} · ${STATUS_TEXT[p.status] || p.status}</option>`).join("") || "<option>Keine Periode vorhanden</option>";
      const t = Cloud.lastSync ? Cloud.lastSync.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" }) : "–";
      $("sync").textContent = Cloud.released ? `Bearbeitung beendet · Stand ${t}` : Cloud.readOnly ? `Nur ansehen – bearbeitet auf „${Cloud.lockedBy}“ · Stand ${t}` : `OneDrive verbunden · Stand ${t}`;
      document.querySelectorAll("nav button").forEach(b => b.classList.toggle("active", b.dataset.tab === UI.tab));
      const editable = UI.canEdit();
      $("fab").style.display = (UI.tab === "customers" ? !Cloud.readOnly : editable) ? "block" : "none";
      let banner = "";
      if (Cloud.released) banner = `<div class="banner info">Bearbeitung beendet, die Sperre ist freigegeben. Zum Bearbeiten oben ⟳ antippen.</div>`;
      else if (Cloud.readOnly) banner = `<div class="banner warn">Nur ansehen: Die Datenbank wird gerade auf „${esc(Cloud.lockedBy)}“ bearbeitet.</div>`;
      else if (UI.pid && !st.editable(UI.pid) && UI.tab !== "customers")
        banner = `<div class="banner info">Die Periode ist abgeschlossen. Es sind keine Änderungen mehr möglich.</div>`;
      if (!UI.pid && UI.tab !== "customers") {
        $("main").innerHTML = banner + `<div class="banner info">Es gibt noch keine Abrechnungsperiode. Bitte in der Windows-App über „Neue Periode“ anlegen.</div>`;
        return;
      }
      $("main").innerHTML = banner + UI["view_" + UI.tab](st);
      $("main").querySelectorAll("[data-id]").forEach(el => el.onclick = () => UI["edit_" + UI.tab](el.dataset.id));
    },
    canEdit() { return Cloud.store && !Cloud.readOnly && UI.pid && Cloud.store.editable(UI.pid); },
    names() { const m = {}; Cloud.store.customers(true).forEach(c => m[c.id] = c); return m; },
    list(rows, fn, sums, emptyText) {
      if (!rows.length) return `<div class="card"><div class="empty">${emptyText}</div></div>`;
      return `<div class="card">${rows.map(fn).join("")}${sums ? `<div class="sum">${sums}</div>` : ""}</div>`;
    },
    sum(rows, key, places) { return unscaled(rows.reduce((s, r) => s + scaled(r[key] || "0", places), 0n), places); },

    view_orders(st) {
      const n = UI.names(), rows = st.orders(UI.pid), chk = st.orderCheck(UI.pid);
      const cls = { NONE: "info", SHORT: "err", SURPLUS: "warn", EXACT: "ok" }[chk.state];
      return `<div class="banner ${cls}">${esc(chk.text)}</div>` + UI.list(rows, r =>
        `<div class="item" data-id="${r.id}"><div class="main"><div class="t">${esc(displayName(n[r.customer_id]))}</div>
         <div class="s">${fmtDate(r.order_date)}</div></div><div class="r"><b>${fmtRm(r.quantity_rm)} RM</b></div></div>`,
        `<span>Summe</span><span>${fmtRm(UI.sum(rows, "quantity_rm", 1))} RM</span>`, "Noch keine Bestellungen.");
    },
    view_purchases(st) {
      const n = UI.names(), rows = st.purchases(UI.pid);
      return UI.list(rows, r => `<div class="item" data-id="${r.id}"><div class="main">
         <div class="t">${fmtRm(r.quantity_rm)} RM à ${fmtEur(r.price_per_rm)}</div>
         <div class="s">${fmtDate(r.purchase_date)} · bezahlt von ${esc(displayName(n[r.paid_by_customer_id]))}</div></div>
         <div class="r"><b>${fmtEur(r.total_price)}</b></div></div>`,
        `<span>${fmtRm(UI.sum(rows, "quantity_rm", 1))} RM</span><span>${fmtEur(UI.sum(rows, "total_price", 2))}</span>`, "Noch keine Holzeinkäufe.");
    },
    view_costs(st) {
      const n = UI.names(), rows = st.costs(UI.pid);
      if (!rows.length) return `<div class="card"><div class="empty">Noch keine Betriebskosten.</div></div>`;
      // Nach Kunde (bezahlt von) gruppieren, alphabetisch nach Name, Vorname; innerhalb nach Datum
      const groups = new Map();
      for (const r of rows) {
        if (!groups.has(r.paid_by_customer_id)) groups.set(r.paid_by_customer_id, []);
        groups.get(r.paid_by_customer_id).push(r);
      }
      const key = cid => { const c = n[cid]; return c ? `${c.last_name} ${c.first_name}`.toLowerCase() : "~"; };
      const order = [...groups.keys()].sort((x, y) => key(x).localeCompare(key(y), "de"));
      return order.map(cid => {
        const list = groups.get(cid);
        return `<div class="group-h"><span>${esc(displayName(n[cid]))}</span>` +
          `<span>${list.length} · ${fmtEur(UI.sum(list, "price", 2))}</span></div>` +
          `<div class="card">${list.map(r => `<div class="item" data-id="${r.id}"><div class="main">
            <div class="t">${esc(r.description)}</div>
            <div class="s">${fmtDate(r.cost_date)} · Menge ${HW.fmtNum(r.quantity)}</div></div>
            <div class="r"><b>${fmtEur(r.price)}</b></div></div>`).join("")}</div>`;
      }).join("") + `<div class="card" style="margin-top:12px"><div class="sum"><span>Summe gesamt</span>` +
        `<span>${fmtEur(UI.sum(rows, "price", 2))}</span></div></div>`;
    },
    view_sales(st) {
      const n = UI.names(), rows = st.sales(UI.pid), types = st.periodCustomerTypes(UI.pid);
      return UI.list(rows, r => {
        const ext = types[r.customer_id] === EXTERNAL;
        const price = r.total_price ? `<b>${fmtEur(r.total_price)}</b><div class="s">${fmtEur(r.price_per_rm)}/RM</div>` : `<span class="s">Preis folgt</span>`;
        return `<div class="item" data-id="${r.id}"><div class="main">
          <div class="t">${esc(displayName(n[r.customer_id]))} <span class="tag ${ext ? "ext" : ""}">${ext ? "extern" : "Crew"}</span></div>
          <div class="s">${fmtDate(r.sale_date)} · ${fmtRm(r.quantity_rm)} RM${ext && r.cash_received_by_customer_id ?
            ` · Einnahme: ${esc(displayName(n[r.cash_received_by_customer_id]))}` : ""}</div></div><div class="r">${price}</div></div>`;
      }, `<span>${fmtRm(UI.sum(rows, "quantity_rm", 1))} RM</span><span>${fmtEur(UI.sum(rows, "total_price", 2))}</span>`, "Noch keine Holzverkäufe.");
    },
    view_customers(st) {
      const rows = st.customers(), types = st.periodCustomerTypes(UI.pid);
      return UI.list(rows, c => {
        const t = types[c.id] || c.customer_type;
        return `<div class="item" data-id="${c.id}"><div class="main"><div class="t">${c.customer_no}. ${esc(personName(c))}</div>
          <div class="s">${esc([c.street, [c.postal_code, c.city].filter(Boolean).join(" ")].filter(Boolean).join(", ")) || "&nbsp;"}</div></div>
          <div class="r"><span class="tag ${t === EXTERNAL ? "ext" : ""}">${TYPE_TEXT[t]}</span></div></div>`;
      }, `<span>Kunden</span><span>${rows.length}</span>`, "Noch keine Kunden.");
    },

    // ------------------------------------------------------------ Formulare
    opts(rows, sel, empty = "– bitte wählen –") {
      return `<option value="">${empty}</option>` + rows.map(c => `<option value="${c.id}" ${c.id === sel ? "selected" : ""}>` +
        `${esc(c.deleted_at ? personName(c) + " (gelöscht)" : personName(c))}</option>`).join("");
    },
    today() {
      const p = Cloud.store.period(UI.pid), d = new Date();
      const iso = new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
      return iso < p.start_date ? p.start_date : iso > p.end_date ? p.end_date : iso;
    },
    de: v => (v === null || v === undefined || v === "") ? "" : String(v).replace(".", ","),
    form(title, fields, row, onSave, onDelete) {
      const ro = row ? (UI.tab === "customers" ? Cloud.readOnly : !UI.canEdit()) : false;
      UI.openSheet(`<h2>${title}</h2>${fields}
        <div class="btns">${ro ? "" : `<button class="btn ${row ? "edit" : ""}" id="fSave">${row ? "Ändern" : "Anlegen"}</button>`}
        <button class="btn sec" id="fCancel">${ro ? "Schließen" : "Abbrechen"}</button></div>
        ${row && onDelete && !ro ? `<div class="btns"><button class="btn del" id="fDel">Löschen</button></div>` : ""}`);
      if (ro) $("sheet").querySelectorAll("input,select").forEach(e => e.disabled = true);
      $("fCancel").onclick = UI.closeSheet;
      const run = async fn => {
        UI.busy(true);
        try { await Cloud.change(fn); UI.closeSheet(); UI.toast("Gespeichert im OneDrive."); }
        catch (e) { UI.error(e); }
        finally { UI.busy(false); UI.render(); }
      };
      if ($("fSave")) $("fSave").onclick = () => run(st => onSave(st, UI.val));
      if ($("fDel")) $("fDel").onclick = () => { if (confirm("Eintrag wirklich löschen?")) run(onDelete); };
    },
    val: id => ($(id) ? $(id).value : ""),
    field(id, label, value = "", type = "text", extra = "") {
      const mode = type === "num" ? `type="text" inputmode="decimal"` : `type="${type}"`;
      return `<label for="${id}">${label}</label><input id="${id}" ${mode} value="${esc(value)}" ${extra}>`;
    },
    select(id, label, options) { return `<label for="${id}">${label}</label><select class="f" id="${id}">${options}</select>`; },
    /** Menge × Preis/RM = Gesamtpreis automatisch ausfüllen. */
    autoTotal(q, p, t) {
      const calc = () => { try {
        const qs = scaled($(q).value, 1), ps = scaled($(p).value, 4);
        if (qs > 0n && $(p).value.trim()) $(t).value = UI.de(money(unscaled(qs * ps, 5)));
      } catch (_) {} };
      const back = () => { try {
        const qs = scaled($(q).value, 1), ts = scaled($(t).value, 2);
        if (qs > 0n && $(t).value.trim()) $(p).value = UI.de(money(unscaled(ts * 1000n / qs, 4)));
      } catch (_) {} };
      $(q).addEventListener("input", calc); $(p).addEventListener("input", calc); $(t).addEventListener("input", back);
    },
    find(rows, id) { const r = rows.find(x => x.id === id); if (!r) throw new AppError("Eintrag nicht gefunden."); return r; },

    edit_orders(id) {
      const st = Cloud.store, r = id ? UI.find(st.orders(UI.pid), id) : null;
      UI.form(r ? "Bestellung" : "Neue Bestellung",
        UI.field("fDate", "Datum", r ? r.order_date : UI.today(), "date") +
        UI.select("fCust", "Kunde*", UI.opts(st.customerChoices(r && r.customer_id), r && r.customer_id)) +
        UI.field("fQty", "Bestellte Menge (RM)", r ? UI.de(r.quantity_rm) : "", "num"), r,
        (s, v) => s.saveOrder(UI.pid, { date: v("fDate"), customer_id: v("fCust"), quantity_rm: v("fQty") }, r && r.id, r && r.version),
        s => s.deleteBooking("customer_order", r.id, r.version));
    },
    edit_purchases(id) {
      const st = Cloud.store, r = id ? UI.find(st.purchases(UI.pid), id) : null;
      UI.form(r ? "Holzeinkauf" : "Neuer Holzeinkauf",
        UI.field("fDate", "Datum", r ? r.purchase_date : UI.today(), "date") +
        UI.field("fQty", "Menge (RM)", r ? UI.de(r.quantity_rm) : "", "num") +
        UI.field("fPrice", "Preis/RM (€)", r ? UI.de(r.price_per_rm) : "", "num") +
        UI.field("fTotal", "Gesamtpreis (€)", r ? UI.de(r.total_price) : "", "num") +
        UI.select("fPaid", "Bezahlt von*", UI.opts(st.crewChoices(UI.pid, r && r.paid_by_customer_id), r && r.paid_by_customer_id)), r,
        (s, v) => s.savePurchase(UI.pid, { date: v("fDate"), quantity_rm: v("fQty"), price_per_rm: v("fPrice"),
          total_price: v("fTotal"), paid_by: v("fPaid") }, r && r.id, r && r.version),
        s => s.deleteBooking("wood_purchase", r.id, r.version));
      UI.autoTotal("fQty", "fPrice", "fTotal");
    },
    edit_costs(id) {
      const st = Cloud.store, r = id ? UI.find(st.costs(UI.pid), id) : null;
      UI.form(r ? "Betriebskosten" : "Neue Betriebskosten",
        UI.field("fDate", "Datum", r ? r.cost_date : UI.today(), "date") +
        UI.field("fQty", "Menge", r ? UI.de(r.quantity) : "1", "num") +
        UI.field("fDesc", "Bezeichnung/Art*", r ? r.description : "") +
        UI.field("fPrice", "Preis (gesamt, €)", r ? UI.de(r.price) : "", "num") +
        UI.select("fPaid", "Bezahlt von*", UI.opts(st.crewChoices(UI.pid, r && r.paid_by_customer_id), r && r.paid_by_customer_id)), r,
        (s, v) => s.saveCost(UI.pid, { date: v("fDate"), quantity: v("fQty"), description: v("fDesc"), price: v("fPrice"),
          paid_by: v("fPaid") }, r && r.id, r && r.version),
        s => s.deleteBooking("operating_cost", r.id, r.version));
    },
    edit_sales(id) {
      const st = Cloud.store, r = id ? UI.find(st.sales(UI.pid), id) : null;
      const recv = r && r.cash_received_by_customer_id;
      UI.form(r ? "Holzverkauf" : "Neuer Holzverkauf",
        UI.field("fDate", "Datum", r ? r.sale_date : UI.today(), "date") +
        UI.select("fCust", "Kunde*", UI.opts(st.customerChoices(r && r.customer_id), r && r.customer_id)) +
        UI.field("fQty", "Menge (RM)", r ? UI.de(r.quantity_rm) : "", "num") +
        `<div id="extBox">` +
        UI.field("fPrice", "Preis/RM (€)", r && r.price_source === "MANUAL" ? UI.de(r.price_per_rm) : "", "num") +
        UI.field("fTotal", "Gesamtpreis (€)", r && r.price_source === "MANUAL" ? UI.de(r.total_price) : "", "num") +
        UI.select("fRecv", LABEL_RECEIVER + "*", UI.opts(st.crewChoices(UI.pid, recv), recv)) +
        `<div class="hint" style="color:#1565c0">Die Rechnung für externe Kunden wird in der Windows-App erstellt.</div></div>` +
        `<div id="crewBox" class="hint">Preis/RM und Gesamtpreis werden in der Windows-App über „Preis für Crew ermitteln“ berechnet.</div>`, r,
        (s, v) => s.saveSale(UI.pid, { date: v("fDate"), customer_id: v("fCust"), quantity_rm: v("fQty"), price_per_rm: v("fPrice"),
          total_price: v("fTotal"), cash_received_by: v("fRecv") }, r && r.id, r && r.version),
        s => s.deleteBooking("wood_sale", r.id, r.version));
      const toggle = () => {
        const cid = $("fCust").value, ext = cid && st.customerTypeIn(UI.pid, cid) === EXTERNAL;
        $("extBox").style.display = ext ? "" : "none"; $("crewBox").style.display = cid && !ext ? "" : "none";
      };
      $("fCust").onchange = toggle; toggle(); UI.autoTotal("fQty", "fPrice", "fTotal");
    },
    edit_customers(id) {
      const st = Cloud.store, c = id ? st.customer(id) : null;
      const fixed = c ? st.fixedTypes(c.id) : [];
      const t = c ? c.customer_type : "";
      UI.form(c ? `Kunde ${c.customer_no}` : "Neuer Kunde",
        UI.field("cFirst", "Vorname", c ? c.first_name : "") + UI.field("cLast", "Name*", c ? c.last_name : "") +
        UI.field("cStreet", "Straße", c ? c.street : "") + UI.field("cPlz", "PLZ", c ? c.postal_code : "", "text", `inputmode="numeric" maxlength="5"`) +
        UI.field("cCity", "Ort", c ? c.city : "") + UI.field("cPaypal", "PayPal (E-Mail)", c ? c.paypal : "", "email", `autocapitalize="off"`) +
        UI.select("cType", "Kundenart*", `<option value="">– bitte wählen –</option>` +
          [CREW, EXTERNAL].map(k => `<option value="${k}" ${k === t ? "selected" : ""}>${TYPE_TEXT[k]}</option>`).join("")) +
        (fixed.length ? `<div class="hint">Festgeschrieben in: ${fixed.map(f => `${esc(f.name)} (${TYPE_TEXT[f.customer_type]})`).join(", ")}.<br>
          Eine geänderte Kundenart gilt nur für Saisons, in denen der Kunde noch keine Buchungen hat.</div>` : ""),
        c,
        (s, v) => s.saveCustomer({ first_name: v("cFirst"), last_name: v("cLast"), street: v("cStreet"), postal_code: v("cPlz"),
          city: v("cCity"), paypal: v("cPaypal"), customer_type: v("cType") }, c && c.id, c && c.version),
        s => s.deleteCustomer(c.id, c.version));
    },

    settings() {
      UI.openSheet(`<h2>Einstellungen</h2>
        <label for="sPath">Datenbank im OneDrive (Pfad ab OneDrive-Hauptordner)</label>
        <input id="sPath" value="${esc(Cloud.path())}" autocapitalize="off" autocorrect="off">
        <div class="hint">Beispiel: Holzwirtschaft/holzwirtschaft.db – derselbe Ordner, den die Windows-App nutzt.</div>
        <div class="btns"><button class="btn blue" id="sLoad">Speichern &amp; neu laden</button></div>
        <div class="btns"><button class="btn edit" id="sRelease">Bearbeitung beenden (Sperre freigeben)</button></div>
        <div class="hint">Danach kann die Windows-App die Datenbank ohne Rückfrage bearbeiten.</div>
        <div class="btns"><button class="btn del" id="sLogout">Abmelden</button><button class="btn sec" id="sClose">Schließen</button></div>
        <div class="hint">Holzwirtschaft iPhone 1.0.1</div>`);
      $("sClose").onclick = UI.closeSheet;
      $("sLoad").onclick = async () => { ls.set("dbPath", $("sPath").value.trim().replace(/^\/+/, "")); UI.closeSheet(); await Cloud.releaseLock().catch(() => {}); UI.connect(); };
      $("sRelease").onclick = async () => {
        try { await Cloud.releaseLock(); Cloud.released = true; Cloud.readOnly = true; Cloud.lockedBy = "Bearbeitung beendet"; UI.closeSheet(); UI.render();
              UI.toast("Sperre freigegeben. Zum Bearbeiten ⟳ antippen."); } catch (e) { UI.error(e); }
      };
      $("sLogout").onclick = async () => { await Cloud.releaseLock().catch(() => {}); Auth.logout(); UI.closeSheet(); UI.showLogin(); };
    }
  };

  // ------------------------------------------------------------ Ereignisse
  document.querySelectorAll("nav button").forEach(b => b.onclick = () => { UI.tab = b.dataset.tab; ls.set("tab", UI.tab); UI.render(); window.scrollTo(0, 0); });
  $("period").onchange = e => { UI.pid = e.target.value; ls.set("period", UI.pid); UI.render(); };
  $("fab").onclick = () => { try { UI["edit_" + UI.tab](null); } catch (e) { UI.error(e); } };
  $("btnMenu").onclick = () => UI.settings();
  $("btnReload").onclick = () => { if (Cloud.store) { Cloud.releaseLock().catch(() => {}).then(UI.connect); } else UI.start(); };
  $("sheetBg").onclick = e => { if (e.target.id === "sheetBg") UI.closeSheet(); };
  document.addEventListener("visibilitychange", () => { if (!document.hidden) Cloud.tick(); });
  window.HW_UI = UI; window.HW_CLOUD = Cloud;
  UI.start();
})();
