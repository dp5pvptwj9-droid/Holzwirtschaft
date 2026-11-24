/* Holzwirtschaft iPhone – Fachlogik (entspricht main.py 1.7.0/1.8.0)
   Läuft im Browser (sql.js) und zum Testen unter Node (node:sqlite). */
(function (root) {
  "use strict";
  const CREW = "CREW", EXTERNAL = "EXTERNAL";
  const TYPE_TEXT = { CREW: "Teil der Crew", EXTERNAL: "externer Kunde" };
  const STATUS_TEXT = { OPEN: "Offen", CALCULATED: "Preis ermittelt",
    INVOICED: "Crew-Rechnungen erstellt", CLOSED: "Abgeschlossen" };
  const EDITABLE_STATUS = ["OPEN", "CALCULATED", "INVOICED"];
  const LABEL_RECEIVER = "Einnahme hat erhalten";
  const KNOWN_PAYPAL = { "torsten|riesenberg": "torsten.riesenberg@web.de" };
  const EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
  const CONFLICT_MSG = "Der Datensatz wurde zwischenzeitlich geändert oder gelöscht " +
    "(eventuell auf einem anderen Gerät).\n\nDie Ansicht wird aktualisiert. Bitte die Änderung erneut vornehmen.";
  const USED_SQL = `
    SELECT paid_by_customer_id AS cid FROM wood_purchase WHERE period_id = ? AND deleted_at IS NULL
    UNION SELECT paid_by_customer_id FROM operating_cost WHERE period_id = ? AND deleted_at IS NULL
    UNION SELECT customer_id FROM wood_sale WHERE period_id = ? AND deleted_at IS NULL
    UNION SELECT cash_received_by_customer_id FROM wood_sale
          WHERE period_id = ? AND deleted_at IS NULL AND cash_received_by_customer_id IS NOT NULL`;

  class AppError extends Error {}

  // ---------------------------------------------------------- Zahlen (exakt, ohne float)
  /** Text (deutsch "1.234,5" oder "1234.5") -> BigInt mit `places` Nachkommastellen, kaufm. gerundet. */
  function scaled(value, places) {
    if (value === null || value === undefined) return 0n;
    let s = String(value).trim().replace(/\s|€/g, "");
    if (s === "") return 0n;
    if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
    if (!/^-?\d*(\.\d*)?$/.test(s) || s === "." || s === "-") throw new AppError(`Ungültige Zahl: ${value}`);
    const neg = s.startsWith("-"); if (neg) s = s.slice(1);
    let [ip, fp = ""] = s.split(".");
    ip = ip || "0";
    const extra = fp.length > places ? fp[places] : "0";
    fp = (fp + "0".repeat(places)).slice(0, places);
    let n = BigInt(ip + fp);
    if (extra >= "5") n += 1n;                         // ROUND_HALF_UP
    return neg ? -n : n;
  }
  /** BigInt mit Nachkommastellen -> Text wie Python str(Decimal), z. B. "12.0", "84.84". */
  function unscaled(n, places) {
    const neg = n < 0n; if (neg) n = -n;
    let s = n.toString().padStart(places + 1, "0");
    const out = places ? s.slice(0, -places) + "." + s.slice(-places) : s;
    return (neg ? "-" : "") + out;
  }
  const money = v => unscaled(scaled(v, 2), 2);
  const rm = v => unscaled(scaled(v, 1), 1);
  function fmtNum(v, places = 2) {
    if (v === null || v === undefined || v === "") return "";
    const s = unscaled(scaled(v, places), places);
    let [ip, fp] = s.split(".");
    const neg = ip.startsWith("-"); if (neg) ip = ip.slice(1);
    ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return (neg ? "-" : "") + ip + (fp ? "," + fp : "");
  }
  const fmtRm = v => fmtNum(v, 1);
  const fmtEur = v => (v === null || v === undefined || v === "") ? "" : fmtNum(v, 2) + " €";
  const fmtDate = iso => iso ? `${iso.slice(8, 10)}.${iso.slice(5, 7)}.${iso.slice(0, 4)}` : "";
  const nowIso = () => new Date().toISOString().replace(/\.\d{3}Z$/, "+00:00");
  const newId = () => (root.crypto && root.crypto.randomUUID) ? root.crypto.randomUUID()
    : require("crypto").randomUUID();
  const personName = r => r ? `${r.first_name} ${r.last_name}`.trim() : "";
  const displayName = r => (!r || r.deleted_at) ? "gelöscht" : personName(r);

  // ---------------------------------------------------------- Store
  class Store {
    /** db: {all(sql, params) -> [obj], run(sql, params) -> changes} */
    constructor(db) { this.db = db; }
    all(sql, p = []) { return this.db.all(sql, p); }
    one(sql, p = []) { return this.db.all(sql, p)[0] || null; }
    run(sql, p = []) { return this.db.run(sql, p); }

    transaction(fn) {
      this.run("BEGIN IMMEDIATE");
      try { const r = fn(); this.run("COMMIT"); return r; }
      catch (e) { try { this.run("ROLLBACK"); } catch (_) {} throw e; }
    }
    checkSchema() {
      const t = new Set(this.all("SELECT name FROM sqlite_master WHERE type='table'").map(r => r.name));
      for (const n of ["accounting_period", "customer", "customer_period_type", "customer_order",
                       "wood_purchase", "operating_cost", "wood_sale"])
        if (!t.has(n)) throw new AppError("Die Datenbank hat einen älteren Stand. Bitte einmal mit der " +
          "Windows-App (ab Version 1.6.0) öffnen, damit sie umgestellt wird.");
    }
    _insert(table, values) {
      const ts = nowIso();
      const v = { id: values.id || newId(), ...values, created_at: ts, updated_at: ts };
      const cols = Object.keys(v);
      this.run(`INSERT INTO ${table} (${cols.join(", ")}) VALUES (${cols.map(() => "?").join(", ")})`,
               cols.map(c => v[c]));
      return v.id;
    }
    _update(table, id, version, values, soft = true) {
      const cols = Object.keys(values);
      let sql = `UPDATE ${table} SET ${cols.map(c => c + " = ?").join(", ")}, version = version + 1, ` +
                "updated_at = ? WHERE id = ? AND version = ?";
      if (soft) sql += " AND deleted_at IS NULL";
      if (this.run(sql, [...cols.map(c => values[c]), nowIso(), id, version]) !== 1)
        throw new AppError(CONFLICT_MSG);
    }

    // Perioden
    periods() { return this.all("SELECT * FROM accounting_period ORDER BY start_date DESC, name"); }
    period(pid) {
      const r = this.one("SELECT * FROM accounting_period WHERE id = ?", [pid]);
      if (!r) throw new AppError("Die Abrechnungsperiode wurde nicht gefunden.");
      return r;
    }
    editable(pid) { return EDITABLE_STATUS.includes(this.period(pid).status); }
    _requireEditable(pid) {
      const p = this.period(pid);
      if (!EDITABLE_STATUS.includes(p.status))
        throw new AppError(`Die Abrechnungsperiode „${p.name}“ ist abgeschlossen. Es sind keine Änderungen mehr möglich.`);
      return p;
    }
    static _checkDate(p, iso) {
      if (!iso) throw new AppError("Bitte ein Datum eingeben.");
      if (!(p.start_date <= iso && iso <= p.end_date))
        throw new AppError(`Das Datum muss im Zeitraum der Abrechnungsperiode liegen (${fmtDate(p.start_date)} – ${fmtDate(p.end_date)}).`);
    }

    // Kunden
    customers(includeDeleted = false) {
      return this.all(`SELECT * FROM customer${includeDeleted ? "" : " WHERE deleted_at IS NULL"} ORDER BY customer_no ASC`);
    }
    customer(cid) {
      const r = this.one("SELECT * FROM customer WHERE id = ?", [cid]);
      if (!r) throw new AppError("Der Kunde wurde nicht gefunden.");
      return r;
    }
    periodCustomerTypes(pid) {
      const t = {};
      for (const r of this.all("SELECT id, customer_type FROM customer")) t[r.id] = r.customer_type;
      if (pid) for (const r of this.all("SELECT customer_id, customer_type FROM customer_period_type WHERE period_id = ?", [pid]))
        t[r.customer_id] = r.customer_type;
      return t;
    }
    customerTypeIn(pid, cid) {
      const r = this.one("SELECT customer_type FROM customer_period_type WHERE period_id = ? AND customer_id = ?", [pid, cid]);
      return r ? r.customer_type : this.customer(cid).customer_type;
    }
    fixedTypes(cid) {
      return this.all("SELECT p.name, cpt.customer_type FROM customer_period_type cpt " +
        "JOIN accounting_period p ON p.id = cpt.period_id WHERE cpt.customer_id = ? ORDER BY p.start_date, p.name", [cid]);
    }
    _syncTypeLocks(pid) {
      if (this.period(pid).status === "CLOSED") return;
      const ts = nowIso(), p4 = [pid, pid, pid, pid];
      this.run("INSERT OR IGNORE INTO customer_period_type (customer_id, period_id, customer_type, fixed_at) " +
               `SELECT c.id, ?, c.customer_type, ? FROM customer c WHERE c.id IN (${USED_SQL})`, [pid, ts, ...p4]);
      this.run(`DELETE FROM customer_period_type WHERE period_id = ? AND customer_id NOT IN (${USED_SQL})`, [pid, ...p4]);
    }
    saveCustomer(data, cid = null, version = null) {
      const v = {};
      for (const k of ["first_name", "last_name", "street", "postal_code", "city", "paypal"]) v[k] = (data[k] || "").trim();
      const ctype = data.customer_type;
      if (ctype !== CREW && ctype !== EXTERNAL)
        throw new AppError("Bitte die Kundenart auswählen: „Teil der Crew“ oder „externer Kunde“.");
      if (!v.last_name) throw new AppError("Bitte mindestens den Namen eingeben.");
      if (v.postal_code && !/^\d{4,5}$/.test(v.postal_code))
        throw new AppError("Die Postleitzahl darf nur aus 4 bis 5 Ziffern bestehen.");
      if (v.paypal && !EMAIL_RE.test(v.paypal)) throw new AppError("Die PayPal-Adresse ist keine gültige E-Mail-Adresse.");
      v.customer_type = ctype;
      if (!cid) {
        if (!v.paypal) v.paypal = KNOWN_PAYPAL[`${v.first_name.toLowerCase()}|${v.last_name.toLowerCase()}`] || "";
        return this.transaction(() => {
          const no = this.one("SELECT COALESCE(MAX(customer_no), 0) + 1 AS n FROM customer").n;
          return this._insert("customer", { ...v, customer_no: no });
        });
      }
      this._update("customer", cid, version, v);
      return cid;
    }
    deleteCustomer(cid, version) { this._update("customer", cid, version, { deleted_at: nowIso() }); }

    _checkCrewPerson(pid, cid, previous, role) {
      if (!cid) throw new AppError(`Bitte im Feld „${role}“ ein Crew-Mitglied auswählen.`);
      const row = this.customer(cid);
      if (this.customerTypeIn(pid, cid) !== CREW)
        throw new AppError(`Im Feld „${role}“ sind nur Kunden möglich, die in dieser Saison „Teil der Crew“ sind.`);
      if (row.deleted_at && cid !== previous) throw new AppError(`Der im Feld „${role}“ gewählte Kunde wurde gelöscht.`);
    }
    static _checkTotal(qty, price, total) {
      // qty: 1 Stelle, price/total: 4 Stellen -> Vergleich mit 5 Stellen
      const q = scaled(qty, 1), p = scaled(price, 4), t = scaled(total, 4);
      let diff = q * p - t * 10n; if (diff < 0n) diff = -diff;
      if (diff > (100n + q * 5n) * 10n) {
        const exact = money(unscaled(q * p, 5));
        throw new AppError(`Der Gesamtpreis passt nicht zu Menge × Preis/RM (rechnerisch ${fmtEur(exact)}).`);
      }
    }
    _booking(table, rid) {
      const r = this.one(`SELECT * FROM ${table} WHERE id = ?`, [rid]);
      if (!r || r.deleted_at) throw new AppError(CONFLICT_MSG);
      return r;
    }
    deleteBooking(table, rid, version) {
      if (!["wood_purchase", "operating_cost", "wood_sale", "customer_order"].includes(table)) throw new Error(table);
      const row = this._booking(table, rid);
      this._requireEditable(row.period_id);
      this.transaction(() => {
        this._update(table, rid, version, { deleted_at: nowIso() });
        this._syncTypeLocks(row.period_id);
      });
    }
    _saveBooking(table, pid, values, rid, version) {
      return this.transaction(() => {
        if (!rid) rid = this._insert(table, { period_id: pid, ...values });
        else this._update(table, rid, version, values);
        this._syncTypeLocks(pid);
        return rid;
      });
    }
    static _positive(v, msg) { if (scaled(v, 4) <= 0n) throw new AppError(msg); }
    static _notNegative(v, msg) { if (scaled(v, 4) < 0n) throw new AppError(msg); }

    // Bestellungen
    orders(pid) { return this.all("SELECT * FROM customer_order WHERE period_id = ? AND deleted_at IS NULL ORDER BY order_date, created_at", [pid]); }
    saveOrder(pid, d, rid = null, version = null) {
      const p = this._requireEditable(pid); Store._checkDate(p, d.date);
      const qty = rm(d.quantity_rm);
      if (scaled(qty, 1) <= 0n) throw new AppError("Die bestellte Menge muss größer als 0 sein.");
      const cid = d.customer_id;
      if (!cid) throw new AppError("Bitte einen Kunden auswählen.");
      const prev = rid ? this._booking("customer_order", rid).customer_id : null;
      if (this.customer(cid).deleted_at && cid !== prev) throw new AppError("Der gewählte Kunde wurde gelöscht.");
      const values = { order_date: d.date, customer_id: cid, quantity_rm: qty };
      return this.transaction(() => {
        if (!rid) return this._insert("customer_order", { period_id: pid, ...values });
        this._update("customer_order", rid, version, values); return rid;
      });
    }
    orderCheck(pid) {
      const orders = this.orders(pid);
      const ordered = orders.reduce((s, o) => s + scaled(o.quantity_rm, 1), 0n);
      const purchased = this.purchases(pid).reduce((s, r) => s + scaled(r.quantity_rm, 1), 0n);
      const diff = purchased - ordered;
      if (!orders.length) return { state: "NONE", text: "Für diese Saison sind noch keine Bestellungen erfasst." };
      const base = `Bestellt: ${fmtRm(unscaled(ordered, 1))} RM (${orders.length}) · Eingekauft: ${fmtRm(unscaled(purchased, 1))} RM – `;
      if (diff < 0n) return { state: "SHORT", text: base + `es fehlen ${fmtRm(unscaled(-diff, 1))} RM.` };
      if (diff > 0n) return { state: "SURPLUS", text: base + `${fmtRm(unscaled(diff, 1))} RM übrig.` };
      return { state: "EXACT", text: base + "deckt die Bestellungen genau." };
    }

    // Holzeinkauf
    purchases(pid) { return this.all("SELECT * FROM wood_purchase WHERE period_id = ? AND deleted_at IS NULL ORDER BY purchase_date, created_at", [pid]); }
    savePurchase(pid, d, rid = null, version = null) {
      const p = this._requireEditable(pid); Store._checkDate(p, d.date);
      const qty = rm(d.quantity_rm);
      if (scaled(qty, 1) <= 0n) throw new AppError("Die Menge muss größer als 0 sein.");
      Store._notNegative(d.price_per_rm, "Preise dürfen nicht negativ sein.");
      Store._notNegative(d.total_price, "Preise dürfen nicht negativ sein.");
      Store._checkTotal(qty, d.price_per_rm, d.total_price);
      const prev = rid ? this._booking("wood_purchase", rid).paid_by_customer_id : null;
      this._checkCrewPerson(pid, d.paid_by, prev, "Bezahlt von");
      return this._saveBooking("wood_purchase", pid, { purchase_date: d.date, quantity_rm: qty,
        price_per_rm: money(d.price_per_rm), total_price: money(d.total_price), paid_by_customer_id: d.paid_by }, rid, version);
    }

    // Betriebskosten
    costs(pid) { return this.all("SELECT * FROM operating_cost WHERE period_id = ? AND deleted_at IS NULL ORDER BY cost_date, created_at", [pid]); }
    saveCost(pid, d, rid = null, version = null) {
      const p = this._requireEditable(pid); Store._checkDate(p, d.date);
      const desc = (d.description || "").trim();
      if (!desc) throw new AppError("Bitte eine Bezeichnung/Art eingeben.");
      Store._positive(d.quantity, "Die Menge muss größer als 0 sein.");
      Store._notNegative(d.price, "Der Preis darf nicht negativ sein.");
      const prev = rid ? this._booking("operating_cost", rid).paid_by_customer_id : null;
      this._checkCrewPerson(pid, d.paid_by, prev, "Bezahlt von");
      return this._saveBooking("operating_cost", pid, { cost_date: d.date, description: desc,
        quantity: money(d.quantity), price: money(d.price), paid_by_customer_id: d.paid_by }, rid, version);
    }

    // Holzverkauf
    sales(pid) { return this.all("SELECT * FROM wood_sale WHERE period_id = ? AND deleted_at IS NULL ORDER BY sale_date, created_at", [pid]); }
    saveSale(pid, d, rid = null, version = null) {
      const p = this._requireEditable(pid); Store._checkDate(p, d.date);
      const qty = rm(d.quantity_rm);
      if (scaled(qty, 1) <= 0n) throw new AppError("Die Menge muss größer als 0 sein.");
      const cid = d.customer_id;
      if (!cid) throw new AppError("Bitte einen Kunden auswählen.");
      const prev = rid ? this._booking("wood_sale", rid) : null;
      const buyer = this.customer(cid);
      if (buyer.deleted_at && (!prev || prev.customer_id !== cid)) throw new AppError("Der gewählte Kunde wurde gelöscht.");
      let v;
      if (this.customerTypeIn(pid, cid) === EXTERNAL) {
        Store._notNegative(d.price_per_rm, "Preise dürfen nicht negativ sein.");
        Store._notNegative(d.total_price, "Preise dürfen nicht negativ sein.");
        if (scaled(d.total_price, 4) === 0n) throw new AppError("Bitte bei externen Kunden Preis/RM bzw. Gesamtpreis eintragen.");
        Store._checkTotal(qty, d.price_per_rm, d.total_price);
        this._checkCrewPerson(pid, d.cash_received_by, prev ? prev.cash_received_by_customer_id : null, LABEL_RECEIVER);
        v = { price_per_rm: money(d.price_per_rm), total_price: money(d.total_price),
              price_source: "MANUAL", cash_received_by_customer_id: d.cash_received_by };
      } else {
        v = { price_per_rm: null, total_price: null, price_source: "PENDING", cash_received_by_customer_id: null };
      }
      return this._saveBooking("wood_sale", pid, { sale_date: d.date, customer_id: cid, quantity_rm: qty,
        ...v, calculation_run_id: null }, rid, version);
    }

    /** Auswahlliste: Crew-Mitglieder der Saison (gelöschte nur, wenn bereits eingetragen). */
    crewChoices(pid, current = null) {
      const types = this.periodCustomerTypes(pid);
      return this.customers(true).filter(c => types[c.id] === CREW && (!c.deleted_at || c.id === current));
    }
    customerChoices(current = null) {
      return this.customers(true).filter(c => !c.deleted_at || c.id === current);
    }
  }

  const api = { Store, AppError, CREW, EXTERNAL, TYPE_TEXT, STATUS_TEXT, LABEL_RECEIVER,
    scaled, unscaled, money, rm, fmtNum, fmtRm, fmtEur, fmtDate, personName, displayName, nowIso };
  if (typeof module !== "undefined") module.exports = api; else root.HW = api;
})(typeof window !== "undefined" ? window : globalThis);
