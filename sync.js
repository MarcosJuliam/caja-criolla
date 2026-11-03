/* Caja Criolla — local store + outbox that syncs to a Google Sheets Apps Script web app.
   Exposes window.CC_DB with the small document API the app already uses:
   doc(path).get/set/update/onSnapshot and collection("dias").where(...).get(). */
(function () {
  "use strict";
  const K = { docs: "cc_docs", out: "cc_outbox", conn: "cc_conexion" };
  const KEEP_DAYS = 62;
  function load(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
  function store(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }
  function clone(o) { return o == null ? o : JSON.parse(JSON.stringify(o)); }
  function isObj(v) { return v && typeof v === "object" && !Array.isArray(v); }
  function merge(a, b) { for (const k in b) { const v = b[k]; if (isObj(v) && isObj(a[k])) merge(a[k], v); else a[k] = clone(v); } return a; }
  function ymd(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }

  let docs = load(K.docs, {});
  let outbox = load(K.out, []);
  let conn = Object.assign({ url: "", clave: "", equipo: "" }, load(K.conn, {}));
  const listeners = {};
  const estadoCbs = new Set();
  const st = { enviando: false, error: "", ultimoOk: 0 };
  let n = 0;
  const oid = () => Date.now().toString(36) + (n++).toString(36) + Math.random().toString(36).slice(2, 5);

  function persist() { const a = store(K.docs, docs), b = store(K.out, outbox); return a && b; }
  function snap(p) { const d = docs[p]; return { exists: !!d, id: p.split("/").pop(), data: () => clone(d) }; }
  function notify(p) { (listeners[p] || []).forEach(cb => { try { cb(snap(p)); } catch (e) { console.error(e); } }); }
  function estado() {
    return { conectado: !!(conn.url && conn.clave), pendientes: outbox.length, enviando: st.enviando, error: st.error,
      online: navigator.onLine !== false, ultimoOk: st.ultimoOk, equipo: conn.equipo };
  }
  function emit() { const e = estado(); estadoCbs.forEach(cb => { try { cb(e); } catch (x) { console.error(x); } }); }

  function opsFromDia(fecha, part) {
    const ops = [];
    ["ventas", "gastos", "merma"].forEach(k => {
      const o = part[k]; if (!isObj(o)) return;
      Object.keys(o).forEach(id => {
        const e = o[id]; if (!e) return;
        if (e.borrado) ops.push({ oid: oid(), op: "del", hoja: k, id, fecha });
        else ops.push({ oid: oid(), op: "add", hoja: k, id, fecha, data: clone(e) });
      });
    });
    if (Array.isArray(part.menu)) {
      outbox = outbox.filter(o => !(o.op === "menu" && o.fecha === fecha));
      ops.push({ oid: oid(), op: "menu", fecha, menu: part.menu.slice() });
    }
    return ops;
  }
  function prune() {
    const lim = ymd(new Date(Date.now() - KEEP_DAYS * 864e5));
    const pend = new Set(outbox.filter(o => o.fecha).map(o => o.fecha));
    Object.keys(docs).forEach(p => { if (p.startsWith("dias/")) { const f = p.slice(5); if (f < lim && !pend.has(f) && !listeners[p]) delete docs[p]; } });
  }
  function write(p, data, replace) {
    const prevDoc = clone(docs[p]), prevOut = outbox.slice();
    docs[p] = replace ? clone(data) : merge(docs[p] || {}, data);
    if (p.startsWith("dias/")) { const f = p.slice(5); docs[p].fecha = docs[p].fecha || f; outbox = outbox.concat(opsFromDia(f, data)); prune(); }
    else if (p === "config/menu") { outbox = outbox.filter(o => o.op !== "config"); outbox.push({ oid: oid(), op: "config", data: clone(docs[p]) }); }
    if (!persist()) {
      if (prevDoc === undefined) delete docs[p]; else docs[p] = prevDoc;
      outbox = prevOut; persist();
      const err = new Error("quota"); err.code = "quota_exceeded"; throw err;
    }
    notify(p); emit(); syncSoon(300);
  }

  /* ---- network ---- */
  function url(params) {
    const u = new URL(conn.url);
    Object.entries(Object.assign({ clave: conn.clave }, params)).forEach(([k, v]) => u.searchParams.set(k, v));
    return u.toString();
  }
  async function getJSON(params) {
    const r = await fetch(url(params), { method: "GET", redirect: "follow", cache: "no-store" });
    if (!r.ok) throw new Error("HTTP " + r.status);
    const j = await r.json();
    if (!j.ok) throw new Error(j.error === "clave" ? "La clave no coincide con la de la hoja." : (j.error || "Error de la hoja"));
    return j;
  }
  async function flush() {
    while (outbox.length) {
      const lote = outbox.slice(0, 40);
      const r = await fetch(conn.url, { method: "POST", redirect: "follow", headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ clave: conn.clave, equipo: conn.equipo || "Equipo", ops: lote }) });
      if (!r.ok) throw new Error("HTTP " + r.status);
      const j = await r.json();
      if (!j.ok) throw new Error(j.error === "clave" ? "La clave no coincide con la de la hoja." : (j.error || "Error de la hoja"));
      const hechos = new Set(j.hechos || []);
      outbox = outbox.filter(o => !hechos.has(o.oid));
      persist(); emit();
      if (!hechos.size) break;
    }
  }
  function overlay(fecha, base) {
    const d = base || { fecha, ventas: {}, gastos: {}, merma: {} };
    ["ventas", "gastos", "merma"].forEach(k => { d[k] = d[k] || {}; });
    outbox.forEach(o => {
      if (o.fecha !== fecha) return;
      if (o.op === "add") d[o.hoja][o.id] = clone(o.data);
      else if (o.op === "del") { if (d[o.hoja][o.id]) d[o.hoja][o.id].borrado = true; }
      else if (o.op === "menu") d.menu = o.menu.slice();
    });
    return d;
  }
  async function pull() {
    const fechas = Object.keys(listeners).filter(p => p.startsWith("dias/") && listeners[p].length).map(p => p.slice(5)).sort();
    const desde = fechas[0] || ymd(new Date()), hasta = fechas[fechas.length - 1] || desde;
    const j = await getJSON({ accion: "dias", desde, hasta, conConfig: 1 });
    const rem = j.dias || {};
    fechas.forEach(f => {
      const p = "dias/" + f, before = JSON.stringify(docs[p] || null);
      const r = rem[f] ? overlay(f, rem[f]) : (docs[p] || outbox.some(o => o.fecha === f) ? overlay(f, null) : null);
      if (r && docs[p] && docs[p].menu && !r.menu) r.menu = docs[p].menu;
      if (r) docs[p] = r; else delete docs[p];
      if (JSON.stringify(docs[p] || null) !== before) notify(p);
    });
    const pendCfg = outbox.some(o => o.op === "config");
    if (!pendCfg) {
      if (j.config) { const b = JSON.stringify(docs["config/menu"] || null); docs["config/menu"] = j.config; if (JSON.stringify(j.config) !== b) notify("config/menu"); }
      else if (docs["config/menu"]) outbox.push({ oid: oid(), op: "config", data: clone(docs["config/menu"]) });
    }
    persist();
  }
  let chain = Promise.resolve(), timer = null, again = false;
  function sync() {
    if (!conn.url || !conn.clave) { emit(); return Promise.resolve(); }
    if (st.enviando) { again = true; return chain; }
    st.enviando = true; emit();
    chain = (async () => {
      try {
        if (navigator.onLine === false) throw new Error("Sin internet");
        await flush(); await pull(); if (outbox.length) await flush();
        st.error = ""; st.ultimoOk = Date.now();
      } catch (e) { st.error = e && e.message ? e.message : "No se pudo enviar"; }
      st.enviando = false; emit();
      if (again) { again = false; syncSoon(500); }
    })();
    return chain;
  }
  function syncSoon(ms) { clearTimeout(timer); timer = setTimeout(sync, ms || 0); }
  setInterval(() => { if (document.visibilityState !== "hidden") sync(); }, 30000);
  window.addEventListener("online", () => syncSoon(200));
  window.addEventListener("offline", emit);
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") syncSoon(200); });

  window.CC_DB = {
    doc(p) {
      return {
        get: async () => snap(p),
        set: async d => write(p, d, true),
        update: async d => write(p, d, false),
        onSnapshot(cb) {
          (listeners[p] = listeners[p] || []).push(cb);
          setTimeout(() => cb(snap(p)), 0); syncSoon(100);
          return () => { listeners[p] = (listeners[p] || []).filter(x => x !== cb); };
        }
      };
    },
    collection(c) {
      const w = [];
      const q = {
        where(f, op, v) { w.push([op, v]); return q; },
        limit() { return q; },
        async get() {
          let desde = "0000", hasta = "9999";
          w.forEach(([op, v]) => { if (op === ">=") desde = v; if (op === "<=") hasta = v; });
          let dias = null;
          if (conn.url && conn.clave && navigator.onLine !== false) {
            try { await chain; dias = (await getJSON({ accion: "dias", desde, hasta })).dias || {}; } catch (e) { dias = null; }
          }
          const out = {};
          if (dias) Object.keys(dias).forEach(f => { out[f] = overlay(f, dias[f]); });
          Object.keys(docs).forEach(p => { if (p.startsWith(c + "/")) { const f = p.slice(c.length + 1); if (f >= desde && f <= hasta && !out[f]) out[f] = clone(docs[p]); } });
          outbox.forEach(o => { if (o.fecha && o.fecha >= desde && o.fecha <= hasta && !out[o.fecha]) out[o.fecha] = overlay(o.fecha, null); });
          return { docs: Object.values(out).map(d => ({ data: () => d })) };
        }
      };
      return q;
    },
    conexion: () => clone(conn),
    async guardarConexion(c) {
      conn = { url: String(c.url || "").trim(), clave: String(c.clave || "").trim(), equipo: String(c.equipo || "").trim() };
      store(K.conn, conn); st.error = ""; emit();
      if (!conn.url) return { ok: true };
      if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(conn.url)) return { ok: false, error: "La URL debe empezar con https://script.google.com/" };
      try { const j = await getJSON({ accion: "ping" }); syncSoon(0); return { ok: true, hoja: j.hoja }; }
      catch (e) { st.error = e.message || "No conecta"; emit(); return { ok: false, error: st.error }; }
    },
    importar(data) {
      let n = 0;
      if (data && isObj(data.config)) { write("config/menu", data.config, true); n++; }
      const dias = (data && data.dias) || {};
      Object.keys(dias).forEach(f => {
        const d = dias[f]; if (!isObj(d)) return;
        const part = {};
        ["ventas", "gastos", "merma"].forEach(k => {
          const o = d[k] || {}; part[k] = {};
          Object.keys(o).forEach(id => { if (o[id] && !o[id].borrado) { part[k][id] = o[id]; n++; } });
        });
        if (Array.isArray(d.menu)) part.menu = d.menu;
        write("dias/" + f, part, false);
      });
      return n;
    },
    respaldo() { const dias = {}; Object.keys(docs).forEach(p => { if (p.startsWith("dias/")) dias[p.slice(5)] = docs[p]; }); return { config: docs["config/menu"] || null, dias }; },
    estado, onEstado(cb) { estadoCbs.add(cb); cb(estado()); return () => estadoCbs.delete(cb); },
    sync
  };
})();
