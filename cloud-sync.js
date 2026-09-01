(() => {
  const SYNC_URL = "https://ytdacypygsfalkixhemj.supabase.co/functions/v1/fishing-sync";
  const STORAGE_KEY = "fishing-day-mvp-v3";
  const POLL_MS = 5000;
  const PUSH_DELAY_MS = 700;
  const baseline = new Map();
  let pushTimer = null;
  let syncing = false;
  let ready = false;

  function dayJson(day) {
    try { return JSON.stringify(day || {}); } catch { return "{}"; }
  }

  function setCloudStatus(text, ok = true) {
    let badge = document.getElementById("cloudSyncBadge");
    if (!badge) {
      badge = document.createElement("div");
      badge.id = "cloudSyncBadge";
      badge.style.cssText = "position:fixed;right:12px;bottom:12px;z-index:9999;padding:8px 11px;border-radius:999px;font:700 11px Arial,sans-serif;box-shadow:0 6px 20px rgba(0,0,0,.12);background:#fff;color:#17342f;border:1px solid #d9dfd7";
      document.body.appendChild(badge);
    }
    badge.textContent = text;
    badge.style.borderColor = ok ? "#9bc7b9" : "#e4b4aa";
    badge.style.background = ok ? "#f3fbf6" : "#fff5f2";

    const syncStatus = document.getElementById("syncStatus");
    if (syncStatus && ready) syncStatus.textContent = `Облачная синхронизация: ${text}. Google Таблица остаётся отдельной выгрузкой.`;
  }

  async function api(method = "GET", body = null) {
    const res = await fetch(SYNC_URL, {
      method,
      headers: { "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined,
      cache: "no-store"
    });
    if (!res.ok) throw new Error(`Cloud sync ${res.status}`);
    return await res.json();
  }

  async function fetchRemoteRows() {
    const result = await api("GET");
    if (!result?.ok) throw new Error(result?.error || "Cloud read failed");
    return Array.isArray(result.rows) ? result.rows : [];
  }

  async function pushRows(rows) {
    if (!rows.length) return;
    const result = await api("POST", { rows });
    if (!result?.ok) throw new Error(result?.error || "Cloud write failed");
    rows.forEach(row => baseline.set(row.day_date, dayJson(row.data)));
  }

  async function pushChangedLocal() {
    if (!window.state?.days) return;
    const changed = [];
    for (const [date, day] of Object.entries(window.state.days)) {
      const json = dayJson(day);
      if (baseline.get(date) !== json) changed.push({ day_date: date, data: day });
    }
    if (changed.length) {
      setCloudStatus("сохраняю изменения…");
      await pushRows(changed);
      setCloudStatus("синхронизировано");
    }
  }

  async function pullRemote({ initial = false } = {}) {
    if (!window.state?.days) return;
    const rows = await fetchRemoteRows();
    const remoteDates = new Set(rows.map(r => r.day_date));
    let changedLocal = false;

    if (initial && rows.length === 0) {
      const localRows = Object.entries(window.state.days).map(([day_date, data]) => ({ day_date, data }));
      if (localRows.length) await pushRows(localRows);
      return;
    }

    for (const row of rows) {
      const remoteJson = dayJson(row.data);
      const baseJson = baseline.get(row.day_date);
      const localJson = dayJson(window.state.days[row.day_date]);
      if (baseJson !== remoteJson && localJson !== remoteJson) {
        window.state.days[row.day_date] = row.data;
        changedLocal = true;
      }
      baseline.set(row.day_date, remoteJson);
    }

    if (initial) {
      const missingRemote = [];
      for (const [date, day] of Object.entries(window.state.days)) {
        if (!remoteDates.has(date)) missingRemote.push({ day_date: date, data: day });
      }
      if (missingRemote.length) await pushRows(missingRemote);
    }

    if (changedLocal) {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(window.state)); } catch {}
      if (typeof window.renderAll === "function") window.renderAll();
      setCloudStatus("получены изменения");
      setTimeout(() => setCloudStatus("синхронизировано"), 900);
    }
  }

  function schedulePush() {
    if (!ready) return;
    clearTimeout(pushTimer);
    pushTimer = setTimeout(() => pushChangedLocal().catch(err => {
      console.error("Cloud push error", err);
      setCloudStatus("ошибка сохранения", false);
    }), PUSH_DELAY_MS);
  }

  function wrapLocalSave() {
    if (typeof window.saveData !== "function" || window.saveData.__cloudWrapped) return;
    const original = window.saveData;
    const wrapped = function(...args) {
      const result = original.apply(this, args);
      schedulePush();
      return result;
    };
    wrapped.__cloudWrapped = true;
    window.saveData = wrapped;
  }

  async function syncLoop() {
    if (syncing || !ready) return;
    syncing = true;
    try {
      await pushChangedLocal();
      await pullRemote();
    } catch (err) {
      console.error("Cloud sync error", err);
      setCloudStatus("нет связи с облаком", false);
    } finally {
      syncing = false;
    }
  }

  async function start() {
    setCloudStatus("подключаю облако…");
    try {
      wrapLocalSave();
      await pullRemote({ initial: true });
      ready = true;
      wrapLocalSave();
      await pushChangedLocal();
      setCloudStatus("синхронизировано");
      setInterval(syncLoop, POLL_MS);
      window.addEventListener("focus", syncLoop);
      document.addEventListener("visibilitychange", () => { if (!document.hidden) syncLoop(); });
    } catch (err) {
      console.error("Cloud bootstrap error", err);
      setCloudStatus("облако недоступно", false);
    }
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start);
  else start();
})();
