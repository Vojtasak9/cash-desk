// Old address of the cash desk. Sends anything not yet in the sheet, then copies all data of this device
// so the new Sidewinder icon can paste it (a Home Screen app on iPad has its own storage).
"use strict";
const $ = s => document.querySelector(s);
const NEW_URL = new URL("../sidewinder/", location.href).href;
const MOVE_TAG = "SIDEWINDER-MOVE:";
const pad = n => String(n).padStart(2, "0");
const today = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
let db = null, state = null;

$("#newUrl").textContent = NEW_URL.replace(/^https?:\/\//, "");
$("#open").href = NEW_URL;

function openDb() {
  return new Promise((resolve, reject) => {
    const r = indexedDB.open("viper", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}
const dbGet = key => new Promise((resolve, reject) => {
  const r = db.transaction("kv").objectStore("kv").get(key);
  r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error);
});
const dbPut = (key, value) => new Promise((resolve, reject) => {
  const tx = db.transaction("kv", "readwrite");
  tx.objectStore("kv").put(value, key);
  tx.oncomplete = () => resolve(); tx.onerror = () => reject(tx.error);
});

function status(el, text, cls) { el.textContent = text; el.className = cls || ""; }

// Changes the old app hadn't sent yet go to the sheet now (the same request the app makes)
async function sendOutbox() {
  const ops = state.outbox || [];
  if (!ops.length) return;
  const el = $("#sendStatus");
  if (!state.sync || !state.sync.url || !state.sync.key) {
    return status(el, `${ops.length} changes aren't in the sheet yet. They move with your data and Sidewinder sends them.`, "bad");
  }
  status(el, `Sending ${ops.length} changes to the sheet…`);
  try {
    const r = await fetch(state.sync.url, {method: "POST", headers: {"Content-Type": "text/plain;charset=utf-8"},
                                           body: JSON.stringify({action: "sync", ops, date: today(), key: state.sync.key})});
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || "error");
    state.outbox = state.outbox.filter(o => !ops.includes(o));
    state.sync.lastSync = new Date().toISOString();
    await dbPut("state", state);
    status(el, `${ops.length} changes sent to the sheet ✓` + (j.skipped && j.skipped.length ? ` (not written, typed by hand there: ${j.skipped.join(", ")})` : ""), "ok");
  } catch (e) {
    status(el, `Couldn't send ${ops.length} changes now (${e.message}). They move with your data and Sidewinder sends them.`, "bad");
  }
}

$("#copy").onclick = async () => {
  if (!state) return status($("#copyStatus"), " Nothing saved on this device - just open Sidewinder.", "muted");
  const text = MOVE_TAG + JSON.stringify(state);
  try {
    await navigator.clipboard.writeText(text);
    status($("#copyStatus"), " Copied ✓", "ok");
  } catch (e) {  // no clipboard access: let them copy it by hand
    const t = $("#copyText");
    t.value = text; t.style.display = ""; t.focus(); t.select();
    status($("#copyStatus"), " Select all the text below and copy it.", "muted");
  }
};

(async () => {
  try {
    db = await openDb();
    state = await dbGet("state");
  } catch (e) {
    return status($("#summary"), "Can't read this device's storage (" + e.message + ").", "bad");
  }
  if (!state) return status($("#summary"), "Nothing saved here. Just open Sidewinder.");
  const days = new Set(state.payments.map(p => p.date)).size;
  $("#summary").textContent = `${state.payments.length} sales (${days} days), ${(state.debts || []).length} Půjčky entries, ` +
    (state.sync && state.sync.key ? "key saved." : "no key saved.");
  await sendOutbox();
})();
