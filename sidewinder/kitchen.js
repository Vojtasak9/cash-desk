// SIDEWINDER kitchen screen: the cook's phone shows today's orders from the sheet "Kuchyně" and ticks them off.
// The cash desk writes every sale there (with its toppings); this page asks the web app every few seconds.
"use strict";

const $ = s => document.querySelector(s);
const DEFAULT_URL = "https://script.google.com/macros/s/AKfycbwX_PD5J5_k-80yH49B47eSKnVwA_duDMtiA3Hn1tzfqz5Nkiq0Ep5JRUtiU1k0BLSc/exec";
const POLL_MS = 5000;
const pad = n => String(n).padStart(2, "0");
const today = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const nowTime = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; };
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c]));

// ---- connection: from the QR code (#u=...&k=...) or typed; the # part stays in the address so
// "Add to Home Screen" keeps it (on iPhone the Home Screen app has its own storage)
function store(k, v) { try { if (v == null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch (e) {} }
function load(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
const hash = new URLSearchParams(location.hash.slice(1));
if (hash.get("k")) { store("viperKitchenKey", hash.get("k")); store("viperKitchenUrl", hash.get("u") || DEFAULT_URL); }
let conf = {url: load("viperKitchenUrl") || DEFAULT_URL, key: load("viperKitchenKey") || ""};

async function call(body) {
  let r, j;
  try {
    r = await fetch(conf.url, {method: "POST", headers: {"Content-Type": "text/plain;charset=utf-8"},
                               body: JSON.stringify({...body, date: today(), key: conf.key})});
  } catch (e) { throw new Error("No connection"); }
  try { j = await r.json(); } catch (e) { throw new Error("Google didn't answer"); }
  if (!j.ok) throw new Error(j.error || "Error");
  return j;
}

// ---- state
let orders = [];          // from the sheet
const pending = new Map(); // number -> true/false (done / undo) not confirmed by the sheet yet
const seen = new Set();    // numbers already shown (a new one beeps)
let firstLoad = true;
let lastOk = null, lastErr = null;

const isDone = o => pending.has(o.number) ? pending.get(o.number) : !!o.done;
const timeSec = t => { const m = String(t).match(/(\d+):(\d+)(?::(\d+))?/); return m ? +m[1] * 3600 + +m[2] * 60 + +(m[3] || 0) : 0; };
const nowSec = () => { const d = new Date(); return d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds(); };

// only touch the page when something changed, so a tap is never lost to a redraw
function setHtml(sel, html) {
  const el = $(sel);
  if (el._html !== html) { el.innerHTML = html; el._html = html; }
}

function render() {
  const open = orders.filter(o => !isDone(o) && !(o.void && !seen.has(o.number)))
    .sort((a, b) => timeSec(a.time) - timeSec(b.time) || a.number.localeCompare(b.number));
  const done = orders.filter(o => isDone(o) && !o.void).sort((a, b) => timeSec(b.done || b.time) - timeSec(a.done || a.time)).slice(0, 10);
  const todo = open.filter(o => !o.void).length;
  $("#count").textContent = todo;
  $("#count").classList.toggle("zero", !todo);
  document.title = todo ? `(${todo}) Kuchyně` : "Kuchyně";
  setHtml("#orders", open.map(o => {
    const age = Math.max(0, Math.floor((nowSec() - timeSec(o.time)) / 60));
    const lines = o.lines.map(l => `<li><div class="item">${l.qty !== "" ? `<span class="q">${esc(l.qty)}×</span>` : ""}${esc(l.name)}</div>` +
      (l.tops && l.tops.length ? `<div class="tops">${l.tops.map(t => `<span>${esc(t)}</span>`).join("")}</div>` : (l.alone ? `<div class="plain">only the topping</div>` : `<div class="plain">nothing on it</div>`)) + "</li>").join("");
    return `<div class="order ${o.void ? "void" : ""} ${o.fresh ? "fresh" : ""} ${pending.has(o.number) ? "sending" : ""}">
      ${o.void ? '<div class="voidtag">CANCELLED - don\'t make it</div>' : ""}
      <div class="oh"><b>${esc(o.number)}</b><span>${esc(String(o.time).slice(0, 5))}</span>
        <span class="age ${age >= 10 ? "late" : ""}">${age} min</span></div>
      <ul>${lines}</ul>
      <button class="done-btn" data-done="${esc(o.number)}">${o.void ? "OK, hide it" : "Done ✓"}</button></div>`;
  }).join("") || (conf.key ? '<div class="empty">No orders waiting 👌</div>' : ""));
  setHtml("#doneBox", done.length ? "<h2>Done</h2>" + done.map(o => `<div class="doneRow">
      <div class="t"><b>${esc(o.number)}</b> · ${esc(o.lines.map(l => (l.qty !== "" ? l.qty + "× " : "") + l.name).join(", "))}
        <br><small>done ${esc(String(o.done || "").slice(0, 5))}</small></div>
      <button data-undo="${esc(o.number)}">Undo</button></div>`).join("") : "");
  document.querySelectorAll("[data-done]").forEach(b => b.onclick = () => setDone(b.dataset.done, true));
  document.querySelectorAll("[data-undo]").forEach(b => b.onclick = () => setDone(b.dataset.undo, false));
  orders.forEach(o => { o.fresh = false; });
  renderStatus();
}

function renderStatus() {
  const s = $("#status");
  s.classList.toggle("err", !!lastErr);
  s.textContent = !conf.key ? "" : lastErr ? lastErr + (lastOk ? " · last " + lastOk : "") : lastOk ? "✓ " + lastOk : "loading…";
}

// ---- new orders: beep + vibrate (sound works after the Start tap)
let audio = null;
function beep() {
  try { if (navigator.vibrate) navigator.vibrate([200, 100, 200]); } catch (e) {}
  if (!audio) return;
  [0, 0.25].forEach((at, i) => {
    const o = audio.createOscillator(), g = audio.createGain();
    o.frequency.value = i ? 1320 : 880;
    g.gain.setValueAtTime(0.0001, audio.currentTime + at);
    g.gain.exponentialRampToValueAtTime(0.4, audio.currentTime + at + 0.02);
    g.gain.exponentialRampToValueAtTime(0.0001, audio.currentTime + at + 0.2);
    o.connect(g).connect(audio.destination);
    o.start(audio.currentTime + at); o.stop(audio.currentTime + at + 0.22);
  });
}

function take(list) {
  let fresh = false;
  for (const o of list) {
    // the sheet now says what we sent: the change is confirmed
    if (pending.has(o.number) && pending.get(o.number) === !!o.done) pending.delete(o.number);
    // an order cancelled before the cook saw it never shows; one he saw shows as CANCELLED until "OK"
    if (!seen.has(o.number) && !o.void) {
      if (!firstLoad && !o.done) { o.fresh = true; fresh = true; }
      seen.add(o.number);
    }
  }
  orders = list;
  firstLoad = false;
  if (fresh) beep();
}

// ---- load and tick off
let loading = false;
async function refresh() {
  if (!conf.key || loading || document.hidden) return;
  loading = true;
  try {
    await sendPending();
    const r = await call({action: "orders"});
    take(r.orders || []);
    lastOk = nowTime().slice(0, 5); lastErr = null;
  } catch (e) {
    lastErr = e.message;
    if (/key/i.test(e.message)) { conf.key = ""; store("viperKitchenKey", null); showSetup(); }
  } finally {
    loading = false;
    render();
  }
}

async function sendPending() {
  for (const [number, done] of [...pending]) {
    const o = orders.find(x => x.number === number);
    if (o && !!o.done === done) { pending.delete(number); continue; }
    await call({action: "done", number, done, time: nowTime()});
  }
}

function setDone(number, done) {
  pending.set(number, done);
  render();
  refresh();
}

// ---- start: unlock sound, keep the screen on
let wake = null;
async function keepAwake() {
  try { if ("wakeLock" in navigator && !document.hidden) wake = await navigator.wakeLock.request("screen"); } catch (e) {}
}
$("#startBtn").onclick = () => {
  try { audio = new (window.AudioContext || window.webkitAudioContext)(); audio.resume(); } catch (e) {}
  $("#start").style.display = "none";
  keepAwake();
  beep();
};
document.addEventListener("visibilitychange", () => { if (!document.hidden) { keepAwake(); refresh(); } });

function showSetup() {
  $("#setup").style.display = "";
  $("#start").style.display = "none";
}
$("#keySave").onclick = () => {
  const k = $("#keyIn").value.trim();
  if (!k) return;
  conf.key = k; store("viperKitchenKey", k); store("viperKitchenUrl", conf.url);
  $("#setup").style.display = "none";
  $("#start").style.display = "";
  refresh();
};

if (!conf.key) showSetup();
refresh();
setInterval(refresh, POLL_MS);
setInterval(render, 30000);  // the "min" counters
// a new version downloads in the background: look for one every few minutes and reload into it (then tap Start again)
if ("serviceWorker" in navigator && location.protocol !== "file:") {
  let firstInstall = !navigator.serviceWorker.controller;  // the first install also "changes" the controller: no reload
  navigator.serviceWorker.register("sw.js").then(reg => setInterval(() => reg.update().catch(() => {}), 3 * 60000)).catch(() => {});
  navigator.serviceWorker.addEventListener("controllerchange", () => { if (firstInstall) firstInstall = false; else location.reload(); });
}
