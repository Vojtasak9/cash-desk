// Cash Desk - runs completely in the browser (no server).
// Data is kept on the device in IndexedDB; "Save Excel file" exports everything to one .xlsx.
"use strict";

const $ = s => document.querySelector(s);
const $$ = s => document.querySelectorAll(s);

// ---------------------------------------------------------------- storage

const DEFAULT_STATE = () => ({
  products: [
    {id: 1, name: "Standard", price: 35, stock: null, active: true},
    {id: 2, name: "Mini", price: 25, stock: null, active: true},
    {id: 3, name: "Hotdog", price: 70, stock: null, active: true},
  ],
  payments: [],      // {receipt, date, time, items, total, method, received, change, status}
  paymentItems: [],  // {receipt, productId, name, qty, unitPrice, lineTotal}
  debts: [],         // cash book: {id, number, date, type (see KINDS), person, amount, note}
  settings: {shop_name: "Jarmark", account_name: "", iban: "", currency: "CZK", start_cash: 0},
  sync: {url: "", key: "", lastSync: null, error: null},  // Google Sheets web app (not exported to Excel)
  outbox: [],        // changes waiting to be sent to Google Sheets
  lastExport: null,  // ISO time of the last Excel export
  changedSinceExport: false,
});

let db = null;
let state = null;

function openDb() {
  return new Promise((resolve, reject) => {
    const r = indexedDB.open("cash-desk", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}

function dbGet(key) {
  return new Promise((resolve, reject) => {
    const r = db.transaction("kv").objectStore("kv").get(key);
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}

function dbPut(key, value) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction("kv", "readwrite");
    tx.objectStore("kv").put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function save(markChanged = true) {
  if (markChanged) state.changedSinceExport = true;
  try {
    await dbPut("state", state);
  } catch (e) {
    toast("Could not save on this device: " + e.message, true);
    throw e;
  }
  renderBackupBanner();
}

// ---------------------------------------------------------------- helpers

class UserError extends Error {}

const round2 = v => Math.round(v * 100) / 100;
const num = (v, def = 0) => { const n = parseFloat(String(v).replace(",", ".")); return isNaN(n) ? def : n; };
const pad = n => String(n).padStart(2, "0");
const today = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const nowTime = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; };
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c]));
const money = v => (v == null || v === "") ? "" :
  Number(v).toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 2}) + " " + (state.settings.currency || "");

function toast(msg, err) {
  const t = $("#toast"); t.textContent = msg; t.className = "toast" + (err ? " err" : ""); t.style.display = "block";
  clearTimeout(t._h); t._h = setTimeout(() => t.style.display = "none", err ? 5000 : 3000);
}

// Show UserErrors as a red message instead of failing silently
function guard(fn) {
  return async (...args) => {
    try { await fn(...args); }
    catch (e) { if (e instanceof UserError) toast(e.message, true); else { console.error(e); toast("Error: " + e.message, true); } }
  };
}

// ---------------------------------------------------------------- QR payment (Czech "QR Platba" / SPAYD)

function ibanValid(iban) {
  const s = iban.slice(4) + iban.slice(0, 4);
  const digits = s.replace(/[A-Z]/g, c => String(c.charCodeAt(0) - 55));
  return BigInt(digits) % 97n === 1n;
}

function toIban(account) {
  const a = (account || "").replace(/\s+/g, "").toUpperCase();
  if (!a) throw new UserError("Set your bank account in Settings first.");
  if (/^[A-Z]{2}\d{2}[A-Z0-9]{10,30}$/.test(a)) {
    if (!ibanValid(a)) throw new UserError("The IBAN is not valid (checksum).");
    return a;
  }
  const m = a.match(/^(?:(\d{1,6})-)?(\d{2,10})\/(\d{4})$/);
  if (!m) throw new UserError("Account must be an IBAN or a Czech account number like 123456789/0800.");
  const bban = m[3] + (m[1] || "0").padStart(6, "0") + m[2].padStart(10, "0");
  const check = 98n - BigInt(bban + "123500") % 97n;  // "CZ00" -> C=12, Z=35, 00
  return "CZ" + String(check).padStart(2, "0") + bban;
}

// SPAYD values must not contain "*"; keep them plain ASCII so every bank app reads them
const spaydText = (s, max) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
  .replace(/[^\x20-\x7e]/g, "").replace(/\*/g, " ").trim().slice(0, max);

function spayd(amount, receipt) {
  const s = state.settings;
  const parts = ["SPD", "1.0", "ACC:" + toIban(s.iban), "AM:" + amount.toFixed(2), "CC:" + (s.currency || "CZK")];
  parts.push("MSG:" + spaydText(`${s.shop_name} ${receipt}`, 60));
  if (s.account_name) parts.push("RN:" + spaydText(s.account_name, 35));
  const vs = receipt.replace(/\D/g, "");
  if (vs) parts.push("X-VS:" + vs.slice(0, 10));
  return parts.join("*");
}

function qrSvg(text) {
  const qr = qrcode(0, "M");
  qr.addData(text);
  qr.make();
  return qr.createSvgTag({cellSize: 6, margin: 2, scalable: true});
}

// ---------------------------------------------------------------- business logic

const receiptNumber = code => parseInt(String(code).replace(/\D/g, ""), 10) || 0;
const formatReceipt = n => "P" + String(n).padStart(3, "0");  // P001, P002, ... as in the Jarmark file
const normalizeReceipt = code => formatReceipt(receiptNumber(code));
const nextReceipt = () => formatReceipt(Math.max(0, ...state.payments.map(p => receiptNumber(p.receipt))) + 1);
const productById = id => state.products.find(p => p.id === +id);

async function createPayment(order, method, received) {
  const lines = [];
  let total = 0;
  for (const [id, qty] of Object.entries(order)) {
    const p = productById(id);
    if (!p || qty <= 0) continue;
    if (p.stock != null && p.stock < qty) throw new UserError(`Not enough "${p.name}" in stock (have ${p.stock}).`);
    const lineTotal = round2(p.price * qty);
    total += lineTotal;
    lines.push({p, qty, lineTotal});
  }
  if (!lines.length) throw new UserError("The order is empty.");
  total = round2(total);
  let change = null;
  if (method === "Cash" && received !== "" && received != null) {
    received = num(received);
    if (received < total) throw new UserError("Received cash is less than the total.");
    change = round2(received - total);
  } else {
    received = null;
  }
  const receipt = nextReceipt();
  for (const {p, qty} of lines) if (p.stock != null) p.stock -= qty;
  state.payments.push({
    receipt, date: today(), time: nowTime(),
    items: lines.map(l => `${l.qty}x ${l.p.name}`).join(", "),
    total, method, received, change, status: "OK",
  });
  for (const {p, qty, lineTotal} of lines) {
    state.paymentItems.push({receipt, productId: p.id, name: p.name, qty, unitPrice: p.price, lineTotal});
  }
  queue(saleOp(state.payments[state.payments.length - 1]));
  await save();
  trySync();
  return {receipt, total, change};
}

async function voidPayment(receipt) {
  const pay = state.payments.find(p => p.receipt === receipt);
  if (!pay || pay.status === "VOID") return;
  pay.status = "VOID";
  for (const it of state.paymentItems.filter(i => i.receipt === receipt)) {
    const p = productById(it.productId);
    if (p && p.stock != null) p.stock += it.qty;
  }
  queue(saleOp(pay));
  await save();
  trySync();
}

function dailyTotals() {
  const days = {};
  for (const p of state.payments) {
    if (p.status === "VOID") continue;
    const d = days[p.date] ||= {date: p.date, receipts: 0, cash: 0, qr: 0};
    d.receipts++;
    if (p.method === "QR") d.qr += p.total; else d.cash += p.total;
  }
  return Object.values(days).sort((a, b) => a.date.localeCompare(b.date))
    .map(d => ({...d, cash: round2(d.cash), qr: round2(d.qr), total: round2(d.cash + d.qr)}));
}

// ---- cash book: money going into or out of the cash box besides sales
// dir +1 = cash comes in, -1 = cash goes out. "cz" is the document type written to the sheets.
const KINDS = {
  LoanIn:       {cz: "Zapůjčka",         en: "Loan into the kasa",     dir: +1, person: true,  text: p => `Půjčeno od: ${p}`},
  LoanInRepaid: {cz: "Vrácení zapůjčky", en: "Kasa repays a loan",     dir: -1, person: true,  text: p => `Vráceno: ${p}`},
  Expense:      {cz: "Vydáno",           en: "Expense / purchase",     dir: -1, person: false, text: p => `Nakoupil(a): ${p}`},
  Lent:         {cz: "Půjčka z kasy",    en: "Lent from the kasa",     dir: -1, person: true,  text: p => `Půjčeno z kasy: ${p}`},
  Repaid:       {cz: "Splátka",          en: "Paid back to the kasa",  dir: +1, person: true,  text: p => `Splaceno od: ${p}`},
};
const kindOf = e => KINDS[e.type] || KINDS.Expense;

// PPD = příjmový pokladní doklad (cash in), VPD = výdajový pokladní doklad (cash out)
const docPrefix = type => (KINDS[type] || KINDS.Expense).dir > 0 ? "PPD" : "VPD";
function nextDocNumber(type) {
  const prefix = docPrefix(type);
  const used = state.debts.filter(d => String(d.number || "").startsWith(prefix)).map(d => receiptNumber(d.number));
  return prefix + String(Math.max(0, ...used) + 1).padStart(3, "0");
}
const entryText = e => [e.person ? kindOf(e).text(e.person) : "", e.note].filter(Boolean).join(" – ") || kindOf(e).cz;

function cashNow() {
  const start = num(state.settings.start_cash);
  const sales = state.payments.filter(p => p.status !== "VOID" && p.method === "Cash").reduce((a, p) => a + p.total, 0);
  const moves = state.debts.reduce((a, e) => a + kindOf(e).dir * e.amount, 0);
  return {start, sales: round2(sales), in: round2(state.debts.filter(e => kindOf(e).dir > 0).reduce((a, e) => a + e.amount, 0)),
          out: round2(state.debts.filter(e => kindOf(e).dir < 0).reduce((a, e) => a + e.amount, 0)),
          total: round2(start + sales + moves)};
}

// Per person: what the kasa still owes them (their loans) and what they still owe the kasa
function debtTotals() {
  const people = {};
  for (const e of state.debts) {
    if (!e.person || e.type === "Expense") continue;
    const p = people[e.person] ||= {person: e.person, kasaOwes: 0, owesKasa: 0};
    if (e.type === "LoanIn") p.kasaOwes += e.amount;
    if (e.type === "LoanInRepaid") p.kasaOwes -= e.amount;
    if (e.type === "Lent") p.owesKasa += e.amount;
    if (e.type === "Repaid") p.owesKasa -= e.amount;
  }
  return Object.values(people).sort((a, b) => a.person.localeCompare(b.person))
    .map(p => ({...p, kasaOwes: round2(p.kasaOwes), owesKasa: round2(p.owesKasa)}));
}

// ---------------------------------------------------------------- Google Sheets sync
// Every change is queued in state.outbox and sent to the Apps Script web app (google-apps-script.js).
// Works offline: the queue is kept on the device and sent when the internet is back.

const opKey = op => op.type === "sale" ? "sale:" + op.receipt : "entry:" + op.number;

function queue(op) {
  if (!state.sync.url) return;
  state.outbox = state.outbox.filter(o => opKey(o) !== opKey(op));  // the newest version wins
  state.outbox.push(op);
}
const saleOp = p => ({type: "sale", receipt: p.receipt, date: p.date, goods: p.items, total: p.total,
                      method: p.method, void: p.status === "VOID"});
const entryOp = e => ({type: "entry", number: e.number, date: e.date, kind: kindOf(e).cz, description: entryText(e),
                       cashIn: kindOf(e).dir > 0 ? e.amount : "", cashOut: kindOf(e).dir < 0 ? e.amount : ""});

async function callSheets(body, sync = state.sync) {
  let r, j;
  try {
    r = await fetch(sync.url, {method: "POST", headers: {"Content-Type": "text/plain;charset=utf-8"},
                               body: JSON.stringify({...body, key: sync.key})});
  } catch (e) { throw new UserError("Can't reach Google - check the internet connection."); }
  try { j = await r.json(); }
  catch (e) { throw new UserError("Google didn't answer as expected - check the web app URL (it ends with /exec) and that access is set to Anyone."); }
  if (!j.ok) throw new UserError(j.error || "Google Sheets returned an error.");
  return j;
}

let syncing = false;
async function trySync() {
  if (syncing || !state.sync.url || !state.outbox.length || !navigator.onLine) return renderSync();
  syncing = true;
  renderSync();
  const ops = state.outbox.slice();
  try {
    await callSheets({action: "sync", ops});
    state.outbox = state.outbox.filter(o => !ops.includes(o));  // changes made meanwhile stay queued
    state.sync.lastSync = new Date().toISOString();
    state.sync.error = null;
  } catch (e) {
    state.sync.error = e.message;
  } finally {
    syncing = false;
    await save(false);
    renderSync();
  }
  if (!state.sync.error && state.outbox.length) trySync();
}

// Names and prices come from Kasa / Produkty; stock and "on sale" stay as set in the app
async function loadProductsFromSheet(sync = state.sync) {
  const {products} = await callSheets({action: "products"}, sync);
  if (!products.length) throw new UserError('No products found in Kasa → sheet "Produkty".');
  let nextId = Math.max(0, ...state.products.map(p => p.id)) + 1;
  state.products = products.map(sp => {
    const old = state.products.find(p => p.name.toLowerCase() === sp.name.toLowerCase());
    return old ? {...old, name: sp.name, price: sp.price} : {id: nextId++, name: sp.name, price: sp.price, stock: null, active: true};
  });
  await save(false);
  return products.length;
}

// ---------------------------------------------------------------- Excel export / import
// Layout follows the school "Jarmark" workbook: Czech sheet names, "Heading 1" headers, amounts in Kč.

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DAY_NAMES = ["Neděle", "Pondělí", "Úterý", "Středa", "Čtvrtek", "Pátek", "Sobota"];
const SH = {
  sales: "Prodej", lines: "Evidence Prodeje", products: "Evidence Pokladny",
  oldReceipts: "Účtenky",  // only in files saved before it was merged into Evidence Prodeje
  days: "Denní uzávěrka", book: "Pokladní kniha", debtors: "Kdo dluží", settings: "Nastavení",
  oldDebts: "Dluhy",  // files saved before the cash book
};
const METHOD_CZ = {Cash: "Hotově", QR: "QR"};

const moneyFmt = () => (state.settings.currency || "CZK") === "CZK"
  ? '#,##0.00 "Kč"' : `#,##0.00 "${state.settings.currency}"`;

// Excel's "Heading 1" look used in the Jarmark file
function heading(row) {
  row.eachCell(c => {
    c.font = {bold: true, size: 15, color: {argb: "FF44546A"}};
    c.border = {bottom: {style: "thick", color: {argb: "FF4472C4"}}};
  });
}
const bold = row => row.eachCell(c => { c.font = {bold: true}; });

function addSheet(wb, name, headers, rows, {style = bold, widths = {}, money = []} = {}) {
  const ws = wb.addWorksheet(name);
  style(ws.addRow(headers));
  rows.forEach(r => ws.addRow(r));
  headers.forEach((h, i) => {
    const col = ws.getColumn(i + 1);
    col.width = widths[h] || Math.max(10, String(h).length + 3);
    if (money.includes(h)) col.numFmt = moneyFmt();
  });
  ws.views = [{state: "frozen", ySplit: 1}];
  return ws;
}

function dayLabels(dates) {
  const count = {};
  dates.forEach(d => { const n = new Date(d + "T12:00").getDay(); count[n] = (count[n] || 0) + 1; });
  return dates.map(d => {
    const dt = new Date(d + "T12:00");
    const name = DAY_NAMES[dt.getDay()];
    return (count[dt.getDay()] > 1 ? `${name} ${dt.getDate()}. ${dt.getMonth() + 1}.` : name) + " (ks)";
  });
}

function buildWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Cash Desk";
  const ok = new Set(state.payments.filter(p => p.status !== "VOID").map(p => p.receipt));

  // 1) Prodej - like "Prodej Propisek": pieces per product per sale day, with formulas
  const dates = [...new Set(state.payments.filter(p => ok.has(p.receipt)).map(p => p.date))].sort();
  const sold = {};  // name -> date -> qty
  const dateOf = Object.fromEntries(state.payments.map(p => [p.receipt, p.date]));
  for (const it of state.paymentItems) {
    if (!ok.has(it.receipt)) continue;
    const byDay = sold[it.name] ||= {};
    byDay[dateOf[it.receipt]] = (byDay[dateOf[it.receipt]] || 0) + it.qty;
  }
  const names = [...new Set([...state.products.map(p => p.name), ...Object.keys(sold)])];
  const priceOf = name => {
    const p = state.products.find(x => x.name === name);
    if (p) return p.price;
    const it = state.paymentItems.find(i => i.name === name);  // product was deleted since
    return it ? it.unitPrice : 0;
  };
  const ws = wb.addWorksheet(SH.sales);
  const nDays = dates.length;
  heading(ws.addRow(["Druh zboží", "Cena za kus", ...dayLabels(dates), "Celkem (ks)", "Tržba za daný druh zboží", "Tržba celkem"]));
  const col = i => ws.getColumn(i).letter;
  const totalCol = 3 + nDays, revCol = totalCol + 1, grandCol = revCol + 1;
  let grand = 0;
  names.forEach((name, i) => {
    const r = i + 2, price = priceOf(name);
    const qtys = dates.map(d => (sold[name] || {})[d] || 0);
    const qty = qtys.reduce((a, b) => a + b, 0);
    grand += price * qty;
    ws.addRow([name, price, ...qtys,
      {formula: nDays ? `SUM(${col(3)}${r}:${col(2 + nDays)}${r})` : "0", result: qty},
      {formula: `B${r}*${col(totalCol)}${r}`, result: round2(price * qty)}]);
  });
  if (names.length) ws.getCell(2, grandCol).value = {formula: `SUM(${col(revCol)}2:${col(revCol)}${names.length + 1})`, result: round2(grand)};
  ws.getColumn(1).width = 18;
  ws.getColumn(2).width = 14;
  ws.getColumn(2).numFmt = moneyFmt();
  for (let c = 3; c <= totalCol; c++) ws.getColumn(c).width = Math.max(14, String(ws.getCell(1, c).value).length + 4);
  ws.getColumn(revCol).width = 30; ws.getColumn(revCol).numFmt = moneyFmt();
  ws.getColumn(grandCol).width = 18; ws.getColumn(grandCol).numFmt = moneyFmt();

  // 2) Evidence Prodeje - one row per product sold, with the payment details of its receipt
  // (Číslo / Produkt / Počet as in the Jarmark file). Received and change go on the first row
  // of a receipt only, so summing those columns doesn't count them twice.
  addSheet(wb, SH.lines, ["Číslo", "Produkt", "Počet", "Cena za kus", "Celkem", "Datum", "Čas", "Platba", "Přijato", "Vráceno", "Stav"],
    state.payments.flatMap(p => state.paymentItems.filter(i => i.receipt === p.receipt).map((i, n) => [
      i.receipt, i.name, i.qty, i.unitPrice, i.lineTotal, p.date, p.time, METHOD_CZ[p.method] || p.method,
      n === 0 ? p.received ?? "" : "", n === 0 ? p.change ?? "" : "", p.status === "VOID" ? "storno" : "OK",
    ])),
    {style: () => {}, widths: {Produkt: 18, Datum: 12}, money: ["Cena za kus", "Celkem", "Přijato", "Vráceno"]});

  // 3) Evidence Pokladny - product list with price and stock
  addSheet(wb, SH.products, ["Produkty:", "Cena za kus", "Skladem", "V prodeji"],
    state.products.map(p => [p.name, p.price, p.stock ?? "", p.active ? "ano" : "ne"]),
    {style: () => {}, widths: {"Produkty:": 18}, money: ["Cena za kus"]});

  // 4) Denní uzávěrka - totals per day with a sum row
  const days = dailyTotals();
  const dws = addSheet(wb, SH.days, ["Datum", "Den", "Účtenek", "Hotově", "QR", "Celkem"],
    days.map(d => [d.date, DAY_NAMES[new Date(d.date + "T12:00").getDay()], d.receipts, d.cash, d.qr, d.total]),
    {style: heading, widths: {Datum: 12, Den: 12}, money: ["Hotově", "QR", "Celkem"]});
  if (days.length) {
    const n = days.length + 1;
    const sum = (c, v) => ({formula: `SUM(${c}2:${c}${n})`, result: v});
    const tot = days.reduce((a, d) => ({r: a.r + d.receipts, c: a.c + d.cash, q: a.q + d.qr, t: a.t + d.total}), {r: 0, c: 0, q: 0, t: 0});
    bold(dws.addRow(["Celkem", "", sum("C", tot.r), sum("D", round2(tot.c)), sum("E", round2(tot.q)), sum("F", round2(tot.t))]));
  }

  // 5) Pokladní kniha (loans, expenses) + Kdo dluží
  addSheet(wb, SH.book, ["Datum", "Číslo", "Doklad", "Osoba", "Popis", "Příjem", "Výdaj", "Poznámka"],
    state.debts.map(e => [e.date, e.number, kindOf(e).cz, e.person || "", entryText(e),
                          kindOf(e).dir > 0 ? e.amount : "", kindOf(e).dir < 0 ? e.amount : "", e.note || ""]),
    {style: heading, widths: {Datum: 12, Doklad: 18, Osoba: 16, Popis: 34, Poznámka: 24}, money: ["Příjem", "Výdaj"]});
  addSheet(wb, SH.debtors, ["Osoba", "Kasa mu/jí dluží", "Dluží kase"],
    debtTotals().map(d => [d.person, d.kasaOwes, d.owesKasa]),
    {style: heading, widths: {Osoba: 18}, money: ["Kasa mu/jí dluží", "Dluží kase"]});

  // 6) Nastavení - needed to load the file back into the app
  addSheet(wb, SH.settings, ["Klíč", "Hodnota"], Object.entries(state.settings), {widths: {Hodnota: 30}});
  return wb;
}

const exportName = () => `${(state.settings.shop_name || "Pokladna").replace(/[\\/:*?"<>|]/g, "")}_${today()}.xlsx`;
let exportFile = null;

async function prepareExport() {
  exportFile = null;
  $$("#exportShare, #exportDownload").forEach(b => b.disabled = true);
  $("#exportName").textContent = "Preparing…";
  const buf = await buildWorkbook().xlsx.writeBuffer();
  exportFile = new File([buf], exportName(), {type: XLSX_TYPE});
  $("#exportName").textContent = exportFile.name;
  $$("#exportShare, #exportDownload").forEach(b => b.disabled = false);
}

async function markExported() {
  state.lastExport = new Date().toISOString();
  state.changedSinceExport = false;
  await save(false);
  renderSettings();
}

function exportDownload() {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(exportFile);
  a.download = exportFile.name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 10000);
  markExported();
}

// The file is built before the tap, so the iPad share sheet opens straight from the tap
async function exportShare() {
  if (navigator.canShare && navigator.canShare({files: [exportFile]})) {
    try {
      await navigator.share({files: [exportFile], title: exportFile.name});
      await markExported();
    } catch (e) {
      if (e.name !== "AbortError") toast("Sharing failed - use Download instead", true);
    }
  } else {
    exportDownload();
  }
}

// ---- import

function cellValue(v) {
  if (v && typeof v === "object") {
    if (v instanceof Date) return v;
    if ("result" in v) return v.result;
    if ("richText" in v) return v.richText.map(t => t.text).join("");
    if ("text" in v) return v.text;
  }
  return v ?? "";
}
function cellDate(v) {
  if (v instanceof Date) return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
  return String(v ?? "");
}
function cellTime(v) {
  if (v instanceof Date) return `${pad(v.getUTCHours())}:${pad(v.getUTCMinutes())}:${pad(v.getUTCSeconds())}`;
  return String(v ?? "");
}
const blankToNull = v => (v === "" || v == null) ? null : num(v);

function sheetRows(wb, name, required = true) {
  const ws = wb.getWorksheet(name);
  if (!ws) {
    if (required) throw new UserError(`This isn't a Cash Desk Excel file (sheet "${name}" is missing).`);
    return [];
  }
  const headers = [];
  ws.getRow(1).eachCell((c, i) => { headers[i] = String(cellValue(c.value)).trim(); });
  const out = [];
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const r = {};
    headers.forEach((h, i) => { if (h) r[h] = cellValue(row.getCell(i).value); });
    if (Object.values(r).some(v => v !== "")) out.push(r);
  });
  return out;
}

async function importExcel(file) {
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(await file.arrayBuffer()); }
  catch (e) { throw new UserError("Can't read this file - is it an .xlsx Excel file?"); }
  const next = DEFAULT_STATE();
  next.products = sheetRows(wb, SH.products).filter(r => r["Produkty:"] !== "").map((r, i) => ({
    id: i + 1, name: String(r["Produkty:"]).trim(), price: num(r["Cena za kus"]),
    stock: r.Skladem === "" || r.Skladem == null ? null : Math.round(num(r.Skladem)),
    active: !["ne", "no", "false", "0"].includes(String(r["V prodeji"]).toLowerCase()),
  }));
  const idOf = name => (next.products.find(p => p.name === name) || {}).id ?? null;
  const fromCz = m => m === "QR" ? "QR" : "Cash";
  const lines = sheetRows(wb, SH.lines).filter(r => r["Číslo"] !== "");
  next.paymentItems = lines.map(r => ({
    receipt: normalizeReceipt(r["Číslo"]), productId: idOf(String(r.Produkt)), name: String(r.Produkt),
    qty: num(r["Počet"]), unitPrice: num(r["Cena za kus"]), lineTotal: num(r.Celkem),
  }));
  const isVoid = r => String(r.Stav).toLowerCase() === "storno";
  if (wb.getWorksheet(SH.oldReceipts)) {
    next.payments = sheetRows(wb, SH.oldReceipts).filter(r => r["Číslo"] !== "").map(r => ({
      receipt: normalizeReceipt(r["Číslo"]), date: cellDate(r.Datum), time: cellTime(r["Čas"]),
      items: String(r["Položky"]), total: num(r["Částka"]), method: fromCz(r.Platba),
      received: blankToNull(r["Přijato"]), change: blankToNull(r["Vráceno"]), status: isVoid(r) ? "VOID" : "OK",
    }));
  } else {
    const receipts = new Map();
    for (const r of lines) {
      const receipt = normalizeReceipt(r["Číslo"]);
      if (!receipts.has(receipt)) receipts.set(receipt, {
        receipt, date: cellDate(r.Datum), time: cellTime(r["Čas"]), items: "", total: 0, method: fromCz(r.Platba),
        received: blankToNull(r["Přijato"]), change: blankToNull(r["Vráceno"]), status: isVoid(r) ? "VOID" : "OK",
      });
      const p = receipts.get(receipt);
      p.items += (p.items ? ", " : "") + `${num(r["Počet"])}x ${r.Produkt}`;
      p.total = round2(p.total + num(r.Celkem));
    }
    next.payments = [...receipts.values()];
  }
  const typeOf = cz => Object.keys(KINDS).find(k => KINDS[k].cz === cz) || "Expense";
  if (wb.getWorksheet(SH.book)) {
    next.debts = sheetRows(wb, SH.book).filter(r => r["Číslo"] !== "").map((r, i) => ({
      id: i + 1, number: String(r["Číslo"]), date: cellDate(r.Datum), type: typeOf(r.Doklad), person: String(r.Osoba || ""),
      amount: num(r["Příjem"]) || num(r["Výdaj"]), note: String(r["Poznámka"] || ""),
    }));
  } else {
    next.debts = sheetRows(wb, SH.oldDebts, false).filter(r => r.Osoba !== "").map((r, i) => ({
      id: i + 1, number: "", date: cellDate(r.Datum), person: String(r.Osoba),
      type: r.Typ === "Vráceno" ? "Repaid" : "Lent", amount: num(r["Částka"]), note: String(r["Poznámka"] || ""),
    }));
  }
  for (const r of sheetRows(wb, SH.settings, false)) {
    if (r["Klíč"] in next.settings) next.settings[r["Klíč"]] = r["Klíč"] === "start_cash" ? num(r.Hodnota) : String(r.Hodnota);
  }
  if (!confirm(`Replace all data on this iPad with this file?\n\n${next.products.length} products, ` +
               `${next.payments.length} receipts, ${next.debts.length} debt entries.`)) return;
  next.sync = state.sync;      // keep the Google Sheets connection and anything not sent yet
  next.outbox = state.outbox;
  state = next;
  migrate();
  state.lastExport = new Date().toISOString();
  await save(false);
  order = {};
  renderAll();
  toast("Data loaded from Excel");
}

// ---------------------------------------------------------------- UI

let order = {};
let editingId = null;

$$("nav button").forEach(b => b.onclick = () => {
  $$("nav button, .tab").forEach(x => x.classList.remove("active"));
  b.classList.add("active");
  $("#" + b.dataset.tab).classList.add("active");
  if (b.dataset.tab === "receipts") renderPayments();
  if (b.dataset.tab === "book") renderBook();
  if (b.dataset.tab === "settings") renderSettings();
});
$$("[data-close]").forEach(b => b.onclick = () => b.closest(".modal").classList.remove("open"));
const closeModals = () => $$(".modal").forEach(m => m.classList.remove("open"));

// ---- register
function renderGrid() {
  $("#prodGrid").innerHTML = state.products.filter(p => p.active).map(p => {
    const left = p.stock == null ? null : p.stock - (order[p.id] || 0);
    return `<button class="prod ${left !== null && left <= 0 ? "out" : ""}" data-id="${p.id}">
      <b>${esc(p.name)}</b>${money(p.price)}<br><small>${left === null ? "" : "stock: " + left}</small></button>`;
  }).join("") || '<p class="muted">No products yet - add some in "Products & stock".</p>';
  $$(".prod").forEach(b => b.onclick = () => changeQty(+b.dataset.id, 1));
}

function changeQty(id, d) {
  const p = productById(id);
  const q = (order[id] || 0) + d;
  if (p.stock != null && q > p.stock) return toast(`Only ${p.stock} × ${p.name} in stock`, true);
  if (q <= 0) delete order[id]; else order[id] = q;
  renderOrder();
}

const orderTotal = () => round2(Object.entries(order).reduce((s, [id, q]) => s + (productById(id)?.price || 0) * q, 0));

function renderOrder() {
  for (const id of Object.keys(order)) if (!productById(id)) delete order[id];
  $("#nextReceipt").textContent = nextReceipt();
  $("#orderLines").innerHTML = Object.entries(order).map(([id, q]) => {
    const p = productById(id);
    return `<tr><td>${esc(p.name)}</td>
      <td class="qty"><button data-d="-1" data-id="${id}" aria-label="less">−</button> <b>${q}</b> <button data-d="1" data-id="${id}" aria-label="more">+</button></td>
      <td class="num">${money(p.price * q)}</td></tr>`;
  }).join("") || '<tr><td class="muted">Tap a product to add it</td></tr>';
  $$(".qty button").forEach(b => b.onclick = () => changeQty(+b.dataset.id, +b.dataset.d));
  $("#orderTotal").textContent = money(orderTotal());
  renderGrid();
}

$("#clearOrder").onclick = () => { order = {}; renderOrder(); };

$("#payQr").onclick = guard(() => {
  const total = orderTotal();
  if (!total) throw new UserError("The order is empty");
  const missing = !state.settings.iban;
  $("#qrAmount").textContent = money(total);
  $("#qrMissing").style.display = missing ? "" : "none";
  $("#qrConfirm").style.display = missing ? "none" : "";
  $("#qrBox").innerHTML = missing ? "" : qrSvg(spayd(total, nextReceipt()));
  $("#qrNote").textContent = missing ? "" : `${nextReceipt()} · ${state.settings.iban}`;
  $("#qrModal").classList.add("open");
});
$("#qrToSettings").onclick = () => { closeModals(); $("[data-tab=settings]").click(); };
$("#qrConfirm").onclick = () => pay("QR");

$("#payCash").onclick = guard(() => {
  if (!orderTotal()) throw new UserError("The order is empty");
  $("#cashAmount").textContent = money(orderTotal());
  $("#cashReceived").value = "";
  $("#cashChange").textContent = "-";
  $("#cashModal").classList.add("open");
  $("#cashReceived").focus();
});
$("#cashReceived").oninput = e => {
  const ch = num(e.target.value) - orderTotal();
  $("#cashChange").textContent = e.target.value === "" ? "-" : ch < 0 ? "not enough" : money(ch);
};
$("#cashConfirm").onclick = () => pay("Cash", $("#cashReceived").value);

let paying = false;
const pay = guard(async (method, received) => {
  if (paying) return;  // a double tap must not save two receipts
  paying = true;
  $$("#qrConfirm, #cashConfirm").forEach(b => b.disabled = true);
  try {
    const r = await createPayment(order, method, received);
    closeModals();
    order = {};
    renderOrder(); renderProducts(); renderCash();
    toast(`${r.receipt} saved · ${money(r.total)}` + (r.change != null ? ` · change ${money(r.change)}` : ""));
  } finally {
    paying = false;
    $$("#qrConfirm, #cashConfirm").forEach(b => b.disabled = false);
  }
});

// ---- receipts / day total
function renderPayments() {
  const day = $("#day").value;
  const s = dailyTotals().find(d => d.date === day) || {receipts: 0, cash: 0, qr: 0, total: 0};
  $("#sCount").textContent = s.receipts;
  $("#sCash").textContent = money(s.cash);
  $("#sQr").textContent = money(s.qr);
  $("#sTotal").textContent = money(s.total);
  const list = state.payments.filter(p => p.date === day).reverse();
  $("#payRows").innerHTML = list.map(p => `<tr class="${p.status === "VOID" ? "void" : ""}">
    <td><b>${esc(p.receipt)}</b></td><td>${esc(p.time)}</td><td>${esc(p.items)}</td><td>${esc(p.method)}</td>
    <td class="num">${money(p.total)}</td>
    <td>${p.status === "VOID" ? "voided" : `<button class="link bad" data-void="${esc(p.receipt)}">void</button>`}</td></tr>`
  ).join("") || '<tr><td colspan="6" class="muted">No receipts this day</td></tr>';
  $$("[data-void]").forEach(b => b.onclick = guard(async () => {
    if (!confirm(`Void ${b.dataset.void}? Its items go back to stock.`)) return;
    await voidPayment(b.dataset.void);
    renderOrder(); renderProducts(); renderPayments(); renderCash();
  }));
}
$("#day").onchange = renderPayments;

// ---- products
function renderProducts() {
  $("#prodRows").innerHTML = state.products.map(p => `<tr>
    <td>${p.id}</td><td>${esc(p.name)}</td><td class="num">${money(p.price)}</td>
    <td class="num qty"><button data-sd="-1" data-sid="${p.id}" aria-label="one less">−</button>
      <input class="stock-in" data-sid="${p.id}" type="number" inputmode="numeric" min="0" step="1" value="${p.stock ?? ""}" placeholder="—" aria-label="stock of ${esc(p.name)}">
      <button data-sd="1" data-sid="${p.id}" aria-label="one more">+</button></td><td>${p.active ? "yes" : "no"}</td>
    <td><button class="link" data-edit="${p.id}">edit</button><button class="link bad" data-del="${p.id}">delete</button></td></tr>`).join("");
  // Stock is edited right in the list: type a number (empty = not counted) or tap − / +
  const setStock = guard(async (id, value) => {
    const p = productById(id);
    p.stock = value == null ? null : Math.max(0, Math.round(value));
    await save();
    $(`.stock-in[data-sid="${id}"]`).value = p.stock ?? "";
    renderGrid();
  });
  $$(".stock-in").forEach(i => {
    i.onchange = () => setStock(+i.dataset.sid, i.value.trim() === "" ? null : num(i.value));
    i.onkeydown = e => { if (e.key === "Enter") i.blur(); };
  });
  $$("[data-sd]").forEach(b => b.onclick = () => {
    const p = productById(b.dataset.sid);
    setStock(p.id, (p.stock ?? 0) + +b.dataset.sd);
  });
  $$("[data-edit]").forEach(b => b.onclick = () => {
    const p = productById(b.dataset.edit);
    editingId = p.id;
    $("#prodFormTitle").textContent = "Edit " + p.name;
    $("#pName").value = p.name; $("#pPrice").value = p.price; $("#pStock").value = p.stock ?? "";
    $("#pActive").checked = p.active; $("#pCancel").style.display = "";
    $("#pName").scrollIntoView({block: "center"});
  });
  $$("[data-del]").forEach(b => b.onclick = guard(async () => {
    if (!confirm("Delete this product?")) return;
    state.products = state.products.filter(p => p.id !== +b.dataset.del);
    await save();
    renderProducts(); renderOrder();
  }));
}
function resetProductForm() {
  editingId = null;
  $("#prodFormTitle").textContent = "Add product";
  ["#pName", "#pPrice", "#pStock"].forEach(s => $(s).value = "");
  $("#pActive").checked = true;
  $("#pCancel").style.display = "none";
}
$("#pCancel").onclick = resetProductForm;
$("#pSave").onclick = guard(async () => {
  const name = $("#pName").value.trim();
  if (!name) throw new UserError("Product name is required.");
  const price = num($("#pPrice").value, NaN);
  if (isNaN(price) || price < 0) throw new UserError("Price must be a number (0 or more).");
  const stockRaw = $("#pStock").value.trim();
  const fields = {name, price, stock: stockRaw === "" ? null : Math.round(num(stockRaw)), active: $("#pActive").checked};
  if (editingId) Object.assign(productById(editingId), fields);
  else state.products.push({id: Math.max(0, ...state.products.map(p => p.id)) + 1, ...fields});
  await save();
  resetProductForm(); renderProducts(); renderOrder();
  toast("Product saved");
});

// ---- cash book + current cash
function renderCash() {
  const c = cashNow();
  $("#cashPill").textContent = "Cash: " + money(c.total);
  $("#cashNow").textContent = money(c.total);
  $("#cashParts").textContent = `Starting cash ${money(c.start)} + cash sales ${money(c.sales)} + money in ${money(c.in)} − money out ${money(c.out)}`;
}

function renderBook() {
  renderCash();
  const totals = debtTotals();
  $("#debtSummary").innerHTML = totals.map(s => `<tr><td>${esc(s.person)}</td>
    <td class="num ${s.kasaOwes > 0 ? "owes" : "settled"}">${money(s.kasaOwes)}</td>
    <td class="num ${s.owesKasa > 0 ? "owes" : "settled"}">${money(s.owesKasa)}</td></tr>`).join("")
    || '<tr><td colspan="3" class="muted">No loans yet</td></tr>';
  const waiting = new Set(state.outbox.map(opKey));
  $("#debtRows").innerHTML = state.debts.slice().reverse().map(e => `<tr><td>${esc(e.date)}</td><td><b>${esc(e.number)}</b></td>
    <td>${esc(kindOf(e).en)}<br><small class="muted">${esc(kindOf(e).cz)}</small></td><td>${esc(entryText(e))}</td>
    <td class="num">${kindOf(e).dir > 0 ? money(e.amount) : ""}</td><td class="num">${kindOf(e).dir < 0 ? money(e.amount) : ""}</td>
    <td>${waiting.has("entry:" + e.number) ? '<small class="muted">not sent yet</small>' : ""}</td>
    <td><button class="link bad" data-ddel="${e.id}">delete</button></td></tr>`).join("")
    || '<tr><td colspan="8" class="muted">Nothing yet</td></tr>';
  $("#peopleList").innerHTML = totals.map(s => `<option value="${esc(s.person)}">`).join("");
  $$("[data-ddel]").forEach(b => b.onclick = guard(async () => {
    const e = state.debts.find(d => d.id === +b.dataset.ddel);
    if (!confirm(`Delete ${e.number}? It is also removed from Google Sheets.`)) return;
    state.debts = state.debts.filter(d => d !== e);
    queue({type: "deleteEntry", number: e.number});
    await save();
    renderBook();
    trySync();
  }));
}
$("#dType").onchange = () => {
  const k = KINDS[$("#dType").value];
  $("#dPerson").placeholder = k.person ? "Person" : "Who bought it (optional)";
  $("#dNote").placeholder = k.person ? "Note (optional)" : "What was bought (e.g. Koupě paragonů)";
};
$("#dSave").onclick = guard(async () => {
  const type = $("#dType").value, k = KINDS[type];
  const person = $("#dPerson").value.trim(), note = $("#dNote").value.trim();
  if (k.person && !person) throw new UserError("Person is required.");
  if (!k.person && !note) throw new UserError("Write what was bought.");
  const amount = num($("#dAmount").value, NaN);
  if (isNaN(amount) || amount <= 0) throw new UserError("Amount must be more than 0.");
  if (k.dir < 0 && amount > cashNow().total) {
    if (!confirm(`There is only ${money(cashNow().total)} in the cash box. Save anyway?`)) return;
  }
  const e = {id: Math.max(0, ...state.debts.map(d => d.id)) + 1, number: nextDocNumber(type), date: today(),
             type, person, amount, note};
  state.debts.push(e);
  queue(entryOp(e));
  await save();
  ["#dPerson", "#dAmount", "#dNote"].forEach(s => $(s).value = "");
  renderBook();
  toast(`${e.number} saved`);
  trySync();
});

// ---- Google Sheets status
function renderSync() {
  const s = state.sync, n = state.outbox.length;
  let pill = "", text;
  if (!s.url) text = "Not connected.";
  else if (syncing) { pill = "Sheets: sending…"; text = "Sending…"; }
  else if (s.error) { pill = `Sheets: ${n || ""} not sent ⚠`; text = "Last try failed: " + s.error; }
  else if (n) { pill = `Sheets: ${n} waiting`; text = `${n} change(s) waiting` + (navigator.onLine ? "" : " - no internet, they are sent when it's back") + "."; }
  else { pill = "Sheets ✓"; text = "Everything is in Google Sheets" + (s.lastSync ? " (last sent " + new Date(s.lastSync).toLocaleTimeString() + ")." : "."); }
  $("#syncPill").textContent = pill;
  $("#syncPill").style.display = pill ? "" : "none";
  $("#syncPill").classList.toggle("warn", !!s.error);
  $("#syncStatus").textContent = text;
}
$("#syncPill").onclick = () => $("[data-tab=settings]").click();

$("#gsSave").onclick = guard(async () => {
  const url = $("#gsUrl").value.trim(), key = $("#gsKey").value.trim();
  if (url && !/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
    throw new UserError("The web app URL starts with https://script.google.com/ and ends with /exec.");
  }
  if (url && !key) throw new UserError("Type the key from the script too.");
  const firstTime = url && !state.sync.url;
  if (url) {
    $("#gsSave").disabled = true;
    try {
      const n = await loadProductsFromSheet({url, key});
      toast(`Connected - ${n} products loaded from Kasa`);
    } finally { $("#gsSave").disabled = false; }
  }
  state.sync = {...state.sync, url, key, error: null};
  if (firstTime && (state.payments.length || state.debts.length) &&
      confirm("Also send the receipts and cash book entries that are already in the app?")) {
    state.payments.forEach(p => queue(saleOp(p)));
    state.debts.forEach(e => queue(entryOp(e)));
  }
  if (!url) state.outbox = [];
  await save(false);
  renderAll();
  trySync();
});
$("#gsSend").onclick = guard(async () => {
  if (!state.sync.url) throw new UserError("Connect Google Sheets first.");
  await trySync();
  if (state.sync.error) throw new UserError(state.sync.error);
  toast(state.outbox.length ? "Some changes are still waiting" : "Everything is in Google Sheets");
});
$("#gsProducts").onclick = guard(async () => {
  if (!state.sync.url) throw new UserError("Connect Google Sheets first.");
  const n = await loadProductsFromSheet();
  renderAll();
  toast(`${n} products loaded from Kasa`);
});

// ---- settings + Excel
function renderSettings() {
  const s = state.settings;
  $("#shopName").textContent = s.shop_name || "Cash Desk";
  $("#sShop").value = s.shop_name; $("#sName").value = s.account_name;
  $("#sIban").value = s.iban; $("#sCur").value = s.currency; $("#sStart").value = s.start_cash || 0;
  $("#gsUrl").value = state.sync.url; $("#gsKey").value = state.sync.key;
  $("#productsNote").style.display = state.sync.url ? "" : "none";
  renderSync();
  $("#lastExport").textContent = state.lastExport
    ? "Last Excel file saved: " + new Date(state.lastExport).toLocaleString()
    : "No Excel file saved yet.";
  renderBackupBanner();
}
$("#sSave").onclick = guard(async () => {
  const iban = $("#sIban").value.trim();
  state.settings = {
    shop_name: $("#sShop").value.trim(), account_name: $("#sName").value.trim(),
    iban: iban ? toIban(iban) : "", currency: $("#sCur").value.trim().toUpperCase() || "CZK",
    start_cash: num($("#sStart").value),
  };
  await save();
  renderAll();
  toast("Settings saved");
});

function renderBackupBanner() {
  const lastDay = state.lastExport ? state.lastExport.slice(0, 10) : null;
  const unsavedSales = state.changedSinceExport && state.payments.some(p => !lastDay || p.date >= lastDay);
  $("#backupBanner").classList.toggle("show", unsavedSales && (!lastDay || lastDay < today() || hoursSinceExport() >= 4));
}
const hoursSinceExport = () => state.lastExport ? (Date.now() - new Date(state.lastExport)) / 36e5 : Infinity;

$$("[data-export]").forEach(b => b.onclick = guard(async () => {
  $("#exportModal").classList.add("open");
  await prepareExport();
}));
$("#exportShare").onclick = guard(async () => { if (exportFile) { await exportShare(); closeModals(); } });
$("#exportDownload").onclick = guard(() => { if (exportFile) { exportDownload(); closeModals(); } });
$("#importBtn").onclick = () => $("#importFile").click();
$("#importFile").onchange = guard(async e => {
  const f = e.target.files[0];
  e.target.value = "";
  if (f) await importExcel(f);
});

async function showStorageInfo() {
  let text = "Data is stored in this browser on this device.";
  if (navigator.storage && navigator.storage.persist) {
    const persisted = await navigator.storage.persisted() || await navigator.storage.persist();
    text += persisted ? " Storage is protected from automatic clean-up."
                      : " Tip: add this app to the Home Screen so iPadOS keeps the data.";
  }
  $("#storageInfo").textContent = text;
}

function renderAll() {
  renderSettings(); renderOrder(); renderProducts(); renderBook(); renderPayments();
}

// Fill in fields that older versions of the app didn't have
function migrate() {
  const d = DEFAULT_STATE();
  state.settings = {...d.settings, ...state.settings};
  state.sync ||= d.sync;
  state.outbox ||= [];
  for (const e of state.debts) {
    if (!KINDS[e.type]) e.type = "Expense";
    if (!e.number) e.number = nextDocNumber(e.type);
  }
}

// ---------------------------------------------------------------- start

(async () => {
  try {
    db = await openDb();
    state = await dbGet("state");
    if (!state) state = DEFAULT_STATE();
    migrate();
    await save(false);
  } catch (e) {
    document.body.innerHTML = `<p style="padding:24px;font-size:18px">This browser can't store data (${esc(e.message)}).
      In Safari, turn off Private Browsing and open the app again.</p>`;
    return;
  }
  $("#day").value = today();
  renderAll();
  showStorageInfo();
  if (state.sync.url && navigator.onLine) {
    loadProductsFromSheet().then(renderAll).catch(e => { state.sync.error = e.message; renderSync(); });
  }
  trySync();
  setInterval(trySync, 60000);
  addEventListener("online", trySync);
  if ("serviceWorker" in navigator && location.protocol !== "file:") navigator.serviceWorker.register("sw.js");
})();
