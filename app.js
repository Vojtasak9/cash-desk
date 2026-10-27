// Cash Desk - runs completely in the browser (no server).
// Data is kept on the device in IndexedDB; "Save Excel file" exports everything to one .xlsx.
"use strict";

const $ = s => document.querySelector(s);
const $$ = s => document.querySelectorAll(s);

// ---------------------------------------------------------------- storage

const DEFAULT_STATE = () => ({
  products: [
    {id: 1, name: "Řez", price: 20, stock: null, active: true},
    {id: 2, name: "Koláček", price: 15, stock: null, active: true},
    {id: 3, name: "Závin", price: 25, stock: null, active: true},
  ],
  payments: [],      // {receipt, date, time, items, total, method, received, change, status}
  paymentItems: [],  // {receipt, productId, name, qty, unitPrice, lineTotal}
  debts: [],         // {id, date, person, type: "Lent"|"Repaid", amount, note}
  settings: {shop_name: "Jarmark", account_name: "", iban: "", currency: "CZK"},
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
  await save();
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
  await save();
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

function debtTotals() {
  const people = {};
  for (const d of state.debts) {
    const p = people[d.person] ||= {person: d.person, lent: 0, repaid: 0};
    if (d.type === "Lent") p.lent += d.amount; else p.repaid += d.amount;
  }
  return Object.values(people).sort((a, b) => a.person.localeCompare(b.person))
    .map(p => ({person: p.person, lent: round2(p.lent), repaid: round2(p.repaid), owes: round2(p.lent - p.repaid)}));
}

// ---------------------------------------------------------------- Excel export / import
// Layout follows the school "Jarmark" workbook: Czech sheet names, "Heading 1" headers, amounts in Kč.

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DAY_NAMES = ["Neděle", "Pondělí", "Úterý", "Středa", "Čtvrtek", "Pátek", "Sobota"];
const SH = {
  sales: "Prodej", lines: "Evidence Prodeje", products: "Evidence Pokladny", receipts: "Účtenky",
  days: "Denní uzávěrka", debts: "Dluhy", debtors: "Kdo dluží", settings: "Nastavení",
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

  // 2) Evidence Prodeje - every receipt line (Číslo / Produkt / Počet as in the Jarmark file)
  addSheet(wb, SH.lines, ["Číslo", "Produkt", "Počet", "Cena za kus", "Celkem", "Datum", "Stav"],
    state.paymentItems.map(i => [i.receipt, i.name, i.qty, i.unitPrice, i.lineTotal, dateOf[i.receipt] || "",
                                 ok.has(i.receipt) ? "" : "storno"]),
    {style: () => {}, widths: {Produkt: 18}, money: ["Cena za kus", "Celkem"]});

  // 3) Evidence Pokladny - product list with price and stock
  addSheet(wb, SH.products, ["Produkty:", "Cena za kus", "Skladem", "V prodeji"],
    state.products.map(p => [p.name, p.price, p.stock ?? "", p.active ? "ano" : "ne"]),
    {style: () => {}, widths: {"Produkty:": 18}, money: ["Cena za kus"]});

  // 4) Účtenky - one row per receipt with payment method
  addSheet(wb, SH.receipts, ["Číslo", "Datum", "Čas", "Položky", "Částka", "Platba", "Přijato", "Vráceno", "Stav"],
    state.payments.map(p => [p.receipt, p.date, p.time, p.items, p.total, METHOD_CZ[p.method] || p.method,
                             p.received ?? "", p.change ?? "", p.status === "VOID" ? "storno" : "OK"]),
    {style: heading, widths: {Položky: 45, Datum: 12}, money: ["Částka", "Přijato", "Vráceno"]});

  // 5) Denní uzávěrka - totals per day with a sum row
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

  // 6) Dluhy + Kdo dluží
  addSheet(wb, SH.debts, ["Datum", "Osoba", "Typ", "Částka", "Poznámka"],
    state.debts.map(d => [d.date, d.person, d.type === "Lent" ? "Půjčeno" : "Vráceno", d.amount, d.note || ""]),
    {style: heading, widths: {Datum: 12, Osoba: 18, Poznámka: 30}, money: ["Částka"]});
  addSheet(wb, SH.debtors, ["Osoba", "Půjčeno", "Vráceno", "Ještě dluží"],
    debtTotals().map(d => [d.person, d.lent, d.repaid, d.owes]),
    {style: heading, widths: {Osoba: 18}, money: ["Půjčeno", "Vráceno", "Ještě dluží"]});

  // 7) Nastavení - needed to load the file back into the app
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
  next.payments = sheetRows(wb, SH.receipts).filter(r => r["Číslo"] !== "").map(r => ({
    receipt: normalizeReceipt(r["Číslo"]), date: cellDate(r.Datum), time: cellTime(r["Čas"]),
    items: String(r["Položky"]), total: num(r["Částka"]), method: fromCz(r.Platba),
    received: blankToNull(r["Přijato"]), change: blankToNull(r["Vráceno"]),
    status: String(r.Stav).toLowerCase() === "storno" ? "VOID" : "OK",
  }));
  next.paymentItems = sheetRows(wb, SH.lines).filter(r => r["Číslo"] !== "").map(r => ({
    receipt: normalizeReceipt(r["Číslo"]), productId: idOf(String(r.Produkt)), name: String(r.Produkt),
    qty: num(r["Počet"]), unitPrice: num(r["Cena za kus"]), lineTotal: num(r.Celkem),
  }));
  next.debts = sheetRows(wb, SH.debts, false).filter(r => r.Osoba !== "").map((r, i) => ({
    id: i + 1, date: cellDate(r.Datum), person: String(r.Osoba),
    type: r.Typ === "Vráceno" ? "Repaid" : "Lent", amount: num(r["Částka"]), note: String(r["Poznámka"] || ""),
  }));
  for (const r of sheetRows(wb, SH.settings, false)) if (r["Klíč"] in next.settings) next.settings[r["Klíč"]] = String(r.Hodnota);
  if (!confirm(`Replace all data on this iPad with this file?\n\n${next.products.length} products, ` +
               `${next.payments.length} receipts, ${next.debts.length} debt entries.`)) return;
  state = next;
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
  if (b.dataset.tab === "debts") renderDebts();
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
    renderOrder(); renderProducts();
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
    renderOrder(); renderProducts(); renderPayments();
  }));
}
$("#day").onchange = renderPayments;

// ---- products
function renderProducts() {
  $("#prodRows").innerHTML = state.products.map(p => `<tr>
    <td>${p.id}</td><td>${esc(p.name)}</td><td class="num">${money(p.price)}</td>
    <td class="num">${p.stock == null ? "—" : p.stock}</td><td>${p.active ? "yes" : "no"}</td>
    <td><button class="link" data-edit="${p.id}">edit</button><button class="link bad" data-del="${p.id}">delete</button></td></tr>`).join("");
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

// ---- debts
function renderDebts() {
  const totals = debtTotals();
  $("#debtSummary").innerHTML = totals.map(s => `<tr><td>${esc(s.person)}</td>
    <td class="num">${money(s.lent)}</td><td class="num">${money(s.repaid)}</td>
    <td class="num ${s.owes > 0 ? "owes" : "settled"}">${money(s.owes)}</td></tr>`).join("")
    || '<tr><td colspan="4" class="muted">Nobody owes you anything</td></tr>';
  $("#debtRows").innerHTML = state.debts.slice().reverse().map(e => `<tr><td>${esc(e.date)}</td><td>${esc(e.person)}</td>
    <td>${e.type === "Lent" ? "Lent" : "Paid back"}</td><td class="num">${money(e.amount)}</td><td>${esc(e.note)}</td>
    <td><button class="link bad" data-ddel="${e.id}">delete</button></td></tr>`).join("");
  $("#peopleList").innerHTML = totals.map(s => `<option value="${esc(s.person)}">`).join("");
  $$("[data-ddel]").forEach(b => b.onclick = guard(async () => {
    if (!confirm("Delete this entry?")) return;
    state.debts = state.debts.filter(d => d.id !== +b.dataset.ddel);
    await save();
    renderDebts();
  }));
}
$("#dSave").onclick = guard(async () => {
  const person = $("#dPerson").value.trim();
  if (!person) throw new UserError("Person is required.");
  const amount = num($("#dAmount").value, NaN);
  if (isNaN(amount) || amount <= 0) throw new UserError("Amount must be more than 0.");
  state.debts.push({id: Math.max(0, ...state.debts.map(d => d.id)) + 1, date: today(), person,
                    type: $("#dType").value, amount, note: $("#dNote").value.trim()});
  await save();
  ["#dAmount", "#dNote"].forEach(s => $(s).value = "");
  renderDebts();
  toast("Saved");
});

// ---- settings + Excel
function renderSettings() {
  const s = state.settings;
  $("#shopName").textContent = s.shop_name || "Cash Desk";
  $("#sShop").value = s.shop_name; $("#sName").value = s.account_name;
  $("#sIban").value = s.iban; $("#sCur").value = s.currency;
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
  renderSettings(); renderOrder(); renderProducts(); renderDebts(); renderPayments();
}

// ---------------------------------------------------------------- start

(async () => {
  try {
    db = await openDb();
    state = await dbGet("state");
    if (!state) { state = DEFAULT_STATE(); await save(false); }
  } catch (e) {
    document.body.innerHTML = `<p style="padding:24px;font-size:18px">This browser can't store data (${esc(e.message)}).
      In Safari, turn off Private Browsing and open the app again.</p>`;
    return;
  }
  $("#day").value = today();
  renderAll();
  showStorageInfo();
  if ("serviceWorker" in navigator && location.protocol !== "file:") navigator.serviceWorker.register("sw.js");
})();
