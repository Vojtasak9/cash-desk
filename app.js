// Cash Desk - runs completely in the browser (no server).
// Data is kept on the device in IndexedDB; "Save Excel file" exports everything to one .xlsx.
"use strict";

const $ = s => document.querySelector(s);
const $$ = s => document.querySelectorAll(s);

// ---------------------------------------------------------------- storage

const DEFAULT_STATE = () => ({
  products: [
    {id: 1, name: "Coffee", price: 45, stock: 100, active: true},
    {id: 2, name: "Tea", price: 35, stock: 100, active: true},
    {id: 3, name: "Croissant", price: 40, stock: 20, active: true},
  ],
  payments: [],      // {receipt, date, time, items, total, method, received, change, status}
  paymentItems: [],  // {receipt, productId, name, qty, unitPrice, lineTotal}
  debts: [],         // {id, date, person, type: "Lent"|"Repaid", amount, note}
  settings: {shop_name: "Cash desk", account_name: "", iban: "", currency: "CZK"},
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
const formatReceipt = n => "P " + String(n).padStart(3, "0");
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

const SHEETS = {
  Products: ["ID", "Name", "Price", "Stock", "Active"],
  Payments: ["Receipt", "Date", "Time", "Items", "Total", "Method", "Received", "Change", "Status"],
  PaymentItems: ["Receipt", "Product ID", "Name", "Qty", "Unit price", "Line total"],
  Debts: ["ID", "Date", "Person", "Type", "Amount", "Note"],
  DebtSummary: ["Person", "Lent", "Repaid", "Still owes"],
  DailySummary: ["Date", "Receipts", "Cash", "QR", "Total"],
  Settings: ["Key", "Value"],
};

function buildWorkbook() {
  const rows = {
    Products: state.products.map(p => [p.id, p.name, p.price, p.stock ?? "", p.active ? "yes" : "no"]),
    Payments: state.payments.map(p => [p.receipt, p.date, p.time, p.items, p.total, p.method, p.received ?? "", p.change ?? "", p.status]),
    PaymentItems: state.paymentItems.map(i => [i.receipt, i.productId, i.name, i.qty, i.unitPrice, i.lineTotal]),
    Debts: state.debts.map(d => [d.id, d.date, d.person, d.type, d.amount, d.note || ""]),
    DebtSummary: debtTotals().map(d => [d.person, d.lent, d.repaid, d.owes]),
    DailySummary: dailyTotals().map(d => [d.date, d.receipts, d.cash, d.qr, d.total]),
    Settings: Object.entries(state.settings),
  };
  const wb = XLSX.utils.book_new();
  for (const [name, headers] of Object.entries(SHEETS)) {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows[name]]);
    ws["!cols"] = headers.map(h => ({wch: h === "Items" ? 50 : Math.max(12, h.length + 4)}));
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  return wb;
}

const exportName = () => `cash_desk_${today()}.xlsx`;

async function markExported() {
  state.lastExport = new Date().toISOString();
  state.changedSinceExport = false;
  await save(false);
  renderSettings();
}

function exportDownload() {
  XLSX.writeFile(buildWorkbook(), exportName());
  markExported();
}

async function exportShare() {
  const data = XLSX.write(buildWorkbook(), {type: "array", bookType: "xlsx"});
  const file = new File([data], exportName(), {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  if (navigator.canShare && navigator.canShare({files: [file]})) {
    try {
      await navigator.share({files: [file], title: exportName()});
      await markExported();
    } catch (e) {
      if (e.name !== "AbortError") toast("Sharing failed - use Download instead", true);
    }
  } else {
    exportDownload();
  }
}

function cellDate(v) {
  if (v instanceof Date) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
  return String(v ?? "");
}
function cellTime(v) {
  if (v instanceof Date) return `${pad(v.getHours())}:${pad(v.getMinutes())}:${pad(v.getSeconds())}`;
  return String(v ?? "");
}
const blankToNull = v => (v === "" || v == null) ? null : num(v);

async function importExcel(file) {
  const wb = XLSX.read(await file.arrayBuffer(), {cellDates: true});
  const sheet = name => {
    const ws = wb.Sheets[name];
    if (!ws) throw new UserError(`This isn't a Cash Desk Excel file (sheet "${name}" is missing).`);
    return XLSX.utils.sheet_to_json(ws, {defval: ""});
  };
  const next = DEFAULT_STATE();
  next.products = sheet("Products").filter(r => r.Name !== "").map(r => ({
    id: parseInt(r.ID, 10), name: String(r.Name), price: num(r.Price),
    stock: r.Stock === "" ? null : Math.round(num(r.Stock)),
    active: !["no", "false", "0"].includes(String(r.Active).toLowerCase()),
  }));
  next.payments = sheet("Payments").filter(r => r.Receipt !== "").map(r => ({
    receipt: String(r.Receipt), date: cellDate(r.Date), time: cellTime(r.Time), items: String(r.Items),
    total: num(r.Total), method: r.Method === "QR" ? "QR" : "Cash",
    received: blankToNull(r.Received), change: blankToNull(r.Change), status: r.Status === "VOID" ? "VOID" : "OK",
  }));
  next.paymentItems = sheet("PaymentItems").filter(r => r.Receipt !== "").map(r => ({
    receipt: String(r.Receipt), productId: parseInt(r["Product ID"], 10), name: String(r.Name),
    qty: num(r.Qty), unitPrice: num(r["Unit price"]), lineTotal: num(r["Line total"]),
  }));
  next.debts = sheet("Debts").filter(r => r.Person !== "").map(r => ({
    id: parseInt(r.ID, 10), date: cellDate(r.Date), person: String(r.Person),
    type: r.Type === "Repaid" ? "Repaid" : "Lent", amount: num(r.Amount), note: String(r.Note || ""),
  }));
  for (const r of sheet("Settings")) if (r.Key in next.settings) next.settings[r.Key] = String(r.Value);
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

$$("[data-export]").forEach(b => b.onclick = () => {
  $("#exportName").textContent = exportName();
  $("#exportModal").classList.add("open");
});
$("#exportShare").onclick = guard(async () => { await exportShare(); closeModals(); });
$("#exportDownload").onclick = guard(() => { exportDownload(); closeModals(); });
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
