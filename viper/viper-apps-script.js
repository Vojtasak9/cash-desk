/**
 * VIPER - the cash desk app's Google Sheet backend (experimental version)
 *
 * The spreadsheet "Viper" does all the calculating with formulas; the app only types in plain data:
 *   - Produkty      what we sell and the price             -> the app reads it
 *   - Sklad         stock: Naskladněno (typed) - sold on each day (formulas) = Zbývá -> the app reads it
 *   - Typy pohybů   what each kind of money movement does to the cash and to debts
 *   - Půjčky        every money movement (loans, expenses, advances...) <- the app writes it
 *   - Den 1, Den 2  every sale (PPD) with one column per product      <- the app writes it
 *                   + at the top: cash in the kasa now and all money (cash + account)
 *
 * setupViper() builds all sheets once (run it from the editor). doPost() is the web app the app calls.
 */

// Pick your own password, put it here AND type the same one into the app.
const KEY = "ZMEN-ME";

const VIPER_ID = "1ZsqWHUY1iF5czN0nYmrUwboyXHMJ65mTZn6niF1sxLk";

const S = {products: "Produkty", stock: "Sklad", types: "Typy pohybů", moves: "Půjčky"};
const DAY_PREFIX = "Den ";
const DAYS = 2;              // Den 1, Den 2
const MAX_PRODUCTS = 10;     // product rows 2..11 in Produkty, columns E..N in the day sheets
const DAY_HEADER_ROW = 6;    // day sheets: summary in rows 1-5, table header in row 6
const DAY_FIRST = 7;
const DAY_LAST = 300;
const MOVES_LAST = 300;
const APP_NOTE = "Zapsáno z pokladny";

// Money movement types: [name, effect on money (cash or account), effect on "Kasa dluží", explanation]
const TYPES = [
  ["Zapůjčka do kasy", 1, 1, "Někdo půjčí kase peníze (např. Kristián 500 Kč na start). Kasa mu je musí vrátit."],
  ["Zaplaceno za kasu", 0, 1, "Člen zaplatil něco za kasu ze svého (např. nákup surovin). V kase se nic nezměnilo, ale kasa mu dluží."],
  ["Záloha na nákup", -1, -1, "Někdo si vzal peníze z kasy, aby něco nakoupil. Dokud nedoloží nákup, dluží je kase."],
  ["Vyúčtování zálohy", 0, 1, "Ten, kdo měl zálohu, doloží, kolik utratil (účtenka). Sníží se, kolik dluží kase."],
  ["Vrácení zbytku", 1, 1, "Ten, kdo měl zálohu, vrátí do kasy, co zbylo."],
  ["Vrácení dluhu", -1, -1, "Kasa vrací peníze tomu, komu dluží (např. Kristiánovi jeho 500 Kč)."],
  ["Výdaj z kasy", -1, 0, "Nákup zaplacený přímo z kasy, nikdo nic nedluží."],
  ["Jiný příjem", 1, 0, "Peníze do kasy, které nikomu nedlužíme."],
];

// ======================================================================= setup (run once)

function setupViper() {
  const ss = SpreadsheetApp.openById(VIPER_ID);
  if (ss.getSheetByName(S.products)) throw new Error("Viper is already set up - nothing changed.");
  const first = ss.getSheets()[0];
  const header = (sh, row, n) => sh.getRange(row, 1, 1, n).setFontWeight("bold").setBackground("#dfe7f5");
  const money = r => r.setNumberFormat('#,##0 "Kč"');

  // --- Produkty
  const pr = first.setName(S.products);
  pr.getRange("A1:B1").setValues([["Produkt", "Cena"]]);
  pr.getRange("A2:B4").setValues([["Párek v rohlíku", 40], ["Hot dog", 55], ["Addon", 5]]);
  header(pr, 1, 2); money(pr.getRange("B2:B" + (MAX_PRODUCTS + 1)));
  pr.getRange("D1").setValue("Nové produkty přidávejte na konec seznamu (max. " + MAX_PRODUCTS +
    "). Pořadí neměňte - podle něj jsou sloupce v listech Den 1 a Den 2.");
  pr.setColumnWidth(1, 170); pr.setFrozenRows(1);

  // --- Typy pohybů
  const ty = ss.insertSheet(S.types);
  ty.getRange("A1:D1").setValues([["Typ", "Peníze (+1 přibude / -1 ubude)", "Kasa dluží (+1 / -1)", "Kdy použít"]]);
  ty.getRange(2, 1, TYPES.length, 4).setValues(TYPES);
  header(ty, 1, 4); ty.setColumnWidth(1, 170); ty.setColumnWidth(4, 520); ty.setFrozenRows(1);

  // --- Den 1, Den 2
  for (let d = 1; d <= DAYS; d++) buildDay(ss, d, header, money);

  // --- Půjčky
  const mv = ss.insertSheet(S.moves, 2);
  mv.getRange("A1:K1").setValues([["Datum", "Den", "Číslo", "Typ", "Osoba", "Popis", "Částka", "Platba",
                                   "Hotovost", "Účet", "Kasa dluží"]]);
  header(mv, 1, 11);
  const f = [];
  for (let r = 2; r <= MOVES_LAST; r++) {
    const look = "VLOOKUP($D" + r + ",'" + S.types + "'!$A:$C,";
    f.push([
      '=IF($D' + r + '="","",IF($H' + r + '="Účet",0,$G' + r + '*' + look + '2,FALSE)))',
      '=IF($D' + r + '="","",IF($H' + r + '="Účet",$G' + r + '*' + look + '2,FALSE),0))',
      '=IF($D' + r + '="","",$G' + r + '*' + look + '3,FALSE))',
    ]);
  }
  mv.getRange(2, 9, f.length, 3).setFormulas(f);
  money(mv.getRange("G2:G" + MOVES_LAST)); money(mv.getRange("I2:K" + MOVES_LAST));
  mv.getRange("A2:A" + MOVES_LAST).setNumberFormat("d.m.yyyy");
  mv.getRange("D2:D" + MOVES_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ty.getRange("A2:A" + (TYPES.length + 1)), true).build());
  mv.getRange("H2:H" + MOVES_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(["Hotově", "Účet"], true).build());
  // starting money: Kristián lends 500 Kč, which the kasa has to give back
  mv.getRange("B2:H2").setValues([[1, "PPD 001", "Zapůjčka do kasy", "Kristián", "Peníze na start (vrátit)", 500, "Hotově"]]);
  // who owes whom
  mv.getRange("M1:N1").setValues([["Osoba", "Kasa dluží (+) / dluží kase (−)"]]); mv.getRange("M1:N1").setFontWeight("bold").setBackground("#dfe7f5");
  mv.getRange("M2").setFormula('=IFERROR(UNIQUE(FILTER(E2:E' + MOVES_LAST + ',E2:E' + MOVES_LAST + '<>"")),"")');
  mv.getRange("N2").setFormula('=ARRAYFORMULA(IF(M2:M50="","",SUMIF(E2:E' + MOVES_LAST + ',M2:M50,K2:K' + MOVES_LAST + ')))');
  money(mv.getRange("N2:N50"));
  mv.setColumnWidth(6, 230); mv.setColumnWidth(13, 140); mv.setColumnWidth(14, 210); mv.setFrozenRows(1);

  // --- Sklad
  const st = ss.insertSheet(S.stock, 1);
  const cols = ["Produkt", "Naskladněno (ks)"];
  for (let d = 1; d <= DAYS; d++) cols.push("Prodáno " + DAY_PREFIX + d);
  cols.push("Prodáno celkem", "Zbývá (ks)", "Hodnota zbytku");
  st.getRange(1, 1, 1, cols.length).setValues([cols]); header(st, 1, cols.length);
  const total = 3 + DAYS, left = total + 1, value = left + 1;
  const L = c => String.fromCharCode(64 + c);
  const rows = [];
  for (let r = 2; r <= MAX_PRODUCTS + 1; r++) {
    const row = ['=IF(' + S.products + '!A' + r + '="","",' + S.products + '!A' + r + ')', null];
    for (let d = 1; d <= DAYS; d++) {
      const day = "'" + DAY_PREFIX + d + "'!";
      row.push('=IF($A' + r + '="","",SUMPRODUCT((' + day + '$E$' + DAY_HEADER_ROW + ':$N$' + DAY_HEADER_ROW + '=$A' + r + ')*' +
               day + '$E$' + DAY_FIRST + ':$N$' + DAY_LAST + '))');
    }
    row.push('=IF($A' + r + '="","",SUM(C' + r + ':' + L(total - 1) + r + '))');
    row.push('=IF(OR($A' + r + '="",$B' + r + '=""),"",$B' + r + '-' + L(total) + r + ')');
    row.push('=IF(' + L(left) + r + '="","",' + L(left) + r + '*' + S.products + '!B' + r + ')');
    rows.push(row);
  }
  for (let i = 0; i < rows.length; i++) {
    const r = i + 2;
    st.getRange(r, 1).setFormula(rows[i][0]);
    st.getRange(r, 3, 1, rows[i].length - 2).setFormulas([rows[i].slice(2)]);
  }
  const sumRow = MAX_PRODUCTS + 2;
  st.getRange(sumRow, 1).setValue("Celkem").setFontWeight("bold");
  for (let c = 2; c <= value; c++) st.getRange(sumRow, c).setFormula('=SUM(' + L(c) + '2:' + L(c) + (sumRow - 1) + ')').setFontWeight("bold");
  money(st.getRange(2, value, sumRow - 1, 1));
  st.getRange("B2:B" + (MAX_PRODUCTS + 1)).setBackground("#fff6db");
  st.getRange(sumRow + 2, 1).setValue("Žluté buňky vyplňte (kolik kusů jsme naskladnili). Ostatní počítají vzorce.");
  st.setColumnWidth(1, 170); st.setFrozenRows(1);

  SpreadsheetApp.flush();
}

function buildDay(ss, d, header, money) {
  const sh = ss.insertSheet(DAY_PREFIX + d);
  const P = "'" + S.moves + "'!";
  sh.getRange("A1").setValue(DAY_PREFIX + d).setFontWeight("bold").setFontSize(16);
  sh.getRange("B1").setValue("Datum:");
  sh.getRange("C1").setNumberFormat("d.m.yyyy").setBackground("#fff6db");
  const prev = d > 1 ? "'" + DAY_PREFIX + (d - 1) + "'!" : null;
  const rangeH = "P" + DAY_FIRST + ":P" + DAY_LAST, rangeQ = "Q" + DAY_FIRST + ":Q" + DAY_LAST;
  sh.getRange("A2:A5").setValues([["Hotovost v kase"], ["Na účtu (QR a převody)"], ["Peníze celkem"], ["Tržba dne"]]);
  sh.getRange("C2").setFormula('=' + (prev ? prev + 'C2+' : '') + 'SUMIF(' + P + '$B:$B,' + d + ',' + P + '$I:$I)+SUM(' + rangeH + ')');
  sh.getRange("C3").setFormula('=' + (prev ? prev + 'C3+' : '') + 'SUMIF(' + P + '$B:$B,' + d + ',' + P + '$J:$J)+SUM(' + rangeQ + ')');
  sh.getRange("C4").setFormula("=C2+C3");
  sh.getRange("C5").setFormula("=SUM(O" + DAY_FIRST + ":O" + DAY_LAST + ")");
  sh.getRange("A2:C5").setFontWeight("bold"); money(sh.getRange("C2:C5"));
  sh.getRange("D2").setValue(prev ? "= konec předchozího dne + pohyby a prodeje v hotovosti dnes" : "= pohyby v hotovosti (Půjčky) + prodeje v hotovosti");
  // table
  sh.getRange(DAY_HEADER_ROW, 1, 1, 4).setValues([["Číslo", "Čas", "Položky", "Platba"]]);
  sh.getRange(DAY_HEADER_ROW, 5).setFormula('=TRANSPOSE(' + S.products + '!A2:A' + (MAX_PRODUCTS + 1) + ')');
  sh.getRange(DAY_HEADER_ROW, 15, 1, 4).setValues([["Částka", "Hotově", "QR", "Stav"]]);
  header(sh, DAY_HEADER_ROW, 18);
  const f = [];
  for (let r = DAY_FIRST; r <= DAY_LAST; r++) {
    f.push(['=IF($A' + r + '="","",TEXTJOIN(", ",TRUE,ARRAYFORMULA(IF($E' + r + ':$N' + r + '>0,$E' + r + ':$N' + r + '&"x "&$E$' + DAY_HEADER_ROW + ':$N$' + DAY_HEADER_ROW + ',""))))']);
  }
  sh.getRange(DAY_FIRST, 3, f.length, 1).setFormulas(f);
  const g = [];
  for (let r = DAY_FIRST; r <= DAY_LAST; r++) {
    g.push(['=IF($A' + r + '="","",IF($D' + r + '="Hotově",$O' + r + ',0))', '=IF($A' + r + '="","",IF($D' + r + '="QR",$O' + r + ',0))']);
  }
  sh.getRange(DAY_FIRST, 16, g.length, 2).setFormulas(g);
  money(sh.getRange("O" + DAY_FIRST + ":Q" + DAY_LAST));
  sh.getRange("D" + DAY_FIRST + ":D" + DAY_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(["Hotově", "QR"], true).build());
  sh.setColumnWidth(3, 260); sh.setFrozenRows(DAY_HEADER_ROW);
}

// ======================================================================= web app

function doGet() {
  return ContentService.createTextOutput("Viper script is running.");
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (KEY === "ZMEN-ME") throw new Error('Change KEY at the top of the script first (it still says "ZMEN-ME").');
    if (req.key !== KEY) throw new Error("Wrong key - type the same key into the app as in the script.");
    const ss = SpreadsheetApp.openById(VIPER_ID);
    if (req.action === "products") {
      out = {ok: true, products: readProducts(ss), numbers: usedNumbers(ss), day: dayInfo(ss, req.date)};
    } else if (req.action === "sync") {
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      const skipped = [];
      try {
        req.ops.forEach(op => applyOp(ss, op, skipped));
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
      out = {ok: true, done: req.ops.length, skipped: skipped, day: dayInfo(ss, req.date)};
    } else {
      throw new Error("Unknown action: " + req.action);
    }
  } catch (err) {
    out = {ok: false, error: String((err && err.message) || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function applyOp(ss, op, skipped) {
  if (op.type === "sale") writeSale(ss, op, skipped);
  else if (op.type === "entry") writeMove(ss, op, skipped);
  else if (op.type === "deleteEntry") deleteMove(ss, op.number, skipped);
  else throw new Error("Unknown operation: " + op.type);
}

// Products with price and what is left in stock (Sklad / Zbývá; null when Naskladněno is empty)
function readProducts(ss) {
  const pr = sheet(ss, S.products).getRange(2, 1, MAX_PRODUCTS, 2).getValues();
  const st = sheet(ss, S.stock);
  const leftCol = 3 + DAYS + 1;
  const left = st.getRange(2, leftCol, MAX_PRODUCTS, 1).getValues();
  const out = [];
  pr.forEach((r, i) => {
    if (String(r[0]).trim() === "") return;
    const v = left[i][0];
    out.push({name: String(r[0]).trim(), price: toNumber(r[1]), stock: v === "" ? null : toNumber(v)});
  });
  return out;
}

// Highest PPD / VPD / ID numbers anywhere in the sheets (also typed by hand)
function usedNumbers(ss) {
  const out = {PPD: 0, VPD: 0, ID: 0};
  const take = v => {
    const m = norm(v).match(/^(PPD|VPD|ID)(\d+)$/);
    if (m) out[m[1]] = Math.max(out[m[1]], +m[2]);
  };
  for (let d = 1; d <= DAYS; d++) {
    sheet(ss, DAY_PREFIX + d).getRange(DAY_FIRST, 1, DAY_LAST - DAY_FIRST + 1, 1).getValues().forEach(r => take(r[0]));
  }
  sheet(ss, S.moves).getRange(2, 3, MOVES_LAST - 1, 1).getValues().forEach(r => take(r[0]));
  return out;
}

// The day sheet for a date: the one with that date in C1, else the first one without a date (it gets the date)
function daySheet(ss, iso, claim) {
  const want = toDate(iso).toDateString();
  let free = null;
  for (let d = 1; d <= DAYS; d++) {
    const sh = sheet(ss, DAY_PREFIX + d);
    const v = sh.getRange("C1").getValue();
    if (v instanceof Date && v.toDateString() === want) return {sh: sh, n: d};
    if (v === "" && !free) free = {sh: sh, n: d};
  }
  if (!free) throw new Error("All day sheets already have another date. Add a new day sheet or change the date in C1.");
  if (claim) free.sh.getRange("C1").setValue(toDate(iso));
  return free;
}

// Cash and money for the app's header, from the day sheet of that date
function dayInfo(ss, iso) {
  if (!iso) return null;
  let day;
  try { day = daySheet(ss, iso, false); } catch (e) { return null; }
  const v = day.sh.getRange("C2:C5").getValues().map(r => toNumber(r[0]));
  return {name: DAY_PREFIX + day.n, cash: v[0], bank: v[1], total: v[2], sales: v[3]};
}

// ---- sales -> Den N:  Číslo | Čas | (Položky = formula) | Platba | one column per product | Částka | (Hotově, QR = formulas) | Stav
function writeSale(ss, op, skipped) {
  let found = null;
  for (let d = 1; d <= DAYS && !found; d++) {
    const sh = sheet(ss, DAY_PREFIX + d);
    const row = findRow(sh, 1, op.number, DAY_FIRST, DAY_LAST);
    if (row) found = {sh: sh, row: row};
  }
  if (found && found.sh.getRange(found.row, 1).getNote() !== APP_NOTE) { skipped.push(op.number); return; }
  let sh, row;
  if (found) { sh = found.sh; row = found.row; }
  else {
    sh = daySheet(ss, op.date, true).sh;
    row = findRow(sh, 1, "", DAY_FIRST, DAY_LAST);
    if (!row) throw new Error(sh.getName() + " is full (" + (DAY_LAST - DAY_FIRST + 1) + " sales).");
  }
  const names = sh.getRange(DAY_HEADER_ROW, 5, 1, MAX_PRODUCTS).getValues()[0].map(v => String(v).trim().toLowerCase());
  const qty = new Array(MAX_PRODUCTS).fill("");
  const unknown = [];
  if (!op.void) {
    op.items.forEach(it => {
      const i = names.indexOf(String(it.name).trim().toLowerCase());
      if (i === -1) unknown.push(it.qty + "x " + it.name);
      else qty[i] = (Number(qty[i]) || 0) + it.qty;
    });
  }
  sh.getRange(row, 1, 1, 2).setValues([[op.number, op.time]]);
  sh.getRange(row, 4).setValue(op.method === "QR" ? "QR" : "Hotově");
  sh.getRange(row, 5, 1, MAX_PRODUCTS).setValues([qty]);
  sh.getRange(row, 15).setValue(op.void ? "" : op.total);
  sh.getRange(row, 18).setValue(op.void ? "STORNO" : (unknown.length ? "Není v Produkty: " + unknown.join(", ") : ""));
  sh.getRange(row, 1).setNote(APP_NOTE);
  sh.getRange(row, 3).setNote(op.void ? "STORNO: " + op.items.map(it => it.qty + "x " + it.name).join(", ") : "");
}

// ---- money movements -> Půjčky:  Datum | Den | Číslo | Typ | Osoba | Popis | Částka | Platba | (formulas)
function writeMove(ss, op, skipped) {
  const sh = sheet(ss, S.moves);
  let row = findRow(sh, 3, op.number, 2, MOVES_LAST);
  if (row && sh.getRange(row, 3).getNote() !== APP_NOTE) { skipped.push(op.number); return; }
  if (!row) row = findRow(sh, 4, "", 2, MOVES_LAST, sh.getRange(2, 3, MOVES_LAST - 1, 1).getValues());
  if (!row) throw new Error("Půjčky is full.");
  const day = daySheet(ss, op.date, true);
  sh.getRange(row, 1, 1, 8).setValues([[toDate(op.date), day.n, op.number, op.kind, op.person || "", op.description || "",
                                        op.amount, op.payment === "Účet" ? "Účet" : "Hotově"]]);
  sh.getRange(row, 3).setNote(APP_NOTE);
}

function deleteMove(ss, number, skipped) {
  const sh = sheet(ss, S.moves);
  const row = findRow(sh, 3, number, 2, MOVES_LAST);
  if (!row) return;
  if (sh.getRange(row, 3).getNote() !== APP_NOTE) { skipped.push(number); return; }
  sh.getRange(row, 1, 1, 8).clearContent();
  sh.getRange(row, 3).setNote("");
}

// ---- helpers

function sheet(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found in Viper - was it renamed?');
  return sh;
}

function norm(v) {
  return String(v).replace(/\s+/g, "").toUpperCase();
}

// First row in first..last whose column equals value ("" = empty). `numbers` (optional) must also be empty for "".
function findRow(sh, col, value, first, last, numbers) {
  const want = norm(value);
  const vals = sh.getRange(first, col, last - first + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (norm(vals[i][0]) === want && (!numbers || want !== "" || norm(numbers[i][0]) === "")) return first + i;
  }
  return null;
}

function toDate(iso) {
  const p = String(iso).split("-").map(Number);
  return new Date(p[0], p[1] - 1, p[2], 12);
}

function toNumber(v) {
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/[^\d,.-]/g, "").replace(",", "."));
  return isNaN(n) ? 0 : n;
}
