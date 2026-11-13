/**
 * VIPER - the cash desk app's Google Sheet (experimental version)
 *
 * The sheet works on its own: you type only WHAT you sold and HOW MANY, formulas do the rest.
 * The app is just a nicer way to type the same rows.
 *
 *   Produkty      what we sell and the price
 *   Sklad         Naskladněno (you type) - Prodáno (SUMIF from the day sheets) = Zbývá
 *   Typy pohybů   lookup table: what each kind of money movement does (Půjčky reads it with VLOOKUP)
 *   Půjčky        loans, expenses, advances...; formulas work out the cash and who owes whom
 *   Den 1, Den 2  one row per product sold: Produkt + Počet -> price, total, cash/QR by formulas;
 *                 at the top: cash in the kasa now, money on the account, all money, today's sales
 *
 * The spreadsheet uses Czech settings, so formulas use ";" between arguments.
 * rebuildViper() deletes all sheets and builds them again (run it from the editor).
 */

// Pick your own password, put it here AND type the same one into the app.
const KEY = "ZMEN-ME";

const VIPER_ID = "1ZsqWHUY1iF5czN0nYmrUwboyXHMJ65mTZn6niF1sxLk";

const S = {products: "Produkty", stock: "Sklad", types: "Typy pohybů", moves: "Půjčky"};
const DAY = d => "Den " + d;
const DAYS = 2;
const FIRST = 7;      // day sheets: summary rows 1-5, header row 6, sales from row 7
const LAST = 400;
const MOVES_LAST = 200;
const PRODUCTS_LAST = 30;
const APP_NOTE = "Zapsáno z pokladny";

// [type, money (+1 comes in / -1 goes out / 0 nothing moves), kasa owes the person (+1 more / -1 less), when to use it]
const TYPES = [
  ["Zapůjčka do kasy", 1, 1, "Někdo půjčí kase peníze (např. Kristián 500 Kč na start). Kasa mu je musí vrátit."],
  ["Zaplaceno za kasu", 0, 1, "Někdo zaplatil něco za kasu ze svého (např. nákup surovin). V kase se nic nezměnilo, ale kasa mu dluží."],
  ["Záloha na nákup", -1, -1, "Někdo si vzal peníze z kasy, aby něco nakoupil. Dokud nedoloží účtenku, dluží je kase."],
  ["Vyúčtování zálohy", 0, 1, "Ten, kdo měl zálohu, ukáže účtenku - kolik utratil."],
  ["Vrácení zbytku", 1, 1, "Ten, kdo měl zálohu, vrátí do kasy, co zbylo."],
  ["Vrácení dluhu", -1, -1, "Kasa vrací peníze tomu, komu dluží (např. Kristiánovi jeho 500 Kč)."],
  ["Výdaj z kasy", -1, 0, "Nákup zaplacený přímo z kasy, nikdo nic nedluží."],
  ["Jiný příjem", 1, 0, "Peníze do kasy, které nikomu nedlužíme."],
];

const BLUE = "#dfe7f5", YELLOW = "#fff6db", GREY = "#f1f3f4";

// ======================================================================= build the sheets

function rebuildViper() {
  const ss = SpreadsheetApp.openById(VIPER_ID);
  const tmp = ss.insertSheet("tmp" + Date.now());
  ss.getSheets().forEach(sh => { if (sh.getSheetId() !== tmp.getSheetId()) ss.deleteSheet(sh); });
  buildProducts(ss);
  buildStock(ss);
  buildTypes(ss);
  buildMoves(ss);
  for (let d = 1; d <= DAYS; d++) buildDay(ss, d);
  ss.deleteSheet(tmp);
  ss.setActiveSheet(ss.getSheetByName(DAY(1)));
  SpreadsheetApp.flush();
}

const money = r => r.setNumberFormat('#,##0 "Kč"');
const head = r => r.setFontWeight("bold").setBackground(BLUE).setWrap(true).setVerticalAlignment("middle");
const typed = r => r.setBackground(YELLOW);   // you type here
const auto = r => r.setBackground(GREY);      // a formula fills this

function buildProducts(ss) {
  const sh = ss.insertSheet(S.products);
  sh.getRange("A1:B1").setValues([["Produkt", "Cena za kus"]]); head(sh.getRange("A1:B1"));
  sh.getRange("A2:B4").setValues([["Párek v rohlíku", 40], ["Hot dog", 55], ["Addon", 5]]);
  typed(sh.getRange("A2:B" + PRODUCTS_LAST)); money(sh.getRange("B2:B" + PRODUCTS_LAST));
  sh.getRange("D1:D3").setValues([["Žlutě = píšete vy."], ["Šedě = počítá vzorec."],
    ["Nový produkt napište na první volný řádek. Objeví se sám v Sklad i v nabídce v Den 1 / Den 2."]]);
  sh.setColumnWidth(1, 180); sh.setColumnWidth(4, 520); sh.setFrozenRows(1);
}

function buildStock(ss) {
  const sh = ss.insertSheet(S.stock);
  const cols = ["Produkt", "Naskladněno (ks)"];
  for (let d = 1; d <= DAYS; d++) cols.push("Prodáno " + DAY(d));
  cols.push("Prodáno celkem", "Zbývá (ks)");
  sh.getRange(1, 1, 1, cols.length).setValues([cols]); head(sh.getRange(1, 1, 1, cols.length));
  const n = PRODUCTS_LAST - 1, totalCol = 3 + DAYS, leftCol = totalCol + 1;
  const L = c => String.fromCharCode(64 + c);
  const f = [];
  for (let r = 2; r <= PRODUCTS_LAST; r++) {
    const row = ['=IF(' + S.products + '!A' + r + '="";"";' + S.products + '!A' + r + ')'];
    f.push(row);
  }
  sh.getRange(2, 1, n, 1).setFormulas(f);
  for (let d = 1; d <= DAYS; d++) {
    const day = "'" + DAY(d) + "'!";
    const g = [];
    for (let r = 2; r <= PRODUCTS_LAST; r++) {
      g.push(['=IF($A' + r + '="";"";SUMIF(' + day + '$C$' + FIRST + ':$C$' + LAST + ';$A' + r + ';' + day + '$D$' + FIRST + ':$D$' + LAST + '))']);
    }
    sh.getRange(2, 2 + d, n, 1).setFormulas(g);
  }
  const t = [], left = [];
  for (let r = 2; r <= PRODUCTS_LAST; r++) {
    t.push(['=IF($A' + r + '="";"";SUM(C' + r + ':' + L(totalCol - 1) + r + '))']);
    left.push(['=IF(OR($A' + r + '="";$B' + r + '="");"";$B' + r + '-' + L(totalCol) + r + ')']);
  }
  sh.getRange(2, totalCol, n, 1).setFormulas(t);
  sh.getRange(2, leftCol, n, 1).setFormulas(left);
  auto(sh.getRange(2, 1, n, 1)); typed(sh.getRange(2, 2, n, 1)); auto(sh.getRange(2, 3, n, leftCol - 2));
  sh.getRange(1, leftCol + 2, 5, 1).setValues([
    ["Jak to funguje:"],
    ["Naskladněno = kolik kusů jsme nakoupili (píšete vy)."],
    ["Prodáno Den 1 = SUMIF: sečte Počet ve všech řádcích listu Den 1, kde je tento produkt."],
    ["Prodáno celkem = SUM všech dnů."],
    ["Zbývá = Naskladněno - Prodáno celkem."]]);
  sh.getRange(1, leftCol + 2).setFontWeight("bold");
  sh.setColumnWidth(1, 180); sh.setColumnWidth(leftCol + 2, 520); sh.setFrozenRows(1);
}

function buildTypes(ss) {
  const sh = ss.insertSheet(S.types);
  sh.getRange("A1").setValue("TAHÁK pro list Půjčky: list Půjčky si odsud bere čísla funkcí VLOOKUP (podle Typu). Neměňte názvy typů.")
    .setFontWeight("bold");
  sh.getRange("A2:D2").setValues([["Typ", "Peníze: +1 přibudou, -1 ubudou, 0 nic", "Kasa dluží: +1 víc, -1 míň, 0 nic", "Kdy použít"]]);
  head(sh.getRange("A2:D2"));
  sh.getRange(3, 1, TYPES.length, 4).setValues(TYPES);
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 150); sh.setColumnWidth(3, 150); sh.setColumnWidth(4, 560); sh.setFrozenRows(2);
}

function buildMoves(ss) {
  const sh = ss.insertSheet(S.moves);
  sh.getRange("A1:K1").setValues([["Datum", "Den (1 nebo 2)", "Číslo", "Typ (co se stalo)", "Kdo", "Popis", "Částka",
                                   "Platba (prázdné = hotově)", "Hotovost +/-", "Účet +/-", "Kasa mu dluží +/-"]]);
  head(sh.getRange("A1:K1"));
  const T = "'" + S.types + "'!$A$3:$C$" + (TYPES.length + 2);
  const f = [];
  for (let r = 2; r <= MOVES_LAST; r++) {
    f.push([
      '=IF($D' + r + '="";"";IF($H' + r + '="Účet";0;$G' + r + '*VLOOKUP($D' + r + ';' + T + ';2;FALSE)))',
      '=IF($D' + r + '="";"";IF($H' + r + '="Účet";$G' + r + '*VLOOKUP($D' + r + ';' + T + ';2;FALSE);0))',
      '=IF($D' + r + '="";"";$G' + r + '*VLOOKUP($D' + r + ';' + T + ';3;FALSE))',
    ]);
  }
  sh.getRange(2, 9, f.length, 3).setFormulas(f);
  typed(sh.getRange("A2:H" + MOVES_LAST)); auto(sh.getRange("I2:K" + MOVES_LAST));
  money(sh.getRange("G2:G" + MOVES_LAST)); money(sh.getRange("I2:K" + MOVES_LAST));
  sh.getRange("A2:A" + MOVES_LAST).setNumberFormat("d.m.yyyy");
  sh.getRange("D2:D" + MOVES_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(S.types).getRange("A3:A" + (TYPES.length + 2)), true).build());
  sh.getRange("H2:H" + MOVES_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(["Hotově", "Účet"], true).build());
  sh.getRange("B2:H2").setValues([[1, "PPD 001", "Zapůjčka do kasy", "Kristián", "Peníze na start (vrátit)", 500, "Hotově"]]);
  // who owes whom
  sh.getRange("M1:N1").setValues([["Kdo", "Kasa mu dluží (+) / dluží kase (−)"]]); head(sh.getRange("M1:N1"));
  sh.getRange("M2").setFormula('=IFERROR(UNIQUE(FILTER(E2:E' + MOVES_LAST + ';E2:E' + MOVES_LAST + '<>""));"")');
  const g = [];
  for (let r = 2; r <= 30; r++) g.push(['=IF(M' + r + '="";"";SUMIF($E$2:$E$' + MOVES_LAST + ';M' + r + ';$K$2:$K$' + MOVES_LAST + '))']);
  sh.getRange(2, 14, g.length, 1).setFormulas(g);
  auto(sh.getRange("M2:N30")); money(sh.getRange("N2:N30"));
  sh.getRange("M32:M35").setValues([
    ["Jak to funguje:"],
    ["Hotovost = Částka × číslo z listu Typy pohybů (VLOOKUP podle Typu)."],
    ["Kasa mu dluží = Částka × druhé číslo z Typy pohybů."],
    ["Vpravo nahoře: SUMIF sečte, kolik kasa dluží každému."]]);
  sh.getRange("M32").setFontWeight("bold");
  sh.setColumnWidth(4, 160); sh.setColumnWidth(6, 220); sh.setColumnWidth(13, 140); sh.setColumnWidth(14, 220); sh.setFrozenRows(1);
}

function buildDay(ss, d) {
  const sh = ss.insertSheet(DAY(d));
  const P = "'" + S.moves + "'!";
  const prev = d > 1 ? "'" + DAY(d - 1) + "'!" : "";
  sh.getRange("A1").setValue(DAY(d)).setFontWeight("bold").setFontSize(16);
  sh.getRange("B1").setValue("Datum:");
  typed(sh.getRange("C1").setNumberFormat("d.m.yyyy"));
  sh.getRange("A2:A5").setValues([["Hotovost v kase"], ["Na účtu (QR)"], ["Peníze celkem"], ["Tržba dne"]]);
  const cash = "SUMIF(" + P + "$B$2:$B$" + MOVES_LAST + ";" + d + ";" + P + "$I$2:$I$" + MOVES_LAST + ")+SUM(I" + FIRST + ":I" + LAST + ")";
  const bank = "SUMIF(" + P + "$B$2:$B$" + MOVES_LAST + ";" + d + ";" + P + "$J$2:$J$" + MOVES_LAST + ")+SUM(J" + FIRST + ":J" + LAST + ")";
  sh.getRange("C2").setFormula("=" + (prev ? prev + "C2+" : "") + cash);
  sh.getRange("C3").setFormula("=" + (prev ? prev + "C3+" : "") + bank);
  sh.getRange("C4").setFormula("=C2+C3");
  sh.getRange("C5").setFormula("=SUM(I" + FIRST + ":J" + LAST + ")");
  sh.getRange("A2:C5").setFontWeight("bold"); auto(sh.getRange("C2:C5")); money(sh.getRange("C2:C5"));
  sh.getRange("E2:E5").setValues([
    ["Vzorec: " + (prev ? "hotovost na konci " + DAY(d - 1) + " + " : "") + "hotovost z Půjčky pro tento den (SUMIF) + SUM sloupce Hotově"],
    ["Vzorec: " + (prev ? "účet na konci " + DAY(d - 1) + " + " : "") + "účet z Půjčky pro tento den (SUMIF) + SUM sloupce QR"],
    ["Vzorec: Hotovost v kase + Na účtu"],
    ["Vzorec: SUM sloupců Hotově a QR (co jsme dnes utržili)"]]);
  sh.getRange("E2:E5").setFontStyle("italic");

  const hdr = ["Číslo (PPD)", "Čas", "Produkt", "Počet", "Cena za kus", "Celkem", "Spropitné", "Platba (prázdné = hotově)", "Hotově", "QR", "Poznámka"];
  sh.getRange(6, 1, 1, hdr.length).setValues([hdr]); head(sh.getRange(6, 1, 1, hdr.length));
  const f = [];
  for (let r = FIRST; r <= LAST; r++) {
    f.push([
      '=IF(C' + r + '="";"";IFERROR(VLOOKUP(C' + r + ';' + S.products + '!$A$2:$B$' + PRODUCTS_LAST + ';2;FALSE);0))',
      '=IF(C' + r + '="";"";D' + r + '*E' + r + ')',
    ]);
  }
  sh.getRange(FIRST, 5, f.length, 2).setFormulas(f);
  const g = [];
  for (let r = FIRST; r <= LAST; r++) {
    g.push([
      '=IF(AND(C' + r + '="";G' + r + '="");"";IF(H' + r + '="QR";0;N(F' + r + ')+N(G' + r + ')))',
      '=IF(AND(C' + r + '="";G' + r + '="");"";IF(H' + r + '="QR";N(F' + r + ')+N(G' + r + ');0))',
    ]);
  }
  sh.getRange(FIRST, 9, g.length, 2).setFormulas(g);
  const rows = LAST - FIRST + 1;
  typed(sh.getRange(FIRST, 1, rows, 4)); typed(sh.getRange(FIRST, 7, rows, 2)); typed(sh.getRange(FIRST, 11, rows, 1));
  auto(sh.getRange(FIRST, 5, rows, 2)); auto(sh.getRange(FIRST, 9, rows, 2));
  money(sh.getRange(FIRST, 5, rows, 3)); money(sh.getRange(FIRST, 9, rows, 2));
  sh.getRange(FIRST, 3, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(S.products).getRange("A2:A" + PRODUCTS_LAST), true).build());
  sh.getRange(FIRST, 8, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(["Hotově", "QR"], true).build());
  sh.getRange("G1").setValue("Prodej = vyplňte jen Produkt a Počet (žlutě). Víc produktů na jednom paragonu = víc řádků se stejným číslem.");
  sh.getRange("G2").setValue("Cena za kus = VLOOKUP z listu Produkty, Celkem = Počet × Cena, Hotově/QR = IF podle Platby.");
  sh.getRange("G1:G2").setFontStyle("italic");
  sh.setColumnWidth(3, 170); sh.setColumnWidth(8, 120); sh.setColumnWidth(11, 200); sh.setFrozenRows(6);
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
        req.ops.forEach(op => {
          if (op.type === "sale") writeSale(ss, op, skipped);
          else if (op.type === "entry") writeMove(ss, op, skipped);
          else if (op.type === "deleteEntry") deleteMove(ss, op.number, skipped);
          else throw new Error("Unknown operation: " + op.type);
        });
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

// Products with price and Zbývá from Sklad (null when Naskladněno is empty)
function readProducts(ss) {
  const pr = sheet(ss, S.products).getRange(2, 1, PRODUCTS_LAST - 1, 2).getValues();
  const st = sheet(ss, S.stock).getRange(2, 2, PRODUCTS_LAST - 1, 4 + DAYS).getValues();
  const out = [];
  pr.forEach((r, i) => {
    if (String(r[0]).trim() === "") return;
    const left = st[i][st[i].length - 1];
    out.push({name: String(r[0]).trim(), price: toNumber(r[1]), stock: st[i][0] === "" || left === "" ? null : toNumber(left)});
  });
  return out;
}

function usedNumbers(ss) {
  const out = {PPD: 0, VPD: 0, ID: 0};
  const take = v => { const m = norm(v).match(/^(PPD|VPD|ID)(\d+)$/); if (m) out[m[1]] = Math.max(out[m[1]], +m[2]); };
  for (let d = 1; d <= DAYS; d++) sheet(ss, DAY(d)).getRange(FIRST, 1, LAST - FIRST + 1, 1).getValues().forEach(r => take(r[0]));
  sheet(ss, S.moves).getRange(2, 3, MOVES_LAST - 1, 1).getValues().forEach(r => take(r[0]));
  return out;
}

// The day sheet with this date in C1, else the first one without a date (it gets the date when `claim`)
function daySheet(ss, iso, claim) {
  const want = toDate(iso).toDateString();
  let free = null;
  for (let d = 1; d <= DAYS; d++) {
    const sh = sheet(ss, DAY(d));
    const v = sh.getRange("C1").getValue();
    if (v instanceof Date && v.toDateString() === want) return {sh: sh, n: d};
    if (v === "" && !free) free = {sh: sh, n: d};
  }
  if (!free) throw new Error("Both day sheets already have another date - change the date in C1 of a day sheet.");
  if (claim) free.sh.getRange("C1").setValue(toDate(iso));
  return free;
}

function dayInfo(ss, iso) {
  if (!iso) return null;
  let day;
  try { day = daySheet(ss, iso, false); } catch (e) { return null; }
  const v = day.sh.getRange("C2:C5").getValues().map(r => toNumber(r[0]));
  return {name: DAY(day.n), cash: v[0], bank: v[1], total: v[2], sales: v[3]};
}

// ---- a sale = one row per product:  Číslo | Čas | Produkt | Počet | (Cena, Celkem) | Spropitné | Platba | (Hotově, QR) | Poznámka
function writeSale(ss, op, skipped) {
  // rows that already have this number (anywhere): only the app's own rows may be changed
  const existing = [];
  for (let d = 1; d <= DAYS; d++) {
    const sh = sheet(ss, DAY(d));
    const a = sh.getRange(FIRST, 1, LAST - FIRST + 1, 1).getValues();
    a.forEach((r, i) => { if (norm(r[0]) === norm(op.number)) existing.push({sh: sh, row: FIRST + i}); });
  }
  if (existing.some(x => x.sh.getRange(x.row, 1).getNote() !== APP_NOTE)) { skipped.push(op.number); return; }
  const pay = op.method === "QR" ? "QR" : "Hotově";
  if (op.void) {
    existing.forEach(x => {
      x.sh.getRange(x.row, 4).setValue(0);
      x.sh.getRange(x.row, 7).setValue("");
      x.sh.getRange(x.row, 11).setValue("STORNO");
    });
    return;
  }
  existing.forEach(x => { clearSaleRow(x.sh, x.row); });
  const sh = existing.length ? existing[0].sh : daySheet(ss, op.date, true).sh;
  const items = op.items.length ? op.items : [{name: "", qty: 0}];
  const tip = Number(op.tip) || 0;
  const a = sh.getRange(FIRST, 1, LAST - FIRST + 1, 1).getValues();
  const c = sh.getRange(FIRST, 3, LAST - FIRST + 1, 1).getValues();
  let i = 0;
  items.forEach((it, k) => {
    while (i < a.length && (String(a[i][0]).trim() !== "" || String(c[i][0]).trim() !== "")) i++;
    if (i >= a.length) throw new Error(sh.getName() + " is full.");
    const row = FIRST + i;
    sh.getRange(row, 1, 1, 4).setValues([[op.number, op.time, it.name, it.qty]]);
    sh.getRange(row, 7, 1, 2).setValues([[k === 0 && tip ? tip : "", pay]]);
    sh.getRange(row, 11).setValue("");
    sh.getRange(row, 1).setNote(APP_NOTE);
    a[i][0] = op.number;
    i++;
  });
}

function clearSaleRow(sh, row) {
  sh.getRange(row, 1, 1, 4).clearContent();
  sh.getRange(row, 7, 1, 2).clearContent();
  sh.getRange(row, 11).clearContent();
  sh.getRange(row, 1).setNote("");
}

// ---- money movements -> Půjčky:  Datum | Den | Číslo | Typ | Kdo | Popis | Částka | Platba | (formulas)
function writeMove(ss, op, skipped) {
  const sh = sheet(ss, S.moves);
  let row = findRow(sh, 3, op.number);
  if (row && sh.getRange(row, 3).getNote() !== APP_NOTE) { skipped.push(op.number); return; }
  if (!row) {
    const v = sh.getRange(2, 3, MOVES_LAST - 1, 2).getValues();
    const i = v.findIndex(r => String(r[0]).trim() === "" && String(r[1]).trim() === "");
    if (i === -1) throw new Error("Půjčky is full.");
    row = 2 + i;
  }
  const day = daySheet(ss, op.date, true);
  sh.getRange(row, 1, 1, 8).setValues([[toDate(op.date), day.n, op.number, op.kind, op.person || "", op.description || "",
                                        op.amount, op.payment === "Účet" ? "Účet" : "Hotově"]]);
  sh.getRange(row, 3).setNote(APP_NOTE);
}

function deleteMove(ss, number, skipped) {
  const sh = sheet(ss, S.moves);
  const row = findRow(sh, 3, number);
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

function norm(v) { return String(v).replace(/\s+/g, "").toUpperCase(); }

function findRow(sh, col, value) {
  const v = sh.getRange(2, col, MOVES_LAST - 1, 1).getValues();
  const i = v.findIndex(r => norm(r[0]) === norm(value));
  return i === -1 ? null : 2 + i;
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
