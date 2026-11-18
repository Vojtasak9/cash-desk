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
 *   Den 1..Den 4  one row per product sold: Produkt + Počet + Platba -> price, total, cash/QR by formulas;
 *                 every sale is a paragon P001.. (cash and QR); end of day: E1 = PPD (cash total), E2 = VBU (QR total)
 *   FULL          everything in time order (PPD, VPD, P, VBU, ID) - only formulas, nobody types here
 *   Kuchyně       the orders for the cook's phone (kitchen.html): what to make, with toppings, and when it was done
 *
 * repairFormulas() writes all formulas again without touching typed data.
 *
 * The spreadsheet uses Czech settings, so formulas use ";" between arguments.
 * rebuildViper() deletes all sheets and builds them again (run it from the editor).
 */

// Pick your own password, put it here AND type the same one into the app.
const KEY = "ZMEN-ME";

const VIPER_ID = "1ZsqWHUY1iF5czN0nYmrUwboyXHMJ65mTZn6niF1sxLk";

const S = {products: "Produkty", stock: "Sklad", types: "Typy pohybů", moves: "Půjčky", full: "FULL", kitchen: "Kuchyně"};
const FULL_LAST = 1800;
const DAY = d => "Den " + d;
const DAYS = 4;
// Produkty column F (after the day prices): "" = product, "přísada" = topping put on a product, "skrýt" = not offered in the app
const KIND_COL = 2 + DAYS;
const KIND_HEAD = "Druh (prázdné = produkt, přísada, skrýt)";
const KINDS_LIST = ["přísada", "skrýt"];
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
  ["Nákup na dluh", 0, 1, "Kupní smlouva: někdo nám něco prodal a ještě jsme nezaplatili. Peníze se nehýbou (číslo ID), kasa mu dluží. Až zaplatíme: Vrácení dluhu."],
];

const BLUE = "#dfe7f5", YELLOW = "#fff6db", GREY = "#f1f3f4";

// ======================================================================= build the sheets

function rebuildViper() {
  const ss = SpreadsheetApp.openById(VIPER_ID);
  const oldProducts = ss.getSheetByName(S.products)
    ? ss.getSheetByName(S.products).getRange(2, 1, PRODUCTS_LAST - 1, 1 + DAYS).getValues().filter(r => String(r[0]).trim() !== "")
      .map(r => r.map((v, i) => i > 1 && v === "" ? r[1] : v)) : [];  // old sheet with one price -> same price every day
  const tmp = ss.insertSheet("tmp" + Date.now());
  ss.getSheets().forEach(sh => { if (sh.getSheetId() !== tmp.getSheetId()) ss.deleteSheet(sh); });
  SpreadsheetApp.flush();
  Utilities.sleep(2000);  // Sheets needs a moment before a deleted sheet's name can be used again
  buildProducts(ss);
  if (oldProducts.length) {  // keep the products that were there
    const sh = ss.getSheetByName(S.products);
    sh.getRange(2, 1, PRODUCTS_LAST - 1, 1 + DAYS).clearContent();
    sh.getRange(2, 1, oldProducts.length, 1 + DAYS).setValues(oldProducts);
  }
  buildStock(ss);
  buildTypes(ss);
  buildMoves(ss);
  for (let d = 1; d <= DAYS; d++) buildDay(ss, d);
  buildMovesForFull(ss);
  buildFull(ss);
  ss.deleteSheet(tmp);
  ss.setActiveSheet(ss.getSheetByName(DAY(1)));
  SpreadsheetApp.flush();
}

// Adds the missing day sheets (up to DAYS) and redoes Sklad, keeping everything already typed in
function addMissingDays() {
  const ss = SpreadsheetApp.openById(VIPER_ID);
  for (let d = 1; d <= DAYS; d++) if (!ss.getSheetByName(DAY(d))) buildDay(ss, d);
  const old = ss.getSheetByName(S.stock);
  const stocked = old ? old.getRange(2, 2, PRODUCTS_LAST - 1, 1).getValues() : null;
  if (old) { old.setName("old" + Date.now()); SpreadsheetApp.flush(); }
  buildStock(ss);
  const sh = ss.getSheetByName(S.stock);
  if (stocked) sh.getRange(2, 2, PRODUCTS_LAST - 1, 1).setValues(stocked);
  ss.setActiveSheet(sh); ss.moveActiveSheet(2);
  if (old) ss.deleteSheet(old);
  SpreadsheetApp.flush();
}

const money = r => r.setNumberFormat('#,##0 "Kč"');
const head = r => r.setFontWeight("bold").setBackground(BLUE).setWrap(true).setVerticalAlignment("middle");
const typed = r => r.setBackground(YELLOW);   // you type here
const auto = r => r.setBackground(GREY);      // a formula fills this

function buildProducts(ss) {
  const sh = ss.insertSheet(S.products);
  // one price column per day: column B = Den 1, C = Den 2, ...
  const hdr = ["Produkt"]; for (let d = 1; d <= DAYS; d++) hdr.push("Cena " + DAY(d));
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]); head(sh.getRange(1, 1, 1, hdr.length));
  const start = [["Standard", 35], ["Mini", 25], ["Hotdog", 70]].map(r => [r[0]].concat(Array(DAYS).fill(r[1])));
  sh.getRange(2, 1, start.length, hdr.length).setValues(start);
  typed(sh.getRange(2, 1, PRODUCTS_LAST - 1, hdr.length)); money(sh.getRange(2, 2, PRODUCTS_LAST - 1, DAYS));
  kindColumn(sh);
  const note = hdr.length + 3;
  sh.getRange(1, note, 4, 1).setValues([["Žlutě = píšete vy."], ["Šedě = počítá vzorec."],
    ["Každý den může mít jinou cenu: napište ji do sloupce toho dne."],
    ["Nový produkt napište na první volný řádek. Objeví se sám v Sklad i v nabídce v listech Den 1 až Den " + DAYS + "."]]);
  sh.setColumnWidth(1, 180); sh.setColumnWidth(note, 520); sh.setFrozenRows(1);
}

// Produkty column F: is it a product, a topping (přísada) or hidden from the app (skrýt)
function kindColumn(sh) {
  head(sh.getRange(1, KIND_COL).setValue(KIND_HEAD));
  typed(sh.getRange(2, KIND_COL, PRODUCTS_LAST - 1, 1)).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(KINDS_LIST, true).setAllowInvalid(false).build());
  sh.setColumnWidth(KIND_COL, 170);
}

// ONE TIME (18.11.2026): toppings instead of "Addon". Addon stays (Den 1 and Den 2 use its price) but is hidden in the app.
function setupToppings() {
  const ss = SpreadsheetApp.openById(VIPER_ID);
  const sh = sheet(ss, S.products);
  kindColumn(sh);
  const v = sh.getRange(2, 1, PRODUCTS_LAST - 1, KIND_COL).getValues();
  const log = [];
  v.forEach((r, i) => { if (norm(r[0]) === "ADDON" && r[KIND_COL - 1] === "") { sh.getRange(2 + i, KIND_COL).setValue("skrýt"); log.push("Addon: skrýt"); } });
  const want = [["Kečup", 0], ["Hořčice", 0], ["Cibulka", 5], ["Jalapeños", 5], ["Okurky", 5]];
  want.forEach(w => {
    const have = v.findIndex(r => norm(r[0]) === norm(w[0]));
    if (have !== -1) { sh.getRange(2 + have, KIND_COL).setValue("přísada"); log.push(w[0] + ": už je, přísada"); return; }
    const i = v.findIndex(r => String(r[0]).trim() === "");
    if (i === -1) throw new Error("Produkty is full.");
    const row = [w[0]].concat(Array(DAYS).fill(w[1]), ["přísada"]);
    sh.getRange(2 + i, 1, 1, KIND_COL).setValues([row]);
    v[i] = row;
    log.push(w[0] + " " + w[1] + " Kč: řádek " + (2 + i));
  });
  sh.getRange(1, KIND_COL + 1, 6, 1).clearContent();
  sh.getRange(1, KIND_COL + 2, 6, 1).setValues([["Žlutě = píšete vy."], ["Šedě = počítá vzorec."],
    ["Každý den může mít jinou cenu: napište ji do sloupce toho dne."],
    ["Nový produkt napište na první volný řádek. Objeví se sám v Sklad i v nabídce v listech Den 1 až Den " + DAYS + "."],
    ["Druh: prázdné = produkt (tlačítko v pokladně), přísada = dává se na produkt (kečup, cibulka…), skrýt = pokladna ho nenabízí."],
    ["Přísada se v listu dne zapíše jako vlastní řádek (např. Cibulka 1 ks × 5 Kč), takže tržba i Sklad sedí."]]);
  sh.setColumnWidth(KIND_COL + 2, 520);
  kitchenSheet(ss);
  SpreadsheetApp.flush();
  Logger.log(log.join("\n"));
  return log;
}

function buildStock(ss) {
  const sh = ss.insertSheet(S.stock);
  const cols = ["Produkt", "Naskladněno (ks)"];
  for (let d = 1; d <= DAYS; d++) cols.push("Prodáno " + DAY(d));
  cols.push("Prodáno celkem", "Zbývá (ks)");
  sh.getRange(1, 1, 1, cols.length).setValues([cols]); head(sh.getRange(1, 1, 1, cols.length));
  const n = PRODUCTS_LAST - 1, totalCol = 3 + DAYS, leftCol = totalCol + 1;
  stockFormulas(sh);
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

function stockFormulas(sh) {
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
  sh.getRange("A1:K1").setValues([["Datum", "Den (1 až " + DAYS + ")", "Číslo", "Typ (co se stalo)", "Kdo", "Popis", "Částka",
                                   "Platba (prázdné = hotově)", "Hotovost +/-", "Účet +/-", "Kasa mu dluží +/-"]]);
  head(sh.getRange("A1:K1"));
  sh.getRange("M1:N1").setValues([["Kdo", "Kasa mu dluží (+) / dluží kase (−)"]]); head(sh.getRange("M1:N1"));
  movesFormulas(sh);
  typed(sh.getRange("A2:H" + MOVES_LAST)); auto(sh.getRange("I2:K" + MOVES_LAST));
  money(sh.getRange("G2:G" + MOVES_LAST)); money(sh.getRange("I2:K" + MOVES_LAST));
  sh.getRange("A2:A" + MOVES_LAST).setNumberFormat("d.m.yyyy");
  sh.getRange("D2:D" + MOVES_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(S.types).getRange("A3:A" + (TYPES.length + 2)), true).build());
  sh.getRange("H2:H" + MOVES_LAST).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(["Hotově", "Účet"], true).build());
  sh.getRange("B2:H2").setValues([[1, "PPD001", "Zapůjčka do kasy", "Kristián", "Peníze na start (vrátit)", 500, "Hotově"]]);
  auto(sh.getRange("M2:N30")); money(sh.getRange("N2:N30"));
  sh.getRange("M32:M35").setValues([
    ["Jak to funguje:"],
    ["Hotovost = Částka × číslo z listu Typy pohybů (VLOOKUP podle Typu)."],
    ["Kasa mu dluží = Částka × druhé číslo z Typy pohybů."],
    ["Vpravo nahoře: SUMIF sečte, kolik kasa dluží každému."]]);
  sh.getRange("M32").setFontWeight("bold");
  sh.setColumnWidth(4, 160); sh.setColumnWidth(6, 220); sh.setColumnWidth(13, 140); sh.setColumnWidth(14, 220); sh.setFrozenRows(1);
}

// Půjčky formulas: cash / account / debt (VLOOKUP into Typy pohybů) and the who-owes-whom table.
// One ARRAYFORMULA per column in row 2 (not one formula per row): inserting, moving or deleting rows
// by hand can't leave a row without its formula or pointing at the wrong row.
const colRange = (c, last) => "$" + c + "$2:$" + c + "$" + (last || MOVES_LAST);
function movesFormulas(sh) {
  const T = "'" + S.types + "'!$A$3:$C$" + (TYPES.length + 2);
  const [D, G, H] = ["D", "G", "H"].map(c => colRange(c));
  sh.getRange("I2:K" + MOVES_LAST).clearContent();
  sh.getRange("I2:K2").setFormulas([[
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + H + '="Účet";0;' + G + '*VLOOKUP(' + D + ';' + T + ';2;FALSE))))',
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + H + '="Účet";' + G + '*VLOOKUP(' + D + ';' + T + ';2;FALSE);0)))',
    '=ARRAYFORMULA(IF(' + D + '="";"";' + G + '*VLOOKUP(' + D + ';' + T + ';3;FALSE)))',
  ]]);
  sh.getRange("M2:N30").clearContent();
  sh.getRange("M2").setFormula('=IFERROR(UNIQUE(FILTER(E2:E' + MOVES_LAST + ';E2:E' + MOVES_LAST + '<>""));"")');
  sh.getRange("N2").setFormula('=ARRAYFORMULA(IF(M2:M30="";"";SUMIF(' + colRange("E") + ';M2:M30;' + colRange("K") + ')))');
}

// Půjčky columns P-X: the same rows prepared for the FULL sheet (built after the day sheets, it reads their dates)
// A row counts as soon as it has a Typ (like the day sheets' SUMIF), even without a number.
function buildMovesForFull(ss) {
  const sh = sheet(ss, S.moves);
  sh.getRange("P1:X1").setValues([FULL_HEAD.slice(0, 9).map(h => "pro FULL: " + h)]); auto(sh.getRange("P1:X1"));
  const [A, B, C, D, E, F, I, J, P] = ["A", "B", "C", "D", "E", "F", "I", "J", "P"].map(c => colRange(c));
  // no Datum typed -> the date of that day's sheet (column B = day number)
  let dayDate = "NA()";
  for (let d = DAYS; d >= 1; d--) dayDate = "IF(" + B + "=" + d + ";'" + DAY(d) + "'!$C$1;" + dayDate + ")";
  const on = x => '=ARRAYFORMULA(IF(' + P + '="";"";' + x + '))';
  const plus = x => 'IF(ISNUMBER(' + x + ')*(' + x + '>0);' + x + ';0)', minus = x => 'IF(ISNUMBER(' + x + ')*(' + x + '<0);-' + x + ';0)';
  sh.getRange("P2:X" + MOVES_LAST).clearContent();
  sh.getRange("P2:X2").setFormulas([[
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + A + '<>"";INT(' + A + ');IFERROR(INT(' + dayDate + ');0))+ROW(' + D + ')/10000000))',
    on('IF(' + C + '="";"(bez čísla)";' + C + ')'),
    on(D + '&IF(' + E + '="";"";" - "&' + E + ')&IF(' + F + '="";"";": "&' + F + ')'), '', '',
    on(plus(I)), on(minus(I)), on(plus(J)), on(minus(J))]]);
  auto(sh.getRange(2, 16, MOVES_LAST - 1, 9)); sh.getRange(2, 16, MOVES_LAST - 1, 1).setNumberFormat("d.m. h:mm");
}

const FULL_HEAD = ["Datum a čas", "Číslo dokladu", "Popis", "Cena za kus", "Počet kusů",
                   "Pokladna - příjmy", "Pokladna - výdaje", "Účet - příjmy", "Účet - výdaje",
                   "Celkem (průběžný zůstatek)", "V kase (průběžně)", "Na účtu (průběžně)"];

function buildDay(ss, d) {
  const sh = ss.insertSheet(DAY(d));
  dayLayout(ss, sh, d);
  dayFormulas(sh, d);
}

// Labels, colours and checks of a day sheet (no formulas, typed values stay)
//   row 1: Datum | Uzávěrka PPD (hotově): E1 = number, F1 = SUM Hotově
//   row 2:         Uzávěrka VBU (QR):     E2 = number, F2 = SUM QR
//   rows 2-5: cash, account, all money, sales of the day;  row 6 header;  sales from row 7
function dayLayout(ss, sh, d) {
  const prev = d > 1 ? DAY(d - 1) : "";
  sh.getRange("A1").setValue(DAY(d)).setFontWeight("bold").setFontSize(16);
  sh.getRange("B1").setValue("Datum:");
  typed(sh.getRange("C1")).setNumberFormat("d.m.yyyy");
  sh.getRange("D1:D2").setValues([["Uzávěrka PPD (hotově):"], ["Uzávěrka VBU (QR):"]]).setFontWeight("bold");
  typed(sh.getRange("E1:E2")).setFontStyle("normal");
  auto(sh.getRange("F1:F2")); money(sh.getRange("F1:F2"));
  sh.getRange("E3:E5").clearContent();
  sh.getRange("A2:A5").setValues([["Hotovost v kase"], ["Na účtu (QR)"], ["Peníze celkem"], ["Tržba dne"]]);
  sh.getRange("A2:C5").setFontWeight("bold"); auto(sh.getRange("C2:C5")); money(sh.getRange("C2:C5"));
  sh.getRange("G1:G5").setValues([
    ["Prodej = Číslo (P001, P002… pro hotovost i QR), Produkt, Počet a Platba (žlutě). Víc produktů na jednom dokladu = víc řádků se stejným číslem."],
    ["Cena = VLOOKUP z Produkty (sloupec Cena " + DAY(d) + "), Celkem = Počet × Cena. Prázdná Platba = hotově."],
    ["Uzávěrka dne: PPD = F1 (součet sloupce Hotově) jde do kasy, VBU = F2 (součet sloupce QR) jde na účet."],
    ["Hotovost v kase = " + (prev ? "konec " + prev + " + " : "") + "Půjčky pro tento den (SUMIF) + SUM Hotově. Na účtu = totéž se sloupcem QR."],
    ["Peníze celkem = Hotovost v kase + Na účtu. Tržba dne = SUM sloupců Hotově a QR."]]).setFontStyle("italic");

  const hdr = ["Číslo (P001…)", "Čas", "Produkt", "Počet", "Cena za kus", "Celkem", "Spropitné", "Platba (Hotově / QR)", "Hotově", "QR", "Poznámka"];
  sh.getRange(6, 1, 1, hdr.length).setValues([hdr]); head(sh.getRange(6, 1, 1, hdr.length));
  const rows = LAST - FIRST + 1;
  typed(sh.getRange(FIRST, 1, rows, 4)); typed(sh.getRange(FIRST, 7, rows, 2)); typed(sh.getRange(FIRST, 11, rows, 1));
  auto(sh.getRange(FIRST, 5, rows, 2)); auto(sh.getRange(FIRST, 9, rows, 2));
  money(sh.getRange(FIRST, 5, rows, 3)); money(sh.getRange(FIRST, 9, rows, 2));
  sh.getRange(FIRST, 3, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(S.products).getRange("A2:A" + PRODUCTS_LAST), true).build());
  sh.getRange(FIRST, 8, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(["Hotově", "QR"], true).build());
  sh.setColumnWidth(1, 120); sh.setColumnWidth(3, 170); sh.setColumnWidth(4, 150); sh.setColumnWidth(8, 120); sh.setColumnWidth(11, 200);
  sh.setFrozenRows(6);
  sh.getRange("R6:Z6").setValues([FULL_HEAD.slice(0, 9).map(h => "pro FULL: " + h)]);
  auto(sh.getRange(4, 18, LAST - 3, 9)); sh.getRange(4, 18, LAST - 3, 1).setNumberFormat("d.m. h:mm");
}

// All formulas of a day sheet (also repairs them when someone overwrote or deleted some)
function dayFormulas(sh, d) {
  const P = "'" + S.moves + "'!";
  const prev = d > 1 ? "'" + DAY(d - 1) + "'!" : "";
  sh.getRange("F1").setFormula("=SUM(I" + FIRST + ":I" + LAST + ")");
  sh.getRange("F2").setFormula("=SUM(J" + FIRST + ":J" + LAST + ")");
  const cash = "SUMIF(" + P + "$B$2:$B$" + MOVES_LAST + ";" + d + ";" + P + "$I$2:$I$" + MOVES_LAST + ")+SUM(I" + FIRST + ":I" + LAST + ")";
  const bank = "SUMIF(" + P + "$B$2:$B$" + MOVES_LAST + ";" + d + ";" + P + "$J$2:$J$" + MOVES_LAST + ")+SUM(J" + FIRST + ":J" + LAST + ")";
  sh.getRange("C2").setFormula("=" + (prev ? prev + "C2+" : "") + cash);
  sh.getRange("C3").setFormula("=" + (prev ? prev + "C3+" : "") + bank);
  sh.getRange("C4").setFormula("=C2+C3");
  sh.getRange("C5").setFormula("=SUM(I" + FIRST + ":J" + LAST + ")");
  closeNumbers(sh, d);
  const f = [], g = [], h = [];
  for (let r = FIRST; r <= LAST; r++) {
    f.push([
      '=IF(C' + r + '="";"";IFERROR(VLOOKUP(C' + r + ';' + S.products + '!$A$2:$' + String.fromCharCode(65 + DAYS) + '$' + PRODUCTS_LAST + ';' + (1 + d) + ';FALSE);0))',
      '=IF(C' + r + '="";"";D' + r + '*E' + r + ')',
    ]);
    g.push([
      '=IF(AND(C' + r + '="";G' + r + '="");"";IF(H' + r + '="QR";0;N(F' + r + ')+N(G' + r + ')))',
      '=IF(AND(C' + r + '="";G' + r + '="");"";IF(H' + r + '="QR";N(F' + r + ')+N(G' + r + ');0))',
    ]);
    // FULL: every P row carries its money (Hotově -> Pokladna příjmy, QR -> Účet příjmy, tip included)
    const on = 'IF($R' + r + '="";"";';
    h.push(['=IF($C' + r + '="";"";INT($C$1)+IF(ISNUMBER($B' + r + ');MOD($B' + r + ';1);IFERROR(TIMEVALUE($B' + r + ');1/2))+ROW()/10000000)',
            '=' + on + '$A' + r + ')',
            '=' + on + '$C' + r + '&IF($H' + r + '="QR";" - prodej (QR)";" - prodej (hotově)")&IF(N($G' + r + ')=0;"";" + spropitné "&$G' + r + '&" Kč"))',
            '=' + on + '$E' + r + ')', '=' + on + '$D' + r + ')',
            '=' + on + 'IF($H' + r + '="QR";0;N($I' + r + ')))', '=' + on + '0)',
            '=' + on + 'IF($H' + r + '="QR";N($J' + r + ');0))', '=' + on + '0)']);
  }
  sh.getRange(FIRST, 5, f.length, 2).setFormulas(f);
  sh.getRange(FIRST, 9, g.length, 2).setFormulas(g);
  sh.getRange(FIRST, 18, h.length, 9).setFormulas(h);
  // rows 4 and 5: the day's VBU (QR total) and PPD (cash total) - only a label in FULL, the money was counted on the P rows
  const zero = c => '=IF(' + c + '="";"";0)';
  sh.getRange("R4:Z5").setFormulas([
    ['=IF(AND(N($F$2)=0;$E$2="");"";INT($C$1)+1-2/100000)', '=IF(R4="";"";IF($E$2="";"VBU (den neuzavřen)";$E$2))',
     '=IF(R4="";"";"Uzávěrka QR za ' + DAY(d) + ': "&$F$2&" Kč (součet QR paragonů P, už započteno výše)")', '', '',
     zero("R4"), zero("R4"), zero("R4"), zero("R4")],
    ['=IF(AND(N($F$1)=0;$E$1="");"";INT($C$1)+1-1/100000)', '=IF(R5="";"";IF($E$1="";"PPD (den neuzavřen)";$E$1))',
     '=IF(R5="";"";"Uzávěrka hotovosti za ' + DAY(d) + ': "&$F$1&" Kč (součet hotovostních paragonů P, už započteno výše)")', '', '',
     zero("R5"), zero("R5"), zero("R5"), zero("R5")]]);
}

// The day's closing numbers count themselves, in time order: E1 = PPD, E2 = VBU.
// Number = the highest PPD/VBU in Půjčky dated up to this day or in an earlier day's close, + 1.
// (The close is at the end of the day, so a VBU typed into Půjčky for that day comes before it.)
// Shown only when the day has cash (F1) / QR (F2) payments.
function closeNumbers(sh, d) {
  const P = "'" + S.moves + "'!$P$2:$P$" + MOVES_LAST, Q = "'" + S.moves + "'!$Q$2:$Q$" + MOVES_LAST;
  const f = (pfx, cell, sum) => {
    let prev = "";
    for (let k = 1; k < d; k++) prev += ";IFERROR(VALUE(MID('" + DAY(k) + "'!" + cell + ";4;9));0)";
    return '=IF(N(' + sum + ')=0;"";"' + pfx + '"&TEXT(MAX(0' + prev + ';IFERROR(ARRAYFORMULA(MAX(IF(LEFT(' + Q + ';3)="' + pfx + '";' +
      'IF(ISNUMBER(' + P + ');IF(' + P + '<INT($C$1)+1;IFERROR(VALUE(MID(' + Q + ';4;9));0);0);0);0)));0))+1;"000"))';
  };
  sh.getRange("E1:E2").setFormulas([[f("PPD", "$E$1", "$F$1")], [f("VBU", "$E$2", "$F$2")]])
    .setNote("Číslo se počítá samo (další PPD/VBU po Půjčkách do tohoto dne a po předchozích dnech). Nepřepisujte.");
}

// FULL: everything from Půjčky and the day sheets in time order. Only formulas - nobody types here.
function buildFull(ss) {
  const sh = ss.insertSheet(S.full);
  sh.getRange(1, 1, 1, FULL_HEAD.length).setValues([FULL_HEAD]); head(sh.getRange(1, 1, 1, FULL_HEAD.length));
  fullFormulas(sh);
  auto(sh.getRange(2, 1, FULL_LAST - 1, FULL_HEAD.length));
  sh.getRange(2, 1, FULL_LAST - 1, 1).setNumberFormat("d.m. h:mm");
  money(sh.getRange(2, 4, FULL_LAST - 1, 1)); money(sh.getRange(2, 6, FULL_LAST - 1, 7));
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 380); sh.setColumnWidth(14, 600); sh.setFrozenRows(1);
  sh.protect().setDescription("FULL se počítá sám").setWarningOnly(true);
}

function fullFormulas(sh) {
  const parts = [], keys = [];
  for (let d = 1; d <= DAYS; d++) {
    const q = "'" + DAY(d) + "'!";
    parts.push(q + "R4:Z5", q + "R" + FIRST + ":Z" + LAST);
    keys.push(q + "R4:R5", q + "R" + FIRST + ":R" + LAST);
  }
  const m = "'" + S.moves + "'!";
  parts.push(m + "P2:X" + MOVES_LAST); keys.push(m + "P2:P" + MOVES_LAST);
  sh.getRange("A2:I" + FULL_LAST).clearContent();
  sh.getRange("A2").setFormula('=IFERROR(SORT(FILTER(VSTACK(' + parts.join(";") + ');VSTACK(' + keys.join(";") + ')<>"");1;TRUE);"")');
  const f = [];
  for (let r = 2; r <= FULL_LAST; r++) {
    const on = '=IF($A' + r + '="";"";';
    f.push([on + 'N(J' + (r - 1) + ')+N(F' + r + ')-N(G' + r + ')+N(H' + r + ')-N(I' + r + '))',
            on + 'N(K' + (r - 1) + ')+N(F' + r + ')-N(G' + r + '))',
            on + 'N(L' + (r - 1) + ')+N(H' + r + ')-N(I' + r + '))']);
  }
  sh.getRange(2, 10, f.length, 3).setFormulas(f);
  sh.getRange("N1:N7").setValues([
    ["Tento list se vyplňuje sám - nic sem nepište."],
    ["A2: SORT(FILTER(VSTACK(...))) spojí řádky z Půjčky a Den 1-" + DAYS + " a seřadí je podle času."],
    ["P = paragon (každý prodej). Příjmy = Počet × Cena (+ spropitné): hotově do Pokladna příjmy, QR do Účet příjmy."],
    ["Uzávěrka dne (PPD hotovost, VBU QR) je tu jen jako popisek se součtem - peníze už jsou započtené u P. VPD = výdej z kasy. ID = bez peněz."],
    ["Celkem = Celkem o řádek výš + Pokladna příjmy - Pokladna výdaje + Účet příjmy - Účet výdaje"],
    ["V kase = předchozí + Pokladna příjmy - Pokladna výdaje"],
    ["Na účtu = předchozí + Účet příjmy - Účet výdaje"]]);
  sh.getRange("N1").setFontWeight("bold");
}

// Writes every formula again (day sheets, Půjčky, Sklad, FULL) without touching what people typed.
// Run it from the editor when a formula was overwritten or deleted by accident.
function repairFormulas(ss) {
  ss = ss || SpreadsheetApp.openById(VIPER_ID);
  for (let d = 1; d <= DAYS; d++) dayFormulas(sheet(ss, DAY(d)), d);
  movesFormulas(sheet(ss, S.moves));
  buildMovesForFull(ss);
  stockFormulas(sheet(ss, S.stock));
  fullFormulas(sheet(ss, S.full));
  protectFormulas(ss);
  SpreadsheetApp.flush();
}

// Grey formula cells warn before someone types over them (they can still click "OK" to do it)
function protectFormulas(ss) {
  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => { if (p.getDescription() === "Vzorec") p.remove(); });
  const add = r => r.protect().setDescription("Vzorec").setWarningOnly(true);
  for (let d = 1; d <= DAYS; d++) {
    const sh = sheet(ss, DAY(d));
    ["C2:C5", "E1:F2", "E" + FIRST + ":F" + LAST, "I" + FIRST + ":J" + LAST, "R4:Z" + LAST].forEach(a => add(sh.getRange(a)));
  }
  const mv = sheet(ss, S.moves);
  ["I2:K" + MOVES_LAST, "M2:N30", "P2:X" + MOVES_LAST].forEach(a => add(mv.getRange(a)));
  const st = sheet(ss, S.stock);
  add(st.getRange(2, 1, PRODUCTS_LAST - 1, 1)); add(st.getRange(2, 3, PRODUCTS_LAST - 1, 2 + DAYS));
}

// ONE TIME (17.11.2026): every sale becomes a paragon P (cash and QR alike, one series in time order),
// every finished day gets PPD (its cash) + VBU (its QR), numbers in Půjčky get their proper series.
// Nothing typed is deleted. Run once from the editor; the log lists every number that changed.
function migrateToP() {
  const ss = SpreadsheetApp.openById(VIPER_ID);
  if (String(sheet(ss, DAY(1)).getRange("D2").getValue()).indexOf("VBU") !== -1) throw new Error("Už je převedeno.");
  const pad = n => String(n).padStart(3, "0");
  const log = [];
  // Den 1: the date was typed into B1 ("Datum: 12.11.2026") and C1 got a later date
  const d1 = sheet(ss, DAY(1)), m1 = String(d1.getRange("B1").getValue()).match(/(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})/);
  if (m1) { d1.getRange("C1").setValue(new Date(+m1[3], +m1[2] - 1, +m1[1], 12)); log.push("Den 1: datum " + m1[1] + "." + m1[2] + "." + m1[3]); }

  // 1. sales -> P001... in the order they are in Den 1..Den 4 (a new receipt = new number or new time)
  let n = 0;
  for (let d = 1; d <= DAYS; d++) {
    const sh = sheet(ss, DAY(d)), rows = LAST - FIRST + 1;
    const v = sh.getRange(FIRST, 1, rows, 3).getValues();
    const A = [], H = [];
    let prev = null;
    v.forEach(r => {
      const old = String(r[0]).trim();
      if (old === "") { A.push([""]); H.push([""]); prev = null; return; }
      const key = norm(old) + "|" + String(r[1]);
      if (key !== prev) {
        n++;
        if (norm(old) !== "P" + pad(n)) log.push(DAY(d) + ": " + old + " -> P" + pad(n));
      }
      prev = key;
      A.push(["P" + pad(n)]);
      H.push([/^VB[UÚ]/.test(norm(old)) ? "QR" : "Hotově"]);
    });
    sh.getRange("E2").clearContent();  // was a note, now the day's VBU
    sh.getRange(FIRST, 8, rows, 1).clearDataValidations();
    sh.getRange(FIRST, 1, rows, 1).setValues(A);
    sh.getRange(FIRST, 8, rows, 1).setValues(H);
    dayLayout(ss, sh, d);
  }

  // 2. Půjčky: "PPD 001" -> "PPD001"; a number from the wrong series gets the next one of the right series
  const mv = sheet(ss, S.moves);
  const mr = mv.getRange(2, 1, MOVES_LAST - 1, 8).getValues();
  const dirOf = {}; TYPES.forEach(t => { dirOf[t[0]] = t[1]; });
  const want = r => { const dir = dirOf[String(r[3]).trim()]; return dir === undefined ? null : dir === 0 ? "ID" : r[7] === "Účet" ? "VBU" : dir > 0 ? "PPD" : "VPD"; };
  const max = {PPD: 0, VPD: 0, VBU: 0, ID: 0};
  const parsed = mr.map(r => { const m = norm(r[2]).replace("VBÚ", "VBU").match(/^(PPD|VPD|VBU|ID)0*(\d+)$/); return m ? {p: m[1], n: +m[2]} : null; });
  parsed.forEach((x, i) => { if (x && (x.p === want(mr[i]) || !want(mr[i]))) max[x.p] = Math.max(max[x.p], x.n); });
  const C = mr.map((r, i) => {
    const x = parsed[i];
    if (!x) return [r[2]];
    const w = want(r) || x.p;
    const nu = x.p === w ? x.p + pad(x.n) : w + pad(++max[w]);
    if (String(r[2]) !== nu) log.push("Půjčky řádek " + (i + 2) + ": " + r[2] + " -> " + nu);
    return [nu];
  });
  mv.getRange(2, 3, C.length, 1).setValues(C);

  // 3. formulas everywhere (some were broken) and FULL
  repairFormulas(ss);

  // 4. finished days (before today) with sales: PPD = their cash, VBU = their QR (only when there was some)
  const today = new Date(); today.setHours(0, 0, 0, 0);
  for (let d = 1; d <= DAYS; d++) {
    const sh = sheet(ss, DAY(d)), date = sh.getRange("C1").getValue();
    const sold = sh.getRange(FIRST, 1, LAST - FIRST + 1, 1).getValues().some(r => String(r[0]).trim() !== "");
    if (!(date instanceof Date) || date >= today || !sold) continue;
    const f = sh.getRange("F1:F2").getValues(), cash = toNumber(f[0][0]), qr = toNumber(f[1][0]);
    const old = [sh.getRange("E1").getValue(), sh.getRange("E2").getValue()];
    const ppd = "PPD" + pad(++max.PPD), vbu = qr ? "VBU" + pad(++max.VBU) : "";
    sh.getRange("E1").setValue(ppd).setNote(APP_NOTE);
    sh.getRange("E2").setValue(vbu).setNote(vbu ? APP_NOTE : "");
    log.push(DAY(d) + " uzávěrka: " + ppd + " (hotově " + cash + " Kč)" + (vbu ? ", " + vbu + " (QR " + qr + " Kč)" : ", bez QR") +
             (old[0] ? " [dříve " + old[0] + "]" : ""));
  }
  SpreadsheetApp.flush();
  Logger.log(log.join("\n"));
  return log;
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
      out = {ok: true, products: readProducts(ss, req.date), numbers: usedNumbers(ss, req.date), day: dayInfo(ss, req.date)};
    } else if (req.action === "sync") {
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      const skipped = [];
      try {
        req.ops.forEach(op => {
          if (op.type === "sale") writeSale(ss, op, skipped);
          else if (op.type === "entry") writeMove(ss, op, skipped);
          else if (op.type === "deleteEntry") deleteMove(ss, op.number, skipped);
          else if (op.type === "closeDay") closeDay(ss, op, skipped);
          else throw new Error("Unknown operation: " + op.type);
        });
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
      out = {ok: true, done: req.ops.length, skipped: skipped, day: dayInfo(ss, req.date)};
    } else if (req.action === "orders") {
      out = {ok: true, orders: readOrders(ss, req.date)};
    } else if (req.action === "done") {
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try { markDone(ss, req.number, req.done, req.time); SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
      out = {ok: true, orders: readOrders(ss, req.date)};
    } else {
      throw new Error("Unknown action: " + req.action);
    }
  } catch (err) {
    out = {ok: false, error: String((err && err.message) || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Products with price and Zbývá from Sklad (null when Naskladněno is empty)
// price = the column of the day sheet for this date (Den 1 when no day matches yet)
function readProducts(ss, iso) {
  let n = 1;
  try { if (iso) n = daySheet(ss, iso, false).n; } catch (e) {}
  const pr = sheet(ss, S.products).getRange(2, 1, PRODUCTS_LAST - 1, KIND_COL).getValues();
  const st = sheet(ss, S.stock).getRange(2, 2, PRODUCTS_LAST - 1, 4 + DAYS).getValues();
  const out = [];
  pr.forEach((r, i) => {
    const kind = String(r[KIND_COL - 1]).trim().toLowerCase();
    if (String(r[0]).trim() === "" || kind === "skrýt") return;
    const left = st[i][st[i].length - 1];
    out.push({name: String(r[0]).trim(), price: toNumber(r[n]), stock: st[i][0] === "" || left === "" ? null : toNumber(left),
              topping: kind === "přísada"});
  });
  return out;
}

// The highest number used per series. A day close that counts itself (formula) and isn't over yet (today or later)
// doesn't count: a PPD/VBU written today takes the next number and that day's close moves after it.
function usedNumbers(ss, iso) {
  const out = {P: 0, VBU: 0, PPD: 0, VPD: 0, ID: 0};
  const take = v => { const m = norm(v).match(/^(PPD|VPD|VBU|ID|P)(\d+)$/); if (m) out[m[1]] = Math.max(out[m[1]], +m[2]); };
  const t = iso ? toDate(iso) : null, todayStart = t ? new Date(t.getFullYear(), t.getMonth(), t.getDate()).getTime() : Infinity;
  for (let d = 1; d <= DAYS; d++) {
    const sh = sheet(ss, DAY(d));
    sh.getRange(FIRST, 1, LAST - FIRST + 1, 1).getValues().forEach(r => take(r[0]));
    const day = sh.getRange("C1").getValue();
    const open = !(day instanceof Date) || new Date(day.getFullYear(), day.getMonth(), day.getDate()).getTime() >= todayStart;
    ["E1", "E2"].forEach(a => { const c = sh.getRange(a); if (!(open && c.getFormula())) take(c.getValue()); });
  }
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
  if (!free) throw new Error("All day sheets already have another date - change the date in C1 of a day sheet.");
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
  if (op.void) {
    existing.forEach(x => {
      x.sh.getRange(x.row, 4).setValue(0);
      x.sh.getRange(x.row, 7).setValue("");
      x.sh.getRange(x.row, 11).setValue("STORNO");
    });
    voidOrder(ss, op.number);
    return;
  }
  if (op.kitchen && op.kitchen.length) writeOrder(ss, op);
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
    sh.getRange(row, 7, 1, 2).setValues([[k === 0 && tip ? tip : "", op.method === "QR" ? "QR" : "Hotově"]]);
    sh.getRange(row, 11).setValue("");
    sh.getRange(row, 1).setNote(APP_NOTE);
    a[i][0] = op.number;
    i++;
  });
}

// ---- Kuchyně: one row per order for the cook's phone
//   Datum | Čas | Číslo | Objednávka (readable) | Hotovo (time) | Stav (STORNO) | data (the same order for the app)
const KITCHEN_HEAD = ["Datum", "Čas", "Číslo", "Objednávka", "Hotovo (čas)", "Stav", "data pro kuchyni (nepřepisovat)"];
function kitchenSheet(ss) {
  let sh = ss.getSheetByName(S.kitchen);
  if (sh) return sh;
  sh = ss.insertSheet(S.kitchen);
  sh.getRange(1, 1, 1, KITCHEN_HEAD.length).setValues([KITCHEN_HEAD]); head(sh.getRange(1, 1, 1, KITCHEN_HEAD.length));
  sh.getRange("A:A").setNumberFormat("d.m.yyyy"); sh.getRange("B:B").setNumberFormat("@"); sh.getRange("E:E").setNumberFormat("@");
  sh.getRange("D:D").setWrap(true);
  sh.setColumnWidth(4, 320); sh.setColumnWidth(7, 200); sh.setFrozenRows(1);
  sh.getRange("I1:I3").setValues([["Objednávky pro kuchaře - zapisuje je pokladna, odškrtává je telefon v kuchyni (kitchen.html)."],
    ["Hotovo = čas, kdy kuchař objednávku odškrtl. Smažete-li čas, objednávka se v kuchyni znovu objeví."],
    ["Peníze se tu nepočítají - ty jsou v listech Den 1 až Den " + DAYS + "."]]);
  sh.setColumnWidth(9, 520);
  return sh;
}

const orderText = lines => lines.map(l => l.qty + "× " + l.name + (l.tops && l.tops.length ? ": " + l.tops.join(", ") : "")).join("\n");

function kitchenRow(sh, number) {
  const last = sh.getLastRow();
  if (last < 2) return null;
  const v = sh.getRange(2, 3, last - 1, 1).getValues();
  const i = v.findIndex(r => norm(r[0]) === norm(number));
  return i === -1 ? null : 2 + i;
}

function writeOrder(ss, op) {
  const sh = kitchenSheet(ss);
  const row = kitchenRow(sh, op.number) || sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 4).setValues([[toDate(op.date), String(op.time || ""), op.number, orderText(op.kitchen)]]);
  sh.getRange(row, 6, 1, 2).setValues([["", JSON.stringify(op.kitchen)]]);  // a re-sent order keeps its Hotovo time
}

function voidOrder(ss, number) {
  const sh = ss.getSheetByName(S.kitchen);
  const row = sh && kitchenRow(sh, number);
  if (row) sh.getRange(row, 6).setValue("STORNO");
}

function markDone(ss, number, done, time) {
  const sh = kitchenSheet(ss);
  const row = kitchenRow(sh, number);
  if (!row) throw new Error(number + " is not in the Kuchyně sheet.");
  sh.getRange(row, 5).setValue(done ? String(time || Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "HH:mm:ss")) : "");
}

// The orders of one day (the app sorts them)
function readOrders(ss, iso) {
  const sh = kitchenSheet(ss);
  const last = sh.getLastRow();
  if (last < 2 || !iso) return [];
  const want = toDate(iso).toDateString();
  const v = sh.getRange(2, 1, last - 1, 7).getValues(), t = sh.getRange(2, 2, last - 1, 1).getDisplayValues();
  const out = [];
  v.forEach((r, i) => {
    if (!(r[0] instanceof Date) || r[0].toDateString() !== want || String(r[2]).trim() === "") return;
    let lines;
    try { lines = JSON.parse(r[6]); } catch (e) { lines = null; }
    if (!Array.isArray(lines)) lines = String(r[3]).split("\n").filter(String).map(s => ({name: s, qty: "", tops: []}));
    out.push({number: String(r[2]).trim(), time: t[i][0], lines: lines, done: String(r[4]).trim(), void: norm(r[5]) === "STORNO"});
  });
  return out;
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

// ---- end of day: PPD number to E1 (amount F1 = SUM Hotově), VBU number to E2 (amount F2 = SUM QR, only when there was QR)
function closeDay(ss, op, skipped) {
  const sh = daySheet(ss, op.date, true).sh;
  const free = (cell, number) => cell.getValue() === "" || cell.getNote() === APP_NOTE || norm(cell.getValue()) === norm(number);
  const ppd = sh.getRange("E1"), vbu = sh.getRange("E2");
  if (ppd.getFormula()) return;  // the sheet numbers the close itself (closeNumbers)
  if (free(ppd, op.number)) ppd.setValue(op.number).setNote(APP_NOTE); else skipped.push(op.number);
  if (!op.vbu) return;
  SpreadsheetApp.flush();
  const qr = toNumber(sh.getRange("F2").getValue());
  if (!free(vbu, op.vbu)) skipped.push(op.vbu);
  else if (qr) vbu.setValue(op.vbu).setNote(APP_NOTE);
  else if (vbu.getNote() === APP_NOTE) vbu.clearContent().setNote("");
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
