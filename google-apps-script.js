/**
 * Cash Desk -> Google Sheets
 *
 * The cash desk app sends its data here and this script types it into the sheets:
 *   - every receipt (P001, P002, ...)  -> "Ten Párek Evidence pokladny" / Prodej
 *   - loans and expenses (PPD / VPD)   -> "Ten Párek Evidence pokladny" / Kompletní evidence
 *   - the app loads products + prices  <- "Kasa" / Produkty
 *
 * It only writes plain values (date, number, goods, amount). Totals, the cash balance,
 * SUM / SUMIF etc. are your own formulas in the sheets - the script never touches them.
 *
 * Setup: see "Google Sheets" in the README, or Settings & Excel in the app.
 */

// 1) Pick your own password, put it here AND type the same one into the app.
const KEY = "ZMEN-ME";

const EVIDENCE_ID = "1_tPsNIFdS_2jo-ZLwsGqWwZQi5xRSp0w2RpxElQPSLw";  // Ten Párek Evidence pokladny
const KASA_ID = "1oowL1M-IIxIWIIewRS1elvtV_3ErwABPJlqN6wu1ASI";      // Kasa

const SALES_SHEET = "Prodej";
const BOOK_SHEET = "Kompletní evidence";
const PRODUCTS_SHEET = "Produkty";
const FIRST_ROW = 3;  // rows 1 and 2 are headers

function doGet() {
  return ContentService.createTextOutput("Cash Desk script is running.");
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (KEY === "ZMEN-ME") throw new Error('Change KEY at the top of the script first (it still says "ZMEN-ME").');
    if (req.key !== KEY) throw new Error("Wrong key - type the same key into the app as in the script.");
    if (req.action === "products") {
      out = {ok: true, products: readProducts()};
    } else if (req.action === "sync") {
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try { req.ops.forEach(applyOp); } finally { lock.releaseLock(); }
      out = {ok: true, done: req.ops.length};
    } else {
      throw new Error("Unknown action: " + req.action);
    }
  } catch (err) {
    out = {ok: false, error: String((err && err.message) || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Every operation can safely be sent twice: rows are found by their number (P001, PPD001, ...)
// and overwritten, so a retry after a lost connection never creates a duplicate.
function applyOp(op) {
  if (op.type === "sale") writeSale(op);
  else if (op.type === "entry") writeEntry(op);
  else if (op.type === "deleteEntry") clearEntry(op.number);
  else throw new Error("Unknown operation: " + op.type);
}

// ---- Kasa / Produkty -> app

function readProducts() {
  const sh = getSheet(KASA_ID, PRODUCTS_SHEET);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 2).getValues()
    .filter(r => String(r[0]).trim() !== "")
    .map(r => ({name: String(r[0]).trim(), price: toNumber(r[1])}));
}

// ---- receipts -> Prodej
// DATUM | ČÍSLO | TYP DOKUMENTU | TYP ZBOŽÍ | POKLADNA příjmy | BĚŽNÝ ÚČET příjmy | (your formulas)

function writeSale(op) {
  const sh = getSheet(EVIDENCE_ID, SALES_SHEET);
  expectHeader(sh, 2, "ČÍSLO");
  expectHeader(sh, 4, "ZBOŽÍ");
  expectHeader(sh, 5, "POKLADNA");
  expectHeader(sh, 6, "ÚČET");
  const last = Math.max(sh.getLastRow(), FIRST_ROW - 1);
  let row = findRow(sh, 2, op.receipt, last + 1) || findRow(sh, 2, "", last + 1);
  if (!row) {
    row = last + 1;
    if (last >= FIRST_ROW) sh.getRange(last, 1, 1, sh.getLastColumn()).copyTo(sh.getRange(row, 1, 1, sh.getLastColumn()));
  }
  const cash = !op.void && op.method === "Cash" ? op.total : "";
  const bank = !op.void && op.method === "QR" ? op.total : "";
  sh.getRange(row, 1, 1, 6).setValues([[toDate(op.date), op.receipt, "Paragon",
    op.void ? "STORNO " + op.goods : op.goods, cash, bank]]);
  sh.getRange(row, 1).setNumberFormat("d.m.yyyy");
}

// ---- loans and expenses -> Kompletní evidence
// DATUM | ČÍSLO | DOKUMENTY | POPIS | POKLADNA příjmy | POKLADNA výdaje | ÚČET příjmy | ÚČET výdaje | (your formulas)

function writeEntry(op) {
  const sh = bookSheet();
  const end = footerRow(sh);
  let row = findRow(sh, 2, op.number, end) || freeBookRow(sh, end, op.kind);
  if (!row) {
    // No empty row left: insert one inside the table (above its last row) so your SUM ranges
    // grow with it, move the last row up into it and write the new entry at the bottom.
    const cols = sh.getLastColumn();
    sh.insertRowBefore(end - 1);
    sh.getRange(end, 1, 1, cols).copyTo(sh.getRange(end - 1, 1, 1, cols));
    row = end;
    sh.getRange(row, 1, 1, 8).clearContent();
  }
  sh.getRange(row, 1, 1, 6).setValues([[toDate(op.date), op.number, op.kind, op.description,
    op.cashIn || "", op.cashOut || ""]]);
  sh.getRange(row, 1).setNumberFormat("d.m.yyyy");
}

function clearEntry(number) {
  const sh = bookSheet();
  const row = findRow(sh, 2, number, footerRow(sh));
  if (row) sh.getRange(row, 1, 1, 8).clearContent();
}

function bookSheet() {
  const sh = getSheet(EVIDENCE_ID, BOOK_SHEET);
  expectHeader(sh, 2, "ČÍSLO");
  expectHeader(sh, 3, "DOKUMENT");
  expectHeader(sh, 4, "POPIS");
  expectHeader(sh, 5, "POKLADNA");
  return sh;
}

// The row with "Celkem" in column I starts the totals at the bottom; entries go above it.
function footerRow(sh) {
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return FIRST_ROW;
  const col = sh.getRange(FIRST_ROW, 9, last - FIRST_ROW + 1, 1).getValues();
  for (let i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim().toLowerCase() === "celkem") return FIRST_ROW + i;
  }
  return last + 1;
}

// A free row has nothing in A..H, except maybe the same document type typed in advance.
function freeBookRow(sh, end, kind) {
  if (end <= FIRST_ROW) return null;
  const rows = sh.getRange(FIRST_ROW, 1, end - FIRST_ROW, 8).getValues();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i].map(v => String(v).trim());
    const rest = [r[0], r[1], r[3], r[4], r[5], r[6], r[7]];
    if (rest.every(v => v === "") && (r[2] === "" || r[2].toLowerCase() === kind.toLowerCase())) return FIRST_ROW + i;
  }
  return null;
}

// ---- helpers

function getSheet(fileId, name) {
  const ss = SpreadsheetApp.openById(fileId);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found in "' + ss.getName() + '" - was it renamed?');
  return sh;
}

function expectHeader(sh, col, text) {
  const value = String(sh.getRange(1, col).getValue()).toUpperCase();
  if (value.indexOf(text) === -1) {
    throw new Error('Sheet "' + sh.getName() + '": column ' + String.fromCharCode(64 + col) +
      ' should be "' + text + '" but is "' + value + '". Were columns moved?');
  }
}

// First row from FIRST_ROW up to (not including) `end` whose column `col` equals `value`
function findRow(sh, col, value, end) {
  if (end <= FIRST_ROW) return null;
  const want = String(value).trim().toUpperCase();
  const vals = sh.getRange(FIRST_ROW, col, end - FIRST_ROW, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === want) return FIRST_ROW + i;
  }
  return null;
}

function toDate(iso) {
  const p = String(iso).split("-").map(Number);
  return new Date(p[0], p[1] - 1, p[2], 12);  // noon, so no time zone can move it to another day
}

function toNumber(v) {
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/[^\d,.-]/g, "").replace(",", "."));
  return isNaN(n) ? 0 : n;
}
