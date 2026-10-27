"""All data of the cash desk lives in one Excel workbook (data/cash_desk.xlsx).

Sheets:
  Products      - editable product list with price and stock
  Payments      - one row per receipt (P 001, P 002, ...)
  PaymentItems  - the individual lines of each receipt
  Debts         - money lent to people and their repayments
  DebtSummary   - who owes how much (regenerated on every save)
  DailySummary  - cash / QR / total per day (regenerated on every save)
  Settings      - IBAN, account name, currency, ...
"""
import io
import os
import threading
from datetime import date, datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

import blob_storage

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
XLSX_PATH = os.environ.get("CASH_DESK_FILE", os.path.join(DATA_DIR, "cash_desk.xlsx"))

SHEETS = {
    "Products": ["ID", "Name", "Price", "Stock", "Active"],
    "Payments": ["Receipt", "Date", "Time", "Items", "Total", "Method", "Received", "Change", "Status"],
    "PaymentItems": ["Receipt", "Product ID", "Name", "Qty", "Unit price", "Line total"],
    "Debts": ["ID", "Date", "Person", "Type", "Amount", "Note"],
    "DebtSummary": ["Person", "Lent", "Repaid", "Still owes"],
    "DailySummary": ["Date", "Receipts", "Cash", "QR", "Total"],
    "Settings": ["Key", "Value"],
}

DEFAULT_SETTINGS = {
    "account_name": "",
    "iban": "",
    "currency": "CZK",
    "shop_name": "Cash desk",
}

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="2F5597")

_lock = threading.Lock()


class StoreError(Exception):
    pass


def _style_header(ws):
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    ws.freeze_panes = "A2"
    for col in ws.columns:
        letter = col[0].column_letter
        ws.column_dimensions[letter].width = max(12, len(str(col[0].value)) + 4)
    if ws.title == "Payments":
        ws.column_dimensions["D"].width = 50


def _new_workbook():
    wb = Workbook()
    wb.remove(wb.active)
    for name, headers in SHEETS.items():
        ws = wb.create_sheet(name)
        ws.append(headers)
        _style_header(ws)
    ws = wb["Settings"]
    for k, v in DEFAULT_SETTINGS.items():
        ws.append([k, v])
    ws = wb["Products"]
    for row in [[1, "Coffee", 45, 100, "yes"], [2, "Tea", 35, 100, "yes"], [3, "Croissant", 40, 20, "yes"]]:
        ws.append(row)
    return wb


def _load():
    if blob_storage.enabled():
        data = blob_storage.load_bytes()
        if data is None:
            wb = _new_workbook()
            _save(wb)
            return wb
        wb = load_workbook(io.BytesIO(data))
    else:
        if not os.path.exists(XLSX_PATH):
            os.makedirs(os.path.dirname(XLSX_PATH), exist_ok=True)
            _save(_new_workbook())
        wb = load_workbook(XLSX_PATH)
    # add sheets that may be missing in an older file
    for name, headers in SHEETS.items():
        if name not in wb.sheetnames:
            ws = wb.create_sheet(name)
            ws.append(headers)
            _style_header(ws)
    return wb


def _save(wb):
    if blob_storage.enabled():
        buf = io.BytesIO()
        wb.save(buf)
        try:
            blob_storage.save_bytes(buf.getvalue())
        except OSError as e:
            raise StoreError("Could not save to cloud storage (%s). Try again." % e)
        return
    try:
        wb.save(XLSX_PATH)
    except PermissionError:
        raise StoreError("Cannot write the Excel file - close it in Excel and try again.")


def workbook_bytes():
    """The whole Excel file, for the Download button."""
    with _lock:
        wb = _load()
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _rows(ws):
    headers = [c.value for c in ws[1]]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue
        r = dict(zip(headers, row))
        # Excel turns dates typed by hand into real dates - keep them as YYYY-MM-DD text
        if isinstance(r.get("Date"), (datetime, date)):
            r["Date"] = r["Date"].strftime("%Y-%m-%d")
        out.append(r)
    return out


def _num(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _rewrite(ws, rows):
    headers = [c.value for c in ws[1]]
    ws.delete_rows(2, ws.max_row)
    for r in rows:
        ws.append([r.get(h) for h in headers])


# ---------- settings ----------

def get_settings():
    with _lock:
        wb = _load()
    s = dict(DEFAULT_SETTINGS)
    for r in _rows(wb["Settings"]):
        if r["Key"]:
            s[r["Key"]] = "" if r["Value"] is None else str(r["Value"])
    return s


def save_settings(new):
    with _lock:
        wb = _load()
        s = dict(DEFAULT_SETTINGS)
        for r in _rows(wb["Settings"]):
            if r["Key"]:
                s[r["Key"]] = r["Value"]
        s.update({k: v for k, v in new.items() if k in DEFAULT_SETTINGS})
        _rewrite(wb["Settings"], [{"Key": k, "Value": v} for k, v in s.items()])
        _save(wb)


# ---------- products ----------

def list_products():
    with _lock:
        wb = _load()
    out = []
    for r in _rows(wb["Products"]):
        out.append({
            "id": int(r["ID"]),
            "name": r["Name"] or "",
            "price": _num(r["Price"]),
            "stock": None if r["Stock"] in (None, "") else int(_num(r["Stock"])),
            "active": str(r["Active"]).lower() not in ("no", "false", "0"),
        })
    return out


def save_product(p):
    """Create (no id) or update (with id) a product."""
    name = (p.get("name") or "").strip()
    if not name:
        raise StoreError("Product name is required.")
    price = _num(p.get("price"), None)
    if price is None or price < 0:
        raise StoreError("Price must be a non-negative number.")
    stock = p.get("stock")
    stock = None if stock in (None, "") else int(_num(stock))
    with _lock:
        wb = _load()
        ws = wb["Products"]
        rows = _rows(ws)
        pid = p.get("id")
        row = {"Name": name, "Price": price, "Stock": stock,
               "Active": "yes" if p.get("active", True) else "no"}
        if pid:
            for r in rows:
                if int(r["ID"]) == int(pid):
                    r.update(row)
                    break
            else:
                raise StoreError("Product not found.")
        else:
            row["ID"] = max([int(r["ID"]) for r in rows] or [0]) + 1
            rows.append(row)
        _rewrite(ws, rows)
        _save(wb)


def delete_product(pid):
    with _lock:
        wb = _load()
        ws = wb["Products"]
        _rewrite(ws, [r for r in _rows(ws) if int(r["ID"]) != int(pid)])
        _save(wb)


# ---------- payments ----------

def format_receipt(n):
    return "P %03d" % n


def _receipt_number(code):
    try:
        return int(str(code).replace("P", "").strip())
    except ValueError:
        return 0


def next_receipt():
    with _lock:
        wb = _load()
    return _next_receipt(wb)


def _next_receipt(wb):
    nums = [_receipt_number(r["Receipt"]) for r in _rows(wb["Payments"])]
    return format_receipt(max(nums or [0]) + 1)


def create_payment(items, method, received=None):
    """items: [{id, qty}]. Checks and lowers stock, writes the receipt."""
    if method not in ("Cash", "QR"):
        raise StoreError("Payment method must be Cash or QR.")
    if not items:
        raise StoreError("The order is empty.")
    with _lock:
        wb = _load()
        products = {int(r["ID"]): r for r in _rows(wb["Products"])}
        lines = []
        total = 0.0
        for it in items:
            qty = int(_num(it.get("qty")))
            if qty <= 0:
                continue
            p = products.get(int(it["id"]))
            if p is None:
                raise StoreError("Unknown product %s." % it["id"])
            stock = p["Stock"]
            if stock not in (None, ""):
                if _num(stock) < qty:
                    raise StoreError("Not enough '%s' in stock (have %d)." % (p["Name"], int(_num(stock))))
                p["Stock"] = int(_num(stock)) - qty
            line_total = round(_num(p["Price"]) * qty, 2)
            total += line_total
            lines.append((p, qty, line_total))
        if not lines:
            raise StoreError("The order is empty.")
        total = round(total, 2)
        change = None
        if method == "Cash" and received not in (None, ""):
            received = _num(received)
            if received < total:
                raise StoreError("Received cash is less than the total.")
            change = round(received - total, 2)
        else:
            received = None

        receipt = _next_receipt(wb)
        now = datetime.now()
        wb["Payments"].append([
            receipt, now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S"),
            ", ".join("%dx %s" % (q, p["Name"]) for p, q, _ in lines),
            total, method, received, change, "OK",
        ])
        for p, q, lt in lines:
            wb["PaymentItems"].append([receipt, int(p["ID"]), p["Name"], q, _num(p["Price"]), lt])
        _rewrite(wb["Products"], list(products.values()))
        _rebuild_daily(wb)
        _save(wb)
    return {"receipt": receipt, "total": total, "change": change}


def void_payment(receipt):
    """Cancel a receipt and put its items back into stock."""
    with _lock:
        wb = _load()
        ws = wb["Payments"]
        rows = _rows(ws)
        target = next((r for r in rows if r["Receipt"] == receipt), None)
        if target is None:
            raise StoreError("Receipt not found.")
        if target["Status"] == "VOID":
            raise StoreError("Receipt is already voided.")
        target["Status"] = "VOID"
        _rewrite(ws, rows)
        products = {int(r["ID"]): r for r in _rows(wb["Products"])}
        for it in _rows(wb["PaymentItems"]):
            if it["Receipt"] == receipt:
                p = products.get(int(it["Product ID"]))
                if p is not None and p["Stock"] not in (None, ""):
                    p["Stock"] = int(_num(p["Stock"])) + int(_num(it["Qty"]))
        _rewrite(wb["Products"], list(products.values()))
        _rebuild_daily(wb)
        _save(wb)


def list_payments(date=None):
    with _lock:
        wb = _load()
    rows = _rows(wb["Payments"])
    if date:
        rows = [r for r in rows if str(r["Date"]) == date]
    return [{k.lower(): v for k, v in r.items()} for r in reversed(rows)]


def _rebuild_daily(wb):
    days = {}
    for r in _rows(wb["Payments"]):
        if r["Status"] == "VOID":
            continue
        d = days.setdefault(str(r["Date"]), {"Receipts": 0, "Cash": 0.0, "QR": 0.0})
        d["Receipts"] += 1
        d[r["Method"]] = d.get(r["Method"], 0.0) + _num(r["Total"])
    out = []
    for date in sorted(days):
        d = days[date]
        out.append({"Date": date, "Receipts": d["Receipts"], "Cash": round(d["Cash"], 2),
                    "QR": round(d["QR"], 2), "Total": round(d["Cash"] + d["QR"], 2)})
    _rewrite(wb["DailySummary"], out)


def daily_summary(date):
    with _lock:
        wb = _load()
    for r in _rows(wb["DailySummary"]):
        if str(r["Date"]) == date:
            return {"date": date, "receipts": r["Receipts"], "cash": _num(r["Cash"]),
                    "qr": _num(r["QR"]), "total": _num(r["Total"])}
    return {"date": date, "receipts": 0, "cash": 0.0, "qr": 0.0, "total": 0.0}


# ---------- debts ----------

def add_debt(person, kind, amount, note=""):
    person = (person or "").strip()
    if not person:
        raise StoreError("Person is required.")
    if kind not in ("Lent", "Repaid"):
        raise StoreError("Type must be Lent or Repaid.")
    amount = _num(amount, None)
    if amount is None or amount <= 0:
        raise StoreError("Amount must be a positive number.")
    with _lock:
        wb = _load()
        ws = wb["Debts"]
        rows = _rows(ws)
        new_id = max([int(r["ID"]) for r in rows] or [0]) + 1
        ws.append([new_id, datetime.now().strftime("%Y-%m-%d"), person, kind, amount, note or ""])
        _rebuild_debts(wb)
        _save(wb)


def delete_debt(did):
    with _lock:
        wb = _load()
        ws = wb["Debts"]
        _rewrite(ws, [r for r in _rows(ws) if int(r["ID"]) != int(did)])
        _rebuild_debts(wb)
        _save(wb)


def _debt_totals(wb):
    people = {}
    for r in _rows(wb["Debts"]):
        p = people.setdefault(r["Person"], {"Lent": 0.0, "Repaid": 0.0})
        p[r["Type"]] = p.get(r["Type"], 0.0) + _num(r["Amount"])
    return [{"Person": name, "Lent": round(v["Lent"], 2), "Repaid": round(v["Repaid"], 2),
             "Still owes": round(v["Lent"] - v["Repaid"], 2)}
            for name, v in sorted(people.items(), key=lambda kv: kv[0].lower())]


def _rebuild_debts(wb):
    _rewrite(wb["DebtSummary"], _debt_totals(wb))


def list_debts():
    with _lock:
        wb = _load()
    entries = [{k.lower(): v for k, v in r.items()} for r in reversed(_rows(wb["Debts"]))]
    summary = [{"person": r["Person"], "lent": r["Lent"], "repaid": r["Repaid"],
                "owes": r["Still owes"]} for r in _debt_totals(wb)]
    return {"entries": entries, "summary": summary}
