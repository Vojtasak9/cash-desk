import hashlib
import hmac
import io
import os
import re
import webbrowser
from datetime import date, timedelta
from threading import Timer

from flask import Flask, jsonify, redirect, render_template, request, send_file, session

import excel_store as store
import qr_payment

app = Flask(__name__)

# Optional password (set APP_PASSWORD on Vercel). Without it the app is open - fine on your own computer.
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
app.secret_key = os.environ.get("SECRET_KEY") or hashlib.sha256(("cash-desk:" + APP_PASSWORD).encode()).hexdigest()
app.permanent_session_lifetime = timedelta(days=365)
app.config.update(SESSION_COOKIE_SAMESITE="Lax", SESSION_COOKIE_SECURE=bool(os.environ.get("VERCEL")))


@app.before_request
def require_login():
    if not APP_PASSWORD or session.get("ok") or request.endpoint in ("login", "static"):
        return None
    if request.path.startswith("/api/"):
        return jsonify(error="Please log in again."), 401
    return redirect("/login")


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if hmac.compare_digest(request.form.get("password", ""), APP_PASSWORD):
            session.permanent = True
            session["ok"] = True
            return redirect("/")
        error = "Wrong password"
    return render_template("login.html", error=error)


@app.get("/logout")
def logout():
    session.clear()
    return redirect("/login")


@app.errorhandler(store.StoreError)
@app.errorhandler(ValueError)
def handle_error(e):
    return jsonify(error=str(e)), 400


@app.errorhandler(500)
def handle_crash(e):
    return jsonify(error="Something went wrong on the server. Try again."), 500


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/state")
def state():
    return jsonify(products=store.list_products(), settings=store.get_settings(),
                   next_receipt=store.next_receipt(), auth=bool(APP_PASSWORD))


@app.post("/api/products")
def save_product():
    store.save_product(request.get_json())
    return jsonify(products=store.list_products())


@app.delete("/api/products/<int:pid>")
def delete_product(pid):
    store.delete_product(pid)
    return jsonify(products=store.list_products())


@app.post("/api/qr")
def qr():
    """QR code for the open order - the amount always matches the order total."""
    body = request.get_json()
    s = store.get_settings()
    receipt = body.get("receipt") or store.next_receipt()
    vs = re.sub(r"\D", "", receipt)
    data = qr_payment.spayd(s["iban"], float(body["amount"]), s["currency"],
                            message="%s %s" % (s["shop_name"], receipt), variable_symbol=vs,
                            recipient=s["account_name"])
    return jsonify(svg=qr_payment.qr_svg(data), spayd=data)


@app.post("/api/payments")
def create_payment():
    body = request.get_json()
    result = store.create_payment(body.get("items", []), body.get("method"), body.get("received"))
    return jsonify(result=result, products=store.list_products(), next_receipt=store.next_receipt())


@app.post("/api/payments/<path:receipt>/void")
def void_payment(receipt):
    store.void_payment(receipt)
    return jsonify(ok=True, products=store.list_products())


@app.get("/api/payments")
def payments():
    d = request.args.get("date") or date.today().isoformat()
    return jsonify(payments=store.list_payments(d), summary=store.daily_summary(d))


@app.get("/api/debts")
def debts():
    return jsonify(store.list_debts())


@app.post("/api/debts")
def add_debt():
    b = request.get_json()
    store.add_debt(b.get("person"), b.get("type"), b.get("amount"), b.get("note"))
    return jsonify(store.list_debts())


@app.delete("/api/debts/<int:did>")
def delete_debt(did):
    store.delete_debt(did)
    return jsonify(store.list_debts())


@app.post("/api/settings")
def save_settings():
    b = request.get_json()
    if b.get("iban"):
        b["iban"] = qr_payment.to_iban(b["iban"])  # validate + normalise to IBAN
    store.save_settings(b)
    return jsonify(settings=store.get_settings())


@app.get("/download")
def download():
    return send_file(io.BytesIO(store.workbook_bytes()), as_attachment=True,
                     download_name="cash_desk_%s.xlsx" % date.today().isoformat(),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    if not os.environ.get("NO_BROWSER"):
        Timer(1.0, lambda: webbrowser.open("http://127.0.0.1:%d" % port)).start()
    print("Cash desk running at http://127.0.0.1:%d  (data: %s)" % (port, store.XLSX_PATH))
    app.run(host="127.0.0.1", port=port)
