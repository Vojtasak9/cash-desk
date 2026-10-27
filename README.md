# Cash Desk

A simple cash register that runs on your computer and keeps **all data in one Excel file**
(`data/cash_desk.xlsx`).

## Features

- **Register** – tap products to build an order, then pay **with QR** or **in cash**
  - QR: shows a payment QR code with the exact order total (Czech *QR Platba* / SPAYD standard –
    every Czech banking app can scan it). Receipt number goes into the message and variable symbol.
  - Cash: type what the customer gave you and it shows the change
- **Receipts** numbered `P 001`, `P 002`, `P 003`, … – a wrong receipt can be voided (items go back to stock)
- **Day total** – number of receipts, cash, QR and total for any day
- **Products & stock** – add / edit / delete products, prices and stock; stock goes down with every sale
  and you can't sell more than you have (leave stock empty for items you don't count)
- **Debts** – write down money you lent to someone and what they paid back; shows who still owes how much
- **Settings** – your bank account (IBAN or Czech number like `123456789/0800`), name, currency

## Excel file

| Sheet | Contents |
|---|---|
| Products | ID, name, price, stock, on sale |
| Payments | receipt no., date, time, items, total, method (QR/Cash), received, change, status |
| PaymentItems | every line of every receipt |
| Debts | every lend / repayment |
| DebtSummary | per person: lent, repaid, still owes |
| DailySummary | per day: receipts, cash, QR, total |
| Settings | account and shop settings |

You can open the file in Excel at any time to look at it. Close it before selling if you are on Windows
(Windows locks open files).

## Online (Vercel) – use it on the iPad

The app is deployed on Vercel. Open the link in Safari on the iPad, log in with the password, then
tap **Share → Add to Home Screen** – it opens full screen like a normal app.

On Vercel the Excel file is kept in **Vercel Blob** storage (private). Use **Receipts → Download Excel file**
to get it. Settings on Vercel (Project → Settings → Environment Variables):

| Variable | Meaning |
|---|---|
| `APP_PASSWORD` | password for the login screen |
| `SECRET_KEY` | random string that signs the login cookie |
| `BLOB_READ_WRITE_TOKEN` | added automatically when the Blob store is connected |

Every push to `main` on GitHub deploys automatically.

## Run it on your computer

Needs Python 3.9+.

- **macOS:** double-click `start.command`
- **Windows:** double-click `start.bat`
- **Manually:**
  ```bash
  python3 -m venv .venv
  .venv/bin/pip install -r requirements.txt
  .venv/bin/python app.py
  ```

The app opens in your browser at http://127.0.0.1:5050. First go to **Settings** and enter your bank account.

## QR payments – why no online generator

Free QR payment options that were looked at:

- [qr-platba.cz](https://qr-platba.cz/pro-vyvojare/specifikace-formatu/) – the official SPAYD spec (the format used here)
- [qrplatba](https://pypi.org/project/qrplatba/) – Python library producing SPAYD QR codes
- [payqr](https://cz-agents.dev/mcp/payqr/) – free SPAYD generator (MCP server)
- PHP packages such as `czeurqrcode` and `qr-payment-string-generator`

The SPAYD format is simple, so this app builds the payment string itself and draws the QR code with the free
[`qrcode`](https://pypi.org/project/qrcode/) library. It works **offline**, has no limits and sends your data nowhere.
