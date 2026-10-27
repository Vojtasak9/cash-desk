# Cash Desk

A cash register made for the iPad. It runs **completely on the device** – no server, no account and no
internet needed while selling – and gives you **everything as one Excel file**.

## Features

- **Register** – tap products to build an order, then pay **with QR** or **in cash**
  - QR: shows a payment QR code with the exact order total (Czech *QR Platba* / SPAYD standard –
    every Czech banking app can scan it). The receipt number goes into the message and variable symbol.
  - Cash: type what the customer gave you and it shows the change
- **Receipts** numbered `P 001`, `P 002`, `P 003`, … – a wrong receipt can be voided (items go back to stock)
- **Day total** – number of receipts, cash, QR and total for any day
- **Products & stock** – add / edit / delete products, prices and stock; stock goes down with every sale
  and you can't sell more than you have (leave stock empty for items you don't count)
- **Debts** – write down money you lent to someone and what they paid back; shows who still owes how much
- **Excel file** – *Save Excel file* exports everything; *Load from Excel file* restores it (e.g. on a new iPad)

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

You can edit products in Excel and load the file back in.

## On the iPad

1. Open the app's web address in **Safari**.
2. Tap **Share → Add to Home Screen**. It now opens full screen like a normal app and works offline.
3. Go to **Settings & Excel** and enter your bank account (IBAN or Czech number like `123456789/0800`).

**Where is the data?** In the app's storage on that iPad (IndexedDB). It's not sent anywhere.
Clearing Safari's website data deletes it, so **save the Excel file at the end of every day** – a yellow
reminder appears when there are sales that aren't in a saved file yet. *Save to Files / Share…* lets you
put it in iCloud Drive, send it by e-mail, AirDrop, etc.

## Hosting

The app is just static files (`index.html`, `app.js`, `vendor/`, …), so any static host works –
GitHub Pages, Vercel, Netlify, Cloudflare Pages. After changing files, bump `VERSION` in `sw.js`
so installed iPads pick up the new version.

To try it on a computer:

```bash
python3 -m http.server 8000
```

then open http://localhost:8000.

## Libraries (included in `vendor/`, so it works offline)

- [SheetJS](https://sheetjs.com/) 0.18.5 – reads and writes the Excel file
- [qrcode-generator](https://github.com/kazuhikoarase/qrcode-generator) 1.4.4 – draws the QR code

The QR payment string follows the [QR Platba spec](https://qr-platba.cz/pro-vyvojare/specifikace-formatu/)
and is built by the app itself – no online QR service is used.
