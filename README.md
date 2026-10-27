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

Laid out like the *Jarmark* workbook (Czech sheet names, Excel "Heading 1" headers, amounts in Kč, receipts `P001`, `P002`, …):

| Sheet | Contents |
|---|---|
| Prodej | pieces sold per product per sale day (like *Prodej Propisek*), with formulas for *Celkem (ks)*, *Tržba za daný druh zboží* and *Tržba celkem* |
| Evidence Prodeje | every receipt line: Číslo, Produkt, Počet, price, total, date |
| Evidence Pokladny | Produkty: name, price, stock, on sale |
| Účtenky | every receipt: number, date, time, items, amount, Hotově / QR, received, change, storno |
| Denní uzávěrka | per day: receipts, Hotově, QR, Celkem (with a total row) |
| Dluhy | every lend (Půjčeno) / repayment (Vráceno) |
| Kdo dluží | per person: Půjčeno, Vráceno, Ještě dluží |
| Nastavení | shop name and bank account (so the file can be loaded back) |

You can edit products in *Evidence Pokladny* and load the file back in.

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

- [ExcelJS](https://github.com/exceljs/exceljs) 4.4.0 – reads and writes the Excel file with formatting
- [qrcode-generator](https://github.com/kazuhikoarase/qrcode-generator) 1.4.4 – draws the QR code

The QR payment string follows the [QR Platba spec](https://qr-platba.cz/pro-vyvojare/specifikace-formatu/)
and is built by the app itself – no online QR service is used.
