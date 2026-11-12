# Cash Desk

A cash register made for the iPad. It runs **completely on the device** – no server, no account and no
internet needed while selling – and gives you **everything as one Excel file**.

## Features

- **Register** – tap products to build an order, then pay **with QR** or **in cash**
  - QR: shows a payment QR code with the exact order total (Czech *QR Platba* / SPAYD standard –
    every Czech banking app can scan it). The receipt number goes into the message and variable symbol.
  - Cash: type what the customer gave you and it shows the change
- **Receipts** numbered `P001`, `P002`, `P003`, … – a wrong receipt can be voided (items go back to stock)
- **Day total** – number of receipts, cash, QR and total for any day
- **Products & stock** – add / edit / delete products, prices and stock; stock goes down with every sale
  and you can't sell more than you have (leave stock empty for items you don't count)
- **Cash book** – money put into the cash box (Zapůjčka), loans paid back, expenses (Vydáno), money lent from the
  box; shows who owes whom
- **Current cash** – always shown at the top: starting cash + cash sales + money in − money out (QR goes to the bank, so it isn't counted)
- **Google Sheets** – every receipt and cash book entry is sent to the team's sheets, which do the calculating (see below)
- **Excel file** – *Save Excel file* exports everything; *Load from Excel file* restores it (e.g. on a new iPad)

## Excel file

Laid out like the *Jarmark* workbook (Czech sheet names, Excel "Heading 1" headers, amounts in Kč, receipts `P001`, `P002`, …):

| Sheet | Contents |
|---|---|
| Prodej | pieces sold per product per sale day (like *Prodej Propisek*), with formulas for *Celkem (ks)*, *Tržba za daný druh zboží* and *Tržba celkem* |
| Evidence Prodeje | one row per product sold: Číslo, Produkt, Počet, price, total, date, time, Hotově / QR, received, change, storno |
| Evidence Pokladny | Produkty: name, price, stock, on sale |
| Denní uzávěrka | per day: receipts, Hotově, QR, Celkem (with a total row) |
| Pokladní kniha | every cash book entry: date, PPD/VPD number, type, person, description, in, out |
| Kdo dluží | per person: what the kasa owes them, what they owe the kasa |
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

## Google Sheets

The app types the plain data into the team's Google Sheets; **the sheets do all the calculating with your own
formulas** (SUM, SUMIF, …). The app never writes into formula columns.

| What | Where it goes | Columns the app fills |
|---|---|---|
| every receipt | *Ten Párek Evidence pokladny* → **Prodej** | DATUM, ČÍSLO (P001…), TYP DOKUMENTU (Paragon), TYP ZBOŽÍ (e.g. `2x Hotdog, 1x Mini`), amount under **POKLADNA** (cash) or **BĚŽNÝ ÚČET** (QR) |
| a cancelled receipt | the same row | TYP ZBOŽÍ starts with `STORNO`, amounts emptied |
| money in / out of the cash box | *Ten Párek Evidence pokladny* → **Kompletní evidence** | DATUM, ČÍSLO (PPD001… in, VPD001… out), DOKUMENTY (Zapůjčka, Vrácení zapůjčky, Vydáno, Půjčka z kasy, Splátka), POPIS, POKLADNA Příjmy or Výdaje |
| products and prices | read from *Ten Párek Evidence pokladny* → **Produkty** (column A Produkt, column B Cena) | – |

A receipt goes into the row that already has its number (P001, P002, …); after P020 new rows are added below.
Cash book entries go into the first empty row above the *Celkem* totals; when none is left, a row is inserted
inside the table so your SUM ranges grow with it. Sending the same thing twice never creates a duplicate.

Without internet the app keeps working and sends everything when the connection is back. The header shows
**Sheets ✓**, **Sheets: 2 waiting** or a warning.

### Setup (once, on a computer)

The person doing this needs edit access to *Ten Párek Evidence pokladny*.

1. Open **Ten Párek Evidence pokladny** → **Rozšíření → Apps Script** (*Extensions → Apps Script*).
2. Delete the code that is there and paste in everything from
   [`google-apps-script.js`](https://vojtasak9.github.io/cash-desk/google-apps-script.js).
3. At the top, change `const KEY = "ZMEN-ME";` to your own password, e.g. `const KEY = "parek-2026";`. Save (⌘S / Ctrl+S).
4. **Nasadit → Nové nasazení** (*Deploy → New deployment*) → gear icon → **Webová aplikace** (*Web app*).
   - *Spustit jako / Execute as*: **Já (Me)**
   - *Kdo má přístup / Who has access*: **Kdokoli (Anyone)**
5. Click **Nasadit / Deploy** and allow access. Google says the app isn't verified – it's your own script:
   **Rozšířené → Přejít na … (nebezpečné)** (*Advanced → Go to … (unsafe)*) → **Povolit / Allow**.
6. Copy the **Web app URL** (it ends with `/exec`).
7. On the iPad: **Settings & Excel → Google Sheets** → paste the URL, type the same key → **Save & connect**.
   The products load from the *Produkty* sheet. It asks whether to also send receipts that are already in the app.
8. In **Settings**, set **Starting cash** to the money that is already in the box and already written in
   *Kompletní evidence* (now 500 + 235 − 35 = 700 Kč), so the app's cash matches the sheets.

Changed the script later? **Nasadit → Spravovat nasazení → ✏️ → Verze: Nová verze** – the URL stays the same.

Good to know:
- Numbers continue after what is already in the sheets, including receipts and PPD/VPD numbers typed by hand.
- The app only overwrites or deletes rows it wrote itself (they have the note *Zapsáno z pokladny* on the number).
  Rows typed by hand are never changed; the app shows a warning if a number is already taken.
- Dropdown columns (TYP ZBOŽÍ, DOKUMENTY) get one of their own choices: e.g. `1x Hotdog` → *Hot dog*, with all
  items in the cell's note. A type the DOKUMENTY dropdown doesn't have is written as *Zapůjčka* (money in) or *Vydáno* (money out).
  Give products in the *Produkty* sheet names that contain a TYP ZBOŽÍ choice (e.g. *Párek v rohlíku Standard*) to have it filled in.
- Don't rename the sheets or move their columns – the script checks the headers and shows an error in the app
  instead of writing into the wrong place.
- The example row P001 (*Párel v rohlíku*) is overwritten by the first real sale.
- Write sums over open ranges like `=SUM(E3:E)` so rows added after P020 are counted too.
- Product names and prices are changed in the *Produkty* sheet of *Ten Párek Evidence pokladny*; the app reloads them whenever it is opened or comes back to the screen.

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
