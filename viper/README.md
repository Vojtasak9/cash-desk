# Viper (experimentální pokladna)

App: https://vojtasak9.github.io/cash-desk/viper/

Pokladna je jen hezký vstup. Všechno počítá Google Sheet „Viper“ pomocí funkcí.

## Listy
- **Produkty** – název a cena pro každý den (Cena Den 1 … Cena Den 4). Pokladna i list Den berou cenu svého dne.
- **Den 1 – Den 4** – prodeje. Člověk vyplní jen **Produkt** a **Počet**. Cena (VLOOKUP), Celkem (Počet × Cena) a součty (SUM, SUMIF) se dopočítají samy. Nahoře je hotovost v kase, peníze na účtu a celkem.
- **FULL** – kompletní přehled pro odevzdání, nic se sem nepíše.
- **Sklad** – Prodáno = SUMIF přes Den 1 až Den 4, Zbývá = Naskladněno − Prodáno.
- **Půjčky** – vklady, zálohy, výdaje (PPD/VPD/ID). Tabulka „kdo komu dluží“ se počítá přes SUMIF.
- **Typy pohybů** – tahák: Půjčky si z něj přes VLOOKUP berou, jestli pohyb mění hotovost, účet nebo dluh.

## Čísla dokladů
- **P001, P002…** – paragon = každý prodej (hotově i QR), jedna řada. Na listu Den sloupec Platba = Hotově / QR.
- **Uzávěrka dne** (tlačítko „Close day (PPD + VBU)“): **PPD** = hotovost ze všech P toho dne do kasy (E1, částka F1),
  **VBU** = QR platby ze všech P toho dne na účet (E2, částka F2; jen když nějaké QR byly)
- **PPD** – i půjčka do kasy v hotovosti
- **VPD** – výdej z kasy
- **VBU** – i jiný pohyb na účtu (Půjčky, Platba = Účet)
- **ID** – bez pohybu peněz

## FULL
List FULL se vyplňuje sám (SORT + FILTER + VSTACK) a obsahuje všechny doklady podle času.
P řádky jsou jen pro přehled, peníze se započtou v uzávěrce PPD (hotovost) a VBU (QR).
Celkem = předchozí + pokladna příjmy − výdaje + účet příjmy − výdaje.

## Opravy
Když někdo přepíše nebo smaže vzorec (šedé buňky), spusť v Apps Script editoru `repairFormulas` – vrátí všechny vzorce, data nechá.

## Spropitné
Když něco stojí 25 Kč a zákazník dá 30, napiš do „Přijato“ 30. Uloží se 30, z toho 5 jako spropitné.

## Připojení
Adresa skriptu je předvyplněná. Klíč se zadá jednou v Nastavení.
Nový začátek: v Apps Script editoru spusť `rebuildViper`.
