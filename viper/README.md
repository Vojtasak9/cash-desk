# Viper (experimentální pokladna)

App: https://vojtasak9.github.io/cash-desk/viper/

Pokladna je jen hezký vstup. Všechno počítá Google Sheet „Viper“ pomocí funkcí.

## Listy
- **Produkty** – název a cena. Odsud si pokladna bere produkty.
- **Den 1 – Den 4** – prodeje. Člověk vyplní jen **Produkt** a **Počet**. Cena (VLOOKUP), Celkem (Počet × Cena) a součty (SUM, SUMIF) se dopočítají samy. Nahoře je hotovost v kase, peníze na účtu a celkem.
- **Sklad** – Prodáno = SUMIF přes Den 1 až Den 4, Zbývá = Naskladněno − Prodáno.
- **Půjčky** – vklady, zálohy, výdaje (PPD/VPD/ID). Tabulka „kdo komu dluží“ se počítá přes SUMIF.
- **Typy pohybů** – tahák: Půjčky si z něj přes VLOOKUP berou, jestli pohyb mění hotovost, účet nebo dluh.

## Čísla dokladů
- Prodeje a příjmy do kasy: **PPD** (společná řada)
- Výdaje z kasy: **VPD**
- Bez pohybu hotovosti: **ID**

## Spropitné
Když něco stojí 25 Kč a zákazník dá 30, napiš do „Přijato“ 30. Uloží se 30, z toho 5 jako spropitné.

## Připojení
Adresa skriptu je předvyplněná. Klíč se zadá jednou v Nastavení.
Nový začátek: v Apps Script editoru spusť `rebuildViper`.
