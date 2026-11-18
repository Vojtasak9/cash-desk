# Viper (experimentální pokladna)

App: https://vojtasak9.github.io/cash-desk/viper/

Pokladna je jen hezký vstup. Všechno počítá Google Sheet „Viper“ pomocí funkcí.

## Listy
- **Produkty** – název a cena pro každý den (Cena Den 1 … Cena Den 4). Pokladna i list Den berou cenu svého dne. Sloupec **Druh**: prázdné = produkt, `přísada` = dává se na produkt (Kečup, Hořčice zdarma; Cibulka, Jalapeños, Okurky +5 Kč), `skrýt` = pokladna ho nenabízí (starý Addon – Den 1 a Den 2 jeho cenu dál používají).
- **Kuchyně** – objednávky pro kuchaře: pokladna sem zapíše každý prodej i s přísadami, telefon v kuchyni je odškrtává (Hotovo = čas). Peníze se tu nepočítají.
- **Den 1 – Den 4** – prodeje. Člověk vyplní jen **Produkt** a **Počet**. Cena (VLOOKUP), Celkem (Počet × Cena) a součty (SUM, SUMIF) se dopočítají samy. Nahoře je hotovost v kase, peníze na účtu a celkem.
- **FULL** – kompletní přehled pro odevzdání, nic se sem nepíše.
- **Sklad** – Prodáno = SUMIF přes Den 1 až Den 4, Zbývá = Naskladněno − Prodáno.
- **Půjčky** – vklady, zálohy, výdaje (PPD/VPD/ID). Tabulka „kdo komu dluží“ se počítá přes SUMIF.
- **Typy pohybů** – tahák: Půjčky si z něj přes VLOOKUP berou, jestli pohyb mění hotovost, účet nebo dluh.
  **Nákup na dluh** = kupní smlouva, zboží máme, ale ještě jsme nezaplatili: číslo ID, peníze 0, kasa prodávajícímu dluží. Až zaplatíme, nový řádek **Vrácení dluhu** (VPD hotově, VBU z účtu).

## Čísla dokladů
- **P001, P002…** – paragon = každý prodej (hotově i QR), jedna řada. Na listu Den sloupec Platba = Hotově / QR.
- **Uzávěrka dne** (tlačítko „Close day (PPD + VBU)“): **PPD** = hotovost ze všech P toho dne do kasy (E1, částka F1),
  **VBU** = QR platby ze všech P toho dne na účet (E2, částka F2; jen když nějaké QR byly)
  Čísla uzávěrky (E1, E2) se **počítají sama** vzorcem: nejvyšší PPD/VBU v Půjčkách s datem do toho dne a v uzávěrkách
  předchozích dnů + 1. Když do Půjček dopíšete např. VBU pro Den 1, uzávěrky se samy posunou. E1/E2 ani FULL nepřepisujte.
- **PPD** – i půjčka do kasy v hotovosti
- **VPD** – výdej z kasy
- **VBU** – i jiný pohyb na účtu (Půjčky, Platba = Účet)
- **ID** – bez pohybu peněz

## FULL
List FULL se vyplňuje sám (SORT + FILTER + VSTACK) a obsahuje všechny doklady podle času.
Každý P řádek nese své peníze (Počet × Cena + spropitné): hotově → Pokladna příjmy, QR → Účet příjmy.
Uzávěrka dne (PPD, VBU) je ve FULL jen jako popisek se součtem, aby se nic nepočítalo dvakrát.
Celkem = předchozí + pokladna příjmy − výdaje + účet příjmy − výdaje.

## Opravy
Když někdo přepíše nebo smaže vzorec (šedé buňky), spusť v Apps Script editoru `repairFormulas` – vrátí všechny vzorce, data nechá.

## Spropitné
Když něco stojí 25 Kč a zákazník dá 30, napiš do „Přijato“ 30. Uloží se 30, z toho 5 jako spropitné.

## Přísady a kuchyně

V pokladně: ťukni na produkt (nový řádek objednávky), pak na přísady – přidají se k vybranému (zeleně označenému) řádku. Do listu dne jde každá přísada jako vlastní řádek (např. Cibulka 1 × 5 Kč), takže tržba i FULL sedí. Na účtence je u produktu „+ cibulka, okurky“.

Telefon kuchaře: v pokladně **Settings → Kitchen phone** naskenuj QR kód a dej „Přidat na plochu“. Stránka `kitchen.html` každých 5 s načte dnešní objednávky, pípne u nové, drží displej rozsvícený a ťuknutím na **Done** objednávku odškrtne (v listu Kuchyně se zapíše čas). Stornovaný prodej se v kuchyni ukáže červeně.

## Připojení
Adresa skriptu je předvyplněná. Klíč se zadá jednou v Nastavení.
Nový začátek: v Apps Script editoru spusť `rebuildViper`.
