# Sjednocené aukce EU ETS (EEX, primární trh) 2012–2025

Skript `build_wide.py` sloučí 14 ročních souborů (2012–2019 `.xls`, 2020–2025 `.xlsx`) do jednoho wide souboru.
Zdroj se jen čte. Skript je deterministický a idempotentní (výstupy se přepisují).

Spuštění: `python build_wide.py [DATA_PATH] [OUT_PATH]` (cesty jsou i konstanty na začátku skriptu). Vyžaduje `pandas`, `openpyxl`, `xlrd`.

## ⚠️ Nejdůležitější zjištění: hodnoty po státech NEJSOU tCO2

Ve všech letech (starý formát `Auction Details` i sloupce států od 2017) je **součet přes státy roven `Total Revenue` v EUR**
(maximální odchylka 0 u všech 2 545 aukcí s rozpisem). Poměr součtu k `Auction Volume` je 2,6–97,5, tedy nejde o tCO2.
Hodnoty jsou **výnosy z aukce (EUR) připadající jednotlivým státům**.

Postup:
- `revenue_eur_<ISO>` = surová hodnota ze zdroje (soubor `eu_ets_auctions_wide_revenue_eur.csv`, list `wide_revenue_eur`).
- `volume_tco2_<ISO>` = **dopočet** `revenue_eur / auction_price_eur`. Součet přes státy pak přesně sedí na `Auction Volume`
  (odchylka 0 %), dopočtené hodnoty jsou celá čísla (odchylka < 1e-9) a zaokrouhlují se na celá tCO2.
- Interpretace zůstává „objem povolenek nabídnutých na účet státu“, ne „nakoupeno státem“.

## Soubory

| Soubor | Obsah |
|---|---|
| `eu_ets_auctions_wide.xlsx` | list `wide` (1 řádek = aukce), `daily` (1 řádek = den), `wide_revenue_eur` |
| `eu_ets_auctions_wide.csv` | totéž co list `wide` (UTF-8) |
| `eu_ets_auctions_wide_daily.csv` | agregace jen podle dne |
| `eu_ets_auctions_wide_revenue_eur.csv` | surové výnosy v EUR po státech |
| `eu_ets_auctions_long.csv` | long tabulka (audit): `date, auction_type, contract, certificate, country_code, revenue_eur, volume_tco2, volume_method, source_file, source_row` |
| `eu_ets_auctions_cancelled.csv` | zrušené aukce (mimo wide/long) |
| `source_survey.csv` | průzkum: rok × formát × odlišnosti |
| `validation_*.csv`, `validation_report.txt` | výsledky validace |
| `build_wide.log` | log běhu |

## Průzkum: rok × formát × odlišnosti

| Rok | Přípona | Hlavička (řádek, 0-based) | Formát | Aukcí | Dnů | Odlišnosti |
|---|---|---|---|---|---|---|
| 2012 | xls | 2 | starý | 75 | 73 | `Auction Name` = kód státu, fáze T2PA/T3PA/EAA2, bez `Certificate`, `Details` jen u společných aukcí, 27 národních aukcí bez `Total Revenue` |
| 2013–2015 | xls | 2 | starý | 211/198/207 | 191/198/205 | `Country` + `Auction Details` (víceřádkový text, tisícové tečky `2.309.100`), `Minimum/Maximum Bid` |
| 2016 | xls | 2 | starý | 195 | 195 | přibývá `Certificate`, `Auction\nDetails` jako `AT:431520,BE:…` bez teček, název `EU (TCAP)` |
| 2017–2018 | xls | 5 | nový | 206/209 | 201/203 | `Unnamed: 0`, 27 sloupců států `Name\n(XX)`, bez `Status` |
| 2019 | xls | 5 | nový | 216 | 208 | sloupec `Innovation\nFund` **bez kódu** (mapován na IF), přibývá IS, LI, NO |
| 2020–2021 | xlsx | 5 | nový | 218/230 | 210/223 | přibývá `Status` (zrušené aukce), IF, MF |
| 2022–2025 | xlsx | 5 | nový | 222/223/221/213 | 222/223/221/213 | XI (2022), IX a MX (2023), SF (2025) |

Všech 14 souborů existuje; hranice `.xls`/`.xlsx` sedí na 2019/2020. Soubory jsou ve vnořené složce (stejné jméno dvakrát), proto se hledá rekurzivně podle roku a obou přípon.

## Rozhodnutí

1. **Hlavička** se hledá jako první řádek, kde některá buňka = `Date`. Názvy sloupců se normalizují (`\n`, dvojité mezery, `€`/`EUR`, `Minimal/Minimum`, `Succesful`).
2. **Datum**: ve všech souborech je to nativní datetime (žádná záměna den/měsíc); skript navíc kontroluje, že rok data = rok souboru.
3. **Kódy států**: regex `\(([A-Z]{2})\)` z hlavičky; `Innovation Fund` bez kódu → `IF`. Sloupec `Country` není stát hodnot, je to **typ aukce** (EU = společná aukce EEX, DE, PL, XI = Severní Irsko) a shoduje se s `Auction Name`.
4. **Zachované kódy (38)**: AT BE BG CY CZ DE DK EE EL ES FI FR HR HU IE IF IS IT IX LI LT LU LV MF MT MX NL NO PL PT RO SE SF SI SK XA XF XI. Nic se nevyhazuje.
5. **Národní aukce bez rozpisu (2012–2016: DE, PL, HU, LT, CZ, NL, XF, XA)**: celý `Auction Volume` je přiřazen pořádajícímu státu (`volume_method = national_auction_full_volume`). V novém formátu je tomu tak také (hodnota jen ve sloupci DE/PL/XI).
6. **Klíč řádku wide**: `Date × auction_type × Certificate × Contract`. Samotné `Date` nestačí (až 2 aukce denně, 5–20 dnů ročně). `Date × auction_type × Certificate` je v datech jednoznačné (0 duplicit), `Contract` je přidán pro jistotu. Duplicitní řádky se nemlčí (vypisují se do logu a sčítají).
7. **Denní soubor**: součet všech aukcí dne (EUA i EUAA). Jsou v něm i sloupce `auction_volume_tco2_EUA/_EUAA`. Součty objemů sedí na wide.
8. **Zrušené aukce** (`Status = cancelled`: 2020-03-17 EU, 2022-02-01 EU, 2022-03-02 PL): nemají cenu ani výnos, nejsou započítány (jsou v `eu_ets_auctions_cancelled.csv`). Starý formát `Status` nemá; úspěšná = má cenu (všechny mají).
9. **Chybějící stát = 0** (konzistentně). 0 znamená „žádný objem přiřazen“, ne „stát v aukci nebyl“; viz `validation_state_history.csv`.
10. **Certifikát** se odvozuje z `Contract` (T2PA/T3PA → EUA, EAA2/EAA3 → EUAA); u 2016 sedí s explicitním sloupcem.
11. **`Auction Name`** je sjednocen do `auction_type`; původní text je v `auction_name_raw`.
12. **`Total Revenue`** chybí u 27 národních aukcí 2012; je doplněn jako cena × objem a označen `total_revenue_imputed = True`.

## Validace (výsledky v `validation_report.txt`)

- Součet objemů přes státy = `Auction Volume`: odchylka 0 % u všech 2 841 aukcí.
- Součet výnosů přes státy = `Total Revenue`: odchylka 0 % (důkaz interpretace výše).
- Řádky zdroj − zrušené = řádky wide pro každý rok; unikátní dny sedí (kromě dnů, kdy byla zrušena jediná aukce dne).
- Namátkově 20 aukcí (5 × 4 období) znovu přečteno přímo ze zdroje přes `xlrd`/`openpyxl`: shoda 20/20.

## Známá omezení a nejistoty

- **VB/UK/GB se v datech nevyskytuje** (ani v `Auction Details`, ani ve sloupcích), přestože list `Caption` kód UK uvádí. Neodhaduje se.
- **Chorvatsko**: první nenulový objem je až 2015-01-08, ne 2013.
- **Lucembursko**: poslední nenulový objem je 2024-11-13 (v 2023 a 2024 jen 6 aukcí). Neověřováno proti jiným zdrojům.
- **XA** (jediná aukce 2013-04-24, T2PA, 865 000 tCO2) není v žádném listu `Caption`. Kód je zachován, význam je nepotvrzený. **XF** je podle `Caption` Flandry (4 aukce, 2013).
- Kódy IF, MF, IX, MX, SF nejsou státy (Inovační fond, Modernizační fond, RRF, Sociální klimatický fond) a v souboru jsou zvlášť.
- `volume_tco2` je dopočet; surová čísla jsou v `revenue_eur_*`. Aukční cena je ve zdroji zaokrouhlena na 2 desetinná místa, výnos = cena × objem sedí přesně.
- Ostatní statistiky z nového formátu (počet nabídek, průměrná velikost nabídky…) se do wide nepřenášejí.
