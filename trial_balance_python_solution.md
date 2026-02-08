
# Automatická konverzia `zostava.xlsx` → `Trial Balance_xxxxx.xlsx` (Python riešenie)

## Zadanie používateľa

> Potrebujem nájsť spôsob, ako urobiť konverziu výstupu v súbore **`zostava.xlsx`**  
> na formát, ktorý je v súbore **`Trial Balance_xxxxx.xlsx`**.  
> Budem to chcieť spúšťať **opakovane každý mesiac**.  
> Chcem **najjednoduchší, ale robustný spôsob**.  
> Power Query sa teraz nechcem učiť – preferujem **Python**.

---

## Navrhované riešenie (prehľad)

Najvhodnejší kompromis medzi jednoduchosťou, robustnosťou a dlhodobou udržateľnosťou je:

**Python skript + konfiguračný súbor + Excel šablóna**

### Základná myšlienka

- `zostava.xlsx` je **mesačný vstup** (mení sa obsah, nie logika)
- `Trial Balance_template.xlsx` je **nemenná šablóna** (formátovanie, hlavičky)
- Python skript:
  - načíta zostavu,
  - nájde hlavičku dynamicky (nie pevný riadok),
  - namapuje stĺpce podľa konfigurácie,
  - normalizuje čísla (MD / D / znamienka),
  - zapíše výsledok do kópie šablóny.

Výsledok: **jeden príkaz = hotový Trial Balance**.

---

## Odporúčaná adresárová štruktúra

```
trial_balance_tool/
│
├─ convert_trial_balance.py
├─ config.yaml
│
├─ input/
│   └─ zostava.xlsx
│
├─ templates/
│   └─ Trial Balance_template.xlsx
│
└─ output/
    └─ Trial Balance_YYYY-MM-DD.xlsx
```

---

## Použité technológie

- `pandas` – práca s dátami
- `python-calamine` – robustné čítanie Excel súborov
- `openpyxl` – zápis do Excel šablóny so zachovaním formátovania
- `pyyaml` – konfiguračný súbor

### Inštalácia

```bash
python -m venv .venv
source .venv/bin/activate
pip install pandas openpyxl pyyaml python-calamine
```

---

## Konfiguračný súbor `config.yaml`

Konfigurácia oddeľuje **logiku mapovania** od kódu.

- stĺpce sa identifikujú pomocou regexov
- hlavička nemusí byť na prvom riadku
- jednoduché doladenie bez zásahu do skriptu

```yaml
source:
  sheet: null
  header_search_rows: 60
  required_columns:
    account: ["^účet$","^account$"]
    name: ["^názov.*$","^name$"]
    opening: ["^počiatoč.*$","^opening.*$"]
    debit: ["^obrat.*md$","^debit$"]
    credit: ["^obrat.*d$","^credit$"]
    closing: ["^koneč.*$","^closing.*$"]

transform:
  account_strip: true
  drop_zero_rows: true
  number_formats:
    decimal_comma: true
    thousands_sep_space: true

template:
  path: "templates/Trial Balance_template.xlsx"
  header_match:
    account: ["Account","Účet"]
    name: ["Account name","Názov účtu"]
    opening: ["Opening balance","Počiatočný stav"]
    debit: ["Debit","MD"]
    credit: ["Credit","D"]
    closing: ["Closing balance","Konečný stav"]
  write:
    clear_below_header: true
```

---

## Logika spracovania (high-level)

1. **Načítanie vstupu**
   - Excel sa načíta bez predpokladu pevnej štruktúry
   - hlavička sa hľadá v prvých N riadkoch

2. **Mapovanie stĺpcov**
   - na základe regexov v `config.yaml`
   - nezávislé od poradia stĺpcov

3. **Normalizácia dát**
   - konverzia čísiel z EU formátu (`1 234,56`)
   - MD a D sú vždy nezáporné
   - odstránenie nulových riadkov (voliteľné)

4. **Validácia**
   - kontrola:  
     `Closing ≈ Opening + Debit − Credit`
   - prípadné rozdiely len ako warning

5. **Zápis do šablóny**
   - zachovanie formátovania
   - prepis len dátovej časti

---

## Spustenie každý mesiac

```bash
source .venv/bin/activate
python convert_trial_balance.py   --input input/zostava.xlsx   --output "output/Trial Balance_2026-01-31.xlsx"
```

Časová náročnosť: **~10 sekúnd**  
Riziko chýb: **minimálne**

---

## Prečo je toto riešenie robustné

- žiadne pevné indexy riadkov ani stĺpcov
- zmeny názvov stĺpcov riešiš v konfigurácii
- formátovanie Excelu ostáva zachované
- vhodné aj pre audit / kontrolu

---

## Možné rozšírenia (neskôr)

- kontrolné súčty (Total Debit = Total Credit)
- logovanie rozdielov do samostatného sheetu
- CI / cron job
- export aj do CSV / Parquet

---

## Záver

Toto riešenie je:

- **jednoduché na používanie**
- **dostatočne robustné pre účtovníctvo**
- **udržateľné dlhodobo bez manuálnej práce**

Je navrhnuté presne na scenár:  
> „každý mesiac nový vstup, rovnaký výstupný formát, minimálny čas“.
