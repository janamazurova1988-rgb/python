# python

## Analýza prodejů z Excelu

Skript `analyza_prodeju.py` načte Excel se sloupci **datum, produkt, region, hodnota**,
spočítá měsíční souhrny podle produktu a regionu, vykreslí grafy a uloží výsledky.

### Rychlý start

```bash
pip install -r requirements.txt
python analyza_prodeju.py                      # použije data/ukazkova_data_prodeje.xlsx
python analyza_prodeju.py moje.xlsx -o vysledky --list Prodeje
```

### Výstupy (složka `vystup/`)

| Soubor | Obsah |
|---|---|
| `souhrn_prodeju.xlsx` | listy *Detail* (měsíc × produkt × region: součet, průměr, počet), *Mesic x Produkt*, *Mesic x Region*, *Produkt x Region* (vč. řádku/sloupce Celkem) a *Grafy* |
| `prodeje_podle_produktu.png` | měsíční vývoj prodejů po produktech |
| `prodeje_podle_regionu.png` | měsíční vývoj prodejů po regionech |
| `produkt_region_heatmapa.png` | součty za celé období: produkt × region |

### Poznámky

- Názvy sloupců nezávisí na velikosti písmen; řádky s chybným datem/hodnotou se vynechají.
- Funguje pro měsíční i denní záznamy (vše se sečte po kalendářních měsících).
- `vytvor_ukazkova_data.py` vygeneruje další smyšlená data (`--denni` pro denní záznamy).

## Výdeje léků pacientům

| Soubor | Obsah |
|---|---|
| `data/ukazkova_data_leky_pacienti.xlsx` | ukázková data 2025 (Léky, Pacienti, Výdeje, Přehled po měsících) |
| `data/ukazkova_data_leky_pacienti_2026.xlsx` | fiktivní data 2026 ve stejné struktuře |
| `reporty/report_2026_12.xlsx` | měsíční report prosinec 2026 |

```bash
python vytvor_data_leky.py                       # data 2025 → fiktivní data 2026
python vytvor_report_leky.py --rok 2026 --mesic 12   # report z dat 2026 → reporty/report_2026_12.xlsx
python vytvor_report_leky.py --data data/ukazkova_data_leky_pacienti.xlsx --rok 2025 --mesic 12
```
