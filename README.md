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

## Fiktivní měsíční report výdejů léků

```bash
python vytvor_report_leky.py --rok 2026 --mesic 12   # → reporty/report_2026_12.xlsx
```

Listy *Report*, *Grafy* a *Výdeje v měsíci* – stejná struktura jako `report_2025_12.xlsx`.
