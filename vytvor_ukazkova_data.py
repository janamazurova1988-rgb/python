"""Vygeneruje ukázkový Excel se smyšlenými daty prodejů.

Výstup: data/ukazkova_data_vygenerovana.xlsx se sloupci datum, produkt, region, hodnota.
Hodnoty jsou náhodné (s pevným seedem, takže výsledek je pokaždé stejný).

Spuštění:
    python vytvor_ukazkova_data.py
    python vytvor_ukazkova_data.py --denni     # denní záznamy místo měsíčních
"""

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

PRODUKTY = {
    # produkt: (základní měsíční prodej, sezónní výkyv 0–1, vyšší v zimě)
    "Analgin Forte": (1000, 0.10),
    "Cardiovit 10 mg": (750, 0.00),
    "Respira Sprej": (700, 0.45),
    "Dermacal Krém": (500, -0.20),  # záporný výkyv = vyšší v létě
    "Imunex Plus": (650, 0.40),
}
REGIONY = {"Praha": 1.6, "Brno": 1.0, "Ostrava": 0.85, "Plzeň": 0.65}


def vytvor_data(rok: int = 2025, denni: bool = False, seed: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    if denni:
        datumy = pd.date_range(f"{rok}-01-01", f"{rok}-12-31", freq="D")
    else:
        datumy = pd.date_range(f"{rok}-01-01", periods=12, freq="MS")

    radky = []
    for datum in datumy:
        # kosinus má maximum v lednu a minimum v červenci
        zima = np.cos((datum.month - 1) / 12 * 2 * np.pi)
        for produkt, (zaklad, sezona) in PRODUKTY.items():
            for region, vaha in REGIONY.items():
                hodnota = zaklad * vaha * (1 + sezona * zima) * rng.normal(1, 0.08)
                if denni:
                    hodnota /= datum.days_in_month
                radky.append((datum, produkt, region, max(0, round(hodnota))))

    return pd.DataFrame(radky, columns=["datum", "produkt", "region", "hodnota"])


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("-o", "--vystup", default="data/ukazkova_data_vygenerovana.xlsx")
    parser.add_argument("--rok", type=int, default=2025)
    parser.add_argument("--denni", action="store_true", help="denní místo měsíčních záznamů")
    args = parser.parse_args()

    df = vytvor_data(args.rok, args.denni)
    cesta = Path(args.vystup)
    cesta.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(cesta, sheet_name="Data", index=False)
    print(f"Uloženo {len(df)} řádků do {cesta}")


if __name__ == "__main__":
    main()
