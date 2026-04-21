"""Genera due file di esempio in examples/."""
from pathlib import Path
import pandas as pd

BASE = Path(__file__).resolve().parent

# Fornitore A - un solo foglio
df_a = pd.DataFrame({
    "Codice":        ["SOL-2000", "EQ-MEC10", "SMT-ST100", "CAV-PRO", "GEO-25", "LIFT-4T"],
    "Descrizione":   ["Sollevatore 2 colonne 4T", "Equilibratrice MEC10",
                       "Smontagomme ST100", "Cavalletto professionale",
                       "GEO 25 assetto ruote", "Sollevatore 4T forbice"],
    "Prezzo":        [2800.00, 1950.00, 2400.00, 180.00, 5600.00, 3200.00],
    "Trasporto":     [150.00, 80.00, 120.00, 15.00, 180.00, 140.00],
    "Installazione": [250.00, 0.00, 0.00, 0.00, 350.00, 250.00],
})
out_a = BASE / "fornitore_A.xlsx"
df_a.to_excel(out_a, index=False, sheet_name="Listino")

# Fornitore B - due fogli, intestazioni leggermente diverse, alcuni codici in comune
df_b_listino = pd.DataFrame({
    "cod.":          ["SOL-2000", "EQ-MEC10", "SMT-ST100", "ACC-KIT", "GEO-25"],
    "Descrizione prodotto": ["Sollevatore 2 col. 4 ton", "Equilibratrice MEC10 PRO",
                              "Smontagomme automatico ST100", "Kit accessori",
                              "GEO 25 - linea assetto"],
    "Prezzo netto":  [2750.00, 2100.00, 2400.00, 95.00, 5450.00],
    "Spese trasporto": [140.00, 80.00, 120.00, 10.00, 180.00],
    "Montaggio":     [280.00, 0.00, 0.00, 0.00, 320.00],
})
df_b_promo = pd.DataFrame({
    "Codice":        ["LIFT-3T", "PROMO-01"],
    "Desc":          ["Sollevatore 3T forbice", "Promo accessori"],
    "Prezzo":        [2900.00, 120.00],
    "Trasporto":     [130.00, 5.00],
    "Installazione": [220.00, 0.00],
})

out_b = BASE / "fornitore_B.xlsx"
with pd.ExcelWriter(out_b, engine="openpyxl") as writer:
    df_b_listino.to_excel(writer, index=False, sheet_name="Listino")
    df_b_promo.to_excel(writer, index=False, sheet_name="Promozioni")

# Fornitore C - CSV con separatore ;
df_c = pd.DataFrame({
    "SKU":           ["SOL-2000", "EQ-MEC10", "GEO-25", "NUOVO-01"],
    "Description":   ["2-post lift 4T", "Wheel balancer MEC10",
                       "GEO 25 alignment", "Nuovo prodotto C"],
    "Price":         [2900.00, 1890.00, 5700.00, 450.00],
    "Shipping":      [160.00, 75.00, 200.00, 25.00],
    "Installation":  [260.00, 0.00, 400.00, 50.00],
})
out_c = BASE / "fornitore_C.csv"
df_c.to_csv(out_c, sep=";", index=False)

print(f"Creati:\n  {out_a}\n  {out_b}\n  {out_c}")
