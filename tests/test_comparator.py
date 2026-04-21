"""Test di base per comparator."""
import os
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from checkapp.comparator import (  # noqa: E402
    CompareOptions,
    compare,
    load_source,
    run_comparison,
)


class TestComparator(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.mkdtemp()

        df1 = pd.DataFrame({
            "Codice": ["A001", "A002", "A003", "B100"],
            "Descrizione": ["Vite M6", "Dado M6", "Rondella", "Bullone"],
            "Prezzo": [0.10, 0.05, 0.02, 0.30],
            "Trasporto": [5, 5, 5, 10],
            "Installazione": [0, 0, 0, 15],
        })
        self.f1 = os.path.join(self.tmp, "fornitoreA.xlsx")
        df1.to_excel(self.f1, index=False)

        df2 = pd.DataFrame({
            "cod.": ["A001", "A002", "C500"],
            "Desc": ["Vite M6 inox", "Dado M6", "Staffa"],
            "Prezzo listino": [0.12, 0.05, 2.50],
            "Spedizione": [4, 4, 8],
            "Montaggio": [0, 0, 25],
        })
        self.f2 = os.path.join(self.tmp, "fornitoreB.xlsx")
        df2.to_excel(self.f2, index=False)

    def test_load_and_compare(self) -> None:
        options = CompareOptions(output_path=os.path.join(self.tmp, "out.xlsx"))
        result = run_comparison([self.f1, self.f2],
                                output_path=options.output_path,
                                options=options)
        self.assertTrue(os.path.exists(result["output"]))
        stats = result["stats"]
        # A001, A002, A003, B100, C500 = 5 codici distinti
        self.assertEqual(stats["totale_codici"], 5)
        self.assertEqual(stats["in_tutti"], 2)   # A001, A002
        self.assertEqual(stats["solo_in_uno"], 3)  # A003, B100, C500

    def test_case_insensitive_codes(self) -> None:
        df3 = pd.DataFrame({
            "Codice": ["a001"],
            "Descrizione": ["Vite M6 lowercase"],
            "Prezzo": [0.11],
        })
        f3 = os.path.join(self.tmp, "c.xlsx")
        df3.to_excel(f3, index=False)
        options = CompareOptions(
            output_path=os.path.join(self.tmp, "out2.xlsx"),
            case_sensitive_codes=False,
        )
        s1 = load_source(self.f1, options=options)
        s3 = load_source(f3, options=options)
        result = compare([s1, s3], options=options)
        codes = set(result["table"]["codice"])
        # 'a001' deve essere normalizzato e fondersi con 'A001'
        self.assertIn("A001", codes)
        self.assertNotIn("a001", codes)


if __name__ == "__main__":
    unittest.main()
