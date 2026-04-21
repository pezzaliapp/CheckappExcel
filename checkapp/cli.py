"""CLI di CheckappExcel.

Esempi:
    python -m checkapp fornitore_a.xlsx fornitore_b.xlsx -o confronto.xlsx
    python -m checkapp *.xlsx -o risultato.xlsx --no-merge-sheets
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List

from .comparator import CompareOptions, run_comparison


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="checkapp",
        description=(
            "Confronta due o più file Excel/CSV di listini prodotti usando "
            "il codice come chiave e produce un Excel con colonne colorate."
        ),
    )
    p.add_argument("files", nargs="+", help="File da confrontare (xlsx, xls, csv).")
    p.add_argument(
        "-o", "--output",
        default="confronto.xlsx",
        help="Percorso file di output (default: confronto.xlsx).",
    )
    p.add_argument(
        "-l", "--labels",
        nargs="+",
        default=None,
        help="Etichette da mostrare nel report (una per file). "
             "Se omesse si usa il nome del file.",
    )
    p.add_argument(
        "--case-sensitive",
        action="store_true",
        help="Confronto codici sensibile a maiuscole/minuscole (default: no).",
    )
    p.add_argument(
        "--no-merge-sheets",
        action="store_true",
        help="Non unire i fogli di uno stesso file: ogni foglio diventa una "
             "colonna separata nel confronto.",
    )
    return p


def main(argv: List[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    files = args.files
    for f in files:
        if not Path(f).exists():
            print(f"[ERRORE] File non trovato: {f}", file=sys.stderr)
            return 2

    if args.labels and len(args.labels) != len(files):
        print("[ERRORE] Numero di --labels diverso dal numero di file.",
              file=sys.stderr)
        return 2

    options = CompareOptions(
        output_path=args.output,
        case_sensitive_codes=args.case_sensitive,
        merge_sheets=not args.no_merge_sheets,
    )

    print(f"→ Confronto di {len(files)} file...")
    result = run_comparison(files, output_path=args.output,
                            labels=args.labels, options=options)
    stats = result["stats"]
    print(f"✓ Creato: {result['output']}")
    print(f"   Codici totali   : {stats['totale_codici']}")
    print(f"   In tutti i file : {stats['in_tutti']}")
    print(f"   Parziali        : {stats['parziali']}")
    print(f"   Solo in uno     : {stats['solo_in_uno']}")
    print("   Per file:")
    for label, n in stats["per_file"].items():
        print(f"     - {label}: {n}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
