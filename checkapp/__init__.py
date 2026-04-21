"""CheckappExcel - confronto prodotti fra file Excel/CSV."""

from .comparator import (
    CompareOptions,
    LoadedSource,
    compare,
    export_to_excel,
    load_source,
    run_comparison,
)

__version__ = "1.0.0"

__all__ = [
    "CompareOptions",
    "LoadedSource",
    "compare",
    "export_to_excel",
    "load_source",
    "run_comparison",
]
