"""GUI Tkinter per CheckappExcel.

Avvio:
    python -m checkapp.gui
oppure:
    python gui.py
"""
from __future__ import annotations

import os
import threading
from pathlib import Path
from tkinter import (
    BOTH, END, LEFT, RIGHT, Y, BooleanVar, StringVar, Tk, Toplevel, X,
    filedialog, messagebox, ttk,
)
from typing import List

from .comparator import CompareOptions, run_comparison


class CheckappGUI(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("CheckappExcel - Confronto listini")
        self.geometry("720x500")
        self.minsize(600, 420)

        self.files: List[str] = []
        self.labels: List[str] = []
        self.output_var = StringVar(value=str(Path.cwd() / "confronto.xlsx"))
        self.case_var = BooleanVar(value=False)
        self.merge_var = BooleanVar(value=True)

        self._build_ui()

    # ---------- UI ----------
    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}
        header = ttk.Label(
            self,
            text="Trascina o aggiungi i file Excel/CSV da confrontare",
            font=("Segoe UI", 12, "bold"),
        )
        header.pack(fill=X, **pad)

        frame_list = ttk.LabelFrame(self, text="File selezionati")
        frame_list.pack(fill=BOTH, expand=True, **pad)

        self.tree = ttk.Treeview(
            frame_list,
            columns=("file", "label"),
            show="headings",
            height=8,
        )
        self.tree.heading("file", text="File")
        self.tree.heading("label", text="Etichetta nel report")
        self.tree.column("file", width=430)
        self.tree.column("label", width=200)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True, padx=(8, 0), pady=8)

        scroll = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree.yview)
        scroll.pack(side=RIGHT, fill=Y, pady=8, padx=(0, 8))
        self.tree.configure(yscrollcommand=scroll.set)

        btns = ttk.Frame(self)
        btns.pack(fill=X, **pad)
        ttk.Button(btns, text="➕ Aggiungi file…", command=self.add_files
                   ).pack(side=LEFT, padx=4)
        ttk.Button(btns, text="✏️  Rinomina etichetta", command=self.rename_label
                   ).pack(side=LEFT, padx=4)
        ttk.Button(btns, text="🗑  Rimuovi selezionato", command=self.remove_selected
                   ).pack(side=LEFT, padx=4)
        ttk.Button(btns, text="🧹 Svuota lista", command=self.clear_all
                   ).pack(side=LEFT, padx=4)

        frame_out = ttk.LabelFrame(self, text="Opzioni")
        frame_out.pack(fill=X, **pad)

        row = ttk.Frame(frame_out)
        row.pack(fill=X, padx=8, pady=6)
        ttk.Label(row, text="File di output:").pack(side=LEFT)
        ttk.Entry(row, textvariable=self.output_var).pack(
            side=LEFT, fill=X, expand=True, padx=6
        )
        ttk.Button(row, text="Sfoglia…", command=self.pick_output).pack(side=LEFT)

        row2 = ttk.Frame(frame_out)
        row2.pack(fill=X, padx=8, pady=4)
        ttk.Checkbutton(
            row2,
            text="Unisci i fogli dello stesso file",
            variable=self.merge_var,
        ).pack(side=LEFT, padx=4)
        ttk.Checkbutton(
            row2,
            text="Codici case-sensitive",
            variable=self.case_var,
        ).pack(side=LEFT, padx=4)

        run_row = ttk.Frame(self)
        run_row.pack(fill=X, **pad)
        self.run_btn = ttk.Button(
            run_row, text="▶  Avvia confronto", command=self.start_compare
        )
        self.run_btn.pack(side=RIGHT)
        self.status = ttk.Label(run_row, text="Pronto.", anchor="w")
        self.status.pack(side=LEFT, fill=X, expand=True)

    # ---------- azioni ----------
    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Seleziona file Excel o CSV",
            filetypes=[
                ("Excel/CSV", "*.xlsx *.xls *.xlsm *.csv *.tsv"),
                ("Tutti i file", "*.*"),
            ],
        )
        for p in paths:
            if p in self.files:
                continue
            self.files.append(p)
            label = Path(p).stem
            self.labels.append(label)
            self.tree.insert("", END, values=(p, label))

    def rename_label(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        idx = self.tree.index(item_id)
        current = self.labels[idx]
        win = Toplevel(self)
        win.title("Rinomina etichetta")
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text="Nuova etichetta:").pack(padx=10, pady=8)
        var = StringVar(value=current)
        entry = ttk.Entry(win, textvariable=var, width=40)
        entry.pack(padx=10, pady=4)
        entry.focus_set()
        entry.select_range(0, END)

        def ok() -> None:
            new_label = var.get().strip() or current
            self.labels[idx] = new_label
            self.tree.item(item_id, values=(self.files[idx], new_label))
            win.destroy()

        ttk.Button(win, text="OK", command=ok).pack(pady=8)
        win.bind("<Return>", lambda _e: ok())

    def remove_selected(self) -> None:
        for item_id in self.tree.selection():
            idx = self.tree.index(item_id)
            del self.files[idx]
            del self.labels[idx]
            self.tree.delete(item_id)

    def clear_all(self) -> None:
        self.files.clear()
        self.labels.clear()
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)

    def pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Salva come",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="confronto.xlsx",
        )
        if path:
            self.output_var.set(path)

    def start_compare(self) -> None:
        if len(self.files) < 2:
            messagebox.showwarning(
                "Attenzione",
                "Servono almeno 2 file per fare un confronto.",
            )
            return
        output = self.output_var.get().strip()
        if not output:
            messagebox.showwarning("Attenzione", "Specifica un file di output.")
            return

        self.run_btn.configure(state="disabled")
        self.status.configure(text="Confronto in corso…")

        thread = threading.Thread(
            target=self._run_compare_thread,
            args=(list(self.files), list(self.labels), output),
            daemon=True,
        )
        thread.start()

    def _run_compare_thread(self, files: List[str], labels: List[str],
                            output: str) -> None:
        try:
            options = CompareOptions(
                output_path=output,
                case_sensitive_codes=self.case_var.get(),
                merge_sheets=self.merge_var.get(),
            )
            result = run_comparison(files, output_path=output,
                                    labels=labels, options=options)
            stats = result["stats"]
            msg = (
                f"File creato:\n{result['output']}\n\n"
                f"Codici totali: {stats['totale_codici']}\n"
                f"In tutti: {stats['in_tutti']}\n"
                f"Parziali: {stats['parziali']}\n"
                f"Solo in uno: {stats['solo_in_uno']}"
            )
            self.after(0, lambda: self._on_done(msg, True, result["output"]))
        except Exception as exc:  # noqa: BLE001 - mostrata all'utente
            self.after(0, lambda: self._on_done(str(exc), False, None))

    def _on_done(self, msg: str, ok: bool, path: str | None) -> None:
        self.run_btn.configure(state="normal")
        if ok:
            self.status.configure(text="Completato.")
            if messagebox.askyesno("Completato",
                                   msg + "\n\nVuoi aprire la cartella?"):
                if path:
                    folder = os.path.dirname(os.path.abspath(path))
                    _open_folder(folder)
        else:
            self.status.configure(text="Errore.")
            messagebox.showerror("Errore", msg)


def _open_folder(path: str) -> None:
    import platform
    import subprocess
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception:
        pass


def main() -> None:
    app = CheckappGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
