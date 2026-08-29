"""
gui.py

Rudimentaere, auf die Kernfunktionen beschraenkte Windows-Oberflaeche
(Tkinter, Teil der Python-Standardbibliothek - keine zusaetzliche
Abhaengigkeit noetig).

Funktionen:
- Datei auswaehlen (Dateidialog)
- Datei scannen und Bericht anzeigen
- Kommentierte Kopie erstellen (Original bleibt unveraendert)

Alle Fehler werden ueber Dialogfenster gemeldet statt das Programm
abstuerzen zu lassen.
"""

from __future__ import annotations

import sys
import traceback
import zipfile
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from comment_writer import annotate_docx
from docx_format_scanner import _format_bericht, scan_docx


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("WordHighLightFormatInspect")
        self.geometry("800x580")
        self.minsize(600, 400)
        self.pfad: Path | None = None

        auswahl_zeile = tk.Frame(self)
        auswahl_zeile.pack(fill="x", padx=8, pady=8)

        self.pfad_label = tk.Label(auswahl_zeile, text="Keine Datei ausgewaehlt", anchor="w")
        self.pfad_label.pack(side="left", fill="x", expand=True)

        tk.Button(auswahl_zeile, text="Datei waehlen...", command=self._datei_waehlen).pack(
            side="right"
        )

        button_zeile = tk.Frame(self)
        button_zeile.pack(fill="x", padx=8, pady=(0, 8))

        self.scan_button = tk.Button(
            button_zeile, text="Scannen", command=self._scannen, state="disabled"
        )
        self.scan_button.pack(side="left")

        self.annotate_button = tk.Button(
            button_zeile,
            text="Kommentierte Kopie erstellen...",
            command=self._annotieren,
            state="disabled",
        )
        self.annotate_button.pack(side="left", padx=(8, 0))

        self.text = scrolledtext.ScrolledText(self, wrap="word", font=("Courier New", 9))
        self.text.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.text.configure(state="disabled")

        self.status = tk.Label(self, text="Bereit.", anchor="w", relief="sunken")
        self.status.pack(fill="x", side="bottom")

    def _set_text(self, inhalt: str) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", inhalt)
        self.text.configure(state="disabled")

    def _datei_waehlen(self) -> None:
        pfad = filedialog.askopenfilename(
            title="Word-Datei auswaehlen",
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")],
        )
        if not pfad:
            return
        self.pfad = Path(pfad)
        self.pfad_label.configure(text=str(self.pfad))
        self.scan_button.configure(state="normal")
        self.annotate_button.configure(state="normal")
        self.status.configure(text="Datei ausgewaehlt. Bitte scannen.")
        self._set_text("")

    def _scannen(self) -> None:
        if self.pfad is None:
            return
        try:
            result = scan_docx(self.pfad)
            bericht = _format_bericht(result, self.pfad.name)
            self._set_text(bericht)
            self.status.configure(text=f"{len(result.findings)} Fundstelle(n) gefunden.")
        except zipfile.BadZipFile:
            messagebox.showerror("Fehler", "Datei ist kein gueltiges .docx (kein ZIP-Archiv).")
        except Exception as exc:  # rudimentaere Fehlerbehandlung fuer die Oberflaeche
            traceback.print_exc()
            messagebox.showerror("Fehler beim Scannen", str(exc))

    def _annotieren(self) -> None:
        if self.pfad is None:
            return
        vorschlag = self.pfad.with_name(self.pfad.stem + "_kommentiert.docx")
        ziel = filedialog.asksaveasfilename(
            title="Kommentierte Kopie speichern unter",
            initialfile=vorschlag.name,
            initialdir=str(vorschlag.parent),
            defaultextension=".docx",
            filetypes=[("Word-Dokumente", "*.docx")],
        )
        if not ziel:
            return
        ziel_pfad = Path(ziel)
        if ziel_pfad.resolve() == self.pfad.resolve():
            messagebox.showerror(
                "Fehler", "Zieldatei darf nicht mit der Originaldatei identisch sein."
            )
            return
        try:
            anzahl = annotate_docx(self.pfad, ziel_pfad)
            self.status.configure(
                text=f"Fertig: {anzahl} Kommentar(e) in {ziel_pfad.name} geschrieben."
            )
            messagebox.showinfo(
                "Fertig",
                f"{anzahl} Kommentar(e) wurden in eine neue Datei geschrieben:\n{ziel_pfad}\n\n"
                "Die Original-Datei wurde nicht veraendert.",
            )
        except zipfile.BadZipFile:
            messagebox.showerror("Fehler", "Datei ist kein gueltiges .docx (kein ZIP-Archiv).")
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror("Fehler beim Erstellen der Kopie", str(exc))


def main() -> int:
    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
