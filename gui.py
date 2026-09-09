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
from i18n import LANGUAGES, detect_system_language, get_translator
from settings import initial_language, save_language


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("WordHighLightFormatInspect")
        self.geometry("800x580")
        self.minsize(600, 400)
        self.pfad: Path | None = None
        self.language = initial_language(system_language=detect_system_language())
        self.translate = get_translator(self.language)

        self.menu_bar = tk.Menu(self, tearoff=False)
        self.language_menu = tk.Menu(self.menu_bar, tearoff=False)
        self.menu_bar.add_cascade(label="Language", menu=self.language_menu)
        self.configure(menu=self.menu_bar)
        self.menu_bar.bind("<Motion>", self._open_language_menu_on_hover)
        self.language_var = tk.StringVar(value=self.language)
        self._update_language_menu()

        auswahl_zeile = tk.Frame(self)
        auswahl_zeile.pack(fill="x", padx=8, pady=8)

        self.pfad_label = tk.Label(auswahl_zeile, anchor="w")
        self.pfad_label.pack(side="left", fill="x", expand=True)

        self.choose_button = tk.Button(auswahl_zeile, command=self._datei_waehlen)
        self.choose_button.pack(side="right")

        button_zeile = tk.Frame(self)
        button_zeile.pack(fill="x", padx=8, pady=(0, 8))

        self.scan_button = tk.Button(
            button_zeile, command=self._scannen, state="disabled"
        )
        self.scan_button.pack(side="left")

        self.annotate_button = tk.Button(
            button_zeile,
            command=self._annotieren,
            state="disabled",
        )
        self.annotate_button.pack(side="left", padx=(8, 0))

        self.text = scrolledtext.ScrolledText(self, wrap="word", font=("Courier New", 9))
        self.text.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.text.configure(state="disabled")

        self.status = tk.Label(self, anchor="w", relief="sunken")
        self.status.pack(fill="x", side="bottom")
        self._update_language_text()

    def _update_language_menu(self) -> None:
        self.language_menu.delete(0, "end")
        for language in LANGUAGES:
            self.language_menu.add_radiobutton(
                label=self.translate(f"language.{language}"),
                variable=self.language_var,
                value=language,
                command=lambda selected=language: self._select_language(selected),
            )
        self.menu_bar.entryconfigure(0, label=self.translate("language"))

    def _open_language_menu_on_hover(self, event: tk.Event) -> None:
        entry_box = self.menu_bar.entrybbox(0)
        if entry_box is None:
            return
        entry_x, entry_y, entry_width, entry_height = entry_box
        if entry_x <= event.x <= entry_x + entry_width and entry_y <= event.y <= entry_y + entry_height:
            self.language_menu.post(
                self.menu_bar.winfo_rootx() + entry_x,
                self.menu_bar.winfo_rooty() + entry_y + entry_height,
            )

    def _update_language_text(self) -> None:
        self.title(self.translate("app.title"))
        self.pfad_label.configure(text=self.translate("no_file"))
        self.choose_button.configure(text=self.translate("choose_file"))
        self.scan_button.configure(text=self.translate("scan"))
        self.annotate_button.configure(text=self.translate("create_copy"))
        self.status.configure(text=self.translate("ready"))

    def _select_language(self, language: str) -> None:
        self.language = language
        self.translate = get_translator(language)
        save_language(language)
        self.language_var.set(language)
        self._update_language_menu()
        self._update_language_text()

    def _set_text(self, inhalt: str) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", inhalt)
        self.text.configure(state="disabled")

    def _datei_waehlen(self) -> None:
        pfad = filedialog.askopenfilename(
            title=self.translate("word_file"),
            filetypes=[
                (self.translate("word_documents"), "*.docx"),
                (self.translate("all_files"), "*.*"),
            ],
        )
        if not pfad:
            return
        self.pfad = Path(pfad)
        self.pfad_label.configure(text=str(self.pfad))
        self.scan_button.configure(state="normal")
        self.annotate_button.configure(state="normal")
        self.status.configure(text=self.translate("file_selected"))
        self._set_text("")

    def _scannen(self) -> None:
        if self.pfad is None:
            return
        try:
            result = scan_docx(self.pfad)
            bericht = _format_bericht(result, self.pfad.name, translate=self.translate)
            self._set_text(bericht)
            self.status.configure(text=self.translate("findings").format(count=len(result.findings)))
        except zipfile.BadZipFile:
            messagebox.showerror(self.translate("error"), self.translate("invalid_docx"))
        except Exception as exc:  # rudimentaere Fehlerbehandlung fuer die Oberflaeche
            traceback.print_exc()
            messagebox.showerror(self.translate("scan_error"), str(exc))

    def _annotieren(self) -> None:
        if self.pfad is None:
            return
        suffix = self.translate("suggested_copy")
        vorschlag = self.pfad.with_name(self.pfad.stem + suffix)
        ziel = filedialog.asksaveasfilename(
            title=self.translate("save_copy"),
            initialfile=vorschlag.name,
            initialdir=str(vorschlag.parent),
            defaultextension=".docx",
            filetypes=[("Word-Dokumente", "*.docx")],
        )
        if not ziel:
            return
        ziel_pfad = Path(ziel)
        if ziel_pfad.resolve() == self.pfad.resolve():
            messagebox.showerror(self.translate("error"), self.translate("same_file"))
            return
        try:
            anzahl = annotate_docx(self.pfad, ziel_pfad)
            self.status.configure(
                text=self.translate("finished").format(count=anzahl, name=ziel_pfad.name)
            )
            messagebox.showinfo(
                self.translate("finished_title"),
                self.translate("comments_written").format(count=anzahl, path=ziel_pfad),
            )
        except zipfile.BadZipFile:
            messagebox.showerror(self.translate("error"), self.translate("invalid_docx"))
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror(self.translate("copy_error"), str(exc))


def main() -> int:
    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
