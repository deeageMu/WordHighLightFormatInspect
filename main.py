"""
Einstiegspunkt der Anwendung.

Ohne Argumente (z.B. per Doppelklick auf die .exe unter Windows) startet
die rudimentaere Tkinter-Oberflaeche zum Auswaehlen und Scannen einer
.docx-Datei sowie zum Erstellen einer kommentierten Kopie.

Mit Argumenten wird weiterhin die bestehende Kommandozeilenvariante
genutzt (siehe docx_format_scanner.main / README.md). '--gui' erzwingt
die Oberflaeche auch bei vorhandenen Argumenten.
"""

import sys

from docx_format_scanner import main as cli_main


def main() -> int:
    argv = sys.argv[1:]
    if not argv or "--gui" in argv:
        from gui import main as gui_main

        return gui_main()
    return cli_main()


if __name__ == "__main__":
    sys.exit(main())
