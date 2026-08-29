"""
Einstiegspunkt der Anwendung.
Delegiert an docx_format_scanner.main(), damit PyInstaller ein
eigenstaendiges main.py als Einstiegsdatei nutzen kann.
"""

import sys

from docx_format_scanner import main

if __name__ == "__main__":
    sys.exit(main())