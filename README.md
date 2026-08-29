# WordHighLightFormatInspect

## Usage (Command Line)

```bash
cd /workspaces/WordHighLightFormatInspect
python3 -m pip install -r requirements.txt
python3 main.py document.docx
python3 main.py document.docx --out-xml
python3 main.py document.docx --out-xml --xml-max-chars 0
python3 main.py document.docx --annotate document_annotated.docx
zip -r WordHighLightFormatInspect.zip . -x "*.git*" "tmp/*" "build/*" "dist/*"
```

`--annotate OUTPUT.docx` creates a **copy** of the inspected file in which
each detected formatting feature (highlighting, character/paragraph/table-cell
shading, text color, and character style) is marked as a native Word comment
at the relevant text location. The original file is opened read-only and
remains unchanged; the source and destination files must not be identical.

## Usage (Windows Interface)

Without arguments (for example, by double-clicking
`WordHighLightFormatInspect.exe` or running `python3 main.py`), a basic Tkinter
interface limited to the core features starts:

1. **Datei waehlen...** - Choose a `.docx` file using the file dialog
2. **Scannen** - Display the formatting report
3. **Kommentierte Kopie erstellen...** - Choose where to save the new
    annotated file (the original remains unchanged)

The interface can also be forced with `--gui` when arguments are present.
Tkinter is part of the Python standard library, so no additional installation
is required.

## Build Windows EXE

```bash
pyinstaller --onefile --name WordHighLightFormatInspect main.py
```

Creates: `dist/WordHighLightFormatInspect.exe`

> Build the Windows EXE only on Windows. In Linux or Codespaces, the output is a Linux binary even if the filename ends with .exe.
> PyInstaller usually includes Tkinter automatically; if the GUI is missing
> from the built `.exe`, add `tkinter` explicitly with
> `--hidden-import tkinter`.