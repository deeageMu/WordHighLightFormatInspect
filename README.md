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
`WordHighLightFormatInspect.exe` or running `python3 main.py`), a Tkinter
interface for the core features starts. The GUI supports Deutsch, English,
and Français. It detects the system language on first start and falls back to
English for unsupported languages.

The language menu opens from the menu bar and offers all three languages.
Manual selection is saved and overrides system-language detection on future
starts. The current report is updated in the selected language on the next
scan.

The setting is stored as UTF-8 JSON under the user's standard configuration
directory, for example:

```json
{ "language": "de" }
```

The GUI report is translated, while the command-line interface always uses
English.

The interface can also be forced with `--gui` when arguments are present.
Tkinter is part of the Python standard library, so no additional installation
is required.

## Build Windows EXE

### Docker

Build the Windows-compatible EXE from Linux or macOS with Docker. The source
directory is mounted into the container and the generated file is written to
`dist/WordHighLightFormatInspect.exe`:

```bash
docker build -f Dockerfile.windows -t wordhighlightformatinspect-windows .
mkdir -p dist
docker run --rm \
    -v "$PWD:/src" \
    -v "$PWD/dist:/output" \
    wordhighlightformatinspect-windows
```

The Docker image uses Wine with Windows Python because PyInstaller cannot
cross-compile a Windows executable from a Linux Python installation.

### Windows

```bash
pyinstaller --onefile --windowed --name WordHighLightFormatInspect main.py
```

Creates: `dist/WordHighLightFormatInspect.exe`

> Build the Windows EXE only on Windows. In Linux or Codespaces, the output is a Linux binary even if the filename ends with .exe.
> PyInstaller usually includes Tkinter automatically; if the GUI is missing
> from the built `.exe`, add `tkinter` explicitly with
> `--hidden-import tkinter`.

The Windows build uses `--windowed`, so starting the EXE by double-clicking it
does not open a DOS console window. A separate console build is available for
CLI usage:

```bash
pyinstaller --onefile --name WordHighLightFormatInspect-console main.py
```

The console build is written to `dist/WordHighLightFormatInspect-console.exe`
on Windows, or `dist/WordHighLightFormatInspect-console` on Linux.

## GitHub Releases

Pushing a tag such as `v1.0.0` starts the release workflow. It builds and
publishes these four assets automatically:

- `WordHighLightFormatInspect.exe` - Windows GUI
- `WordHighLightFormatInspect-console.exe` - Windows CLI
- `WordHighLightFormatInspect` - Linux GUI
- `WordHighLightFormatInspect-console` - Linux CLI

The same workflow can be started manually to create downloadable workflow
artifacts without publishing a release.