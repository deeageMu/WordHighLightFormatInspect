# WordHighLightFormatInspect

## Usage

```bash
cd /workspaces/WordHighLightFormatInspect
python3 -m pip install -r requirements.txt
python3 main.py document.docx
python3 main.py document.docx --out-xml
python3 main.py document.docx --out-xml --xml-max-chars 0
zip -r WordHighLightFormatInspect.zip . -x "*.git*" "tmp/*" "build/*" "dist/*"
```

## Build Windows EXE

```bash
pyinstaller --onefile --name WordHighLightFormatInspect main.py
```

Creates: `dist/WordHighLightFormatInspect.exe`

> Build the Windows EXE only on Windows. In Linux or Codespaces, the output is a Linux binary even if the filename ends with .exe.