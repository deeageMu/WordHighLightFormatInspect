# WordHighLightFormatInspect

cd /workspaces/WordHighLightFormatInspect
python3 -m pip install -r requirements.txt
pyinstaller --onefile --name WordHighLightFormatInspect main.py

Hinweis: PyInstaller ergänzt auf Windows automatisch die .exe-Dateiendung.
Ein Windows-EXE muss auf einem Windows-Runner bzw. Windows-System gebaut werden; ein Build im Linux-Codespace erzeugt eine Linux-Anwendung, auch wenn der Name .exe endet.