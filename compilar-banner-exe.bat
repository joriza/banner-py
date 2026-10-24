@echo off
rem Compila el ejecutable portable del cartel de letras huecas.
rem Requiere Python 3 con PyInstaller instalado (pip install pyinstaller).
rem Produce banner-portable.exe en esta misma carpeta.
python -m PyInstaller --onefile --console --name banner-portable --distpath . --workpath build --specpath build banner-portable.py
