@echo off
REM Ejecutar este archivo en Windows (doble clic) para generar un .exe
REM independiente que el departamento legal pueda usar sin instalar Python.
REM Requiere tener Python instalado UNA VEZ en el equipo donde se genera el .exe.

echo Instalando dependencias...
pip install -r requirements.txt
pip install pyinstaller

echo Generando ejecutable...
pyinstaller --noconfirm --onefile --windowed --name "LawyGen" --icon "assets\lawygen_icon.ico" --add-data "assets;assets" app.py

echo.
echo Listo. El ejecutable esta en la carpeta dist\LawyGen.exe
echo Puedes copiar ese unico archivo .exe a cualquier PC del despacho y
echo funcionara sin necesidad de instalar Python en esos equipos.
pause
