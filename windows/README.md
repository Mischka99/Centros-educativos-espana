# Paquete portable para Windows

Así se monta `CentrosEducativos.zip`, el paquete que se publica en Releases. Solo lleva el Python oficial de python.org, firmado por la Python Software Foundation, más archivos de texto e imágenes. No hay ningún ejecutable propio, así que el Control inteligente de aplicaciones de Windows 11 no lo bloquea.

## Estructura

```
CentrosEducativos/
├── Centros educativos.exe   ← pythonw.exe oficial con otro nombre
├── LEEME.txt                ← windows/LEEME.txt
├── python313.dll, python3.dll, vcruntime140.dll, vcruntime140_1.dll, LICENSE-Python.txt
├── DLLs/  tcl/  Lib/        ← biblioteca estándar, Tkinter y Tcl/Tk
│   └── Lib/site-packages/   ← openpyxl, et_xmlfile, xlrd y sitecustomize.py
└── programa/                ← el programa, la lista, README.md y LICENSE.txt
```

## Pasos (en Windows)

1. Descarga el instalador oficial de 64 bits de Python 3.13 desde python.org. La versión 1.0.0 usa `python-3.13.17-amd64.exe`.
2. Descarga sus paquetes MSI sin instalar nada:
   ```
   python-3.13.17-amd64.exe /layout C:\py313
   ```
3. Extrae en la carpeta del paquete los MSI `core`, `exe`, `lib` y `tcltk` con una instalación administrativa:
   ```
   msiexec /a C:\py313\core.msi TARGETDIR=C:\CentrosEducativos /qn
   msiexec /a C:\py313\exe.msi TARGETDIR=C:\CentrosEducativos /qn
   msiexec /a C:\py313\lib.msi TARGETDIR=C:\CentrosEducativos /qn
   msiexec /a C:\py313\tcltk.msi TARGETDIR=C:\CentrosEducativos /qn
   ```
   `msiexec` deja una copia de cada `.msi` en esa carpeta: bórralas.
4. Quita lo que el programa no usa. Así, los únicos ejecutables que quedan son los 37 firmados de Python:
   - `python.exe` y `NEWS.txt`.
   - `DLLs\py.ico`, `DLLs\pyc.ico` y `DLLs\pyd.ico`.
   - La carpeta `libs`.
   - `Lib\ensurepip`, `Lib\idlelib` y `Lib\turtledemo`.
   - `tcl\nmake` (lleva un `.exe` sin firmar que solo sirve para compilar), los `tcl\*.lib` y los `tcl\*.sh`.
5. Cambia dos nombres:
   - `pythonw.exe` pasa a llamarse `Centros educativos.exe`. La firma sigue siendo válida, porque cambiar el nombre no modifica el archivo.
   - `LICENSE.txt` pasa a llamarse `LICENSE-Python.txt`.
6. Añade las dependencias, que son Python puro, con cualquier Python que tengas instalado:
   ```
   py -m pip install --target C:\CentrosEducativos\Lib\site-packages --no-compile openpyxl==3.1.5 xlrd==2.0.2
   ```
   pip instala también et_xmlfile, que necesita openpyxl (la versión 1.0.0 lleva et_xmlfile 2.0.0).
7. Copia `windows\sitecustomize.py` en `Lib\site-packages`.
8. Crea la carpeta `programa\` y copia en ella, desde este repositorio:
   - `centros_educativos.py`, `centros_gui.pyw`, `interfaz.py`, `mapa.py`, `icono.ico`, `icono.png` y `README.md`.
   - Las carpetas `img\` y `datos\`.
   - `LICENSE`, con el nombre `LICENSE.txt`.

   Copia también `windows\LEEME.txt` en la raíz del paquete.
9. Opcional: marca como ocultos los archivos de Python para que la carpeta solo muestre «Centros educativos», «LEEME» y «programa»:
   ```
   attrib +h C:\CentrosEducativos\*.dll
   attrib +h C:\CentrosEducativos\LICENSE-Python.txt
   attrib +h C:\CentrosEducativos\DLLs
   attrib +h C:\CentrosEducativos\Lib
   attrib +h C:\CentrosEducativos\tcl
   ```
10. Comprime la carpeta `CentrosEducativos`. Comprueba antes que no quede ninguna carpeta `__pycache__`.

Desde Linux se puede hacer lo mismo. `cabextract` saca los MSI del instalador con nombres como `a0` o `a2`, y el manifiesto `0` indica cuál es cada uno (`core`, `exe`, `lib`, `tcltk`…). Después, `msiextract` (del paquete msitools) los extrae.

## Cómo arranca

Cuando se abre `Centros educativos.exe` sin argumentos, Python carga `Lib\site-packages\sitecustomize.py`. Ese archivo comprueba el nombre del ejecutable y lanza `programa\centros_gui.pyw`. Si el programa falla al arrancar, el error se guarda en `programa\error.log`.

Con cualquier otro nombre, o con argumentos, el Python del paquete funciona con normalidad: `sitecustomize.py` solo actúa en ese caso concreto.
