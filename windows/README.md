# Paquete de Windows

Cada versión publica dos archivos con el mismo contenido:

- **`CentrosEducativos.msi`:** instalador. Instala para el usuario, sin permisos de administrador, en `%LOCALAPPDATA%\Programs\Centros educativos`, y crea los accesos directos del menú Inicio y del escritorio. Se desinstala desde Configuración > Aplicaciones.
- **`CentrosEducativos.zip`:** portable. Se descomprime en cualquier carpeta y se abre con doble clic.

Los dos llevan solo el Python oficial de python.org, firmado por la Python Software Foundation y Microsoft, más archivos de texto e imágenes. No hay ningún ejecutable propio, así que el Control inteligente de aplicaciones de Windows 11 no bloquea el programa. El `.msi` es la excepción: va sin firmar, y el Control inteligente de aplicaciones puede bloquearlo.

Además, la compilación genera **`CentrosEducativos.msix`**, el paquete de la Microsoft Store. Va sin firmar porque lo firma Microsoft al publicarlo: se sube a Partner Center como explica [`TIENDA.md`](TIENDA.md).

## Compilación automática

[`.github/workflows/windows.yml`](../.github/workflows/windows.yml) lo hace todo en GitHub, en cada cambio y en cada versión:

1. Descarga de python.org los MSI oficiales `core`, `exe`, `lib` y `tcltk` de Python 3.13.17 (64 bits) y comprueba su SHA-256 y su firma.
2. Los extrae con `msiexec /a`.
3. Monta el paquete y el zip con [`montar_paquete.py`](montar_paquete.py):
   - Quita lo que el programa no usa: `python.exe`, IDLE, `ensurepip`, `turtledemo`, las bibliotecas de compilación de Tcl/Tk y algunos iconos. Así, los únicos ejecutables que quedan son los 37 firmados de Python.
   - Renombra `pythonw.exe` a `Centros educativos.exe` (la firma sigue siendo válida) y `LICENSE.txt` a `LICENSE-Python.txt`.
   - Instala openpyxl, et_xmlfile y xlrd con la versión y el hash fijados en [`requisitos.txt`](requisitos.txt).
   - Copia [`sitecustomize.py`](sitecustomize.py) en `Lib\site-packages`, el programa en `programa\` y [`LEEME.txt`](LEEME.txt) en la raíz.
   - En el zip, marca como ocultos los archivos de Python para que solo se vean «Centros educativos», «LEEME» y «programa».
4. Comprueba que todos los `.exe`, `.dll` y `.pyd` tienen una firma válida de la Python Software Foundation o de Microsoft.
5. Compila el `.msi` con WiX Toolset 5 a partir de [`instalador.wxs`](instalador.wxs). Para el `.msi`, el paquete lleva los `.py` ya compilados: así Python no escribe nada en la carpeta del programa y la desinstalación no deja restos.
6. Marca el `.msi` como instalable sin administrador y lo valida (ICE).
7. Lo instala, abre el programa, comprueba que arranca sin errores y lo desinstala.
8. Monta el paquete de la Microsoft Store con [`montar_msix.py`](montar_msix.py) a partir del paquete del `.msi`:
   - El ejecutable se llama `CentrosEducativos.exe`, sin espacios, y no lleva `LEEME.txt`.
   - Añade [`msix/AppxManifest.xml`](msix/AppxManifest.xml), con la identidad de Partner Center (las variables `STORE_*` del repositorio; sin ellas, una de prueba), y los iconos de [`msix/Assets`](msix/Assets).
   - Crea `resources.pri` con makepri y el `.msix` con makeappx, del Windows SDK.
9. Firma una copia del `.msix` con un certificado de prueba, la instala y abre el programa como desde el menú Inicio. Comprueba que arranca con la identidad del paquete, que guarda sus datos en la carpeta del paquete y que no crea el acceso directo del escritorio. Después lo desinstala.
10. Al publicar una versión, adjunta el zip y el `.msi` a la versión. El `.msix` queda en los artefactos de la ejecución, para subirlo a Partner Center. Los tres archivos están también en los artefactos de cada ejecución, sin comprimir.

## Versión de la Microsoft Store

Es el mismo programa, pero lo instala y lo actualiza la Tienda:

- Se instala en `C:\Program Files\WindowsApps`, en una carpeta de solo lectura que cambia con cada versión. Por eso el programa no crea el acceso directo del escritorio: se abre desde el menú Inicio.
- Lo que el programa guarda en `%LOCALAPPDATA%\CentrosEducativos` va a la carpeta de datos del paquete, `%LOCALAPPDATA%\Packages\<paquete>\LocalCache\Local\CentrosEducativos`. Windows la borra al desinstalarlo.
- Se actualiza desde la Tienda.

El programa sabe que es la versión de la Tienda porque Windows le da la identidad del paquete (`GetCurrentPackageFamilyName`).

## Estructura

```
CentrosEducativos/           ← con el .msi: %LOCALAPPDATA%\Programs\Centros educativos
├── Centros educativos.exe   ← pythonw.exe oficial con otro nombre
├── LEEME.txt
├── python313.dll, python3.dll, vcruntime140.dll, vcruntime140_1.dll, LICENSE-Python.txt
├── DLLs/  tcl/  Lib/        ← biblioteca estándar, Tkinter y Tcl/Tk
│   └── Lib/site-packages/   ← openpyxl, et_xmlfile, xlrd y sitecustomize.py
└── programa/                ← el programa, la lista, README.md y LICENSE.txt
```

## Compilar a mano

En Windows, con Python 3.13:

```
mkdir C:\py313
:: Descarga en C:\py313 los MSI core, exe, lib y tcltk de https://www.python.org/ftp/python/3.13.17/amd64/
for %m in (core exe lib tcltk) do msiexec /a C:\py313\%m.msi TARGETDIR=C:\py313\python /qn
py windows\montar_paquete.py C:\py313\python C:\paquete --zip CentrosEducativos.zip
```

Para el `.msi`, monta el paquete con `--pyc`, compílalo con WiX 5 y márcalo como instalable sin administrador, como en el paso «Marcar el .msi como instalable sin administrador» del workflow:

```
py windows\montar_paquete.py C:\py313\python C:\paquete-msi --pyc
dotnet tool install --global wix --version 5.0.2
wix build -arch x64 -d Version=1.2.0 -d Origen=C:\paquete-msi windows\instalador.wxs -o CentrosEducativos.msi
```

El paquete de la Tienda se monta a partir del paquete del `.msi`. Después, makepri y makeappx (del Windows SDK) crean `resources.pri` y el `.msix`, como en el paso «Crear el .msix de la Microsoft Store» del workflow:

```
py windows\montar_msix.py C:\paquete-msi C:\paquete-msix --version 1.2.0 --nombre ... --editor CN=... --editor-visible ...
```

En Linux se puede montar el zip extrayendo los MSI con `msiextract` (paquete msitools) en lugar de `msiexec /a`. El `.msi` y el `.msix` solo se pueden compilar en Windows.

## Cómo arranca

Cuando se abre `Centros educativos.exe` (en la versión de la Tienda, `CentrosEducativos.exe`) sin argumentos, Python carga `Lib\site-packages\sitecustomize.py`. Ese archivo comprueba el nombre del ejecutable y lanza `programa\centros_gui.pyw`. Si el programa falla al arrancar, el error se guarda en `programa\error.log` o, si esa carpeta es de solo lectura, en `%LOCALAPPDATA%\CentrosEducativos\error.log`.

Con cualquier otro nombre, o con argumentos, el Python del paquete funciona con normalidad: `sitecustomize.py` solo actúa en ese caso concreto.
