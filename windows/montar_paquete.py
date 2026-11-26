# -*- coding: utf-8 -*-
"""Monta el paquete de Windows de «Centros educativos»: el Python oficial de python.org más el programa.

Lo usa la compilación automática (.github/workflows/windows.yml), y se puede ejecutar a mano en cualquier sistema:

    python windows/montar_paquete.py PYTHON SALIDA [--zip ARCHIVO.zip] [--pyc]

PYTHON  Carpeta con el contenido de los MSI oficiales core, exe, lib y tcltk de Python 3.13 para Windows (64 bits),
        extraídos con «msiexec /a» (Windows) o «msiextract» (Linux).
SALIDA  Carpeta que se crea con el paquete: «Centros educativos.exe», LEEME.txt, programa\\ y el Python.
--zip   Además, comprime la carpeta (CentrosEducativos.zip). Los archivos de Python van ocultos para que, al
        descomprimir, solo se vean «Centros educativos», «LEEME» y «programa».
--pyc   Compila de antemano todos los .py (para el instalador .msi). Python ya no escribe nada en la carpeta del
        programa, y la desinstalación la deja vacía.

Las dependencias (openpyxl, et_xmlfile y xlrd, Python puro) se instalan con pip desde PyPI, con la versión y el
hash fijados en windows/requisitos.txt.
"""
import argparse
import compileall
import os
import py_compile
import shutil
import subprocess
import sys
import zipfile

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # el repositorio
VENTANAS = os.path.join(RAIZ, "windows")

# Lo que el programa no usa. Así, los únicos ejecutables que quedan son los firmados de Python.
QUITAR_CARPETAS = ["libs", "Lib/ensurepip", "Lib/idlelib", "Lib/turtledemo", "tcl/nmake"]
QUITAR_ARCHIVOS = ["python.exe", "NEWS.txt", "DLLs/py.ico", "DLLs/pyc.ico", "DLLs/pyd.ico"]
QUITAR_EXTENSIONES = {"tcl": (".lib", ".sh"), "": (".msi",)}  # solo en esa carpeta, sin entrar en subcarpetas
RENOMBRAR = {"pythonw.exe": "Centros educativos.exe", "LICENSE.txt": "LICENSE-Python.txt"}

# Programa: lo que se copia del repositorio a programa\.
PROGRAMA = ["centros_educativos.py", "centros_gui.pyw", "interfaz.py", "mapa.py", "icono.ico", "icono.png", "README.md"]
CARPETAS_PROGRAMA = ["img", "datos"]

# En el zip, ocultos (atributo de Windows) para que la carpeta se vea limpia.
OCULTOS = {"python313.dll", "python3.dll", "vcruntime140.dll", "vcruntime140_1.dll", "LICENSE-Python.txt",
           "DLLs", "Lib", "tcl"}

IMPRESCINDIBLES = ["Centros educativos.exe", "python313.dll", "LEEME.txt", "LICENSE-Python.txt",
                   "Lib/os.py", "Lib/tkinter/__init__.py", "DLLs/_tkinter.pyd", "DLLs/_ssl.pyd",
                   "tcl/tcl8.6/init.tcl", "tcl/tk8.6/tk.tcl",
                   "Lib/site-packages/sitecustomize.py", "Lib/site-packages/openpyxl/__init__.py",
                   "Lib/site-packages/xlrd/__init__.py", "Lib/site-packages/et_xmlfile/__init__.py",
                   "programa/centros_gui.pyw", "programa/centros_educativos.py", "programa/interfaz.py",
                   "programa/mapa.py", "programa/LICENSE.txt", "programa/img/100/marca.png",
                   "programa/datos/centros_inicial.json.gz"]


def fallo(texto):
    sys.exit(f"ERROR: {texto}")


def ruta(base, rel):
    return os.path.join(base, *rel.split("/"))


def montar(python, salida, pyc=False):
    if not os.path.isfile(ruta(python, "pythonw.exe")) or not os.path.isfile(ruta(python, "python313.dll")):
        fallo(f"{python} no parece el Python 3.13 extraído de los MSI (falta pythonw.exe o python313.dll)")
    if os.path.exists(salida):
        fallo(f"{salida} ya existe")
    shutil.copytree(python, salida)

    for rel in QUITAR_CARPETAS:
        if not os.path.isdir(ruta(salida, rel)):
            fallo(f"no está {rel}: ¿ha cambiado la versión de Python? Revisa QUITAR_CARPETAS")
        shutil.rmtree(ruta(salida, rel))
    for rel in QUITAR_ARCHIVOS:
        if not os.path.isfile(ruta(salida, rel)):
            fallo(f"no está {rel}: ¿ha cambiado la versión de Python? Revisa QUITAR_ARCHIVOS")
        os.remove(ruta(salida, rel))
    for carpeta, extensiones in QUITAR_EXTENSIONES.items():
        d = ruta(salida, carpeta) if carpeta else salida
        for nombre in os.listdir(d):
            if nombre.lower().endswith(extensiones) and os.path.isfile(os.path.join(d, nombre)):
                os.remove(os.path.join(d, nombre))
    for antes, despues in RENOMBRAR.items():
        os.rename(ruta(salida, antes), ruta(salida, despues))

    # Dependencias (Python puro), con versión y hash fijados.
    site = ruta(salida, "Lib/site-packages")
    subprocess.run([sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "--no-input",
                    "--target", site, "--no-compile", "--no-deps", "--only-binary=:all:", "--require-hashes",
                    "-r", os.path.join(VENTANAS, "requisitos.txt")], check=True)
    shutil.rmtree(os.path.join(site, "bin"), ignore_errors=True)  # script de consola de xlrd: no se usa
    shutil.copy2(os.path.join(VENTANAS, "sitecustomize.py"), site)

    programa = ruta(salida, "programa")
    os.makedirs(programa)
    for nombre in PROGRAMA:
        shutil.copy2(os.path.join(RAIZ, nombre), programa)
    for nombre in CARPETAS_PROGRAMA:
        shutil.copytree(os.path.join(RAIZ, nombre), os.path.join(programa, nombre))
    shutil.copy2(os.path.join(RAIZ, "LICENSE"), os.path.join(programa, "LICENSE.txt"))
    shutil.copy2(os.path.join(VENTANAS, "LEEME.txt"), salida)

    limpiar(salida)
    if pyc:
        compilar(salida)
    verificar(salida, pyc)


def limpiar(salida):
    for d, carpetas, archivos in os.walk(salida):
        for c in list(carpetas):
            if c == "__pycache__":
                shutil.rmtree(os.path.join(d, c))
                carpetas.remove(c)
        for a in archivos:
            if a.endswith((".pyc", ".log")):
                os.remove(os.path.join(d, a))


def compilar(salida):
    """Compila todos los .py con «unchecked-hash»: Python usa el .pyc sin comprobar la fuente y no reescribe nada.
    Hay que hacerlo con Python 3.13, la misma versión que la del paquete."""
    if sys.version_info[:2] != (3, 13):
        fallo(f"--pyc necesita Python 3.13 (este es {sys.version.split()[0]})")
    modo = py_compile.PycInvalidationMode.UNCHECKED_HASH
    for rel in ("Lib", "programa"):
        if not compileall.compile_dir(ruta(salida, rel), quiet=1, workers=0, invalidation_mode=modo):
            fallo(f"no se ha podido compilar todo {rel}")


def es_ejecutable(archivo):
    with open(archivo, "rb") as f:
        return f.read(2) == b"MZ"


def verificar(salida, pyc):
    for rel in IMPRESCINDIBLES:
        if not os.path.isfile(ruta(salida, rel)):
            fallo(f"falta {rel}")
    if os.path.exists(ruta(salida, "pythonw.exe")) or os.path.exists(ruta(salida, "python.exe")):
        fallo("quedan python.exe o pythonw.exe")
    ejecutables = []
    for d, _, archivos in os.walk(salida):
        for a in archivos:
            p = os.path.join(d, a)
            rel = os.path.relpath(p, salida).replace(os.sep, "/")
            if a.endswith(".log") or (a.endswith(".pyc") and not pyc):
                fallo(f"no debería estar {rel}")
            if es_ejecutable(p):
                ejecutables.append(rel)
                if rel.startswith(("programa/", "Lib/site-packages/")):
                    fallo(f"{rel} es un ejecutable que no es de Python: no puede ir en el paquete")
    print(f"{len(ejecutables)} ejecutables (.exe, .dll, .pyd), todos del Python oficial")


def comprimir(salida, archivo_zip):
    """Zip con la carpeta «CentrosEducativos» y los archivos de Python ocultos (atributo de Windows)."""
    if os.path.exists(archivo_zip):
        os.remove(archivo_zip)
    raiz_zip = "CentrosEducativos"
    n = 0
    with zipfile.ZipFile(archivo_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for d, carpetas, archivos in os.walk(salida):
            carpetas.sort()
            archivos.sort()
            rel_d = os.path.relpath(d, salida).replace(os.sep, "/")
            primero = rel_d.split("/")[0]
            if rel_d != ".":
                zi = zipfile.ZipInfo(f"{raiz_zip}/{rel_d}/")
                zi.create_system = 0
                zi.external_attr = 0x10 | (0x02 if primero in OCULTOS else 0)
                z.writestr(zi, b"")
            for a in archivos:
                p = os.path.join(d, a)
                rel = os.path.relpath(p, salida).replace(os.sep, "/")
                zi = zipfile.ZipInfo.from_file(p, f"{raiz_zip}/{rel}")
                zi.create_system = 0
                zi.external_attr = 0x20 | (0x02 if rel.split("/")[0] in OCULTOS else 0)
                zi.compress_type = zipfile.ZIP_DEFLATED
                with open(p, "rb") as f:
                    z.writestr(zi, f.read(), compresslevel=9)
                n += 1
    with zipfile.ZipFile(archivo_zip) as z:
        if z.testzip() is not None:
            fallo("el zip está dañado")
        visibles = sorted({x.split("/")[1] for x in z.namelist()
                           if x.count("/") >= 1 and x.split("/")[1] and not z.getinfo(x).external_attr & 0x02})
    print(f"{archivo_zip}: {n} archivos, {os.path.getsize(archivo_zip) / 1e6:.1f} MB; se ven: {', '.join(visibles)}")


def main():
    p = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    p.add_argument("python", help="carpeta con los MSI oficiales de Python extraídos")
    p.add_argument("salida", help="carpeta que se crea con el paquete")
    p.add_argument("--zip", help="crea también este zip")
    p.add_argument("--pyc", action="store_true", help="compila los .py de antemano (para el .msi)")
    a = p.parse_args()
    montar(a.python, a.salida, a.pyc)
    if a.zip:
        comprimir(a.salida, a.zip)
    print("OK")


if __name__ == "__main__":
    main()
