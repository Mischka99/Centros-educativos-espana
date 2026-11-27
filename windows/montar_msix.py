# -*- coding: utf-8 -*-
"""Monta la carpeta del paquete MSIX de «Centros educativos» para la Microsoft Store.

Lo usa la compilación automática (.github/workflows/windows.yml):

    python windows/montar_msix.py PAQUETE SALIDA --version X.Y.Z [--nombre N --editor CN=... --editor-visible E]
                                  [--nombre-visible NOMBRE]

PAQUETE  El paquete montado con «windows/montar_paquete.py ... --pyc», el mismo del instalador .msi.
SALIDA   Carpeta que se crea: el paquete con CentrosEducativos.exe, AppxManifest.xml y Assets\\.
Después, makepri y makeappx (del Windows SDK) crean resources.pri y el .msix.

La identidad (--nombre, --editor y --editor-visible, las tres o ninguna) la da Partner Center al reservar el nombre de
la aplicación (windows/TIENDA.md). Sin ella se usa una de prueba, que sirve para probar el paquete pero no para subirlo
a la Tienda. --nombre-visible es el nombre reservado en la Tienda; por defecto, «Centros educativos».

    python windows/montar_msix.py --iconos

vuelve a generar los iconos de windows/msix/Assets a partir de icono.png (necesita Pillow).
"""
import argparse
import os
import re
import shutil
import sys
from xml.sax.saxutils import escape

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # el repositorio
MSIX = os.path.join(RAIZ, "windows", "msix")
EJECUTABLE = "CentrosEducativos.exe"   # sin espacios: así lo prefiere el sistema de paquetes de Windows
PRUEBA = {"nombre": "CentrosEducativos.Prueba", "editor": "CN=Centros educativos (prueba)",
          "editor_visible": "Centros educativos (prueba)"}
NOMBRE_VISIBLE = "Centros educativos"

# Iconos del paquete: nombre → [(calificador, lado en píxeles, parte del cuadro que ocupa el icono)].
# Las variantes «scale» son para cada escalado de pantalla; las «targetsize», para la barra de tareas y las listas.
ICONOS = {
    "Square44x44Logo": [(f".scale-{e}", round(44 * e / 100), 1) for e in (100, 125, 150, 200, 400)]
                       + [(f".targetsize-{n}", n, 1) for n in (16, 24, 32, 48, 256)]
                       + [(f".targetsize-{n}_altform-unplated", n, 1) for n in (16, 24, 32, 48, 256)],
    "Square150x150Logo": [(f".scale-{e}", round(150 * e / 100), 0.66) for e in (100, 125, 150, 200)],
    "StoreLogo": [(f".scale-{e}", round(50 * e / 100), 1) for e in (100, 125, 150, 200, 400)],
}


def fallo(texto):
    sys.exit(f"ERROR: {texto}")


def generar_iconos():
    from PIL import Image  # solo hace falta para regenerar los iconos
    origen = Image.open(os.path.join(RAIZ, "icono.png")).convert("RGBA")
    destino = os.path.join(MSIX, "Assets")
    shutil.rmtree(destino, ignore_errors=True)
    os.makedirs(destino)
    for nombre, variantes in ICONOS.items():
        for calificador, lado, parte in variantes:
            tam = round(lado * parte)
            if tam > origen.width:
                fallo(f"icono.png ({origen.width} px) es pequeño para {nombre}{calificador} ({tam} px)")
            icono = origen.convert("RGBa").resize((tam, tam), Image.LANCZOS).convert("RGBA")  # alfa premultiplicado
            lienzo = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
            lienzo.alpha_composite(icono, ((lado - tam) // 2, (lado - tam) // 2))
            lienzo.save(os.path.join(destino, f"{nombre}{calificador}.png"), optimize=True)
    print(f"{sum(len(v) for v in ICONOS.values())} iconos en {destino}")


def version_msix(texto):
    partes = texto.split(".")
    if len(partes) != 3 or not all(p.isdigit() and int(p) <= 65535 for p in partes):
        fallo(f"la versión debe ser X.Y.Z (es «{texto}»)")
    return ".".join(str(int(p)) for p in partes) + ".0"  # la Tienda exige que el cuarto número sea 0


def identidad(a):
    recibidos = {"nombre": (a.nombre or "").strip(), "editor": (a.editor or "").strip(),
                 "editor_visible": (a.editor_visible or "").strip()}
    faltan = [f"--{clave.replace('_', '-')}" for clave, valor in recibidos.items() if not valor]
    if len(faltan) == len(recibidos):
        datos = dict(PRUEBA)
    elif faltan:
        fallo(f"{'faltan' if len(faltan) > 1 else 'falta'} {' y '.join(faltan)}: la identidad de Partner Center "
              "son tres valores y hacen falta todos")
    else:
        datos = recibidos
    if not re.fullmatch(r"[A-Za-z0-9.\-]{3,50}", datos["nombre"]):
        fallo(f"nombre de paquete no válido: «{datos['nombre']}» (Package/Identity/Name en Partner Center)")
    if not datos["editor"].startswith("CN="):
        fallo(f"editor no válido: «{datos['editor']}» (Package/Identity/Publisher, empieza por CN=)")
    datos["nombre_visible"] = (a.nombre_visible or "").strip() or NOMBRE_VISIBLE
    if len(datos["nombre_visible"]) > 256:
        fallo("el nombre visible es demasiado largo")
    return datos


def montar(paquete, salida, version, datos):
    if not os.path.isfile(os.path.join(paquete, "Centros educativos.exe")) \
            or not os.path.isfile(os.path.join(paquete, "programa", "centros_gui.pyw")):
        fallo(f"{paquete} no parece el paquete montado con montar_paquete.py")
    if os.path.exists(salida):
        fallo(f"{salida} ya existe")
    shutil.copytree(paquete, salida)
    os.rename(os.path.join(salida, "Centros educativos.exe"), os.path.join(salida, EJECUTABLE))
    os.remove(os.path.join(salida, "LEEME.txt"))  # guía del zip: no sirve en la versión de la Tienda

    with open(os.path.join(MSIX, "AppxManifest.xml"), encoding="utf-8") as f:
        manifiesto = f.read()
    valores = {"NOMBRE": datos["nombre"], "EDITOR": datos["editor"], "EDITOR_VISIBLE": datos["editor_visible"],
               "NOMBRE_VISIBLE": datos["nombre_visible"], "VERSION": version}
    for clave, valor in valores.items():
        manifiesto = manifiesto.replace("{{" + clave + "}}", escape(valor, {'"': "&quot;"}))
    if "{{" in manifiesto:
        fallo("quedan valores sin rellenar en AppxManifest.xml")
    with open(os.path.join(salida, "AppxManifest.xml"), "w", encoding="utf-8") as f:
        f.write(manifiesto)
    shutil.copytree(os.path.join(MSIX, "Assets"), os.path.join(salida, "Assets"))

    prueba = datos["nombre"] == PRUEBA["nombre"]
    print(f"{salida}: «{datos['nombre_visible']}» versión {version}, identidad {datos['nombre']} · {datos['editor']}"
          f" · {datos['editor_visible']}" + (" (de prueba: no sirve para subirlo a la Tienda)" if prueba else ""))


def main():
    p = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    p.add_argument("paquete", nargs="?", help="paquete montado con montar_paquete.py --pyc")
    p.add_argument("salida", nargs="?", help="carpeta que se crea con el paquete MSIX")
    p.add_argument("--version", help="versión X.Y.Z")
    p.add_argument("--nombre", help="Package/Identity/Name de Partner Center")
    p.add_argument("--editor", help="Package/Identity/Publisher de Partner Center (CN=...)")
    p.add_argument("--editor-visible", help="Package/Properties/PublisherDisplayName de Partner Center")
    p.add_argument("--nombre-visible", help=f"nombre reservado en la Tienda (por defecto, «{NOMBRE_VISIBLE}»)")
    p.add_argument("--iconos", action="store_true", help="regenera los iconos a partir de icono.png")
    a = p.parse_args()
    if a.iconos:
        generar_iconos()
        return
    if not (a.paquete and a.salida and a.version):
        p.error("hacen falta PAQUETE, SALIDA y --version")
    montar(a.paquete, a.salida, version_msix(a.version), identidad(a))


if __name__ == "__main__":
    main()
