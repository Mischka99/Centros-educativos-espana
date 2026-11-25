# -*- coding: utf-8 -*-
"""
Genera los iconos PNG de la ventana (programa/img/<escala>/<nombre>.png) para cada escalado de pantalla de Windows.

Solo hace falta al preparar el programa (usa rsvg-convert y Pillow); el programa únicamente lee los PNG,
así no se añade ningún ejecutable ni librería nueva al paquete.

    python generar_imagenes.py [carpeta_destino]
"""

import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image

ESCALAS = {"100": 1.0, "125": 1.25, "150": 1.5, "175": 1.75, "200": 2.0, "250": 2.5, "300": 3.0}
ACENTO = "#C2410C"
COLORES = {
    "texto": "#1c1917", "suave": "#57534e", "tenue": "#78716c", "lat": "#a8a29e", "claro": "#fafaf9",
    "blanco": "#ffffff", "acento": ACENTO, "acento_osc": "#9a3412", "marcado": "#923109",
    "verde": "#15803d", "verde_osc": "#166534", "ambar": "#b45309", "ambar_osc": "#92400e",
    "rojo": "#b91c1c", "rojo_osc": "#991b1b", "apagado": "#a8a29e",
}
# trazos estilo «Lucide» (viewBox 24×24); (dibujo, grosor)
TRAZOS = {
    "buscar": ('<circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/>', 1.8),
    "cerrar": ('<path d="M6 6l12 12M18 6L6 18"/>', 2),
    "chevron": ('<path d="M6 9l6 6 6-6"/>', 2),
    "check": ('<path d="M5 12l5 5L20 7"/>', 2.2),
    "actualizar": ('<path d="M20 11a8 8 0 1 0-2.3 5.7"/><path d="M20 4v7h-7"/>', 1.8),
    "copiar": ('<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a1 1 0 0 1 1-1h10"/>', 1.8),
    "email": ('<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/>', 1.8),
    "telefono": ('<path d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2"/>',
                 1.8),
    "web": ('<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18"/>', 1.8),
    "mapa": ('<path d="M12 21s-7-6.2-7-11a7 7 0 0 1 14 0c0 4.8-7 11-7 11z"/><circle cx="12" cy="10" r="2.5"/>', 1.8),
    "codigo": ('<path d="M5 9h14M5 15h14M10 4L8 20M16 4l-2 16"/>', 1.8),
    "descargar": ('<path d="M12 4v11"/><path d="M7 10l5 5 5-5"/><path d="M5 20h14"/>', 2),
    "ok": ('<circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/>', 2),
    "alerta": ('<circle cx="12" cy="12" r="9"/><path d="M12 7.5v5.5"/><path d="M12 16.5h.01"/>', 2),
    "ampliar": ('<path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/>', 1.8),
    "externo": ('<path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/>', 1.8),
}
# (icono, color, tamaño en píxeles a escala 100 %)
ICONOS = [
    ("buscar", "suave", 18), ("buscar", "lat", 16), ("buscar", "suave", 16),
    ("cerrar", "tenue", 14), ("cerrar", "lat", 14), ("cerrar", "verde_osc", 14), ("cerrar", "ambar_osc", 14),
    ("cerrar", "rojo_osc", 14),
    ("chevron", "texto", 12), ("chevron", "acento_osc", 12), ("chevron", "tenue", 12),
    ("check", "marcado", 16), ("check", "acento_osc", 14), ("check", "acento", 16),
    ("actualizar", "claro", 16), ("actualizar", "apagado", 16), ("actualizar", "ambar", 18),
    ("copiar", "texto", 16), ("copiar", "texto", 15), ("copiar", "apagado", 16),
    ("email", "texto", 15),
    ("telefono", "suave", 18), ("telefono", "texto", 15),
    ("web", "suave", 18), ("mapa", "suave", 18), ("mapa", "texto", 15), ("codigo", "suave", 18),
    ("descargar", "blanco", 16), ("descargar", "texto", 16),
    ("ok", "verde", 18), ("alerta", "rojo", 18), ("alerta", "ambar_osc", 16),
    ("ampliar", "texto", 15), ("ampliar", "apagado", 15), ("externo", "texto", 15),
]
# mapa: chincheta, puntos de los otros centros (por tipo; con borde oscuro si están marcados) y esquinas redondeadas
PUNTO = '<circle cx="8" cy="8" r="7.2" fill="none" stroke="#1c1917" stroke-opacity="0.22" stroke-width="1"/>' \
        '<circle cx="8" cy="8" r="5.6" fill="{c}" stroke="#ffffff" stroke-width="2"/>'
PUNTO_M = '<circle cx="9" cy="9" r="6.6" fill="{c}" stroke="#1c1917" stroke-width="2.6"/>'
ESQUINA = ('<path d="M0 0H10A10 10 0 0 0 0 10Z" fill="#ffffff"/>'
           '<path d="M10 0.5A9.5 9.5 0 0 0 0.5 10" fill="none" stroke="#e7e5e4" stroke-width="1"/>')
MAPA = {  # nombre: (ancho, alto, svg)
    "pin": (28, 36, '<path d="M14 34.5S2.5 24 2.5 14a11.5 11.5 0 0 1 23 0c0 10-11.5 20.5-11.5 20.5z" fill="none" '
                    'stroke="#1c1917" stroke-opacity="0.3" stroke-width="3.6"/>'
                    '<path d="M14 34.5S2.5 24 2.5 14a11.5 11.5 0 0 1 23 0c0 10-11.5 20.5-11.5 20.5z" '
                    f'fill="{ACENTO}" stroke="#ffffff" stroke-width="2"/><circle cx="14" cy="14" r="4.5" fill="#ffffff"/>'),
    "punto-ip": (16, 16, PUNTO.format(c=ACENTO)), "punto-sec": (16, 16, PUNTO.format(c="#1d4ed8")),
    "punto-otros": (16, 16, PUNTO.format(c="#a8a29e")),
    "punto-ip-m": (18, 18, PUNTO_M.format(c=ACENTO)), "punto-sec-m": (18, 18, PUNTO_M.format(c="#1d4ed8")),
    "punto-otros-m": (18, 18, PUNTO_M.format(c="#a8a29e")),
    "mapa-esq-ii": (10, 10, ESQUINA), "mapa-esq-id": (10, 10, f'<g transform="rotate(90 5 5)">{ESQUINA}</g>'),
    "mapa-esq-bd": (10, 10, f'<g transform="rotate(180 5 5)">{ESQUINA}</g>'),
    "mapa-esq-bi": (10, 10, f'<g transform="rotate(270 5 5)">{ESQUINA}</g>'),
}
# casillas de 16 px (viewBox 16×16)
CASILLAS = {
    "casilla-claro-no": '<rect x="1.75" y="1.75" width="12.5" height="12.5" rx="3.5" fill="#ffffff" stroke="#a8a29e" '
                        'stroke-width="1.5"/>',
    "casilla-claro-hover": '<rect x="1.75" y="1.75" width="12.5" height="12.5" rx="3.5" fill="#ffffff" stroke="#78716c" '
                           'stroke-width="1.5"/>',
    "casilla-claro-si": f'<rect x="1" y="1" width="14" height="14" rx="4" fill="{ACENTO}"/>'
                        '<path d="M4.4 8.3l2.5 2.5 4.8-5.2" fill="none" stroke="#ffffff" stroke-width="1.9" '
                        'stroke-linecap="round" stroke-linejoin="round"/>',
    "casilla-claro-mixto": f'<rect x="1" y="1" width="14" height="14" rx="4" fill="{ACENTO}"/>'
                           '<path d="M4.5 8h7" fill="none" stroke="#ffffff" stroke-width="1.9" stroke-linecap="round"/>',
    "casilla-oscuro-no": '<rect x="1.75" y="1.75" width="12.5" height="12.5" rx="3.5" fill="none" stroke="#8a827c" '
                         'stroke-width="1.5"/>',
    "casilla-oscuro-si": f'<rect x="1" y="1" width="14" height="14" rx="4" fill="{ACENTO}"/>'
                         '<path d="M4.4 8.3l2.5 2.5 4.8-5.2" fill="none" stroke="#ffffff" stroke-width="1.9" '
                         'stroke-linecap="round" stroke-linejoin="round"/>',
}


def svg_icono(dibujo, grosor, color):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" '
            f'stroke="{color}" stroke-width="{grosor}" stroke-linecap="round" stroke-linejoin="round">{dibujo}</svg>')


def rasterizar(svg, lado, destino, tmp, alto=None):
    ruta = os.path.join(tmp, "x.svg")
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(svg)
    subprocess.run(["rsvg-convert", "-w", str(lado), "-h", str(alto or lado), "-o", destino, ruta], check=True)


def main(destino):
    icono_app = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icono.png")
    if os.path.isdir(destino):
        shutil.rmtree(destino)
    with tempfile.TemporaryDirectory() as tmp:
        for carpeta, f in ESCALAS.items():
            d = os.path.join(destino, carpeta)
            os.makedirs(d)
            for nombre, color, tam in ICONOS:
                dibujo, grosor = TRAZOS[nombre]
                rasterizar(svg_icono(dibujo, grosor, COLORES[color]), round(tam * f),
                           os.path.join(d, f"{nombre}-{color}-{tam}.png"), tmp)
            for nombre, dibujo in CASILLAS.items():
                svg = f'<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16">{dibujo}</svg>'
                rasterizar(svg, round(16 * f), os.path.join(d, nombre + ".png"), tmp)
            for nombre, (an, al, dibujo) in MAPA.items():
                svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{an}" height="{al}" '
                       f'viewBox="0 0 {an} {al}">{dibujo}</svg>')
                rasterizar(svg, round(an * f), os.path.join(d, nombre + ".png"), tmp, round(al * f))
            with Image.open(icono_app) as im:
                lado = round(34 * f)
                im.convert("RGBA").resize((lado, lado), Image.LANCZOS).save(os.path.join(d, "marca.png"), optimize=True)
    n = sum(len(fs) for _, _, fs in os.walk(destino))
    print(f"{n} imágenes en {destino}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "img"))
