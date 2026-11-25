# -*- coding: utf-8 -*-
"""
Piezas de la ventana de «Centros educativos» (diseño F): colores, escalado de pantalla, formas redondeadas,
iconos, campos de búsqueda, listas con casillas (comunidades y desplegables), tabla de centros, desplegables,
panel desplazable y texto con negritas.

Todo es Tkinter puro: las esquinas redondeadas se calculan al arrancar (suavizadas contra el color de fondo, sin
transparencias) y los iconos son PNG de la carpeta img/. No hace falta ninguna librería ni ejecutable adicional.
"""

import math
import os
import sys
import time
import unicodedata
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

ESCALAS = {"100": 1.0, "125": 1.25, "150": 1.5, "175": 1.75, "200": 2.0, "250": 2.5, "300": 3.0}


def _rgb(c):
    c = c.lstrip("#")
    return int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)


def _hex(r, g, b):
    return "#%02x%02x%02x" % (int(round(r)), int(round(g)), int(round(b)))


def mezcla(a, b, t):
    """Color entre a (t=0) y b (t=1)."""
    (r1, g1, b1), (r2, g2, b2) = _rgb(a), _rgb(b)
    return _hex(r1 + (r2 - r1) * t, g1 + (g2 - g1) * t, b1 + (b2 - b1) * t)


ACENTO = "#C2410C"
C = {
    "fondo": "#fafaf9", "blanco": "#ffffff", "texto": "#1c1917", "suave": "#57534e", "sec": "#44403c",
    "tenue": "#78716c", "apagado": "#a8a29e", "borde": "#e7e5e4", "borde2": "#d6d3d1", "gris": "#f5f5f4",
    "gris2": "#e7e5e4",
    "lat": "#1c1917", "lat2": "#292524", "lat_borde": "#44403c", "lat_borde2": "#57534e", "lat_texto": "#e7e5e4",
    "lat_tenue": "#a8a29e", "lat_claro": "#fafaf9", "lat_hover": "#231f1d", "lat_pulsado": "#3a3532",
    "acento": ACENTO, "acento_hover": "#9a3412", "acento_pulsado": "#7c2d12",
    "acento_suave": mezcla(ACENTO, "#ffffff", 0.9), "acento_suave2": mezcla(ACENTO, "#ffffff", 0.82),
    "acento_borde": mezcla(ACENTO, "#ffffff", 0.6), "acento_texto": mezcla(ACENTO, "#000000", 0.25),
    "acento_claro": mezcla(ACENTO, "#ffffff", 0.45), "acento_apagado": mezcla(ACENTO, "#ffffff", 0.55),
    "fila_email": "#fffbeb", "fila_nuevo": "#f0fdf4", "fila_hover": "#f7f6f5",
}
AVISOS = {  # estado: (fondo, borde, texto, borde del botón, fondo del botón al pasar)
    "ok": ("#f0fdf4", "#bbf7d0", "#14532d", "#86efac", "#dcfce7"),
    "act": ("#fffbeb", "#fde68a", "#78350f", "#fcd34d", "#fef3c7"),
    "error": ("#fef2f2", "#fecaca", "#7f1d1d", "#fca5a5", "#fee2e2"),
}


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).lower()
    return "".join(c for c in s if not unicodedata.combining(c)).strip()


def espaciado(texto):
    """Mayúsculas con un poco de espacio entre letras (Tk no tiene letter-spacing)."""
    return " ".join(texto)


# ---------------------------------------------------------------------------------------------------------------
# escalado, fuentes, iconos y formas
# ---------------------------------------------------------------------------------------------------------------
class Escala:
    """Factor de escalado de Windows (100 % = 1.0) para convertir los píxeles del diseño."""

    def __init__(self, root):
        try:
            f = float(root.tk.call("tk", "scaling")) * 72 / 96
        except (tk.TclError, ValueError):
            f = 1.0
        self.f = min(max(f, 1.0), 4.0)
        self.carpeta = min(ESCALAS, key=lambda k: abs(ESCALAS[k] - self.f))

    def px(self, n):
        return int(round(n * self.f))


class Fuentes:
    def __init__(self, root):
        familias = set(tkfont.families(root))
        base = next((f for f in ("Segoe UI", "Inter", "Fira Sans", "Noto Sans", "DejaVu Sans") if f in familias),
                    tkfont.nametofont("TkDefaultFont").actual("family"))
        semi = next((f for f in (base + " Semibold", base + " SemiBold") if f in familias), None)

        def fuente(tam, negrita=False, subrayada=False):
            if negrita and semi:
                return tkfont.Font(root, family=semi, size=tam, underline=subrayada)
            return tkfont.Font(root, family=base, size=tam, weight="bold" if negrita else "normal",
                               underline=subrayada)

        self.normal = fuente(10)
        self.semi = fuente(10, True)
        self.peque = fuente(9)
        self.peque_semi = fuente(9, True)
        self.peque_sub = fuente(9, subrayada=True)
        self.mini = fuente(8)
        self.tipo = fuente(8, True)
        self.caps = fuente(8, True)
        self.titulo = fuente(17, True)
        self.kpi = fuente(16, True)
        self.kpi_extra = fuente(10)
        self.nombre = fuente(15, True)
        self.email = fuente(11, True)
        self.marca = fuente(11, True)
        self.mono = tkfont.Font(root, family="Consolas" if sys.platform == "win32" else "DejaVu Sans Mono", size=9)

    @staticmethod
    def alto(fuente):
        return fuente.metrics("linespace")


class Iconos:
    """Carga (una vez) los PNG de img/<escala>/; si falta alguno, se usa una imagen vacía."""

    def __init__(self, root, carpeta, escala):
        self.root = root
        self.dir = os.path.join(carpeta, escala.carpeta)
        self.cache = {}

    def __call__(self, nombre):
        if nombre not in self.cache:
            try:
                self.cache[nombre] = tk.PhotoImage(master=self.root, file=os.path.join(self.dir, nombre + ".png"))
            except (tk.TclError, OSError):  # sin el PNG: imagen vacía (el botón sale igual, sin icono)
                self.cache[nombre] = tk.PhotoImage(master=self.root, width=1, height=1)
        return self.cache[nombre]


def redondeado(root, ancho, alto, radio, relleno, fondo, borde=None, grosor=1.0):
    """PhotoImage opaca: rectángulo redondeado con el contorno suavizado y mezclado con el color de fondo."""
    radio = max(0.0, min(float(radio), ancho / 2, alto / 2))
    fondo_c, relleno_c = _rgb(fondo), _rgb(relleno)
    borde_c = _rgb(borde) if borde else relleno_c
    g = grosor if borde else 0.0
    mx, my = ancho / 2, alto / 2
    cache = {}

    def color(x, y):
        qx = abs(x + 0.5 - mx) - (mx - radio)
        qy = abs(y + 0.5 - my) - (my - radio)
        d = math.hypot(max(qx, 0.0), max(qy, 0.0)) + min(max(qx, qy), 0.0) - radio
        fuera = min(max(0.5 - d, 0.0), 1.0)
        dentro = min(max(0.5 - (d + g), 0.0), 1.0) if g else fuera
        clave = (round(fuera, 3), round(dentro, 3))
        if clave not in cache:
            cache[clave] = _hex(*(fc * (1 - fuera) + bc * (fuera - dentro) + rc * dentro
                                  for fc, bc, rc in zip(fondo_c, borde_c, relleno_c)))
        return cache[clave]

    lado_x = min(ancho // 2 + 1, int(radio + g) + 3)
    lado_y = min(alto // 2 + 1, int(radio + g) + 3)

    def fila(y):
        izq = [color(x, y) for x in range(lado_x)]
        medio = ancho - 2 * lado_x
        if medio >= 0:
            return izq + [color(ancho // 2, y)] * medio + izq[::-1]
        return [color(x, y) for x in range(ancho)]

    arriba = [fila(y) for y in range(lado_y)]
    medio = alto - 2 * lado_y
    filas = arriba + ([fila(alto // 2)] * medio if medio >= 0 else []) + arriba[::-1]
    if medio < 0:
        filas = [fila(y) for y in range(alto)]
    img = tk.PhotoImage(master=root, width=ancho, height=alto)
    img.put(" ".join("{" + " ".join(f) + "}" for f in filas))
    return img


def esquina(root, radio, fondo, borde, relleno, cuadrante):
    """Esquina (radio × radio) de un borde redondeado de 1 px: «ii», «id», «bi», «bd» (arriba/abajo, izq./dcha.)."""
    lado = radio * 2
    img = redondeado(root, lado, lado, radio, relleno, fondo, borde)
    trozo = tk.PhotoImage(master=root, width=radio, height=radio)
    x = 0 if cuadrante[1] == "i" else radio
    y = 0 if cuadrante[0] == "i" else radio
    trozo.tk.call(trozo, "copy", img, "-from", x, y, x + radio, y + radio)
    return trozo


class Formas:
    """Imágenes de 9 trozos para los estilos ttk (botones, chips, campos, tarjetas)."""

    def __init__(self, root):
        self.root = root
        self.cache = {}

    def nueve(self, radio, relleno, fondo, borde=None, grosor=1.0):
        """ttk repite (no estira) el centro y los lados de la imagen: con un centro amplio, pocas repeticiones
        y la ventana se redibuja al instante aunque la tarjeta sea grande."""
        clave = (radio, relleno, fondo, borde, grosor)
        if clave not in self.cache:
            b = radio + 1
            lado = 2 * b + 64
            self.cache[clave] = (redondeado(self.root, lado, lado, radio, relleno, fondo, borde, grosor), b)
        return self.cache[clave]

    def rect(self, ancho, alto, radio, relleno, fondo, borde=None):
        clave = ("r", ancho, alto, radio, relleno, fondo, borde)
        if clave not in self.cache:
            self.cache[clave] = redondeado(self.root, ancho, alto, radio, relleno, fondo, borde)
        return self.cache[clave]


class Estilos:
    """Estilos ttk del diseño. Cada estilo sabe sobre qué color se dibuja, para suavizar sus esquinas."""

    def __init__(self, root, escala, fuentes):
        self.root, self.e, self.fu = root, escala, fuentes
        self.st = ttk.Style(root)
        try:
            self.st.theme_use("clam")
        except tk.TclError:
            pass
        self.formas = Formas(root)
        self._n = 0
        st = self.st
        st.configure(".", font=fuentes.normal, background=C["fondo"], foreground=C["texto"],
                     selectbackground=C["acento_suave2"], selectforeground=C["texto"], troughcolor=C["fondo"],
                     focuscolor=C["acento"], bordercolor=C["borde2"], lightcolor=C["fondo"], darkcolor=C["fondo"])

    def _elemento(self, nombre, imagenes, border):
        normal = imagenes.pop("normal")
        estados = [(estado, img) for estado, img in imagenes.items()]
        # tamaño mínimo = las esquinas (si no, ttk pediría el tamaño de toda la imagen)
        self.st.element_create(nombre, "image", normal, *estados, border=border, padding=0, sticky="nsew",
                               width=2 * border, height=2 * border)

    def boton(self, nombre, fondo, alto, radio, colores, fuente, fg, padx=12, fg_estados=None, anchor="center"):
        """colores: {'normal': (relleno, borde), 'active': ..., 'pressed': ..., 'disabled': ...}."""
        r = min(self.e.px(radio), self.e.px(alto) // 2 - 1)
        imgs, b = {}, 0
        for estado in ("disabled", "pressed", "active", "normal"):
            if estado in colores:
                relleno, borde = colores[estado]
                img, b = self.formas.nueve(r, relleno, fondo, borde)
                imgs[estado] = img
        elem = f"{nombre}.borde"
        self._elemento(elem, imgs, b)
        self.st.layout(f"{nombre}.TButton", [(elem, {"sticky": "nsew", "children": [
            ("Button.padding", {"sticky": "nsew", "children": [("Button.label", {"sticky": "nsew"})]})]})])
        pady = max(0, (self.e.px(alto) - Fuentes.alto(fuente)) // 2)
        self.st.configure(f"{nombre}.TButton", font=fuente, foreground=fg, background=fondo, anchor=anchor,
                          padding=(max(self.e.px(padx), r), pady), relief="flat", borderwidth=0, width=0)
        if fg_estados:
            self.st.map(f"{nombre}.TButton", foreground=fg_estados)
        return f"{nombre}.TButton"

    def enlace(self, nombre, fondo, fuente, fuente_hover, fg, fg_hover, fg_off=None):
        self.st.layout(f"{nombre}.TButton", [("Button.padding", {"sticky": "nsew", "children": [
            ("Button.label", {"sticky": "nsew"})]})])
        self.st.configure(f"{nombre}.TButton", font=fuente, foreground=fg, background=fondo, padding=0, anchor="w",
                          relief="flat", borderwidth=0, width=0)
        self.st.map(f"{nombre}.TButton", foreground=[("disabled", fg_off or C["apagado"]), ("active", fg_hover)],
                    font=[("active", fuente_hover)], background=[("disabled", fondo), ("active", fondo),
                                                                  ("pressed", fondo)])
        return f"{nombre}.TButton"

    def entrada(self, nombre, fondo, relleno, borde, borde_foco, radio, alto, fg, izq=36, dcha=30, fuente=None):
        r = self.e.px(radio)
        img_n, b = self.formas.nueve(r, relleno, fondo, borde)
        img_f, _ = self.formas.nueve(r, relleno, fondo, borde_foco, 1.6)
        elem = f"{nombre}.campo"
        self._elemento(elem, {"normal": img_n, "focus": img_f}, b)
        self.st.layout(f"{nombre}.TEntry", [(elem, {"sticky": "nsew", "children": [
            ("Entry.padding", {"sticky": "nsew", "children": [("Entry.textarea", {"sticky": "nsew"})]})]})])
        pady = max(0, (self.e.px(alto) - Fuentes.alto(fuente or self.fu.normal)) // 2)
        self.st.configure(f"{nombre}.TEntry", padding=(self.e.px(izq), pady, self.e.px(dcha), pady), foreground=fg,
                          fieldbackground=relleno, insertcolor=fg, selectbackground=C["acento_suave2"],
                          selectforeground=C["texto"], insertwidth=max(1, self.e.px(1)))
        return f"{nombre}.TEntry"

    def tarjeta(self, nombre, fondo, relleno, borde=None, radio=10):
        img, b = self.formas.nueve(self.e.px(radio), relleno, fondo, borde)
        elem = f"{nombre}.fondo"
        self._elemento(elem, {"normal": img}, b)
        self.st.layout(f"{nombre}.TFrame", [(elem, {"sticky": "nsew"})])
        self.st.configure(f"{nombre}.TFrame", background=relleno)
        return f"{nombre}.TFrame"

    def pastilla(self, nombre, fondo, relleno, fg, fuente, alto=22, padx=10):
        r = self.e.px(alto) // 2 - 1
        img, b = self.formas.nueve(r, relleno, fondo)
        elem = f"{nombre}.fondo"
        self._elemento(elem, {"normal": img}, b)
        self.st.layout(f"{nombre}.TLabel", [(elem, {"sticky": "nsew", "children": [
            ("Label.padding", {"sticky": "nsew", "children": [("Label.label", {"sticky": "nsew"})]})]})])
        pady = max(0, (self.e.px(alto) - Fuentes.alto(fuente)) // 2)
        self.st.configure(f"{nombre}.TLabel", font=fuente, foreground=fg, background=relleno,
                          padding=(max(self.e.px(padx), r), pady))
        return f"{nombre}.TLabel"

    def barra(self, nombre, fondo, pulgar, pulgar_hover, ancho=10):
        """Barra de desplazamiento plana (sin flechas)."""
        st = self.st
        st.layout(f"{nombre}.Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
            ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
        st.configure(f"{nombre}.Vertical.TScrollbar", troughcolor=fondo, background=pulgar, bordercolor=fondo,
                     lightcolor=pulgar, darkcolor=pulgar, arrowsize=self.e.px(ancho), gripcount=0, relief="flat",
                     borderwidth=0)
        st.map(f"{nombre}.Vertical.TScrollbar", background=[("pressed", pulgar_hover), ("active", pulgar_hover)],
               lightcolor=[("pressed", pulgar_hover), ("active", pulgar_hover)],
               darkcolor=[("pressed", pulgar_hover), ("active", pulgar_hover)])
        return f"{nombre}.Vertical.TScrollbar"

    def progreso(self, nombre, fondo, color, alto=6):
        self.st.configure(f"{nombre}.Horizontal.TProgressbar", troughcolor=fondo, background=color, bordercolor=fondo,
                          lightcolor=color, darkcolor=color, thickness=self.e.px(alto), borderwidth=0)
        return f"{nombre}.Horizontal.TProgressbar"


def etiqueta(padre, texto="", fuente=None, fg=C["texto"], bg=C["fondo"], **kw):
    kw.setdefault("anchor", "w")
    kw.setdefault("justify", "left")
    return tk.Label(padre, text=texto, font=fuente, fg=fg, bg=bg, bd=0, padx=0, pady=0, highlightthickness=0, **kw)


class Recortador:
    """Recorta textos con «…» para que quepan en un ancho (con caché)."""

    def __init__(self):
        self.cache = {}

    def __call__(self, texto, fuente, ancho):
        if ancho <= 0 or not texto:
            return ""
        clave = (texto, str(fuente), ancho)
        r = self.cache.get(clave)
        if r is None:
            if fuente.measure(texto) <= ancho:
                r = texto
            else:
                lo, hi = 0, len(texto)
                while lo < hi:
                    mid = (lo + hi + 1) // 2
                    if fuente.measure(texto[:mid].rstrip() + "…") <= ancho:
                        lo = mid
                    else:
                        hi = mid - 1
                r = texto[:lo].rstrip() + "…" if lo else "…"
            if len(self.cache) > 30000:
                self.cache.clear()
            self.cache[clave] = r
        return r


RECORTAR = Recortador()


# ---------------------------------------------------------------------------------------------------------------
# campos y textos
# ---------------------------------------------------------------------------------------------------------------
class CampoBusqueda(tk.Frame):
    """Campo de búsqueda redondeado con lupa, texto de ayuda y botón para borrar."""

    def __init__(self, padre, ui, estilo, fondo, relleno, ayuda, icono, icono_borrar, color_ayuda, fuente=None,
                 al_cambiar=None, ancho=None):
        super().__init__(padre, bg=fondo, bd=0, highlightthickness=0)
        self.ui = ui
        self.var = tk.StringVar()
        self.entrada = ttk.Entry(self, textvariable=self.var, style=estilo, font=fuente or ui.fu.normal,
                                 width=ancho or 10)
        self.entrada.pack(fill="x", expand=True)
        try:
            self.izq = int(str(ui.estilos.st.lookup(estilo, "padding")).split()[0]) + 1
        except (ValueError, IndexError, tk.TclError):
            self.izq = ui.e.px(36)
        self.lupa = tk.Label(self, image=icono, bg=relleno, bd=0, cursor="xterm")
        self.lupa.place(in_=self.entrada, x=ui.e.px(11), rely=0.5, anchor="w")
        self.ayuda = etiqueta(self, ayuda, fuente or ui.fu.normal, color_ayuda, relleno, cursor="xterm")
        self.borrar = tk.Label(self, image=icono_borrar, bg=relleno, bd=0, cursor="hand2")
        for w in (self.lupa, self.ayuda):
            w.bind("<Button-1>", lambda e: self.entrada.focus_set())
        self.borrar.bind("<Button-1>", lambda e: self.limpiar())
        self.al_cambiar = al_cambiar
        self.var.trace_add("write", self._cambio)
        self._cambio(avisar=False)

    def _cambio(self, *_, avisar=True):
        hay = bool(self.var.get())
        if hay:
            self.ayuda.place_forget()
            self.borrar.place(in_=self.entrada, relx=1.0, x=-self.ui.e.px(10), rely=0.5, anchor="e")
        else:
            self.borrar.place_forget()
            self.ayuda.place(in_=self.entrada, x=self.izq, rely=0.5, anchor="w")
        if avisar and self.al_cambiar:
            self.al_cambiar(self.var.get())

    def get(self):
        return self.var.get()

    def limpiar(self):
        self.var.set("")
        self.entrada.focus_set()


class TextoRico(tk.Text):
    """Texto de solo lectura con partes en negrita que ajusta su altura al contenido."""

    def __init__(self, padre, fuente, fuente_negrita, fg, bg):
        super().__init__(padre, font=fuente, fg=fg, bg=bg, bd=0, highlightthickness=0, wrap="word", height=1,
                         cursor="arrow", padx=0, pady=0, relief="flat", takefocus=0, spacing1=0, spacing3=0)
        self.tag_configure("negrita", font=fuente_negrita)
        self.configure(state="disabled")
        self.bind("<Configure>", lambda e: self.after_idle(self._ajustar))

    def poner(self, partes, fg=None, bg=None):
        self.configure(state="normal")
        if fg:
            self.configure(fg=fg)
        if bg:
            self.configure(bg=bg)
        self.delete("1.0", "end")
        for texto, negrita in partes:
            self.insert("end", texto, ("negrita",) if negrita else ())
        self.configure(state="disabled")
        self.after_idle(self._ajustar)

    def _ajustar(self):
        try:
            n = self.count("1.0", "end", "displaylines")
        except tk.TclError:
            return
        if isinstance(n, tuple):
            n = n[0]
        n = max(1, n or 1)
        if int(self.cget("height")) != n:
            self.configure(height=n)


class Flujo(tk.Frame):
    """Coloca sus hijos en filas que saltan de línea si no caben (como los chips de filtro)."""

    def __init__(self, padre, bg, sep_x, sep_y):
        super().__init__(padre, bg=bg, bd=0, highlightthickness=0, height=1)
        self.sep_x, self.sep_y = sep_x, sep_y
        self.hijos = []
        self._pendiente = None
        self.bind("<Configure>", lambda e: self.recolocar())

    def agregar(self, w):
        self.hijos.append(w)
        self.recolocar()

    def recolocar(self):
        if self._pendiente is None:
            self._pendiente = self.after_idle(self._colocar)

    def _colocar(self):
        self._pendiente = None
        ancho = max(self.winfo_width(), 1)
        x = y = alto_fila = 0
        for w in self.hijos:
            if not getattr(w, "visible", True):
                w.place_forget()
                continue
            ww, hh = w.winfo_reqwidth(), w.winfo_reqheight()
            if x and x + ww > ancho:
                x, y = 0, y + alto_fila + self.sep_y
                alto_fila = 0
            w.place(x=x, y=y)
            x += ww + self.sep_x
            alto_fila = max(alto_fila, hh)
        alto = y + alto_fila
        if int(self.cget("height")) != max(alto, 1):
            self.configure(height=max(alto, 1))


class MarcoDesplazable(tk.Frame):
    """Marco con contenido desplazable con la rueda (la barra solo aparece si hace falta)."""

    def __init__(self, padre, bg, estilo_barra):
        super().__init__(padre, bg=bg, bd=0, highlightthickness=0)
        self.lienzo = tk.Canvas(self, bg=bg, bd=0, highlightthickness=0)
        self.barra = ttk.Scrollbar(self, orient="vertical", command=self.lienzo.yview, style=estilo_barra)
        self.contenido = tk.Frame(self.lienzo, bg=bg, bd=0, highlightthickness=0)
        self._win = self.lienzo.create_window(0, 0, window=self.contenido, anchor="nw")
        self.lienzo.configure(yscrollcommand=self._barra)
        self.lienzo.pack(side="left", fill="both", expand=True)
        self.contenido.bind("<Configure>", lambda e: self._medidas())
        self.lienzo.bind("<Configure>", lambda e: (self.lienzo.itemconfigure(self._win, width=e.width), self._medidas()))

    def _medidas(self):
        self.lienzo.configure(scrollregion=(0, 0, 0, self.contenido.winfo_reqheight()))

    def _barra(self, a, b):
        if float(a) <= 0.0 and float(b) >= 1.0:
            self.barra.pack_forget()
        elif not self.barra.winfo_ismapped():
            self.barra.pack(side="right", fill="y", before=self.lienzo)
        self.barra.set(a, b)

    def rueda(self, w):
        """Hace que la rueda del ratón sobre w (y sus hijos) desplace el marco."""
        w.bind("<MouseWheel>", self._rueda, add="+")
        w.bind("<Button-4>", lambda e: self.lienzo.yview_scroll(-3, "units"), add="+")
        w.bind("<Button-5>", lambda e: self.lienzo.yview_scroll(3, "units"), add="+")
        for h in w.winfo_children():
            self.rueda(h)

    def _rueda(self, e):
        if self.barra.winfo_ismapped():
            self.lienzo.yview_scroll(int(-e.delta / 40) or (-1 if e.delta > 0 else 1), "units")

    def arriba(self):
        self.lienzo.yview_moveto(0)


# ---------------------------------------------------------------------------------------------------------------
# listas virtuales (solo se dibujan las filas visibles)
# ---------------------------------------------------------------------------------------------------------------
class ListaVirtual(tk.Canvas):
    def __init__(self, padre, alto_fila, bg, barra=None):
        super().__init__(padre, bg=bg, bd=0, highlightthickness=0, takefocus=1)
        self.alto = alto_fila
        self.n = 0
        self.y = 0.0
        self.barra = barra
        self.hover = -1
        self.bind("<Configure>", lambda e: self._redibujar())
        self.bind("<MouseWheel>", self._rueda)
        self.bind("<Button-4>", lambda e: self.desplazar(-3 * self.alto))
        self.bind("<Button-5>", lambda e: self.desplazar(3 * self.alto))
        self.bind("<Motion>", self._movimiento)
        self.bind("<Leave>", lambda e: self._poner_hover(-1))

    # -- desplazamiento
    def total(self):
        return self.n * self.alto

    def _limitar(self):
        self.y = max(0.0, min(self.y, float(max(0, self.total() - self.winfo_height()))))

    def yview(self, *args):
        if not args:
            return self._fracciones()
        if args[0] == "moveto":
            self.y = float(args[1]) * self.total()
        elif args[0] == "scroll":
            paso = self.alto if args[2] == "units" else max(self.alto, self.winfo_height() - self.alto)
            self.y += int(args[1]) * paso
        self._limitar()
        self._redibujar()

    def _fracciones(self):
        tot = max(self.total(), 1)
        return self.y / tot, min(1.0, (self.y + self.winfo_height()) / tot)

    def desplazar(self, dy):
        self.y += dy
        self._limitar()
        self._redibujar()

    def _rueda(self, e):
        if e.delta:
            self.desplazar(-e.delta / 120 * 3 * self.alto)

    def ver(self, i):
        if i < 0:
            return
        top, h = i * self.alto, self.winfo_height()
        if top < self.y:
            self.y = top
        elif top + self.alto > self.y + h:
            self.y = top + self.alto - h
        self._limitar()
        self._redibujar()

    def indice(self, y):
        i = int((self.y + y) // self.alto)
        return i if 0 <= i < self.n else -1

    def _movimiento(self, e):
        self._poner_hover(self.indice(e.y))

    def _poner_hover(self, i):
        if i != self.hover:
            antes, self.hover = self.hover, i
            self.cambio_hover(antes, i)

    def cambio_hover(self, antes, ahora):
        self._redibujar()

    # -- dibujo
    def _redibujar(self):
        self.delete("all")
        h = self.winfo_height()
        if self.n:
            primera = int(self.y // self.alto)
            ultima = min(self.n, int((self.y + h) // self.alto) + 1)
            for i in range(primera, ultima):
                self.dibujar_fila(i, i * self.alto - self.y)
        self.dibujar_extra()
        if self.barra is not None:
            a, b = self._fracciones()
            if self.n and (a > 0 or b < 1):
                if not self.barra.winfo_ismapped():
                    self.barra.grid()
            else:
                self.barra.grid_remove()
            self.barra.set(a, b)

    def dibujar_fila(self, i, y):
        pass

    def dibujar_extra(self):
        pass


class ListaMarcable(ListaVirtual):
    """Lista con casillas (selección múltiple) o con marca (una sola opción), en tema claro u oscuro."""

    def __init__(self, padre, ui, tema="claro", multiple=True, al_cambiar=None, barra=None, alto=34, margen=0):
        self.ui = ui
        self.oscuro = tema == "oscuro"
        bg = C["lat"] if self.oscuro else C["blanco"]
        super().__init__(padre, ui.e.px(alto), bg, barra)
        self.multiple = multiple
        self.al_cambiar = al_cambiar
        self.items = []
        self.seleccion = set()
        self.margen = margen
        self.foco = -1
        self._fondos = {}
        self.bind("<Button-1>", self._clic)
        self.bind("<Up>", lambda e: self._mover(-1))
        self.bind("<Down>", lambda e: self._mover(1))
        self.bind("<space>", lambda e: self._pulsar_foco())
        self.bind("<Return>", lambda e: self._pulsar_foco())
        self.bind("<FocusOut>", lambda e: self._redibujar())

    def poner(self, items, seleccion=None):
        """items: [(clave, texto, cuenta|None, apagado)]."""
        self.items = items
        self.n = len(items)
        if seleccion is not None:
            self.seleccion = set(seleccion)
        self.foco = min(self.foco, self.n - 1)
        self._limitar()
        self._redibujar()

    def _fondo(self, color):
        ancho = max(10, self.winfo_width() - 2 * self.margen)
        clave = (ancho, color)
        if clave not in self._fondos:
            self._fondos[clave] = self.ui.estilos.formas.rect(ancho, self.alto - self.ui.e.px(2), self.ui.e.px(6),
                                                             color, self.cget("bg"))
        return self._fondos[clave]

    def dibujar_fila(self, i, y):
        ui, px = self.ui, self.ui.e.px
        clave, texto, cuenta, apagado = self.items[i]
        sel = clave in self.seleccion
        ancho = self.winfo_width()
        x0 = self.margen
        resaltar = (self.oscuro and sel) or i == self.hover or (i == self.foco and self.focus_get() is self)
        if resaltar:
            if self.oscuro:
                color = C["lat2"] if sel else C["lat_hover"]
            else:
                color = C["gris"]
            self.create_image(x0, y + px(1), image=self._fondo(color), anchor="nw")
        cy = y + self.alto / 2
        if self.multiple:
            nombre = ("casilla-oscuro-" if self.oscuro else "casilla-claro-") + ("si" if sel else "no")
            self.create_image(x0 + px(10), cy, image=ui.iconos(nombre), anchor="w")
        elif sel:
            self.create_image(x0 + px(8), cy, image=ui.iconos("check-acento-16"), anchor="w")
        if self.oscuro:
            fg = C["blanco"] if sel else C["lat_texto"]
            fg_n = C["lat_tenue"]
        else:
            fg = C["apagado"] if apagado else C["texto"]
            fg_n = C["apagado"] if apagado else C["tenue"]
        fuente = ui.fu.semi if sel and (self.oscuro or not self.multiple) else ui.fu.normal
        xt = x0 + px(36)
        derecha = ancho - self.margen - px(10)
        ancho_n = 0
        if cuenta is not None:
            tn = f"{cuenta:,}".replace(",", ".")
            ancho_n = ui.fu.peque.measure(tn) + px(10)
            self.create_text(derecha, cy, text=tn, anchor="e", font=ui.fu.peque, fill=fg_n)
        self.create_text(xt, cy, text=RECORTAR(texto, fuente, derecha - ancho_n - xt), anchor="w", font=fuente,
                         fill=fg)

    def dibujar_extra(self):
        if not self.n and self.vacio:
            self.create_text(self.winfo_width() / 2, self.ui.e.px(24), text=self.vacio, font=self.ui.fu.peque,
                             fill=C["lat_tenue"] if self.oscuro else C["tenue"], anchor="n")

    vacio = ""

    def _clic(self, e):
        self.focus_set()
        i = self.indice(e.y)
        if i >= 0:
            self.foco = i
            self.pulsar(i)

    def pulsar(self, i):
        clave = self.items[i][0]
        if self.multiple:
            self.seleccion.symmetric_difference_update({clave})
        else:
            self.seleccion = {clave}
        self._redibujar()
        if self.al_cambiar:
            self.al_cambiar(set(self.seleccion), clave)

    def _mover(self, d):
        if self.n:
            self.foco = max(0, min(self.n - 1, (self.foco if self.foco >= 0 else -1) + d))
            self.ver(self.foco)
        return "break"

    def _pulsar_foco(self):
        if 0 <= self.foco < self.n:
            self.pulsar(self.foco)
        return "break"


class Columna:
    def __init__(self, cid, titulo, ancho=0, flex=0, minimo=0, ordenable=True):
        self.cid, self.titulo, self.ancho, self.flex, self.minimo, self.ordenable = cid, titulo, ancho, flex, minimo, \
            ordenable
        self.x = self.w = 0


class TablaCentros(tk.Frame):
    """Tabla de centros (cabecera, filas virtuales y barra), pensada para ir dentro de una tarjeta redondeada.

    fila(i) devuelve (clave, nombre, tipo, municipio, teléfono, email, marcado, estado) con estado
    '' / 'nuevo' / 'email'."""

    def __init__(self, padre, ui, fila, al_seleccionar, al_marcar, al_doble, al_ordenar, al_marcar_todos,
                 estado_todos, ancho_email=860):
        super().__init__(padre, bg=C["blanco"], bd=0, highlightthickness=0)
        self.ui, self.fila = ui, fila
        self.al_seleccionar, self.al_marcar, self.al_doble = al_seleccionar, al_marcar, al_doble
        self.al_ordenar, self.al_marcar_todos, self.estado_todos = al_ordenar, al_marcar_todos, estado_todos
        self.ancho_email = ui.e.px(ancho_email)
        px = ui.e.px
        self.activa = None   # clave del centro elegido (su ficha está a la derecha)
        self.cursor = -1     # su posición en la lista
        self.orden = (None, False)
        self.mensaje, self.enlace = "", None
        fu = ui.fu
        ancho_tipo = max(px(84), max(fu.tipo.measure(t) for t in (
            "INFANTIL PRIV.", "ED. ESPECIAL", "CONCERTADO", "INST. ESCOLA", "EXTRANJERO")) + px(14))
        self.columnas = [Columna("sel", "", px(44), ordenable=False), Columna("nombre", "Centro", flex=2, minimo=px(170)),
                         Columna("tipo", "Tipo", ancho_tipo), Columna("mun", "Municipio", flex=1, minimo=px(90)),
                         Columna("tel", "Teléfono", px(112)), Columna("email", "Email", flex=1.5, minimo=px(190))]
        self.visibles = self.columnas[:-1]
        self.radio = px(10)
        self.cab = tk.Canvas(self, height=px(36), bg=C["fondo"], bd=0, highlightthickness=0)
        self.cab.pack(fill="x")
        cuerpo = tk.Frame(self, bg=C["blanco"])
        cuerpo.pack(fill="both", expand=True, padx=1)
        cuerpo.rowconfigure(0, weight=1)
        cuerpo.columnconfigure(0, weight=1)
        self.barra = ttk.Scrollbar(cuerpo, orient="vertical", style=ui.estilo_barra)
        self.lista = _Filas(cuerpo, self)
        self.lista.grid(row=0, column=0, sticky="nsew")
        self.barra.grid(row=0, column=1, sticky="ns")
        self.barra.configure(command=self.lista.yview)
        self.lista.barra = self.barra
        self.esquinas = (esquina(ui.root, self.radio, C["fondo"], C["borde"], C["fondo"], "ii"),
                         esquina(ui.root, self.radio, C["fondo"], C["borde"], C["fondo"], "id"))
        self.cab.bind("<Configure>", lambda e: self._medir())
        self.cab.bind("<Button-1>", self._clic_cabecera)
        self.cab.bind("<Motion>", self._cursor_cabecera)

    # -- datos
    def poner(self, n, mensaje="", enlace=None):
        """n filas; mensaje si no hay ninguna (y un enlace opcional (texto, acción) debajo)."""
        self.lista.n = n
        self.mensaje, self.enlace = mensaje, enlace
        self.lista._limitar()
        self.refrescar()

    def refrescar(self):
        self.lista._redibujar()
        self._dibujar_cabecera()

    def arriba(self):
        self.lista.y = 0
        self.lista._redibujar()

    def seleccionar(self, i, clave, ver=True):
        self.cursor, self.activa = i, clave
        if ver and i >= 0:
            self.lista.ver(i)
        else:
            self.lista._redibujar()

    # -- columnas
    def _medir(self):
        ancho = self.cab.winfo_width() - 2
        self.visibles = self.columnas if ancho >= self.ancho_email else self.columnas[:-1]
        margen = self.ui.e.px(16)
        fijo = sum(c.ancho for c in self.visibles if not c.flex) + margen
        libre = max(0, ancho - fijo)
        flex = sum(c.flex for c in self.visibles)
        x = 0
        for c in self.visibles:
            c.x = x
            c.w = c.ancho if not c.flex else max(c.minimo, int(libre * c.flex / flex))
            x += c.w
        self._dibujar_cabecera()
        self.lista._redibujar()

    def _dibujar_cabecera(self):
        cv, ui, px = self.cab, self.ui, self.ui.e.px
        cv.delete("all")
        w, h, r = cv.winfo_width(), cv.winfo_height(), self.radio
        if w < 4:
            return
        cv.create_image(0, 0, image=self.esquinas[0], anchor="nw")
        cv.create_image(w, 0, image=self.esquinas[1], anchor="ne")
        cv.create_line(r, 0, w - r, 0, fill=C["borde"])
        cv.create_line(0, r, 0, h, fill=C["borde"])
        cv.create_line(w - 1, r, w - 1, h, fill=C["borde"])
        cv.create_line(0, h - 1, w, h - 1, fill=C["borde"])
        cy = h / 2
        estado = self.estado_todos()
        if estado:
            cv.create_image(1 + px(14), cy, image=ui.iconos(f"casilla-claro-{estado}"), anchor="w", tags=("todos",))
        col, desc = self.orden
        for c in self.visibles:
            if not c.titulo:
                continue
            x = 1 + c.x
            cv.create_text(x, cy, text=c.titulo, anchor="w", font=ui.fu.peque_semi, fill=C["suave"])
            if c.cid == col:
                xa = x + ui.fu.peque_semi.measure(c.titulo) + px(9)
                s = px(4)
                pts = (xa - s, cy - s / 2, xa + s, cy - s / 2, xa, cy + s / 2) if desc else \
                    (xa - s, cy + s / 2, xa + s, cy + s / 2, xa, cy - s / 2)
                cv.create_polygon(*pts, fill=C["suave"], outline="")

    def _columna_en(self, x):
        for c in self.visibles:
            if c.x <= x - 1 < c.x + c.w:
                return c
        return None

    def _clic_cabecera(self, e):
        c = self._columna_en(e.x)
        if c is None:
            return
        if c.cid == "sel":
            if self.estado_todos():
                self.al_marcar_todos()
        elif c.ordenable:
            actual, desc = self.orden
            self.orden = (c.cid, not desc if actual == c.cid else False)
            self.al_ordenar(*self.orden)

    def _cursor_cabecera(self, e):
        c = self._columna_en(e.x)
        self.cab.configure(cursor="hand2" if c is not None and (c.ordenable or c.cid == "sel") else "")


class _Filas(ListaVirtual):
    def __init__(self, padre, tabla):
        self.t = tabla
        super().__init__(padre, tabla.ui.e.px(38), C["blanco"])
        self.bind("<Button-1>", self._clic)
        self.bind("<Double-Button-1>", self._doble)
        for tecla, d in (("<Up>", -1), ("<Down>", 1), ("<Prior>", "-p"), ("<Next>", "p"), ("<Home>", "ini"),
                         ("<End>", "fin")):
            self.bind(tecla, lambda e, d=d: self._tecla(d))
        self.bind("<space>", self._espacio)
        self.bind("<Return>", lambda e: (self.t.al_doble(self.t.cursor) if self.t.cursor >= 0 else None, "break")[1])

    def _color(self, i, estado, activa):
        if activa:
            return C["acento_suave"]
        base = C["fila_nuevo"] if estado == "nuevo" else C["fila_email"] if estado == "email" else C["blanco"]
        if i == self.hover:
            return C["fila_hover"] if base == C["blanco"] else mezcla(base, "#000000", 0.035)
        return base

    def cambio_hover(self, antes, ahora):
        for i in (antes, ahora):
            if 0 <= i < self.n:
                k, *_, estado = self.t.fila(i)
                self.itemconfigure(f"bg{i}", fill=self._color(i, estado, k == self.t.activa))

    def dibujar_fila(self, i, y):
        t, ui, px = self.t, self.t.ui, self.t.ui.e.px
        k, nombre, tipo, mun, tel, email, marcado, estado = t.fila(i)
        activa = k == t.activa
        w = self.winfo_width()
        self.create_rectangle(0, y, w, y + self.alto, fill=self._color(i, estado, activa), outline="",
                              tags=(f"bg{i}",))
        self.create_line(0, y + self.alto - 1, w, y + self.alto - 1, fill=C["gris"])
        cy = y + self.alto / 2
        self.create_image(px(14), cy, image=ui.iconos("casilla-claro-si" if marcado else "casilla-claro-no"),
                          anchor="w")
        fuente = ui.fu.semi if activa else ui.fu.normal
        for c in t.visibles[1:]:
            if c.cid == "nombre":
                txt, f, color, hueco = nombre, fuente, C["texto"], px(12)
            elif c.cid == "tipo":
                txt, f, color, hueco = tipo.upper(), ui.fu.tipo, C["suave"], px(8)
            elif c.cid == "mun":
                txt, f, color, hueco = mun, ui.fu.normal, C["texto"], px(10)
            elif c.cid == "tel":
                txt, f, color, hueco = tel, ui.fu.normal, C["texto"], px(8)
            else:
                txt, f, color, hueco = email, ui.fu.normal, C["sec"], px(8)
            if txt:
                self.create_text(c.x, cy, text=RECORTAR(txt, f, c.w - hueco), anchor="w", font=f, fill=color)

    def dibujar_extra(self):
        t = self.t
        if self.n or not t.mensaje:
            return
        ui, px = t.ui, t.ui.e.px
        cx, cy = self.winfo_width() / 2, max(px(60), self.winfo_height() * 0.38)
        self.create_text(cx, cy, text=t.mensaje, font=ui.fu.normal, fill=C["suave"], justify="center",
                         width=max(200, self.winfo_width() - px(80)))
        if t.enlace:
            texto, accion = t.enlace
            lineas = t.mensaje.count("\n") + 1
            item = self.create_text(cx, cy + (lineas / 2 + 1) * ui.fu.normal.metrics("linespace"), text=texto,
                                    font=ui.fu.semi, fill=C["acento"], tags=("enlace",))
            self.tag_bind(item, "<Button-1>", lambda e: accion())
            self.tag_bind(item, "<Enter>", lambda e: self.configure(cursor="hand2"))
            self.tag_bind(item, "<Leave>", lambda e: self.configure(cursor=""))

    def _clic(self, e):
        self.focus_set()
        i = self.indice(e.y)
        if i < 0:
            return
        if e.x < self.t.ui.e.px(44):
            self.t.al_marcar(i)
        else:
            self.t.al_seleccionar(i)

    def _doble(self, e):
        i = self.indice(e.y)
        if i >= 0 and e.x >= self.t.ui.e.px(44):
            self.t.al_doble(i)

    def _tecla(self, d):
        if not self.n:
            return "break"
        actual = self.t.cursor
        pagina = max(1, int(self.winfo_height() // self.alto) - 1)
        if d == "ini":
            i = 0
        elif d == "fin":
            i = self.n - 1
        elif d in ("p", "-p"):
            i = (actual if actual >= 0 else 0) + (pagina if d == "p" else -pagina)
        else:
            i = actual + d if actual >= 0 else 0
        self.t.al_seleccionar(max(0, min(self.n - 1, i)))
        return "break"

    def _espacio(self, e):
        if 0 <= self.t.cursor < self.n:
            self.t.al_marcar(self.t.cursor)
        return "break"


# ---------------------------------------------------------------------------------------------------------------
# desplegable de los chips de filtro
# ---------------------------------------------------------------------------------------------------------------
class Desplegable(tk.Toplevel):
    """Lista flotante bajo un chip: una opción (se cierra al elegir) o varias (se aplican al momento)."""

    def __init__(self, ui, ancla, items, seleccion, multiple, al_cambiar, buscador=False, ayuda="Buscar",
                 al_cerrar=None, ancho=300):
        super().__init__(ui.root)
        self.ui, self.ancla, self.todos = ui, ancla, items
        self.al_cambiar, self.al_cerrar = al_cambiar, al_cerrar
        px = ui.e.px
        self.withdraw()
        self.overrideredirect(True)
        self.configure(bg=C["borde2"])
        try:
            self.attributes("-topmost", True)
        except tk.TclError:
            pass
        cuerpo = tk.Frame(self, bg=C["blanco"])
        cuerpo.pack(fill="both", expand=True, padx=1, pady=1)
        self.campo = None
        if buscador:
            self.campo = CampoBusqueda(cuerpo, ui, ui.estilo_buscar_popup, C["blanco"], C["blanco"], ayuda,
                                       ui.iconos("buscar-suave-16"), ui.iconos("cerrar-tenue-14"), C["apagado"],
                                       al_cambiar=self._buscar)
            self.campo.pack(fill="x", padx=px(8), pady=(px(8), px(4)))
            self.campo.entrada.bind("<Down>", lambda e: (self.lista.focus_set(), self.lista._mover(1), "break")[2])
            self.campo.entrada.bind("<Return>", self._intro)
        marco = tk.Frame(cuerpo, bg=C["blanco"])
        marco.pack(fill="both", expand=True, padx=px(6), pady=px(6))
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)
        barra = ttk.Scrollbar(marco, orient="vertical", style=ui.estilo_barra)
        self.lista = ListaMarcable(marco, ui, "claro", multiple, self._cambio, barra, alto=32)
        self.lista.vacio = "Ninguna coincidencia"
        self.lista.grid(row=0, column=0, sticky="nsew")
        barra.grid(row=0, column=1, sticky="ns")
        barra.configure(command=self.lista.yview)
        self.lista.seleccion = set(seleccion)
        if multiple:
            pie = tk.Frame(cuerpo, bg=C["blanco"])
            pie.pack(fill="x", padx=px(10), pady=(0, px(10)))
            tk.Frame(cuerpo, bg=C["borde"], height=1).pack(fill="x", before=pie, pady=(0, px(8)))
            self.btn_borrar = ttk.Button(pie, text="Quitar selección", style=ui.estilo_enlace, command=self._borrar,
                                         cursor="hand2")
            self.btn_borrar.pack(side="left")
            ttk.Button(pie, text="Listo", style=ui.estilo_primario_peq, command=self.cerrar, cursor="hand2").pack(
                side="right")
        self.lista.poner(items)
        self._estado_borrar()
        # medidas y posición
        filas = min(len(items), 10) or 1
        alto = filas * self.lista.alto + px(12) + (px(50) if buscador else 0) + (px(56) if multiple else 0)
        ancho = max(px(ancho), ancla.winfo_width())
        x, y = ancla.winfo_rootx(), ancla.winfo_rooty() + ancla.winfo_height() + px(4)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        if y + alto > sh - px(40):
            y = max(0, ancla.winfo_rooty() - alto - px(4))
        x = max(0, min(x, sw - ancho - px(8)))
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.deiconify()
        self.lift()
        self.bind("<Escape>", lambda e: self.cerrar())
        # sin «grab» (en algunos sistemas tarda): un clic en cualquier otro sitio de la ventana lo cierra
        self.cerrado = False
        ui.root.bind_all("<Button-1>", self._clic_fuera, add="+")
        self._id_mover = ui.root.bind("<Configure>", self._ventana_movida, add="+")
        self.bind("<FocusOut>", lambda e: self.after(120, self._comprobar_foco))
        (self.campo.entrada if self.campo else self.lista).focus_force()

    def _estado_borrar(self):
        if hasattr(self, "btn_borrar"):
            self.btn_borrar.state(["!disabled"] if self.lista.seleccion else ["disabled"])

    def _buscar(self, texto):
        t = norm(texto)
        items = [it for it in self.todos if t in norm(it[1])] if t else self.todos
        self.lista.y = 0
        self.lista.foco = 0 if t and items else -1
        self.lista.poner(items)

    def _intro(self, e):
        if self.lista.n == 1 or (self.lista.foco >= 0 and self.lista.n):
            self.lista.pulsar(max(self.lista.foco, 0))
        return "break"

    def _cambio(self, seleccion, clave):
        self._estado_borrar()
        self.al_cambiar(seleccion)
        if not self.lista.multiple:
            self.cerrar()

    def _borrar(self):
        self.lista.seleccion = set()
        self.lista._redibujar()
        self._estado_borrar()
        self.al_cambiar(set())

    def _clic_fuera(self, e):
        if self.cerrado:
            return
        try:
            x, y = e.x_root, e.y_root
            dentro = (self.winfo_rootx() <= x < self.winfo_rootx() + self.winfo_width()
                      and self.winfo_rooty() <= y < self.winfo_rooty() + self.winfo_height())
        except tk.TclError:
            dentro = False
        if not dentro:
            # si el clic es en el propio chip, que no se vuelva a abrir al soltar el botón
            self.ui._cierre_desplegable = (self.ancla, time.monotonic())
            self.cerrar()

    def _ventana_movida(self, e):
        if e.widget is self.ui.root:
            self.cerrar()

    def _comprobar_foco(self):
        if self.cerrado:
            return
        try:
            if self.focus_get() is None:
                self.cerrar()
        except (tk.TclError, KeyError):
            pass

    def cerrar(self):
        if self.cerrado:
            return
        self.cerrado = True
        try:
            self.ui.root.unbind_all("<Button-1>")
            self.ui.root.unbind("<Configure>", self._id_mover)
        except tk.TclError:
            pass
        try:
            self.destroy()
        except tk.TclError:
            pass
        if self.al_cerrar:
            cb, self.al_cerrar = self.al_cerrar, None
            cb()
