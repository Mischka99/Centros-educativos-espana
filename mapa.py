# -*- coding: utf-8 -*-
"""
Mapa de la ficha del centro.

- Plano: callejero del Instituto Geográfico Nacional (servicio WMTS «IGN Base», teselas PNG de 256 px en la
  proyección de Google Maps). Tkinter muestra PNG sin ayuda, así que no hace falta ninguna librería ni programa nuevo.
- Situación de cada centro: geocodificador oficial CartoCiudad (IGN/CNIG) con la dirección del centro. Solo se busca
  el centro que se abre (y, en el mapa ampliado, los demás del mismo municipio que hay en la lista).
- Todo se guarda en %LOCALAPPDATA%\\CentrosEducativos\\mapa: las teselas 90 días y las coordenadas hasta que cambie
  la dirección del centro. Licencia de los datos: CC BY 4.0 (© Instituto Geográfico Nacional · CartoCiudad).

Las descargas van en hilos aparte; sus resultados vuelven al hilo de la ventana con la función «entregar».
"""

import hashlib
import json
import math
import os
import queue
import re
import threading
import time
import urllib.parse
from collections import OrderedDict

import tkinter as tk

import centros_educativos as ce
from interfaz import C

URL_TESELA = ("https://www.ign.es/wmts/ign-base?SERVICE=WMTS&REQUEST=GetTile&VERSION=1.0.0&LAYER=IGNBaseTodo"
              "&STYLE=default&TILEMATRIXSET=GoogleMapsCompatible&TILEMATRIX={z}&TILEROW={y}&TILECOL={x}"
              "&FORMAT=image/png")
URL_GEOCODER = "https://www.cartociudad.es/geocoder/api/geocoder/find?q="
ATRIBUCION_CORTA = "© IGN · CartoCiudad"
ATRIBUCION = "© Instituto Geográfico Nacional · CartoCiudad"
TAM = 256
ZOOM_MIN, ZOOM_MAX = 6, 19
DIAS_TESELAS = 90          # después se vuelve a descargar (mientras, se muestra la guardada)
DIAS_BORRAR = 180          # las teselas sin usar desde hace más se borran
DIAS_SIN_RESULTADO = 30    # una dirección no encontrada se vuelve a buscar pasado este tiempo
PAUSA_GEO = 0.2            # entre consultas a CartoCiudad
MAX_MEMORIA = 150          # teselas en memoria (cada una ocupa ~260 KB en Tk)
ESPERA_SIN_RED = 20        # segundos sin volver a intentarlo tras un fallo de conexión
FIRMA_PNG = b"\x89PNG\r\n\x1a\n"
TIPOS_EXACTOS = {"portal", "pk", "punto kilometrico", "punto_kilometrico", "refcatastral"}
TIPOS_CALLE = {"callejero", "vial", "carretera", "via"}


# ---------------------------------------------------------------------------------------------- proyección
def a_pixeles(lat, lon, z):
    """Coordenadas geográficas → píxeles del mundo a ese zoom (proyección Web Mercator)."""
    n = TAM * 2 ** z
    s = math.sin(math.radians(max(-85.0, min(85.0, lat))))
    return (lon + 180.0) / 360.0 * n, (0.5 - math.log((1 + s) / (1 - s)) / (4 * math.pi)) * n


def a_coordenadas(x, y, z):
    n = TAM * 2 ** z
    return math.degrees(math.atan(math.sinh(math.pi * (1 - 2 * y / n)))), x / n * 360.0 - 180.0


# ---------------------------------------------------------------------------------------------- direcciones
_ABREVIATURAS = [
    (r"^\s*c\s*/\s*", "Calle "), (r"\bc\s*/\s*", "Calle "), (r"\bcl\.\s*", "Calle "), (r"\bcalle\.\s*", "Calle "),
    (r"\bavda?\.?\s+", "Avenida "), (r"\bav\.\s*", "Avenida "), (r"\bavgda\.?\s+", "Avinguda "),
    (r"\bpza\.?\s+", "Plaza "), (r"\bplza\.?\s+", "Plaza "), (r"\bpl\.\s*", "Plaza "), (r"\bpça\.?\s+", "Plaça "),
    (r"\bctra\.?\s+", "Carretera "), (r"\bp[º°o]\.?\s+", "Paseo "), (r"\bpg\.\s*", "Passeig "),
    (r"\burb\.\s*", "Urbanización "), (r"\bbda\.\s*", "Barriada "), (r"\bcmno\.\s*", "Camino "),
    (r"\bn[º°]\.?\s*", ""), (r"\bnúm\.?\s*", ""), (r"\bnum\.?\s*", ""),
    (r"\bs\s*/\s*n\b", ""), (r"\bsin\s+n[uú]mero\b", ""),
]


def consulta_direccion(r):
    """Dirección del centro escrita como la entiende CartoCiudad: «Calle Perú 4, 07007 Palma»."""
    d = re.sub(r"\([^)]*\)", " ", r.get("dir") or "")
    for patron, cambio in _ABREVIATURAS:
        d = re.sub(patron, cambio, d, flags=re.I)
    d = re.sub(r"\s+", " ", d).strip(" ,.-")
    lugar = " ".join(x for x in (r.get("cp") or "", r.get("mun") or "") if x)
    return ", ".join(x for x in (d, lugar) if x)


def clave_geo(r):
    """Clave de las coordenadas de un centro: cambia si cambia su dirección (y entonces se vuelve a buscar)."""
    base = "|".join(ce.norm(r.get(c) or "") for c in ("dir", "cp", "mun", "prov"))
    return (r.get("code") or r.get("nombre") or "?") + "|" + hashlib.sha1(base.encode("utf-8")).hexdigest()[:10]


def _respuesta(data):
    """Primer resultado de CartoCiudad (JSON simple o GeoJSON) con lat/lng numéricas, o None."""
    if isinstance(data, list):
        data = data[0] if data else None
    if isinstance(data, dict) and "features" in data:
        feats = data.get("features") or []
        if not feats:
            return None
        props = dict(feats[0].get("properties") or {})
        geo = feats[0].get("geometry") or {}
        if (props.get("lat") is None or props.get("lng") is None) and geo.get("type") == "Point":
            props["lng"], props["lat"] = geo["coordinates"][:2]
        data = props
    if not isinstance(data, dict):
        return None
    try:
        lat, lon = float(data.get("lat")), float(data.get("lng"))
    except (TypeError, ValueError):
        return None
    if not (27.0 <= lat <= 44.5 and -19.0 <= lon <= 5.0):   # fuera de España
        return None
    data["lat"], data["lng"] = lat, lon
    return data


def _mismo_sitio(props, d):
    """El resultado está en la provincia (por el código postal) o en el municipio del centro."""
    cp = (d.get("cp") or "").strip()
    cod = str(props.get("provinceCode") or props.get("postalCode") or "")[:2]
    if len(cp) == 5 and cp.isdigit() and cod.isdigit():
        return cod == cp[:2]
    a, b = ce.norm(props.get("muni") or ""), ce.norm(d.get("mun") or "")
    return not a or not b or a in b or b in a


# ---------------------------------------------------------------------------------------------- servicio
class ServicioMapa:
    """Teselas del IGN y coordenadas de CartoCiudad, con caché en disco y descargas en segundo plano.

    tesela() y situar() se llaman desde el hilo de la ventana; los avisos llegan también a ese hilo."""

    def __init__(self, carpeta, entregar, http=None, registrar=None):
        self.dir = os.path.join(carpeta, "mapa")
        self.dir_teselas = os.path.join(self.dir, "teselas")
        self.ruta_coord = os.path.join(self.dir, "coordenadas.json")
        self.entregar = entregar
        self.http = http
        self.memoria = OrderedDict()      # (z, x, y) → PhotoImage (solo en el hilo de la ventana)
        self.pedidas = {}                 # (z, x, y) → [avisos]
        self.geo_pedidas = {}             # clave → {"avisos": [...], "prioridad": n}
        self.cola_t = queue.LifoQueue()   # lo último que se mira, primero
        self.cola_g = queue.PriorityQueue()
        self.coords = None
        self.cambios = 0
        self.lock = threading.Lock()
        self.parar = threading.Event()
        self.hilos = []
        self.sin_red_hasta = 0.0
        self.ultimo_error = ""
        self.registrar = registrar        # registrar(texto) en el hilo de la ventana (historial)
        self._ultimo_aviso = 0.0
        self._n = 0

    # -- hilos
    def _arrancar(self):
        if self.hilos:
            return
        if self.http is None:
            self.http = ce.Http(timeout=20)
        for objetivo in (self._trabajo_teselas, self._trabajo_teselas, self._trabajo_geo):
            h = threading.Thread(target=objetivo, daemon=True)
            h.start()
            self.hilos.append(h)

    def cerrar(self):
        self.parar.set()
        for _ in range(2):
            self.cola_t.put(None)
        self.cola_g.put((-1, -1, None, None))
        self._guardar_coords()

    def descargar(self, url):
        """Descarga una URL (aparte para poder simularla en las pruebas)."""
        return self.http.fetch(url, timeout=20, retries=1)

    # -- teselas
    def _ruta(self, k):
        z, x, y = k
        return os.path.join(self.dir_teselas, str(z), str(x), f"{y}.png")

    def tesela(self, z, x, y, aviso):
        """PhotoImage de la tesela si está a mano (memoria o disco); si no, la pide y avisa al llegar."""
        k = (z, x, y)
        img = self.memoria.get(k)
        if img is not None:
            self.memoria.move_to_end(k)
            return img
        ruta = self._ruta(k)
        fresca = False
        try:
            edad = time.time() - os.path.getmtime(ruta)
            img = tk.PhotoImage(file=ruta)
            fresca = edad < DIAS_TESELAS * 86400
            self.memoria[k] = img
            while len(self.memoria) > MAX_MEMORIA:
                self.memoria.popitem(last=False)
        except (OSError, tk.TclError):
            img = None
            if os.path.exists(ruta):  # dañada: fuera
                try:
                    os.remove(ruta)
                except OSError:
                    pass
        if (img is None or not fresca) and time.time() >= self.sin_red_hasta:
            self._pedir_tesela(k, aviso)
        return img

    def _fallo(self, e, que):
        """Tras un fallo de conexión se espera un poco antes de reintentar; se anota en el historial."""
        self.sin_red_hasta = time.time() + ESPERA_SIN_RED
        self.ultimo_error = f"{type(e).__name__}: {e}"
        if self.registrar and time.time() - self._ultimo_aviso > 60:
            self._ultimo_aviso = time.time()
            self.entregar(self.registrar, f"Mapa: no se ha podido {que} ({self.ultimo_error}).")

    def _pedir_tesela(self, k, aviso):
        if k in self.pedidas:
            if aviso not in self.pedidas[k]:
                self.pedidas[k].append(aviso)
            return
        self.pedidas[k] = [aviso]
        self._arrancar()
        self.cola_t.put(k)

    def _trabajo_teselas(self):
        while not self.parar.is_set():
            k = self.cola_t.get()
            if k is None or self.parar.is_set():
                return
            if time.time() < self.sin_red_hasta:
                self.entregar(self._tesela_lista, k, False)
                continue
            z, x, y = k
            try:
                data = self.descargar(URL_TESELA.format(z=z, x=x, y=y))
                if not data.startswith(FIRMA_PNG):
                    raise ValueError("el servidor de mapas no ha devuelto una imagen")
                ruta = self._ruta(k)
                os.makedirs(os.path.dirname(ruta), exist_ok=True)
                tmp = f"{ruta}.{threading.get_ident()}.tmp"
                with open(tmp, "wb") as f:
                    f.write(data)
                os.replace(tmp, ruta)
                self.entregar(self._tesela_lista, k, True)
            except Exception as e:  # sin conexión, certificado, servidor caído…
                self._fallo(e, "descargar el plano del IGN")
                self.entregar(self._tesela_lista, k, False)

    def _tesela_lista(self, k, ok):
        if ok:
            self.memoria.pop(k, None)  # se vuelve a leer del disco la versión nueva
        for aviso in self.pedidas.pop(k, []):
            try:
                aviso(k, ok)
            except tk.TclError:
                pass

    def limpiar_viejas(self):
        """Borra teselas que no se han usado en mucho tiempo (se llama en segundo plano al abrir)."""
        limite = time.time() - DIAS_BORRAR * 86400
        for d, _, fs in os.walk(self.dir_teselas):
            for f in fs:
                p = os.path.join(d, f)
                try:
                    if os.path.getmtime(p) < limite or f.endswith(".tmp"):
                        os.remove(p)
                except OSError:
                    pass

    # -- coordenadas
    def _cargar_coords(self):
        if self.coords is not None:
            return
        with self.lock:
            if self.coords is not None:
                return
            try:
                with open(self.ruta_coord, encoding="utf-8") as f:
                    datos = json.load(f)
                self.coords = datos if isinstance(datos, dict) else {}
            except (OSError, ValueError):
                self.coords = {}

    def _guardar_coords(self):
        with self.lock:
            if not self.cambios or self.coords is None:
                return
            datos = dict(self.coords)
            self.cambios = 0
        try:
            os.makedirs(self.dir, exist_ok=True)
            tmp = self.ruta_coord + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(datos, f, ensure_ascii=False, separators=(",", ":"))
            os.replace(tmp, self.ruta_coord)
        except OSError:
            pass

    def conocida(self, r):
        """Coordenadas ya guardadas del centro (sin buscar nada), o None."""
        self._cargar_coords()
        with self.lock:
            res = self.coords.get(clave_geo(r))
        if res and (res.get("prec") != "no" or time.time() - res.get("t", 0) < DIAS_SIN_RESULTADO * 86400):
            return res
        return None

    def situar(self, r, aviso, prioridad=0):
        """Coordenadas del centro si ya se conocen; si no, None y se buscan: aviso(clave, resultado, error)."""
        res = self.conocida(r)
        if res is not None:
            return res
        k = clave_geo(r)
        datos = {c: r.get(c) or "" for c in ("dir", "cp", "mun", "prov", "loc", "nombre")}
        pedida = self.geo_pedidas.get(k)
        if pedida is not None:
            if aviso not in pedida["avisos"]:
                pedida["avisos"].append(aviso)
            if prioridad < pedida["prioridad"]:  # la ficha no espera detrás de los del mapa ampliado
                pedida["prioridad"] = prioridad
                self._n += 1
                self.cola_g.put((prioridad, self._n, k, datos))
            return None
        self.geo_pedidas[k] = {"avisos": [aviso], "prioridad": prioridad}
        self._arrancar()
        self._n += 1
        self.cola_g.put((prioridad, self._n, k, datos))
        return None

    def cancelar_lote(self):
        """Quita de la cola las búsquedas del mapa ampliado (prioridad 1) que aún no han empezado."""
        quedan = []
        while True:
            try:
                item = self.cola_g.get_nowait()
            except queue.Empty:
                break
            pedida = self.geo_pedidas.get(item[2]) if item[2] is not None else None
            if item[0] >= 1 and pedida is not None and pedida["prioridad"] >= 1:
                self.geo_pedidas.pop(item[2], None)
            else:
                quedan.append(item)
        for item in quedan:
            self.cola_g.put(item)

    def _buscar(self, texto):
        data = json.loads(self.descargar(URL_GEOCODER + urllib.parse.quote(texto)).decode("utf-8"))
        return _respuesta(data)

    def geocodificar(self, d):
        """Busca la dirección y, si no aparece, el municipio. Devuelve {lat, lon, prec} (prec: exacta, calle,
        aprox o no). Los fallos de red se propagan (no se guardan)."""
        props = None
        q = consulta_direccion(d)
        if d.get("dir"):
            props = self._buscar(q)
            if props is not None and not _mismo_sitio(props, d):
                props = None
        prec = None
        if props is not None:
            tipo = ce.norm(props.get("type") or "")
            prec = "exacta" if tipo in TIPOS_EXACTOS else "calle" if tipo in TIPOS_CALLE else "aprox"
        elif d.get("mun"):
            time.sleep(PAUSA_GEO)
            props = self._buscar(", ".join(x for x in (d.get("mun"), d.get("prov")) if x))
            if props is not None and not _mismo_sitio(props, d):
                props = None
            prec = "aprox"
        if props is None:
            return {"prec": "no", "t": int(time.time())}
        return {"lat": round(props["lat"], 6), "lon": round(props["lng"], 6), "prec": prec, "t": int(time.time())}

    def _trabajo_geo(self):
        ultimo_guardado = time.time()
        while not self.parar.is_set():
            try:
                _p, _n, k, d = self.cola_g.get(timeout=5)
            except queue.Empty:
                if self.cambios:
                    self._guardar_coords()
                continue
            if k is None or self.parar.is_set():
                return
            self._cargar_coords()
            with self.lock:
                hecho = self.coords.get(k)
            if hecho is not None and time.time() - hecho.get("t", 0) < 60:  # repetida en la cola: ya está
                self.entregar(self._situado, k, hecho, False)
                continue
            if time.time() < self.sin_red_hasta:
                self.entregar(self._situado, k, None, True)
                continue
            try:
                res = self.geocodificar(d)
            except Exception as e:
                self._fallo(e, "consultar la dirección en CartoCiudad")
                self.entregar(self._situado, k, None, True)
                continue
            self._cargar_coords()
            with self.lock:
                self.coords[k] = res
                self.cambios += 1
            self.entregar(self._situado, k, res, False)
            if self.cambios >= 25 or time.time() - ultimo_guardado > 10:
                self._guardar_coords()
                ultimo_guardado = time.time()
            time.sleep(PAUSA_GEO)

    def _situado(self, k, res, error):
        pedida = self.geo_pedidas.pop(k, None)
        for aviso in (pedida or {}).get("avisos", []):
            try:
                aviso(k, res, error)
            except tk.TclError:
                pass


# ---------------------------------------------------------------------------------------------- visor
class VisorMapa(tk.Canvas):
    """Plano con teselas, chincheta del centro y puntos de otros centros. Se arrastra para moverlo; botones +/−,
    doble clic para acercar y, si rueda=True, la rueda del ratón acerca y aleja."""

    def __init__(self, padre, ui, servicio, alto=None, zoom=16, rueda=False, atribucion=ATRIBUCION_CORTA,
                 esquinas=False, al_pulsar=None):
        super().__init__(padre, bg="#eeebe5", height=alto or ui.e.px(170), highlightthickness=0, bd=0,
                         takefocus=0)
        self.ui, self.s = ui, servicio
        self.z = zoom
        self.centro = None        # (lat, lon) del centro de la vista
        self.pin = None           # (lat, lon) del centro elegido
        self.puntos = []          # [(lat, lon, imagen, clave, nombre)]
        self.texto = ""           # mensaje sobre el plano (buscando, sin conexión…)
        self.atribucion = atribucion
        self.esquinas = esquinas
        self.al_pulsar = al_pulsar
        self.fallo_red = False
        self._dibujados = []      # puntos visibles: (x, y, clave, nombre)
        self._arrastre = None
        self._pendiente = None
        self._sobre = None
        px = ui.e.px
        self.lado = px(28)
        self.bind("<Configure>", lambda e: self.redibujar())
        self.bind("<ButtonPress-1>", self._pulsar)
        self.bind("<B1-Motion>", self._mover)
        self.bind("<ButtonRelease-1>", self._soltar)
        self.bind("<Double-Button-1>", self._doble)
        self.bind("<Motion>", self._movimiento)
        self.bind("<Leave>", lambda e: self._poner_sobre(None))
        if rueda:
            self.bind("<MouseWheel>", lambda e: self._zoom_en(e.x, e.y, 1 if e.delta > 0 else -1))
            self.bind("<Button-4>", lambda e: self._zoom_en(e.x, e.y, 1))
            self.bind("<Button-5>", lambda e: self._zoom_en(e.x, e.y, -1))

    # -- API
    def ver(self, lat, lon, zoom=None):
        self.centro = (lat, lon)
        if zoom is not None:
            self.z = max(ZOOM_MIN, min(ZOOM_MAX, zoom))
        self.texto = ""
        self.redibujar()

    def poner_pin(self, lat, lon):
        self.pin = (lat, lon)
        self.redibujar()

    def poner_puntos(self, puntos):
        self.puntos = puntos
        self.redibujar()

    def mensaje(self, texto, limpiar=True):
        """Muestra un aviso sobre el plano; con limpiar=True quita también el plano (aún no se sabe dónde ir)."""
        if limpiar:
            self.centro = self.pin = None
            self.puntos = []
        self.texto = texto
        self.redibujar()

    def recentrar(self):
        if self.pin:
            self.centro = self.pin
            self.redibujar()

    def redibujar(self):
        if self._pendiente is None:
            self._pendiente = self.after_idle(self._dibujar)

    # -- dibujo
    def _origen(self):
        cx, cy = a_pixeles(self.centro[0], self.centro[1], self.z)
        return cx - self.winfo_width() / 2, cy - self.winfo_height() / 2

    def _dibujar(self):
        self._pendiente = None
        try:
            self.delete("all")
        except tk.TclError:
            return
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4:
            return
        px = self.ui.e.px
        self._dibujados = []
        cargadas = 0
        if self.centro is not None:
            x0, y0 = self._origen()
            n = 2 ** self.z
            for tx in range(int(x0 // TAM), int((x0 + w - 1) // TAM) + 1):
                for ty in range(int(y0 // TAM), int((y0 + h - 1) // TAM) + 1):
                    if not 0 <= ty < n:
                        continue
                    img = self.s.tesela(self.z, tx % n, ty, self._llega)
                    X, Y = tx * TAM - x0, ty * TAM - y0
                    if img is not None:
                        self.create_image(X, Y, image=img, anchor="nw")
                        cargadas += 1
                    else:
                        self.create_rectangle(X, Y, X + TAM, Y + TAM, fill="#eeebe5", outline="#e6e2da")
            for lat, lon, img, clave, nombre in self.puntos:
                x, y = a_pixeles(lat, lon, self.z)
                x, y = x - x0, y - y0
                if -10 <= x <= w + 10 and -10 <= y <= h + 10:
                    self.create_image(x, y, image=img)
                    self._dibujados.append((x, y, clave, nombre))
            if self.pin is not None:
                x, y = a_pixeles(self.pin[0], self.pin[1], self.z)
                self.create_image(x - x0, y - y0, image=self.ui.iconos("pin"), anchor="s")
            self._controles(w)
        texto = self.texto
        if not texto and self.centro is not None and not cargadas and self.fallo_red:
            texto = "No se puede cargar el mapa. Comprueba la conexión a internet."
        if texto:
            self._aviso(texto, w, h)
        self._sobre_dibujar()
        # atribución (obligatoria por la licencia) y bordes redondeados
        if self.centro is not None:
            t = self.create_text(w - px(5), h - px(2), text=self.atribucion, anchor="se", font=self.ui.fu.mini,
                                 fill=C["sec"])
            x1, y1, x2, y2 = self.bbox(t)
            self.create_rectangle(x1 - px(5), y1 - px(1), w, h, fill="#f7f6f3", outline="")
            self.tag_raise(t)
        if self.esquinas:
            for nombre, x, y, anc in (("mapa-esq-ii", 0, 0, "nw"), ("mapa-esq-id", w, 0, "ne"),
                                      ("mapa-esq-bi", 0, h, "sw"), ("mapa-esq-bd", w, h, "se")):
                self.create_image(x, y, image=self.ui.iconos(nombre), anchor=anc)
            r = px(10)
            borde = C["borde"]
            self.create_line(r, 0, w - r, 0, fill=borde)
            self.create_line(r, h - 1, w - r, h - 1, fill=borde)
            self.create_line(0, r, 0, h - r, fill=borde)
            self.create_line(w - 1, r, w - 1, h - r, fill=borde)

    def _controles(self, w):
        px = self.ui.e.px
        l, m = self.lado, px(8)
        x1, x0 = w - m, w - m - l
        for n, (tag, activo) in enumerate((("mas", self.z < ZOOM_MAX), ("menos", self.z > ZOOM_MIN),
                                          ("centrar", self.pin is not None))):
            y0 = m + n * l + (px(6) if tag == "centrar" else 0)
            color = C["texto"] if activo else C["apagado"]
            self.create_rectangle(x0, y0, x1, y0 + l, fill=C["blanco"], outline=C["borde2"], tags=("ctl", tag))
            cx, cy, a = (x0 + x1) / 2, y0 + l / 2, px(5)
            ancho = max(1, px(2))
            if tag == "mas":
                self.create_line(cx - a, cy, cx + a, cy, fill=color, width=ancho, tags=("ctl", tag))
                self.create_line(cx, cy - a, cx, cy + a, fill=color, width=ancho, tags=("ctl", tag))
            elif tag == "menos":
                self.create_line(cx - a, cy, cx + a, cy, fill=color, width=ancho, tags=("ctl", tag))
            else:
                self.create_oval(cx - a + 1, cy - a + 1, cx + a - 1, cy + a - 1, outline=color, width=ancho,
                                 tags=("ctl", tag))
                self.create_oval(cx - 1, cy - 1, cx + 1, cy + 1, fill=color, outline=color, tags=("ctl", tag))

    def _aviso(self, texto, w, h):
        px = self.ui.e.px
        t = self.create_text(w / 2, h / 2, text=texto, font=self.ui.fu.peque, fill=C["sec"], justify="center",
                             width=max(px(120), w - px(60)))
        x1, y1, x2, y2 = self.bbox(t)
        self.create_rectangle(x1 - px(10), y1 - px(6), x2 + px(10), y2 + px(6), fill=C["blanco"], outline=C["borde"])
        self.tag_raise(t)

    def _sobre_dibujar(self):
        if not self._sobre:
            return
        x, y, _clave, nombre = self._sobre
        px = self.ui.e.px
        t = self.create_text(x + px(12), y - px(12), text=nombre, anchor="sw", font=self.ui.fu.peque_semi,
                             fill=C["texto"])
        x1, y1, x2, y2 = self.bbox(t)
        self.create_rectangle(x1 - px(6), y1 - px(3), x2 + px(6), y2 + px(3), fill=C["blanco"], outline=C["borde2"])
        self.tag_raise(t)

    def _llega(self, k, ok):
        if not ok:
            if not self.fallo_red:  # se avisa una vez y se reintenta pasado un rato (sin bucles)
                self.fallo_red = True
                self.redibujar()
                self.after(int(ESPERA_SIN_RED * 1000) + 500, self._reintentar)
            return
        if k[0] == self.z:
            self.fallo_red = False
            self.redibujar()

    def _reintentar(self):
        self.fallo_red = False
        self.redibujar()

    # -- ratón
    def _control_en(self, x, y):
        for item in self.find_overlapping(x, y, x, y):
            tags = self.gettags(item)
            if "ctl" in tags:
                return next(t for t in tags if t != "ctl" and t != "current")
        return None

    def _punto_en(self, x, y):
        radio = self.ui.e.px(10)
        mejor = None
        for px_, py_, clave, nombre in self._dibujados:
            d = math.hypot(px_ - x, py_ - y)
            if d <= radio and (mejor is None or d < mejor[0]):
                mejor = (d, (px_, py_, clave, nombre))
        return mejor[1] if mejor else None

    def _pulsar(self, e):
        ctl = self._control_en(e.x, e.y) if self.centro is not None else None
        if ctl == "mas":
            self._zoom_en(self.winfo_width() / 2, self.winfo_height() / 2, 1)
        elif ctl == "menos":
            self._zoom_en(self.winfo_width() / 2, self.winfo_height() / 2, -1)
        elif ctl == "centrar":
            self.recentrar()
        if ctl or self.centro is None:
            self._arrastre = None
            return
        self._arrastre = (e.x, e.y, a_pixeles(self.centro[0], self.centro[1], self.z), False)

    def _mover(self, e):
        if not self._arrastre:
            return
        x, y, (cx, cy), movido = self._arrastre
        dx, dy = e.x - x, e.y - y
        if not movido and math.hypot(dx, dy) < self.ui.e.px(3):
            return
        self._arrastre = (x, y, (cx, cy), True)
        self.configure(cursor="fleur")
        self.centro = a_coordenadas(cx - dx, cy - dy, self.z)
        self.redibujar()

    def _soltar(self, e):
        arrastre, self._arrastre = self._arrastre, None
        self.configure(cursor="")
        if arrastre and not arrastre[3] and self.al_pulsar:
            p = self._punto_en(e.x, e.y)
            if p:
                self.al_pulsar(p[2])

    def _doble(self, e):
        if self.centro is not None and not self._control_en(e.x, e.y):
            self._zoom_en(e.x, e.y, 1)

    def _zoom_en(self, x, y, d):
        if self.centro is None:
            return
        z = max(ZOOM_MIN, min(ZOOM_MAX, self.z + d))
        if z == self.z:
            return
        x0, y0 = self._origen()
        lat, lon = a_coordenadas(x0 + x, y0 + y, self.z)       # lo que está bajo el ratón se queda en su sitio
        px_, py_ = a_pixeles(lat, lon, z)
        self.z = z
        self.centro = a_coordenadas(px_ - x + self.winfo_width() / 2, py_ - y + self.winfo_height() / 2, z)
        self.redibujar()

    def _movimiento(self, e):
        if self._arrastre:
            return
        p = self._punto_en(e.x, e.y) if self.al_pulsar else None
        self._poner_sobre(p)
        ctl = self._control_en(e.x, e.y) if self.centro is not None else None
        self.configure(cursor="hand2" if (p or ctl) else "")

    def _poner_sobre(self, p):
        if (p and p[2]) != (self._sobre and self._sobre[2]):
            self._sobre = p
            self.redibujar()
