# -*- coding: utf-8 -*-
"""
Centros educativos de España — ventana de consulta y descarga.

- La lista de toda España viene incluida con el programa y se abre al instante.
- Mientras el programa está abierto, la lista se actualiza sola en segundo plano una vez por semana:
  avisa al empezar y, al terminar, dice cuántos centros nuevos hay y cuántos emails y teléfonos se han añadido.
- Comunidades a la izquierda; filtros por provincia, municipio, tipo de centro, titularidad y texto; ficha del
  centro a la derecha (copiar o escribir el email, llamar, abrir la web y un mapa con su situación, que se puede
  ampliar con los demás centros del municipio); marca centros y descárgalos a Excel.

La lista actualizada se guarda en la carpeta del usuario (%LOCALAPPDATA%\\CentrosEducativos).
Usa centros_educativos.py (misma carpeta) para descargar los datos de las fuentes oficiales, interfaz.py para
las piezas de la ventana y mapa.py para el mapa (plano del IGN y direcciones de CartoCiudad).
"""

import ctypes
import os
import queue
import re
import shutil
import sys
import threading
import time
import traceback
import urllib.parse
import webbrowser
from collections import Counter
from datetime import datetime, timedelta

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

if sys.platform == "win32":  # nitidez en pantallas con escalado
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

import tkinter as tk  # noqa: E402
from tkinter import filedialog, messagebox, ttk  # noqa: E402

import centros_educativos as ce  # noqa: E402
import interfaz as ui  # noqa: E402
import mapa  # noqa: E402
from interfaz import C, etiqueta  # noqa: E402

ICONO_ICO = os.path.join(APP_DIR, "icono.ico")
ICONO_PNG = os.path.join(APP_DIR, "icono.png")
IMG_DIR = os.path.join(APP_DIR, "img")
NOMBRE_ACCESO = "Centros educativos"
DESCRIPCION_ACCESO = "Centros educativos de España: teléfono y email"
# «Centros educativos.exe» (junto a la carpeta «programa») es el pythonw.exe oficial, firmado por la Python Software
# Foundation, con otro nombre; al abrirlo sin argumentos, sitecustomize.py arranca esta ventana.
LANZADOR = os.path.join(os.path.dirname(APP_DIR), "Centros educativos.exe")
FOLDERID_DOCUMENTOS = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
FOLDERID_ESCRITORIO = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
RUTA_INICIAL = os.path.join(APP_DIR, "datos", "centros_inicial.json.gz")   # lista incluida con el programa
DIAS_ACTUALIZACION = 7     # la lista se actualiza sola una vez por semana
DIAS_FICHAS = 90           # cada cuánto se vuelve a consultar en el Ministerio un centro al que le falta el email
MIN_PENDIENTES = 50        # centros por completar que justifican seguir una actualización interrumpida
HILOS_FICHAS, PAUSA_FICHAS = 4, 0.2
REVISION_MS = 60 * 60 * 1000   # cada hora se comprueba si toca actualizar (o reintentar si falló)
ANCHO_LATERAL, ANCHO_FICHA = 272, 352
ANCHO_CON_FICHA = 1180     # por debajo, la ficha se abre en una ventana aparte (doble clic)
TITULARIDADES = ["Pública", "Privada concertada", "Privada"]
ICONO_AVISO = {"ok": "ok-verde-18", "act": "actualizar-ambar-18", "error": "alerta-rojo-18"}
CERRAR_AVISO = {"ok": "cerrar-verde_osc-14", "act": "cerrar-ambar_osc-14", "error": "cerrar-rojo_osc-14"}
MAX_PUNTOS_MAPA = 400      # centros del municipio que se sitúan en el mapa ampliado
PRECISION = {"calle": "Situado en su calle: el número exacto no aparece en el callejero.",
             "aprox": "Situación aproximada: no se ha encontrado la dirección exacta y se muestra el municipio."}


def miles(n):
    return f"{n:,}".replace(",", ".")


def plural(n, singular, plural_):
    return f"{miles(n)} {singular if n == 1 else plural_}"


def porcentaje(parte, total):
    return f"{round(100 * parte / total)} %" if total else "0 %"


def fecha_iso(texto):
    try:
        return datetime.fromisoformat(str(texto))
    except (TypeError, ValueError):
        return None


def tiempo_restante(seg):
    m = round(seg / 60)
    if seg < 60:
        return "queda menos de un minuto"
    if m <= 1:
        return "queda 1 minuto aprox."
    if m < 60:
        return f"quedan unos {m} minutos"
    h, m = divmod(m, 60)
    return f"quedan aprox. {h} h {m:02d} min"


class GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort), ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8)]

    def __init__(self, texto):
        import uuid
        u = uuid.UUID(texto)
        super().__init__(u.fields[0], u.fields[1], u.fields[2], (ctypes.c_ubyte * 8)(*u.bytes[8:]))


def carpeta_conocida(folderid):
    """Ruta de una carpeta especial de Windows (Documentos, Escritorio…), también si está en OneDrive."""
    if sys.platform != "win32":
        return ""
    try:
        g = GUID(folderid)
        p = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(g), 0, None, ctypes.byref(p)) == 0:
            ruta = p.value
            ctypes.windll.ole32.CoTaskMemFree(p)
            if ruta and os.path.isdir(ruta):
                return ruta
    except Exception:
        pass
    return ""


def documentos():
    """Carpeta Documentos del usuario (también si está en OneDrive)."""
    ruta = carpeta_conocida(FOLDERID_DOCUMENTOS)
    if ruta:
        return ruta
    d = os.path.join(os.path.expanduser("~"), "Documents")
    return d if os.path.isdir(d) else os.path.expanduser("~")


def abrir(ruta):
    try:
        if hasattr(os, "startfile"):
            os.startfile(ruta)
        else:
            webbrowser.open("file://" + os.path.abspath(ruta))
    except OSError:
        pass


def carpeta_usuario():
    """Carpeta escribible del usuario para la lista actualizada y la caché de descargas."""
    for d in (os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "CentrosEducativos"),
              os.path.join(APP_DIR, "datos_usuario")):
        try:
            os.makedirs(d, exist_ok=True)
            prueba = os.path.join(d, ".escritura")
            with open(prueba, "w") as f:
                f.write("ok")
            os.remove(prueba)
            return d
        except OSError:
            continue
    import tempfile
    return tempfile.mkdtemp(prefix="centros_")


def preparar(rows):
    """Clave estable de cada centro (las marcas se conservan aunque la lista se actualice) y datos que la ventana usa
    a menudo: tipo breve, categoría, teléfono legible e índice de búsqueda. Se calcula en el hilo que carga la lista."""
    for r in rows:
        r["_k"] = ce.clave_centro(r)
        ce.categoria_centro(r)
        r["_corto"] = ce.tipo_corto(r)
        r["_corto_nivel"] = ce.tipo_corto(r, False)
        r["_tel"] = ce.telefono_legible(r["tel"])
    ce.filtrar(rows, "x")  # prepara el índice de búsqueda de todas las filas
    return rows


def _carpetas_datos():
    """Primero el perfil del usuario (la marca es por usuario y equipo); si no, la carpeta del programa."""
    return [os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "CentrosEducativos"), APP_DIR]


def _hay_marca_acceso():
    return any(os.path.exists(os.path.join(d, ".acceso_directo_creado")) for d in _carpetas_datos())


def _guardar_marca_acceso():
    for d in _carpetas_datos():
        try:
            os.makedirs(d, exist_ok=True)
            with open(os.path.join(d, ".acceso_directo_creado"), "w", encoding="utf-8") as f:
                f.write(datetime.now().isoformat())
            return
        except OSError:
            continue


def _metodo_com(objeto, indice, restype, *argtypes):
    """Función número 'indice' de la tabla de métodos (vtable) de un objeto COM."""
    tabla = ctypes.cast(objeto, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    return ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)(tabla[indice])


def crear_lnk(ruta_lnk, destino, argumentos="", carpeta="", icono="", descripcion=""):
    """Crea o actualiza un acceso directo (.lnk) con la API de Windows (IShellLinkW + IPersistFile).
    No usa PowerShell: con el control de aplicaciones activo, PowerShell puede no dejar crear objetos COM."""
    from ctypes import HRESULT, POINTER, byref, c_int, c_ulong, c_void_p, c_wchar_p
    ole32 = ctypes.OleDLL("ole32")
    ole32.CoCreateInstance.argtypes = [POINTER(GUID), c_void_p, ctypes.c_uint, POINTER(GUID), POINTER(c_void_p)]
    ole32.CoUninitialize.restype = None  # devuelve void (OleDLL lo leería como código de error)
    try:
        ole32.CoInitialize(None)   # S_OK o S_FALSE (ya iniciado en este hilo): hay que emparejarlo con CoUninitialize
        iniciado = True
    except OSError:                # ya iniciado en otro modo en este hilo: se puede usar igual
        iniciado = False
    try:
        enlace = c_void_p()
        ole32.CoCreateInstance(byref(GUID("{00021401-0000-0000-C000-000000000046}")), None, 1,   # CLSID_ShellLink
                               byref(GUID("{000214F9-0000-0000-C000-000000000046}")), byref(enlace))  # IShellLinkW
        try:
            _metodo_com(enlace, 20, HRESULT, c_wchar_p)(enlace, destino)                # SetPath
            _metodo_com(enlace, 11, HRESULT, c_wchar_p)(enlace, argumentos)             # SetArguments
            _metodo_com(enlace, 9, HRESULT, c_wchar_p)(enlace, carpeta)                 # SetWorkingDirectory
            if icono:
                _metodo_com(enlace, 17, HRESULT, c_wchar_p, c_int)(enlace, icono, 0)   # SetIconLocation
            _metodo_com(enlace, 7, HRESULT, c_wchar_p)(enlace, descripcion)             # SetDescription
            fichero = c_void_p()
            _metodo_com(enlace, 0, HRESULT, POINTER(GUID), POINTER(c_void_p))(          # QueryInterface
                enlace, byref(GUID("{0000010B-0000-0000-C000-000000000046}")), byref(fichero))  # IPersistFile
            try:
                _metodo_com(fichero, 6, HRESULT, c_wchar_p, c_int)(fichero, ruta_lnk, 1)  # Save
            finally:
                _metodo_com(fichero, 2, c_ulong)(fichero)                                 # Release
        finally:
            _metodo_com(enlace, 2, c_ulong)(enlace)                                        # Release
    finally:
        if iniciado:
            ole32.CoUninitialize()


# Plan B si falla lo anterior. Sin comillas dobles dentro: así llega intacto a PowerShell como un único argumento.
PS_ACCESO = r"""
$ErrorActionPreference = 'Stop'
$s = (New-Object -ComObject WScript.Shell).CreateShortcut($env:CE_LNK)
$s.TargetPath = $env:CE_DESTINO
$s.Arguments = $env:CE_ARGS
$s.WorkingDirectory = $env:CE_DIR
$s.IconLocation = $env:CE_ICONO + ',0'
$s.Description = $env:CE_DESC
$s.Save()
'ok'
"""


def _lnk_powershell(lnk, destino, argumentos, carpeta, icono, descripcion):
    import subprocess
    env = dict(os.environ, CE_LNK=lnk, CE_DESTINO=destino, CE_ARGS=argumentos, CE_DIR=carpeta, CE_ICONO=icono,
               CE_DESC=descripcion)
    ps = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
    r = subprocess.run([ps if os.path.exists(ps) else "powershell", "-NoProfile", "-NonInteractive", "-Command", PS_ACCESO],
                       env=env, capture_output=True, encoding="utf-8", errors="replace", timeout=60,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if r.returncode != 0 or "ok" not in (r.stdout or ""):
        raise RuntimeError((r.stderr or r.stdout or f"código {r.returncode}").strip()[:300])


def acceso_directo(forzar=False):
    """Crea el acceso directo del escritorio la primera vez (o si forzar) y lo mantiene apuntando a esta carpeta.
    Devuelve 'creado', 'actualizado', 'omitido' o 'error: ...'."""
    if sys.platform != "win32":
        return "omitido"
    if os.path.exists(LANZADOR):
        destino, argumentos, carpeta = LANZADOR, "", os.path.dirname(LANZADOR)
    else:  # versiones anteriores: Python en programa\python y la ventana como argumento
        destino = os.path.join(APP_DIR, "python", "pythonw.exe")
        argumentos, carpeta = '"' + os.path.join(APP_DIR, "centros_gui.pyw") + '"', APP_DIR
        if not os.path.exists(destino):
            return "omitido"
    escritorio = carpeta_conocida(FOLDERID_ESCRITORIO) or os.path.join(os.path.expanduser("~"), "Desktop")
    lnk = os.path.join(escritorio, NOMBRE_ACCESO + ".lnk")
    existe = os.path.exists(lnk)
    if not existe and not (forzar or not _hay_marca_acceso()):
        return "omitido"  # se creó una vez y el usuario lo borró: no insistir
    datos = (lnk, destino, argumentos, carpeta, ICONO_ICO, DESCRIPCION_ACCESO)
    try:
        crear_lnk(*datos)
    except Exception as e1:
        try:
            _lnk_powershell(*datos)
        except Exception as e2:
            return f"error: {e1} / {e2}"
    _guardar_marca_acceso()
    return "actualizado" if existe else "creado"


def registrar_error(texto):
    """Guarda el detalle de un error en error.log (carpeta del programa o perfil del usuario)."""
    for d in (APP_DIR, os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "CentrosEducativos")):
        try:
            os.makedirs(d, exist_ok=True)
            ruta = os.path.join(d, "error.log")
            with open(ruta, "a", encoding="utf-8") as f:
                f.write(f"--- {datetime.now():%Y-%m-%d %H:%M:%S}\n{texto}\n")
            return ruta
        except OSError:
            continue
    return ""


class Ficha(tk.Frame):
    """Ficha del centro elegido: email, teléfono, web, dirección, código y botón para marcarlo."""

    def __init__(self, padre, app):
        super().__init__(padre, bg=C["blanco"])
        self.app = app
        self.r = None
        px, fu = app.e.px, app.fu
        blanco = C["blanco"]
        self.pie = tk.Frame(self, bg=blanco)
        self.pie.pack(side="bottom", fill="x", padx=px(22), pady=(px(8), px(20)))
        self.btn_marcar = ttk.Button(self.pie, text="Marcar este centro", style=app.estilo_marcar,
                                     command=self._alternar, cursor="hand2")
        self.btn_marcar.pack(fill="x")
        self.nota = etiqueta(self.pie, "Para enviar publicidad por email hace falta el consentimiento previo del "
                                       "destinatario (LSSI, art. 21).", fu.peque, C["suave"], blanco,
                             wraplength=px(ANCHO_FICHA - 48))
        self.nota.pack(fill="x", pady=(px(10), 0))
        self.marco = ui.MarcoDesplazable(self, blanco, app.estilo_barra)
        self.marco.pack(fill="both", expand=True)
        c = tk.Frame(self.marco.contenido, bg=blanco)
        c.pack(fill="both", expand=True, padx=px(22), pady=(px(24), px(8)))
        c.columnconfigure(0, weight=1)
        self.c = c
        etiqueta(c, ui.espaciado("FICHA DEL CENTRO"), fu.caps, C["suave"], blanco).grid(row=0, column=0, sticky="w")
        self.vacio = etiqueta(c, "Haz clic en un centro de la lista para ver aquí su email, teléfono, web y "
                                 "dirección.", fu.normal, C["suave"], blanco)
        self.vacio.grid(row=1, column=0, sticky="ew", pady=(px(16), 0))
        self.pastillas = tk.Frame(c, bg=blanco)
        self.pastillas.grid(row=2, column=0, sticky="w", pady=(px(16), 0))
        self.lbl_nombre = etiqueta(c, "", fu.nombre, C["texto"], blanco)
        self.lbl_nombre.grid(row=3, column=0, sticky="ew", pady=(px(12), 0))
        self.lbl_tipo = etiqueta(c, "", fu.peque, C["suave"], blanco)
        self.lbl_tipo.grid(row=4, column=0, sticky="ew", pady=(px(4), 0))
        # email destacado
        self.caja = ttk.Frame(c, style="Caja.TFrame", padding=px(14))
        self.caja.grid(row=5, column=0, sticky="ew", pady=(px(16), 0))
        etiqueta(self.caja, "Email", fu.peque, C["suave"], C["fondo"]).pack(anchor="w")
        self.lbl_email = etiqueta(self.caja, "", fu.email, C["texto"], C["fondo"])
        self.lbl_email.pack(fill="x", pady=(px(4), 0))
        self.btns_email = tk.Frame(self.caja, bg=C["fondo"])
        self.btns_email.columnconfigure((0, 1), weight=1, uniform="em")
        self.btn_copiar_email = ttk.Button(self.btns_email, text="  Copiar", image=app.iconos("copiar-texto-15"),
                                           compound="left", style=app.estilo_caja, cursor="hand2",
                                           command=lambda: self._copiar(self.r["email"], self.btn_copiar_email))
        self.btn_copiar_email.grid(row=0, column=0, sticky="ew", padx=(0, px(4)))
        self.btn_escribir = ttk.Button(self.btns_email, text="  Escribir", image=app.iconos("email-texto-15"),
                                       compound="left", style=app.estilo_caja, cursor="hand2", command=self._escribir)
        self.btn_escribir.grid(row=0, column=1, sticky="ew", padx=(px(4), 0))
        # resto de datos, cada uno con su icono
        self.f_tel, self.lbl_tel, (self.btn_copiar_tel, self.btn_llamar) = self._fila(
            6, "telefono-suave-18", [("Copiar", lambda: self._copiar(self.r["tel"], self.btn_copiar_tel)),
                                     ("Llamar", self._llamar)])
        self.f_web, self.lbl_web, (self.btn_web,) = self._fila(7, "web-suave-18", [("Abrir", self._abrir_web)])
        self.f_dir, self.lbl_dir, _ = self._fila(8, "mapa-suave-18", [])
        self._construir_mapa(c, 9)
        self.f_cod, self.lbl_cod, _ = self._fila(10, "codigo-suave-18", [])
        self.info = tk.Frame(c, bg=blanco)
        self.info.grid(row=11, column=0, sticky="ew", pady=(px(16), 0))
        self.lbl_info = etiqueta(self.info, "", fu.peque, C["suave"], blanco)
        self.lbl_info.pack(fill="x")
        self.lbl_revisar = etiqueta(self.info, "", fu.peque, "#92400e", blanco,
                                    image=app.iconos("alerta-ambar_osc-16"), compound="left")
        self.ocultables = (self.pastillas, self.lbl_nombre, self.lbl_tipo, self.caja, self.f_tel, self.f_web,
                           self.f_dir, self.f_mapa, self.f_cod, self.info)
        c.bind("<Configure>", lambda e: self._anchos(e.width))
        self.marco.rueda(self)
        self.mostrar(None)

    def _fila(self, fila, icono, botones):
        px = self.app.e.px
        f = tk.Frame(self.c, bg=C["blanco"])
        f.grid(row=fila, column=0, sticky="ew", pady=(px(16) if fila == 6 else px(12), 0))
        f.columnconfigure(1, weight=1)
        tk.Label(f, image=self.app.iconos(icono), bg=C["blanco"], bd=0).grid(row=0, column=0, sticky="nw",
                                                                            padx=(0, px(10)), pady=(px(1), 0))
        lbl = etiqueta(f, "", self.app.fu.normal, C["texto"], C["blanco"])
        lbl.grid(row=0, column=1, sticky="new")
        lista = []
        for n, (texto, accion) in enumerate(botones):
            b = ttk.Button(f, text=texto, style=self.app.estilo_ficha, command=accion, cursor="hand2")
            b.grid(row=0, column=2 + n, sticky="ne", padx=(px(8) if n == 0 else px(4), 0))
            b.texto = texto
            lista.append(b)
        return f, lbl, tuple(lista)

    def _construir_mapa(self, c, fila):
        """Mapa con la situación del centro (plano del IGN) y botones para ampliarlo o abrir Google Maps."""
        app, px, blanco = self.app, self.app.e.px, C["blanco"]
        self.f_mapa = tk.Frame(c, bg=blanco)
        self.f_mapa.grid(row=fila, column=0, sticky="ew", pady=(px(12), 0))
        self.visor = mapa.VisorMapa(self.f_mapa, app, app.servicio_mapa, alto=px(170), zoom=app.zoom_mapa,
                                    esquinas=True)
        self.visor.pack(fill="x")
        botones = tk.Frame(self.f_mapa, bg=blanco)
        botones.pack(fill="x", pady=(px(8), 0))
        botones.columnconfigure((0, 1), weight=1, uniform="mapa")
        self.btn_ampliar = ttk.Button(botones, text="  Ampliar mapa", compound="left", style=app.estilo_ficha,
                                      image=(app.iconos("ampliar-texto-15"), "disabled",
                                             app.iconos("ampliar-apagado-15")),
                                      command=self._ampliar, cursor="hand2")
        self.btn_ampliar.grid(row=0, column=0, sticky="ew", padx=(0, px(4)))
        ttk.Button(botones, text="  Google Maps", image=app.iconos("externo-texto-15"), compound="left",
                   style=app.estilo_ficha, command=self._mapa, cursor="hand2").grid(row=0, column=1, sticky="ew",
                                                                                     padx=(px(4), 0))
        self.lbl_precision = etiqueta(self.f_mapa, "", app.fu.peque, C["suave"], blanco)

    def _situar(self, r):
        if not (r["dir"] or r["mun"]):
            self.f_mapa.grid_remove()
            return
        res = self.app.servicio_mapa.situar(r, self._situado)
        if res is None:
            self.visor.mensaje("Buscando la dirección del centro…")
            self.lbl_precision.pack_forget()
            self.btn_ampliar.state(["disabled"])
        else:
            self._poner_situacion(res)

    def _situado(self, clave, res, error):
        if self.r is None or mapa.clave_geo(self.r) != clave:
            return  # mientras tanto se ha elegido otro centro
        if error:
            self.visor.mensaje("No se ha podido situar el centro. Comprueba la conexión a internet y vuelve a "
                               "abrir la ficha.")
            return
        self._poner_situacion(res)

    def _poner_situacion(self, res):
        if res.get("prec") == "no":
            self.visor.mensaje("No se ha encontrado la dirección de este centro en el callejero oficial.")
            self.lbl_precision.pack_forget()
            self.btn_ampliar.state(["disabled"])
            return
        self.visor.ver(res["lat"], res["lon"], self.app.zoom_mapa)
        self.visor.poner_pin(res["lat"], res["lon"])
        self.btn_ampliar.state(["!disabled"])
        texto = PRECISION.get(res.get("prec"), "")
        self.lbl_precision.configure(text=texto)
        if texto:
            self.lbl_precision.pack(fill="x", pady=(self.app.e.px(6), 0))
        else:
            self.lbl_precision.pack_forget()

    def _ampliar(self):
        if self.r is not None:
            self.app._ampliar_mapa(self.r)

    def _anchos(self, ancho):
        px = self.app.e.px
        for w in (self.vacio, self.lbl_nombre, self.lbl_tipo, self.lbl_info):
            w.configure(wraplength=max(px(120), ancho - px(4)))
        self.lbl_revisar.configure(wraplength=max(px(100), ancho - px(30)))
        self.lbl_precision.configure(wraplength=max(px(120), ancho - px(4)))
        self.lbl_email.configure(wraplength=max(px(120), ancho - px(32)))
        for f, lbl in ((self.f_tel, self.lbl_tel), (self.f_web, self.lbl_web), (self.f_dir, self.lbl_dir),
                       (self.f_cod, self.lbl_cod)):
            botones = sum(w.winfo_reqwidth() + px(8) for w in f.grid_slaves(row=0) if isinstance(w, ttk.Button)
                          and w.winfo_manager())
            lbl.configure(wraplength=max(px(90), ancho - px(30) - botones))

    def mostrar(self, r):
        self.r = r
        app, px = self.app, self.app.e.px
        if r is None:
            for w in self.ocultables:
                w.grid_remove()
            self.vacio.grid()
            self.btn_marcar.pack_forget()
            return
        self.vacio.grid_remove()
        for w in self.ocultables:
            w.grid()
        for w in self.pastillas.winfo_children():
            w.destroy()
        pastillas = [(f"{r['_corto_nivel']} · {r['nat']}" if r["nat"] else r["_corto_nivel"], "PillGris")]
        if app._es_nuevo(r):
            pastillas.append(("Centro nuevo", "PillVerde"))
        elif app._email_nuevo(r):
            pastillas.append(("Email nuevo", "PillAmbar"))
        for texto, estilo in pastillas:
            ttk.Label(self.pastillas, text=texto, style=f"{estilo}.TLabel").pack(side="left", padx=(0, px(6)))
        for w in self.pastillas.winfo_children():
            self.marco.rueda(w)
        self.lbl_nombre.configure(text=r["nombre"] or "(sin nombre)")
        self.lbl_tipo.configure(text=r["tipo"])
        if r["email"]:
            self.lbl_email.configure(text=r["email"], font=app.fu.email, fg=C["texto"])
            self.btns_email.pack(fill="x", pady=(px(10), 0))
        else:
            self.lbl_email.configure(text="Este centro no tiene email en las fuentes oficiales.", font=app.fu.normal,
                                     fg=C["suave"])
            self.btns_email.pack_forget()
        self._dato(self.lbl_tel, r["_tel"], "Sin teléfono", (self.btn_copiar_tel, self.btn_llamar))
        web = re.sub(r"^https?://", "", r["web"] or "", flags=re.I).rstrip("/")
        self._dato(self.lbl_web, web, "Sin web", (self.btn_web,))
        self._dato(self.lbl_dir, direccion(r), "Sin dirección", ())
        self._situar(r)
        self._dato(self.lbl_cod, f"Código de centro {r['code']}" if r["code"] else "", "Sin código", ())
        lineas = []
        alta = fecha_iso(r.get("alta"))
        if alta:
            lineas.append(f"En la lista desde el {alta:%d/%m/%Y}" + (" (centro nuevo)." if app._es_nuevo(r) else "."))
        alta_email = fecha_iso(r.get("email_alta"))
        if r["email"] and alta_email:
            lineas.append(f"Email añadido el {alta_email:%d/%m/%Y}.")
        if r.get("fuente"):
            lineas.append(texto_fuente(r["fuente"]))
        self.lbl_info.configure(text="\n".join(lineas))
        self.lbl_info.pack(fill="x") if lineas else self.lbl_info.pack_forget()
        if r.get("revisar"):
            self.lbl_revisar.configure(text="  " + r["revisar"])
            self.lbl_revisar.pack(fill="x", pady=(px(6) if lineas else 0, 0))
        else:
            self.lbl_revisar.pack_forget()
        self.marcado()
        self.btn_marcar.pack(fill="x", before=self.nota)
        self._anchos(self.c.winfo_width())
        self.marco.arriba()

    def _dato(self, lbl, texto, sin, botones):
        lbl.configure(text=texto or sin, fg=C["texto"] if texto else C["suave"])
        for b in botones:
            b.grid() if texto else b.grid_remove()

    def marcado(self):
        if self.r is None:
            return
        if self.r["_k"] in self.app.marcados:
            self.btn_marcar.configure(text="  Marcado para descargar", image=self.app.iconos("check-marcado-16"),
                                      compound="left", style=self.app.estilo_marcado)
        else:
            self.btn_marcar.configure(text="Marcar este centro", image="", style=self.app.estilo_marcar)

    # -- acciones
    def _alternar(self):
        if self.r is not None:
            self.app._alternar_marca(self.r["_k"])

    def _copiar(self, texto, boton):
        self.app.clipboard_clear()
        self.app.clipboard_append(texto)
        original = getattr(boton, "texto", None) or boton.cget("text")
        boton.texto = original
        boton.configure(text=("  " if original.startswith("  ") else "") + "Copiado")
        boton.after(1500, lambda: boton.winfo_exists() and boton.configure(text=original))

    def _escribir(self):
        webbrowser.open("mailto:" + self.r["email"])

    def _abrir_web(self):
        web = self.r["web"]
        webbrowser.open(web if re.match(r"https?://", web, re.I) else "http://" + web)

    def _llamar(self):
        numero = re.sub(r"\D", "", self.r["tel"])
        enlace = "tel:+34" + numero if len(numero) == 9 else "tel:" + numero
        try:
            if sys.platform == "win32":
                os.startfile(enlace)
            elif not webbrowser.open(enlace):
                raise OSError("sin aplicación")
        except OSError:
            self.app.clipboard_clear()
            self.app.clipboard_append(self.r["_tel"])
            messagebox.showinfo("Llamar", "Este equipo no tiene ninguna aplicación para hacer llamadas (por ejemplo, "
                                          "Enlace Móvil o Teams).\n\nHe copiado el número para que lo marques desde el "
                                          f"teléfono: {self.r['_tel']}", parent=self.winfo_toplevel())

    def _mapa(self):
        if self.r is not None:
            abrir_google_maps(self.r)


class VentanaMapa(tk.Toplevel):
    """Mapa grande del centro elegido con los demás centros de la lista en su municipio (clic: ver su ficha)."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.r = None
        self.otros = {}           # clave de coordenadas → fila
        self.pendientes = set()
        self.sin_red = False
        self._pintar_pendiente = None
        px, fu, blanco = app.e.px, app.fu, C["blanco"]
        self.title("Mapa del centro")
        self.configure(bg=blanco)
        self.transient(app)
        ancho = min(px(980), self.winfo_screenwidth() - px(80))
        alto = min(px(680), self.winfo_screenheight() - px(120))
        x = app.winfo_rootx() + max(0, (app.winfo_width() - ancho) // 2)
        y = app.winfo_rooty() + max(0, (app.winfo_height() - alto) // 3)
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.minsize(px(520), px(380))
        try:
            if sys.platform == "win32" and os.path.exists(ICONO_ICO):
                self.iconbitmap(ICONO_ICO)
        except tk.TclError:
            pass
        cab = tk.Frame(self, bg=blanco)
        cab.pack(fill="x", padx=px(20), pady=(px(14), px(12)))
        self.lbl_nombre = etiqueta(cab, "", fu.email, C["texto"], blanco)
        self.lbl_nombre.pack(fill="x")
        self.lbl_dir = etiqueta(cab, "", fu.peque, C["suave"], blanco)
        self.lbl_dir.pack(fill="x", pady=(px(2), 0))
        tk.Frame(self, bg=C["borde"], height=1).pack(fill="x")
        pie = tk.Frame(self, bg=blanco)
        pie.pack(side="bottom", fill="x", padx=px(20), pady=(px(10), px(14)))
        tk.Frame(self, bg=C["borde"], height=1).pack(side="bottom", fill="x")
        self.visor = mapa.VisorMapa(self, app, app.servicio_mapa, alto=px(400), zoom=app.zoom_mapa - 1, rueda=True,
                                    atribucion=mapa.ATRIBUCION, al_pulsar=self._pulsar_punto)
        self.visor.pack(fill="both", expand=True)
        izq = tk.Frame(pie, bg=blanco)
        izq.pack(side="left", fill="x", expand=True)
        leyenda = tk.Frame(izq, bg=blanco)
        leyenda.pack(anchor="w")
        for img, texto in (("punto-ip", "Infantil y primaria"), ("punto-sec", "Secundaria, FP y colegios"),
                           ("punto-otros", "Otros"), ("punto-otros-m", "Marcado")):
            tk.Label(leyenda, image=app.iconos(img), bg=blanco, bd=0).pack(side="left")
            etiqueta(leyenda, texto, fu.peque, C["sec"], blanco).pack(side="left", padx=(px(4), px(12)))
        self.lbl_estado = etiqueta(izq, "", fu.peque, C["suave"], blanco)
        self.lbl_estado.pack(anchor="w", pady=(px(4), 0))
        ttk.Button(pie, text="Cerrar", style=app.estilo_primario, command=self.destroy, cursor="hand2").pack(
            side="right")
        ttk.Button(pie, text="  Abrir en Google Maps", image=app.iconos("externo-texto-15"), compound="left",
                   style=app.estilo_dialogo, command=self._google, cursor="hand2").pack(side="right",
                                                                                         padx=(0, px(8)))
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Destroy>", lambda e: app.servicio_mapa.cancelar_lote() if e.widget is self else None)
        cab.bind("<Configure>", lambda e: self.lbl_nombre.configure(wraplength=max(px(200), e.width)))

    def mostrar(self, r):
        primera = self.r is None
        self.r = r
        self.title(f"Mapa — {r['nombre']}")
        self.lbl_nombre.configure(text=r["nombre"] or "(sin nombre)")
        self.lbl_dir.configure(text=direccion(r).replace("\n", " · "))
        res = self.app.servicio_mapa.situar(r, self._situado)
        if res is None:
            self.visor.mensaje("Buscando la dirección del centro…", limpiar=primera)
        else:
            self._poner(res, primera)
        self._preparar_otros()

    def _situado(self, clave, res, error):
        if self.r is None or mapa.clave_geo(self.r) != clave or not self.winfo_exists():
            return
        if error:
            self.visor.mensaje("No se ha podido situar el centro. Comprueba la conexión a internet.")
        else:
            self._poner(res, True)

    def _poner(self, res, centrar):
        if res.get("prec") == "no":
            self.visor.mensaje("No se ha encontrado la dirección de este centro en el callejero oficial.",
                               limpiar=False)
            return
        self.visor.texto = ""
        self.visor.poner_pin(res["lat"], res["lon"])
        if centrar or self.visor.centro is None:
            self.visor.ver(res["lat"], res["lon"])
        else:
            self.visor.recentrar()
        self._pintar_puntos()

    def _preparar_otros(self):
        """Los demás centros de la lista en el mismo municipio: se sitúan los que falten (en segundo plano)."""
        app, r = self.app, self.r
        candidatos = [app.rows[i] for i in app.vista if app.rows[i]["mun"] == r["mun"]
                      and app.rows[i]["prov"] == r["prov"] and app.rows[i]["_k"] != r["_k"]]
        candidatos.sort(key=lambda o: o["cp"] != r["cp"])  # primero los del mismo código postal
        app.servicio_mapa.cancelar_lote()   # lo que quedara de otro municipio o de otra lista
        self.otros = {mapa.clave_geo(o): o for o in candidatos[:MAX_PUNTOS_MAPA]}
        self.pendientes = set()
        self.sin_red = False
        for clave, o in self.otros.items():
            if app.servicio_mapa.conocida(o) is None:
                self.pendientes.add(clave)
                app.servicio_mapa.situar(o, self._situado_otro, prioridad=1)
        self._pintar_puntos()

    def _situado_otro(self, clave, res, error):
        if clave not in self.otros or not self.winfo_exists():
            return
        self.pendientes.discard(clave)
        if error:
            self.sin_red = True
        if self._pintar_pendiente is None:
            self._pintar_pendiente = self.after(250, self._pintar_puntos)

    def _pintar_puntos(self):
        self._pintar_pendiente = None
        if not self.winfo_exists():
            return
        app = self.app
        puntos = []
        sin_situar = 0
        for clave, o in self.otros.items():
            res = app.servicio_mapa.conocida(o)
            if res is None:            # aún buscándose, o sin conexión
                continue
            if res.get("prec") == "no":
                sin_situar += 1
                continue
            cat = o.get("_cat") or ce.categoria_centro(o)
            grupo = "ip" if cat in ("infantil", "primaria") else "sec" if cat in ("secundaria", "colegio", "fp") \
                else "otros"
            img = app.iconos(f"punto-{grupo}" + ("-m" if o["_k"] in app.marcados else ""))
            puntos.append((res["lat"], res["lon"], img, o["_k"], o["nombre"]))
        self.visor.poner_puntos(puntos)
        mun = self.r["mun"] or "el municipio"
        if self.pendientes:
            hechos = len(self.otros) - len(self.pendientes)
            texto = f"Situando en el mapa los centros de la lista en {mun}: {miles(hechos)} de {miles(len(self.otros))}…"
        elif self.otros:
            texto = (f"{plural(len(puntos), 'centro', 'centros')} más de la lista en {mun}. Haz clic en un punto "
                     "para ver su ficha.")
            if sin_situar:
                texto += f" {plural(sin_situar, 'centro no tiene', 'centros no tienen')} dirección en el callejero."
        else:
            texto = f"No hay más centros de la lista en {mun}."
        if self.sin_red:
            texto += " Sin conexión: faltan algunos centros por situar."
        self.lbl_estado.configure(text=texto)

    def _pulsar_punto(self, clave_centro):
        self.app._elegir_centro(clave_centro)

    def _google(self):
        if self.r is not None:
            abrir_google_maps(self.r)


def abrir_google_maps(r):
    """Busca el centro en Google Maps (nombre y dirección) en el navegador."""
    nombre = r["nombre"]
    if r["_corto_nivel"] not in ("Otro",) and ce.norm(r["_corto_nivel"]) not in ce.norm(nombre):
        nombre = f"{r['_corto_nivel']} {nombre}"
    partes = [nombre, r["dir"], " ".join(x for x in (r["cp"], r["loc"]) if x)]
    if r["mun"] and ce.norm(r["mun"]) != ce.norm(r["loc"]):
        partes.append(r["mun"])
    partes += [r["prov"], "España"]
    consulta = ", ".join(p for p in partes if p)
    webbrowser.open("https://www.google.com/maps/search/?api=1&query=" + urllib.parse.quote(consulta))


def texto_fuente(fuente):
    """«Teléfono: Registro Estatal (listado) · Email: directorio autonómico» → frase legible."""
    if " · " in fuente:
        partes = []
        for trozo in fuente.split(" · "):
            que, _, de = trozo.partition(": ")
            partes.append(f"{que.lower()} del {de}" if de else trozo)
        return "Fuente: " + "; ".join(partes) + "."
    return "Fuente: " + (fuente[:1].lower() + fuente[1:] if fuente.startswith("Directorio") else fuente) + "."


def direccion(r):
    linea = ", ".join(x for x in (r["dir"], " ".join(x for x in (r["cp"], r["loc"]) if x)) if x)
    if r["mun"] and ce.norm(r["mun"]) != ce.norm(r["loc"]):
        linea += f" ({r['mun']})"
    if r["prov"] and ce.norm(r["prov"]) not in ce.norm(linea):
        linea += ("\n" if linea else "") + r["prov"]
    return linea


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.title("Centros educativos de España — teléfono y email")
        self.root = self
        self.e = ui.Escala(self)
        self.fu = ui.Fuentes(self)
        self.iconos = ui.Iconos(self, IMG_DIR, self.e)
        self.configure(bg=C["fondo"])
        ancho, alto = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{min(self.e.px(1400), ancho - 60)}x{min(self.e.px(860), alto - 100)}+20+20")
        self.minsize(min(self.e.px(900), ancho - 40), min(self.e.px(600), alto - 80))
        self.q = queue.Queue()
        self.cancel_act = threading.Event()
        self.hilo_act = None
        self.actualizando = False
        self.guardando = False
        self.cerrando = False
        self.destruida = False
        self.cargada = False          # ya se ha intentado leer la lista guardada
        self.rows, self.info, self.meta = [], {}, {}
        self.por_clave = {}
        self.vista = []
        self.marcados = set()         # claves de centro (se conservan al actualizar la lista)
        self.activa = None            # centro cuya ficha se ve
        self.orden = (None, False)
        self.log_lineas = []
        self.ventana_log = None
        self.texto_log = None
        self.ventana_ficha = None
        self.desplegable = None
        self._cierre_desplegable = (None, 0.0)
        self.pendientes = 0           # centros sin email/teléfono aún sin consultar en el Ministerio
        self.no_antes_de = 0.0        # tras un fallo, no reintentar antes de esta hora
        self.nuevos_desde = ""        # lo que apareció desde esta fecha cuenta como «nuevo»
        self.n_nuevos = 0
        self.n_emails_nuevos = 0
        self.cuenta_ccaa = {}
        self.provs_ccaa = {}
        self.sel_ccaa = set()
        # filtros
        self.provincia = ""
        self.municipios = set()       # (provincia, municipio)
        self.categorias = set()
        self.titularidad = ""
        self.solo_email = False
        self.solo_nuevos = False
        self.solo_emails_nuevos = False
        self.ficha_visible = True
        self._filtro_pendiente = None
        self._estado_tmp = None
        self.carpeta = documentos()
        self.dir_usuario = carpeta_usuario()
        self.ruta_datos = os.path.join(self.dir_usuario, "centros.json.gz")
        self.cache = ce.carpeta_cache(os.path.join(self.dir_usuario, "cache"))
        self._migrar_cache()
        self.servicio_mapa = mapa.ServicioMapa(self.dir_usuario, lambda fn, *a: self.q.put(("tk", fn) + a),
                                               registrar=self._log)
        self.zoom_mapa = 16 if self.e.f < 1.75 else 17
        self.ventana_mapa = None
        try:  # mismo icono en la ventana y en la barra de tareas que en el acceso directo
            if sys.platform == "win32" and os.path.exists(ICONO_ICO):
                self.iconbitmap(default=ICONO_ICO)
            elif os.path.exists(ICONO_PNG):
                self._icono = tk.PhotoImage(file=ICONO_PNG)
                self.iconphoto(True, self._icono)
        except tk.TclError:
            pass
        self._estilos()
        self._construir()
        self._pintar_ccaa()
        self._filtrar()
        self.deiconify()
        if sys.platform == "win32":
            try:
                self.state("zoomed")
            except tk.TclError:
                pass
        self._gancho_log = lambda m: self.q.put(("log", m))
        ce._LOG_HOOKS.append(self._gancho_log)
        threading.Thread(target=self._tarea_cargar, daemon=True).start()
        threading.Thread(target=self.servicio_mapa.limpiar_viejas, daemon=True).start()
        self.after(100, self._sondear)
        self.after(1500, lambda: self._lanzar_acceso(False))
        self.after(REVISION_MS, self._revision_periodica)
        self.protocol("WM_DELETE_WINDOW", self._cerrar)

    def report_callback_exception(self, exc, val, tb):
        ruta = registrar_error("".join(traceback.format_exception(exc, val, tb)))
        messagebox.showerror("Error inesperado", f"{val}\n\nDetalle guardado en:\n{ruta}", parent=self)

    def _migrar_cache(self):
        """Aprovecha los emails ya consultados por la versión anterior (caché en la carpeta del programa)."""
        viejo = os.path.join(APP_DIR, "cache", "fichas_rcd.json")
        nuevo = os.path.join(self.cache, "fichas_rcd.json")
        try:
            if os.path.abspath(viejo) != os.path.abspath(nuevo) and os.path.exists(viejo) and not os.path.exists(nuevo):
                shutil.copy2(viejo, nuevo)
        except OSError:
            pass

    # ------------------------------------------------------------------ estilos
    def _estilos(self):
        es = self.estilos = ui.Estilos(self, self.e, self.fu)
        fu = self.fu
        blanco, fondo = C["blanco"], C["fondo"]
        prim = {"normal": (C["acento"], None), "active": (C["acento_hover"], None),
                "pressed": (C["acento_pulsado"], None), "disabled": (C["acento_apagado"], None)}
        fg_prim = [("disabled", "#fdf1ec")]
        self.estilo_primario = es.boton("Primario", blanco, 36, 8, prim, fu.semi, blanco, 14, fg_prim)
        self.estilo_primario_peq = es.boton("PrimarioPeq", blanco, 32, 6, prim, fu.semi, blanco, 14, fg_prim)
        self.estilo_marcar = es.boton("Marcar", blanco, 42, 8, prim, fu.semi, blanco, 14, fg_prim)
        marcado = {"normal": (C["acento_suave"], C["acento_borde"]), "active": (C["acento_suave2"], C["acento_borde"]),
                   "pressed": (C["acento_suave2"], C["acento"])}
        self.estilo_marcado = es.boton("Marcado", blanco, 42, 8, marcado, fu.semi, C["acento_texto"], 14)
        sec = {"normal": (blanco, C["borde2"]), "active": (C["gris"], C["borde2"]), "pressed": (C["gris2"], C["borde2"]),
               "disabled": (fondo, C["borde"])}
        fg_sec = [("disabled", C["apagado"])]
        self.estilo_secundario = es.boton("Secundario", blanco, 36, 8, sec, fu.normal, C["texto"], 12, fg_sec)
        self.estilo_ficha = es.boton("FichaBtn", blanco, 30, 6, sec, fu.peque, C["texto"], 10, fg_sec)
        sec_fondo = dict(sec, disabled=(C["gris"], C["borde"]))
        self.estilo_caja = es.boton("CajaBtn", fondo, 34, 6, sec_fondo, fu.normal, C["texto"], 10, fg_sec)
        self.estilo_dialogo = es.boton("DialogoBtn", blanco, 34, 6, sec, fu.normal, C["texto"], 14, fg_sec)
        fant = {"normal": (blanco, None), "active": (C["gris"], None), "pressed": (C["gris2"], None),
                "disabled": (blanco, None)}
        self.estilo_fantasma = es.boton("Fantasma", blanco, 32, 6, fant, fu.normal, C["texto"], 10, fg_sec)
        osc = {"normal": (C["lat2"], C["lat_borde2"]), "active": (C["lat_pulsado"], C["lat_borde2"]),
               "pressed": (C["lat_borde"], C["lat_borde2"]), "disabled": (C["lat2"], C["lat_borde"])}
        self.estilo_oscuro = es.boton("Oscuro", C["lat2"], 32, 6, osc, fu.normal, C["lat_claro"], 12,
                                      [("disabled", C["lat_tenue"])])
        chip = {"normal": (blanco, C["borde2"]), "active": (fondo, C["apagado"]), "pressed": (C["gris"], C["apagado"]),
                "disabled": (C["gris"], C["borde"])}
        self.estilo_chip = es.boton("Chip", fondo, 32, 16, chip, fu.normal, C["texto"], 12, [("disabled", C["tenue"])])
        chip_on = {"normal": (C["acento_suave"], C["acento_borde"]), "active": (C["acento_suave2"], C["acento_borde"]),
                   "pressed": (C["acento_suave2"], C["acento"]), "disabled": (C["gris"], C["borde"])}
        self.estilo_chip_on = es.boton("ChipOn", fondo, 32, 16, chip_on, fu.normal, C["acento_hover"], 12,
                                       [("disabled", C["tenue"])])
        limpiar = {"normal": (fondo, None), "active": (C["gris"], None), "pressed": (C["gris2"], None)}
        self.estilo_limpiar = es.boton("Limpiar", fondo, 32, 16, limpiar, fu.normal, C["acento"], 10)
        for estado, (bg, borde, texto, borde_btn, hover) in ui.AVISOS.items():
            es.tarjeta(f"Aviso{estado}", fondo, bg, borde, 8)
            es.boton(f"AvisoBtn{estado}", bg, 30, 6, {"normal": (bg, borde_btn), "active": (hover, borde_btn),
                                                     "pressed": (hover, texto)}, fu.semi, texto, 10)
            es.boton(f"AvisoX{estado}", bg, 30, 6, {"normal": (bg, None), "active": (hover, None),
                                                   "pressed": (hover, None)}, fu.normal, texto, 8)
        self.estilo_enlace = es.enlace("Enlace", blanco, fu.peque, fu.peque_sub, C["acento"], C["acento_hover"])
        self.estilo_enlace_osc = es.enlace("EnlaceOsc", C["lat2"], fu.peque, fu.peque_sub, C["acento_claro"], "#f3c9b6")
        self.estilo_enlace_lat = es.enlace("EnlaceLat", C["lat"], fu.peque, fu.peque_sub, C["acento_claro"], "#f3c9b6")
        self.estilo_buscar = es.entrada("Buscar", fondo, blanco, C["borde2"], C["acento"], 8, 40, C["texto"], 38, 34)
        self.estilo_buscar_osc = es.entrada("BuscarOsc", C["lat"], C["lat2"], C["lat_borde"], C["acento_claro"], 6, 34,
                                            C["lat_claro"], 34, 30)
        self.estilo_buscar_popup = es.entrada("BuscarPopup", blanco, blanco, C["borde2"], C["acento"], 6, 34,
                                              C["texto"], 34, 30)
        es.tarjeta("Tarjeta", fondo, blanco, C["borde"], 10)
        es.tarjeta("Caja", blanco, fondo, C["borde"], 10)
        es.tarjeta("LatCaja", C["lat"], C["lat2"], None, 8)
        es.pastilla("PillGris", blanco, C["gris"], C["sec"], fu.peque_semi)
        es.pastilla("PillAmbar", blanco, "#fef3c7", "#92400e", fu.peque_semi)
        es.pastilla("PillVerde", blanco, "#dcfce7", "#166534", fu.peque_semi)
        self.estilo_barra = es.barra("Fina", blanco, C["borde2"], C["apagado"])
        self.estilo_barra_osc = es.barra("FinaOsc", C["lat"], C["lat_borde2"], C["tenue"])
        self.estilo_progreso = es.progreso("Aviso", "#fde68a", "#d97706")

    # ------------------------------------------------------------------ ventana
    def _construir(self):
        px, fu = self.e.px, self.fu
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._construir_lateral()
        # ---- zona central
        centro = tk.Frame(self, bg=C["fondo"])
        centro.grid(row=0, column=1, sticky="nsew")
        cont = self.cont = tk.Frame(centro, bg=C["fondo"])
        cont.pack(fill="both", expand=True, padx=px(20), pady=(px(20), px(16)))
        cab = self.cabecera = tk.Frame(cont, bg=C["fondo"])
        cab.pack(fill="x")
        cab.columnconfigure(0, weight=1)
        tit = self.f_titulo = tk.Frame(cab, bg=C["fondo"])
        tit.grid(row=0, column=0, sticky="w")
        self.lbl_titulo = etiqueta(tit, "Toda España", fu.titulo)
        self.lbl_titulo.pack(side="left")
        self.lbl_total = etiqueta(tit, "", fu.normal, C["suave"])
        self.lbl_total.pack(side="left", padx=(px(12), 0), pady=(px(7), 0))
        self.caja_buscar = tk.Frame(cab, bg=C["fondo"], width=px(400))
        self.caja_buscar.pack_propagate(False)
        self.buscar = ui.CampoBusqueda(self.caja_buscar, self, self.estilo_buscar, C["fondo"], C["blanco"],
                                       "Nombre, municipio, tipo, email o código", self.iconos("buscar-suave-18"),
                                       self.iconos("cerrar-tenue-14"), C["apagado"],
                                       al_cambiar=lambda t: self._programar_filtro(), ancho=10)
        self.buscar.pack(fill="both", expand=True)
        self._buscar_abajo = None
        cab.bind("<Configure>", lambda e: self._ancho_buscador())
        self.buscar.entrada.bind("<Down>", lambda e: (self.tabla.lista.focus_set(), self.tabla.lista._tecla(1),
                                                      "break")[2])
        self.buscar.entrada.bind("<Escape>", lambda e: self.buscar.var.set(""))
        self.bind("<Control-f>", lambda e: (self.buscar.entrada.focus_set(),
                                           self.buscar.entrada.select_range(0, "end")))
        self._construir_aviso(cont)
        # ---- cifras
        kp = self.f_kpis = tk.Frame(cont, bg=C["fondo"])
        kp.pack(fill="x", pady=(px(14), 0))
        self.kpis = {}
        for n, (clave, titulo) in enumerate((("centros", "Centros"), ("email", "Con email"),
                                             ("tel", "Con teléfono"), ("marcados", "Marcados"))):
            kp.columnconfigure(n, weight=1, uniform="kpi")
            t = ttk.Frame(kp, style="Tarjeta.TFrame", padding=(px(14), px(10)))
            t.grid(row=0, column=n, sticky="nsew", padx=(0 if n == 0 else px(5), 0 if n == 3 else px(5)))
            lt = etiqueta(t, titulo, fu.peque, C["suave"], C["blanco"])
            lt.pack(anchor="w")
            fila = tk.Frame(t, bg=C["blanco"])
            fila.pack(anchor="w")
            v = etiqueta(fila, "—", fu.kpi, C["acento"] if clave == "marcados" else C["texto"], C["blanco"])
            v.pack(side="left")
            extra = etiqueta(fila, "", fu.kpi_extra, C["suave"], C["blanco"])
            extra.pack(side="left", anchor="s", padx=(px(6), 0), pady=(0, px(4)))
            self.kpis[clave] = (lt, v, extra)
        # ---- filtros
        self.flujo = ui.Flujo(cont, C["fondo"], px(8), px(8))
        self.flujo.pack(fill="x", pady=(px(14), 0))
        self.chip_prov = self._chip("Provincia", self._abrir_provincia, True)
        self.chip_mun = self._chip("Municipio", self._abrir_municipio, True)
        self.chip_tipo = self._chip("Tipo de centro", self._abrir_tipo, True)
        self.chip_tit = self._chip("Titularidad", self._abrir_titularidad, True)
        self.chip_email = self._chip("Con email", self._alternar_email)
        self.chip_nuevos = self._chip("Centros nuevos", lambda: self._alternar_novedad("centros"))
        self.chip_emails = self._chip("Emails nuevos", lambda: self._alternar_novedad("emails"))
        self.btn_limpiar = ttk.Button(self.flujo, text="Limpiar filtros", style=self.estilo_limpiar,
                                      command=self._limpiar_filtros, cursor="hand2")
        self.btn_limpiar.visible = False
        self.flujo.agregar(self.btn_limpiar)
        # ---- tabla
        tarjeta = ttk.Frame(cont, style="Tarjeta.TFrame", padding=(0, 0, 0, px(10)))
        tarjeta.pack(fill="both", expand=True, pady=(px(14), 0))
        pie = self.pie = tk.Frame(tarjeta, bg=C["blanco"])
        pie.pack(side="bottom", fill="x", padx=(px(12), px(12)), pady=(px(10), 0))
        tk.Frame(tarjeta, bg=C["borde"], height=1).pack(side="bottom", fill="x", padx=1)
        self.tabla = ui.TablaCentros(tarjeta, self, self._fila_tabla, self._seleccionar_fila, self._marcar_fila,
                                     self._doble_fila, self._ordenar, self._marcar_todos, self._estado_todos)
        self.tabla.pack(fill="both", expand=True)
        izq = self.pie_izq = tk.Frame(pie, bg=C["blanco"])
        der = self.pie_der = tk.Frame(pie, bg=C["blanco"])
        self.lbl_marcados = etiqueta(izq, "", fu.semi, C["texto"], C["blanco"])
        self.lbl_marcados.pack(side="left", padx=(px(6), px(6)))
        self.btn_copiar = ttk.Button(izq, text="  Copiar emails", compound="left", style=self.estilo_fantasma,
                                     image=(self.iconos("copiar-texto-16"), "disabled", self.iconos("copiar-apagado-16")),
                                     command=self._copiar_emails, cursor="hand2")
        self.btn_copiar.pack(side="left")
        self.btn_desmarcar = ttk.Button(izq, text="Desmarcar todos", style=self.estilo_fantasma,
                                        command=self._desmarcar_todos, cursor="hand2")
        self.btn_desmarcar.pack(side="left")
        self.lbl_estado = etiqueta(izq, "", fu.peque, C["suave"], C["blanco"])
        self.lbl_estado.pack(side="left", padx=(px(8), 0))
        self.btn_desc_marc = ttk.Button(der, text="  Descargar marcados", compound="left", style=self.estilo_primario,
                                        image=self.iconos("descargar-blanco-16"), command=lambda: self._descargar(True),
                                        cursor="hand2")
        self.btn_desc_marc.pack(side="right")
        self.btn_desc_lista = ttk.Button(der, text="Descargar lista", style=self.estilo_secundario,
                                         command=lambda: self._descargar(False), cursor="hand2")
        self.btn_desc_lista.pack(side="right", padx=(0, px(8)))
        pie.columnconfigure(1, weight=1)
        izq.grid(row=0, column=0, sticky="w")
        der.grid(row=0, column=2, sticky="e")
        self._pie_en_dos = False
        pie.bind("<Configure>", lambda e: self._colocar_pie(e.width))
        self.nota = etiqueta(cont, "", fu.peque, C["suave"], C["fondo"])
        self.nota.pack(fill="x", pady=(px(10), 0))
        cont.bind("<Configure>", lambda e: self.nota.configure(wraplength=max(px(200), e.width - px(4))), add="+")
        # ---- ficha (derecha)
        self.sep_ficha = tk.Frame(self, bg=C["borde"], width=1)
        self.sep_ficha.grid(row=0, column=2, sticky="ns")
        self.panel = tk.Frame(self, bg=C["blanco"], width=px(ANCHO_FICHA))
        self.panel.grid(row=0, column=3, sticky="ns")
        self.panel.pack_propagate(False)
        self.ficha = Ficha(self.panel, self)
        self.ficha.pack(fill="both", expand=True)
        self.bind("<Configure>", self._ajustar_ancho, add="+")
        self._actualizar_cuenta()

    def _construir_lateral(self):
        px, fu = self.e.px, self.fu
        lat = tk.Frame(self, bg=C["lat"], width=px(ANCHO_LATERAL))
        lat.grid(row=0, column=0, sticky="ns")
        lat.pack_propagate(False)
        interior = tk.Frame(lat, bg=C["lat"])
        interior.pack(fill="both", expand=True, padx=px(14), pady=px(20))
        marca = tk.Frame(interior, bg=C["lat"])
        marca.pack(fill="x", padx=px(4))
        tk.Label(marca, image=self.iconos("marca"), bg=C["lat"], bd=0).pack(side="left")
        tx = tk.Frame(marca, bg=C["lat"])
        tx.pack(side="left", padx=(px(12), 0))
        etiqueta(tx, "Centros educativos", fu.marca, C["lat_claro"], C["lat"]).pack(anchor="w")
        etiqueta(tx, "de España · teléfono y email", fu.peque, C["lat_tenue"], C["lat"]).pack(anchor="w")
        # tarjeta del estado de la lista (abajo)
        tarjeta = ttk.Frame(interior, style="LatCaja.TFrame", padding=px(12))
        tarjeta.pack(side="bottom", fill="x", pady=(px(16), 0))
        ancho = px(ANCHO_LATERAL - 28 - 24)
        self.lbl_lista = etiqueta(tarjeta, "Cargando…", fu.semi, C["lat_claro"], C["lat2"], wraplength=ancho)
        self.lbl_lista.pack(fill="x")
        self.lbl_prox = etiqueta(tarjeta, "Se actualiza sola cada semana.", fu.peque, C["lat_tenue"], C["lat2"],
                                 wraplength=ancho)
        self.lbl_prox.pack(fill="x", pady=(px(4), px(10)))
        self.btn_actualizar = ttk.Button(tarjeta, text="  Actualizar ahora", compound="left", style=self.estilo_oscuro,
                                         image=(self.iconos("actualizar-claro-16"), "disabled",
                                                self.iconos("actualizar-apagado-16")),
                                         command=self._actualizar_ahora, cursor="hand2")
        self.btn_actualizar.pack(fill="x")
        enlaces = tk.Frame(tarjeta, bg=C["lat2"])
        enlaces.pack(fill="x", pady=(px(10), 0))
        if sys.platform == "win32":
            ttk.Button(enlaces, text="Crear acceso directo en el escritorio", style=self.estilo_enlace_osc,
                       command=lambda: self._lanzar_acceso(True), cursor="hand2").pack(anchor="w")
        ttk.Button(enlaces, text="Ver el historial de actualizaciones", style=self.estilo_enlace_osc,
                   command=self._ver_log, cursor="hand2").pack(anchor="w", pady=(px(6), 0))
        # comunidades
        cab = tk.Frame(interior, bg=C["lat"])
        cab.pack(fill="x", pady=(px(18), px(10)))
        etiqueta(cab, ui.espaciado("COMUNIDADES"), fu.caps, C["lat_tenue"], C["lat"]).pack(side="left",
                                                                                          padx=(px(4), 0))
        self.btn_quitar_ccaa = ttk.Button(cab, text="Ver toda España", style=self.estilo_enlace_lat,
                                          command=lambda: self._elegir_ccaa(set()), cursor="hand2")
        self.buscar_ccaa = ui.CampoBusqueda(interior, self, self.estilo_buscar_osc, C["lat"], C["lat2"],
                                            "Buscar comunidad", self.iconos("buscar-lat-16"), self.iconos("cerrar-lat-14"),
                                            C["lat_tenue"], al_cambiar=lambda t: self._pintar_ccaa())
        self.buscar_ccaa.pack(fill="x")
        marco = tk.Frame(interior, bg=C["lat"])
        marco.pack(fill="both", expand=True, pady=(px(10), 0))
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)
        barra = ttk.Scrollbar(marco, orient="vertical", style=self.estilo_barra_osc)
        self.lista_ccaa = ui.ListaMarcable(marco, self, "oscuro", True, lambda sel, k: self._elegir_ccaa(sel), barra,
                                           alto=34)
        self.lista_ccaa.vacio = "Ninguna comunidad coincide"
        self.lista_ccaa.grid(row=0, column=0, sticky="nsew")
        barra.grid(row=0, column=1, sticky="ns", padx=(px(2), 0))
        barra.configure(command=self.lista_ccaa.yview)
        entrada = self.buscar_ccaa.entrada
        entrada.bind("<Down>", lambda e: (self.lista_ccaa.focus_set(), self.lista_ccaa._mover(1), "break")[2])
        entrada.bind("<Return>", lambda e: self.lista_ccaa.pulsar(0) if self.lista_ccaa.n == 1 else None)

    def _construir_aviso(self, cont):
        px, fu = self.e.px, self.fu
        b = self.banner = ttk.Frame(cont, style="Avisook.TFrame", padding=(px(12), px(8)))
        b.columnconfigure(1, weight=1)
        self.ban_icono = tk.Label(b, bd=0)
        self.ban_icono.grid(row=0, column=0, sticky="nw", padx=(0, px(10)), pady=(px(4), 0))
        self.ban_txt = ui.TextoRico(b, fu.normal, fu.semi, C["texto"], C["fondo"])
        self.ban_txt.grid(row=0, column=1, sticky="ew", pady=(px(4), px(4)))
        self.ban_barra = ttk.Progressbar(b, mode="indeterminate", length=px(160), style=self.estilo_progreso)
        self.ban_barra.grid(row=0, column=2, padx=(px(12), 0))
        self.ban_botones = tk.Frame(b, bg=C["fondo"])
        self.ban_botones.grid(row=0, column=3, sticky="e", padx=(px(8), 0))
        self.ban_btn_nuevos = ttk.Button(self.ban_botones, text="Ver centros nuevos", command=self._ver_nuevos,
                                         cursor="hand2")
        self.ban_btn_emails = ttk.Button(self.ban_botones, text="Ver emails nuevos", command=self._ver_emails_nuevos,
                                         cursor="hand2")
        self.ban_btn_detalles = ttk.Button(self.ban_botones, text="Ver detalles", command=self._ver_log, cursor="hand2")
        self.ban_btn_cerrar = ttk.Button(self.ban_botones, command=self._cerrar_aviso, cursor="hand2")
        self._aviso_partes = ("", "")

    def _chip(self, texto, accion, desplegable=False):
        b = ttk.Button(self.flujo, text=texto, style=self.estilo_chip, command=accion, cursor="hand2")
        b.desplegable = desplegable
        b.visible = True
        self._poner_chip(b, texto, False)
        self.flujo.agregar(b)
        return b

    def _poner_chip(self, b, texto, activo, habilitado=True):
        if b.desplegable:
            imagen = (self.iconos("chevron-acento_osc-12") if activo else self.iconos("chevron-texto-12"),
                      "disabled", self.iconos("chevron-tenue-12"))
            b.configure(text=texto + "  ", image=imagen, compound="right")
        elif activo:
            b.configure(text="  " + texto, image=self.iconos("check-acento_osc-14"), compound="left")
        else:
            b.configure(text=texto, image="", compound="none")
        b.configure(style=self.estilo_chip_on if activo else self.estilo_chip)
        b.state(["!disabled"] if habilitado else ["disabled"])

    def _ancho_buscador(self):
        """El buscador mide entre 260 y 420 px; si no cabe junto al título, baja a su propia línea."""
        px = self.e.px
        disponible = self.cabecera.winfo_width() - self.f_titulo.winfo_reqwidth() - px(16)
        abajo = disponible < px(260)
        if abajo != self._buscar_abajo:
            self._buscar_abajo = abajo
            if abajo:
                self.caja_buscar.grid(row=1, column=0, columnspan=2, sticky="ew", padx=0, pady=(px(12), 0))
            else:
                self.caja_buscar.grid(row=0, column=1, columnspan=1, sticky="e", padx=(px(16), 0), pady=0)
        alto = self.buscar.entrada.winfo_reqheight()
        self.caja_buscar.configure(height=alto, width=self.cabecera.winfo_width() if abajo else min(px(420), disponible))

    def _colocar_pie(self, ancho):
        dos = self.pie_izq.winfo_reqwidth() + self.pie_der.winfo_reqwidth() + self.e.px(16) > ancho
        if dos != self._pie_en_dos:
            self._pie_en_dos = dos
            if dos:
                self.pie_der.grid(row=1, column=0, columnspan=3, sticky="e", pady=(self.e.px(8), 0))
            else:
                self.pie_der.grid(row=0, column=2, columnspan=1, sticky="e", pady=0)

    def _ajustar_ancho(self, e):
        if e.widget is not self:
            return
        visible = e.width >= self.e.px(ANCHO_CON_FICHA)
        if visible != self.ficha_visible:
            self.ficha_visible = visible
            if visible:
                self.sep_ficha.grid()
                self.panel.grid()
            else:
                self.sep_ficha.grid_remove()
                self.panel.grid_remove()
            self._pintar_nota()

    # ------------------------------------------------------------------ comunidades
    def _orden_ccaa(self):
        return sorted(ce.CCAA, key=lambda k: ce.norm(ce.CCAA[k]["nombre"]))

    def _pintar_ccaa(self):
        filtro = ce.norm(self.buscar_ccaa.get())
        items = []
        for k in self._orden_ccaa():
            d = ce.CCAA[k]
            if filtro and filtro not in ce.norm(d["nombre"]) and not any(filtro in a for a in d["alias"]):
                continue
            items.append((k, d["nombre"], self.cuenta_ccaa.get(k) or None, False))
        self.lista_ccaa.poner(items, self.sel_ccaa)
        if self.sel_ccaa:
            self.btn_quitar_ccaa.pack(side="right")
        else:
            self.btn_quitar_ccaa.pack_forget()

    def _elegir_ccaa(self, seleccion):
        self.sel_ccaa = set(seleccion)
        if self.sel_ccaa:  # la provincia y los municipios elegidos tienen que estar en esas comunidades
            provs = set().union(*(self.provs_ccaa.get(k, set()) for k in self.sel_ccaa))
            if self.provincia not in provs:
                self.provincia = ""
            self.municipios = {m for m in self.municipios if m[0] in provs}
        self._pintar_ccaa()
        self._filtrar()

    # ------------------------------------------------------------------ lista guardada
    def _fichas(self, http=None):
        return ce.Fichas(os.path.join(self.cache, "fichas_rcd.json"), http, DIAS_FICHAS, offline=http is None)

    def _tarea_cargar(self):
        """Lee la última lista guardada (o la incluida con el programa) sin bloquear la ventana."""
        try:
            rows = info = meta = None
            origen = ""
            for ruta in (self.ruta_datos, RUTA_INICIAL):
                if not os.path.exists(ruta):
                    continue
                try:
                    rows, info, meta = ce.cargar_datos(ruta)
                    origen = ruta
                    break
                except Exception:
                    registrar_error(f"No se ha podido leer {ruta}:\n{traceback.format_exc()}")
                    self.q.put(("log", f"No se ha podido leer {ruta}; pruebo con otra copia."))
            if rows is None:
                self.q.put(("datos", [], {}, {}, 0, ""))
                return
            pendientes = 0
            try:
                fichas = self._fichas()
                if fichas.data:  # emails que se consultaron en una actualización que no llegó a terminar
                    ce.completar_con_fichas(rows, fichas)
                pendientes = len(ce.por_consultar(rows, fichas))
            except Exception:
                registrar_error(traceback.format_exc())
            self.q.put(("datos", preparar(rows), info, meta, pendientes, origen))
        except Exception:
            registrar_error(traceback.format_exc())
            self.q.put(("datos", [], {}, {}, 0, ""))

    def _datos_listos(self, rows, info, meta, pendientes, origen):
        self.cargada = True
        self.pendientes = pendientes
        if rows:
            self._log(f"Lista cargada: {miles(len(rows))} centros ({origen}).")
            self._poner_datos(rows, info, meta)
        else:
            self._log("No hay ninguna lista guardada en este equipo: hay que descargarla.")
            self._pintar_info_datos()
            self._estado_vacio()
        self.after(1500, self._comprobar_actualizacion)

    def _poner_datos(self, rows, info, meta):
        """Muestra una lista nueva conservando filtros, centros marcados y la ficha abierta."""
        self.rows, self.info, self.meta = rows, info, meta or {}
        self.por_clave = {r["_k"]: r for r in rows}
        self.marcados &= set(self.por_clave)
        self.cuenta_ccaa = Counter(r["key"] for r in rows)
        self.provs_ccaa = {}
        for r in rows:
            self.provs_ccaa.setdefault(r["key"], set()).add(r["prov"])
        self.nuevos_desde = self.meta.get("nuevos_desde") or ""
        self.n_nuevos = sum(1 for r in rows if self._es_nuevo(r))
        self.n_emails_nuevos = sum(1 for r in rows if self._email_nuevo(r))
        self.solo_nuevos = self.solo_nuevos and bool(self.n_nuevos)
        self.solo_emails_nuevos = self.solo_emails_nuevos and bool(self.n_emails_nuevos)
        ambito = self.sel_ccaa or set(self.provs_ccaa)
        provs = set().union(*(self.provs_ccaa.get(k, set()) for k in ambito)) if ambito else set()
        if self.provincia not in provs:
            self.provincia = ""
        if self.municipios:
            self.municipios &= {(r["prov"], r["mun"]) for r in rows}
        if self.activa not in self.por_clave:
            self.activa = None
        self._pintar_ccaa()
        self._pintar_info_datos()
        self._filtrar(arriba=False)
        self._mostrar_ficha(self.por_clave.get(self.activa))

    def _pintar_info_datos(self):
        n = len(self.rows)
        act = fecha_iso(self.meta.get("actualizado"))
        if not n:
            self.lbl_lista.configure(text="Todavía no hay lista de centros en este equipo." if self.cargada
                                     else "Cargando…")
            self.lbl_prox.configure(text="Se descarga de las fuentes oficiales (hace falta internet).")
            return
        self.lbl_lista.configure(text=f"Lista del {act:%d/%m/%Y}" if act else "Lista de centros")
        if act:
            prox = act + timedelta(days=DIAS_ACTUALIZACION)
            cuando = f"{prox:%d/%m/%Y}" if prox > datetime.now() else "hoy, con el programa abierto"
            self.lbl_prox.configure(text=f"Se actualiza sola cada semana. Próxima: {cuando}.")
        else:
            self.lbl_prox.configure(text="Se actualiza sola cada semana.")

    # ------------------------------------------------------------------ actualización en segundo plano
    def _motivo_actualizacion(self):
        if self.actualizando or self.cerrando or not self.cargada or time.time() < self.no_antes_de:
            return None
        if not self.rows:
            return "primera"
        act = fecha_iso(self.meta.get("actualizado"))
        if act is None or datetime.now() - act >= timedelta(days=DIAS_ACTUALIZACION):
            return "semanal"
        if not self.meta.get("fichas_completas"):
            if self.meta.get("origen") == "actualizada":
                return "completar"   # la última actualización se interrumpió al cerrar: se termina
            if self.pendientes >= MIN_PENDIENTES:
                return "completar"   # lista incluida con el programa: faltan emails por buscar
        return None

    def _comprobar_actualizacion(self):
        motivo = self._motivo_actualizacion()
        if motivo:
            self._iniciar_actualizacion(motivo)

    def _revision_periodica(self):
        if self.destruida:
            return
        self._comprobar_actualizacion()
        self.after(REVISION_MS, self._revision_periodica)

    def _actualizar_ahora(self):
        if self.actualizando or self.cerrando:
            return
        if not self.cargada:
            messagebox.showinfo("Actualizar", "Espera un momento: se está abriendo la lista.", parent=self)
            return
        self.no_antes_de = 0.0
        self._iniciar_actualizacion("manual")

    def _iniciar_actualizacion(self, motivo):
        self.actualizando = True
        self.cancel_act = threading.Event()
        self.btn_actualizar.configure(state="disabled", text="  Actualizando…")
        if motivo == "primera":
            titulo = "Descargando la lista de centros por primera vez…"
            detalle = "Tardará unos minutos. Te aviso cuando termine."
        else:
            titulo = "Actualizando la lista de centros en segundo plano…"
            detalle = "Puedes seguir trabajando con normalidad; te aviso cuando termine."
            if motivo == "completar":
                detalle += " Esta vez también se buscan los emails y teléfonos que faltan (puede tardar un rato)."
        self._aviso("act", titulo, detalle, barra=True)
        self._log(f"Empieza la actualización ({motivo}).")
        self._estado_vacio()
        self.hilo_act = threading.Thread(target=self._tarea_actualizar, args=(self.rows, self.info, dict(self.meta)),
                                         daemon=True)
        self.hilo_act.start()

    def _tarea_actualizar(self, antes, info_antes, meta_antes):
        """Hilo de fondo: descarga las fuentes oficiales, completa con el Ministerio, compara y guarda."""
        cancel = self.cancel_act
        inicio = datetime.now().isoformat(timespec="seconds")
        try:
            nuevas, info_n, _res = ce.obtener_centros(list(ce.CCAA), sin_fichas=True, cache=self.cache, cancel=cancel,
                                                      on_fase=lambda t: self.q.put(("act_fase", t)))
            if cancel.is_set():
                self.q.put(("act_cancelada",))
                return
            filas, info, conservadas, pares = ce.fusionar(antes, info_antes, nuevas, info_n, inicio)
            if not {r["key"] for r in nuevas} - set(conservadas):
                self.q.put(("act_error", "No se ha podido descargar ninguna comunidad (¿hay conexión a internet?)."))
                return
            if not antes:  # primera descarga: se enseña ya lo que hay mientras se completan los emails
                self.q.put(("act_parcial", preparar([dict(r) for r in filas]), dict(info)))
            fichas = self._fichas(ce.Http())
            ce.completar_con_fichas(filas, fichas)
            todo = ce.por_consultar(filas, fichas)
            if todo:
                self.q.put(("act_fase", f"Buscando en el Ministerio de Educación los emails y teléfonos que faltan "
                                        f"({plural(len(todo), 'centro', 'centros')})…"))
                fichas.run(todo, HILOS_FICHAS, PAUSA_FICHAS, cancel=cancel,
                           on_progress=lambda h, t, eta: self.q.put(("act_prog", h, t, eta)))
                ce.completar_con_fichas(filas, fichas)
            ce.marcar_emails_nuevos(pares, inicio, bool(antes))
            completas = not cancel.is_set()
            # si la vez anterior se cerró a medias, se suman sus cambios (aún no se habían avisado)
            previos = meta_antes.get("cambios_sin_avisar") or {}
            cambios = ce.resumir_cambios(pares, len(antes))
            for k in ("nuevos", "bajas", "emails_nuevos", "telefonos_nuevos"):
                cambios[k] += int(previos.get(k, 0))
            if previos:
                cambios["antes"] = int(previos.get("antes", cambios["antes"]))
            desde = previos.get("desde") or inicio
            meta = {"actualizado": datetime.now().isoformat(timespec="seconds"), "origen": "actualizada",
                    "fichas_completas": completas, "conservadas": conservadas,
                    "nuevos_desde": (desde if cambios["antes"] else "") if completas else meta_antes.get("nuevos_desde", "")}
            if not completas:
                meta["cambios_sin_avisar"] = dict(cambios, desde=desde)
            self.q.put(("act_fase", "Guardando la lista actualizada…"))
            ce.guardar_datos(self.ruta_datos, filas, info, meta)
            self._historial(cambios, conservadas, completas)
            self.q.put(("act_fin", preparar(filas), info, meta, cambios, conservadas, len(ce.por_consultar(filas, fichas))))
        except Exception as e:
            registrar_error(traceback.format_exc())
            self.q.put(("act_error", f"{type(e).__name__}: {e}"))

    def _historial(self, cambios, conservadas, completas):
        try:
            with open(os.path.join(self.dir_usuario, "actualizaciones.log"), "a", encoding="utf-8") as f:
                f.write(f"{datetime.now():%Y-%m-%d %H:%M} total {cambios['total']} · nuevos {cambios['nuevos']} · "
                        f"bajas {cambios['bajas']} · emails +{cambios['emails_nuevos']} · teléfonos "
                        f"+{cambios['telefonos_nuevos']}" + (f" · sin actualizar: {', '.join(conservadas)}" if conservadas else "")
                        + ("" if completas else " · interrumpida (sigue la próxima vez)") + "\n")
        except OSError:
            pass

    def _fin_hilo(self):
        self.actualizando = False
        self.btn_actualizar.configure(state="normal", text="  Actualizar ahora")
        if self.cerrando:
            self._destruir()
            return True
        return False

    def _act_parcial(self, filas, info):
        if not self.cerrando:
            self._poner_datos(filas, info, {})

    def _act_fin(self, filas, info, meta, cambios, conservadas, pendientes):
        if self._fin_hilo():
            return
        self.pendientes = pendientes
        self._poner_datos(filas, info, meta)
        nuevos, bajas = cambios["nuevos"], cambios["bajas"]
        em, te, total = cambios["emails_nuevos"], cambios["telefonos_nuevos"], cambios["total"]
        con_email = sum(1 for r in filas if r["email"])
        if not cambios.get("antes"):  # primera descarga en este equipo: no hay con qué comparar
            self._aviso("ok", f"Lista descargada: {miles(total)} centros, {miles(con_email)} con email.",
                        "A partir de ahora se actualizará sola una vez por semana.")
            self._log(f"Primera descarga terminada: {total} centros.")
            self._parpadear()
            messagebox.showinfo("Lista descargada", f"Ya tienes la lista de centros: {miles(total)} centros "
                                f"({miles(con_email)} con email).\n\nA partir de ahora se actualizará sola una vez "
                                "por semana mientras el programa esté abierto.", parent=self)
            return
        partes = [plural(nuevos, "centro nuevo", "centros nuevos") if nuevos else "no hay centros nuevos"]
        if em:
            partes.append(plural(em, "email nuevo", "emails nuevos"))
        if te:
            partes.append(plural(te, "teléfono añadido", "teléfonos añadidos"))
        nombres = [ce.CCAA[k]["nombre"] for k in conservadas]
        detalle = f"Total: {miles(total)} centros, {miles(con_email)} con email."
        if bajas:
            detalle += f" {plural(bajas, 'centro ya no aparece', 'centros ya no aparecen')} en las fuentes oficiales."
        if nombres:
            detalle += f" Sin actualizar esta vez (se mantienen sus datos): {', '.join(nombres)}."
        self._aviso("ok", "Lista actualizada: " + " · ".join(partes) + ".", detalle, nuevos=bool(self.n_nuevos),
                    emails=bool(self.n_emails_nuevos))
        self._log("Actualización terminada: " + " · ".join(partes) + f" · bajas {bajas}"
                  + (f" · sin actualizar: {', '.join(nombres)}" if nombres else ""))
        self._parpadear()
        msg = (f"La lista de centros se ha actualizado.\n\n"
               f"Centros nuevos: {miles(nuevos)}\n"
               f"Emails nuevos: {miles(em)}\n"
               f"Teléfonos añadidos: {miles(te)}\n")
        if bajas:
            msg += f"Centros que ya no aparecen: {miles(bajas)}\n"
        msg += f"\nTotal: {miles(total)} centros ({miles(con_email)} con email)."
        if nombres:
            msg += (f"\n\nNo se han podido actualizar: {', '.join(nombres)}. Se mantienen sus datos anteriores "
                    "y se volverá a intentar en la próxima actualización.")
        botones = ([("centros", "Ver centros nuevos")] if self.n_nuevos else []) + \
                  ([("emails", "Ver emails nuevos")] if self.n_emails_nuevos else [])
        if not botones:
            messagebox.showinfo("Lista actualizada", msg, parent=self)
            return
        msg += ("\n\nEn la lista, los centros nuevos salen en verde y los emails nuevos en amarillo. "
                "Puedes volver a verlos cuando quieras con los filtros «Centros nuevos» y «Emails nuevos».")
        eleccion = self._dialogo("Lista actualizada", msg, botones + [("ok", "Aceptar")])
        if eleccion == "centros":
            self._ver_nuevos()
        elif eleccion == "emails":
            self._ver_emails_nuevos()

    def _act_error(self, detalle):
        if self._fin_hilo():
            return
        self.no_antes_de = time.time() + REVISION_MS / 1000
        self.after(REVISION_MS + 1000, self._comprobar_actualizacion)
        self._log("La actualización no se ha podido completar: " + detalle)
        if self.rows:
            texto = ("Comprueba la conexión a internet. Se volverá a intentar automáticamente dentro de una hora; "
                     "mientras, puedes seguir usando la lista actual.")
        else:
            texto = ("Comprueba la conexión a internet. Se volverá a intentar dentro de una hora, "
                     "o pulsa «Actualizar ahora».")
        self._aviso("error", "No se ha podido actualizar la lista de centros.", texto)
        self._estado_vacio()

    def _act_fase(self, texto):
        if self.cerrando:
            return
        self._detalle_aviso(texto + "  Puedes seguir trabajando; te aviso cuando termine.")
        if str(self.ban_barra.cget("mode")) != "indeterminate":
            self.ban_barra.configure(mode="indeterminate", value=0)
            self.ban_barra.start(15)

    def _act_prog(self, hechas, total, eta):
        if self.cerrando:
            return
        if str(self.ban_barra.cget("mode")) != "determinate":
            self.ban_barra.stop()
            self.ban_barra.configure(mode="determinate")
        self.ban_barra.configure(maximum=max(total, 1), value=hechas)
        self._detalle_aviso(f"Completando emails y teléfonos: {miles(hechas)} de {miles(total)} centros · "
                            f"{tiempo_restante(eta)}. Puedes seguir trabajando; te aviso cuando termine.")

    # ------------------------------------------------------------------ aviso
    def _aviso(self, estado, texto, detalle="", barra=False, nuevos=False, emails=False):
        bg, _borde, fg, _, _ = ui.AVISOS[estado]
        px = self.e.px
        self.banner.configure(style=f"Aviso{estado}.TFrame")
        self.ban_icono.configure(image=self.iconos(ICONO_AVISO[estado]), bg=bg)
        self._aviso_partes = (texto, detalle)
        self.ban_txt.poner([(texto + (" " if detalle else ""), True), (detalle, False)], fg=fg, bg=bg)
        self.ban_botones.configure(bg=bg)
        if barra:
            self.ban_barra.grid()
            self.ban_barra.configure(mode="indeterminate", value=0)
            self.ban_barra.start(15)
        else:
            self.ban_barra.stop()
            self.ban_barra.grid_remove()
        for b in (self.ban_btn_nuevos, self.ban_btn_emails, self.ban_btn_detalles, self.ban_btn_cerrar):
            b.pack_forget()
            b.configure(style=f"AvisoBtn{estado}.TButton")
        self.ban_btn_cerrar.configure(style=f"AvisoX{estado}.TButton", image=self.iconos(CERRAR_AVISO[estado]))
        if nuevos:
            self.ban_btn_nuevos.pack(side="left", padx=(px(6), 0))
        if emails:
            self.ban_btn_emails.pack(side="left", padx=(px(6), 0))
        if estado != "ok":
            self.ban_btn_detalles.pack(side="left", padx=(px(6), 0))
        if estado != "act":
            self.ban_btn_cerrar.pack(side="left", padx=(px(4), 0))
        if not self.banner.winfo_ismapped():
            self.banner.pack(fill="x", pady=(px(14), 0), after=self.cabecera)

    def _detalle_aviso(self, detalle):
        texto = self._aviso_partes[0]
        self._aviso_partes = (texto, detalle)
        self.ban_txt.poner([(texto + " ", True), (detalle, False)])

    def _cerrar_aviso(self):
        self.ban_barra.stop()
        self.banner.pack_forget()

    def _ver_novedades(self, centros=False, emails=False):
        """Quita los demás filtros (también las comunidades) y deja solo los centros o emails nuevos."""
        self.sel_ccaa.clear()
        self._quitar_filtros()
        self.solo_nuevos = centros
        self.solo_emails_nuevos = emails
        self._pintar_ccaa()
        self._filtrar()

    def _ver_nuevos(self):
        if self.n_nuevos:
            self._ver_novedades(centros=True)

    def _ver_emails_nuevos(self):
        if self.n_emails_nuevos:
            self._ver_novedades(emails=True)

    def _dialogo(self, titulo, texto, botones):
        """Ventana de aviso con botones propios. botones: [(clave, texto)]. Devuelve la clave pulsada (o None)."""
        px = self.e.px
        w = tk.Toplevel(self)
        w.withdraw()  # se muestra ya centrado, sin saltos
        w.title(titulo)
        w.transient(self)
        w.resizable(False, False)
        w.configure(bg=C["blanco"])
        elegido = {"v": None}
        cuerpo = tk.Frame(w, bg=C["blanco"])
        cuerpo.pack(fill="both", expand=True, padx=px(22), pady=px(20))
        etiqueta(cuerpo, texto, self.fu.normal, C["texto"], C["blanco"], wraplength=px(460)).pack(anchor="w")
        fb = tk.Frame(cuerpo, bg=C["blanco"])
        fb.pack(anchor="e", pady=(px(18), 0))

        def elegir(v):
            elegido["v"] = v
            w.destroy()

        for n, (clave, txt) in enumerate(botones):
            b = ttk.Button(fb, text=txt, style=self.estilo_primario if n == 0 else self.estilo_dialogo,
                           command=lambda c=clave: elegir(c), cursor="hand2")
            b.pack(side="left", padx=(px(8) if n else 0, 0))
            if n == 0:
                b.focus_set()
        w.bind("<Escape>", lambda e: elegir(None))
        w.bind("<Return>", lambda e: elegir(botones[0][0]))
        w.protocol("WM_DELETE_WINDOW", lambda: elegir(None))
        try:
            w.update_idletasks()
            x = self.winfo_rootx() + (self.winfo_width() - w.winfo_reqwidth()) // 2
            y = self.winfo_rooty() + (self.winfo_height() - w.winfo_reqheight()) // 3
            w.geometry(f"+{max(0, x)}+{max(0, y)}")
        except tk.TclError:
            pass
        w.deiconify()
        try:
            w.wait_visibility()
            w.grab_set()
            w.lift()
        except tk.TclError:
            pass
        self.wait_window(w)
        return elegido["v"]

    def _parpadear(self):
        """Hace parpadear el botón de la barra de tareas si la ventana no está delante."""
        if sys.platform != "win32":
            return
        try:
            class FLASHWINFO(ctypes.Structure):
                _fields_ = [("cbSize", ctypes.c_uint), ("hwnd", ctypes.c_void_p), ("dwFlags", ctypes.c_uint),
                            ("uCount", ctypes.c_uint), ("dwTimeout", ctypes.c_uint)]
            fi = FLASHWINFO(ctypes.sizeof(FLASHWINFO), int(self.wm_frame(), 16), 0x3 | 0xC, 0, 0)
            ctypes.windll.user32.FlashWindowEx(ctypes.byref(fi))
        except Exception:
            pass

    # ------------------------------------------------------------------ acceso directo
    def _lanzar_acceso(self, forzar):
        def tarea():
            self.q.put(("acceso", acceso_directo(forzar), forzar))
        threading.Thread(target=tarea, daemon=True).start()

    def _fin_acceso(self, res, forzar):
        self._log(f"Acceso directo del escritorio: {res}")
        if res == "creado" or (forzar and res == "actualizado"):
            messagebox.showinfo("Acceso directo", "Tienes el acceso directo «Centros educativos» en el escritorio.\n"
                                                  "A partir de ahora puedes abrir el programa desde ahí.", parent=self)
        elif forzar and res.startswith("error"):
            messagebox.showwarning("Acceso directo", "No se ha podido crear el acceso directo en el escritorio.\n"
                                                     "Puedes seguir abriendo el programa con «Abrir Centros educativos».",
                                   parent=self)

    # ------------------------------------------------------------------ mensajes de los hilos
    def _sondear(self):
        if self.destruida:
            return
        try:
            for _ in range(500):
                msg = self.q.get_nowait()
                self._mensaje(msg)
                if self.destruida:
                    return
        except queue.Empty:
            pass
        self.after(100, self._sondear)

    def _mensaje(self, msg):
        tipo, args = msg[0], msg[1:]
        if tipo == "tk":  # resultado de un hilo que debe tratarse en la ventana (mapa)
            args[0](*args[1:])
        elif tipo == "log":
            self._log(args[0])
        elif tipo == "datos":
            self._datos_listos(*args)
        elif tipo == "act_fase":
            self._act_fase(*args)
        elif tipo == "act_prog":
            self._act_prog(*args)
        elif tipo == "act_parcial":
            self._act_parcial(*args)
        elif tipo == "act_fin":
            self._act_fin(*args)
        elif tipo == "act_error":
            self._act_error(*args)
        elif tipo == "act_cancelada":
            if not self._fin_hilo():
                self._cerrar_aviso()
        elif tipo == "acceso":
            self._fin_acceso(*args)
        elif tipo == "guardado":
            self._fin_guardado(*args)
        elif tipo == "error_guardar":
            self._error_guardar(*args)

    # ------------------------------------------------------------------ filtros y tabla
    def _programar_filtro(self):
        if self._filtro_pendiente:
            self.after_cancel(self._filtro_pendiente)
        self._filtro_pendiente = self.after(250, self._filtrar)

    def _quitar_filtros(self):
        self.provincia, self.titularidad = "", ""
        self.municipios, self.categorias = set(), set()
        self.solo_email = self.solo_nuevos = self.solo_emails_nuevos = False
        self.buscar.var.set("")
        if self._filtro_pendiente:
            self.after_cancel(self._filtro_pendiente)
            self._filtro_pendiente = None

    def _limpiar_filtros(self):
        self._quitar_filtros()
        self._filtrar()

    def _hay_filtros(self):
        return bool(self.buscar.get().strip() or self.provincia or self.municipios or self.categorias
                    or self.titularidad or self.solo_email)

    def _indices(self, excluir=()):
        """Centros que cumplen los filtros (menos los de 'excluir', para contar las opciones de un desplegable)."""
        vista = ce.filtrar(self.rows, "" if "texto" in excluir else self.buscar.get(),
                           "" if "provincia" in excluir else self.provincia,
                           "" if "titularidad" in excluir else self.titularidad,
                           self.solo_email and "email" not in excluir, comunidades=self.sel_ccaa or None,
                           municipios=None if "municipios" in excluir else (self.municipios or None),
                           categorias=None if "categorias" in excluir else (self.categorias or None))
        if self.nuevos_desde and self.solo_nuevos:
            vista = [i for i in vista if self._es_nuevo(self.rows[i])]
        if self.nuevos_desde and self.solo_emails_nuevos:
            vista = [i for i in vista if self._email_nuevo(self.rows[i])]
        return vista

    def _filtrar(self, arriba=True):
        if self._filtro_pendiente:
            self.after_cancel(self._filtro_pendiente)
        self._filtro_pendiente = None
        vista = self._indices()
        col, desc = self.orden
        if col:
            vista.sort(key=self._clave_orden(col), reverse=desc)
        self.vista = vista
        cursor = -1
        if self.activa is not None:
            cursor = next((n for n, i in enumerate(vista) if self.rows[i]["_k"] == self.activa), -1)
        self.tabla.activa, self.tabla.cursor = self.activa, cursor
        if arriba:
            self.tabla.lista.y = 0
        self._estado_vacio()
        self._pintar_resumen()

    def _clave_orden(self, col):
        campo = {"nombre": "nombre", "tipo": "_corto", "mun": "mun", "tel": "tel", "email": "email"}[col]
        cache = "_o_" + col
        rows = self.rows

        def clave(i):
            r = rows[i]
            v = r.get(cache)
            if v is None:
                v = r[cache] = (r[campo] == "", ce.norm(r[campo]))
            return v
        return clave

    def _ordenar(self, col, desc):
        self.orden = (col, desc)
        self._filtrar(arriba=False)
        if self.tabla.cursor >= 0:
            self.tabla.lista.ver(self.tabla.cursor)

    def _es_nuevo(self, r):
        return bool(self.nuevos_desde) and (r.get("alta") or "") >= self.nuevos_desde

    def _email_nuevo(self, r):
        return bool(self.nuevos_desde) and bool(r.get("email")) and (r.get("email_alta") or "") >= self.nuevos_desde

    def _estado_vacio(self):
        if not self.cargada:
            texto, enlace = "Abriendo la lista de centros…", None
        elif not self.rows:
            texto = ("Todavía no hay ninguna lista de centros en este equipo.\nSe está descargando; tardará unos "
                     "minutos y te aviso al terminar." if self.actualizando else
                     "Todavía no hay ninguna lista de centros en este equipo.\n"
                     "Pulsa «Actualizar ahora», abajo a la izquierda (hace falta conexión a internet).")
            enlace = None
        elif not self.vista:
            texto = "Ningún centro cumple estos filtros."
            enlace = ("Limpiar filtros", self._limpiar_filtros) if (self._hay_filtros() or self.solo_nuevos
                                                                    or self.solo_emails_nuevos) else None
        else:
            texto, enlace = "", None
        self.tabla.poner(len(self.vista), texto, enlace)

    def _titulo_ambito(self):
        sel = [k for k in self._orden_ccaa() if k in self.sel_ccaa]
        if not sel or len(sel) == len(ce.CCAA):
            return "Toda España"
        nombres = [ce.CCAA[k]["nombre"] for k in sel]
        if len(sel) == 1:
            return nombres[0]
        if len(sel) == 2 and len(nombres[0]) + len(nombres[1]) <= 30:
            return f"{nombres[0]} y {nombres[1]}"
        return f"{len(sel)} comunidades"

    def _pintar_resumen(self):
        rows, vista = self.rows, self.vista
        self.lbl_titulo.configure(text=self._titulo_ambito())
        self.after_idle(self._ancho_buscador)
        en_ambito = sum(self.cuenta_ccaa.get(k, 0) for k in self.sel_ccaa) if self.sel_ccaa else len(rows)
        self.lbl_total.configure(text=plural(en_ambito, "centro", "centros") if rows else "")
        n = len(vista)
        email = sum(1 for i in vista if rows[i]["email"])
        tel = sum(1 for i in vista if rows[i]["tel"])
        filtrado = self._hay_filtros() or self.solo_nuevos or self.solo_emails_nuevos
        lt, v, x = self.kpis["centros"]
        lt.configure(text="En la lista" if filtrado else "Centros")
        v.configure(text=miles(n) if rows else "—")
        for clave, cuenta in (("email", email), ("tel", tel)):
            _, v, x = self.kpis[clave]
            v.configure(text=miles(cuenta) if rows else "—")
            x.configure(text=porcentaje(cuenta, n) if n else "")
        # chips
        if self.provincia:
            self._poner_chip(self.chip_prov, f"Provincia: {self.provincia}", True)
        else:
            self._poner_chip(self.chip_prov, "Provincia", False)
        if len(self.municipios) == 1:
            self._poner_chip(self.chip_mun, f"Municipio: {next(iter(self.municipios))[1]}", True)
        else:
            self._poner_chip(self.chip_mun, f"Municipios: {len(self.municipios)}" if self.municipios else "Municipio",
                             bool(self.municipios))
        if len(self.categorias) == 1:
            nombre = ce.NOMBRE_CATEGORIA[next(iter(self.categorias))].split(" (")[0]
            self._poner_chip(self.chip_tipo, f"Tipo: {nombre}", True)
        else:
            self._poner_chip(self.chip_tipo, f"Tipos: {len(self.categorias)}" if self.categorias else "Tipo de centro",
                             bool(self.categorias))
        self._poner_chip(self.chip_tit, f"Titularidad: {self.titularidad}" if self.titularidad else "Titularidad",
                         bool(self.titularidad))
        self._poner_chip(self.chip_email, "Con email", self.solo_email, bool(rows))
        self._poner_chip(self.chip_nuevos, f"Centros nuevos · {miles(self.n_nuevos)}", self.solo_nuevos,
                         bool(self.n_nuevos))
        self._poner_chip(self.chip_emails, f"Emails nuevos · {miles(self.n_emails_nuevos)}", self.solo_emails_nuevos,
                         bool(self.n_emails_nuevos))
        self.btn_limpiar.visible = filtrado
        self.flujo.recolocar()
        self._pintar_nota()
        self._actualizar_cuenta()

    def _pintar_nota(self):
        texto = ("Haz clic en un centro para ver su ficha a la derecha" if self.ficha_visible else
                 "Haz doble clic en un centro para ver su ficha") + "; la casilla lo marca para descargarlo."
        if self.n_emails_nuevos or self.n_nuevos:
            partes = []
            if self.n_emails_nuevos:
                partes.append("en amarillo, los emails nuevos")
            if self.n_nuevos:
                partes.append("en verde, los centros nuevos")
            texto += " De la última actualización: " + " y ".join(partes) + "."
        self.nota.configure(text=texto)

    def _actualizar_cuenta(self):
        n = len(self.marcados)
        self.kpis["marcados"][1].configure(text=miles(n))
        self.lbl_marcados.configure(text=plural(n, "marcado", "marcados"))
        self.btn_desmarcar.state(["!disabled"] if n else ["disabled"])
        self.btn_copiar.state(["!disabled"] if (self.marcados or self.vista) else ["disabled"])
        self.btn_desc_lista.state(["!disabled"] if self.vista and not self.guardando else ["disabled"])
        self.btn_desc_marc.state(["!disabled"] if n and not self.guardando else ["disabled"])
        if self.guardando:
            self.btn_desc_marc.configure(text="  Guardando…")
        else:
            self.btn_desc_marc.configure(text=f"  Descargar {plural(n, 'marcado', 'marcados')}" if n
                                         else "  Descargar marcados")
        self._colocar_pie(self.pie.winfo_width())

    def _fila_tabla(self, i):
        r = self.rows[self.vista[i]]
        estado = "nuevo" if self._es_nuevo(r) else "email" if self._email_nuevo(r) else ""
        return r["_k"], r["nombre"], r["_corto"], r["mun"], r["_tel"], r["email"], r["_k"] in self.marcados, estado

    def _seleccionar_fila(self, i):
        r = self.rows[self.vista[i]]
        self.activa = r["_k"]
        self.tabla.seleccionar(i, r["_k"])
        self._mostrar_ficha(r)

    def _doble_fila(self, i):
        if i < 0:
            return
        self._seleccionar_fila(i)
        if not self.ficha_visible:
            self._ficha_flotante(self.rows[self.vista[i]])

    def _marcar_fila(self, i):
        self._alternar_marca(self.rows[self.vista[i]]["_k"])

    def _alternar_marca(self, k):
        self.marcados.symmetric_difference_update({k})
        self.tabla.refrescar()
        self._actualizar_cuenta()
        self._refrescar_mapa()
        for f in self._fichas_abiertas():
            if f.r is not None and f.r["_k"] == k:
                f.marcado()

    def _estado_todos(self):
        if not self.vista:
            return ""
        marcados = sum(1 for i in self.vista if self.rows[i]["_k"] in self.marcados) if self.marcados else 0
        return "si" if marcados == len(self.vista) else "mixto" if marcados else "no"

    def _marcar_todos(self):
        claves = {self.rows[i]["_k"] for i in self.vista}
        if self._estado_todos() == "si":
            self.marcados -= claves
        else:
            self.marcados |= claves
        self.tabla.refrescar()
        self._actualizar_cuenta()
        self._refrescar_mapa()
        for f in self._fichas_abiertas():
            f.marcado()

    def _desmarcar_todos(self):
        self.marcados.clear()
        self.tabla.refrescar()
        self._actualizar_cuenta()
        self._refrescar_mapa()
        for f in self._fichas_abiertas():
            f.marcado()

    def _filas_marcadas(self):
        return [r for r in self.rows if r["_k"] in self.marcados]

    def _copiar_emails(self):
        filas = self._filas_marcadas() if self.marcados else [self.rows[i] for i in self.vista]
        emails = list(dict.fromkeys(r["email"] for r in filas if r["email"]))
        origen = "marcados" if self.marcados else "de la lista"
        if not emails:
            messagebox.showinfo("Copiar emails", f"No hay emails que copiar en los centros {origen}.", parent=self)
            return
        self.clipboard_clear()
        self.clipboard_append("; ".join(emails))
        self._estado(f"Copiados {miles(len(emails))} emails de los centros {origen}: pégalos con Ctrl+V.", 8000)

    def _estado(self, texto, ms=0):
        if self._estado_tmp:
            self.after_cancel(self._estado_tmp)
            self._estado_tmp = None
        self.lbl_estado.configure(text=texto)
        if ms and texto:
            self._estado_tmp = self.after(ms, lambda: self._estado(""))
        self._colocar_pie(self.pie.winfo_width())

    # ------------------------------------------------------------------ desplegables de los filtros
    def _abrir_desplegable(self, chip, items, seleccion, multiple, al_cambiar, buscador=False, ayuda="Buscar",
                           ancho=300):
        ancla, cuando = self._cierre_desplegable
        if ancla is chip and time.monotonic() - cuando < 0.5:  # el clic en el chip acaba de cerrarlo
            return
        if self.desplegable is not None:
            self.desplegable.cerrar()
        self.desplegable = ui.Desplegable(self, chip, items, seleccion, multiple, al_cambiar, buscador, ayuda,
                                          al_cerrar=self._fin_desplegable, ancho=ancho)

    def _fin_desplegable(self):
        self.desplegable = None

    def _abrir_provincia(self):
        cuenta = Counter(self.rows[i]["prov"] for i in self._indices(("provincia", "municipios")))
        items = [("", "Todas las provincias", sum(cuenta.values()), False)]
        items += [(p, p, cuenta[p], False) for p in sorted((p for p in cuenta if p), key=ce.norm)]
        if self.provincia and self.provincia not in cuenta:
            items.append((self.provincia, self.provincia, 0, True))
        self._abrir_desplegable(self.chip_prov, items, {self.provincia}, False, self._elegir_provincia,
                                buscador=len(items) > 9, ayuda="Buscar provincia")

    def _elegir_provincia(self, seleccion):
        prov = next(iter(seleccion), "")
        if prov != self.provincia:
            self.provincia = prov
            if prov:
                self.municipios = {m for m in self.municipios if m[0] == prov}
            self._filtrar()

    def _abrir_municipio(self):
        cuenta = Counter((self.rows[i]["prov"], self.rows[i]["mun"]) for i in self._indices(("municipios",))
                         if self.rows[i]["mun"])
        for m in self.municipios:
            cuenta.setdefault(m, 0)
        varias = len({p for p, _ in cuenta}) > 1
        orden = sorted(cuenta.items(), key=lambda kv: (ce.norm(kv[0][1]), ce.norm(kv[0][0])))
        items = [((p, m), f"{m} ({p})" if varias else m, n, n == 0) for (p, m), n in orden]
        ayuda = "Buscar municipio" + ("" if self.provincia or not varias else " (elige antes la provincia para acotar)")
        self._abrir_desplegable(self.chip_mun, items, set(self.municipios), True, self._elegir_municipios,
                                buscador=True, ayuda="Buscar municipio" if len(ayuda) > 40 else ayuda, ancho=340)

    def _elegir_municipios(self, seleccion):
        self.municipios = set(seleccion)
        self._filtrar()

    def _abrir_tipo(self):
        cuenta = Counter(self.rows[i]["_cat"] for i in self._indices(("categorias",)))
        items = [(k, nombre, cuenta.get(k, 0), not cuenta.get(k)) for k, nombre in ce.CATEGORIAS]
        self._abrir_desplegable(self.chip_tipo, items, set(self.categorias), True, self._elegir_tipos, ancho=330)

    def _elegir_tipos(self, seleccion):
        self.categorias = set(seleccion)
        self._filtrar()

    def _abrir_titularidad(self):
        cuenta = Counter(self.rows[i]["nat"] for i in self._indices(("titularidad",)))
        items = [("", "Todas", sum(cuenta.values()), False)]
        items += [(t, t, cuenta.get(t, 0), not cuenta.get(t)) for t in TITULARIDADES]
        self._abrir_desplegable(self.chip_tit, items, {self.titularidad}, False, self._elegir_titularidad, ancho=260)

    def _elegir_titularidad(self, seleccion):
        tit = next(iter(seleccion), "")
        if tit != self.titularidad:
            self.titularidad = tit
            self._filtrar()

    def _alternar_email(self):
        self.solo_email = not self.solo_email
        self._filtrar()

    def _alternar_novedad(self, cual):
        if cual == "centros":
            self.solo_nuevos = not self.solo_nuevos
        else:
            self.solo_emails_nuevos = not self.solo_emails_nuevos
        self._filtrar()

    # ------------------------------------------------------------------ ficha
    def _fichas_abiertas(self):
        fichas = [self.ficha]
        if self.ventana_ficha is not None:
            w, f = self.ventana_ficha
            try:
                if w.winfo_exists():
                    fichas.append(f)
            except tk.TclError:
                pass
        return fichas

    def _mostrar_ficha(self, r):
        for f in self._fichas_abiertas():
            f.mostrar(r)
        if r is not None and self._mapa_abierto():
            self.ventana_mapa.mostrar(r)

    def _mapa_abierto(self):
        try:
            return self.ventana_mapa is not None and bool(self.ventana_mapa.winfo_exists())
        except tk.TclError:
            return False

    def _ampliar_mapa(self, r):
        if not self._mapa_abierto():
            self.ventana_mapa = VentanaMapa(self)
        self.ventana_mapa.mostrar(r)
        self.ventana_mapa.deiconify()
        self.ventana_mapa.lift()
        self.ventana_mapa.focus_set()

    def _refrescar_mapa(self):
        if self._mapa_abierto():
            self.ventana_mapa._pintar_puntos()

    def _elegir_centro(self, k):
        """Elige un centro por su clave (desde el mapa): ficha a la derecha y, si está en la lista, su fila."""
        r = self.por_clave.get(k)
        if r is None:
            return
        self.activa = k
        i = next((n for n, j in enumerate(self.vista) if self.rows[j]["_k"] == k), -1)
        self.tabla.seleccionar(i, k)
        self._mostrar_ficha(r)

    def _ficha_flotante(self, r):
        px = self.e.px
        abierta = len(self._fichas_abiertas()) > 1
        if not abierta:
            w = tk.Toplevel(self)
            w.title("Ficha del centro")
            w.configure(bg=C["blanco"])
            w.transient(self)
            alto = min(px(700), self.winfo_screenheight() - px(120))
            x = self.winfo_rootx() + max(0, self.winfo_width() - px(ANCHO_FICHA + 40))
            w.geometry(f"{px(ANCHO_FICHA + 20)}x{alto}+{x}+{self.winfo_rooty() + px(40)}")
            f = Ficha(w, self)
            f.pack(fill="both", expand=True)
            w.bind("<Escape>", lambda e: w.destroy())
            self.ventana_ficha = (w, f)
        w, f = self.ventana_ficha
        f.mostrar(r)
        w.deiconify()
        w.lift()

    # ------------------------------------------------------------------ descarga
    def _etiqueta(self, marcados):
        if marcados:
            return f"{len(self.marcados)} centros marcados"
        sel = [k for k in self._orden_ccaa() if k in self.sel_ccaa]
        if len(self.municipios) == 1:
            base = next(iter(self.municipios))[1]
        elif self.provincia:
            base = self.provincia
        elif not sel or len(sel) == len(ce.CCAA):
            base = "España"
        elif len(sel) <= 3:
            base = ", ".join(ce.CCAA[k]["nombre"] for k in sel)
        else:
            base = f"{len(sel)} comunidades"
        if len(self.categorias) == 1:
            base += " - " + ce.NOMBRE_CATEGORIA[next(iter(self.categorias))].split(" (")[0]
        novedades = [t for v, t in ((self.solo_nuevos, "centros nuevos"), (self.solo_emails_nuevos, "emails nuevos"))
                     if v]
        return (base + (" - " + " y ".join(novedades) if novedades else "")
                + (" (filtrado)" if self._hay_filtros() else ""))

    def _descargar(self, marcados):
        if self.guardando:
            return
        filas = self._filas_marcadas() if marcados else [self.rows[i] for i in self.vista]
        if not filas:
            return
        nombre = f"Centros educativos - {self._etiqueta(marcados)} - {datetime.now():%Y-%m-%d}.xlsx"
        ruta = filedialog.asksaveasfilename(parent=self, title="Guardar como", initialdir=self.carpeta,
                                            initialfile=nombre, defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx"), ("CSV separado por ;", "*.csv")])
        if not ruta:
            return
        self.carpeta = os.path.dirname(ruta)
        self.guardando = True
        self._actualizar_cuenta()
        self._estado(f"Guardando {miles(len(filas))} centros…")
        info = self.info

        def tarea():
            try:
                self.q.put(("guardado", ce.guardar(ruta, filas, info), len(filas)))
            except Exception as e:
                registrar_error(traceback.format_exc())
                self.q.put(("error_guardar", f"{type(e).__name__}: {e}"))

        threading.Thread(target=tarea, daemon=True).start()

    def _fin_guardado(self, destino, n):
        self.guardando = False
        self._actualizar_cuenta()
        self._estado(f"Guardado: {os.path.basename(destino)} ({miles(n)} centros).")
        self._log(f"Guardado {destino} ({n} centros).")
        if self.cerrando:
            if not self.actualizando:
                self._destruir()
            return
        if messagebox.askyesno("Descarga terminada", f"Se ha guardado el fichero:\n{destino}\n\n¿Quieres abrirlo ahora?",
                               parent=self):
            abrir(destino)

    def _error_guardar(self, detalle):
        self.guardando = False
        self._actualizar_cuenta()
        self._estado("No se ha podido guardar el fichero.")
        self._log("Error al guardar: " + detalle)
        if self.cerrando:
            if not self.actualizando:
                self._destruir()
            return
        messagebox.showerror("No se ha podido guardar", "No se ha podido guardar el fichero.\n"
                             "Comprueba que no esté abierto en Excel y que puedas escribir en esa carpeta.\n\n"
                             "Detalle: " + detalle, parent=self)

    # ------------------------------------------------------------------ varios
    def _log(self, texto):
        linea = f"{datetime.now():%H:%M:%S}  {texto}"
        self.log_lineas.append(linea)
        del self.log_lineas[:-3000]
        if self.ventana_log is not None and self.ventana_log.winfo_exists():
            self.texto_log.configure(state="normal")
            self.texto_log.insert("end", linea + "\n")
            self.texto_log.see("end")
            self.texto_log.configure(state="disabled")

    def _ver_log(self):
        """Historial de actualizaciones de este equipo y detalles de la sesión actual."""
        if self.ventana_log is not None and self.ventana_log.winfo_exists():
            self.ventana_log.deiconify()
            self.ventana_log.lift()
            return
        px = self.e.px
        w = tk.Toplevel(self)
        w.title("Historial de actualizaciones")
        w.geometry(f"{px(780)}x{px(440)}")
        w.configure(bg=C["blanco"])
        self.texto_log = tk.Text(w, wrap="word", font=self.fu.mono, bd=0, highlightthickness=0, padx=px(14),
                                 pady=px(12), bg=C["blanco"], fg=C["texto"], spacing1=1)
        sb = ttk.Scrollbar(w, orient="vertical", command=self.texto_log.yview, style=self.estilo_barra)
        self.texto_log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.texto_log.pack(side="left", fill="both", expand=True)
        historial = ""
        try:
            with open(os.path.join(self.dir_usuario, "actualizaciones.log"), encoding="utf-8") as f:
                historial = "".join(f.readlines()[-300:])
        except OSError:
            pass
        self.texto_log.tag_configure("titulo", font=self.fu.semi, spacing1=px(4), spacing3=px(6))
        self.texto_log.insert("end", "Actualizaciones de la lista en este equipo\n", "titulo")
        self.texto_log.insert("end", (historial or "Todavía no se ha actualizado la lista en este equipo.\n") + "\n")
        self.texto_log.insert("end", "Detalles desde que se abrió el programa\n", "titulo")
        self.texto_log.insert("end", "\n".join(self.log_lineas) + ("\n" if self.log_lineas else ""))
        self.texto_log.see("end")
        self.texto_log.configure(state="disabled")
        self.ventana_log = w

    def _cerrar(self):
        if self.cerrando:
            return
        if self.actualizando:
            if not messagebox.askyesno("Salir", "Se está actualizando la lista de centros.\n"
                                                "Si sales ahora, la actualización seguirá sola la próxima vez que "
                                                "abras el programa.\n\n¿Quieres salir?", parent=self):
                return
            self.cerrando = True
            self.cancel_act.set()
            self._aviso("act", "Cerrando…", "Guardando lo descargado hasta ahora.", barra=True)
            self.after(15000, self._destruir)  # por si alguna conexión tarda en responder
            return
        if self.guardando:
            self.cerrando = True
            self._estado("Terminando de guardar el fichero…")
            self.after(10000, self._destruir)
            return
        self._destruir()

    def _destruir(self):
        if self.destruida:
            return
        self.destruida = True
        if self._gancho_log in ce._LOG_HOOKS:
            ce._LOG_HOOKS.remove(self._gancho_log)
        self.servicio_mapa.cerrar()
        try:
            self.ban_barra.stop()
            self.destroy()
        except tk.TclError:
            pass


def main():
    try:
        app = App()
    except Exception:
        ruta = registrar_error(traceback.format_exc())
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror("Centros educativos", "No se ha podido abrir el programa.\n\n"
                                 f"Detalle guardado en:\n{ruta}")
            r.destroy()
        except Exception:
            pass
        return
    app.mainloop()
    hilo = app.hilo_act
    if hilo is not None and hilo.is_alive():  # una conexión lenta no debe dejar el programa abierto por detrás
        hilo.join(timeout=3)
        os._exit(0)


if __name__ == "__main__":
    main()
