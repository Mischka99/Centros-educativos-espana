# -*- coding: utf-8 -*-
"""Arranque de «Centros educativos».

«Centros educativos.exe» (en la versión de la Microsoft Store, «CentrosEducativos.exe») es el pythonw.exe oficial de
Python con otro nombre, firmado por la Python Software Foundation. Python carga este módulo siempre al arrancar: si se
ha abierto ese ejecutable con doble clic o desde el menú Inicio (sin argumentos), abre el programa. Así no hace falta
ningún .bat, un tipo de archivo que el Control inteligente de aplicaciones de Windows 11 bloquea cuando viene de
internet.
"""
import os
import sys

NOMBRES = {"centros educativos", "centroseducativos"}


def _registrar_error(base, texto):
    """En programa\\error.log; si la carpeta es de solo lectura (Microsoft Store), en la del usuario."""
    usuario = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "CentrosEducativos")
    for carpeta in (os.path.join(base, "programa"), usuario):
        try:
            os.makedirs(carpeta, exist_ok=True)
            with open(os.path.join(carpeta, "error.log"), "a", encoding="utf-8") as f:
                f.write(texto)
            return
        except OSError:
            continue


def _abrir_programa():
    nombre = os.path.splitext(os.path.basename(sys.executable or ""))[0].lower()
    if nombre not in NOMBRES or sys.argv != [""]:
        return
    base = os.path.dirname(os.path.abspath(sys.executable))
    script = os.path.join(base, "programa", "centros_gui.pyw")
    codigo = 0
    try:
        import runpy
        sys.argv = [script]
        runpy.run_path(script, run_name="__main__")
    except SystemExit as e:
        codigo = e.code if isinstance(e.code, int) else 0
    except BaseException:
        codigo = 1
        try:
            import traceback
            from datetime import datetime
            _registrar_error(base, f"--- {datetime.now():%Y-%m-%d %H:%M:%S} (arranque)\n{traceback.format_exc()}\n")
        except Exception:
            pass
    finally:
        os._exit(codigo)  # el programa ya terminó: no seguir arrancando el intérprete


_abrir_programa()
