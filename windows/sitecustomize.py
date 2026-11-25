# -*- coding: utf-8 -*-
"""Arranque de «Centros educativos».

«Centros educativos.exe» es el pythonw.exe oficial de Python con otro nombre (firmado por la Python Software
Foundation). Python carga este módulo siempre al arrancar: si se ha abierto «Centros educativos.exe» con doble clic
(sin argumentos), abre el programa. Así no hace falta ningún .bat, un tipo de archivo que el Control inteligente de
aplicaciones de Windows 11 bloquea cuando viene de internet.
"""
import os
import sys


def _abrir_programa():
    nombre = os.path.splitext(os.path.basename(sys.executable or ""))[0].lower()
    if nombre != "centros educativos" or sys.argv != [""]:
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
            with open(os.path.join(base, "programa", "error.log"), "a", encoding="utf-8") as f:
                f.write(f"--- {datetime.now():%Y-%m-%d %H:%M:%S} (arranque)\n{traceback.format_exc()}\n")
        except OSError:
            pass
    finally:
        os._exit(codigo)  # el programa ya terminó: no seguir arrancando el intérprete


_abrir_programa()
