#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
centros_educativos.py
=====================

Directorio de centros educativos no universitarios de España con teléfono, email,
web y titularidad, a partir de fuentes oficiales, filtrando por comunidad autónoma.

Fuentes
-------
1. Directorio de datos abiertos de cada comunidad (teléfono, email, web).
2. Registro Estatal de Centros Docentes del Ministerio (educacion.gob.es/centros):
   - listado por comunidad: titularidad, tipo de centro y teléfono de respaldo;
   - ficha de cada centro: completa email, teléfono y web cuando faltan.

Uso rápido
----------
    python centros_educativos.py --listar
    python centros_educativos.py --ccaa baleares
    python centros_educativos.py --ccaa valencia galicia -o centros.xlsx
    python centros_educativos.py --ccaa todas -o espana.csv
    python centros_educativos.py --ccaa madrid --sin-fichas

Requisitos: Python 3.9+ y `pip install openpyxl xlrd`.
Ejecútalo desde una conexión española: algunas comunidades bloquean IP extranjeras.
"""

import argparse
import csv
import html
import http.client
import http.cookiejar
import io
import json
import os
import re
import ssl
import sys
import threading
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

VERSION = "1.0.0"
UA = "Mozilla/5.0 (compatible; centros-educativos/1.0)"

# ---------------------------------------------------------------------------
# Comunidades
# ---------------------------------------------------------------------------
# rcd: identificador de la comunidad en el buscador del Ministerio (no es el código INE).
CCAA = {
    "and": {"nombre": "Andalucía", "rcd": "01", "alias": ["andalucia"]},
    "ara": {"nombre": "Aragón", "rcd": "02", "alias": ["aragon"]},
    "ast": {"nombre": "Asturias", "rcd": "03", "alias": ["asturias", "principado de asturias"]},
    "bal": {"nombre": "Illes Balears", "rcd": "04", "alias": ["baleares", "illes balears", "islas baleares", "balears"]},
    "can": {"nombre": "Canarias", "rcd": "05", "alias": ["canarias"]},
    "ctb": {"nombre": "Cantabria", "rcd": "06", "alias": ["cantabria"]},
    "cyl": {"nombre": "Castilla y León", "rcd": "07", "alias": ["castilla y leon", "castilla-leon", "castillayleon"]},
    "clm": {"nombre": "Castilla-La Mancha", "rcd": "08", "alias": ["castilla-la mancha", "castilla la mancha"]},
    "cat": {"nombre": "Cataluña", "rcd": "09", "alias": ["cataluna", "catalunya", "catalonia"]},
    "ext": {"nombre": "Extremadura", "rcd": "10", "alias": ["extremadura"]},
    "gal": {"nombre": "Galicia", "rcd": "11", "alias": ["galicia"]},
    "rio": {"nombre": "La Rioja", "rcd": "12", "alias": ["la rioja", "rioja"]},
    "mad": {"nombre": "Comunidad de Madrid", "rcd": "13", "alias": ["madrid", "comunidad de madrid"]},
    "mur": {"nombre": "Región de Murcia", "rcd": "14", "alias": ["murcia", "region de murcia"]},
    "nav": {"nombre": "Navarra", "rcd": "15", "alias": ["navarra", "comunidad foral de navarra"]},
    "val": {"nombre": "Comunitat Valenciana", "rcd": "16", "alias": ["valencia", "comunitat valenciana", "comunidad valenciana", "cv"]},
    "pvs": {"nombre": "País Vasco", "rcd": "17", "alias": ["pais vasco", "euskadi", "pv"]},
    "ceu": {"nombre": "Ceuta", "rcd": "18", "alias": ["ceuta"]},
    "mel": {"nombre": "Melilla", "rcd": "19", "alias": ["melilla"]},
}

PROV = {"01": "Araba/Álava", "02": "Albacete", "03": "Alicante/Alacant", "04": "Almería", "05": "Ávila", "06": "Badajoz",
        "07": "Illes Balears", "08": "Barcelona", "09": "Burgos", "10": "Cáceres", "11": "Cádiz", "12": "Castellón/Castelló",
        "13": "Ciudad Real", "14": "Córdoba", "15": "A Coruña", "16": "Cuenca", "17": "Girona", "18": "Granada",
        "19": "Guadalajara", "20": "Gipuzkoa", "21": "Huelva", "22": "Huesca", "23": "Jaén", "24": "León", "25": "Lleida",
        "26": "La Rioja", "27": "Lugo", "28": "Madrid", "29": "Málaga", "30": "Murcia", "31": "Navarra", "32": "Ourense",
        "33": "Asturias", "34": "Palencia", "35": "Las Palmas", "36": "Pontevedra", "37": "Salamanca",
        "38": "Santa Cruz de Tenerife", "39": "Cantabria", "40": "Segovia", "41": "Sevilla", "42": "Soria", "43": "Tarragona",
        "44": "Teruel", "45": "Toledo", "46": "Valencia/València", "47": "Valladolid", "48": "Bizkaia", "49": "Zamora",
        "50": "Zaragoza", "51": "Ceuta", "52": "Melilla"}
PROV_CCAA = {"and": "04 11 14 18 21 23 29 41", "ara": "22 44 50", "ast": "33", "bal": "07", "can": "35 38", "ctb": "39",
             "cyl": "05 09 24 34 37 40 42 47 49", "clm": "02 13 16 19 45", "cat": "08 17 25 43", "ext": "06 10",
             "gal": "15 27 32 36", "rio": "26", "mad": "28", "mur": "30", "nav": "31", "val": "03 12 46",
             "pvs": "01 20 48", "ceu": "51", "mel": "52"}
PROV_CCAA = {k: set(v.split()) for k, v in PROV_CCAA.items()}

RCD_BASE = "https://www.educacion.gob.es/centros/"
RCD_EXPORT = RCD_BASE + "exportarListadoCentrosExcel"
RCD_FICHA = RCD_BASE + "detalleCentro"
RCD_HOME = RCD_BASE + "home.do"

# Unidades sin alumnado (equipos de orientación, aulas hospitalarias, centros del profesorado...).
# Se excluyen salvo --incluir-no-docentes.
NO_DOCENTES = ("equipo de orientacion", "equipos de orientacion", "equipo de atencion temprana", "equipo especifico",
               "aula hospitalaria", "aulas hospitalarias", "unidad hospitalaria", "atencion hospitalaria",
               "establecimiento administrativo", "servicios centrales", "centro del profesorado",
               "centro de profesores", "centro de apoyo al profesorado", "centre de professorat")
SIGLAS_NO_DOCENTES = {  # siglas propias de cada directorio
    "bal": {"EAP", "EOEP", "CEP", "CEFIRDE", "SAED", "EPEP", "MEIB"},
    "mad": {"A. HOSP.", "A. HOSP. PRIV.", "C. DOC. EN EXTRANJ.", "CTIF", "EOEP ATEN-TEMPR", "EOEP ESPECÍFICO",
            "EOEP GENERAL", "EST ADMI", "ZNAS.CN.", "ISMIE", "CFII", "CREACIM", "UNIVERSIDAD"},
    "mur": {"AHOS"},
}
PREFIJOS_NO_DOCENTES = {"cat": ("HDA ", "AH ", "URPIJ ", "UTCA "),
                        "nav": ("CAP ", "ORG ", "CREENA", "CRikasNOVA", "EIBZ", "EPA Euskarabidea")}

# Certificados intermedios que algunos servidores no envían (catedu.es: Sectigo DV R36;
# servidores con certificado de la FNMT). Se AÑADEN a los del sistema; no se desactiva la verificación.
EXTRA_CERTS = """
-----BEGIN CERTIFICATE-----
MIIGTDCCBDSgAwIBAgIQOXpmzCdWNi4NqofKbqvjsTANBgkqhkiG9w0BAQwFADBf
MQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMTYwNAYDVQQD
Ey1TZWN0aWdvIFB1YmxpYyBTZXJ2ZXIgQXV0aGVudGljYXRpb24gUm9vdCBSNDYw
HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBgMQswCQYDVQQGEwJHQjEY
MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMTcwNQYDVQQDEy5TZWN0aWdvIFB1Ymxp
YyBTZXJ2ZXIgQXV0aGVudGljYXRpb24gQ0EgRFYgUjM2MIIBojANBgkqhkiG9w0B
AQEFAAOCAY8AMIIBigKCAYEAljZf2HIz7+SPUPQCQObZYcrxLTHYdf1ZtMRe7Yeq
RPSwygz16qJ9cAWtWNTcuICc++p8Dct7zNGxCpqmEtqifO7NvuB5dEVexXn9RFFH
12Hm+NtPRQgXIFjx6MSJcNWuVO3XGE57L1mHlcQYj+g4hny90aFh2SCZCDEVkAja
EMMfYPKuCjHuuF+bzHFb/9gV8P9+ekcHENF2nR1efGWSKwnfG5RawlkaQDpRtZTm
M64TIsv/r7cyFO4nSjs1jLdXYdz5q3a4L0NoabZfbdxVb+CUEHfB0bpulZQtH1Rv
38e/lIdP7OTTIlZh6OYL6NhxP8So0/sht/4J9mqIGxRFc0/pC8suja+wcIUna0HB
pXKfXTKpzgis+zmXDL06ASJf5E4A2/m+Hp6b84sfPAwQ766rI65mh50S0Di9E3Pn
2WcaJc+PILsBmYpgtmgWTR9eV9otfKRUBfzHUHcVgarub/XluEpRlTtZudU5xbFN
xx/DgMrXLUAPaI60fZ6wA+PTAgMBAAGjggGBMIIBfTAfBgNVHSMEGDAWgBRWc1hk
lfmSGrASKgRieaFAFYghSTAdBgNVHQ4EFgQUaMASFhgOr872h6YyV6NGUV3LBycw
DgYDVR0PAQH/BAQDAgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0lBBYwFAYI
KwYBBQUHAwEGCCsGAQUFBwMCMBsGA1UdIAQUMBIwBgYEVR0gADAIBgZngQwBAgEw
VAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdv
UHVibGljU2VydmVyQXV0aGVudGljYXRpb25Sb290UjQ2LmNybDCBhAYIKwYBBQUH
AQEEeDB2ME8GCCsGAQUFBzAChkNodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3Rp
Z29QdWJsaWNTZXJ2ZXJBdXRoZW50aWNhdGlvblJvb3RSNDYucDdjMCMGCCsGAQUF
BzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTANBgkqhkiG9w0BAQwFAAOCAgEA
YtOC9Fy+TqECFw40IospI92kLGgoSZGPOSQXMBqmsGWZUQ7rux7cj1du6d9rD6C8
ze1B2eQjkrGkIL/OF1s7vSmgYVafsRoZd/IHUrkoQvX8FZwUsmPu7amgBfaY3g+d
q1x0jNGKb6I6Bzdl6LgMD9qxp+3i7GQOnd9J8LFSietY6Z4jUBzVoOoz8iAU84OF
h2HhAuiPw1ai0VnY38RTI+8kepGWVfGxfBWzwH9uIjeooIeaosVFvE8cmYUB4TSH
5dUyD0jHct2+8ceKEtIoFU/FfHq/mDaVnvcDCZXtIgitdMFQdMZaVehmObyhRdDD
4NQCs0gaI9AAgFj4L9QtkARzhQLNyRf87Kln+YU0lgCGr9HLg3rGO8q+Y4ppLsOd
unQZ6ZxPNGIfOApbPVf5hCe58EZwiWdHIMn9lPP6+F404y8NNugbQixBber+x536
WrZhFZLjEkhp7fFXf9r32rNPfb74X/U90Bdy4lzp3+X1ukh1BuMxA/EEhDoTOS3l
7ABvc7BYSQubQ2490OcdkIzUh3ZwDrakMVrbaTxUM2p24N6dB+ns2zptWCva6jzW
r8IWKIMxzxLPv5Kt3ePKcUdvkBU/smqujSczTzzSjIoR5QqQA6lN1ZRSnuHIWCvh
JEltkYnTAH41QJ6SAWO66GrrUESwN/cgZzL4JLEqz1Y=
-----END CERTIFICATE-----
-----BEGIN CERTIFICATE-----
MIIDkDCCAxegAwIBAgIUaHsIgxrdrDnMKCPwQBsy5/2gEVYwCgYIKoZIzj0EAwMw
eDELMAkGA1UEBhMCRVMxETAPBgNVBAoMCEZOTVQtUkNNMQ4wDAYDVQQLDAVDZXJl
czEYMBYGA1UEYQwPVkFURVMtUTI4MjYwMDRKMSwwKgYDVQQDDCNBQyBSQUlaIEZO
TVQtUkNNIFNFUlZJRE9SRVMgU0VHVVJPUzAeFw0yNTEyMTYxMDQ0MDNaFw0zNTEy
MTQxMDQ0MDJaMGMxCzAJBgNVBAYTAkVTMREwDwYDVQQKDAhGTk1ULVJDTTEYMBYG
A1UEYQwPVkFURVMtUTI4MjYwMDRKMScwJQYDVQQDDB5BQyBTRVJWSURPUkVTIFNF
R1VST1MgVElQTzIgRzIwdjAQBgcqhkjOPQIBBgUrgQQAIgNiAARwl+9qu5Its/0+
FLNvLi4KBBXQfITkQd24UOZEJrvpzEBvmTAJtidl1YIJuFgQxI1Mq7ueOEAu1arR
BKw/IT3x9VhdXk4wIn2yj4IjrZQKPSNSFa5WIwA4gzuo1tvMSRujggF1MIIBcTAS
BgNVHRMBAf8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFAG5L++/EYZg8k/QQW6rcx/n
0m5JMIGaBggrBgEFBQcBAQSBjTCBijBFBggrBgEFBQcwAoY5aHR0cDovL3d3dy5j
ZXJ0LmZubXQuZXMvY2VydHMvQUNSQUlaU0VSVklET1JFU1NFR1VST1MuY3J0MEEG
CCsGAQUFBzABhjVodHRwOi8vb2NzcGZubXRzc3IuY2VydC5mbm10LmVzL29jc3Bz
c3IvT2NzcFJlc3BvbmRlcjARBgNVHSAECjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYI
KwYBBQUHAwEwRgYDVR0fBD8wPTA7oDmgN4Y1aHR0cDovL3d3dy5jZXJ0LmZubXQu
ZXMvY3Jscy9BUkxTRVJWSURPUkVTU0VHVVJPUy5jcmwwHQYDVR0OBBYEFEJKN+hE
rOw5bTWY4aoWL9gIp9oxMA4GA1UdDwEB/wQEAwIBBjAKBggqhkjOPQQDAwNnADBk
AjBg8RX7HVLA9LPPgFSBhR7IaKmmC4gOxMcZY5cp1Le48MscEyCLj3FqzX3rrEHe
PsMCMGWFuZFZU513iVJxE3th9IDdDpJ9uE04KOySHRHOtOmx3HE4Uae1Mw6mPn6W
BZawzQ==
-----END CERTIFICATE-----
"""


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
_LOG_HOOKS = []


def log(msg):
    if sys.stderr is not None:
        try:
            print(f"[{datetime.now():%H:%M:%S}] {msg}", file=sys.stderr, flush=True)
        except Exception:
            pass
    for h in list(_LOG_HOOKS):
        try:
            h(msg)
        except Exception:
            pass


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).lower()
    return "".join(c for c in s if not unicodedata.combining(c)).strip()


def code8(c):
    c = str(c or "").strip()
    if re.fullmatch(r"\d+\.0", c):
        c = c[:-2]
    return c.zfill(8) if c.isdigit() and 7 <= len(c) <= 8 else c


PHONE_RE = re.compile(r"(?<!\d)[6789](?:[\s.\-/]?\d){8}(?!\d)")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def phone(v):
    s = str(v or "").strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"^(\+|00)\s*34", "", s)
    m = PHONE_RE.search(s)
    return re.sub(r"\D", "", m.group(0)) if m else ""


def email(v):
    m = EMAIL_RE.search(str(v or ""))
    return m.group(0).lower().rstrip(".") if m else ""


def web(v):
    s = str(v or "").strip()
    if not s or norm(s) in ("-", "n/a", "nan", "none") or "." not in s or " " in s:
        return ""
    return s


def titularidad(v, conc=""):
    s, c = re.sub(r"^centro\s+", "", norm(v)), norm(conc)
    if "concert" in s or s.startswith("priv. conc") or c in ("s", "si", "con niveles concertados"):
        return "Privada concertada"
    if s.startswith("pub") or "public" in s or s in ("p", "depart. educacion", "otros publicos"):
        return "Pública"
    if s.startswith("priv") or s in ("privat", "n"):
        return "Privada"
    return ""


MINUS = {"de", "del", "la", "las", "los", "el", "y", "e", "i", "o", "u", "en", "a", "al", "da", "do", "das", "dos",
         "per", "amb", "sa", "es", "ses", "na", "en"}
SIGLAS = {"CEIP", "IES", "IESO", "CEO", "CEE", "CEPA", "EOI", "CIFP", "CPR", "CPEIPS", "CPEIP", "CRA", "CC", "EI", "EEI",
          "EIM", "CEPER", "SEP", "CEIPSO", "CER", "CEI", "CP", "CEP", "ESO", "FP", "EA", "EM", "CDP", "CPI", "CRIE", "EBI",
          "CEPI", "CPEPS", "CPES", "CPEI", "CPC", "HLHI", "BHI", "LHI", "EPA", "EIC", "EMPU", "IKT", "CIP", "CPIFP",
          "CPA", "CEAD", "ESMU", "EMUS", "CPM", "CSM", "EASD", "ESAD", "EOEP", "II", "III", "IV", "VI", "VII", "XXI", "SL",
          "SA", "SLU", "SAL", "SCL", "AMPA", "APA", "ONCE", "CEU", "UNED", "S/N", "SN", "N.º", "Nº"}


def titulo(v):
    """Pasa a formato título solo los textos que vienen en MAYÚSCULAS ('AINA MOLL I MARQUÉS' -> 'Aina Moll i Marqués')."""
    s = re.sub(r"\s+", " ", str(v or "")).strip()
    letras = [c for c in s if c.isalpha() and c not in "ºª"]
    if not letras or any(c.islower() for c in letras):
        return s
    out = []
    for i, w in enumerate(s.split(" ")):
        lw = w.lower()
        if w.strip(".,;:()") in SIGLAS or any(ch.isdigit() for ch in w):
            out.append(w)
        elif lw == "i" and i and i == len(s.split(" ")) - 1:
            out.append("I")  # 'JAUME I' (romano al final); en medio es la conjunción catalana 'i'
        elif i and lw in MINUS:
            out.append(lw)
        else:
            out.append(re.sub(r"(^|[-('/.\"«])(\w)", lambda m: m.group(1) + m.group(2).upper(), lw))
    return " ".join(out)


capital = titulo

TIPOS_CAT = [("Institut Escola", "Instituto Escuela"), ("Institut", "Instituto"), ("Escola Bressol", "Escuela Infantil (0-3 años)"),
             ("EB ", "Escuela Infantil (0-3 años)"), ("Llar d'infants", "Escuela Infantil (0-3 años)"),
             ("Llar inf", "Escuela Infantil (0-3 años)"), ("LLI ", "Escuela Infantil (0-3 años)"),
             ("Escola de Música", "Escuela de Música"), ("Esc. de Música", "Escuela de Música"), ("Esc. Mús", "Escuela de Música"),
             ("Escola de Dansa", "Escuela de Danza"), ("Esc. de Dansa", "Escuela de Danza"), ("Escola d'Art", "Escuela de Arte"),
             ("EA ", "Escuela de Arte"), ("Escola Oficial d'Idiomes", "Escuela Oficial de Idiomas"),
             ("EOI ", "Escuela Oficial de Idiomas"), ("CFA ", "Centro de Formación de Adultos"),
             ("Centre de Formació d'Adults", "Centro de Formación de Adultos"), ("AFA ", "Aula de Formación de Adultos"),
             ("CEE ", "Centro de Educación Especial"), ("Escola d'Educació Especial", "Centro de Educación Especial"),
             ("Conservatori", "Conservatorio"), ("Cons. ", "Conservatorio"), ("ZER ", "Zona Escolar Rural")]


def tipo_cat(nombre, naturalesa):
    for pre, t in TIPOS_CAT:
        if nombre.startswith(pre):
            return t
    publico = norm(naturalesa).startswith("public")
    if nombre.startswith(("Escola", "Esc.")) and publico:
        return "Escuela (Infantil y Primaria)"
    return "Centro público" if publico else "Centro privado"


def decode(raw):
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("latin-1")


def read_csv(raw, delim=None):
    t = decode(raw)
    first = t.split("\n", 1)[0]
    d = delim or (";" if first.count(";") > first.count(",") else ",")
    return list(csv.DictReader(io.StringIO(t), delimiter=d))


def read_xls(raw):
    import xlrd  # .xls binario (Ministerio, Asturias)
    sh = xlrd.open_workbook(file_contents=raw).sheet_by_index(0)
    h = [str(x).strip() for x in sh.row_values(0)]
    return [dict(zip(h, sh.row_values(i))) for i in range(1, sh.nrows)]


def read_xlsx(raw):
    from openpyxl import load_workbook
    ws = load_workbook(io.BytesIO(raw), read_only=True, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    h = [str(x or "").strip() for x in next(it)]
    return [{k: ("" if v is None else str(v)) for k, v in zip(h, row)} for row in it]


# ---------------------------------------------------------------------------
# HTTP con reintentos, cookies por hilo y certificados intermedios
# ---------------------------------------------------------------------------
class Http:
    def __init__(self, timeout=120):
        self.ctx = ssl.create_default_context()
        try:
            self.ctx.load_verify_locations(cadata=EXTRA_CERTS)
        except (ssl.SSLError, ValueError):
            pass
        self.timeout = timeout
        self.local = threading.local()

    def _opener(self):
        op = getattr(self.local, "op", None)
        if op is None:
            op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()),
                                             urllib.request.HTTPSHandler(context=self.ctx))
            self.local.op = op
        return op

    def fetch(self, url, data=None, timeout=None, retries=3):
        headers = {"User-Agent": UA, "Accept": "*/*", "Accept-Language": "es-ES,es;q=0.9"}
        body = None
        if data is not None:
            body = urllib.parse.urlencode(data).encode()
            headers["Content-Type"] = "application/x-www-form-urlencoded"
        last = None
        for i in range(retries):
            try:
                req = urllib.request.Request(url, data=body, headers=headers)
                with self._opener().open(req, timeout=timeout or self.timeout) as r:
                    return r.read()
            except urllib.error.HTTPError as e:
                last = e
                if e.code < 500 and e.code not in (408, 429):
                    raise
            except (urllib.error.URLError, http.client.HTTPException, ConnectionError, TimeoutError, OSError) as e:
                last = e
            time.sleep(2 * (i + 1) ** 2)
        raise last

    def warm_rcd(self):
        """Primera visita al buscador del Ministerio en este hilo (cookie de sesión)."""
        if not getattr(self.local, "rcd", False):
            try:
                self.fetch(RCD_HOME, timeout=30, retries=2)
            except Exception:
                pass
            self.local.rcd = True


class Store:
    """Caché en disco de las descargas grandes. En modo carpeta (--desde-carpeta) no usa la red."""

    def __init__(self, folder, http, max_age_h=24, refresh=False, offline=False):
        self.folder, self.http, self.offline = folder, http, offline
        self.max_age = float("inf") if offline else max_age_h * 3600
        self.refresh = refresh and not offline
        os.makedirs(folder, exist_ok=True)

    def get(self, name, url_fn, data=None):
        p = os.path.join(self.folder, name)
        if os.path.exists(p) and not self.refresh and time.time() - os.path.getmtime(p) < self.max_age:
            with open(p, "rb") as f:
                return f.read()
        if self.offline:
            raise FileNotFoundError(f"falta {name} en {self.folder}")
        url = url_fn() if callable(url_fn) else url_fn
        raw = self.http.fetch(url, data=data)
        tmp = p + ".tmp"
        with open(tmp, "wb") as f:
            f.write(raw)
        os.replace(tmp, p)
        return raw


# ---------------------------------------------------------------------------
# Lectores de cada comunidad -> lista de registros normalizados
# ---------------------------------------------------------------------------
def rec(code, nombre="", tipo="", dir="", cp="", loc="", mun="", tel="", mail="", url="", nat="", conc="", sigla=""):
    return {"code": code, "nombre": str(nombre or "").strip(), "tipo": str(tipo or "").strip(),
            "dir": re.sub(r"\s+", " ", str(dir or "")).strip(), "cp": str(cp or "").strip().split(".")[0],
            "loc": str(loc or "").strip(), "mun": str(mun or "").strip(), "tel_src": tel, "email_src": mail,
            "web_src": url, "nat_src": nat, "conc_src": conc, "sigla": str(sigla or "").strip()}


def ckan_resources(st, url):
    return json.loads(st.http.fetch(url))["result"]["resources"]


def load_and(st):
    def url():
        try:
            rs = ckan_resources(st, "https://www.juntadeandalucia.es/datosabiertos/portal/api/3/action/package_show"
                                    "?id=directorio-de-centros-docentes-de-andalucia")
            cursos = [(m.group(1), r) for r in rs if (m := re.search(r"Curso (\d{4}/\d{4})", r.get("name", "")))]
            u = max(cursos)[1]["url"]
            return re.sub(r"^https://[^/]+", "https://www.juntadeandalucia.es", u)
        except Exception:
            return ("https://www.juntadeandalucia.es/datosabiertos/portal/dataset/e039df22-4b82-4d0d-9884-0ab5952e24e4/"
                    "resource/b5924e81-0b53-4418-9d93-b1f39ba1ef65/download/da_centros.csv")
    rows = read_csv(st.get("and.csv", url), ";")
    return [rec(code8(r["codigo"]), r["D_ESPECIFICA"], r["D_DENOMINA"], r["D_DOMICILIO"], r["C_POSTAL"], r["D_LOCALIDAD"],
                r["D_MUNICIPIO"], r["N_TELEFONO"], r["Correo_e"], "", r["D_TIPO"]) for r in rows]


def load_ara(st):
    rows = read_csv(st.get("ara.csv", "https://centrosdocentes.catedu.es/ee/php/export/export_centros.php"), ";")
    return [rec(code8(r["codcentro"]), titulo(r["nombre"]), r["tipo_centro"], titulo(r["direccion"]), r["cod_postal"],
                titulo(r["localidad"]), titulo(r["municipio"]), r["telefono"], r["email"], r["web"], r["naturaleza"]) for r in rows]


def load_ast(st):
    def url():
        pag = "https://www.educastur.es/directorio-de-centros"
        h = decode(st.http.fetch(pag))
        return urllib.parse.urljoin(pag, html.unescape(re.search(r'href="([^"]+\.xls[^"]*)"', h).group(1)))
    rows = read_xls(st.get("ast.xls", url))
    return [rec(code8(r["CÓDIGO"]), r["CENTRO"], r["TIPO"], r["DOMICILIO"], r["C_POSTAL"], r["LOCALIDAD"], r["MUNICIPIO"],
                r["N_TELEFONO"], r["CORREO_E"], "", r["TITULARIDAD"], r["CONCIERTO"]) for r in rows]


def load_bal(st):
    def url():
        pag = ("https://intranet.caib.es/sites/planificacioicentres/es/"
               "llistat_de_centres_educatius_de_les_illes_balears/")
        h = decode(st.http.fetch(pag))
        return urllib.parse.urljoin(pag, html.unescape(re.search(r'href="(archivopub\.do\?[^"]+)">Listado', h).group(1)))
    rows = read_xlsx(st.get("bal.xlsx", url))
    return [rec(code8(r["CODI"]), titulo(r["NOM"]), r["DENOMINACIÓ GENÈRICA"], titulo(r["ADREÇA"]), r["CP"], titulo(r["LOCALITAT"]), titulo(r["MUNICIPI"]),
                "", r["CORREU ELECTRÒNIC 1"], r["PÀGINA WEB"],
                "Pública" if "CONSELLERIA" in r["TITULAR"].upper() else "", sigla=r["DENOMINACIÓ GENÈRICA"]) for r in rows]


def load_can(st):
    rows = read_csv(st.get("can.csv", "https://datos.canarias.es/catalogos/general/dataset/f6b15811-014b-46f7-a858-"
                                      "fe48b062ed05/resource/b5e08adf-841b-4ba5-a599-4339e772d792/download/centros.csv"), ",")
    return [rec(code8(r["Codigo"]), titulo(r["Denominacion"]), r["DescripcionEtapaCentro"], titulo(r["Direccion"]), r["CodigoPostal"],
                titulo(r["Localidad"]), titulo(r["Municipio"]), r["Telefono"], r["CorreoElectronico"] or r["SegundoCorreoElectronico"],
                r["PaginaWeb"], r["Naturaleza"], r["Concierto"]) for r in rows]


def load_cyl(st):
    rows = read_csv(st.get("cyl.csv", "https://directorio.educa.jcyl.es/es/datos-abiertos-centros-docentes"), ";")
    out = []
    for r in rows:
        if r.get("SITUACIÓN") and norm(r["SITUACIÓN"]) != "alta":
            continue
        num = r.get("NÚMERO") or r.get("NÚMERO(EXT)") or ""
        d = " ".join(x for x in (r.get("VÍA", ""), r.get("NOMBRE DE LA VÍA", ""), num) if x)
        out.append(rec(code8(r["CÓDIGO"]), titulo(r["DENOMINACIÓN ESPECÍFICA"]), titulo(r["DENOMINACIÓN GENÉRICA"]), titulo(d), r["C.POSTAL"],
                       titulo(r["LOCALIDAD"]), titulo(r["MUNICIPIO"]), r["TELÉFONO"], r["CORREO ELECTRÓNICO"], r["WEB"], r["NATURALEZA"],
                       r["CONCIERTO"]))
    return out


def load_cat(st):
    api = "https://analisi.transparenciacatalunya.cat/resource/kvmv-ahh4"

    def url():
        q = urllib.parse.urlencode({"$select": "max(curs) AS curs"})
        curs = json.loads(st.http.fetch(f"{api}.json?{q}"))[0]["curs"]
        return f"{api}.csv?" + urllib.parse.urlencode({"$where": f"curs='{curs}'", "$limit": "50000"})
    rows = read_csv(st.get("cat.csv", url), ",")
    return [rec(code8(r["codi_centre"]), r["denominaci_completa"], tipo_cat(r["denominaci_completa"], r["nom_naturalesa"]),
                r["adre_a"], r["codi_postal"], r["nom_localitat"],
                r["nom_municipi"], r["tel_fon"], r["e_mail_centre"], r.get("url", ""), r["nom_naturalesa"],
                sigla=r["denominaci_completa"]) for r in rows]


def load_val(st):
    def url():
        try:
            return ckan_resources(st, "https://dadesobertes.gva.es/api/3/action/package_show?id=edu-centros")[0]["url"]
        except Exception:
            return ("https://dadesobertes.gva.es/dataset/68eb1d94-76d3-4305-8507-e1aab7717d0e/resource/1aa53c3a-4639-41aa-"
                    "ac85-d58254c428c0/download/centros-docentes-de-la-comunitat-valenciana.csv")
    rows = read_csv(st.get("val.csv", url), ";")
    return [rec(code8(r["codigo"]), titulo(r["denominacion_especifica"] or r["denominacion"]), titulo(r["denominacion_generica_es"]),
                titulo(" ".join(x for x in (r["tipo_via"], r["direccion"], r["numero"]) if x)), r["codigo_postal"],
                titulo(r["localidad"]), titulo(r["localidad"]), r["telefono"], "", "", r["regimen"]) for r in rows]


def load_gal(st):
    base = "https://abertos.xunta.gal/catalogo/ensino-formacion/-/dataset/0257/centros-educativos-galicia/"
    out = []
    for n in ("001", "003", "005", "007"):
        for r in read_csv(st.get(f"gal_{n}.csv", f"{base}{n}/descarga-directa-ficheiro.csv"), ","):
            out.append(rec(code8(r["Código"]), r["Nome"], r["Tipo de centro"], r["Enderezo"], r["Cód. postal"], r["Concello"],
                           r["Concello"], r["Teléfono"], "", "", r["TITULARIDADE"], r["ENSINO_CONCERTADO"]))
    return out


def load_mad(st):
    url = ("https://idem.comunidad.madrid/geosrvg/wfs?" + urllib.parse.urlencode({
        "service": "WFS", "version": "2.0.0", "request": "GetFeature", "typeNames": "WPAD_V_CENTROS_GIS",
        "outputFormat": "application/json", "srsName": "EPSG:4326"}))
    feats = json.loads(st.get("mad.json", url))["features"]
    out = []
    for f in feats:
        p = f["properties"]
        out.append(rec(code8(p.get("CD_CENTRO")), titulo(p.get("DS_CENTRO")), titulo(p.get("DSEXTEND")), titulo(p.get("DOMICILIO")),
                       p.get("CDPOSTAL"), p.get("DSMUNI"), p.get("DSMUNI"), p.get("TELEFONO"), p.get("E_MAIL"),
                       p.get("WEB"), p.get("DS_NATURALEZA"), p.get("ITCONCIERTO"), sigla=p.get("DS_GEN")))
    return out


def load_mur(st):
    def url():
        try:
            return ckan_resources(st, "https://datosabiertos.regiondemurcia.es/api/3/action/package_show"
                                      "?id=sedes-de-los-centros-educativos-de-la-region-de-murcia")[0]["url"]
        except Exception:
            return "https://mapaescolar.murciaeduca.es/mapaescolar-api/api/centros"
    rows = json.loads(st.get("mur.json", url))
    nat = {"P": "Pública", "C": "Privada concertada", "N": "Privada"}
    return [rec(code8(r.get("codcen")), titulo(r.get("dencen")), r.get("denLarga"), titulo(r.get("domcen")), r.get("cpcen"), titulo(r.get("loccen")),
                titulo(r.get("muncen")), r.get("telcen"), r.get("email"), r.get("web"), nat.get(r.get("titularidad"), ""),
                sigla=r.get("denCorta")) for r in rows]


def load_nav(st):
    def url():
        try:
            rs = ckan_resources(st, "https://datosabiertos.navarra.es/api/3/action/package_show?id=centros-educativos")
            return max(rs, key=lambda r: r.get("name", ""))["url"]
        except Exception:
            return ("https://datosabiertos.navarra.es/dataset/37044915-b9fe-47f6-8e1e-f6bb9a3e1cc8/resource/68bf4158-9db6-"
                    "4dc4-9e37-a45f0bf12fbd/download/centros-educativos-2026-2027.csv")
    out, seen = [], set()
    for r in read_csv(st.get("nav.csv", url), ","):
        k = (r["Código"], r["Dirección"])
        if k in seen:
            continue
        seen.add(k)
        out.append(rec(code8(r["Código"]), titulo(r["Nombre"] or r["Nombre corto"]), "", r["Dirección"], r["CP"], r["Localidad"],
                       r["Localidad"], r["Telefono"], r["Mail"], r["URL"], r["Naturaleza"], sigla=r["Nombre corto"]))
    return out


TIPOS_PVS = {"CEIP": "Colegio de Educación Infantil y Primaria", "IES": "Instituto de Educación Secundaria",
             "CPEIPS": "Centro Privado de Educación Infantil, Primaria y Secundaria",
             "CPEIP": "Centro Privado de Educación Infantil y Primaria", "CPES": "Centro Privado de Educación Secundaria",
             "CPEI": "Centro Privado de Educación Infantil", "CIFP": "Centro Integrado de Formación Profesional",
             "CPIFP": "Centro Privado Integrado de Formación Profesional", "CEPA": "Centro de Educación de Personas Adultas",
             "EOI": "Escuela Oficial de Idiomas", "CEE": "Centro de Educación Especial",
             "CPEE": "Centro Privado de Educación Especial", "EEI": "Escuela de Educación Infantil",
             "EIC": "Escuela Infantil (Consorcio Haurreskolak)", "EIPR": "Escuela Infantil Privada",
             "EIMU": "Escuela Infantil Municipal", "EMPU": "Escuela de Música Pública", "EMPR": "Escuela de Música Privada",
             "CPM": "Conservatorio Profesional de Música", "EASD": "Escuela de Arte y Superior de Diseño"}


def load_pvs(st):
    rows = read_csv(st.get("pvs.csv", "https://opendata.euskadi.eus/contenidos/ds_localizaciones/"
                                      "centros_docentes_no_universit/es_euskadi/adjuntos/dirgennouniv.csv"), ";")
    out = []
    for r in rows:
        nat = "Pública" if str(r.get("DEPE", "")).startswith("1") else ("Privada" if r.get("DEPE") else "")
        sig = str(r.get("DGENRC") or "").strip()
        out.append(rec(str(r["CCEN"]).strip(), titulo(r["NOM"]), TIPOS_PVS.get(sig, sig), titulo(r["DOMI"]), r["CPOS"], titulo(r["DMUNIC"]),
                       capital(r["DMUNIC"]), r["TEL1"], r["EMAIL"], r["PAGINA"], nat, sigla=r["DGENRC"]))
    return out


LOADERS = {"and": load_and, "ara": load_ara, "ast": load_ast, "bal": load_bal, "can": load_can, "cyl": load_cyl,
           "cat": load_cat, "val": load_val, "gal": load_gal, "mad": load_mad, "mur": load_mur, "nav": load_nav,
           "pvs": load_pvs}
FUENTE_TXT = {
    "and": "Junta de Andalucía · Directorio de centros docentes (datos abiertos)",
    "ara": "Gobierno de Aragón · centrosdocentes.catedu.es",
    "ast": "Educastur · Directorio de centros",
    "bal": "Govern de les Illes Balears · Llistat de centres educatius",
    "can": "Gobierno de Canarias · datos.canarias.es",
    "cyl": "Junta de Castilla y León · Datos abiertos de centros docentes",
    "cat": "Generalitat de Catalunya · Directori de centres docents",
    "val": "Generalitat Valenciana · dadesobertes.gva.es",
    "gal": "Xunta de Galicia · abertos.xunta.gal",
    "mad": "Comunidad de Madrid · IDEM (WPAD_V_CENTROS_GIS)",
    "mur": "Región de Murcia · Mapa Escolar (murciaeduca)",
    "nav": "Gobierno de Navarra · datosabiertos.navarra.es",
    "pvs": "Gobierno Vasco · Open Data Euskadi",
}


# ---------------------------------------------------------------------------
# Registro Estatal de Centros Docentes (Ministerio)
# ---------------------------------------------------------------------------
def load_rcd(st, key):
    rid = CCAA[key]["rcd"]
    form = {"nivel": "0", "idComunidad": rid, "idProvincia": "00", "naturaleza": "0", "concertado": "0", "familia": "0",
            "ensenanza": "0", "modalidad": "0", "tipoCentro": "0", "provincial": "", "comarca": "", "pais": "",
            "localidad": "", "codCentro": "", "nombreCentro": "", "denominacion": "0", "ensenanzaFP": ""}
    rows = read_xls(st.get(f"rcd_{rid}.xls", RCD_EXPORT, data=form))
    if not rows:
        return {}
    # El listado termina con algunos centros "de ámbito estatal" de otras comunidades.
    from collections import Counter
    main = Counter(r["COMUNIDAD AUTÓNOMA"] for r in rows).most_common(1)[0][0]
    return {code8(r["CÓDIGO"]): r for r in rows if r["COMUNIDAD AUTÓNOMA"] == main}


def rcd_as_records(rcd):
    out = []
    for c, r in rcd.items():
        out.append(rec(c, capital(r["DENOMINACIÓN ESPECÍFICA"]), r["DENOMINACIÓN GENÉRICA"], capital(r["DOMICILIO"]),
                       str(r["CÓD POSTAL"]).split(".")[0].zfill(5) if str(r["CÓD POSTAL"]).strip() else "",
                       capital(r["LOCALIDAD"]), capital(r["LOCALIDAD"]), "", "", "", ""))
    return out


NAT_RCD = {"centro publico": "Pública", "centro privado": "Privada", "centro concertado": "Privada concertada"}
FICHA_CAMPO = r"<b>\s*{}:?\s*</b>(?:&nbsp;|\s)*(?:<span>([^<]*)</span>|<a [^>]*>([^<]*)</a>)"


def parse_ficha(h):
    def campo(label):
        m = re.search(FICHA_CAMPO.format(label), h, re.I)
        return html.unescape((m.group(1) or m.group(2) or "")).strip() if m else ""
    return {"tel": phone(campo("Tel[ée]fono")), "email": email(campo("(?:Direcci[óo]n de )?[Cc]orreo electr[óo]nico")),
            "web": web(campo("P[áa]gina [Ww]eb del [Cc]entro"))}


class Fichas:
    """Fichas individuales del Ministerio con caché persistente (permite reanudar)."""

    def __init__(self, path, http, max_age_dias=30, offline=False):
        self.path, self.http, self.offline = path, http, offline
        self.max_age = max_age_dias * 86400
        self.lock = threading.Lock()
        self.data = {}
        if os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    self.data = json.load(f)
            except (OSError, ValueError):
                self.data = {}

    def fresh(self, code):
        d = self.data.get(code)
        return d is not None and time.time() - d.get("t", 0) < self.max_age

    def save(self):
        with self.lock:
            tmp = self.path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False)
            os.replace(tmp, self.path)

    def fetch_one(self, code, rid, pause, cancel=None):
        if cancel is not None and cancel.is_set():
            return None
        self.http.warm_rcd()
        h = decode(self.http.fetch(RCD_FICHA, data={"codCentro": code, "idComunidad": rid}, timeout=45))
        d = parse_ficha(h)
        d["t"] = time.time()
        with self.lock:
            self.data[code] = d
        time.sleep(pause)
        return d

    def run(self, todo, hilos, pause, cancel=None, on_progress=None):
        """todo: lista de (code, rid). Devuelve nº de fichas descargadas."""
        todo = [(c, r) for c, r in todo if not self.fresh(c)]
        if not todo or self.offline:
            return 0
        log(f"Fichas del Ministerio: {len(todo)} centros ({hilos} conexiones). Se guardan en {self.path}")
        done, errors, t0, last = 0, 0, time.time(), time.time()
        ex = ThreadPoolExecutor(max_workers=hilos)
        try:
            futs = [ex.submit(self.fetch_one, c, r, pause, cancel) for c, r in todo]
            for f in as_completed(futs):
                try:
                    f.result()
                except Exception:
                    errors += 1
                done += 1
                rate = done / max(1e-6, time.time() - t0)
                eta = (len(todo) - done) / max(rate, 1e-6)
                if on_progress:
                    on_progress(done, len(todo), eta)
                if done % 200 == 0:
                    self.save()
                if cancel is not None and cancel.is_set():
                    ex.shutdown(wait=False, cancel_futures=True)
                    log("Cancelado: se guardan las fichas descargadas hasta ahora.")
                    break
                if time.time() - last > 10 or done == len(todo):
                    last = time.time()
                    log(f"  fichas {done}/{len(todo)} ({done / len(todo):.0%}) · {rate:.1f}/s · quedan ~{eta / 60:.0f} min"
                        + (f" · errores {errors}" if errors else ""))
        except KeyboardInterrupt:
            ex.shutdown(wait=False, cancel_futures=True)
            self.save()
            raise
        ex.shutdown(wait=True)
        self.save()
        return done - errors


# ---------------------------------------------------------------------------
# Construcción del listado
# ---------------------------------------------------------------------------
INST = ("juntadeandalucia.es", "xtec.cat", "educa.jcyl.es", "gobiernodecanarias.org", "murciaeduca.es", "hezkuntza.net",
        "edu.gva.es")


def revisar(code, e):
    if not e:
        return ""
    local, dom = e.split("@", 1)
    dig = re.sub(r"\D", "", local)
    if re.fullmatch(r"\d{7,8}\.edu", local) and dom != "juntadeandalucia.es":
        return "Dominio del email posiblemente erróneo"
    if dom in INST and len(dig) >= 5 and not (code in dig or code[1:] in dig or dig in code):
        return "El email no coincide con el código del centro"
    return ""


def no_docente(r, key=""):
    sig = r["sigla"].strip()
    if sig.upper() in SIGLAS_NO_DOCENTES.get(key, ()) or sig.startswith(PREFIJOS_NO_DOCENTES.get(key, ("\0",))):
        return True
    t = norm(r["tipo"]) + " | " + norm(r["nombre"])
    return any(k in t for k in NO_DOCENTES)


def provincia(key, code, cp):
    p = cp[:2] if re.fullmatch(r"\d{5}", cp) else ""
    ok = PROV_CCAA[key]
    if p not in ok and len(code) == 8 and code[:2] in ok:
        p = code[:2]
    if p not in ok:
        p = next(iter(ok)) if len(ok) == 1 else ""
    return PROV.get(p, "")


def construir(key, st, incluir_no_docentes, info):
    nombre = CCAA[key]["nombre"]
    try:
        rcd = load_rcd(st, key)
    except Exception as e:
        log(f"{nombre}: no se pudo leer el Registro Estatal ({e}); sigo sin él.")
        rcd = {}
    regs, fuente = None, ""
    if key in LOADERS:
        try:
            regs = LOADERS[key](st)
            fuente = FUENTE_TXT[key]
        except Exception as e:
            log(f"{nombre}: falló el directorio autonómico ({type(e).__name__}: {e}). Uso el Registro Estatal.")
    degradada = regs is None and key in LOADERS
    if regs is None:
        if not rcd and key in LOADERS:
            raise RuntimeError("no se pudo descargar ni el directorio autonómico ni el Registro Estatal")
        regs = rcd_as_records(rcd)
        fuente = "Registro Estatal de Centros Docentes (Ministerio)"
    info[key] = {"fuente": fuente, "n_fuente": len(regs), "rcd": len(rcd), "degradada": degradada}

    out = []
    for r in regs:
        if not r["code"]:
            continue
        if not incluir_no_docentes and no_docente(r, key):
            continue
        m = rcd.get(r["code"])
        tipo = (str(m["DENOMINACIÓN GENÉRICA"]).strip() if m else "") or r["tipo"]
        nat = NAT_RCD.get(norm(m["NATURALEZA"]), "") if m else ""
        nat = nat or titularidad(r["nat_src"], r["conc_src"])
        t_a, t_m = phone(r["tel_src"]), phone(m.get("TELÉFONO")) if m else ""
        e_a = email(r["email_src"])
        out.append({
            "ccaa": nombre, "key": key, "prov": provincia(key, r["code"], r["cp"]), "code": r["code"],
            "nombre": r["nombre"], "tipo": tipo, "nat": nat, "dir": r["dir"], "cp": r["cp"], "loc": r["loc"], "mun": r["mun"],
            "tel": t_a or t_m, "email": e_a, "web": web(r["web_src"]),
            "src_tel": "aut" if t_a else ("rcd" if t_m else ""), "src_email": "aut" if e_a else "",
        })
    return out


def aplicar_fichas(rows, fichas):
    for r in rows:
        d = fichas.data.get(r["code"])
        if not d:
            continue
        if not r["email"] and d.get("email"):
            r["email"], r["src_email"] = d["email"], "ficha"
        if not r["tel"] and d.get("tel"):
            r["tel"], r["src_tel"] = d["tel"], "ficha"
        if not r["web"] and d.get("web"):
            r["web"] = d["web"]


FUENTE_NOMBRE = {"aut": "directorio autonómico", "rcd": "Registro Estatal (listado)", "ficha": "Registro Estatal (ficha)"}


def fuente_contacto(r):
    a, b = r["src_tel"], r["src_email"]
    if not a and not b:
        return ""
    if a == b or not b or not a:
        t = FUENTE_NOMBRE[a or b]
        return t[:1].upper() + t[1:]
    return f"Teléfono: {FUENTE_NOMBRE[a]} · Email: {FUENTE_NOMBRE[b]}"


# ---------------------------------------------------------------------------
# Salida
# ---------------------------------------------------------------------------
COLS = ["Comunidad autónoma", "Provincia", "Código centro", "Nombre", "Tipo de centro", "Titularidad", "Dirección", "CP",
        "Localidad", "Municipio", "Teléfono", "Email", "Web", "Fuente contacto", "Revisar"]


def fila(r):
    return [r["ccaa"], r["prov"], r["code"], r["nombre"], r["tipo"], r["nat"], r["dir"], r["cp"], r["loc"], r["mun"],
            r["tel"], r["email"], r["web"], r.get("fuente") or fuente_contacto(r), r.get("revisar") or revisar(r["code"], r["email"])]


def escribir_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(COLS)
        for r in rows:
            w.writerow(fila(r))


def escribir_xlsx(path, rows, info):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.table import Table, TableStyleInfo

    A = "Arial"
    fn, fb = Font(name=A, size=10), Font(name=A, size=10, bold=True)
    fh, fill = Font(name=A, size=10, bold=True, color="FFFFFF"), PatternFill("solid", start_color="1F4E78")
    wb = Workbook()
    ws = wb.active
    ws.title = "Centros"
    ws.append(COLS)
    for r in rows:
        ws.append([v if v != "" else None for v in fila(r)])
    n = len(rows)
    for c in ws[1]:
        c.font = fb
    for row in ws.iter_rows(min_row=2, max_row=n + 1):
        for c in row:
            c.font = fn
        row[2].number_format = row[7].number_format = row[10].number_format = "@"
    for col, w in zip("ABCDEFGHIJKLMNO", [20, 22, 12, 42, 40, 18, 40, 8, 24, 24, 12, 36, 34, 34, 34]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    if n:
        t = Table(displayName="Centros", ref=f"A1:O{n + 1}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)

    wr = wb.create_sheet("Resumen")
    wr.append(["Comunidad / ciudad autónoma", "Nº de centros", "Con teléfono", "% teléfono", "Con email", "% email",
               "Públicos", "Privados", "Privados concertados"])
    end = max(n + 1, 2)
    rA, rF, rK, rL = (f"Centros!${c}$2:${c}${end}" for c in "AFKL")
    orden = sorted({r["ccaa"] for r in rows}, key=norm)
    for i, cc in enumerate(orden, 2):
        wr.cell(row=i, column=1, value=cc)
        wr.cell(row=i, column=2, value=f"=COUNTIF({rA},A{i})")
        wr.cell(row=i, column=3, value=f'=COUNTIFS({rA},A{i},{rK},"<>")')
        wr.cell(row=i, column=4, value=f"=IF(B{i}=0,0,C{i}/B{i})")
        wr.cell(row=i, column=5, value=f'=COUNTIFS({rA},A{i},{rL},"<>")')
        wr.cell(row=i, column=6, value=f"=IF(B{i}=0,0,E{i}/B{i})")
        wr.cell(row=i, column=7, value=f'=COUNTIFS({rA},A{i},{rF},"Pública")')
        wr.cell(row=i, column=8, value=f'=COUNTIFS({rA},A{i},{rF},"Privada")')
        wr.cell(row=i, column=9, value=f'=COUNTIFS({rA},A{i},{rF},"Privada concertada")')
    t = len(orden) + 2
    wr.cell(row=t, column=1, value="Total")
    for col in "BCEGHI":
        wr[f"{col}{t}"] = f"=SUM({col}2:{col}{t - 1})"
    wr[f"D{t}"], wr[f"F{t}"] = f"=IF(B{t}=0,0,C{t}/B{t})", f"=IF(B{t}=0,0,E{t}/B{t})"
    for row in wr.iter_rows(min_row=2, max_row=t):
        for c in row:
            c.font = fb if c.row == t else fn
            c.number_format = "0.0%" if c.column in (4, 6) else ("#,##0" if c.column > 1 else "General")
    for c in wr[1]:
        c.font, c.fill, c.alignment = fh, fill, Alignment(wrap_text=True, vertical="center")
    for col, w in zip("ABCDEFGHI", [30, 13, 13, 11, 12, 10, 11, 11, 13]):
        wr.column_dimensions[col].width = w

    wf = wb.create_sheet("Fuentes")
    wf.append(["Comunidad / ciudad autónoma", "Directorio usado", "Centros en la fuente", "Centros en el Registro Estatal"])
    presentes = {r["key"] for r in rows}
    for i, k in enumerate(sorted((k for k in info if k in presentes), key=lambda k: norm(CCAA[k]["nombre"])), 2):
        d = info[k]
        for j, v in enumerate([CCAA[k]["nombre"], d["fuente"], d["n_fuente"], d["rcd"]], 1):
            wf.cell(row=i, column=j, value=v).font = fn
    nt = len(presentes) + 3
    notas = [
        f"Generado el {datetime.now():%d/%m/%Y %H:%M} con centros_educativos.py v{VERSION}.",
        "Teléfono y email: directorio autonómico; si falta, Registro Estatal de Centros Docentes "
        "(listado y ficha de cada centro, educacion.gob.es/centros).",
        "Titularidad y tipo de centro: Registro Estatal; si el centro no aparece allí, el directorio autonómico.",
        "La columna Revisar marca emails que parecen tener una errata en origen.",
        "Contactos institucionales públicos. Para enviar publicidad por email, la LSSI (art. 21) exige "
        "consentimiento previo del destinatario.",
    ]
    for j, t_ in enumerate(notas):
        wf.cell(row=nt + j, column=1, value=t_).font = Font(name=A, size=9, italic=True)
    for c in wf[1]:
        c.font, c.fill = fh, fill
    for col, w in zip("ABCD", [28, 70, 18, 26]):
        wf.column_dimensions[col].width = w
    wb.save(path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def resolver_ccaa(valores):
    """Admite códigos y nombres de varias palabras: 'castilla y leon madrid', 'bal,mad', 'todas'."""
    tokens = norm(" ".join(valores or [])).replace(",", " ").replace(";", " ").split()
    if not tokens:
        return []
    if any(t in ("todas", "all", "espana", "*") for t in tokens):
        return list(CCAA)
    nombres = sorted({(n, k) for k, d in CCAA.items() for n in [k, norm(d["nombre"])] + d["alias"]},
                     key=lambda x: -len(x[0].split()))
    claves, i = [], 0
    while i < len(tokens):
        for n, k in nombres:
            partes = n.split()
            if tokens[i:i + len(partes)] == partes:
                if k not in claves:
                    claves.append(k)
                i += len(partes)
                break
        else:
            raise SystemExit(f"Comunidad desconocida: '{tokens[i]}'. Usa --listar para ver las opciones.")
    return claves


def autotest():
    assert phone("928-478864") == "928478864"
    assert phone("968680031.") == "968680031"
    assert phone("+34 971 12 34 56 / 971654321") == "971123456"
    assert phone("33021902.0") == "" and phone("") == ""
    assert email("04001187.edu@ieslosangeles.es ieslosangeles@gmail.com") == "04001187.edu@ieslosangeles.es"
    assert email("alcaldia@lamojonera.es        migutier@lamojonera.es") == "alcaldia@lamojonera.es"
    assert code8("7002774") == "07002774" and code8("33021902.0") == "33021902" and code8("010002") == "010002"
    assert titularidad("PRIV. CONC.") == "Privada concertada" and titularidad("Públic") == "Pública"
    assert titularidad("Privado", "S") == "Privada concertada" and titularidad("CENTRO PRIVADO") == "Privada"
    assert revisar("04003721", "0403721.edu@juntadeandalucia.es") == "El email no coincide con el código del centro"
    assert revisar("11013314", "11013314.edu@juntadeandalucia.com") == "Dominio del email posiblemente erróneo"
    assert revisar("08005771", "a8005771@xtec.cat") == ""
    ficha = ('<div class="col-md-6"><b>Teléfono:</b>&nbsp;&nbsp;<span>924 020508</span></div>'
             '<div class="col-md-6"><b>Fax:</b>&nbsp;&nbsp;<span>924 020509 / 05</span></div>'
             '<div class="col-md-6"><b>Correo electrónico:</b>&nbsp;&nbsp;<span>cp.joaquintenaartigas@edu.juntaextremadura.net</span></div>'
             '<div class="col-md-6"><b>Página Web del centro:</b>&nbsp;&nbsp;<a href="http://x.net" target="_blank">x.net</a></div>')
    assert parse_ficha(ficha) == {"tel": "924020508", "email": "cp.joaquintenaartigas@edu.juntaextremadura.net", "web": "x.net"}
    vacia = ('<b>Teléfono:</b>&nbsp;&nbsp;<span></span><b>Fax:</b>&nbsp;&nbsp;<span></span>'
             '<b>Correo electrónico:</b>&nbsp;&nbsp;<span></span><b>Página Web del centro:</b>&nbsp;&nbsp;<span></span>')
    assert parse_ficha(vacia) == {"tel": "", "email": "", "web": ""}
    assert resolver_ccaa(["Baleares", "madrid,cv"]) == ["bal", "mad", "val"]
    assert resolver_ccaa(["todas"]) == list(CCAA)
    assert resolver_ccaa(["castilla", "y", "leon", "Comunidad", "Valenciana", "la", "rioja", "castilla-la", "mancha"]) == ["cyl", "val", "rio", "clm"]
    assert no_docente(rec("07000001", "EOEP Palma", "EOEP", sigla="EOEP"), "bal")
    assert not no_docente(rec("07000002", "Aina Moll", "CEIP", sigla="CEIP"), "bal")
    assert no_docente(rec("28000001", "X", "Equipo de Orientación Educativa y Psicopedagógica"), "clm")
    assert not no_docente(rec("20000001", "X", "CEP", sigla="CEP"), "pvs")
    assert titulo("AINA MOLL I MARQUÉS") == "Aina Moll i Marqués" and titulo("JAUME I") == "Jaume I"
    assert titulo("CEIP JOSÉ SÁNCHEZ Y SÁNCHEZ") == "CEIP José Sánchez y Sánchez"
    assert titulo("C/ LAUREADO GENERAL SERRA ALGARRA, 14") == "C/ Laureado General Serra Algarra, 14"
    assert titulo("Ya Está Bien") == "Ya Está Bien" and titulo("") == ""
    assert titulo("AVDA. LA FERIA, Nº 3") == "Avda. la Feria, Nº 3"
    assert tipo_cat("Institut Escola Sant Martí", "Públic") == "Instituto Escuela"
    assert tipo_cat("Escola Els Porxos", "Públic") == "Escuela (Infantil y Primaria)"
    a = [dict(key="bal", code="07000001", dir="C/ A", email="", tel="971000001"),
         dict(key="bal", code="07000002", dir="C/ B", email="x@y.es", tel="")]
    b = [dict(key="bal", code="07000001", dir="C/ A", email="a@b.es", tel="971000001"),
         dict(key="bal", code="07000003", dir="C/ C", email="", tel="")]
    assert comparar(a, b) == {"antes": 2, "nuevos": 1, "bajas": 1, "emails_nuevos": 1, "telefonos_nuevos": 0, "total": 2}
    c = [dict(key="bal", code="07000001", dir="C/ A nueva", email="a@b.es", tel=""),  # cambio de dirección
         dict(key="bal", code="07000002", dir="C/ B", email="x@y.es", tel="971000002")]
    assert comparar(a, c) == {"antes": 2, "nuevos": 0, "bajas": 0, "emails_nuevos": 1, "telefonos_nuevos": 1, "total": 2}
    base = [dict(dict.fromkeys(CAMPOS_DATOS, ""), key="bal", ccaa="Illes Balears", code=f"070{i:05d}", dir=f"C/ {i}",
                 alta="2026-01-01" if i == 0 else "") for i in range(30)]
    nueva = [dict(r, email="e@x.es") for r in base[:29]] + [dict(base[0], code="07099999", dir="C/ Z", alta="")]
    filas, inf, cons, pares = fusionar(base, {"bal": {"fuente": "v"}}, nueva, {"bal": {"fuente": "n"}}, "2026-02-01")
    assert len(filas) == 30 and not cons and inf["bal"]["fuente"] == "n"
    assert sorted(r["alta"] for r in filas if r["alta"]) == ["2026-01-01", "2026-02-01"]
    assert resumir_cambios(pares, len(base)) == {"antes": 30, "nuevos": 1, "bajas": 1, "emails_nuevos": 29,
                                                 "telefonos_nuevos": 0, "total": 30}
    filas, inf, cons, _ = fusionar(base, {"bal": {"fuente": "v"}}, nueva[:10], {"bal": {"fuente": "n"}}, "2026-02-01")
    assert cons == ["bal"] and len(filas) == 30 and inf["bal"]["fuente"] == "v" and filas[0] is not base[0]
    filas, inf, cons, _ = fusionar(base, {}, nueva, {"bal": {"fuente": "n", "degradada": True}}, "2026-02-01")
    assert cons == ["bal"] and not any(r["email"] for r in filas)
    viejo = [dict(base[1], email="f@m.es", src_email="ficha", tel="971", src_tel="aut", web="w.es")]
    filas, _, _, _ = fusionar(viejo, {}, [dict(base[1])], {"bal": {}}, "2026-03-01")
    assert (filas[0]["email"], filas[0]["src_email"], filas[0]["tel"], filas[0]["web"]) == ("f@m.es", "ficha", "", "w.es")
    filas, _, _, _ = fusionar([], {}, [dict(base[1])], {"bal": {}}, "2026-03-01")
    assert filas[0]["alta"] == ""
    pruebas = {("Colegio de Educación Infantil y Primaria", "", "Pública"): "primaria",
               ("Col.legi Públic d'Educació Infantil i Primària", "", "Pública"): "primaria",
               ("Colexio de Educación Infantil e Primaria", "", "Pública"): "primaria",
               ("Escuela (Infantil y Primaria)", "Escola Mar Nova", "Pública"): "primaria",
               ("Instituto de Educación Secundaria/Bigarren Hezkuntzako Institutua", "", "Pública"): "secundaria",
               ("Instituto", "Institut Comte de Rius", "Pública"): "secundaria",
               ("Instituto Escuela", "", "Pública"): "colegio",
               ("Centro Privado de Educación Infantil Primaria y Secundaria", "", "Privada concertada"): "colegio",
               ("Colegio de Educación Infantil Primaria y Primer Ciclo de ESO", "", "Pública"): "colegio",
               ("Escuela Infantil (Consorcio Haurreskolak)", "", "Pública"): "infantil",
               ("Centro Privado de Educación Infantil de Primer Ciclo", "", "Privada"): "infantil",
               ("Centro Público Integrado de Formación Profesional", "", "Pública"): "fp",
               ("Centre de Formació d'Adults", "", "Pública"): "adultos",
               ("Escola Oficial d'Idiomes", "", "Pública"): "idiomas",
               ("Conservatorio Profesional de Música", "", "Pública"): "musica",
               ("Escuela de Arte y Superior de Diseño", "", "Pública"): "arte",
               ("Centro Privado Autorizado de Enseñanzas Deportivas", "", "Privada"): "deporte",
               ("Col·legi d'Educació Especial", "", "Pública"): "especial",
               ("Centro Docente Privado Extranjero en España", "", "Privada"): "extranjero",
               ("Centro privado", "Mare del Diví Pastor", "Privada concertada"): "colegio",
               ("Centro Docente Privado", "Formación Galeno", "Privada"): "fp",
               ("Centro público", "Aula Municipal de Música d'Artesa de Segre", "Pública"): "musica",
               ("Centro público", "Sant Martí de Provençals", "Pública"): "otros",
               ("CEIP", "", "Pública"): "primaria", ("", "", ""): "otros"}
    for (t, n, nat), esperada in pruebas.items():
        assert categoria(t, n, nat) == esperada, (t, n, categoria(t, n, nat), esperada)
    assert tipo_corto({"tipo": "Colegio Rural Agrupado", "nombre": "", "nat": "Pública"}) == "CRA"
    assert tipo_corto({"tipo": "Instituto de Educación Secundaria", "nombre": "", "nat": "Privada"}) == "Privado"
    assert tipo_corto({"tipo": "Centro privado", "nombre": "Y", "nat": "Privada concertada"}) == "Concertado"
    assert tipo_corto({"tipo": "Centro privado", "nombre": "Y", "nat": "Privada concertada"}, False) == "Colegio"
    assert tipo_corto({"tipo": "Instituto Escuela", "nombre": "", "nat": "Pública"}) == "Inst. escola"
    assert telefono_legible("971242034") == "971 242 034" and telefono_legible("+34 971") == "+34 971"
    fil = [dict(key="bal", prov="Illes Balears", mun="Palma", tipo="Instituto de Educación Secundaria", nombre="X",
                nat="Pública", email="", tel="", code="1", dir="", loc="", ccaa=""),
           dict(key="bal", prov="Illes Balears", mun="Inca", tipo="Escuela Infantil", nombre="Y", nat="Pública",
                email="", tel="", code="2", dir="", loc="", ccaa="")]
    assert filtrar(fil, municipios={("Illes Balears", "Inca")}) == [1]
    assert filtrar(fil, categorias={"secundaria"}) == [0] and filtrar(fil, categorias={"fp"}) == []
    antes = [dict(base[i], email=e, email_alta=f) for i, e, f in ((1, "", ""), (2, "A@b.es", "2026-01-01"),
                                                                  (3, "v@b.es", "2026-01-01"), (4, "q@b.es", ""))]
    despues = [dict(base[1], email="n@b.es"), dict(base[2], email="a@b.es"), dict(base[3], email="otro@b.es"),
               dict(base[4], email=""), dict(base[5], code="07088888", dir="C/ N", email="c@n.es")]
    pares = emparejar(antes, despues)
    marcar_emails_nuevos(pares, "2026-04-01")
    assert [r["email_alta"] for r in despues] == ["2026-04-01", "2026-01-01", "2026-04-01", "", "2026-04-01"]
    assert resumir_cambios(pares, len(antes))["emails_nuevos"] == 3
    print("autotest OK")


def carpeta_cache(preferida):
    """Carpeta de caché escribible: la preferida o, si no se puede, una en el perfil del usuario."""
    for d in (preferida, os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "CentrosEducativos", "cache")):
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


def obtener_centros(claves, sin_fichas=False, hilos=4, pausa=0.2, cache=".cache_centros", refrescar=False,
                    dias_fichas=30, incluir_no_docentes=False, desde_carpeta=None, cancel=None,
                    on_progress=None, on_fase=None):
    """Descarga y cruza los datos de las comunidades indicadas. Devuelve (filas, info, resumen)."""
    fase = on_fase or (lambda texto: None)
    hilos = max(1, min(8, int(hilos)))
    http = Http()
    offline = bool(desde_carpeta)
    st = Store(desde_carpeta or cache, http, refresh=refrescar, offline=offline)
    t0 = time.time()
    fase(f"Descargando los listados de {len(claves)} comunidad{'es' if len(claves) != 1 else ''}…")
    log(f"Comunidades: {', '.join(CCAA[k]['nombre'] for k in claves)}")
    rows, info, errores = [], {}, []
    with ThreadPoolExecutor(max_workers=max(1, min(6, len(claves)))) as ex:
        futs = {ex.submit(construir, k, st, incluir_no_docentes, info): k for k in claves}
        for f in as_completed(futs):
            k = futs[f]
            try:
                rs = f.result()
                rows.extend(rs)
                log(f"{CCAA[k]['nombre']}: {len(rs)} centros · teléfono {sum(1 for r in rs if r['tel'])} · "
                    f"email {sum(1 for r in rs if r['email'])}")
            except Exception as e:
                errores.append(CCAA[k]["nombre"])
                log(f"{CCAA[k]['nombre']}: ERROR {type(e).__name__}: {e}")
    cancelado = bool(cancel is not None and cancel.is_set())
    if rows and not sin_fichas and not cancelado:
        fase("Buscando los emails y teléfonos que faltan (fichas del Ministerio)…")
        fichas = Fichas(os.path.join(cache, "fichas_rcd.json"), http, dias_fichas, offline=offline)
        todo = list(dict.fromkeys((r["code"], CCAA[r["key"]]["rcd"]) for r in rows
                                  if (not r["email"] or not r["tel"]) and len(r["code"]) == 8))
        fichas.run(todo, hilos, pausa, cancel=cancel, on_progress=on_progress)
        aplicar_fichas(rows, fichas)
        cancelado = bool(cancel is not None and cancel.is_set())

    from collections import Counter, defaultdict
    muni = defaultdict(Counter)
    for r in rows:
        if r["prov"]:
            muni[(r["key"], norm(r["mun"]))][r["prov"]] += 1
    for r in rows:
        if not r["prov"] and muni.get((r["key"], norm(r["mun"]))):
            r["prov"] = muni[(r["key"], norm(r["mun"]))].most_common(1)[0][0]
        r["fuente"] = fuente_contacto(r)
        r["revisar"] = revisar(r["code"], r["email"])
    rows.sort(key=lambda r: (norm(r["ccaa"]), norm(r["prov"]), norm(r["mun"]), norm(r["nombre"])))
    n = len(rows)
    degradadas = [CCAA[k]["nombre"] for k, d in info.items() if d.get("degradada")]
    resumen = {"n": n, "tel": sum(1 for r in rows if r["tel"]), "email": sum(1 for r in rows if r["email"]),
               "errores": errores, "degradadas": degradadas, "cancelado": cancelado, "segundos": time.time() - t0}
    return rows, info, resumen


def guardar(path, rows, info=None):
    """Escribe .xlsx o .csv. Si el fichero está abierto (Excel), guarda con otro nombre. Devuelve la ruta final."""
    base, ext = os.path.splitext(path)
    ext = ext.lower() if ext.lower() in (".xlsx", ".csv") else ".xlsx"
    carpeta = os.path.dirname(os.path.abspath(path))
    os.makedirs(carpeta, exist_ok=True)
    for i in range(1, 30):
        destino = f"{base}{ext}" if i == 1 else f"{base} ({i}){ext}"
        try:
            if ext == ".csv":
                escribir_csv(destino, rows)
            else:
                escribir_xlsx(destino, rows, info or {})
            return destino
        except PermissionError:
            continue
    raise PermissionError(f"No se puede escribir en {carpeta}")


# ---------------------------------------------------------------------------
# Lista local (para el programa con ventana): se guarda comprimida y se compara al actualizar
# ---------------------------------------------------------------------------
CAMPOS_DATOS = ["ccaa", "key", "prov", "code", "nombre", "tipo", "nat", "dir", "cp", "loc", "mun", "tel", "email", "web",
                "src_tel", "src_email", "fuente", "revisar", "alta", "email_alta"]
VERSION_DATOS = 1


def guardar_datos(path, rows, info, meta=None):
    """Guarda la lista (gzip + JSON). Escribe en un temporal y lo renombra: nunca deja el fichero a medias."""
    import gzip
    doc = {"version": VERSION_DATOS, "meta": meta or {}, "info": info, "cols": CAMPOS_DATOS,
           "rows": [[r.get(c, "") or "" for c in CAMPOS_DATOS] for r in rows]}
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp"
    with gzip.open(tmp, "wt", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, path)


def cargar_datos(path):
    """Lee una lista guardada con guardar_datos. Devuelve (filas, info, meta)."""
    import gzip
    with gzip.open(path, "rt", encoding="utf-8") as f:
        doc = json.load(f)
    if int(doc.get("version", 0)) > VERSION_DATOS:
        raise ValueError(f"lista de una versión más nueva del programa ({doc.get('version')})")
    cols = doc["cols"]
    rows = []
    for v in doc["rows"]:
        r = dict.fromkeys(CAMPOS_DATOS, "")
        r.update(zip(cols, v))
        rows.append(r)
    if not rows or not all(r["key"] in CCAA and r["code"] for r in rows[:50]):
        raise ValueError("lista vacía o con un formato desconocido")
    meta = doc.get("meta") or {}
    if "email_alta" not in cols and meta.get("nuevos_desde"):
        # lista guardada antes de que existiera email_alta: lo que trajo la última actualización son los emails
        # de las fichas del Ministerio (la lista incluida con el programa no tiene ninguno)
        for r in rows:
            if r["email"] and r["src_email"] == "ficha":
                r["email_alta"] = meta["nuevos_desde"]
    return rows, doc.get("info") or {}, meta


def clave_centro(r):
    return (r["key"], r["code"], norm(r["dir"]))


def emparejar(antes, despues):
    """Lista de (fila_antes o None, fila_despues): el mismo centro por clave o, si cambió la dirección, por código."""
    from collections import defaultdict
    a = {}
    for r in antes:
        a.setdefault(clave_centro(r), r)
    usados, pares, pendientes = set(), [], []
    for r in despues:
        k = clave_centro(r)
        if k in a and k not in usados:
            usados.add(k)
            pares.append((a[k], r))
        else:
            pendientes.append(r)
    libres = defaultdict(list)
    for k, r in a.items():
        if k not in usados:
            libres[k[:2]].append(r)
    for r in pendientes:
        lst = libres.get((r["key"], r["code"]))
        pares.append((lst.pop() if lst else None, r))
    return pares


def email_nuevo(x, r):
    """El email actual del centro r no estaba antes (centro nuevo con email, email añadido o email cambiado)."""
    return bool(r.get("email")) and (x is None or r["email"].strip().lower() != (x.get("email") or "").strip().lower())


def resumir_cambios(pares, n_antes):
    emparejados = sum(1 for x, _ in pares if x is not None)
    return {"antes": n_antes, "nuevos": len(pares) - emparejados, "bajas": max(0, n_antes - emparejados),
            "emails_nuevos": sum(1 for x, r in pares if email_nuevo(x, r)),
            "telefonos_nuevos": sum(1 for x, r in pares if x is not None and r["tel"] and not x["tel"]),
            "total": len(pares)}


def comparar(antes, despues):
    """Cambios entre dos listas: centros nuevos, dados de baja, emails y teléfonos nuevos."""
    return resumir_cambios(emparejar(antes, despues), len(antes))


def fusionar(antes, info_antes, despues, info_despues, hoy, umbral=0.8):
    """Une la lista anterior con la recién descargada (programa con ventana).

    Conserva los datos anteriores de una comunidad si ahora ha fallado, si solo se ha podido usar el Registro Estatal
    (su directorio autonómico no respondía) o si trae muchos menos centros que antes (descarga incompleta).
    Marca con 'hoy' (fecha y hora, campo alta) los centros que no estaban. Devuelve (filas, info, conservadas, pares)."""
    from collections import defaultdict
    viejas, nuevas = defaultdict(list), defaultdict(list)
    for r in antes:
        viejas[r["key"]].append(r)
    for r in despues:
        nuevas[r["key"]].append(r)
    filas, info, conservadas = [], {}, []
    for k in CCAA:
        d = info_despues.get(k)
        n_a, n_d = len(viejas[k]), len(nuevas[k])
        mala = d is None or d.get("degradada") or (n_a >= 20 and n_d < umbral * n_a)
        if mala and viejas[k]:
            conservadas.append(k)
            filas.extend({c: r.get(c, "") for c in CAMPOS_DATOS} for r in viejas[k])
            if k in info_antes:
                info[k] = info_antes[k]
            if d is not None:
                log(f"{CCAA[k]['nombre']}: se mantienen los datos anteriores ({n_a} centros; ahora {n_d}"
                    f"{', solo Registro Estatal' if d.get('degradada') else ''}).")
        else:
            filas.extend(nuevas[k])
            if d is not None:
                info[k] = d
    filas.sort(key=lambda r: (norm(r["ccaa"]), norm(r["prov"]), norm(r["mun"]), norm(r["nombre"])))
    pares = emparejar(antes, filas)
    for x, r in pares:
        if x is None:
            r["alta"] = hoy if antes else ""   # en la primera descarga no hay con qué comparar
            continue
        r["alta"] = x.get("alta") or ""
        # lo que ya se sabía por la ficha del Ministerio no se pierde aunque falte la caché de fichas
        if not r["email"] and x.get("email") and x.get("src_email") == "ficha":
            r["email"], r["src_email"] = x["email"], "ficha"
        if not r["tel"] and x.get("tel") and x.get("src_tel") == "ficha":
            r["tel"], r["src_tel"] = x["tel"], "ficha"
        if not r["web"] and x.get("web"):
            r["web"] = x["web"]
    return filas, info, conservadas, pares


def marcar_emails_nuevos(pares, marca, hay_antes=True):
    """Fecha (email_alta) en que apareció el email actual de cada centro. Llamar tras completar con las fichas."""
    for x, r in pares:
        if not r.get("email"):
            r["email_alta"] = ""
        elif email_nuevo(x, r):
            r["email_alta"] = marca if hay_antes else ""
        else:
            r["email_alta"] = x.get("email_alta") or ""


def completar_con_fichas(rows, fichas):
    """Rellena con las fichas del Ministerio ya descargadas los emails/teléfonos que faltan y recalcula fuente y aviso."""
    aplicar_fichas(rows, fichas)
    for r in rows:
        r["fuente"] = fuente_contacto(r)
        r["revisar"] = revisar(r["code"], r["email"])


def por_consultar(rows, fichas):
    """(código, id comunidad) de los centros sin email o teléfono cuya ficha falta o está caducada."""
    return list(dict.fromkeys((r["code"], CCAA[r["key"]]["rcd"]) for r in rows
                              if (not r["email"] or not r["tel"]) and len(r["code"]) == 8 and not fichas.fresh(r["code"])))


def fichas_pendientes(rows, cache, dias_fichas=90):
    """Cuántos centros sin email o teléfono tienen la ficha del Ministerio sin consultar (o caducada)."""
    return len(por_consultar(rows, Fichas(os.path.join(cache, "fichas_rcd.json"), None, dias_fichas, offline=True)))


# ---------------------------------------------------------------------------
# Tipo de centro: categorías comunes a partir de las denominaciones oficiales (en las lenguas cooficiales)
# ---------------------------------------------------------------------------
CATEGORIAS = [  # (clave, nombre en el filtro), en el orden en que se muestran
    ("infantil", "Escuela infantil / guardería"),
    ("primaria", "Infantil y Primaria (CEIP y colegios)"),
    ("colegio", "Colegio con varias etapas"),
    ("secundaria", "Secundaria y Bachillerato (IES y colegios)"),
    ("fp", "Formación Profesional"),
    ("adultos", "Personas adultas"),
    ("idiomas", "Idiomas (EOI)"),
    ("musica", "Música y danza"),
    ("arte", "Arte, diseño y artes escénicas"),
    ("especial", "Educación especial"),
    ("deporte", "Enseñanzas deportivas"),
    ("extranjero", "Centros extranjeros"),
    ("otros", "Otros"),
]
NOMBRE_CATEGORIA = dict(CATEGORIAS)
_SIGLAS_TIPO = {
    "ceip": "primaria", "cei": "infantil", "cpi": "colegio", "ceo": "colegio", "cpfpb": "fp", "imfpb": "fp",
    "cpfp": "fp", "ifp": "fp", "esda": "arte", "esmu": "musica", "emuspr": "musica", "emr": "musica",
    "caed": "deporte", "caem": "musica", "capm": "musica", "cpd": "musica", "casm": "musica", "caapd": "arte",
    "casdi": "arte", "cpepa": "adultos", "cepa": "adultos", "cpeps": "colegio", "cpep": "primaria", "eh": "otros",
}
_RE_PALABRA = {k: re.compile(r"\b(" + v + r")\b") for k, v in {
    "eso": "eso", "eoi": "eoi", "fp": "fp", "canto": "canto", "art": "art|arte|artes|arts|art\\.",
}.items()}


def _hay(t, *claves):
    return any(k in t for k in claves)


def _categoria_por_nombre(n, privado):
    if _hay(n, "formacion", "formacio", "business school", "escuela profesional", "escuelas profesionales",
            "academia", "cidep", "ilerna") or _RE_PALABRA["fp"].search(n):
        return "fp"
    if _hay(n, "musica", "danza", "dansa", "conservatori"):
        return "musica"
    if _hay(n, "llar d'infants", "bressol", "infantil", "guarderia", "jardin de infancia", "kids", "nursery"):
        return "infantil"
    if _hay(n, "idiomas", "idiomes", "language", "english"):
        return "idiomas"
    if _hay(n, "diseno", "disseny") or _RE_PALABRA["art"].search(n):
        return "arte"
    if _hay(n, "hospital", "justicia", "c.educ."):
        return "otros"
    return "colegio" if privado else "otros"


def categoria(tipo, nombre="", nat=""):
    """Categoría común del centro (clave de CATEGORIAS) según su denominación oficial y, si es genérica, su nombre."""
    t = " ".join(norm(tipo).split())
    privado = norm(nat).startswith("priv") or _hay(t, "privad", "privat")
    if not t:
        return _categoria_por_nombre(norm(nombre), privado)
    if t in _SIGLAS_TIPO:
        return _SIGLAS_TIPO[t]
    if "eoep" in t:
        return "otros"
    if _hay(t, "extranjer", "estranger"):
        return "extranjero"
    prim = _hay(t, "primaria", "lehen hezkuntza", "pri+")
    secu = _hay(t, "secundaria", "bachillerato", "bigarren") or bool(_RE_PALABRA["eso"].search(t))
    if (prim and secu) or _hay(t, "instituto escuela", "institut escola", "educacion obligatoria", "educacion basica",
                                "diversos ensenyaments", "ensenanzas integradas", "escuela europea", "ikastetxe bateratua"):
        return "colegio"
    if "integrad" in t and not _hay(t, "profesional", "musica", "danza", "dansa"):
        return "colegio"
    if _hay(t, "educacion especial", "educacio especial", "hezkuntza berezi", "terapeutico"):
        return "especial"
    if _hay(t, "idiomas", "idiomes") or _RE_PALABRA["eoi"].search(t):
        return "idiomas"
    if _hay(t, "musica", "danza", "dansa", "conservatori") or _RE_PALABRA["canto"].search(t):
        return "musica"
    if _hay(t, "deportiv", "esport", "futbol", "hipica", "salvamento"):
        return "deporte"
    if _hay(t, "adult", "helduen", "educacion permanente", "penitenciari", "distancia"):
        return "adultos"
    if _hay(t, "formacion profesional", "formacio professional", "capacitacion agraria", "capacitacio agraria",
            "empleo", "insercion laboral", "agrari", "acuicola", "pesquer", "viticultura", "energias renovables",
            "formacion ambiental") or _RE_PALABRA["fp"].search(t):
        return "fp"
    if _hay(t, "diseno", "disseny", "deseno", "conservacion y restauracion", "conservacio i restauracio", "ceramica",
            "vidrio", "dramatic", "escenicas", "plastiques", "plasticas") or _RE_PALABRA["art"].search(t):
        return "arte"
    if prim or _hay(t, "infantil y primaria", "infantil e primaria", "infantil i primaria", "rural agrupado",
                    "publico rural", "colegio publico"):
        return "primaria"
    if secu or _hay(t, "instituto", "institut", "batxillerat"):
        return "secundaria"
    if _hay(t, "infantil", "bressol", "llar d'infants", "haur", "preescolar", "casa de ninos", "guarderia",
            "centro de convenio"):
        return "infantil"
    if _hay(t, "hogar", "fogar", "residencia", "hospital", "militar"):
        return "otros"
    return _categoria_por_nombre(norm(nombre), privado)


def categoria_centro(r):
    """Categoría de una fila (se guarda en la propia fila para no recalcularla)."""
    c = r.get("_cat")
    if c is None:
        c = r["_cat"] = categoria(r.get("tipo", ""), r.get("nombre", ""), r.get("nat", ""))
    return c


def tipo_corto(r, titularidad=True):
    """Etiqueta breve del tipo (CEIP, IES, E. infantil…). Con titularidad=True, los colegios privados se distinguen
    como «Concertado» o «Privado» (para la tabla); sin ella, queda el nivel (para acompañar a la titularidad)."""
    c = categoria_centro(r)
    t = norm(r.get("tipo", ""))
    nat = norm(r.get("nat", ""))
    publico = nat.startswith("public")
    privado = "concertad" in nat and "Concertado" or "Privado"
    if c == "infantil":
        return "E. infantil" if publico or not titularidad else "Infantil priv."
    if c in ("primaria", "colegio", "secundaria") and not publico:
        return privado if titularidad else "Colegio"
    if c == "primaria":
        return "CRA" if _hay(t, "rural agrupado", "publico rural") else "CEIP"
    if c == "colegio":
        if _hay(t, "instituto escuela", "institut escola"):
            return "Inst. escola"
        if _hay(t, "obligatoria", "basica"):
            return "CEO"
        return "CPI" if "integrad" in t else "CEIP + ESO"
    if c == "secundaria":
        return "IES"
    if c == "adultos":
        return "CEPA" if publico else "Adultos"
    if c == "idiomas":
        return "EOI" if publico else "Idiomas"
    return {"fp": "FP", "musica": "Música", "arte": "Arte", "especial": "Ed. especial", "deporte": "Deporte",
            "extranjero": "Extranjero"}.get(c, "Otro")


def telefono_legible(tel):
    """971242034 → 971 242 034 (solo para mostrar)."""
    d = re.sub(r"\D", "", tel or "")
    return f"{d[:3]} {d[3:6]} {d[6:]}" if len(d) == 9 else (tel or "")


def filtrar(rows, texto="", provincia="", titularidad="", solo_email=False, solo_telefono=False, comunidades=None,
            municipios=None, categorias=None):
    """Índices de las filas que cumplen los filtros. El texto busca todas sus palabras (sin tildes ni mayúsculas).
    comunidades: conjunto de claves ('bal', 'mad'...); municipios: conjunto de (provincia, municipio);
    categorias: conjunto de claves de CATEGORIAS. Vacío o None = sin ese filtro."""
    palabras = norm(texto).split()
    out = []
    for i, r in enumerate(rows):
        if comunidades and r["key"] not in comunidades:
            continue
        if provincia and r["prov"] != provincia:
            continue
        if municipios and (r["prov"], r["mun"]) not in municipios:
            continue
        if categorias and categoria_centro(r) not in categorias:
            continue
        if titularidad and r["nat"] != titularidad:
            continue
        if solo_email and not r["email"]:
            continue
        if solo_telefono and not r["tel"]:
            continue
        if palabras:
            idx = r.get("_idx")
            if idx is None:
                idx = r["_idx"] = norm(" ".join((r["nombre"], r["tipo"], r["mun"], r["loc"], r["prov"], r["ccaa"],
                                                 r["email"], r["code"], r["tel"], r["dir"], tipo_corto(r),
                                                 tipo_corto(r, False), NOMBRE_CATEGORIA[categoria_centro(r)])))
            if not all(p in idx for p in palabras):
                continue
        out.append(i)
    return out


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser(description="Centros educativos de España con teléfono y email (fuentes oficiales).",
                                 formatter_class=argparse.RawDescriptionHelpFormatter,
                                 epilog="Ejemplos:\n  python centros_educativos.py --ccaa baleares\n"
                                        "  python centros_educativos.py --ccaa valencia galicia -o centros.xlsx\n"
                                        "  python centros_educativos.py --ccaa todas -o espana.csv")
    ap.add_argument("--ccaa", nargs="+", metavar="CCAA", help="comunidades (nombre o código; varias separadas por espacio) o 'todas'")
    ap.add_argument("-o", "--salida", help="fichero de salida .xlsx o .csv (por defecto centros_<ccaa>_<fecha>.xlsx)")
    ap.add_argument("--sin-fichas", action="store_true", help="no consultar las fichas del Ministerio (más rápido, menos emails)")
    ap.add_argument("--hilos", type=int, default=4, help="conexiones simultáneas para las fichas (por defecto 4, máximo 8)")
    ap.add_argument("--pausa", type=float, default=0.2, help="pausa en segundos tras cada ficha y conexión (por defecto 0,2)")
    ap.add_argument("--cache", default=".cache_centros", help="carpeta de caché (por defecto .cache_centros)")
    ap.add_argument("--refrescar", action="store_true", help="volver a descargar los listados aunque estén en caché (<24 h)")
    ap.add_argument("--dias-fichas", type=int, default=30, help="días que vale una ficha guardada (por defecto 30)")
    ap.add_argument("--incluir-no-docentes", action="store_true",
                    help="incluir equipos de orientación, aulas hospitalarias, centros del profesorado, etc.")
    ap.add_argument("--desde-carpeta", metavar="DIR", help="leer los listados de una carpeta, sin red (and.csv, rcd_01.xls...)")
    ap.add_argument("--abrir", action="store_true", help="abrir el fichero al terminar (Windows)")
    ap.add_argument("--listar", action="store_true", help="mostrar las comunidades disponibles")
    ap.add_argument("--autotest", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    if args.autotest:
        autotest()
        return 0
    if args.listar or not args.ccaa:
        print("Comunidades (usa el código o el nombre):")
        for k, d in CCAA.items():
            extra = "" if k in LOADERS else "  [solo Registro Estatal + fichas]"
            print(f"  {k}  {d['nombre']}{extra}")
        print("  todas")
        return 0 if args.listar else 2

    claves = resolver_ccaa(args.ccaa)
    salida = args.salida or f"centros_{'todas' if len(claves) == len(CCAA) else '-'.join(claves)}_{datetime.now():%Y%m%d}.xlsx"
    if not salida.lower().endswith((".xlsx", ".csv")):
        salida += ".xlsx"
    cache = carpeta_cache(args.cache)
    try:
        rows, info, res = obtener_centros(claves, sin_fichas=args.sin_fichas, hilos=args.hilos, pausa=args.pausa,
                                          cache=cache, refrescar=args.refrescar, dias_fichas=args.dias_fichas,
                                          incluir_no_docentes=args.incluir_no_docentes, desde_carpeta=args.desde_carpeta)
    except KeyboardInterrupt:
        log("Interrumpido: las fichas descargadas se han guardado; vuelve a ejecutar para continuar.")
        return 130
    if not rows:
        log("No se ha obtenido ningún centro.")
        return 1
    salida = guardar(salida, rows, info)
    n = res["n"]
    log(f"Listo en {res['segundos'] / 60:.1f} min: {n} centros · teléfono {res['tel']} ({res['tel'] / n:.0%}) · "
        f"email {res['email']} ({res['email'] / n:.0%}) -> {os.path.abspath(salida)}")
    if res["errores"]:
        log("Comunidades con error: " + ", ".join(res["errores"]))
    if args.abrir and hasattr(os, "startfile"):
        try:
            os.startfile(os.path.abspath(salida))
        except OSError:
            pass
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log("Cancelado.")
        sys.exit(130)
