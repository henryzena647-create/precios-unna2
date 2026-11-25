#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
actualizar_precios.py
=====================
Servicio de actualización de precios de lista (Petroperú + Repsol) para el
reporte HTML de CDA / Survial - UNNA Transporte.

Qué hace en cada corrida (pensado para 7:00, 13:00 y 19:00):
  1. Petroperú combustible  -> última lista COMB (tabla uso interno S-50, por planta)
  2. Petroperú asfalto       -> última lista ASFA (banda ex-planta)
  3. Repsol asfalto          -> última "Lista de precios de asfaltos" (PDF público)
  4. Repsol combustible      -> reporte dinámico PrecioLima (se renderiza con navegador)
  5. Ensambla ./data.json en el MISMO esquema que usa el reporte HTML (v5)
  6. Mantiene ./historico.json para que la línea de tiempo persista entre corridas

IMPORTANTE
----------
- Este script corre FUERA del navegador (por eso puede leer el reporte dinámico
  de Repsol y descargar los PDF de Petroperú, cosa que el HTML no puede por CORS).
- Robustez: si una fuente falla, se conserva el último valor conocido (del data.json
  previo). Una corrida con error NO borra los datos buenos.
- La parte más propensa a requerir un ajuste fino es el selector del reporte SSRS
  de Repsol (paso 4): está marcada con  # >>> AJUSTAR .  Se ajusta una sola vez,
  viendo el HTML real del reporte, y queda estable.

Requisitos:
    pip install playwright requests pdfplumber
    playwright install chromium
"""

import json
import re
import sys
import io
import os
import ssl
import smtplib
import datetime as dt
from email.message import EmailMessage
from pathlib import Path

import requests

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

BASE = Path(__file__).resolve().parent
OUT = BASE / "data.json"
HIST = BASE / "historico.json"
LOG = BASE / "actualizar_precios.log"

TZ_HINT = "America/Lima"  # informativo; la programación horaria la da el SO

# ---------------- URLs de fuentes ----------------
PETRO_LISTAS_PAGE   = "https://www.petroperu.com.pe/productos/lista-de-precios-en-nuestras-plantas/"
PETRO_ASFALTOS_PAGE = "https://www.petroperu.com.pe/productos/lista-de-precios-en-nuestras-plantas/asfaltos"
REPSOL_ASF_PAGE     = "https://www.repsol.pe/es/soluciones-para-empresas/asfaltos/index.cshtml"
REPSOL_FUEL_REPORT  = "https://relapasaa.cloudapp.repsol.com/Reportes/PrecioLima"

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; UNNA-PriceBot/1.0)"}

# ---------------- Alertas por correo (credenciales por variables de entorno) ----------------
# Configura estas variables en el sistema o en GitHub Secrets. Si faltan, no se envía correo
# (pero igual se registra el cambio en alertas.log). NUNCA escribas la contraseña aquí.
SMTP_HOST = os.environ.get("SMTP_HOST", "")
SMTP_PORT = int(os.environ.get("SMTP_PORT") or 587)
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASS = os.environ.get("SMTP_PASS", "")
ALERTA_FROM = os.environ.get("ALERTA_EMAIL_FROM", SMTP_USER)
ALERTA_LOG = BASE / "alertas.log"
CONFIG_ALERTAS = BASE / "config_alertas.json"   # lo escribe el botón "Guardar" del reporte


def _destinatarios():
    """Correos de alerta: variable de entorno + config_alertas.json (unidos, sin repetir)."""
    dest = [x.strip() for x in os.environ.get("ALERTA_EMAIL_TO", "").split(",") if x.strip()]
    try:
        cfg = json.loads(CONFIG_ALERTAS.read_text(encoding="utf-8"))
        for e in cfg.get("destinatarios", []):
            if e and e not in dest:
                dest.append(e)
    except Exception:
        pass
    return dest


def _webhook_url():
    """URL del flujo de Power Automate: variable de entorno o config_alertas.json."""
    url = os.environ.get("POWERAUTOMATE_URL", "").strip()
    if url:
        return url
    try:
        cfg = json.loads(CONFIG_ALERTAS.read_text(encoding="utf-8"))
        return (cfg.get("webhook_url") or "").strip()
    except Exception:
        return ""

# Plantas Petroperú relevantes y a qué sede abastecen (se conserva del reporte)
PLANTAS_SEDE = {
    "TALARA": "Ruta norte / Canchaque",
    "PIURA": "Canchaque",
    "CONCHAN": "Ancón (Lima)",
    "CONCHÁN": "Ancón (Lima)",
    "CALLAO": "Ancón (Lima)",
    "PISCO": "Ruta a Puquio",
    "MOLLENDO": "Sur / Abancay–Puquio",
    "CUSCO": "Abancay",
}


def log(msg):
    line = f"{dt.datetime.now().isoformat(timespec='seconds')}  {msg}"
    print(line)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


# =====================================================================
#  Utilidades de navegador (Playwright) - se usan solo donde hace falta
# =====================================================================
def render_html(url, wait_selector=None, timeout=45000):
    """Devuelve el HTML ya renderizado (JS ejecutado) de una URL."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent=HEADERS["User-Agent"])
        page.goto(url, timeout=timeout, wait_until="networkidle")
        if wait_selector:
            try:
                page.wait_for_selector(wait_selector, timeout=timeout)
            except Exception:
                pass
        html = page.content()
        browser.close()
        return html


def render_text(url, wait_selector=None, timeout=45000):
    """Devuelve el texto visible (innerText) de una URL renderizada."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent=HEADERS["User-Agent"])
        page.goto(url, timeout=timeout, wait_until="networkidle")
        if wait_selector:
            try:
                page.wait_for_selector(wait_selector, timeout=timeout)
            except Exception:
                pass
        txt = page.inner_text("body")
        browser.close()
        return txt


def find_pdf_links(html):
    """Extrae (url, fecha) de enlaces a PDF que aparezcan en el HTML."""
    links = re.findall(r'href="([^"]+\.pdf[^"]*)"', html, flags=re.I)
    out = []
    for l in links:
        if l.startswith("/"):
            # resolver relativo a petroperu (ajustar si otra fuente)
            l = "https://www.petroperu.com.pe" + l
        out.append(l)
    return list(dict.fromkeys(out))  # dedup preservando orden


def _petro_price_pdfs(html):
    """Solo los PDF que son LISTAS DE PRECIOS (carpeta tbl_listas_de_precio),
    ordenados del más nuevo al más viejo por el número de archivo (1737 > 1719...).
    Evita agarrar el informe de transparencia u otros PDF que hay en la página."""
    price = [l for l in find_pdf_links(html) if "tbl_listas_de_precio" in l]
    def _id(u):
        m = re.search(r"/(\d+)-[^/]*\.pdf", u, flags=re.I)
        return int(m.group(1)) if m else -1
    return sorted(price, key=_id, reverse=True)


def download_pdf(url):
    r = requests.get(url, headers=HEADERS, timeout=60)
    r.raise_for_status()
    return r.content


def pdf_text(data):
    if not pdfplumber:
        raise RuntimeError("Falta pdfplumber (pip install pdfplumber)")
    txt = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for pg in pdf.pages:
            txt.append(pg.extract_text() or "")
    return "\n".join(txt)


def nums(line):
    """Todos los números tipo 1234.56 en una línea."""
    return [float(x) for x in re.findall(r"\d+\.\d+", line)]


# =====================================================================
#  1) PETROPERÚ COMBUSTIBLE  (tabla uso interno S-50, por planta)
# =====================================================================
def scrape_petro_comb():
    html = render_html(PETRO_LISTAS_PAGE)          # vista "Automotriz" por defecto
    pdfs = _petro_price_pdfs(html)
    if not pdfs:
        raise RuntimeError("No se hallaron PDF de LISTA DE PRECIOS de combustible en Petroperú")
    latest = pdfs[0]                               # el de mayor número = más reciente
    log(f"Petroperú combustible: {latest}")
    text = pdf_text(download_pdf(latest))
    lines = text.splitlines()

    # GLP (S/kg): primera columna de la tabla principal (solo si el 1er número es < 5)
    glp = {}
    for ln in lines:
        u = ln.upper().strip()
        for plant in PLANTAS_SEDE:
            if u.startswith(plant):
                v = nums(ln)
                if v and v[0] < 5:
                    glp.setdefault(plant, v[0])
                break

    # Tabla USO INTERNO vehicular (ADDENDUM N°1):
    #   DIESEL B5 UV S-50 | DIESEL B5 S-50 | GASOHOL PREMIUM | GASOHOL REGULAR
    start = None
    for i, ln in enumerate(lines):
        u = ln.upper()
        if "DIESEL B5 UV" in u and "GASOHOL" in u:   # encabezado de esa tabla
            start = i
            break
    plantas = []
    if start is not None:
        for ln in lines[start + 1:]:
            u = ln.upper().strip()
            if "IMPUESTO" in u or "LEY DE" in u or "ADDENDUM" in u:
                break
            for plant in PLANTAS_SEDE:
                if u.startswith(plant):
                    v = nums(ln)
                    if len(v) >= 2:
                        plantas.append({
                            "planta": plant.capitalize(),
                            "abastece": PLANTAS_SEDE[plant],
                            "diesel_uv":   v[0],
                            "diesel_s50":  v[1],
                            "gasohol_prem": v[2] if len(v) > 2 else None,
                            "gasohol_reg":  v[3] if len(v) > 3 else None,
                            "glp_kg": glp.get(plant),
                        })
                    break
    # dedup por planta
    seen, uniq = set(), []
    for p in plantas:
        if p["planta"] not in seen:
            seen.add(p["planta"]); uniq.append(p)
    lista = _find_code(text, r"COMB-\d+-\d+")
    mv = re.search(r"VIGENCIA A PARTIR DEL\s*(\d{1,2})[./](\d{1,2})[./](\d{4})", text.upper())
    vig = f"{mv.group(3)}-{int(mv.group(2)):02d}-{int(mv.group(1)):02d}" if mv else _find_date(text)
    return {"lista": lista or "COMB", "vigente_desde": vig, "plantas": uniq}


# =====================================================================
#  2) PETROPERÚ ASFALTO  (banda ex-planta)
# =====================================================================
def scrape_petro_asf():
    html = render_html(PETRO_ASFALTOS_PAGE)
    pdfs = _petro_price_pdfs(html)
    if not pdfs:
        raise RuntimeError("No se hallaron PDF de LISTA DE PRECIOS de asfalto en Petroperú")
    latest = pdfs[0]
    log(f"Petroperú asfalto: {latest}")
    text = pdf_text(download_pdf(latest))

    # precios reales = números seguidos de "Soles/Gal" (evita fechas como 24.09.2026)
    precios = [float(x) for x in re.findall(r"(\d+\.\d+)\s*Soles/?\s*Gal", text, flags=re.I)]
    banda = {"min": min(precios), "max": max(precios)} if precios else {"min": None, "max": None}

    # mapear grado -> precio (Carros Cisterna) por bloques "Despachos ... / Carros Cisterna ..."
    items_map = {}
    for blk in re.split(r"Despachos", text)[1:]:
        hdr = blk.splitlines()[0] if blk.splitlines() else ""
        grades = re.findall(r"PEN\s*\d+/\d+(?:\s*AL\s*\d+/\d+)?|RC-\d+|MC-\d+", hdr, flags=re.I)
        m = re.search(r"Carros Cisterna(.*?)(?:Cilindro|$)", blk, flags=re.S)
        if m:
            prices = [float(x) for x in re.findall(r"(\d+\.\d+)\s*Soles", m.group(1))]
            for g, p in zip(grades, prices):
                items_map[re.sub(r"\s+", " ", g.upper().strip())] = p

    def _g(*keys):
        for k in keys:
            kk = re.sub(r"\s+", " ", k.upper().strip())
            if kk in items_map:
                return items_map[kk]
        return None

    items = [
        {"familia": "Cemento asfáltico", "grado": "PEN 40/50 a 120/150", "precio": _g("PEN 40/50 AL 120/150")},
        {"familia": "Cemento asfáltico", "grado": "PEN 20/30", "precio": _g("PEN 20/30")},
        {"familia": "Cemento asfáltico", "grado": "PEN 10/20", "precio": _g("PEN 10/20")},
        {"familia": "Asfalto líquido", "grado": "RC-70 / RC-250", "precio": _g("RC-70", "RC-250")},
        {"familia": "Asfalto líquido", "grado": "MC-30 / MC-70", "precio": _g("MC-30", "MC-70")},
    ]
    lista = _find_code(text, r"ASFA-\d+-\d+")
    mv = re.search(r"VIGENTE DESDE[:\s]*(\d{1,2})[/.](\d{1,2})[/.](\d{4})", text.upper())
    vig = f"{mv.group(3)}-{int(mv.group(2)):02d}-{int(mv.group(1)):02d}" if mv else _find_date(text)
    return {"lista": lista or "ASFA", "vigente_desde": vig,
            "plantas": "Talara / Conchán", "items": items, "banda": banda}


# =====================================================================
#  3) REPSOL ASFALTO  (PDF público en repsol.pe)
# =====================================================================
def scrape_repsol_asf():
    html = render_html(REPSOL_ASF_PAGE)
    # enlaces del tipo .../Precio de Lista Asfaltos- Recosac DDMMYY.pdf
    raw = re.findall(r'href="([^"]+\.pdf[^"]*)"', html, flags=re.I)
    cand = [l for l in raw if re.search(r"[Pp]recio", l) and re.search(r"[Aa]sfalto", l)]
    cand = [l for l in cand if "Ficha" not in l and "Seguridad" not in l]
    if not cand:
        raise RuntimeError("No se halló PDF de LISTA DE PRECIOS de asfalto Repsol (solo ficha u otros)")
    def _rdate(u):
        m = re.search(r"(\d{2})(\d{2})(\d{2})\.pdf", u)   # DDMMYY -> AAAAMMDD
        return int("20" + m.group(3) + m.group(2) + m.group(1)) if m else -1
    cand.sort(key=_rdate, reverse=True)
    latest = cand[0]
    if latest.startswith("/"):
        latest = "https://www.repsol.pe" + latest
    log(f"Repsol asfalto: {latest}")
    text = pdf_text(download_pdf(requests.utils.requote_uri(latest)))
    vals = [x for x in nums(text.replace("\n", " ")) if 5 < x < 40]
    precio = max(set(vals), key=vals.count) if vals else None   # precio único más frecuente
    vig = _find_date(text)
    items = [
        {"familia": "Cemento asfáltico", "grado": "CA 60/70", "precio": precio},
        {"familia": "Cemento asfáltico", "grado": "CA 85/100", "precio": precio},
        {"familia": "Cemento asfáltico", "grado": "CA 120/150", "precio": precio},
    ]
    return {"lista": "Recosac (público)", "vigente_desde": vig,
            "nota": "Repsol/RECOSAC · Terminal La Pampilla (Callao). Precios sin IGV.",
            "items": items}


# =====================================================================
#  4) REPSOL COMBUSTIBLE  (reporte dinámico PrecioLima - SSRS)
# =====================================================================
def _repsol_report_extract(url, timeout=70000):
    """
    Devuelve (rows, fulltext) del reporte SSRS.
    - rows: lista de filas, cada una es lista de celdas (texto) leídas de las <tr>/<td>
            recorriendo TODOS los frames (el reporte vive dentro de un iframe).
      Leer por celdas evita el problema de que inner_text ponga cada celda en su
      propia línea y se rompa el emparejamiento producto->precio.
    - fulltext: texto concatenado (para metadatos como la vigencia).
    """
    from playwright.sync_api import sync_playwright
    rows, textparts = [], []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent=HEADERS["User-Agent"])
        page.goto(url, timeout=timeout, wait_until="networkidle")
        page.wait_for_timeout(7000)  # SSRS pinta por postbacks asíncronos
        for _ in range(10):          # esperar a que algún frame muestre un producto
            if any(re.search(r"DIESEL|GASOHOL|GLP", (fr.inner_text("body") or "").upper())
                   for fr in page.frames):
                break
            page.wait_for_timeout(2000)
        for fr in page.frames:
            try:
                textparts.append(fr.inner_text("body"))
            except Exception:
                pass
            try:
                frows = fr.evaluate(
                    """() => {
                        const out = [];
                        document.querySelectorAll('tr').forEach(tr => {
                            const cells = [...tr.querySelectorAll('td,th')]
                                .map(td => (td.innerText || '').trim());
                            if (cells.some(c => c)) out.push(cells);
                        });
                        return out;
                    }"""
                )
                rows.extend(frows or [])
            except Exception:
                pass
        browser.close()
    return rows, "\n".join([t for t in textparts if t])


def _last_number(cells):
    """Última celda que sea un número (columna PRECIO DE LISTA)."""
    for c in reversed(cells):
        m = re.search(r"\d+[.,]\d+", c or "")
        if m:
            return float(m.group(0).replace(",", "."))
    return None


def scrape_repsol_fuel():
    """
    El reporte es un ReportViewer (SSRS) que arma la tabla por JavaScript dentro de un iframe.
    Se renderiza con navegador, se leen las filas por celdas y se toma la columna
    PRECIO DE LISTA (con IGV) = última celda numérica de la fila del producto.
    """
    rows, text = _repsol_report_extract(REPSOL_FUEL_REPORT, timeout=70000)
    if not rows and not re.search(r"DIESEL|GASOHOL", text.upper()):
        raise RuntimeError("El reporte PrecioLima no expuso la tabla (revisar render/timeout)")
    # Método robusto: en este reporte la columna PRECIO DE LISTA siempre tiene 5
    # decimales (p.ej. 21.67660, 16.34583), mientras las demás columnas tienen 2-4.
    # Por cada producto tomamos el primer número de 5 decimales que aparece tras su
    # nombre, sobre todo el texto junto. Así no importa cómo se partan las celdas.
    patrones = {
        "diesel_uv":    r"DIESEL\s*B5[\s\-]*S-?50\s*UV",     # DIESEL B5-S50 UV
        "diesel_s50":   r"DIESEL\s*B5\s*\(S-?50\)",          # DIESEL B5 (S-50)
        "gasohol_prem": r"GASOHOL\s*PREMIUM",
        "gasohol_reg":  r"GASOHOL\s*REGULAR",
        "glp_kg":       r"GLP\s*GRANEL",
    }
    fila = {"planta": "La Pampilla (Callao)", "abastece": "Ancón (Lima)",
            "diesel_uv": None, "diesel_s50": None, "gasohol_reg": None,
            "gasohol_prem": None, "glp_kg": None}

    def _plausible(campo, val):
        if val is None:
            return None
        if campo == "glp_kg":
            return val if 1.2 <= val <= 8 else None         # S/kg
        return val if 10 <= val <= 45 else None             # S/galón (diésel/gasohol)

    blob = " ".join(" ".join(c for c in r) for r in rows) + " " + (text or "")
    blob = re.sub(r"\s+", " ", blob).upper()
    for campo, pat in patrones.items():
        m = re.search(pat + r".{0,70}?(\d+\.\d{5})", blob)   # 1er precio de 5 decimales
        if m:
            fila[campo] = _plausible(campo, float(m.group(1)))
    log(f"Repsol combustible: diesel_uv={fila['diesel_uv']} gasohol_reg={fila['gasohol_reg']} prem={fila['gasohol_prem']} glp={fila['glp_kg']}")
    # vigencia del propio reporte ("Vigencia: 4/08/2026")
    mv = re.search(r"VIGENCIA[:\s]*(\d{1,2})[/.](\d{1,2})[/.](\d{4})", text.upper())
    vig = f"{mv.group(3)}-{int(mv.group(2)):02d}-{int(mv.group(1)):02d}" if mv else _today()
    return {"lista": "PrecioLima (La Pampilla)", "vigente_desde": vig,
            "nota": "Reporte oficial Repsol/RELAPASAA · Lima. Precio de lista con IGV (no incluye recargo FISE).",
            "plantas": [fila]}


# =====================================================================
#  Helpers de metadatos
# =====================================================================
def _find_code(text, pat):
    m = re.search(pat, text, flags=re.I)
    return m.group(0).upper() if m else None


_MESES = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,"julio":7,
          "agosto":8,"setiembre":9,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}

def _find_date(text):
    m = re.search(r"(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚ]+)\s+de[l]?\s+(\d{4})", text, flags=re.I)
    if m:
        d, mes, y = int(m.group(1)), _MESES.get(m.group(2).lower()), int(m.group(3))
        if mes:
            return f"{y:04d}-{mes:02d}-{d:02d}"
    m = re.search(r"(\d{2})[/.](\d{2})[/.](\d{4})", text)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return _today()

def _today():
    return dt.date.today().isoformat()


# =====================================================================
#  Historial (línea de tiempo persistente)
# =====================================================================
def load_json(path, default):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return default

def append_point(hist, key, fecha, valor):
    """Registra un punto solo cuando el precio CAMBIA (no a diario), para que la
    línea de tendencia quede limpia y gerencial."""
    serie = hist.setdefault(key, [])
    if valor is None:
        return
    valor = round(float(valor), 4)
    # si ya hay un punto con esa fecha, actualiza su valor en vez de duplicar
    for p in serie:
        if p["fecha"] == fecha:
            p["valor"] = valor
            return
    # solo agrega un punto nuevo si el valor cambió respecto al último registrado
    if serie and abs(serie[-1]["valor"] - valor) < 0.0001:
        return
    serie.append({"fecha": fecha, "valor": valor})
    serie.sort(key=lambda p: p["fecha"])


# =====================================================================
#  Ensamblado del data.json (esquema idéntico al reporte v5)
# =====================================================================
# ---------------- Proyección del crudo (escenarios; revisar con cada STEO mensual de la EIA) ----------------
PROYECCION_BASE = {
    # Proyección base del Brent (EIA STEO nov-2026): ~US$100 hasta fin de 2026,
    # bajando a ~US$86 promedio en 2027 (por debajo de 90 en el 2T27).
    # Trayectoria de los próximos 6 meses (niveles US$/bbl, se fechan hacia adelante).
    "proj_valores": [98, 95, 92, 90, 88, 86],
    "passthrough": 0.5,                       # traslado Brent -> diésel local (proyección)
    "umbral_alza": 100, "umbral_base": 85,    # para el tablero gerencial
    "esc_ref": {"base": 90, "alza": 105, "baja": 80},
}


def _meses_adelante(n):
    """Devuelve n fechas mensuales (día 15) a partir del próximo mes."""
    hoy = dt.date.today()
    y, m = hoy.year, hoy.month
    out = []
    for _ in range(n):
        m += 1
        if m > 12:
            m = 1; y += 1
        out.append(f"{y:04d}-{m:02d}-15")
    return out


def _eia_rbrte(key, frequency, length, timeout=60, intentos=3):
    """Consulta la serie RBRTE (Europe Brent Spot) a la API v2 de la EIA.
    Devuelve filas [{period, value}] ordenadas de más reciente a más antigua.
    Reintenta ante timeouts transitorios."""
    ultimo = None
    for i in range(intentos):
        try:
            r = requests.get(
                "https://api.eia.gov/v2/petroleum/pri/spt/data/",
                params={
                    "api_key": key, "frequency": frequency, "data[0]": "value",
                    "facets[series][]": "RBRTE",
                    "sort[0][column]": "period", "sort[0][direction]": "desc",
                    "offset": 0, "length": length,
                }, timeout=timeout)
            r.raise_for_status()
            return [row for row in r.json().get("response", {}).get("data", [])
                    if row.get("value") is not None]
        except Exception as e:
            ultimo = e
            if i < intentos - 1:
                import time; time.sleep(3 * (i + 1))
    raise ultimo


def scrape_brent(prev):
    """Brent (Europe Brent Spot Price FOB, serie RBRTE) desde la API oficial de la EIA.
    Requiere EIA_API_KEY (gratis en https://www.eia.gov/opendata/).
    Devuelve dict: spot (último diario), fecha, e hist (realidad mensual ~12 meses).
    Si no hay clave o falla, conserva lo último conocido en el data.json previo."""
    py_prev = prev.get("proyeccion", {}) if isinstance(prev, dict) else {}
    out = {"spot": None, "fecha": None, "hist": []}
    key = os.environ.get("EIA_API_KEY", "").strip()
    if key:
        # 1) spot diario más reciente
        try:
            diario = _eia_rbrte(key, "daily", 5)
            if diario:
                out["spot"] = round(float(diario[0]["value"]), 2)
                out["fecha"] = diario[0].get("period")
                log(f"[BRENT] EIA RBRTE spot = {out['spot']} ({out['fecha']})")
        except Exception as e:
            log(f"[BRENT] API EIA (diario) falló: {e}. Se conserva el Brent previo.")
        # 2) histórico mensual (realidad de los últimos ~12 meses) -> línea única
        try:
            mens = _eia_rbrte(key, "monthly", 12)
            hist = []
            for row in reversed(mens):            # de más antiguo a más reciente
                per = row.get("period")            # "YYYY-MM"
                if per and len(per) == 7:
                    hist.append({"fecha": per + "-15", "valor": round(float(row["value"]), 2)})
            if hist:
                out["hist"] = hist
                log(f"[BRENT] EIA RBRTE mensual: {len(hist)} meses "
                    f"({hist[0]['fecha'][:7]} → {hist[-1]['fecha'][:7]})")
        except Exception as e:
            log(f"[BRENT] API EIA (mensual) falló: {e}. Uso histórico previo.")
    else:
        log("[BRENT] Falta EIA_API_KEY (gratis en eia.gov/opendata). Se conserva el Brent previo.")

    # respaldos desde el data.json anterior
    if out["spot"] is None and py_prev.get("brent_spot"):
        out["spot"] = float(py_prev["brent_spot"]); out["fecha"] = py_prev.get("brent_fecha")
    if not out["hist"] and py_prev.get("brent_hist"):
        out["hist"] = py_prev["brent_hist"]
    if out["spot"] is None:
        out["spot"] = 88.80                       # semilla inicial
    return out


def assemble(comb_petro, asf_petro, asf_repsol, comb_repsol, brent, prev):
    hoy = _today()
    hist = load_json(HIST, {})

    # ---- "foto" del día para cada serie (se acumula solo si el valor cambió) ----
    def callao_uv():
        # Diésel B5 UV S-50 de Petroperú · planta Callao (abastece Ancón, con impuestos)
        for p in comb_petro["plantas"]:
            if p["planta"].upper().startswith("CALLAO"):
                return p.get("diesel_uv")
        # respaldo: Conchán (también abastece Ancón) y luego la primera planta disponible
        for p in comb_petro["plantas"]:
            if p["planta"].upper().startswith("CONCH"):
                return p.get("diesel_uv")
        return comb_petro["plantas"][0].get("diesel_uv") if comb_petro["plantas"] else None
    def repsol_uv():
        return comb_repsol["plantas"][0]["diesel_uv"] if comb_repsol["plantas"] else None
    def pen_petro():
        # Cemento asfáltico PEN 40/50 a 120/150 (el grado que lleva Henry en su Excel)
        items = asf_petro.get("items") or []
        return items[0]["precio"] if items else None

    append_point(hist, "diesel_petroperu", comb_petro.get("vigente_desde") or hoy, callao_uv())
    append_point(hist, "diesel_repsol", comb_repsol.get("vigente_desde") or hoy, repsol_uv())
    append_point(hist, "asfalto_petroperu", asf_petro.get("vigente_desde") or hoy, pen_petro())
    append_point(hist, "asfalto_repsol", asf_repsol.get("vigente_desde") or hoy,
                 asf_repsol["items"][0]["precio"] if asf_repsol.get("items") else None)
    Path(HIST).write_text(json.dumps(hist, ensure_ascii=False, indent=2), encoding="utf-8")

    series = [
        {"clave":"diesel_petroperu","etiqueta":"Diésel B5 UV S-50 · Petroperú (Callao)","color":"#E9A020",
         "fuente":{"tipo":"comb_planta","planta":"Callao","campo":"diesel_uv"},
         "puntos":hist.get("diesel_petroperu",[])},
        {"clave":"diesel_repsol","etiqueta":"Diésel B5 S-50 · Repsol (La Pampilla)","color":"#C77D3A",
         "fuente":{"tipo":"comb_repsol","campo":"diesel_uv"},"puntos":hist.get("diesel_repsol",[])},
        {"clave":"asfalto_petroperu","etiqueta":"Cemento asfáltico PEN · Petroperú","color":"#5B7288",
         "fuente":{"tipo":"asfalto_petro_item","idx":0,"campo":"precio"},
         "puntos":hist.get("asfalto_petroperu",[])},
        {"clave":"asfalto_repsol","etiqueta":"Cemento asfáltico CA 60/70 · Repsol","color":"#3E8E7E",
         "fuente":{"tipo":"asfalto_repsol_item","grado":"CA 60/70","campo":"precio"},
         "puntos":hist.get("asfalto_repsol",[])},
    ]

    # ---------------- Proyección: UNA línea = realidad (Brent mensual) + 6 meses hacia adelante ----------------
    proyeccion = dict(PROYECCION_BASE)
    proyeccion["brent_spot"]  = brent.get("spot")
    proyeccion["brent_fecha"] = brent.get("fecha")
    proyeccion["brent_hist"]  = brent.get("hist", [])          # realidad (línea continua)

    # trayectoria futura anclada al último dato real de Brent
    base_hist = brent.get("hist", [])
    ancla = base_hist[-1]["valor"] if base_hist else (brent.get("spot") or 90.0)
    fechas_fut = _meses_adelante(len(PROYECCION_BASE["proj_valores"]))
    niveles = PROYECCION_BASE["proj_valores"]
    # suaviza el empalme: desplaza la trayectoria para que arranque desde el último real
    ajuste = ancla - niveles[0]
    proj = [{"fecha": f, "valor": round(v + ajuste * (1 - i/len(niveles)), 2)}
            for i, (f, v) in enumerate(zip(fechas_fut, niveles))]
    # el primer punto proyectado se une a la realidad (mismo ancla el mes previo)
    proyeccion["brent_proj"] = ([{"fecha": base_hist[-1]["fecha"], "valor": ancla}] if base_hist else []) + proj

    # proyección del diésel local (línea punteada en Tendencia) por traslado del Brent
    dpp = hist.get("diesel_petroperu", [])
    if dpp and base_hist:
        d_ancla = dpp[-1]["valor"]
        pt = PROYECCION_BASE["passthrough"]
        proyeccion["diesel_proj"] = ([{"fecha": dpp[-1]["fecha"], "valor": d_ancla}] +
            [{"fecha": p["fecha"], "valor": round(d_ancla + (p["valor"] - ancla) * pt, 2)}
             for p in proj])
    else:
        proyeccion["diesel_proj"] = []

    # destinatarios / URL de Power Automate guardados desde el dashboard
    alertas = load_json(CONFIG_ALERTAS, {"destinatarios": [], "webhook_url": ""})

    data = {
        "meta": {"fecha_datos": hoy, "generado": dt.datetime.now().isoformat(timespec="minutes")},
        "combustible": {
            "unidad": "S/ galón · con impuestos (sin FISE ni descuentos) · uso interno S-50",
            "petroperu": comb_petro,
            "repsol": comb_repsol,
        },
        "asfalto": {
            "unidad": "S/ galón + IGV · ex-planta", "igv": 0.18,
            "petroperu": asf_petro,
            "repsol": asf_repsol,
        },
        "timeline": {"unidad": "S/ galón", "series": series},
        "proyeccion": proyeccion,
        "alertas": alertas,
        "fuentes": [
            {"t":"Combustible — Petroperú","d":"Lista de precios en plantas (Automotriz).","f":"3×/día"},
            {"t":"Asfalto — Petroperú","d":"Lista de precios en plantas (Asfaltos).","f":"3×/día"},
            {"t":"Asfalto — Repsol","d":"Lista pública de asfaltos (La Pampilla).","f":"3×/día"},
            {"t":"Combustible — Repsol","d":"Reporte PrecioLima (RELAPASAA).","f":"3×/día"},
            {"t":"Brent — EIA","d":"Europe Brent Spot (RBRTE), API oficial EIA.","f":"3×/día"},
        ],
    }
    return data


def with_fallback(fn, prev_getter, label):
    """Ejecuta un scraper; si falla, usa el valor previo del data.json anterior."""
    try:
        return fn()
    except Exception as e:
        log(f"[AVISO] {label} falló: {e}. Se conserva el último valor conocido.")
        return prev_getter()


def serie_actual(data, fuente):
    """Valor actual de una serie según su 'fuente' (espejo del resolver del reporte)."""
    t = (fuente or {}).get("tipo")
    try:
        if t == "comb_planta":
            for p in data["combustible"]["petroperu"]["plantas"]:
                if p["planta"].upper().startswith(fuente["planta"].upper()):
                    return p.get(fuente["campo"])
        elif t == "asfalto_banda":
            return data["asfalto"]["petroperu"]["banda"].get(fuente["campo"])
        elif t == "asfalto_petro_item":
            items = data["asfalto"]["petroperu"].get("items") or []
            idx = fuente.get("idx", 0)
            return items[idx].get(fuente["campo"]) if idx < len(items) else None
        elif t == "asfalto_repsol_item":
            for it in data["asfalto"]["repsol"]["items"]:
                if it["grado"] == fuente["grado"]:
                    return it.get(fuente["campo"])
        elif t == "comb_repsol":
            p = data["combustible"]["repsol"]["plantas"][0]
            return p.get(fuente["campo"])
    except Exception:
        return None
    return None


def detectar_cambios(prev, data):
    """Compara precios actuales (prev vs nuevo) por cada serie de la línea de tiempo."""
    cambios = []
    for s in data.get("timeline", {}).get("series", []):
        f = s.get("fuente")
        nueva = serie_actual(data, f)
        antes = serie_actual(prev, f) if prev else None
        if nueva is not None and antes is not None and abs(nueva - antes) > 0.0001:
            cambios.append({
                "serie": s.get("etiqueta", s.get("clave", "")),
                "antes": round(float(antes), 4),
                "ahora": round(float(nueva), 4),
                "delta": round(float(nueva) - float(antes), 4),
                "dir": "ALZA" if nueva > antes else "BAJA",
            })
    return cambios


def enviar_alerta(cambios):
    """Registra los cambios en alertas.log y, si hay SMTP configurado, manda correo."""
    if not cambios:
        return
    fecha = dt.datetime.now().isoformat(timespec="minutes")
    lineas = [f"[{fecha}]"]
    for c in cambios:
        signo = "▲" if c["dir"] == "ALZA" else "▼"
        lineas.append(f"  {c['dir']} {signo}  {c['serie']}: "
                      f"S/ {c['antes']:.4f} -> S/ {c['ahora']:.4f} "
                      f"({'+' if c['delta']>=0 else ''}{c['delta']:.4f})")
    cuerpo = "\n".join(lineas)
    try:
        with open(ALERTA_LOG, "a", encoding="utf-8") as f:
            f.write(cuerpo + "\n")
    except Exception:
        pass
    log(f"Cambios detectados:\n{cuerpo}")

    destinatarios = _destinatarios()
    n = len(cambios)
    subj = f"[Precios UNNA] {n} cambio(s) de precio · " + \
           ", ".join(sorted(set(c["dir"] for c in cambios)))

    # 1) Vía Power Automate (recomendada en entorno Microsoft 365)
    url = _webhook_url()
    if url and destinatarios:
        try:
            requests.post(url, json={
                "asunto": subj,
                "cuerpo": "Cambios en los precios de lista (Petroperú / Repsol):\n\n" + cuerpo,
                "destinatarios": ";".join(destinatarios),
            }, timeout=30)
            log(f"[ALERTA] Enviado por Power Automate a {', '.join(destinatarios)}")
            return
        except Exception as e:
            log(f"[ALERTA] Power Automate falló ({e}); intento SMTP si está configurado.")

    # 2) Vía SMTP (si no hay flujo o falló)
    if not (SMTP_HOST and SMTP_USER and SMTP_PASS and destinatarios):
        log("[ALERTA] Sin Power Automate ni SMTP configurados: cambio registrado, sin correo.")
        return
    try:
        msg = EmailMessage()
        msg["Subject"] = subj
        msg["From"] = ALERTA_FROM
        msg["To"] = ", ".join(destinatarios)
        msg.set_content(
            "Se detectaron cambios en los precios de lista (Petroperú / Repsol):\n\n"
            + cuerpo +
            "\n\nEste aviso lo genera el servicio automático de precios (CDA / Survial)."
        )
        ctx = ssl.create_default_context()
        if SMTP_PORT == 465:
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx) as s:
                s.login(SMTP_USER, SMTP_PASS); s.send_message(msg)
        else:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
                s.starttls(context=ctx); s.login(SMTP_USER, SMTP_PASS); s.send_message(msg)
        log(f"[ALERTA] Correo enviado a {', '.join(destinatarios)}")
    except Exception as e:
        log(f"[ALERTA] No se pudo enviar el correo: {e}")


def main():
    prev = load_json(OUT, {})
    def prev_or(path_keys, default):
        node = prev
        for k in path_keys:
            node = node.get(k, {}) if isinstance(node, dict) else {}
        return node or default

    comb_petro = with_fallback(scrape_petro_comb,
                    lambda: prev_or(["combustible","petroperu"], {"lista":"COMB","vigente_desde":_today(),"plantas":[]}),
                    "Petroperú combustible")
    asf_petro = with_fallback(scrape_petro_asf,
                    lambda: prev_or(["asfalto","petroperu"], {"lista":"ASFA","vigente_desde":_today(),"plantas":"Talara / Conchán","items":[],"banda":{"min":None,"max":None}}),
                    "Petroperú asfalto")
    asf_repsol = with_fallback(scrape_repsol_asf,
                    lambda: prev_or(["asfalto","repsol"], {"lista":"Recosac","vigente_desde":_today(),"nota":"","items":[]}),
                    "Repsol asfalto")
    comb_repsol = with_fallback(scrape_repsol_fuel,
                    lambda: prev_or(["combustible","repsol"], {"lista":"PrecioLima","vigente_desde":_today(),"nota":"","plantas":[{"planta":"La Pampilla (Callao)","abastece":"Ancón (Lima)","diesel_uv":None,"diesel_s50":None,"gasohol_reg":None,"gasohol_prem":None,"glp_kg":None}]}),
                    "Repsol combustible")

    brent = scrape_brent(prev)   # (valor, fecha); con respaldo interno al último conocido

    data = assemble(comb_petro, asf_petro, asf_repsol, comb_repsol, brent, prev)
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"OK -> {OUT}  (fecha_datos={data['meta']['fecha_datos']})")

    # Alertas de alza/baja vs la corrida anterior
    cambios = detectar_cambios(prev, data)
    if cambios:
        enviar_alerta(cambios)
    else:
        log("Sin cambios de precio respecto a la corrida anterior.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"[ERROR FATAL] {e}")
        sys.exit(1)
