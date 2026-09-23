"""
Descargador automático de cartas glosa — Seguros Mundial
Flujo: login + agregar cartas + descarga de ZIPs del portal.
Estructura final: un solo ZIP por IPS con carpetas DEV/ y LIQ/ + Excel.
"""

import os
import re
import json
import csv
import time
import threading
import logging
import zipfile
import tempfile
import shutil
from datetime import datetime, timezone
from pathlib import Path
from flask import Blueprint, render_template, request, jsonify, send_from_directory
from . import concurrency
from . import historial_db
from . import registro_rutas
from .catalogo_ips import resolver, validar_identidad, normalizar_nit, listar_catalogo_por_responsable
from io import BytesIO

try:
    import openpyxl
    from openpyxl.styles import Font
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    print("openpyxl no instalado. No se generara el archivo Excel.")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

bp = Blueprint("mundial", __name__, url_prefix="/mundial")

BASE_DIR = Path(__file__).resolve().parent.parent  # raíz del proyecto (un nivel arriba de bots/)
port = int(os.environ.get("PORT", 8080))

# En Railway se monta un volumen persistente en /data.
# Localmente (sin volumen) usa la carpeta downloads/ del proyecto.
_data_root = Path(os.environ.get("DATA_DIR", "/data"))
if _data_root.exists():
    DOWNLOAD_DIR = _data_root / "downloads" / "mundial"
else:
    DOWNLOAD_DIR = BASE_DIR / "downloads" / "mundial"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

LOGIN_URL = "https://a2m-mundial.iqdigital.com.co/ATF_Site/"
CONSULTA_CARTAS_URL = "https://a2m-mundial.iqdigital.com.co/ATF_Site/wallet/settlement-letters"

# ==================== MAPA DE IPS POR NIT ====================
MAPA_IPS = {
    "900267064": "INVERSIONES_AZALUD_CLINICA_BAHIA",
    "900827065": "CENTRO_DE_DIAGNOSTICO_E_IMAGENES_BAHIA",
    "900657731": "CENTRO_MEDICO_Y_DE_REHABILITACION_BAHIA",
    "900826509": "RED_DE_URGENCIAS_DEL_MAGDALENA",
    "900513306": "FUNDACION_MARIA_REINA",
    "900600550": "INVERSIONES_MEDICAS_BARU",
    "900954800": "CENTRO_MEDICO_Y_DE_REHABILITACION_BARU",
    "900631361": "INVERSIONES_MEDICAS_VALLESALUD",
    "900257333": "ODONTOTRANS",
    "901081281": "URGETRAUMA",
    "900792417": "RED_DE_URGENCIAS_DE_LA_COSTA_PACIFICA",
    "901959993": "CLINICA_CORDIALIDAD",
    "900002780": "FUNDACION_CAMPBELL",
    "901523868": "MOVID_IPS_SAS",
    "901057487": "TECNOLOGIA_DIAGNOSTICA_DEL_VALLE",
    "900558595": "FUNDACION_MEDICA_CAMPBELL",
    "901149757": "UNIDAD_MEDICA_DE_TRAUMA_VALLE_SALUD",
    "900900754": "CLINICA_VALLE_SALUD_SAN_FERNANDO",
    "900469882": "CENTRO_MEDICO_SERVISALUD_INTEGRAL_IPS_SAS",
    "802024329": "RED_DE_URGENCIA_DE_LA_COSTA_LTDA",
    "900847382": "CENTRO_MEDICO_Y_DE_REHABILITACION_VALLE_SALUD",
}

# Mapa de correos electrónicos a IPS (para usuarios que usan email en lugar de NIT)
MAPA_CORREOS = {
    "carteracdibahia@gmail.com":          ("900827065", "CENTRO_DE_DIAGNOSTICO_E_IMAGENES_BAHIA"),
    "carteraclinicabahia@gmail.com":       ("900267064", "INVERSIONES_AZALUD_CLINICA_BAHIA"),
    "carteracentromedicobahia@gmail.com":  ("900657731", "CENTRO_MEDICO_Y_DE_REHABILITACION_BAHIA"),
    "carteraimbbaru@gmail.com":            ("900600550", "INVERSIONES_MEDICAS_BARU"),
    "carteracentromedicobahia@gmail.com":  ("900657731", "CENTRO_MEDICO_Y_DE_REHABILITACION_BAHIA"),
    "carterarucmagdalena@gmail.com":       ("900826509", "RED_DE_URGENCIAS_DEL_MAGDALENA"),
    "glosas@salud-net.com":               ("901081281", "URGETRAUMA"),
    "auditoria@salud-net.com":             ("900631361", "INVERSIONES_MEDICAS_VALLESALUD"),
    "devolucion.objecion@salud-net.com":   ("900002780", "CMQ_ALVERNIA"),
    "vallesaludc@gmail.com":              ("900847382", "CENTRO_MEDICO_Y_DE_REHABILITACION_VALLE_SALUD"),
}

def resolver_ips_por_usuario(usuario: str):
    u = (usuario or "").strip().lower()
    # 1. Si el usuario es un correo electrónico, buscar en el mapa de correos
    if "@" in u:
        if u in MAPA_CORREOS:
            nit, nombre = MAPA_CORREOS[u]
            return nit, nombre
        # Buscar por coincidencia parcial del correo
        for correo, (nit, nombre) in MAPA_CORREOS.items():
            if u == correo.lower():
                return nit, nombre
        return None, "IPS_DESCONOCIDA"
    # 2. Si tiene dígitos, extraer NIT
    m = re.search(r"(\d{9,12})", usuario or "")
    nit = m.group(1) if m else None
    nombre = MAPA_IPS.get(nit, "IPS_DESCONOCIDA") if nit else "IPS_DESCONOCIDA"
    return nit, nombre

def detectar_tipo_solicitud(valor: str):
    v = (valor or "").strip().upper()
    if v.startswith("DEV") or v.startswith("OBJ"):
        return "No Dev/Obj"
    if v.startswith("LIQ"):
        return "No Liquidación"
    # Cualquier otro prefijo (CMV, CMVIQ, o lo que sea) es "No Radicado".
    return "No Radicado"

# ==================== REGISTRO DE JOBS POR EMPRESA/IPS ====================
# Cada empresa (resuelta por NIT o correo del usuario) tiene su propio job
# aislado: su propio estado, lock y navegador. Dos empresas distintas pueden
# correr en paralelo; la MISMA empresa sigue bloqueada si ya tiene un proceso
# en curso.
MAX_CONCURRENT_JOBS = int(os.environ.get("MAX_CONCURRENT_JOBS", "2"))

jobs = {}  # empresa_id -> job dict
jobs_registry_lock = threading.RLock()


def resolve_empresa_id(usuario: str) -> str:
    """Resuelve el identificador de 'empresa' reutilizando la misma lógica
    que ya usa el bot para mapear usuario -> NIT/IPS (resolver_ips_por_usuario).
    Si no se puede resolver un NIT, se usa el propio usuario (aislado)."""
    nit, _nombre = resolver_ips_por_usuario(usuario)
    if nit:
        return nit
    return (usuario or "").strip().lower()


def new_job_state():
    return {
        "running": False,
        "stopping": False,
        "logs": [],
        "stats": {"total": 0, "descargadas": 0, "errores": 0},
        "finished": False,
        "error": None,
        "errores_detalle": [],
        "descargas_exitosas": [],
        "errores_excel_url": None,
        "zip_url": None,
    }


def get_or_create_job(empresa_id: str):
    with jobs_registry_lock:
        if empresa_id not in jobs:
            jobs[empresa_id] = {
                "state": new_job_state(),
                "lock": threading.Lock(),
                "browser": None,
                "context": None,
                "dl_dir": None,
                "lote": None,
                "ips_nombre": None,
            }
        return jobs[empresa_id]


def get_job_or_none(empresa_id: str):
    with jobs_registry_lock:
        return jobs.get(empresa_id)


def count_running_jobs() -> int:
    with jobs_registry_lock:
        return sum(1 for j in jobs.values() if j["state"]["running"])

# ==================== LOGGING CON COLORES ====================
def log(job, msg, level="info"):
    ts = datetime.now().strftime("%H:%M:%S")
    colors = {
        "info": "\033[94m",
        "success": "\033[92m",
        "warn": "\033[93m",
        "error": "\033[91m",
    }
    reset = "\033[0m"
    color = colors.get(level, colors["info"])
    if level == "success":
        print(f"{color}[{ts}] {msg}{reset}")
    elif level == "error":
        print(f"{color}[{ts}] ERROR: {msg}{reset}")
    elif level == "warn":
        print(f"{color}[{ts}] ADVERTENCIA: {msg}{reset}")
    else:
        print(f"{color}[{ts}] {msg}{reset}")

    entry = {"ts": ts, "msg": msg, "level": level}
    if job is not None:
        with job["lock"]:
            job["state"]["logs"].append(entry)
    if level == "error":
        logger.error(msg)
    else:
        logger.info(msg)
    if level == "error" and job is not None:
        try:
            historial_db.registrar_error_ejecucion(
                job["state"].get("ejecucion_id"), None, "mundial", "Mundial",
                job["state"].get("ips_identity"), "bot", msg, "log", True, 1, "error"
            )
        except Exception:
            pass

def reset_state(job):
    with job["lock"]:
        job["state"] = new_job_state()

def stop_job(job):
    with job["lock"]:
        job["state"]["stopping"] = True
    log(job, "Solicitando detencion del proceso...", "warn")
    # No cerrar el browser desde este hilo — Playwright requiere que se cierre
    # desde el mismo hilo que lo creó. El job lo cerrará al detectar stopping=True.

def generar_zip_parcial(job):
    """Genera un ZIP con lo que se alcanzó a descargar hasta el momento
    (carpetas DEV/LIQ/RAD + un Excel parcial), para no perder el trabajo
    si el proceso se cae por un error antes de llegar al final normal."""
    dl_dir = job.get("dl_dir")
    ips_nombre = job.get("ips_nombre")
    if not dl_dir or not ips_nombre:
        return
    ips_dir = Path(dl_dir) / ips_nombre
    if not ips_dir.exists():
        return

    dev_dir = ips_dir / "DEV"
    liq_dir = ips_dir / "LIQ"
    rad_dir = ips_dir / "RAD"

    with job["lock"]:
        exitosas = job["state"]["descargas_exitosas"].copy()
        errores = job["state"]["errores_detalle"].copy()

    if not exitosas and not errores:
        return  # nada que empacar todavía

    excel_path = None
    if EXCEL_AVAILABLE:
        try:
            excel_path = ips_dir / "reporte_parcial.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Cartas procesadas"
            ws.append(["Consecutivo", "Tipo solicitud", "Carpeta", "Fecha/Hora"])
            for ex in exitosas:
                carpeta = {"No Dev/Obj": "DEV", "No Liquidación": "LIQ", "No Radicado": "RAD"}.get(ex.get("tipo"), "DEV")
                ws.append([ex.get("consecutivo"), ex.get("tipo"), carpeta, ex.get("timestamp")])
            if errores:
                wse = wb.create_sheet("Errores")
                wse.append(["Consecutivo", "Tipo solicitud", "Error", "Captura", "Fecha/Hora"])
                for er in errores:
                    wse.append([er.get("consecutivo"), er.get("tipo"), er.get("error"), er.get("captura"), er.get("timestamp")])
            wb.save(excel_path)
        except Exception as e:
            log(job, f"No se pudo generar el Excel parcial: {e}", "warn")
            excel_path = None

    try:
        for viejo in Path(dl_dir).glob(f"{ips_nombre}_PARCIAL_*.zip"):
            try:
                viejo.unlink()
            except Exception:
                pass
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_path = Path(dl_dir) / f"{ips_nombre}_PARCIAL_{timestamp}.zip"
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for carpeta, nombre in ((dev_dir, "DEV"), (liq_dir, "LIQ"), (rad_dir, "RAD")):
                if carpeta.exists():
                    for file_path in carpeta.rglob("*"):
                        if file_path.is_file():
                            zf.write(file_path, f"{nombre}/{file_path.relative_to(carpeta)}")
            if excel_path and excel_path.exists():
                zf.write(excel_path, "reporte_parcial.xlsx")
            errores_dir = ips_dir / "Errores"
            if errores_dir.exists():
                for file_path in errores_dir.rglob("*"):
                    if file_path.is_file():
                        zf.write(file_path, f"Errores/{file_path.relative_to(errores_dir)}")
        log(job, f"📦 ZIP parcial generado con lo descargado hasta el momento: {zip_path.name}", "warn")
    except Exception as e:
        log(job, f"No se pudo generar el ZIP parcial: {e}", "error")
    finally:
        # No se borran DEV/LIQ/RAD ni progreso.json: si se reinicia el proceso,
        # debe poder seguir acumulando y saltar lo ya hecho.
        if excel_path and excel_path.exists():
            try:
                excel_path.unlink()
            except Exception:
                pass

# ==================== PERSISTENCIA ====================
def cargar_progreso(job, ips_dir):
    p = ips_dir / "progreso.json"
    if p.exists():
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            completadas = data.get("completadas", [])
            return set(completadas) if isinstance(completadas, list) else set()
        except Exception as e:
            log(job, f"Error al leer progreso: {e}", "warn")
    return set()

def guardar_progreso(job, ips_dir, completadas, nuevo_item=None, meta=None):
    """
    Además de la lista plana de completadas (para reanudar, igual que
    siempre), guarda un detalle por consecutivo (tipo + fecha real) y
    metadata (identidad, ips, aseguradora, lote) — necesario para poder
    migrar esto a la base de datos de historial sin perder información.
    Lectura-fusión-escritura: cada llamada solo pasa el ítem NUEVO.
    """
    p = ips_dir / "progreso.json"
    detalle = {}
    if p.exists():
        try:
            with open(p, "r", encoding="utf-8") as f:
                detalle = json.load(f).get("detalle", {}) or {}
        except Exception:
            pass
    if nuevo_item:
        detalle[str(nuevo_item["factura"])] = {
            "tipo": nuevo_item.get("tipo"),
            "fecha_descarga": nuevo_item.get("fecha_descarga") or datetime.now(timezone.utc).isoformat(),
        }
    try:
        data = {"completadas": list(completadas), "detalle": detalle, "actualizado": datetime.now(timezone.utc).isoformat()}
        if meta:
            data["meta"] = meta
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        log(job, f"⚠️ Error al escribir progreso.json: {e}", "warn")

    try:
        historial_db.registrar_descargas(
            aseguradora=(meta or {}).get("aseguradora", "Desconocida"),
            ips_nombre=(meta or {}).get("ips_nombre", "IPS_NO_IDENTIFICADA"),
            periodo=(meta or {}).get("periodo"), identidad=(meta or {}).get("identidad"),
            items=[{"factura": k, "siniestro": v.get("tipo"), "fecha_descarga": v.get("fecha_descarga")} for k, v in detalle.items()],
            ips=(meta or {}).get("ips"),
        )
    except Exception as e:
        log(job, f"⚠️ Error al registrar en el historial (tabla descargas): {e}", "warn")

    try:
        historial_db.registrar_facturas_ejecucion(
            (meta or {}).get("ejecucion_id"), "mundial", (meta or {}).get("aseguradora"),
            (meta or {}).get("ips"), (meta or {}).get("periodo"),
            [{"factura": k, "estado": "exitosa", "fecha_fin": v.get("fecha_descarga")} for k, v in detalle.items()],
        )
    except Exception as e:
        log(job, f"⚠️ Error al registrar la ejecución (tabla ejecucion_facturas): {e}", "warn")

# ==================== UTILIDADES PLAYWRIGHT ====================
def _texto_pagina(page):
    try:
        return (page.evaluate("() => document.body ? document.body.innerText : ''") or "")
    except Exception:
        return ""

def _esperar_texto(job, page, regex, timeout=30, intervalo=0.5):
    fin = time.time() + timeout
    pat = re.compile(regex, re.I)
    while time.time() < fin:
        if job["state"].get("stopping"):
            return False
        if pat.search(_texto_pagina(page)):
            return True
        time.sleep(intervalo)
    return False

def _click_por_texto(page, texto, exacto=False, timeout=8000):
    estrategias = []
    if exacto:
        estrategias.append(lambda: page.get_by_text(texto, exact=True).first)
    estrategias.append(lambda: page.get_by_text(texto).first)
    estrategias.append(lambda: page.locator(f"text={texto}").first)
    estrategias.append(lambda: page.get_by_role("button", name=re.compile(re.escape(texto), re.I)).first)
    for estrategia in estrategias:
        try:
            loc = estrategia()
            loc.click(timeout=timeout)
            return True
        except Exception:
            continue
    try:
        clicked = page.evaluate(
            """([t, ex]) => {
                const norm = s => (s||'').toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').trim();
                const target = norm(t);
                const els = document.querySelectorAll('button, a, span, div, li, td, [role="button"]');
                for (const el of els) {
                    const txt = norm(el.textContent);
                    if (ex ? (txt === target) : txt.includes(target)) {
                        if (txt.length <= target.length + 40) { el.click(); return true; }
                    }
                }
                return false;
            }""",
            [texto, exacto],
        )
        return bool(clicked)
    except Exception:
        return False

# ==================== LOGIN ====================
def _hacer_login(job, page, usuario, password, login_timeout):
    log(job, "Abriendo portal de Seguros Mundial...")
    page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
    time.sleep(2)

    # Detectar si el usuario es correo electrónico
    es_correo = "@" in (usuario or "")
    tipo_label = "Correo electrónico" if es_correo else "Usuario"
    log(job, f"  Tipo de acceso detectado: {tipo_label}")

    try:
        select_locator = page.locator('mat-select[name="typeAccess"]').first
        select_locator.click(timeout=5000)
        log(job, "  Clic en desplegable 'Tipo de Acceso'.")
        time.sleep(0.5)
        if es_correo:
            # "Correo electrónico" es la segunda opción: ArrowDown x2
            page.keyboard.press("ArrowDown")
            time.sleep(0.2)
            page.keyboard.press("ArrowDown")
            time.sleep(0.2)
            page.keyboard.press("Enter")
            log(job, "  Teclas: Flecha Abajo x2 + Enter para seleccionar 'Correo electrónico'.")
        else:
            # "Usuario" es la primera opción: ArrowDown x1
            page.keyboard.press("ArrowDown")
            time.sleep(0.3)
            page.keyboard.press("Enter")
            log(job, "  Teclas: Flecha Abajo + Enter para seleccionar 'Usuario'.")
        time.sleep(0.5)
    except Exception as e:
        log(job, f"  No se pudo seleccionar tipo con teclado: {e}.", "warn")
        try:
            texto_opcion = re.compile(r"^Correo electrónico$", re.I) if es_correo else re.compile(r"^Usuario$", re.I)
            page.locator('mat-option, .mat-option', has_text=texto_opcion).first.click(timeout=5000)
            log(job, f"  Opción '{tipo_label}' seleccionada por clic directo (fallback).")
        except Exception as e2:
            log(job, f"  Fallback también falló: {e2}", "warn")

    try:
        page.fill('input[name="user"]', usuario, timeout=5000)
        log(job, "  Campo 'Usuario' llenado.")
    except Exception:
        page.locator('input:not([type="password"])').first.fill(usuario)
        log(job, "  Campo 'Usuario' llenado (fallback).")

    try:
        page.fill('input[name="password"]', password, timeout=5000)
        log(job, "  Campo 'Contraseña' llenado.")
    except Exception:
        page.locator('input[type="password"]').first.fill(password)
        log(job, "  Campo 'Contraseña' llenado (fallback).")

    try:
        page.click('button.btn-login_v2, button[type="submit"]', timeout=5000)
        log(job, "  Clic en 'Ingresar' enviado.")
    except Exception as e:
        raise Exception(f"No se pudo hacer clic en Ingresar: {e}")

    log(job, "🔐 Inicio de sesión automático para Seguros Mundial. Si aparece un reCAPTCHA, resuélvelo manualmente en el navegador.")
    time.sleep(5)
    fin = time.time() + login_timeout
    while time.time() < fin:
        if job["state"].get("stopping"):
            return False
        txt = _texto_pagina(page)
        if re.search(r"Inicio|DOCUMENTOS DE AYUDA|Cartera IPS|Consulta de Cartas", txt, re.I):
            log(job, "Login exitoso, menú principal visible.")
            time.sleep(2)
            return True
        if "/home" in page.url or "/wallet" in page.url:
            # URL cambió pero Angular puede estar aún inicializando.
            # Esperar a que el menú principal sea visible antes de continuar.
            log(job, "URL interna detectada. Esperando que Angular cargue el menú...")
            for _ in range(15):
                if job["state"].get("stopping"):
                    return False
                txt2 = _texto_pagina(page)
                if re.search(r"Inicio|DOCUMENTOS DE AYUDA|Cartera IPS|Menú Principal|Cerrar Sesión", txt2, re.I):
                    log(job, "Login exitoso, menú principal visible.")
                    time.sleep(2)
                    return True
                time.sleep(1)
            # Si tras 15s no aparece el menú, igual continuamos (página puede estar cargada)
            log(job, "Login exitoso por cambio de URL (menú no detectado, continuando).")
            time.sleep(3)
            return True
        if "Tipo de Acceso" in txt and "Ingresar" in txt:
            log(job, "Aún en página de login. Si hay reCAPTCHA, resuélvelo manualmente en el navegador.", "warn")
        time.sleep(2)

    log(job, "Tiempo de espera normal agotado. Se concede 60 segundos adicionales para resolver captcha manualmente...", "warn")
    extra_time = 60
    fin_extra = time.time() + extra_time
    while time.time() < fin_extra:
        if job["state"].get("stopping"):
            return False
        txt = _texto_pagina(page)
        if re.search(r"Inicio|DOCUMENTOS DE AYUDA|Cartera IPS", txt, re.I):
            log(job, "Login exitoso después de intervención manual.")
            return True
        time.sleep(2)
    raise Exception("Tiempo de espera total agotado. No se pudo completar el login.")

# ==================== FUNCIONES PARA CONSULTA DE CARTAS ====================
def _input_por_label(page, etiqueta_regex):
    try:
        handle = page.evaluate_handle(
            """(re) => {
                const rx = new RegExp(re, 'i');
                const fields = document.querySelectorAll('mat-form-field');
                for (const f of fields) {
                    const label = f.querySelector('mat-label');
                    if (label && rx.test(label.textContent || '')) {
                        const inp = f.querySelector('input');
                        if (inp) return inp;
                    }
                }
                return null;
            }""",
            etiqueta_regex,
        )
        return handle.as_element()
    except Exception:
        return None

def _seleccionar_tipo_solicitud(job, page, tipo):
    # Buscar el campo con reintentos: Angular puede tardar en renderizar el formulario.
    inp = None
    for intento in range(12):  # hasta ~12 segundos
        if job["state"].get("stopping"):
            return
        inp = _input_por_label(page, r"tipo de solicitud|seleccione tipo")
        if inp is None:
            try:
                inp = page.locator('input[role="combobox"], input.mat-autocomplete-trigger, mat-form-field input').first.element_handle(timeout=2000)
            except Exception:
                inp = None
        if inp is not None:
            break
        time.sleep(1)
    if inp is None:
        raise Exception("No se encontró el campo 'Seleccione tipo de solicitud'.")

    inp.click()
    time.sleep(0.4)
    try:
        inp.fill("")
    except Exception:
        pass
    clave = {"No Dev/Obj": "Dev", "No Liquidación": "Liquid", "No Radicado": "Radic"}.get(tipo, tipo)
    inp.type(clave, delay=40)
    time.sleep(1.2)

    opcion_texto = tipo
    seleccionado = False
    for _ in range(3):
        try:
            clicked = page.evaluate(
                """(target) => {
                    const norm = s => (s||'').toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').replace(/\\s+/g,' ').trim();
                    const t = norm(target);
                    const opts = document.querySelectorAll('mat-option, .mat-option, [role="option"]');
                    for (const o of opts) {
                        if (norm(o.textContent) === t) { o.click(); return true; }
                    }
                    for (const o of opts) {
                        if (norm(o.textContent).includes(t)) { o.click(); return true; }
                    }
                    if (opts.length === 1) { opts[0].click(); return true; }
                    return false;
                }""",
                opcion_texto,
            )
            if clicked:
                seleccionado = True
                break
        except Exception:
            pass
        time.sleep(0.8)

    if not seleccionado:
        try:
            page.locator('mat-option', has_text=re.compile(re.escape(tipo.split('/')[0].split()[-1]), re.I)).first.click(timeout=3000)
            seleccionado = True
        except Exception:
            pass

    if not seleccionado:
        raise Exception(f"No se pudo seleccionar el tipo de solicitud '{tipo}'.")
    time.sleep(0.6)
    return True

def _escribir_valor_y_consultar(page, valor):
    inp = _input_por_label(page, r"^\s*valor\s*$|valor")
    if inp is None:
        try:
            inputs = page.locator('app-input input, mat-form-field input')
            inp = inputs.nth(inputs.count() - 1).element_handle(timeout=4000)
        except Exception:
            inp = None
    if inp is None:
        raise Exception("No se encontró el campo 'Valor'.")
    inp.click()
    try:
        inp.fill("")
    except Exception:
        pass
    inp.type(valor, delay=20)
    time.sleep(0.5)

    consultado = False
    for sel in ['button:has-text("Consultar")', 'button.btn-secundario_v2']:
        try:
            page.locator(sel).first.click(timeout=4000)
            consultado = True
            break
        except Exception:
            continue
    if not consultado and not _click_por_texto(page, "Consultar", timeout=4000):
        raise Exception("No se pudo hacer clic en 'Consultar'.")
    time.sleep(2.5)

def _click_agregar(job, page):
    for _ in range(8):
        if job["state"].get("stopping"):
            return False
        try:
            clicked = page.evaluate(
                """() => {
                    const norm = s => (s||'').trim().toLowerCase();
                    const addIcons = document.querySelectorAll('mat-icon, button, a, [role="button"]');
                    for (const el of addIcons) {
                        const txt = norm(el.textContent);
                        const title = norm(el.getAttribute('title') || '');
                        const aria = norm(el.getAttribute('aria-label') || '');
                        if (txt === 'add' || txt === 'add_circle' || txt === 'add_circle_outline' ||
                            title.includes('agregar') || aria.includes('agregar')) {
                            const target = el.closest('button, a, [role="button"]') || el;
                            target.click();
                            return true;
                        }
                    }
                    const cells = document.querySelectorAll('td, th, div[role="cell"]');
                    for (const cell of cells) {
                        if (norm(cell.textContent) === 'agregar') {
                            const btn = cell.querySelector('button, a, [role="button"]');
                            if (btn) { btn.click(); return true; }
                            else { cell.click(); return true; }
                        }
                    }
                    return false;
                }"""
            )
            if clicked:
                time.sleep(1.5)
                return True
        except Exception:
            pass
        time.sleep(1)
    raise Exception("No se pudo dar clic en AGREGAR (¿sin resultados para este valor?).")

def _abrir_buzon(job, page):
    for _ in range(12):
        if job["state"].get("stopping"):
            return False
        try:
            estado = page.evaluate(
                """() => {
                    const icons = document.querySelectorAll('mat-icon');
                    for (const el of icons) {
                        const txt = (el.textContent||'');
                        if (txt.includes('mail_outline') || txt.includes('mail') || txt.includes('email')) {
                            const badge = el.querySelector('.mat-badge-content');
                            const n = badge ? parseInt((badge.textContent||'0').trim()||'0', 10) : 1;
                            return { found: true, count: isNaN(n) ? 0 : n };
                        }
                    }
                    return { found: false, count: 0 };
                }"""
            )
            if estado and estado.get("found") and estado.get("count", 0) >= 1:
                page.evaluate(
                    """() => {
                        const icons = document.querySelectorAll('mat-icon');
                        for (const el of icons) {
                            const txt = (el.textContent||'');
                            if (txt.includes('mail')) {
                                const target = el.closest('button, a, [role="button"]') || el;
                                target.click();
                                return true;
                            }
                        }
                        return false;
                    }"""
                )
                time.sleep(1)
                return True
        except Exception:
            pass
        time.sleep(1)
    raise Exception("El buzón no registró ninguna carta agregada (contador sigue en 0).")

def _descargar_zip_y_extraer(job, page, context, destino_dir, sufijo):
    """
    Descarga el ZIP del buzón, lo descomprime en destino_dir (carpeta con nombre sufijo)
    y devuelve la ruta de la carpeta donde están los PDFs.
    """
    log(job, f"Descargando ZIP para {sufijo}...")
    _abrir_buzon(job, page)

    if not _esperar_texto(job, page, r"Buz[oó]n de cartas|Descargar", timeout=10):
        raise Exception("No apareció el modal del buzón de cartas.")

    # Descargar el ZIP a un archivo temporal
    try:
        with page.expect_download(timeout=60000) as dl_info:
            if not _click_por_texto(page, "Descargar", timeout=8000):
                raise Exception("No se encontró el botón Descargar")
        download = dl_info.value
        # Guardar temporalmente
        temp_zip = destino_dir / f"temp_{sufijo}.zip"
        download.save_as(str(temp_zip))
        log(job, f"ZIP descargado: {temp_zip.name} ({temp_zip.stat().st_size // 1024} KB)")
    except Exception as e:
        log(job, f"Fallo descarga estándar, intentando método alternativo: {e}", "warn")
        pdf_url = None
        try:
            with context.expect_page(timeout=15000) as np_info:
                _click_por_texto(page, "Descargar", timeout=4000)
            np = np_info.value
            for _ in range(20):
                u = np.url
                if u and u != "about:blank":
                    pdf_url = u
                    break
                time.sleep(0.5)
            try:
                np.close()
            except Exception:
                pass
        except Exception:
            pass
        if pdf_url:
            resp = context.request.get(pdf_url, timeout=60000)
            if resp.ok:
                data = resp.body()
                temp_zip = destino_dir / f"temp_{sufijo}.zip"
                temp_zip.write_bytes(data)
                log(job, f"ZIP capturado por URL: {temp_zip.name}")
            else:
                raise Exception("No se pudo obtener el ZIP desde la URL")
        else:
            raise Exception("No se pudo descargar el ZIP del buzón")

    _click_por_texto(page, "Volver", timeout=2500)
    _click_por_texto(page, "\u00d7", timeout=1500)
    time.sleep(0.8)

    # Descomprimir en la carpeta destino
    extract_dir = destino_dir / sufijo
    extract_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(temp_zip, 'r') as zf:
        zf.extractall(extract_dir)
    log(job, f"ZIP descomprimido en: {extract_dir}")
    # Eliminar el ZIP temporal
    temp_zip.unlink()
    return extract_dir

def _agregar_carta_sin_tipo(job, page, valor, etiqueta=""):
    pref = f"[{etiqueta}] " if etiqueta else ""
    log(job, f"    {pref}Agregando {valor} al buzón...")
    _escribir_valor_y_consultar(page, valor)
    _click_agregar(job, page)
    log(job, f"    {pref}Carta {valor} agregada (contador +1).", "success")
    return detectar_tipo_solicitud(valor)

# ==================== AUTOMATIZACIÓN PRINCIPAL ====================
def run_automation(job, usuario, password, tipo_acceso, lote, valores, download_path, login_timeout):
    from playwright.sync_api import sync_playwright

    dl_dir = Path(download_path)
    dl_dir.mkdir(parents=True, exist_ok=True)
    nit_detectado, ips_nombre = resolver_ips_por_usuario(usuario)
    if nit_detectado:
        log(job, f"NIT detectado en el usuario: {nit_detectado} -> IPS: {ips_nombre}")
    else:
        log(job, "No se detecto NIT en el usuario. Carpeta: IPS_DESCONOCIDA", "warn")
    job["dl_dir"] = dl_dir
    job["lote"] = lote
    job["ips_nombre"] = ips_nombre
    ips_dir = dl_dir / ips_nombre
    ips_dir.mkdir(parents=True, exist_ok=True)

    # Estructura final: dentro de ips_dir tendremos DEV/, LIQ/ y RAD/ (carpetas) y luego el Excel
    dev_extract_dir = ips_dir / "DEV"
    liq_extract_dir = ips_dir / "LIQ"
    radicado_extract_dir = ips_dir / "RAD"
    dev_extract_dir.mkdir(exist_ok=True)
    liq_extract_dir.mkdir(exist_ok=True)
    radicado_extract_dir.mkdir(exist_ok=True)

    zip_ya_generado = False

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox","--disable-dev-shm-usage"])
            context = browser.new_context(accept_downloads=True, viewport={"width": 1500, "height": 900})
            page = context.new_page()
            job["browser"] = browser
            job["context"] = context

            if not _hacer_login(job, page, usuario, password, login_timeout):
                if job["state"].get("stopping"):
                    return
            if job["state"].get("stopping"):
                return

            log(job, "Navegando directamente a Consulta de Cartas...")
            page.goto(CONSULTA_CARTAS_URL, wait_until="domcontentloaded", timeout=60000)
            time.sleep(5)  # dar tiempo a Angular para renderizar
            # Detectar cualquier elemento característico de esa pantalla
            if not _esperar_texto(job, page, r"(?i)seleccione tipo|BUSQUEDA|Busqueda|Consultar|CONSULTAR", timeout=30):
                # Último recurso: si la URL es correcta, continuar igual
                if "settlement-letters" not in (page.url or ""):
                    raise Exception("No se detectó la pantalla de Consulta de Cartas después de navegación directa.")
                log(job, "URL correcta, continuando aunque el texto no se detectó.", "warn")
            else:
                log(job, "Pantalla 'Consulta de Cartas' cargada correctamente.")

            completadas = cargar_progreso(job, ips_dir)
            pendientes = []
            for v in valores:
                if v in completadas:
                    log(job, f"Omitiendo (ya procesada): {v}")
                    with job["lock"]:
                        job["state"]["stats"]["descargadas"] += 1
                else:
                    pendientes.append(v)
            with job["lock"]:
                job["state"]["stats"]["total"] = len(valores)

            lista_dev = [v for v in pendientes if detectar_tipo_solicitud(v) == "No Dev/Obj"]
            lista_liq = [v for v in pendientes if detectar_tipo_solicitud(v) == "No Liquidación"]
            lista_radicado = [v for v in pendientes if detectar_tipo_solicitud(v) == "No Radicado"]

            log(job, f"Total a procesar: {len(pendientes)} | DEV/Obj: {len(lista_dev)} | LIQ: {len(lista_liq)} | Radicado: {len(lista_radicado)}")

            # Procesar DEV (agregar cartas y luego descargar ZIP y extraer)
            if lista_dev:
                log(job, f"=== Procesando {len(lista_dev)} cartas DEV/Obj ===")
                log(job, "Seleccionando tipo 'No Dev/Obj' (solo una vez)...")
                _seleccionar_tipo_solicitud(job, page, "No Dev/Obj")
                # Registrar cartas para el Excel posterior
                cartas_dev_registradas = []
                exitos_dev = 0
                for idx, valor in enumerate(lista_dev, 1):
                    if job["state"].get("stopping"):
                        break
                    log(job, f"[DEV {idx}/{len(lista_dev)}]")
                    try:
                        _agregar_carta_sin_tipo(job, page, valor, etiqueta="DEV")
                        exitos_dev += 1
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        with job["lock"]:
                            job["state"]["descargas_exitosas"].append({
                                "consecutivo": valor,
                                "tipo": "No Dev/Obj",
                                "archivo": "",  # se llenará después con la ruta dentro del ZIP final
                                "timestamp": timestamp,
                            })
                            job["state"]["stats"]["descargadas"] += 1
                        completadas.add(valor)
                        guardar_progreso(
                            job, ips_dir, completadas,
                            nuevo_item={"factura": valor, "tipo": "No Dev/Obj"},
                            meta={
                                "identidad": job["state"].get("identidad"),
                                "ips": job["state"].get("ips_identity"),
                                "ejecucion_id": job["state"].get("ejecucion_id"),
                                "ips_nombre": ips_nombre,
                                "aseguradora": "Mundial",
                                "periodo": job["state"].get("lote"),
                            },
                        )
                        cartas_dev_registradas.append({
                            "consecutivo": valor,
                            "tipo": "No Dev/Obj",
                            "timestamp": timestamp
                        })
                    except Exception as e:
                        log(job, f"  Error al agregar {valor}: {e}", "error")
                        with job["lock"]:
                            job["state"]["errores_detalle"].append({
                                "consecutivo": valor,
                                "tipo": "No Dev/Obj",
                                "error": str(e),
                                "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
                                "captura": "",
                            })
                            job["state"]["stats"]["errores"] += 1
                        try:
                            errores_dir = ips_dir / "Errores"
                            errores_dir.mkdir(parents=True, exist_ok=True)
                            cap = errores_dir / f"ERROR_{re.sub(r'[^A-Za-z0-9_-]','_', valor)}.png"
                            page.screenshot(path=str(cap))
                            with job["lock"]:
                                job["state"]["errores_detalle"][-1]["captura"] = str(cap)
                        except Exception:
                            pass

                if not job["state"].get("stopping") and exitos_dev > 0:
                    log(job, f"Descargando y extrayendo ZIP de DEV ({exitos_dev} carta(s) en buzón)...")
                    try:
                        _descargar_zip_y_extraer(job, page, context, ips_dir, "DEV")
                        # Las cartas ahora están en ips_dir/DEV/
                        # Actualizar las rutas en job_state para el Excel final
                        with job["lock"]:
                            for ex in job["state"]["descargas_exitosas"]:
                                if ex["tipo"] == "No Dev/Obj" and not ex["archivo"]:
                                    # Asignar una ruta simbólica dentro del ZIP final
                                    ex["archivo"] = f"DEV/{ex['consecutivo']}.pdf"  # aproximado, solo para referencia
                    except Exception as e:
                        log(job, f"Error al procesar ZIP de DEV: {e}", "error")

            # Procesar LIQ
            if lista_liq and not job["state"].get("stopping"):
                log(job, f"=== Procesando {len(lista_liq)} cartas LIQ ===")
                # Recargar la página para resetear el buzón
                page.goto(CONSULTA_CARTAS_URL, wait_until="domcontentloaded", timeout=60000)
                time.sleep(2)
                if not _esperar_texto(job, page, r"Seleccione tipo de solicitud", timeout=10):
                    raise Exception("No se detectó la pantalla de Consulta de Cartas para LIQ.")
                log(job, "Seleccionando tipo 'No Liquidación' (solo una vez)...")
                _seleccionar_tipo_solicitud(job, page, "No Liquidación")
                cartas_liq_registradas = []
                exitos_liq = 0
                for idx, valor in enumerate(lista_liq, 1):
                    if job["state"].get("stopping"):
                        break
                    log(job, f"[LIQ {idx}/{len(lista_liq)}]")
                    try:
                        _agregar_carta_sin_tipo(job, page, valor, etiqueta="LIQ")
                        exitos_liq += 1
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        with job["lock"]:
                            job["state"]["descargas_exitosas"].append({
                                "consecutivo": valor,
                                "tipo": "No Liquidación",
                                "archivo": "",
                                "timestamp": timestamp,
                            })
                            job["state"]["stats"]["descargadas"] += 1
                        completadas.add(valor)
                        guardar_progreso(
                            job, ips_dir, completadas,
                            nuevo_item={"factura": valor, "tipo": "No Liquidación"},
                            meta={
                                "identidad": job["state"].get("identidad"),
                                "ips": job["state"].get("ips_identity"),
                                "ejecucion_id": job["state"].get("ejecucion_id"),
                                "ips_nombre": ips_nombre,
                                "aseguradora": "Mundial",
                                "periodo": job["state"].get("lote"),
                            },
                        )
                        cartas_liq_registradas.append({
                            "consecutivo": valor,
                            "tipo": "No Liquidación",
                            "timestamp": timestamp
                        })
                    except Exception as e:
                        log(job, f"  Error al agregar {valor}: {e}", "error")
                        with job["lock"]:
                            job["state"]["errores_detalle"].append({
                                "consecutivo": valor,
                                "tipo": "No Liquidación",
                                "error": str(e),
                                "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
                                "captura": "",
                            })
                            job["state"]["stats"]["errores"] += 1
                        try:
                            errores_dir = ips_dir / "Errores"
                            errores_dir.mkdir(parents=True, exist_ok=True)
                            cap = errores_dir / f"ERROR_{re.sub(r'[^A-Za-z0-9_-]','_', valor)}.png"
                            page.screenshot(path=str(cap))
                            with job["lock"]:
                                job["state"]["errores_detalle"][-1]["captura"] = str(cap)
                        except Exception:
                            pass

                if not job["state"].get("stopping") and exitos_liq > 0:
                    log(job, f"Descargando y extrayendo ZIP de LIQ ({exitos_liq} carta(s) en buzón)...")
                    try:
                        _descargar_zip_y_extraer(job, page, context, ips_dir, "LIQ")
                        with job["lock"]:
                            for ex in job["state"]["descargas_exitosas"]:
                                if ex["tipo"] == "No Liquidación" and not ex["archivo"]:
                                    ex["archivo"] = f"LIQ/{ex['consecutivo']}.pdf"
                    except Exception as e:
                        log(job, f"Error al procesar ZIP de LIQ: {e}", "error")

            # Procesar Radicado (CMV/CMVIQ y cualquier otro prefijo distinto de DEV/OBJ/LIQ)
            if lista_radicado and not job["state"].get("stopping"):
                log(job, f"=== Procesando {len(lista_radicado)} cartas Radicado ===")
                # Recargar la página para resetear el buzón
                page.goto(CONSULTA_CARTAS_URL, wait_until="domcontentloaded", timeout=60000)
                time.sleep(2)
                if not _esperar_texto(job, page, r"Seleccione tipo de solicitud", timeout=10):
                    raise Exception("No se detectó la pantalla de Consulta de Cartas para Radicado.")
                log(job, "Seleccionando tipo 'No Radicado' (solo una vez)...")
                _seleccionar_tipo_solicitud(job, page, "No Radicado")
                cartas_radicado_registradas = []
                exitos_radicado = 0
                for idx, valor in enumerate(lista_radicado, 1):
                    if job["state"].get("stopping"):
                        break
                    log(job, f"[Radicado {idx}/{len(lista_radicado)}]")
                    try:
                        _agregar_carta_sin_tipo(job, page, valor, etiqueta="RAD")
                        exitos_radicado += 1
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        with job["lock"]:
                            job["state"]["descargas_exitosas"].append({
                                "consecutivo": valor,
                                "tipo": "No Radicado",
                                "archivo": "",
                                "timestamp": timestamp,
                            })
                            job["state"]["stats"]["descargadas"] += 1
                        completadas.add(valor)
                        guardar_progreso(
                            job, ips_dir, completadas,
                            nuevo_item={"factura": valor, "tipo": "No Radicado"},
                            meta={
                                "identidad": job["state"].get("identidad"),
                                "ips": job["state"].get("ips_identity"),
                                "ejecucion_id": job["state"].get("ejecucion_id"),
                                "ips_nombre": ips_nombre,
                                "aseguradora": "Mundial",
                                "periodo": job["state"].get("lote"),
                            },
                        )
                        cartas_radicado_registradas.append({
                            "consecutivo": valor,
                            "tipo": "No Radicado",
                            "timestamp": timestamp
                        })
                    except Exception as e:
                        log(job, f"  Error al agregar {valor}: {e}", "error")
                        with job["lock"]:
                            job["state"]["errores_detalle"].append({
                                "consecutivo": valor,
                                "tipo": "No Radicado",
                                "error": str(e),
                                "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
                                "captura": "",
                            })
                            job["state"]["stats"]["errores"] += 1
                        try:
                            errores_dir = ips_dir / "Errores"
                            errores_dir.mkdir(parents=True, exist_ok=True)
                            cap = errores_dir / f"ERROR_{re.sub(r'[^A-Za-z0-9_-]','_', valor)}.png"
                            page.screenshot(path=str(cap))
                            with job["lock"]:
                                job["state"]["errores_detalle"][-1]["captura"] = str(cap)
                        except Exception:
                            pass

                if not job["state"].get("stopping") and exitos_radicado > 0:
                    log(job, f"Descargando y extrayendo ZIP de Radicado ({exitos_radicado} carta(s) en buzón)...")
                    try:
                        _descargar_zip_y_extraer(job, page, context, ips_dir, "RAD")
                        with job["lock"]:
                            for ex in job["state"]["descargas_exitosas"]:
                                if ex["tipo"] == "No Radicado" and not ex["archivo"]:
                                    ex["archivo"] = f"RAD/{ex['consecutivo']}.pdf"
                    except Exception as e:
                        log(job, f"Error al procesar ZIP de Radicado: {e}", "error")

            if job["state"].get("stopping"):
                browser.close()
                return

            browser.close()

            # Ahora crear el Excel general con todas las cartas
            with job["lock"]:
                exitosas = job["state"]["descargas_exitosas"].copy()
                errores = job["state"]["errores_detalle"].copy()
            if EXCEL_AVAILABLE:
                excel_path = ips_dir / "reporte_cartas.xlsx"
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Cartas procesadas"
                ws.append(["Consecutivo", "Tipo solicitud", "Carpeta", "Fecha/Hora"])
                for ex in exitosas:
                    carpeta = {"No Dev/Obj": "DEV", "No Liquidación": "LIQ", "No Radicado": "RAD"}.get(ex["tipo"], "DEV")
                    ws.append([ex.get("consecutivo"), ex.get("tipo"), carpeta, ex.get("timestamp")])
                if errores:
                    wse = wb.create_sheet("Errores")
                    wse.append(["Consecutivo", "Tipo solicitud", "Error", "Captura", "Fecha/Hora"])
                    for er in errores:
                        wse.append([er.get("consecutivo"), er.get("tipo"), er.get("error"), er.get("captura"), er.get("timestamp")])
                for col in ws.columns:
                    max_length = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    ws.column_dimensions[col_letter].width = adjusted_width
                wb.save(excel_path)
                log(job, f"Reporte Excel generado: {excel_path.name}")
            else:
                excel_path = None

            # Excel aparte, solo de errores, para descarga automática desde
            # el frontend — nunca debe poder tumbar el proceso ya exitoso.
            try:
                if errores and EXCEL_AVAILABLE:
                    nombre_err = f"errores_{ips_nombre}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    path_err = dl_dir / nombre_err
                    wb_err = openpyxl.Workbook()
                    ws_err = wb_err.active
                    ws_err.title = "Errores"
                    ws_err.append(["Consecutivo", "Tipo solicitud", "Error", "Fecha/Hora"])
                    for er in errores:
                        ws_err.append([er.get("consecutivo"), er.get("tipo"), er.get("error"), er.get("timestamp")])
                    wb_err.save(path_err)
                    try:
                        lote_rel = str(dl_dir.relative_to(DOWNLOAD_DIR))
                    except Exception:
                        lote_rel = dl_dir.name
                    with job["lock"]:
                        job["state"]["errores_excel_url"] = f"/mundial/downloads/{lote_rel}/{nombre_err}"
                else:
                    with job["lock"]:
                        job["state"]["errores_excel_url"] = None
            except Exception:
                pass

            # Crear el ZIP único final con nombre de la IPS
            final_zip_path = dl_dir / f"{ips_nombre}.zip"
            with zipfile.ZipFile(final_zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                # Agregar carpeta DEV
                if dev_extract_dir.exists():
                    for file_path in dev_extract_dir.rglob("*"):
                        if file_path.is_file():
                            arcname = f"DEV/{file_path.relative_to(dev_extract_dir)}"
                            zf.write(file_path, arcname)
                # Agregar carpeta LIQ
                if liq_extract_dir.exists():
                    for file_path in liq_extract_dir.rglob("*"):
                        if file_path.is_file():
                            arcname = f"LIQ/{file_path.relative_to(liq_extract_dir)}"
                            zf.write(file_path, arcname)
                # Agregar carpeta Radicado
                if radicado_extract_dir.exists():
                    for file_path in radicado_extract_dir.rglob("*"):
                        if file_path.is_file():
                            arcname = f"RAD/{file_path.relative_to(radicado_extract_dir)}"
                            zf.write(file_path, arcname)
                # Agregar Excel
                if excel_path and excel_path.exists():
                    zf.write(excel_path, "reporte_cartas.xlsx")
                # Agregar carpeta de errores si existe
                errores_dir = ips_dir / "Errores"
                if errores_dir.exists():
                    for file_path in errores_dir.rglob("*"):
                        if file_path.is_file():
                            arcname = f"Errores/{file_path.relative_to(errores_dir)}"
                            zf.write(file_path, arcname)
                # Excel dedicado SOLO con consecutivos en error persistente
                # (recargable directo al bot, sin ambigüedad de hoja activa).
                if errores and EXCEL_AVAILABLE:
                    try:
                        nombre_solo_errores = f"errores_{ips_nombre}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        path_solo_errores = dl_dir / nombre_solo_errores
                        wb_solo = openpyxl.Workbook()
                        ws_solo = wb_solo.active
                        ws_solo.title = "Errores"
                        ws_solo.append(["Consecutivo", "Tipo solicitud", "Error", "Fecha/Hora"])
                        for er in errores:
                            ws_solo.append([er.get("consecutivo"), er.get("tipo"), er.get("error"), er.get("timestamp")])
                        wb_solo.save(path_solo_errores)
                        zf.write(path_solo_errores, nombre_solo_errores)
                    except Exception as e:
                        log(job, f"⚠️ No se pudo generar el Excel dedicado de errores: {e}", "warn")
            log(job, f"ZIP final creado: {final_zip_path.name}", "success")
            try:
                zip_rel = str(final_zip_path.relative_to(DOWNLOAD_DIR))
                with job["lock"]:
                    job["state"]["zip_url"] = f"/mundial/downloads/{zip_rel}"
            except Exception:
                pass
            zip_ya_generado = True
            # Ya terminó bien: cualquier ZIP parcial que haya quedado de
            # intentos/ciclos anteriores queda obsoleto, se borra.
            for viejo in dl_dir.glob(f"{ips_nombre}_PARCIAL_*.zip"):
                try:
                    viejo.unlink()
                except Exception:
                    pass
            # NO borrar las carpetas DEV/LIQ si hay un reintento activo:
            # los reintentos acumulan PDFs sobre estas carpetas. Solo se limpian
            # cuando ya no quedan errores pendientes (proceso realmente terminado).
            hay_errores_pendientes = len(errores) > 0 and job["state"].get("_reintento_activo")
            if not hay_errores_pendientes:
                shutil.rmtree(dev_extract_dir, ignore_errors=True)
                shutil.rmtree(liq_extract_dir, ignore_errors=True)
                shutil.rmtree(radicado_extract_dir, ignore_errors=True)
            if excel_path and excel_path.exists():
                excel_path.unlink()
            # El progreso.json lo dejamos

            total_errores = len(errores)
            if total_errores:
                log(job, f"Proceso completado con {total_errores} error(es). Ver Excel en el ZIP.", "warn")
            else:
                log(job, "Proceso completado sin errores.", "success")

    except Exception as e:
        if not job["state"].get("stopping"):
            log(job, f"Error critico: {e}", "error")
            with job["lock"]:
                job["state"]["error"] = str(e)
        else:
            log(job, "Proceso detenido por el usuario.")
    finally:
        # Red de seguridad final: sin importar CÓMO se salió de la función
        # (éxito, error, reintentos agotados, o Detener manual del usuario),
        # si todavía no se generó ningún ZIP, se genera uno parcial aquí con
        # lo que se haya descargado hasta el momento.
        if not zip_ya_generado:
            generar_zip_parcial(job)
        historial_db.cerrar_ejecucion(
            job["state"].get("ejecucion_id"), "error" if job["state"].get("error") else ("cancelada" if job["state"].get("stopping") else "completada"),
            total_detectadas=job["state"].get("stats", {}).get("total", 0),
            total_procesadas=job["state"].get("stats", {}).get("descargadas", 0) + job["state"].get("stats", {}).get("errores", 0),
            total_exitosas=job["state"].get("stats", {}).get("descargadas", 0),
            total_fallidas=job["state"].get("stats", {}).get("errores", 0),
        )
        # Solo marcar como terminado si no hay un wrapper de reintentos controlando el estado
        with job["lock"]:
            if not job["state"].get("_reintento_activo"):
                job["state"]["running"] = False
                job["state"]["finished"] = True
                job["state"]["stopping"] = False
            # Si SÍ hay un wrapper de reintentos activo, "stopping" se deja
            # tal cual está -- es el wrapper quien debe verla para decidir si
            # detiene el ciclo de reintentos o si arranca uno nuevo. Antes se
            # reseteaba aquí incondicionalmente, y por eso "Detener" durante
            # un reintento automático no evitaba que arrancara el siguiente
            # ciclo completo (login + navegación) de todos modos.
        job["browser"] = None
        job["context"] = None
        job["dl_dir"] = None
        job["lote"] = None
        job["ips_nombre"] = None

# ==================== PARSEO DE VALORES ====================
def parse_valores(texto):
    if not texto:
        return []
    crudos = re.split(r"[\s,;]+", texto.strip())
    out = [c.strip() for c in crudos if c.strip()]
    vistos = set()
    res = []
    for c in out:
        if c not in vistos:
            vistos.add(c)
            res.append(c)
    return res

# ==================== REINTENTOS AUTOMÁTICOS ====================
MAX_CICLOS_REINTENTO = 5  # máximo de veces que el bot se relanza solo

def run_automation_con_reintentos(job, usuario, password, tipo_acceso, lote, valores, dl_path, login_timeout):
    """Envuelve run_automation con lógica de reintento automático.
    Si al terminar quedan errores, vuelve a lanzar el bot SOLO con las
    cartas que fallaron, hasta MAX_CICLOS_REINTENTO veces en total.
    Las credenciales y el lote se conservan entre ciclos."""

    ciclo = 1
    valores_actuales = list(valores)

    with job["lock"]:
        job["state"]["_reintento_activo"] = True

    while ciclo <= MAX_CICLOS_REINTENTO:
        if job["state"].get("stopping"):
            break

        if ciclo > 1:
            log(job, f"🔄 Reintento automático {ciclo}/{MAX_CICLOS_REINTENTO} — {len(valores_actuales)} carta(s) con error...", "warn")
            time.sleep(4)  # pausa breve antes de relanzar el navegador
            # Resetear stats para el nuevo ciclo (sin borrar el progreso en disco)
            with job["lock"]:
                job["state"]["running"] = True
                job["state"]["finished"] = False
                job["state"]["error"] = None
                job["state"]["stats"]["errores"] = 0
                job["state"]["errores_detalle"] = []

        # Ejecutar el bot con la lista actual
        run_automation(job, usuario, password, tipo_acceso, lote, valores_actuales, dl_path, login_timeout)

        if job["state"].get("stopping"):
            break

        # Revisar qué quedó con error
        with job["lock"]:
            errores = [e["consecutivo"] for e in job["state"].get("errores_detalle", [])]

        if not errores:
            # Sin errores — terminó limpio
            if ciclo > 1:
                log(job, f"✅ Todos los errores resueltos en el ciclo {ciclo}.", "success")
            break

        ciclo += 1
        if ciclo > MAX_CICLOS_REINTENTO:
            log(job, f"⚠️ Se alcanzó el máximo de {MAX_CICLOS_REINTENTO} intentos. "
                f"{len(errores)} carta(s) persisten con error y quedan en el Excel.", "warn")
            break

        # Preparar la siguiente ronda solo con los que fallaron
        valores_actuales = errores

    # Marcar como terminado y desactivar bandera de reintento
    with job["lock"]:
        job["state"]["_reintento_activo"] = False
        job["state"]["running"] = False
        job["state"]["finished"] = True
        job["state"]["stopping"] = False
    concurrency.registrar_fin("mundial")

# ==================== RUTAS FLASK ====================
@bp.route("/")
def index():
    return render_template("mundial_index.html")

@bp.route("/api/start", methods=["POST"])
def start_job():
    data = request.get_json(silent=True) or request.form or {}
    usuario = data.get("usuario", "").strip()
    password = data.get("password", "").strip()
    tipo_acceso = data.get("tipo_acceso", "").strip()
    lote = data.get("lote", "").strip() or datetime.now().strftime("Lote_%Y%m%d_%H%M")
    valores_texto = data.get("valores", "").strip()
    custom_path = data.get("download_path", "").strip()
    login_timeout = int(data.get("login_timeout", 180) or 180)
    identidad = validar_identidad(data.get("identidad"))
    if not identidad:
        return jsonify({"ok": False, "error": "Selecciona quién eres antes de iniciar el proceso.", "campo": "identidad"}), 400

    if not all([usuario, password]):
        return jsonify({"ok": False, "error": "Faltan usuario o contraseña"}), 400

    valores = parse_valores(valores_texto)
    if not valores:
        return jsonify({"ok": False, "error": "No hay consecutivos para procesar. Carga el Excel o pega la lista."}), 400

    lote_safe = re.sub(r"[^\w\-]", "_", lote)

    empresa_id = resolve_empresa_id(usuario)
    job = get_or_create_job(empresa_id)

    confirmar_duplicados = str(data.get("confirmar_duplicados", "")).lower() in ("true", "1", "on", "si", "sí")
    decision_redescarga = data.get("decision_redescarga", "")
    facturas_redescarga = {str(v) for v in (data.get("facturas_redescarga") or [])}
    ips_nit_manual = normalizar_nit(data.get("ips_nit_manual", ""))
    _nit_check, ips_check = resolver_ips_por_usuario(usuario)
    if ips_nit_manual:
        _nit_check = ips_nit_manual
        ips_check = resolver(nit=ips_nit_manual).get("nombre_estandar", ips_check)
    if resolver(nit=_nit_check, nombre_detectado=ips_check).get("metodo") == "NO_IDENTIFICADA":
        return jsonify({"ok": False, "requiere_seleccion_ips": True,
                         "catalogo_ips": listar_catalogo_por_responsable()})
    try:
        if _nit_check:
            ya_descargadas = historial_db.buscar_ya_descargadas_por_nit("Mundial", _nit_check, valores)
        else:
            ya_descargadas = historial_db.buscar_ya_descargadas("Mundial", ips_check, valores) if ips_check else {}
    except Exception as e:
        log(job, f"⚠️ No se pudo consultar el historial de duplicados: {e}", "warn")
        ya_descargadas = {}
    if ya_descargadas:
        if not confirmar_duplicados:
            return jsonify({"ok": False, "requiere_confirmacion": True, "ya_descargadas": ya_descargadas, "total_filas": len(valores)})
        if decision_redescarga == "ninguna":
            valores = [v for v in valores if v not in ya_descargadas]
        elif decision_redescarga == "seleccionadas":
            valores = [v for v in valores if v not in ya_descargadas or v in facturas_redescarga]

    with jobs_registry_lock:
        with job["lock"]:
            if job["state"]["running"]:
                return jsonify({"ok": False, "error": "Ya hay un proceso en ejecución para esta empresa/IPS. Espera a que termine antes de iniciar otro."}), 409
            if count_running_jobs() >= MAX_CONCURRENT_JOBS:
                return jsonify({"ok": False, "error": f"Se alcanzó el máximo de {MAX_CONCURRENT_JOBS} procesos simultáneos. Intenta de nuevo en unos minutos."}), 429
            if not concurrency.puede_iniciar():
                return jsonify({"ok": False, "error": "El panel está al máximo de descargas simultáneas entre todas las aseguradoras en este momento. Intenta de nuevo en unos minutos."}), 429
            job["state"]["running"] = True
            job["state"]["finished"] = False
            job["state"]["error"] = None
            job["state"]["stats"] = {"total": len(valores), "descargadas": 0, "errores": 0}
            job["state"]["errores_detalle"] = []
            job["state"]["descargas_exitosas"] = []
            job["state"]["logs"] = []
            job["state"]["_reintento_activo"] = False
            job["state"]["errores_excel_url"] = None
            job["state"]["zip_url"] = None
            job["state"]["lote"] = lote

    dl_path = custom_path if custom_path else str(DOWNLOAD_DIR / lote_safe)
    if custom_path:
        registro_rutas.registrar_ruta("Mundial", custom_path)

    job["state"]["identidad"] = identidad
    nit_identidad, nombre_identidad = resolver_ips_por_usuario(usuario)
    job["state"]["ips_identity"] = resolver(nit=nit_identidad, nombre_detectado=nombre_identidad)
    job["state"]["ips_nit"] = nit_identidad
    job["state"]["ejecucion_id"] = historial_db.iniciar_ejecucion(
        identidad, "mundial", "Mundial", job["state"]["ips_identity"], lote_safe, dl_path
    )

    concurrency.registrar_inicio("mundial")
    t = threading.Thread(
        target=run_automation_con_reintentos,
        args=(job, usuario, password, tipo_acceso, lote_safe, valores, dl_path, login_timeout),
        daemon=True,
    )
    t.start()
    return jsonify({"ok": True, "download_path": dl_path, "total": len(valores), "lote": lote_safe, "empresa_id": empresa_id})

@bp.route("/api/stop", methods=["POST"])
def stop_job_route():
    empresa_id = resolve_empresa_id((request.get_json(silent=True) or {}).get("empresa_id") if request.is_json else request.args.get("empresa_id", ""))
    job = get_job_or_none(empresa_id)
    if not job:
        return jsonify({"ok": False, "message": "No hay proceso en ejecucion"}), 400
    with job["lock"]:
        if not job["state"]["running"]:
            return jsonify({"ok": False, "message": "No hay proceso en ejecucion"}), 400
    stop_job(job)
    return jsonify({"ok": True, "message": "Deteniendo proceso..."})

@bp.route("/api/reset", methods=["POST"])
def reset_job_route():
    data = request.get_json(silent=True) or request.form or {}
    lote = data.get("lote", "").strip()
    empresa_id = resolve_empresa_id(data.get("empresa_id", ""))
    job = get_job_or_none(empresa_id)
    if job:
        with job["lock"]:
            estaba_corriendo = job["state"]["running"]
        if estaba_corriendo:
            # Importante: NO retener job["lock"] aquí; stop_job() lo adquiere
            # internamente. threading.Lock no es reentrante -- si lo
            # retuviéramos, el reset quedaría bloqueado para siempre.
            stop_job(job)
            # Espera activa: hasta 10s para que el hilo confirme running=False,
            # en vez de un sleep(2) fijo que podía no ser suficiente.
            for _ in range(20):
                time.sleep(0.5)
                with job["lock"]:
                    if not job["state"]["running"]:
                        break
    # Si hay lote, borra solo el de ese lote. Si NO hay lote, borra TODOS los progreso.json.
    if lote:
        lote_safe = re.sub(r"[^\w\-]", "_", lote)
        base = DOWNLOAD_DIR / lote_safe
    else:
        base = DOWNLOAD_DIR
    borrados = 0
    if base.exists():
        for cand in list(base.glob("**/progreso.json")):
            try:
                n_migradas = historial_db.migrar_progreso_json(cand.parent)
                cand.unlink()
                borrados += 1
                log(job, f"Progreso eliminado: {cand} ({n_migradas} factura(s) archivadas en el historial)")
            except Exception as e:
                log(job, f"Error al borrar progreso: {e}", "warn")
    if borrados == 0:
        log(job, "No se encontro progreso para borrar.", "warn")
    if job:
        reset_state(job)
    return jsonify({"ok": True, "message": f"Progreso eliminado ({borrados} archivo(s)). Los soportes se conservaron."})

@bp.route("/api/status")
def get_status():
    empresa_id_raw = request.args.get("empresa_id", "")
    if empresa_id_raw:
        empresa_id = resolve_empresa_id(empresa_id_raw)
        job = get_job_or_none(empresa_id)
        if not job:
            return jsonify({"running": False, "finished": False, "error": None,
                             "stats": {"total": 0, "descargadas": 0, "errores": 0}})
        with job["lock"]:
            return jsonify({
                "running": job["state"]["running"],
                "finished": job["state"]["finished"],
                "error": job["state"]["error"],
                "stats": job["state"]["stats"],
                "errores_excel_url": job["state"].get("errores_excel_url"),
                "zip_url": job["state"].get("zip_url"),
            })
    with jobs_registry_lock:
        activos = [j for j in jobs.values() if j["state"]["running"]]
        stats = {
            "total": sum(j["state"]["stats"]["total"] for j in activos),
            "descargadas": sum(j["state"]["stats"]["descargadas"] for j in activos),
            "errores": sum(j["state"]["stats"]["errores"] for j in activos),
        }
        procesos = [{
            "ips": j.get("ips_nombre") or "Detectando IPS...",
            "stats": j["state"]["stats"],
        } for j in activos]
        return jsonify({"running": len(activos) > 0, "finished": False, "error": None, "stats": stats, "procesos": procesos})

@bp.route("/api/logs")
def get_logs():
    since = int(request.args.get("since", 0))
    empresa_id = resolve_empresa_id(request.args.get("empresa_id", ""))
    job = get_job_or_none(empresa_id)
    if not job:
        return jsonify({"logs": []})
    with job["lock"]:
        return jsonify({"logs": job["state"]["logs"][since:]})

@bp.route("/api/logs", methods=["DELETE"])
def clear_logs():
    empresa_id = resolve_empresa_id(request.args.get("empresa_id", ""))
    job = get_job_or_none(empresa_id)
    if job:
        with job["lock"]:
            job["state"]["logs"] = []
    return jsonify({"ok": True})

@bp.route("/api/files")
def list_files():
    lote = request.args.get("lote", "")
    folder = DOWNLOAD_DIR / lote if lote else DOWNLOAD_DIR
    files = []
    if folder.exists():
        # Buscar el ZIP final con nombre de IPS (no los ZIPs temporales)
        for zip_file in folder.rglob("*.zip"):
            # Excluir cualquier ZIP que no sea el final (por ejemplo, los que contengan "temp" o "PARCIAL")
            if "temp_" in zip_file.name or "PARCIAL" in zip_file.name:
                continue
            # También evitar el ZIP completo de lote si se generara (pero ya no)
            files.append({
                "name": zip_file.name,
                "size": zip_file.stat().st_size,
                "lote": lote,
                "path": zip_file.name
            })
    return jsonify({"files": files})

@bp.route("/api/files", methods=["DELETE"])
def delete_all_files():
    lote = request.args.get("lote", "")
    folder = DOWNLOAD_DIR / lote if lote else DOWNLOAD_DIR
    if not folder.exists():
        return jsonify({"ok": True, "message": "No hay archivos que eliminar"})
    try:
        eliminados = 0
        for item in list(folder.iterdir()):
            if item.is_file():
                item.unlink()
                eliminados += 1
            elif item.is_dir():
                for sub in item.rglob("*"):
                    if sub.is_file() and sub.name != "progreso.json":
                        sub.unlink()
                        eliminados += 1
        log(None, f"Soportes eliminados: {eliminados} (progreso conservado)")
        return jsonify({"ok": True, "message": f"Se eliminaron {eliminados} archivo(s). El progreso se conservo.", "eliminados": eliminados})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@bp.route("/api/files/soportes", methods=["DELETE"])
def delete_soportes():
    """Borra ZIPs, PDFs y carpetas de soportes pero conserva progreso.json intacto."""
    lote = request.args.get("lote", "")
    folder = DOWNLOAD_DIR / lote if lote else DOWNLOAD_DIR
    if not folder.exists():
        return jsonify({"ok": True, "message": "No hay soportes que eliminar", "eliminados": 0})
    CONSERVAR = {"progreso.json"}
    try:
        eliminados = 0
        for path in sorted(folder.rglob("*"), reverse=True):
            if path.name in CONSERVAR:
                continue
            if path.is_file():
                path.unlink()
                eliminados += 1
            elif path.is_dir():
                try:
                    path.rmdir()  # solo borra si quedó vacía
                except OSError:
                    pass
        log(None, f"Soportes eliminados: {eliminados} archivos. progreso.json conservado.")
        return jsonify({"ok": True, "message": f"{eliminados} soporte(s) eliminado(s). El progreso se conservó.", "eliminados": eliminados})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@bp.route("/api/progreso")
def get_progreso():
    """Devuelve el historial de progreso de todos los lotes guardados en disco."""
    lote = request.args.get("lote", "")
    resultados = []
    base = DOWNLOAD_DIR / lote if lote else DOWNLOAD_DIR
    for prog_file in sorted(base.rglob("progreso.json")):
        try:
            data = json.loads(prog_file.read_text(encoding="utf-8"))
            # Reconstruir ruta relativa para mostrar: lote/IPS
            partes = prog_file.parent.relative_to(DOWNLOAD_DIR).parts
            resultados.append({
                "ruta": "/".join(partes),
                "lote": partes[0] if partes else "",
                "ips": partes[1] if len(partes) > 1 else "",
                "completadas": sorted(data) if isinstance(data, list) else sorted(data.get("completadas", data) if isinstance(data, dict) else []),
                "total": len(data) if isinstance(data, (list, set)) else len(data.get("completadas", data) if isinstance(data, dict) else []),
            })
        except Exception:
            pass
    return jsonify({"progreso": resultados})


@bp.route("/api/cruce", methods=["POST"])
def api_cruce():
    """Cruza los 3 Excel y devuelve los consecutivos que cumplen ambas condiciones:
    1. La factura NO tiene fecha de glosa en Cartera (o no existe en Cartera).
    2. La factura NO aparece en Esculapio.
    """
    try:
        import pandas as pd
        from io import BytesIO

        f_mundial   = request.files.get("mundial")
        f_cartera   = request.files.get("cartera")
        f_esculapio = request.files.get("esculapio")

        if not all([f_mundial, f_cartera, f_esculapio]):
            return jsonify({"ok": False, "error": "Faltan archivos. Se requieren los 3 Excel."}), 400

        # ── Leer Excel Mundial (header en fila 3, índice 3) ──
        df_mundial = pd.read_excel(BytesIO(f_mundial.read()), header=3)
        # Columna A = Numero Factura, Columna M (índice 12) = Consecutivo
        col_factura_m = df_mundial.columns[0]
        col_consec    = df_mundial.columns[12]
        df_mundial = df_mundial[[col_factura_m, col_consec]].dropna(subset=[col_consec])
        df_mundial[col_factura_m] = df_mundial[col_factura_m].astype(str).str.strip()
        df_mundial[col_consec]    = df_mundial[col_consec].astype(str).str.strip()
        # Filtrar solo filas que parezcan consecutivos reales: prefijo de letras + guion + número
        # (antes solo aceptaba DEV/LIQ/CMV y descartaba silenciosamente OBJ y otros prefijos)
        df_mundial = df_mundial[df_mundial[col_consec].str.match(r"^[A-Z]+-\d+")]

        # ── Leer Excel Cartera (header en fila 0 de datos, skiprows=1) ──
        df_cartera = pd.read_excel(BytesIO(f_cartera.read()), skiprows=1, header=0)
        # Col A = No. FACTURA, Col P (índice 15) = FECHA DE GLOSA
        col_fac_c   = df_cartera.columns[0]
        col_glosa   = df_cartera.columns[15]
        df_cartera[col_fac_c] = df_cartera[col_fac_c].astype(str).str.strip()
        # Construir set de facturas CON fecha de glosa (estas se excluyen)
        con_glosa = set(
            df_cartera[df_cartera[col_glosa].notna()][col_fac_c].tolist()
        )
        todas_cartera = set(df_cartera[col_fac_c].tolist())

        # ── Leer Excel Esculapio (col B = nofactura) ──
        df_esc = pd.read_excel(BytesIO(f_esculapio.read()), header=0)
        col_nofac = df_esc.columns[1]  # columna B
        # Normalizar: quitar guion del prefijo "71-73582" → "7173582"
        esc_facturas = set(
            df_esc[col_nofac].astype(str).str.replace("-", "", regex=False).str.strip().tolist()
        )

        # ── Aplicar filtros ──
        consecutivos = []
        for _, row in df_mundial.iterrows():
            factura   = row[col_factura_m]
            consec    = row[col_consec]
            # Normalizar factura mundial también (por si acaso)
            fac_norm  = factura.replace("-", "").strip()

            # Condición 1: sin fecha de glosa en Cartera (o no existe en Cartera)
            cond1 = factura not in con_glosa

            # Condición 2: no aparece en Esculapio
            cond2 = fac_norm not in esc_facturas

            if cond1 and cond2:
                consecutivos.append(consec)

        # Eliminar duplicados conservando orden
        vistos = set()
        unicos = []
        for c in consecutivos:
            if c not in vistos:
                vistos.add(c)
                unicos.append(c)

        return jsonify({
            "ok": True,
            "consecutivos": unicos,
            "stats": {
                "mundial":    len(df_mundial),
                "cartera":    len(todas_cartera),
                "sin_glosa":  len(df_mundial) - sum(
                    1 for _, r in df_mundial.iterrows() if r[col_factura_m] in con_glosa
                ),
                "esculapio":  len(esc_facturas),
            }
        })

    except Exception as e:
        log(None, f"Error en cruce de Excel: {e}", "error")
        return jsonify({"ok": False, "error": str(e)}), 500

@bp.route("/downloads/<path:filename>")
def download_file(filename):
    for file_path in DOWNLOAD_DIR.rglob(filename):
        if file_path.is_file():
            return send_from_directory(file_path.parent, file_path.name, as_attachment=True)
    return jsonify({"error": "Archivo no encontrado"}), 404

@bp.route("/api/upload", methods=["POST"])
def upload_consecutivos():
    if 'file' not in request.files:
        return jsonify({"ok": False, "error": "No se envio ningun archivo"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"ok": False, "error": "Archivo vacio"}), 400
    try:
        filename = file.filename.lower()
        valores = []
        if filename.endswith('.csv'):
            raw = file.read()
            for enc in ('utf-8-sig', 'utf-8', 'latin-1', 'cp1252'):
                try:
                    txt = raw.decode(enc)
                    break
                except (UnicodeDecodeError, LookupError):
                    continue
            else:
                txt = raw.decode('latin-1', errors='replace')
            reader = csv.reader(txt.splitlines())
            rows = list(reader)
            header = rows[0] if rows else []
            col_idx = 0
            for i, h in enumerate(header):
                palabras = re.findall(r'[a-záéíóúñ]+', (h or '').lower())
                if any(p.startswith('consecutiv') or p == 'radicado' for p in palabras):
                    col_idx = i
                    break
            # Igual que en el flujo de Excel: decidir por FORMA (letras+guion+número),
            # no por adivinar la palabra exacta del título del encabezado.
            start = 0 if (header and re.match(r'^[A-Z]+-?\d', (header[col_idx] or '').upper())) else 1
            for r in rows[start:]:
                if len(r) > col_idx and r[col_idx].strip():
                    valores.append(r[col_idx].strip())
        elif filename.endswith(('.xls', '.xlsx')):
            if not EXCEL_AVAILABLE:
                return jsonify({"ok": False, "error": "openpyxl no instalado"}), 500
            wb = openpyxl.load_workbook(BytesIO(file.read()), data_only=True)
            ws = wb.active
            col_idx = 1
            for cell in ws[1]:
                palabras = re.findall(r'[a-záéíóúñ]+', str(cell.value or '').lower())
                if any(p.startswith('consecutiv') or p == 'radicado' for p in palabras):
                    col_idx = cell.column
                    break
            # Decidir si la fila 1 es encabezado (texto) o ya es un dato real,
            # mirando si TIENE FORMA de consecutivo (letras+guion+número) —
            # así no depende de adivinar la palabra exacta del título
            # (ej: un encabezado truncado como "Cons" antes se colaba como dato).
            first = ws.cell(row=1, column=col_idx).value
            start_row = 1 if (first and re.match(r'^[A-Z]+-?\d', str(first).upper())) else 2
            for row in ws.iter_rows(min_row=start_row, values_only=True):
                val = row[col_idx - 1] if len(row) >= col_idx else None
                if val and str(val).strip():
                    valores.append(str(val).strip())
        else:
            return jsonify({"ok": False, "error": "Formato no soportado. Use CSV o Excel"}), 400

        vistos = set()
        limpios = []
        for v in valores:
            v = v.strip()
            if v and v not in vistos:
                vistos.add(v)
                limpios.append(v)
        if not limpios:
            return jsonify({"ok": False, "error": "No se encontraron consecutivos validos en el archivo"}), 400
        resumen = {}
        for v in limpios:
            t = detectar_tipo_solicitud(v)
            resumen[t] = resumen.get(t, 0) + 1
        return jsonify({"ok": True, "count": len(limpios), "valores": limpios, "resumen": resumen})
    except Exception as e:
        return jsonify({"ok": False, "error": f"Error al procesar archivo: {str(e)}"}), 500

if __name__ == "__main__":
    print("\n" + "=" * 55)
    print("  Descargador de Cartas Glosa - Seguros Mundial")
    print("  Portal A3M / iqdigital (un solo ZIP por IPS con carpetas DEV/LIQ y Excel)")
    print("=" * 55)
    print(f"  Carpeta de descargas: {DOWNLOAD_DIR}")
    print(f"  Puerto: {port}")
    print("=" * 55 + "\n")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)