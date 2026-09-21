"""
Activa IT - Descargador automático de cartas glosa (Previsora SOAT)
Versión mejorada con:
- Detener/Reiniciar
- Carpetas por IPS (forzando nombre exacto del mapa)
- Reporte Excel (descargadas + errores)
- Búsqueda flexible (Envios_D / ActaDevolucion / Carta de Objeción) con filtro correcto
- Persistencia (reanudación automática)
- Importación opcional de lista de facturas (CSV/Excel)
- Generación de ZIP parcial al detener o ante error (incluye Excel parcial y Errores)
- ZIP final incluye Excel y carpeta Errores
- API para consultar progreso y exportar a Excel
- Soporte para períodos individuales y rangos masivos
- Descarga con nombres mejorados: Factura Y [tipo_soporte]
"""


import os
import re
import json
import csv
import time
import threading
import logging
import zipfile
from datetime import datetime
from pathlib import Path
from flask import Blueprint, render_template, request, jsonify, send_from_directory, send_file
from . import concurrency
from . import historial_db
from . import registro_rutas
from .catalogo_ips import resolver, validar_identidad
from io import BytesIO

# Para generar Excel
try:
    import openpyxl
    from openpyxl.styles import Font
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    print("⚠️ openpyxl no instalado. No se generará el archivo Excel.")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

bp = Blueprint("previsora", __name__, url_prefix="/previsora")

BASE_DIR = Path(__file__).resolve().parent.parent  # raíz del proyecto (un nivel arriba de bots/)
port = int(os.environ.get("PORT", 8080))
DOWNLOAD_DIR = BASE_DIR / "downloads" / "previsora"
DOWNLOAD_DIR.mkdir(exist_ok=True)

# ==================== MAPA DE IPS POR NIT ====================
MAPA_IPS = {
    # IPS existentes
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
    # Nuevas IPS agregadas
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

# ==================== REGISTRO DE JOBS POR EMPRESA/IPS ====================
# En vez de un único estado global compartido por todos los usuarios, cada
# empresa (resuelta a partir del NIT dentro del usuario) tiene su propio job
# aislado: su propio estado, su propio lock y su propio navegador. Esto
# permite que dos empresas distintas corran en paralelo, mientras que la
# MISMA empresa sigue bloqueada si ya tiene un proceso en curso.
MAX_CONCURRENT_JOBS = int(os.environ.get("MAX_CONCURRENT_JOBS", "2"))

jobs = {}  # empresa_id -> job dict
jobs_registry_lock = threading.RLock()


def resolve_empresa_id(usuario: str) -> str:
    """Resuelve el identificador de 'empresa' extrayendo el NIT embebido en el
    usuario (ej: PREV900600550 -> 900600550). Si no se encuentra, se usa el
    propio usuario como identificador (aislado)."""
    u = (usuario or "").strip().upper()
    m = re.search(r"(\d{9,12})", u)
    return m.group(1) if m else u


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
        "facturas_permitidas": [],
        "errores_excel_url": None,
        "duplicados_pendientes": None,
        "duplicate_event": threading.Event(),
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
                "periodo": None,
                "ips_nombre": None,
            }
        return jobs[empresa_id]


def get_job_or_none(empresa_id: str):
    with jobs_registry_lock:
        return jobs.get(empresa_id)


def count_running_jobs() -> int:
    with jobs_registry_lock:
        return sum(1 for j in jobs.values() if j["state"]["running"])

# ==================== UTILIDADES DE PERÍODOS ====================
MESES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

def validar_periodo(p):
    """Valida que un string sea un período válido (MMMYY)"""
    if not p or len(p) < 5:
        return False
    mes = p[:3]
    anio = p[3:]
    return mes in MESES and re.match(r'^\d{2}$', anio)

def generar_rango_periodos(inicio, fin):
    """Genera lista de períodos entre inicio y fin (ambos inclusive)"""
    if not validar_periodo(inicio) or not validar_periodo(fin):
        return []

    mes_inicio = MESES.index(inicio[:3])
    anio_inicio = int(inicio[3:])
    mes_fin = MESES.index(fin[:3])
    anio_fin = int(fin[3:])

    # Convertir a fecha comparable (año * 100 + mes)
    fecha_inicio = anio_inicio * 100 + mes_inicio
    fecha_fin = anio_fin * 100 + mes_fin

    if fecha_fin < fecha_inicio:
        return []

    periodos = []
    anio = anio_inicio
    mes = mes_inicio

    while True:
        anio_str = str(anio).zfill(2)
        periodos.append(MESES[mes] + anio_str)

        if anio == anio_fin and mes == mes_fin:
            break

        mes += 1
        if mes > 11:
            mes = 0
            anio += 1

    return periodos

def parse_periodo_input(periodo_input):
    """Parsea el input de período y retorna lista de períodos.
    Soporta:
    - Período único: May26
    - Rango: Dic25-May26
    """
    periodo_input = periodo_input.strip()
    if not periodo_input:
        return []

    # Detectar rango con "-"
    if '-' in periodo_input:
        parts = [p.strip() for p in periodo_input.split('-')]
        if len(parts) == 2:
            return generar_rango_periodos(parts[0], parts[1])
        return []

    # Período individual
    if validar_periodo(periodo_input):
        return [periodo_input]

    return []

# ==================== LOGGING ====================
def log(job, msg, level="info"):
    ts = datetime.now().strftime("%H:%M:%S")
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
                job["state"].get("ejecucion_id"), None, "previsora", "Previsora",
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
    log(job, "🛑 Solicitando detención del proceso...", "warn")
    if job["browser"]:
        try:
            job["browser"].close()
            log(job, "  → Navegador cerrado por solicitud de stop.")
        except Exception as e:
            log(job, f"  → Error al cerrar navegador: {e}", "error")
    # Generar ZIP parcial si hay archivos descargados y tenemos los datos necesarios
    generar_zip_parcial(job)

def generar_zip_parcial(job):
    """Genera un ZIP con los PDFs ya descargados hasta el momento,
       incluyendo un reporte Excel parcial y la carpeta Errores."""
    if not job["dl_dir"] or not job["periodo"] or not job["ips_nombre"]:
        return
    ips_dir = job["dl_dir"] / job["ips_nombre"]
    if not ips_dir.exists():
        return

    # Obtener los datos actuales de descargas y errores
    with job["lock"]:
        exitosas = job["state"]["descargas_exitosas"].copy()
        errores = job["state"]["errores_detalle"].copy()

    # Generar un Excel parcial (si hay datos o si openpyxl está disponible)
    excel_parcial_path = None
    if EXCEL_AVAILABLE:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_name = f"reporte_parcial_{timestamp}.xlsx"
            excel_parcial_path = ips_dir / excel_name
            wb = openpyxl.Workbook()
            ws_exit = wb.active
            ws_exit.title = "Descargadas"
            ws_exit.append(["N° Factura", "Estado", "IPS", "Archivo Descargado", "Fecha/Hora"])
            for ex in exitosas:
                ws_exit.append([ex.get("factura"), ex.get("estado"), job["ips_nombre"], ex.get("archivo"), ex.get("timestamp")])
            ws_err = wb.create_sheet("Errores")
            ws_err.append(["N° Factura", "Estado", "IPS", "Error", "Captura pantalla", "Fecha/Hora"])
            for err in errores:
                ws_err.append([err.get("factura"), err.get("estado"), job["ips_nombre"], err.get("error"), err.get("captura"), err.get("timestamp")])
            wb.save(excel_parcial_path)
            log(job, f"📊 Reporte Excel parcial generado: {excel_parcial_path}")
        except Exception as e:
            log(job, f"⚠️ No se pudo generar Excel parcial: {e}", "warn")
            excel_parcial_path = None

    # Recopilar archivos a incluir
    archivos_a_incluir = []
    # PDFs
    archivos_a_incluir.extend(ips_dir.rglob("*.pdf"))
    # Excel parcial (si existe)
    if excel_parcial_path and excel_parcial_path.exists():
        archivos_a_incluir.append(excel_parcial_path)
    # Carpeta Errores
    errores_dir = ips_dir / "Errores"
    if errores_dir.exists():
        archivos_a_incluir.extend(errores_dir.rglob("*"))

    if not archivos_a_incluir:
        return

    try:
        # Antes de crear el nuevo parcial, borrar los parciales anteriores de
        # este mismo período: el nuevo siempre es un superconjunto del viejo
        # (progreso.json nunca repite lo ya descargado), así que el anterior
        # queda redundante y solo generaría confusión si se deja ahí.
        for viejo in job["dl_dir"].glob(f'facturas_{job["periodo"]}_{job["ips_nombre"]}_PARCIAL_*.zip'):
            try:
                viejo.unlink()
            except Exception:
                pass
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_name = f'facturas_{job["periodo"]}_{job["ips_nombre"]}_PARCIAL_{timestamp}.zip'
        zip_path = job["dl_dir"] / zip_name
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for archivo in archivos_a_incluir:
                arcname = archivo.relative_to(job["dl_dir"])
                zf.write(archivo, arcname=str(arcname))
        log(job, f"📦 ZIP parcial generado (detención/error): {zip_path}")
    except Exception as e:
        log(job, f"⚠️ No se pudo generar ZIP parcial: {e}", "warn")

def crear_zip_completo(job, dl_dir, periodo, ips_nombre):
    """Crea ZIP final incluyendo PDFs, Excel final y carpeta Errores."""
    try:
        zip_final_name = f"facturas_{periodo}_{ips_nombre}.zip"
        zip_final_path = dl_dir / zip_final_name
        with zipfile.ZipFile(zip_final_path, "w", zipfile.ZIP_DEFLATED) as zf:
            ips_dir = dl_dir / ips_nombre
            if ips_dir.exists():
                # PDFs
                for pdf in ips_dir.rglob("*.pdf"):
                    zf.write(pdf, arcname=str(pdf.relative_to(dl_dir)))
                # Excel final (sin "_PARCIAL_" en el nombre)
                for excel in ips_dir.glob("reporte_*.xlsx"):
                    if "_PARCIAL_" not in excel.name:
                        zf.write(excel, arcname=str(excel.relative_to(dl_dir)))
                # Errores
                errores_dir = ips_dir / "Errores"
                if errores_dir.exists():
                    for err_file in errores_dir.rglob("*"):
                        zf.write(err_file, arcname=str(err_file.relative_to(dl_dir)))
                # Excel dedicado SOLO con errores persistentes (recargable
                # directo al bot, sin ambigüedad de hoja activa).
                with job["lock"]:
                    errores_persistentes = job["state"]["errores_detalle"].copy()
                excel_solo_errores_nombre = generar_excel_solo_errores(ips_dir, ips_nombre, errores_persistentes)
                if excel_solo_errores_nombre:
                    excel_solo_errores_path = ips_dir / excel_solo_errores_nombre
                    if excel_solo_errores_path.exists():
                        zf.write(excel_solo_errores_path, arcname=str(excel_solo_errores_path.relative_to(dl_dir)))
        log(job, f"📦 ZIP final generado: {zip_final_path}")
        # Ya terminó bien: cualquier ZIP parcial que haya quedado de intentos
        # anteriores queda obsoleto (el final ya lo incluye todo), se borra
        # para no dejarlo compitiendo visualmente con el ZIP bueno.
        for viejo in dl_dir.glob(f"facturas_{periodo}_{ips_nombre}_PARCIAL_*.zip"):
            try:
                viejo.unlink()
            except Exception:
                pass
        return str(zip_final_path)
    except Exception as e:
        log(job, f"⚠️ No se pudo generar el ZIP final: {e}", "warn")
        return None

# ==================== PERSISTENCIA (REANUDACIÓN) ====================
def cargar_progreso(job, ips_dir):
    progreso_path = ips_dir / "progreso.json"
    if progreso_path.exists():
        try:
            with open(progreso_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                # El progreso puede ser una lista simple o un diccionario con timestamps
                completadas = data.get("completadas", [])
                if isinstance(completadas, list):
                    return set(completadas)
                elif isinstance(completadas, dict):
                    return set(completadas.keys())
                else:
                    return set()
        except Exception as e:
            log(job, f"⚠️ Error al leer progreso: {e}", "warn")
    return set()

def guardar_progreso(job, ips_dir, completadas, nuevo_item=None, meta=None):
    """
    Además de la lista plana de completadas (para reanudar, igual que
    siempre), guarda un detalle por factura (tipo + fecha real) y metadata
    (identidad, ips, aseguradora, período) — necesario para poder migrar
    esto a la base de datos de historial sin perder información.
    Usa lectura-fusión-escritura: cada llamada solo necesita pasar el
    ítem NUEVO, el detalle de los anteriores se preserva del archivo.
    """
    progreso_path = ips_dir / "progreso.json"
    detalle = {}
    if progreso_path.exists():
        try:
            with open(progreso_path, "r", encoding="utf-8") as f:
                detalle = json.load(f).get("detalle", {}) or {}
        except Exception:
            pass
    if nuevo_item:
        detalle[str(nuevo_item["factura"])] = {
            "tipo": nuevo_item.get("tipo"),
            "fecha_descarga": nuevo_item.get("fecha_descarga") or datetime.now().isoformat(),
        }
    try:
        data = {"completadas": list(completadas), "detalle": detalle, "actualizado": datetime.now().isoformat()}
        if meta:
            data["meta"] = meta
        with open(progreso_path, "w", encoding="utf-8") as f:
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
            (meta or {}).get("ejecucion_id"), "previsora", (meta or {}).get("aseguradora"),
            (meta or {}).get("ips"), (meta or {}).get("periodo"),
            [{"factura": k, "estado": "exitosa", "fecha_fin": v.get("fecha_descarga")} for k, v in detalle.items()],
        )
    except Exception as e:
        log(job, f"⚠️ Error al registrar la ejecución (tabla ejecucion_facturas): {e}", "warn")

# ==================== GENERADOR DE EXCEL ====================
def generar_reporte_excel(dl_dir, periodo, ips_nombre, exitosas, errores):
    if not EXCEL_AVAILABLE:
        return None
    excel_path = dl_dir / ips_nombre / f"reporte_{periodo}.xlsx"
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()

    ws_exit = wb.active
    ws_exit.title = "Descargadas"
    ws_exit.append(["N° Factura", "Estado", "IPS", "Archivo Descargado", "Fecha/Hora"])
    for ex in exitosas:
        ws_exit.append([ex.get("factura"), ex.get("estado"), ips_nombre, ex.get("archivo"), ex.get("timestamp")])

    ws_err = wb.create_sheet("Errores")
    ws_err.append(["N° Factura", "Estado", "IPS", "Error", "Captura pantalla", "Fecha/Hora"])
    for err in errores:
        ws_err.append([err.get("factura"), err.get("estado"), ips_nombre, err.get("error"), err.get("captura"), err.get("timestamp")])

    wb.save(excel_path)
    return excel_path


def generar_excel_solo_errores(dl_dir, ips_nombre, errores):
    """
    Excel aparte, SOLO con las facturas que quedaron en error después de
    agotar los reintentos programados. Pensado para descarga automática;
    cualquier fallo aquí se ignora en silencio (el reporte combinado ya
    tiene esta misma información como respaldo).
    """
    if not errores or not EXCEL_AVAILABLE:
        return None
    try:
        nombre = f"errores_{ips_nombre}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = dl_dir / nombre
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Errores"
        ws.append(["N° Factura", "Estado", "Error", "Fecha/Hora"])
        for err in errores:
            ws.append([err.get("factura"), err.get("estado"), err.get("error"), err.get("timestamp")])
        wb.save(path)
        return nombre
    except Exception:
        return None

# ==================== FUNCIONES AUXILIARES ====================

def _find_frame_with_text(page, regex_text: str):
    js = f"() => {{ const re = new RegExp({json.dumps(regex_text)}, 'i'); return re.test(document.body?.innerText || ''); }}"
    for fr in page.frames:
        try:
            if fr.evaluate(js):
                return fr
        except:
            continue
    return None

def _cerrar_traza_factura(page):
    js = """
        () => {
            const headers = document.querySelectorAll('.ui-dialog-titlebar, .modal-header, [class*="header"]');
            for (const h of headers) {
                if (h.textContent && h.textContent.includes('Traza de Factura')) {
                    const dlg = h.closest('.ui-dialog, .modal, [role="dialog"]');
                    if (dlg) {
                        const closeBtn = dlg.querySelector('.ui-dialog-titlebar-close, button.close, [aria-label*="lose"], [class*="close"]');
                        if (closeBtn) { closeBtn.click(); return true; }
                    }
                }
            }
            return false;
        }
    """
    for fr in page.frames:
        try:
            if fr.evaluate(js):
                time.sleep(0.5)
                return
        except:
            continue

# ==================== FUNCIÓN MEJORADA DE EXTRACCIÓN DE NOMBRE DE IPS ====================
def _extraer_nombre_ips(job, page, target_frame, nit_usuario=None):
    """
    Extrae el nombre de la IPS forzando el uso del nombre exacto del mapa.
    Orden de prioridad:
    0. NIT extraído del nombre de usuario (ej: PREV900600550 → 900600550)
    1. Carpeta previa del mismo período (solo si existe en job["dl_dir"])
    2. NIT encontrado en el HTML de la página
    3. Nombre candidato por palabras clave (con búsqueda de NIT embebido)
    4. Título de la página
    5. Fallback: IPS_DESCONOCIDA
    """
    # 0. PRIMERO: si el NIT viene del nombre de usuario, úsalo directamente
    if nit_usuario and nit_usuario in MAPA_IPS:
        nombre = MAPA_IPS[nit_usuario]
        log(job, f"    🏥 IPS identificada por NIT del usuario ({nit_usuario}) -> nombre del mapa: {nombre}")
        return nombre

    def _buscar_nit_en_frame(frame):
        try:
            nit = frame.evaluate("() => { const match = document.body.innerText.match(/NIT\\s*:\\s*([\\d\\-\\s]+)/i); if(match) return match[1].replace(/[^0-9]/g, ''); return ''; }").strip()
            return nit if nit else ""
        except:
            return ""

    def _buscar_nombre_por_palabras(frame):
        keywords = ["IPS","CLINICA","HOSPITAL","CENTRO","FUNDACIÓN","URGENCIAS","SALUD","ODONTOTRANS","URGETRAUMA","CORDIALIDAD"]
        try:
            js = f"""
                () => {{
                    const keywords = {json.dumps(keywords)};
                    const elementos = document.querySelectorAll('h1, h2, h3, h4, p, div');
                    for (const el of elementos) {{
                        let txt = el.innerText.trim();
                        if (txt.length > 5 && txt.length < 100) {{
                            for (const kw of keywords) {{
                                if (txt.toUpperCase().includes(kw)) {{
                                    return txt;
                                }}
                            }}
                        }}
                    }}
                    return "";
                }}
            """
            nombre = frame.evaluate(js).strip()
            return nombre
        except:
            return ""

    # 1. Buscar NIT en todos los frames
    nit = ""
    for fr in [page] + page.frames:
        nit = _buscar_nit_en_frame(fr)
        if nit:
            log(job, f"    🔍 NIT encontrado en frame: {fr.name or 'principal'}")
            break

    # Si encontramos NIT y está en el mapa, usamos el nombre del mapa (sin añadir nada extra)
    if nit and nit in MAPA_IPS:
        nombre = MAPA_IPS[nit]
        log(job, f"    🏥 IPS identificada por NIT {nit} -> nombre forzado del mapa: {nombre}")
        return nombre

    # 2. Si no se encontró NIT o no está en el mapa, buscar por palabras clave
    nombre_candidato = ""
    for fr in [page] + page.frames:
        nombre_candidato = _buscar_nombre_por_palabras(fr)
        if nombre_candidato:
            log(job, f"    🔍 Nombre candidato encontrado: '{nombre_candidato}'")
            break

    if nombre_candidato:
        # Limpiar caracteres no válidos
        nombre_candidato = re.sub(r'[\\/*?:"<>|]', "", nombre_candidato).strip()
        # Buscar si dentro del nombre candidato hay un NIT conocido
        nit_embedded = re.search(r'\b(\d{9})\b', nombre_candidato)
        if nit_embedded and nit_embedded.group(1) in MAPA_IPS:
            nombre = MAPA_IPS[nit_embedded.group(1)]
            log(job, f"    🏥 IPS identificada por NIT embebido en texto: {nombre}")
            return nombre
        # Si no, devolver el nombre candidato limpio, pero sin añadir "_Previsora" ni similares
        nombre = re.sub(r'\s+', ' ', nombre_candidato).strip()
        log(job, f"    🏥 IPS identificada por texto: {nombre}")
        return nombre

    # 3. Fallback: usar el título de la página
    try:
        title = page.evaluate("() => document.title").strip()
        if title and len(title) > 5 and len(title) < 100:
            title = re.sub(r'Activa IT|BI IPS|Inteligencia de Negocio|Previsora|SOAT|Inicio', '', title, flags=re.I).strip()
            if title:
                log(job, f"    🏥 IPS obtenida del título: {title}")
                return title
    except:
        pass

    # 4. Fallback definitivo
    log(job, "    ⚠️ No se pudo determinar la IPS, se usará 'IPS_DESCONOCIDA'", "warn")
    return "IPS_DESCONOCIDA"

# ==================== FUNCIÓN _download_factura (sin cambios) ====================
def _download_factura(job, page, context, modal_frame, fac: dict, dl_dir: Path, ips_nombre: str):
    import re
    num = fac["num"]
    tipo = fac["tipo"]

    # Determinar etiquetas según el tipo
    if tipo == "devolucion":
        target_label = "ActaDevolucion"
        target_label_norm = target_label.replace('ó', 'o').replace('í', 'i')
        subcarpeta = "Devolucion"
        nombre_soporte = "ActaDevolución"
    else:
        target_label = "Envios_D"
        target_label_norm = target_label.replace('í', 'i')
        subcarpeta = "Auditada"
        nombre_soporte = "Envios_D"

    ips_dir = dl_dir / ips_nombre
    dl_subdir = ips_dir / subcarpeta
    dl_subdir.mkdir(parents=True, exist_ok=True)

    bot_id = fac.get("botId")
    log(job, f"    🔗 Abriendo factura {num}...")
    num_solo_digitos = re.sub(r'\D', '', str(num))

    js_click_robusto = f"""
        () => {{
            const botId = '{bot_id}';
            const targetDigits = '{num_solo_digitos}';

            function dispararClick(el) {{
                if (!el) return false;
                try {{ el.click(); }} catch (e) {{}}
                try {{ el.dispatchEvent(new MouseEvent('click', {{bubbles: true, cancelable: true, view: window}})); }} catch (e) {{}}
                return true;
            }}

            function clickearFila(fila, metodo) {{
                fila.scrollIntoView({{block: 'center'}});
                const candidatos = [];
                for (const a of fila.querySelectorAll('a')) {{
                    const t = (a.textContent || '').trim();
                    if (t.replace(/\\D/g, '') === targetDigits || candidatos.length === 0)
                        candidatos.push(a);
                }}
                for (const el of fila.querySelectorAll('[onclick]')) {{
                    if (!candidatos.includes(el)) candidatos.push(el);
                }}
                candidatos.push(fila);
                for (const td of fila.querySelectorAll('td')) candidatos.push(td);
                for (const c of candidatos) dispararClick(c);
                return {{ ok: true, clickedWith: metodo, candidates: candidatos.length }};
            }}

            let fila = botId ? document.querySelector(`[data-bot-row-id="${{botId}}"]`) : null;
            if (fila) return clickearFila(fila, 'botId');

            for (const row of document.querySelectorAll('tr')) {{
                for (const a of row.querySelectorAll('a')) {{
                    if ((a.textContent || '').replace(/\\D/g, '') === targetDigits) {{
                        return clickearFila(row, 'numero_factura');
                    }}
                }}
            }}

            for (const row of document.querySelectorAll('tr')) {{
                for (const td of row.querySelectorAll('td')) {{
                    if ((td.textContent || '').replace(/\\D/g, '') === targetDigits) {{
                        return clickearFila(row, 'celda_numero');
                    }}
                }}
            }}

            return {{ ok: false, reason: "fila_no_encontrada" }};
        }}
    """
    result = None
    try:
        result = modal_frame.evaluate(js_click_robusto)
    except Exception as e:
        log(job, f"    ⚠️ Click falló: {e}", "warn")
    if not result or not result.get("ok"):
        for fr in page.frames:
            try:
                r = fr.evaluate(js_click_robusto)
                if r and r.get("ok"):
                    result = r
                    break
            except:
                continue
    if not result or not result.get("ok"):
        raise Exception(f"Click totalmente fallido para factura {num}.")
    log(job, f"    ✓ Click en factura {num} OK.")
    time.sleep(1.5)

    detalle_state = None
    detalle_frame = None
    for _ in range(60):
        if job["state"].get("stopping"): return
        f = _find_frame_with_text(page, "Adjuntos por Factura")
        if f:
            try:
                has_traza = f.evaluate("() => /Traza de Factura/i.test(document.body?.innerText || '')")
                detalle_state = "traza" if has_traza else "adjuntos_directo"
            except:
                detalle_state = "adjuntos_directo"
            detalle_frame = f
            break
        f = _find_frame_with_text(page, "Traza de Factura")
        if f:
            detalle_state = "traza"
            detalle_frame = f
        time.sleep(0.5)
    if not detalle_frame:
        raise Exception("No apareció 'Traza de Factura' ni 'Adjuntos por Factura'.")
    time.sleep(1.5)
    log(job, f"    ✅ Detalle abierto (modo: {detalle_state}).")

    if detalle_state == "traza":
        log(job, "    📑 Forzando cambio a pestaña 'Soportes'...")
        soportes_ok = False
        for intento in range(5):
            if job["state"].get("stopping"): return
            for fr in page.frames:
                try:
                    has_tabs = fr.evaluate(r"""() => {
                        const txt = (document.body?.innerText || '').replace(/\n/g, ' ');
                        return /Factura.*Detalles.*Soportes/i.test(txt);
                    }""")
                    if has_tabs:
                        try:
                            fr.locator("text=Soportes").first.click(timeout=5000)
                            soportes_ok = True
                            break
                        except:
                            clicked = fr.evaluate("""() => {
                                for (const el of document.querySelectorAll('*')) {
                                    if ((el.textContent||'').trim() === 'Soportes') {
                                        el.click(); return true;
                                    }
                                }
                                return false;
                            }""")
                            if clicked:
                                soportes_ok = True
                                break
                except:
                    continue
            if soportes_ok:
                break
            time.sleep(1)
        if not soportes_ok:
            log(job, "    ⚠️ No se pudo clickear Soportes", "warn")
        else:
            time.sleep(3)

    log(job, "    ⏳ Esperando 'Adjuntos por Factura'...")
    adjuntos_frame = None
    for _ in range(90):
        if job["state"].get("stopping"): return
        for fr in page.frames:
            try:
                if fr.evaluate("() => /Adjuntos por Factura|Buscar por.*Fecha/i.test(document.body?.innerText || '')"):
                    adjuntos_frame = fr
                    break
            except:
                continue
        if adjuntos_frame:
            break
        time.sleep(0.5)
    if not adjuntos_frame:
        raise Exception("No se encontró sección 'Adjuntos por Factura'.")
    for _ in range(35):
        if job["state"].get("stopping"): return
        try:
            busy = adjuntos_frame.evaluate("() => /Procesando Solicitud/i.test(document.body?.innerText || '')")
            if not busy:
                break
        except:
            pass
        time.sleep(1)
    time.sleep(1)
    log(job, "    ✅ Adjuntos cargados.")

    search_frame = adjuntos_frame

    # Función mejorada para escribir en el buscador y disparar la lupa
    def _escribir_buscador(texto):
        # Limpiar input antes
        search_frame.evaluate("""
            () => {
                const inputs = document.querySelectorAll('input');
                for (const input of inputs) {
                    const ph = (input.placeholder || '').toLowerCase();
                    if (ph.includes('buscar') || ph.includes('filtrar') || ph.includes('nombre')) {
                        input.value = '';
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        break;
                    }
                }
            }
        """)
        time.sleep(0.5)
        # Escribir el texto y disparar eventos
        search_frame.evaluate(f"""
            () => {{
                const target = '{texto.replace('í', 'i')}';
                const inputs = document.querySelectorAll('input');
                let searchInput = null;
                for (const input of inputs) {{
                    const ph = (input.placeholder || '').toLowerCase();
                    if (ph.includes('buscar') || ph.includes('filtrar') || ph.includes('nombre')) {{
                        searchInput = input;
                        break;
                    }}
                }}
                if (!searchInput) return;
                searchInput.focus();
                searchInput.select();
                const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                if (nativeSetter) nativeSetter.call(searchInput, target);
                else searchInput.value = target;
                searchInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                searchInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
                // Buscar botón de lupa
                let parent = searchInput.closest('div, td, form, span');
                if (parent) {{
                    const btns = parent.querySelectorAll('button, a, [role="button"], span');
                    for (const btn of btns) {{
                        const html = (btn.outerHTML || '').toLowerCase();
                        const title = (btn.title || '').toLowerCase();
                        if (html.includes('search') || html.includes('lup') || title.includes('search')) {{
                            btn.click();
                            return;
                        }}
                    }}
                }}
                const svgs = document.querySelectorAll('svg');
                for (const svg of svgs) {{
                    if ((svg.outerHTML || '').toLowerCase().includes('search')) {{
                        const container = svg.closest('button, a, [role="button"]');
                        if (container) {{ container.click(); return; }}
                    }}
                }}
                // Fallback: presionar Enter
                searchInput.dispatchEvent(new KeyboardEvent('keypress', {{ key: 'Enter', bubbles: true }}));
            }}
        """)
        time.sleep(2)
        # Esperar a que termine "Procesando Solicitud"
        for _ in range(40):
            if job["state"].get("stopping"): return
            processing = False
            for fr in page.frames:
                try:
                    if fr.evaluate("() => /Procesando Solicitud/i.test(document.body?.innerText || '')"):
                        processing = True
                        break
                except:
                    pass
            if not processing:
                break
            time.sleep(0.5)
        time.sleep(2)

    # ---------- BÚSQUEDA DE SOPORTES CON MEJORAS ----------
    # Intentar primero con la etiqueta principal (Envios_D o ActaDevolucion)
    log(job, f"    🔍 Buscando '{target_label}'...")
    _escribir_buscador(target_label)
    archivo_seleccionado = False
    tipo_encontrado = None  # Rastrear qué tipo de soporte se encontró
    posibles_nombres = list({target_label, target_label_norm})

    for intento in range(4):
        if job["state"].get("stopping"): return
        for fr in page.frames:
            try:
                resultado = fr.evaluate(f"""
                    () => {{
                        const nombres = {json.dumps(posibles_nombres)};
                        let contenedor = null;
                        const elementos = document.querySelectorAll('td, div, span, li, p, tr');
                        for (const el of elementos) {{
                            const txt = (el.innerText || '').trim();
                            for (const nombre of nombres) {{
                                if (txt === nombre) {{
                                    contenedor = el.closest('div[class*="file"], li[class*="file"], tr, div[class*="item"], div[class*="attach"], div[class*="row"]');
                                    if (!contenedor) contenedor = el.closest('div, li, tr');
                                    break;
                                }}
                            }}
                            if (contenedor) break;
                        }}
                        if (!contenedor) return {{ ok: false }};
                        let check = contenedor.querySelector('input[type="checkbox"], input[type="radio"], [role="checkbox"]');
                        if (!check) check = contenedor.parentElement?.querySelector('input[type="checkbox"], input[type="radio"], [role="checkbox"]');
                        if (check) {{
                            if (!check.checked) {{
                                check.click();
                                check.checked = true;
                                check.dispatchEvent(new Event('change', {{ bubbles: true }}));
                            }}
                            return {{ ok: true, metodo: 'checkbox' }};
                        }}
                        let iconoPdf = null;
                        const candidatosPdf = contenedor.querySelectorAll('img, svg, i, div');
                        for (const el of candidatosPdf) {{
                            const src = el.getAttribute('src') || '';
                            const lbl = el.getAttribute('aria-label') || '';
                            const cls = el.className || '';
                            if (src.toLowerCase().includes('pdf') || lbl.toLowerCase().includes('pdf') ||
                                cls.toLowerCase().includes('pdf') || cls.toLowerCase().includes('file')) {{
                                iconoPdf = el; break;
                            }}
                        }}
                        if (iconoPdf) {{
                            iconoPdf.click();
                            return {{ ok: true, metodo: 'icono_pdf' }};
                        }}
                        contenedor.click();
                        contenedor.dispatchEvent(new MouseEvent('click', {{ bubbles: true, cancelable: true }}));
                        contenedor.dispatchEvent(new MouseEvent('dblclick', {{ bubbles: true, cancelable: true }}));
                        return {{ ok: true, metodo: 'contenedor_forzado' }};
                    }}
                """)
                if resultado and resultado.get('ok'):
                    log(job, f"    ✅ Selección realizada (método: {resultado.get('metodo')})")
                    archivo_seleccionado = True
                    tipo_encontrado = nombre_soporte  # Guardar el tipo encontrado
                    break
            except Exception as e:
                log(job, f"    ⚠️ Error en intento {intento+1}: {e}", "warn")
        if archivo_seleccionado:
            break
        log(job, f"    🔄 Reintentando selección ({intento+1}/4)...")
        time.sleep(0.8)

    # ---------- SEGUNDA BÚSQUEDA: CARTA DE OBJECIÓN (si no se encontró Envios_D/ActaDevolucion) ----------
    if not archivo_seleccionado:
        log(job, f"    ⚠️ No se encontró '{target_label}'. Intentando con 'Carta de'...")
        texto_busqueda = "Carta de"
        _escribir_buscador(texto_busqueda)

        archivo_seleccionado = False
        for intento in range(4):
            if job["state"].get("stopping"): return
            for fr in page.frames:
                try:
                    resultado = fr.evaluate(f"""
                        () => {{
                            const buscarTexto = '{texto_busqueda}';
                            function normalizar(s) {{
                                return s.toLowerCase().normalize("NFD").replace(/[\\u0300-\\u036f]/g, "");
                            }}
                            const elementos = document.querySelectorAll('td, div, span, li, p, tr');
                            for (const el of elementos) {{
                                const txt = (el.innerText || '').trim();
                                if (normalizar(txt).includes(normalizar(buscarTexto))) {{
                                    let contenedor = el.closest('div[class*="file"], li[class*="file"], tr, div[class*="item"], div[class*="attach"], div[class*="row"]');
                                    if (!contenedor) contenedor = el.closest('div, li, tr');
                                    if (contenedor) {{
                                        let check = contenedor.querySelector('input[type="checkbox"], input[type="radio"], [role="checkbox"]');
                                        if (!check) check = contenedor.parentElement?.querySelector('input[type="checkbox"], input[type="radio"], [role="checkbox"]');
                                        if (check) {{
                                            if (!check.checked) {{
                                                check.click();
                                                check.checked = true;
                                                check.dispatchEvent(new Event('change', {{ bubbles: true }}));
                                            }}
                                            return {{ ok: true, metodo: 'checkbox', texto: txt }};
                                        }}
                                        let iconoPdf = null;
                                        const candidatosPdf2 = contenedor.querySelectorAll('img, svg, i, div');
                                        for (const el of candidatosPdf2) {{
                                            const src = el.getAttribute('src') || '';
                                            const lbl = el.getAttribute('aria-label') || '';
                                            const cls = el.className || '';
                                            if (src.toLowerCase().includes('pdf') || lbl.toLowerCase().includes('pdf') ||
                                                cls.toLowerCase().includes('pdf')) {{
                                                iconoPdf = el; break;
                                            }}
                                        }}
                                        if (iconoPdf) {{
                                            iconoPdf.click();
                                            return {{ ok: true, metodo: 'icono_pdf', texto: txt }};
                                        }}
                                        contenedor.click();
                                        contenedor.dispatchEvent(new MouseEvent('click', {{ bubbles: true, cancelable: true }}));
                                        contenedor.dispatchEvent(new MouseEvent('dblclick', {{ bubbles: true, cancelable: true }}));
                                        return {{ ok: true, metodo: 'contenedor_forzado', texto: txt }};
                                    }}
                                }}
                            }}
                            return {{ ok: false }};
                        }}
                    """)
                    if resultado and resultado.get('ok'):
                        log(job, f"    ✅ Selección realizada con '{texto_busqueda}' (método: {resultado.get('metodo')}) - Texto encontrado: '{resultado.get('texto')}'")
                        archivo_seleccionado = True
                        tipo_encontrado = "Carta de Objecion"  # Guardar el tipo encontrado
                        break
                except Exception as e:
                    log(job, f"    ⚠️ Error en intento {intento+1} para '{texto_busqueda}': {e}", "warn")
            if archivo_seleccionado:
                break
            log(job, f"    🔄 Reintentando '{texto_busqueda}' ({intento+1}/4)...")
            time.sleep(2)

    if not archivo_seleccionado:
        raise Exception(f"No se pudo seleccionar el archivo (intentó '{target_label}' y 'Carta de')")

    # Confirmar que no haya mensaje de error "Debe seleccionar..."
    log(job, "    ⏳ Esperando confirmación de selección...")
    for _ in range(20):
        if job["state"].get("stopping"): return
        hay_error = False
        for fr in page.frames:
            try:
                if fr.evaluate("() => /Debe seleccionar por lo menos un documento/i.test(document.body?.innerText || '')"):
                    hay_error = True
                    break
            except:
                pass
        if not hay_error:
            log(job, "    ✅ Selección confirmada")
            break
        time.sleep(1)

    # ---------- ABRIR DOCUMENTO ----------
    log(job, f"    👁️ Buscando botón 'Abrir Documento'...")
    pdf_data = None
    pdf_url = None

    boton_encontrado = False
    start_time = time.time()
    while time.time() - start_time < 15:
        if job["state"].get("stopping"): return
        for fr in page.frames:
            try:
                btn = fr.locator('button[title="Abrir Documento"], button[aria-label="Abrir Documento"], button:has(i.fa-eye), button:has(i.bi-eye)').first
                if btn.is_visible(timeout=2000):
                    boton_encontrado = True
                    break
            except:
                pass
        if boton_encontrado:
            break
        time.sleep(0.5)
    else:
        raise Exception("Botón 'Abrir Documento' no encontrado")

    for reintento in range(2):
        if job["state"].get("stopping"): return
        new_page = None
        try:
            with context.expect_page(timeout=30000) as page_info:
                for fr in page.frames:
                    try:
                        btn = fr.locator('button[title="Abrir Documento"], button[aria-label="Abrir Documento"], button:has(i.fa-eye), button:has(i.bi-eye)').first
                        if btn.is_visible(timeout=5000):
                            for _ in range(10):
                                if btn.is_enabled():
                                    break
                                time.sleep(0.5)
                            btn.click()
                            log(job, "    ✅ Clic en botón 'Abrir Documento'")
                            break
                    except:
                        pass
            new_page = page_info.value
            for _ in range(30):
                if job["state"].get("stopping"): return
                url = new_page.url
                if url and url != "about:blank" and ("amazonaws" in url or ".pdf" in url.lower()):
                    pdf_url = url
                    break
                time.sleep(0.5)
        except Exception as e:
            log(job, f"    ⚠️ Intento {reintento+1}: No se abrió nueva pestaña: {e}", "warn")
        finally:
            if new_page:
                try:
                    new_page.close()
                except:
                    pass

        if pdf_url:
            try:
                response = context.request.get(pdf_url, timeout=60000)
                if response.ok:
                    pdf_data = response.body()
                    log(job, f"    ✅ PDF descargado ({len(pdf_data)//1024} KB)")
                    break
            except Exception as e:
                log(job, f"    ⚠️ Error descargando: {e}", "warn")

        if not pdf_data:
            log(job, "    ⏳ Intentando descarga directa...")
            try:
                with page.expect_download(timeout=30000) as download_info:
                    for fr in page.frames:
                        try:
                            btn = fr.locator('button[title="Abrir Documento"], button:has(i.fa-eye), button:has(i.bi-eye)').first
                            if btn.is_visible(timeout=3000):
                                btn.click()
                                break
                        except:
                            pass
                download = download_info.value
                pdf_data = download.path().read_bytes() if download.path() else None
                log(job, "    ✅ Descarga directa capturada")
                break
            except Exception as e:
                log(job, f"    ⚠️ No se capturó descarga: {e}", "warn")

        if not pdf_data:
            log(job, f"    🔄 Reintento {reintento+1}/2...")
            time.sleep(2)

    if not pdf_data:
        raise Exception("No se pudo obtener el PDF")

    # ---------- MEJORAR NOMBRE DEL ARCHIVO ----------
    # Determinar el tipo de soporte encontrado para el nombre del archivo
    soporte_encontrado = tipo_encontrado if tipo_encontrado else nombre_soporte

    # El nombre del archivo será: {num}_{soporte_encontrado}.pdf
    safe_name = re.sub(r"[^\w\-_.]", "_", f"{num}_{soporte_encontrado}.pdf")
    out_path = dl_subdir / safe_name
    out_path.write_bytes(pdf_data)
    log(job, f"    💾 PDF guardado: {out_path.name} ({len(pdf_data)//1024} KB)")

    with job["lock"]:
        job["state"]["descargas_exitosas"].append({
            "factura": num,
            "estado": fac["estado"],
            "archivo": str(out_path),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    _cerrar_traza_factura(page)
    time.sleep(0.8)

# ==================== AUTOMATIZACIÓN PRINCIPAL ====================
# (El resto del código de run_automation y rutas Flask es idéntico al original,
#  no se modifica nada más. Para ahorrar espacio, se incluye tal cual estaba,
#  pero asegurando que la función _extraer_nombre_ips es la nueva.)

def run_automation(job, usuario: str, password: str, periodo: str, download_path: str, reintentos_largos: bool = True):
    from playwright.sync_api import sync_playwright

    dl_dir = Path(download_path)
    dl_dir.mkdir(parents=True, exist_ok=True)
    ips_nombre_actual = "IPS_SIN_NOMBRE"
    zip_parcial_generado = False

    job["dl_dir"] = dl_dir
    job["periodo"] = periodo

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
            context = browser.new_context(accept_downloads=True, viewport={"width": 1500, "height": 900})
            page = context.new_page()
            job["browser"] = browser
            job["context"] = context

            log(job, "🔐 Iniciando sesión en Activa IT...")
            if job["state"].get("stopping"): return
            page.goto("https://activa-it.net/Login.aspx", wait_until="networkidle", timeout=60000)
            log(job, f"  → Usuario: {usuario}")
            page.fill('input[placeholder="Usuario"]', usuario)
            page.fill('input[placeholder="Contraseña"]', password)
            try:
                checkbox = page.locator('input[type="checkbox"]').first
                if not checkbox.is_checked():
                    checkbox.check()
            except:
                pass
            page.click('button:has-text("Inicio de sesión"), input[value="Inicio de sesión"]')
            page.wait_for_url("**/Index.aspx", timeout=60000)
            try:
                page.wait_for_selector("text=BI IPS, text=Inteligencia de Negocio", timeout=10000)
            except:
                pass
            # ── Cerrar popups/alertas inesperados (mantenimiento, cookies, etc.) ──
            def _cerrar_popups():
                try:
                    page.evaluate("""() => {
                        // 1. Buscar específicamente el modal de Activa IT (Corte de Sistema, etc.)
                        for (const el of document.querySelectorAll('*')) {
                            const txt = el.innerText || '';
                            if (/Corte de Sistema|mantenimiento|suspendido/i.test(txt) && el.offsetParent !== null) {
                                const btns = [...el.querySelectorAll('button, a, input[type="button"]')];
                                const close = btns.find(b => /cerrar|close|aceptar|ok|entendido|continuar/i.test(b.textContent || b.value || ''));
                                if (close) { close.click(); return true; }
                            }
                        }
                        // 2. Buscar cualquier modal/dialog visible genérico
                        const selectores = [
                            '.modal', '.dialog', '.alert', '.popup', '.overlay',
                            '[role="dialog"]', '[role="alertdialog"]',
                            '.ui-dialog', '.sweet-alert', '.swal2-container',
                            '.modal-dialog', '.modal-content'
                        ];
                        for (const sel of selectores) {
                            for (const el of document.querySelectorAll(sel)) {
                                if (el.offsetParent === null) continue;
                                const btns = [...el.querySelectorAll('button, a, input[type="button"]')];
                                const close = btns.find(b => /cerrar|close|aceptar|ok|entendido|continuar|×|✕/i.test(b.textContent || b.getAttribute('aria-label') || b.value || ''));
                                if (close) { close.click(); return true; }
                            }
                        }
                        return false;
                    }""")
                except:
                    pass
                # También buscar en frames internos
                for fr in page.frames:
                    try:
                        found = fr.evaluate("""() => {
                            for (const el of document.querySelectorAll('*')) {
                                const txt = el.innerText || '';
                                if (/Corte de Sistema|mantenimiento|suspendido/i.test(txt) && el.offsetParent !== null) {
                                    const btns = [...el.querySelectorAll('button, a, input[type="button"]')];
                                    const close = btns.find(b => /cerrar|close|aceptar|ok/i.test(b.textContent || b.value || ''));
                                    if (close) { close.click(); return true; }
                                }
                            }
                            return false;
                        }""")
                        if found:
                            break
                    except:
                        continue
                try:
                    page.keyboard.press("Escape")
                except:
                    pass

            _cerrar_popups()
            log(job, "✅ Sesión iniciada correctamente.")
            if job["state"].get("stopping"): return

            log(job, "📂 Navegando a módulo BI IPS...")

            def _find_periodo_in_frames():
                js_check = f"""
                    () => {{
                        const bodyText = (document.body?.innerText || '').toLowerCase();
                        const periodo = '{periodo}'.toLowerCase();
                        if (bodyText.includes(periodo)) return true;
                        const variaciones = ['abr26', 'abr-26', 'abr.26', 'abr/26', 'abr2026'];
                        return variaciones.some(v => bodyText.includes(v));
                    }}
                """
                for fr in page.frames:
                    try:
                        if fr.evaluate(js_check):
                            return fr
                    except:
                        continue
                return None

            if job["state"].get("stopping"): return
            clicked = False
            for intento in range(3):
                try:
                    page.locator("text=BI IPS").first.click(timeout=15000)
                    clicked = True
                    log(job, "  ✓ Click directo en 'BI IPS' OK.")
                    break
                except:
                    pass
                try:
                    page.click("text=Inteligencia de Negocio", timeout=8000)
                    time.sleep(1)
                    page.click("text=BI IPS", timeout=8000)
                    clicked = True
                    log(job, "  ✓ Click vía 'Inteligencia de Negocio' + 'BI IPS' OK.")
                    break
                except:
                    pass
                try:
                    page.click("[class*='menu-toggle'], [class*='hamburger'], .sidebar-toggle", timeout=5000)
                    time.sleep(2)
                    page.click("text=BI IPS", timeout=8000)
                    clicked = True
                    log(job, "  ✓ Click vía hamburguesa + 'BI IPS' OK.")
                    break
                except Exception as e:
                    log(job, f"    ⚠️ Intento {intento+1} falló: {e}", "warn")
                    time.sleep(2)
            if not clicked:
                raise Exception("No se encontró el módulo BI IPS en el menú.")

            try:
                page.wait_for_load_state("networkidle", timeout=8000)
            except:
                pass
            _cerrar_popups()
            log(job, "✅ Módulo BI IPS abierto. Buscando período...")
            target_frame = None
            for i in range(120):
                if job["state"].get("stopping"): return
                target_frame = _find_periodo_in_frames()
                if target_frame:
                    log(job, f"✅ Período '{periodo}' detectado tras {(i+1)*0.5:.1f}s.")
                    break
                time.sleep(0.5)
            if not target_frame:
                raise Exception(f"No se pudo localizar el período '{periodo}' tras 60s.")

            log(job, "🏥 Obteniendo nombre de la IPS...")
            # Intentar extraer NIT del nombre de usuario (ej: PREV900600550 → 900600550)
            nit_from_usuario = re.search(r'(\d{9,12})', usuario)
            nit_from_usuario = nit_from_usuario.group(1) if nit_from_usuario else None
            if nit_from_usuario:
                log(job, f"    🔑 NIT extraído del usuario '{usuario}': {nit_from_usuario}")
            ips_nombre_actual = _extraer_nombre_ips(job, page, target_frame, nit_usuario=nit_from_usuario)
            job["ips_nombre"] = ips_nombre_actual
            # Refresca la misma ips_identity que ya usan guardar_progreso() y
            # el registro de errores, con lo realmente detectado en esta
            # corrida (no solo la suposición hecha antes de arrancar).
            with job["lock"]:
                job["state"]["ips_identity"] = resolver(nit=nit_from_usuario, nombre_detectado=ips_nombre_actual)
                job["state"]["ips_nit"] = nit_from_usuario

            if job["state"].get("stopping"): return
            log(job, f"📅 Click en columna Cant del período '{periodo}'...")
            click_result = target_frame.evaluate(f"""
                () => {{
                    const rows = document.querySelectorAll('tr');
                    for (const row of rows) {{
                        const cells = row.querySelectorAll('td');
                        if (cells.length < 3) continue;
                        const firstText = cells[0].textContent.trim();
                        if (firstText !== '{periodo}') continue;
                        const links = row.querySelectorAll('a');
                        if (links.length === 0) return {{ ok: false, reason: 'sin_links' }};
                        const firstLink = links[0];
                        const value = firstLink.textContent.trim();
                        if (value === '0') return {{ ok: false, reason: 'cant_cero', value: '0' }};
                        firstLink.scrollIntoView({{block: 'center'}});
                        firstLink.click();
                        return {{ ok: true, value: value }};
                    }}
                    return {{ ok: false, reason: 'fila_no_encontrada' }};
                }}
            """)
            if click_result.get("reason") == "cant_cero":
                log(job, f"ℹ️ El período '{periodo}' tiene 0 facturas radicadas.", "warn")
                browser.close()
                return
            if not click_result.get("ok"):
                raise Exception(f"No se pudo hacer click en Cant de '{periodo}': {click_result.get('reason')}")
            log(job, f"  → Click en Cant: {click_result.get('value')}")

            log(job, "⏳ Esperando modal 'Listado de facturas recibidas'...")
            modal_frame = None
            for _ in range(60):
                if job["state"].get("stopping"): return
                for fr in page.frames:
                    try:
                        if fr.evaluate("() => /Listado de facturas recibidas/i.test(document.body?.innerText || '')"):
                            modal_frame = fr
                            break
                    except:
                        continue
                if modal_frame:
                    break
                time.sleep(0.5)
            if not modal_frame:
                raise Exception("El modal 'Listado de facturas recibidas' no apareció.")

            log(job, "⏳ Esperando datos del listado...")
            data_frame = None
            tiempo_espera = 0
            while tiempo_espera < 60:
                if job["state"].get("stopping"): return
                for fr in page.frames:
                    try:
                        if fr.evaluate("() => /Pendiente de recibir Informaci|Devoluci[oó]n de entrada/i.test(document.body?.innerText || '')"):
                            data_frame = fr
                            break
                    except:
                        continue
                if data_frame:
                    break
                time.sleep(0.5)
                tiempo_espera += 0.5

            if not data_frame:
                log(job, "⚠️ No se encontraron facturas con los estados objetivo.", "warn")
                browser.close()
                return

            log(job, f"✅ Datos detectados en frame '{data_frame.name or '(main)'}'.")
            time.sleep(2)

            log(job, "🔍 Extrayendo facturas...")
            js_extract = r"""
            (state) => {
                const ESTADOS = [
                    { nombre: 'Auditada: Pendiente de recibir Informacion', regex: /auditada\s*:\s*pendiente\s+de\s+recibir\s+informaci[oó]n/i, tipo: 'auditada' },
                    { nombre: 'En radicacion: Devolución de entrada', regex: /en\s+radicaci[oó]n\s*:\s*devoluci[oó]n\s+de\s+entrada/i, tipo: 'devolucion' },
                    { nombre: 'En auditoria: Pendiente de informar Orden de pago al Pagador', regex: /en\s+auditori?a\s*:\s*pendiente\s+de\s+informar\s+orden\s+de\s+pago\s+al\s+pagador/i, tipo: 'auditada' },
                ];
                const filas = document.querySelectorAll('tr, [role="row"], li');
                const nuevas = [];
                for (const fila of filas) {
                    const fullText = (fila.innerText || '').replace(/\s+/g, ' ').trim();
                    if (!fullText || fullText.length < 20 || fullText.length > 400) continue;
                    if (!/\d{2}\/\d{2}\/\d{4}/.test(fullText)) continue;
                    let tipoDetectado = null, nombreEstado = null;
                    for (const e of ESTADOS) {
                        if (e.regex.test(fullText)) { tipoDetectado = e.tipo; nombreEstado = e.nombre; break; }
                    }
                    if (!tipoDetectado) continue;
                    const tokens = fullText.split(/\s+/);
                    const candidatosNum = tokens.filter(t => { const digits = t.replace(/\D/g, ''); return digits.length >= 6 && digits.length <= 10; });
                    if (candidatosNum.length === 0 || candidatosNum.length > 6) continue;
                    const numNorm = candidatosNum[0].replace(/\D/g, '');
                    if (state.seen.includes(numNorm)) continue;
                    const botId = 'bot_' + state.nextId;
                    state.nextId++;
                    fila.setAttribute('data-bot-row-id', botId);
                    nuevas.push({
                        botId: botId, num: numNorm, rawNum: candidatosNum[0],
                        tipo: tipoDetectado, estado: nombreEstado,
                        textoFila: fullText.slice(0, 150), tagName: fila.tagName.toLowerCase(),
                    });
                    state.seen.push(numNorm);
                }
                return { nuevas: nuevas, total: state.seen.length };
            }
            """
            extract_state = {"nextId": 0, "seen": []}
            facturas_acumuladas = []
            rondas_sin_nuevos = 0
            for ronda in range(20):
                if job["state"].get("stopping"): return
                try:
                    res = data_frame.evaluate(js_extract, extract_state)
                except:
                    res = {"nuevas": []}
                nuevas = res.get("nuevas", [])
                if nuevas:
                    facturas_acumuladas.extend(nuevas)
                    rondas_sin_nuevos = 0
                    log(job, f"  Ronda {ronda+1}: +{len(nuevas)} (Total: {len(facturas_acumuladas)})")
                else:
                    rondas_sin_nuevos += 1
                extract_state["seen"] = list(set(extract_state["seen"] + [n["num"] for n in nuevas]))
                if rondas_sin_nuevos >= 5:
                    break
                try:
                    data_frame.evaluate("() => { const scrollables = document.querySelectorAll('div, table, tbody, [class*=\"scroll\"]'); for (const s of scrollables) { if (s.scrollHeight > s.clientHeight + 20) s.scrollTop += s.clientHeight * 0.8; } window.scrollBy(0, window.innerHeight * 0.8); }")
                except:
                    pass
                time.sleep(0.5)
            log(job, f"📊 {len(facturas_acumuladas)} facturas detectadas.")
            facturas_objetivo = facturas_acumuladas

            # ========== PERSISTENCIA Y FILTRO ==========
            ips_dir = dl_dir / ips_nombre_actual
            completadas = cargar_progreso(job, ips_dir)

            try:
                _ips_actual = job["state"].get("ips_identity") or {}
                if _ips_actual.get("nit"):
                    ya_en_historial = historial_db.buscar_ya_descargadas_por_nit(
                        "Previsora", _ips_actual["nit"], [fac['num'] for fac in facturas_objetivo]
                    )
                else:
                    ya_en_historial = historial_db.buscar_ya_descargadas(
                        "Previsora", ips_nombre_actual, [fac['num'] for fac in facturas_objetivo]
                    )
                nuevas_en_historial = {f: v for f, v in ya_en_historial.items() if f not in completadas}
                if nuevas_en_historial:
                    with job["lock"]:
                        job["state"]["duplicados_pendientes"] = {"ya_descargadas": nuevas_en_historial}
                        job["state"]["duplicate_event"].clear()
                    while not job["state"].get("duplicate_event").wait(0.5):
                        if job["state"].get("stopping"):
                            return
                    decision = job["state"].get("decision_redescarga") or "ninguna"
                    seleccionadas = set(job["state"].get("facturas_redescarga") or [])
                    if decision == "ninguna":
                        facturas_objetivo = [f for f in facturas_objetivo if f["num"] not in nuevas_en_historial]
                    elif decision == "seleccionadas":
                        facturas_objetivo = [f for f in facturas_objetivo if f["num"] not in nuevas_en_historial or f["num"] in seleccionadas]
                    with job["lock"]:
                        job["state"]["duplicados_pendientes"] = None
            except Exception:
                pass

            # Filtrar facturas ya descargadas
            facturas_pendientes = []
            for fac in facturas_objetivo:
                if fac['num'] in completadas:
                    log(job, f"⏭️ Factura {fac['num']} ya descargada en ejecución anterior, omitiendo.")
                    with job["lock"]:
                        job["state"]["stats"]["descargadas"] += 1
                        job["state"]["descargas_exitosas"].append({
                            "factura": fac['num'],
                            "estado": fac['estado'],
                            "archivo": str(ips_dir / ("Auditada" if fac['tipo']=='auditada' else "Devolucion") / f"Factura_{fac['num']}_{('Envios_D' if fac['tipo']=='auditada' else 'ActaDevolucion')}.pdf"),
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                else:
                    facturas_pendientes.append(fac)

            # Aplicar filtro opcional por lista de facturas permitidas
            with job["lock"]:
                permitidas = job["state"].get("facturas_permitidas", [])
            if permitidas:
                original_count = len(facturas_pendientes)
                facturas_pendientes = [fac for fac in facturas_pendientes if fac['num'] in permitidas]
                log(job, f"📋 Filtro activo: solo {len(facturas_pendientes)} de {original_count} facturas están en la lista permitida.")

            log(job, f"📋 Facturas pendientes por procesar en esta ejecución: {len(facturas_pendientes)}")

            # Actualizar estadísticas totales
            with job["lock"]:
                job["state"]["stats"]["total"] = len(facturas_pendientes) + job["state"]["stats"]["descargadas"]
                job["state"]["stats"]["errores"] = 0

            cnt_aud = sum(1 for f in facturas_pendientes if f["tipo"] == "auditada")
            cnt_dev = sum(1 for f in facturas_pendientes if f["tipo"] == "devolucion")
            log(job, "📋 RESUMEN DE FACTURAS PENDIENTES:")
            log(job, f"  • Auditada: {cnt_aud}")
            log(job, f"  • Devolucion: {cnt_dev}")
            log(job, f"  TOTAL: {len(facturas_pendientes)}")
            if not facturas_pendientes:
                log(job, "ℹ️ No hay facturas pendientes por procesar.")
                browser.close()
                with job["lock"]:
                    exitosas = job["state"]["descargas_exitosas"].copy()
                    errores = job["state"]["errores_detalle"].copy()
                generar_reporte_excel(dl_dir, periodo, ips_nombre_actual, exitosas, errores)
                crear_zip_completo(job, dl_dir, periodo, ips_nombre_actual)
                zip_parcial_generado = True
                return

            # ========== PROCESAR FACTURAS PENDIENTES ==========
            # ── Función interna: procesar una lista de facturas en la sesión activa ──
            def _reconectar_data_frame():
                """Reconectar el data_frame si fue destruido al cerrar el visor."""
                nonlocal data_frame
                try:
                    data_frame.evaluate("() => true")
                    return True
                except:
                    pass
                for fr in page.frames:
                    try:
                        if fr.evaluate("() => /Pendiente de recibir Informaci|Devoluci[oó]n de entrada/i.test(document.body?.innerText || '')"):
                            data_frame = fr
                            return True
                    except:
                        continue
                return False

            def _procesar_lista(lista, intento_num):
                nonlocal zip_parcial_generado
                fallidas = []
                for idx, fac in enumerate(lista, 1):
                    if job["state"].get("stopping"):
                        log(job, "🛑 Proceso detenido por el usuario.")
                        if not zip_parcial_generado:
                            generar_zip_parcial(job)
                            zip_parcial_generado = True
                        return None  # señal de detención
                    log(job, f"[Intento {intento_num}][{idx}/{len(lista)}] Factura {fac['num']} ({fac['tipo']})...")
                    if not _reconectar_data_frame():
                        log(job, f"  ⚠️ data_frame perdido, reintentando en siguiente ciclo.", "warn")
                        fallidas.append(fac)
                        continue
                    try:
                        _download_factura(job, page, context, data_frame, fac, dl_dir, ips_nombre_actual)
                        with job["lock"]:
                            job["state"]["stats"]["descargadas"] += 1
                            job["state"]["stats"]["errores"] = max(0, job["state"]["stats"]["errores"] - 1) if intento_num > 1 else job["state"]["stats"]["errores"]
                        completadas.add(fac['num'])
                        guardar_progreso(
                            job, ips_dir, completadas,
                            nuevo_item={"factura": fac['num'], "tipo": fac.get('tipo')},
                            meta={
                                "identidad": job["state"].get("identidad"),
                                "ips": job["state"].get("ips_identity"),
                                "ejecucion_id": job["state"].get("ejecucion_id"),
                                "ips_nombre": ips_nombre_actual,
                                "aseguradora": "Previsora",
                                "periodo": job.get("periodo"),
                            },
                        )
                        log(job, f"  ✅ Descargada: {fac['num']}", "success")
                    except Exception as e:
                        error_msg = str(e)
                        es_definitivo = "no encontrada en el sistema" in error_msg
                        with job["lock"]:
                            if "No se pudo seleccionar el archivo" in error_msg:
                                if fac['tipo'] == 'auditada':
                                    error_msg = f"En la factura {fac['num']} no se encontró soporte Envios_D ni Carta de Objecion"
                                else:
                                    error_msg = f"En la factura {fac['num']} no se encontró soporte ActaDevolucion ni Carta de Objecion"
                        log(job, f"  ⚠️ Error intento {intento_num}: {error_msg}", "error")
                        _cerrar_traza_factura(page)
                        time.sleep(1)
                        if es_definitivo:
                            with job["lock"]:
                                nums_existentes = {er["factura"] for er in job["state"]["errores_detalle"]}
                                if fac['num'] not in nums_existentes:
                                    job["state"]["stats"]["errores"] += 1
                                    job["state"]["errores_detalle"].append({
                                        "factura": fac['num'],
                                        "estado": fac['estado'],
                                        "error": "Factura no existe en el sistema (No se encontraron registros)",
                                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        "captura": ""
                                    })
                                    try:
                                        historial_db.registrar_error_ejecucion(
                                            job["state"].get("ejecucion_id"), fac['num'], "previsora", "Previsora",
                                            job["state"].get("ips_identity"), "factura_inexistente",
                                            "Factura no existe en el sistema (No se encontraron registros)",
                                            "descarga_factura", False, intento_num, "fallida_definitiva",
                                        )
                                    except Exception as e:
                                        log(job, f"⚠️ No se pudo registrar el error de la factura {fac['num']} en el historial: {e}", "warn")
                            log(job, f"  ⏭️ Factura {fac['num']} marcada como inexistente — no se reintentará.", "warn")
                        else:
                            with job["lock"]:
                                if intento_num == 1:
                                    job["state"]["stats"]["errores"] += 1
                            fallidas.append(fac)
                return fallidas

            # ── Primer pase ──
            MAX_REINTENTOS = 8
            fallidas = _procesar_lista(facturas_pendientes, 1)

            # ── Reintentos automáticos ──
            if fallidas is None:  # usuario detuvo
                browser.close()
                return

            def _relogin_y_procesar(fallidas_lista, intento_num, espera_seg):
                """Cierra browser, espera, hace login completo y reintenta solo las fallidas."""
                nonlocal page, context, browser
                log(job, f"🔒 Cerrando browser para reinicio completo (intento {intento_num}/{MAX_REINTENTOS})...", "warn")
                try:
                    browser.close()
                except:
                    pass
                log(job, f"⏳ Esperando {espera_seg//60} minuto(s) antes de reiniciar...", "warn")
                for _ in range(espera_seg):
                    if job["state"].get("stopping"):
                        return None
                    time.sleep(1)
                log(job, f"🔄 Reiniciando browser y sesión (intento {intento_num}/{MAX_REINTENTOS})...", "warn")
                # Reutilizar el playwright (p) ya existente — no crear uno nuevo dentro del hilo
                browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
                context = browser.new_context(accept_downloads=True, viewport={"width": 1500, "height": 900})
                page = context.new_page()
                # Login completo
                page.goto("https://activa-it.net/Login.aspx", wait_until="networkidle", timeout=60000)
                page.fill('input[placeholder="Usuario"]', usuario)
                page.fill('input[placeholder="Contraseña"]', password)
                try:
                    checkbox = page.locator('input[type="checkbox"]').first
                    if not checkbox.is_checked():
                        checkbox.check()
                except:
                    pass
                page.click('button:has-text("Inicio de sesión"), input[value="Inicio de sesión"]')
                page.wait_for_url("**/Index.aspx", timeout=60000)
                try:
                    page.wait_for_selector("text=BI IPS, text=Inteligencia de Negocio", timeout=10000)
                except:
                    pass
                log(job, "✅ Sesión reiniciada correctamente.")
                # Navegar al módulo BI IPS
                for _ in range(3):
                    try:
                        page.click("text=Inteligencia de Negocio", timeout=8000)
                        time.sleep(1)
                        page.click("text=BI IPS", timeout=8000)
                        break
                    except:
                        time.sleep(2)
                try:
                    page.wait_for_load_state("networkidle", timeout=8000)
                except:
                    pass

                # ── Reconstruir el listado de facturas (data_frame nuevo) ──
                nonlocal data_frame
                log(job, "🔄 Reabriendo listado de facturas tras relogin...", "warn")
                target_frame_re = None
                for _ in range(120):
                    if job["state"].get("stopping"): return None
                    for fr in page.frames:
                        try:
                            if fr.evaluate(f"() => {{ for (const row of document.querySelectorAll('tr')) {{ const c = row.querySelectorAll('td'); if (c.length >= 3 && c[0].textContent.trim() === '{periodo}') return true; }} return false; }}"):
                                target_frame_re = fr
                                break
                        except:
                            continue
                    if target_frame_re:
                        break
                    time.sleep(0.5)
                if not target_frame_re:
                    log(job, "⚠️ No se pudo reabrir el período tras relogin.", "error")
                    return fallidas_lista

                try:
                    target_frame_re.evaluate(f"""
                        () => {{
                            for (const row of document.querySelectorAll('tr')) {{
                                const cells = row.querySelectorAll('td');
                                if (cells.length < 3) continue;
                                if (cells[0].textContent.trim() !== '{periodo}') continue;
                                const links = row.querySelectorAll('a');
                                if (links.length === 0) return;
                                links[0].scrollIntoView({{block: 'center'}});
                                links[0].click();
                                return;
                            }}
                        }}
                    """)
                except:
                    pass

                time.sleep(2)
                nuevo_data = None
                for _ in range(120):
                    if job["state"].get("stopping"): return None
                    for fr in page.frames:
                        try:
                            if fr.evaluate("() => /Pendiente de recibir Informaci|Devoluci[oó]n de entrada/i.test(document.body?.innerText || '')"):
                                nuevo_data = fr
                                break
                        except:
                            continue
                    if nuevo_data:
                        break
                    time.sleep(0.5)
                if nuevo_data:
                    data_frame = nuevo_data
                    log(job, "✅ Listado reabierto correctamente.")
                    time.sleep(2)
                else:
                    log(job, "⚠️ No se pudo reabrir el listado tras relogin.", "error")
                    return fallidas_lista

                # Limpiar errores anteriores y procesar
                nums_fallidas = {f['num'] for f in fallidas_lista}
                with job["lock"]:
                    job["state"]["errores_detalle"] = [e for e in job["state"]["errores_detalle"] if e["factura"] not in nums_fallidas]
                return _procesar_lista(fallidas_lista, intento_num)

            intento = 2
            while fallidas and intento <= MAX_REINTENTOS and not job["state"].get("stopping") and (reintentos_largos or intento <= 4):
                nums_fallidas = {f['num'] for f in fallidas}
                with job["lock"]:
                    job["state"]["errores_detalle"] = [e for e in job["state"]["errores_detalle"] if e["factura"] not in nums_fallidas]

                if intento <= 4:
                    # Intentos 2-4: reintento rápido sin cerrar browser
                    log(job, f"🔄 {len(fallidas)} factura(s) con error. Reintentando sin cerrar sesión ({intento}/{MAX_REINTENTOS})...", "warn")
                    time.sleep(3)
                    fallidas_nuevo = _procesar_lista(fallidas, intento)
                else:
                    # Intentos 5-7: esperar 5 min | Intento 8: esperar 10 min
                    espera = 600 if intento == MAX_REINTENTOS else 300
                    mins = espera // 60
                    log(job, f"🔄 {len(fallidas)} factura(s) persisten. Reiniciando sesión completa ({intento}/{MAX_REINTENTOS}) — espera {mins} min...", "warn")
                    fallidas_nuevo = _relogin_y_procesar(fallidas, intento, espera_seg=espera)

                if fallidas_nuevo is None:
                    try:
                        browser.close()
                    except:
                        pass
                    return
                fallidas = fallidas_nuevo
                intento += 1

            # ── Registrar errores persistentes al final ──
            if fallidas:
                if not reintentos_largos and intento > 4:
                    log(job, f"⏹️ {len(fallidas)} factura(s) con error tras los reintentos rápidos (4/4). Reintentos largos desactivados — no se intentaron los reinicios de sesión de 5-10 min.", "error")
                else:
                    log(job, f"⛔ {len(fallidas)} factura(s) no pudieron descargarse tras {MAX_REINTENTOS} intentos.", "error")
                with job["lock"]:
                    for fac in fallidas:
                        nums_existentes = {e["factura"] for e in job["state"]["errores_detalle"]}
                        if fac['num'] not in nums_existentes:
                            error_info = {
                                "factura": fac['num'],
                                "estado": fac['estado'],
                                "error": (
                                    "Sin descargar tras los 4 reintentos rápidos (reintentos largos desactivados)"
                                    if not reintentos_largos and intento > 4
                                    else f"Error persistente tras {MAX_REINTENTOS} intentos automáticos"
                                ),
                                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "captura": ""
                            }
                            try:
                                errores_dir = ips_dir / "Errores"
                                errores_dir.mkdir(parents=True, exist_ok=True)
                                cap_path = errores_dir / f"ERROR_PERSISTENTE_{fac['num']}.png"
                                page.screenshot(path=str(cap_path))
                                error_info["captura"] = str(cap_path)
                            except:
                                pass
                            job["state"]["errores_detalle"].append(error_info)
                            try:
                                historial_db.registrar_error_ejecucion(
                                    job["state"].get("ejecucion_id"), fac['num'], "previsora", "Previsora",
                                    job["state"].get("ips_identity"), "descarga_persistente",
                                    error_info["error"], "descarga_factura", False, MAX_REINTENTOS, "fallida",
                                )
                            except Exception as e:
                                log(job, f"⚠️ No se pudo registrar el error de la factura {fac['num']} en el historial: {e}", "warn")
                    job["state"]["stats"]["errores"] = len(job["state"]["errores_detalle"])

            browser.close()

            # ========== GENERAR EXCEL FINAL ==========
            with job["lock"]:
                exitosas = job["state"]["descargas_exitosas"].copy()
                errores = job["state"]["errores_detalle"].copy()
            excel_path = generar_reporte_excel(dl_dir, periodo, ips_nombre_actual, exitosas, errores)
            if excel_path:
                log(job, f"📊 Reporte Excel generado: {excel_path}")
            else:
                log(job, "⚠️ No se pudo generar el Excel (openpyxl no instalado o error).", "warn")

            # Excel aparte, solo de errores, para descarga automática desde
            # el frontend — nunca debe poder tumbar el proceso ya exitoso.
            try:
                nombre_excel_errores = generar_excel_solo_errores(dl_dir, ips_nombre_actual, errores)
                with job["lock"]:
                    job["state"]["errores_excel_url"] = (
                        f"/previsora/downloads/{periodo}/{nombre_excel_errores}" if nombre_excel_errores else None
                    )
            except Exception:
                pass

            # ========== ZIP FINAL (incluye Excel y Errores) ==========
            crear_zip_completo(job, dl_dir, periodo, ips_nombre_actual)
            zip_parcial_generado = True

            if fallidas:
                log(job, f"⚠️ Proceso completado con {len(fallidas)} factura(s) con error persistente. Ver Excel para detalle.")
            else:
                log(job, "🎉 Proceso completado sin errores.")

    except Exception as e:
        if not job["state"].get("stopping"):
            log(job, f"💥 Error crítico: {e}", "error")
            with job["lock"]:
                job["state"]["error"] = str(e)
        else:
            log(job, "Proceso detenido por el usuario.")
    finally:
        # Red de seguridad final: sin importar CÓMO se salió de la función
        # (éxito, error, reintentos agotados, o Detener manual del usuario),
        # si todavía no se generó ningún ZIP, se genera uno parcial aquí con
        # lo que se haya descargado hasta el momento.
        if not zip_parcial_generado:
            generar_zip_parcial(job)
        historial_db.cerrar_ejecucion(
            job["state"].get("ejecucion_id"), "error" if job["state"].get("error") else ("cancelada" if job["state"].get("stopping") else "completada"),
            total_detectadas=job["state"].get("stats", {}).get("total", 0),
            total_procesadas=job["state"].get("stats", {}).get("descargadas", 0) + job["state"].get("stats", {}).get("errores", 0),
            total_exitosas=job["state"].get("stats", {}).get("descargadas", 0),
            total_fallidas=job["state"].get("stats", {}).get("errores", 0),
        )
        with job["lock"]:
            if not job["state"].get("_lote_activo"):
                job["state"]["running"] = False
                job["state"]["finished"] = True
                job["state"]["stopping"] = False
                concurrency.registrar_fin("previsora")
        job["browser"] = None
        job["context"] = None
        job["dl_dir"] = None
        job["periodo"] = None
        job["ips_nombre"] = None

def run_automation_lote(job, usuario, password, periodos, download_path_base, reintentos_largos: bool = True):
    """Ejecuta run_automation secuencialmente para cada período de un rango
    (ej: Dic25-May26 -> primero Dic25 completo, luego Ene26, etc.), en vez de
    intentar buscar un período literal 'Dic25-May26' en el portal."""
    with job["lock"]:
        job["state"]["_lote_activo"] = True
    try:
        for idx, periodo in enumerate(periodos, 1):
            if job["state"].get("stopping"):
                log(job, "🛑 Rango detenido por el usuario antes de continuar con el siguiente período.", "warn")
                break
            log(job, f"{'='*50}")
            log(job, f"📅 Período {idx}/{len(periodos)} del rango: {periodo}")
            log(job, f"{'='*50}")
            dl_path_periodo = str(Path(download_path_base) / periodo)
            run_automation(job, usuario, password, periodo, dl_path_periodo, reintentos_largos)
    finally:
        with job["lock"]:
            job["state"]["_lote_activo"] = False
            job["state"]["running"] = False
            job["state"]["finished"] = True
            job["state"]["stopping"] = False
        concurrency.registrar_fin("previsora")
        log(job, "🏁 Rango de períodos finalizado.")

# ==================== RUTAS FLASK ====================
@bp.route("/")
def index():
    return render_template("previsora_index.html")

def _parsear_archivo_facturas(file_storage):
    """Parsea un CSV/XLSX con una columna 'factura' y devuelve los números limpios."""
    filename = file_storage.filename.lower()
    facturas = []
    if filename.endswith('.csv'):
        raw = file_storage.read()
        for enc in ('utf-8-sig', 'utf-8', 'latin-1', 'cp1252'):
            try:
                csv_text = raw.decode(enc)
                break
            except (UnicodeDecodeError, LookupError):
                continue
        else:
            csv_text = raw.decode('latin-1', errors='replace')
        reader = csv.DictReader(csv_text.splitlines())
        for row in reader:
            for col, val in row.items():
                if 'factura' in col.lower():
                    facturas.append(val.strip())
                    break
    elif filename.endswith(('.xls', '.xlsx')):
        if not EXCEL_AVAILABLE:
            raise ValueError("openpyxl no instalado")
        wb = openpyxl.load_workbook(BytesIO(file_storage.read()), data_only=True)
        ws = wb.active
        col_idx = None
        for cell in ws[1]:
            if cell.value and 'factura' in str(cell.value).lower():
                col_idx = cell.column
                break
        if col_idx is None:
            raise ValueError("No se encontró columna con 'factura'")
        for row in ws.iter_rows(min_row=2, values_only=True):
            val = row[col_idx - 1]
            if val:
                facturas.append(str(val).strip())
    else:
        raise ValueError("Formato no soportado. Use CSV o Excel")
    return [re.sub(r'\D', '', f) for f in facturas if re.sub(r'\D', '', f)]


@bp.route("/api/start", methods=["POST"])
def start_job():
    is_multipart = request.content_type and "multipart/form-data" in request.content_type
    data = request.form if is_multipart else (request.get_json(silent=True) or request.form or {})
    usuario = data.get("usuario", "").strip()
    password = data.get("password", "").strip()
    periodo_input = data.get("periodo", "").strip()
    custom_path = data.get("download_path", "").strip()
    reintentos_largos = str(data.get("reintentos_largos", "true")).lower() not in ("false", "off", "0", "no")
    identidad = validar_identidad(data.get("identidad"))
    if not identidad:
        return jsonify({"ok": False, "error": "Selecciona quién eres antes de iniciar el proceso.", "campo": "identidad"}), 400

    if not all([usuario, password, periodo_input]):
        return jsonify({"ok": False, "error": "Faltan campos requeridos"}), 400

    periodos = parse_periodo_input(periodo_input)
    if not periodos:
        return jsonify({"ok": False, "error": f"Formato de período inválido: '{periodo_input}'. Use MMMYY (ej: May26) o rango MMMYY-MMMYY"}), 400

    # Igual que en Bolívar: el archivo (si viene) se procesa AQUÍ MISMO, junto
    # con el arranque, para que el filtro de facturas sea atómico y no
    # dependa de un /api/upload previo que podría perderse si el contenedor
    # entró en "sleep" mientras tanto.
    archivo_facturas = request.files.get("file") if is_multipart else None
    facturas_nuevas = None
    if archivo_facturas and archivo_facturas.filename:
        try:
            facturas_nuevas = _parsear_archivo_facturas(archivo_facturas)
        except ValueError as e:
            return jsonify({"ok": False, "error": f"Archivo de facturas: {e}"}), 400
        except Exception as e:
            return jsonify({"ok": False, "error": f"Error al procesar archivo de facturas: {e}"}), 500

    manual_json = data.get("manual_entries", "").strip() if hasattr(data.get("manual_entries", ""), "strip") else ""
    if manual_json:
        try:
            manual_facturas = [str(f).strip() for f in json.loads(manual_json) if str(f).strip()]
        except Exception:
            return jsonify({"ok": False, "error": "Formato inválido en las facturas agregadas manualmente"}), 400
        if manual_facturas:
            facturas_nuevas = (facturas_nuevas or []) + manual_facturas

    empresa_id = resolve_empresa_id(usuario)
    job = get_or_create_job(empresa_id)
    nit_usuario = (re.search(r"(\d{9,12})", usuario) or [None, None])[1]
    ips_previa = MAPA_IPS.get(nit_usuario, "IPS_NO_IDENTIFICADA") if nit_usuario else "IPS_NO_IDENTIFICADA"
    decision_redescarga = str(data.get("decision_redescarga", ""))
    raw_selected = data.get("facturas_redescarga", "[]")
    try:
        facturas_redescarga = {str(v) for v in (json.loads(raw_selected) if isinstance(raw_selected, str) else (raw_selected or []))}
    except Exception:
        facturas_redescarga = set()
    confirmar_duplicados = str(data.get("confirmar_duplicados", "")).lower() in ("true", "1", "on", "si", "sí")
    try:
        if nit_usuario:
            repetidas = historial_db.buscar_ya_descargadas_por_nit("Previsora", nit_usuario, facturas_nuevas or [])
        else:
            repetidas = historial_db.buscar_ya_descargadas("Previsora", ips_previa, facturas_nuevas or [])
    except Exception as e:
        log(job, f"⚠️ No se pudo consultar el historial de duplicados: {e}", "warn")
        repetidas = {}
    if repetidas and not confirmar_duplicados:
        return jsonify({"ok": False, "requiere_confirmacion": True, "ya_descargadas": repetidas, "total_filas": len(facturas_nuevas or [])}), 409
    if repetidas and facturas_nuevas is not None:
        if decision_redescarga == "ninguna":
            facturas_nuevas = [f for f in facturas_nuevas if f not in repetidas]
        elif decision_redescarga == "seleccionadas":
            facturas_nuevas = [f for f in facturas_nuevas if f not in repetidas or f in facturas_redescarga]

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
            job["state"]["stats"] = {"total": 0, "descargadas": 0, "errores": 0}
            job["state"]["errores_detalle"] = []
            job["state"]["descargas_exitosas"] = []
            job["state"]["errores_excel_url"] = None
            if facturas_nuevas is not None:
                job["state"]["facturas_permitidas"] = facturas_nuevas

    if facturas_nuevas is not None:
        log(job, f"📄 Filtro de {len(facturas_nuevas)} facturas aplicado junto con el arranque.")
    elif job["state"].get("facturas_permitidas"):
        log(job, f"📄 Usando filtro de {len(job['state']['facturas_permitidas'])} facturas cargado previamente.")

    dl_path = custom_path if custom_path else str(DOWNLOAD_DIR / periodo_input)
    if custom_path:
        registro_rutas.registrar_ruta("Previsora", custom_path)

    job["state"]["identidad"] = identidad
    job["state"]["ips_identity"] = resolver(nit=nit_usuario, nombre_detectado=ips_previa)
    job["state"]["ips_nit"] = nit_usuario
    job["state"]["ejecucion_id"] = historial_db.iniciar_ejecucion(
        identidad, "previsora", "Previsora", job["state"]["ips_identity"], periodo_input, dl_path
    )

    concurrency.registrar_inicio("previsora")
    if len(periodos) > 1:
        log(job, f"📅 Procesando rango de {len(periodos)} períodos: {periodos[0]} → {periodos[-1]}")
        job["state"]["periodos_rango"] = periodos
        t = threading.Thread(target=run_automation_lote, args=(job, usuario, password, periodos, dl_path, reintentos_largos), daemon=True)
    else:
        job["state"]["periodos_rango"] = None
        t = threading.Thread(target=run_automation, args=(job, usuario, password, periodos[0], dl_path, reintentos_largos), daemon=True)
    t.start()
    return jsonify({"ok": True, "download_path": dl_path, "periodos_detectados": periodos, "empresa_id": empresa_id})

@bp.route("/api/stop", methods=["POST"])
def stop_job_route():
    empresa_id = resolve_empresa_id((request.get_json(silent=True) or {}).get("empresa_id") if request.is_json else request.args.get("empresa_id", ""))
    job = get_job_or_none(empresa_id)
    if not job:
        return jsonify({"ok": False, "message": "No hay proceso en ejecución"}), 400
    with job["lock"]:
        if not job["state"]["running"]:
            return jsonify({"ok": False, "message": "No hay proceso en ejecución"}), 400
    stop_job(job)
    return jsonify({"ok": True, "message": "Deteniendo proceso..."})


@bp.route("/api/duplicate-decision", methods=["POST"])
def duplicate_decision_route():
    data = request.get_json(silent=True) or request.form or {}
    job = get_job_or_none(resolve_empresa_id(data.get("empresa_id", "")))
    if not job or not job["state"].get("duplicados_pendientes"):
        return jsonify({"ok": False, "error": "No hay una decisión de duplicados pendiente"}), 409
    decision = data.get("decision")
    if decision not in ("ninguna", "todas", "seleccionadas"):
        return jsonify({"ok": False, "error": "Decisión inválida"}), 400
    with job["lock"]:
        job["state"]["decision_redescarga"] = decision
        job["state"]["facturas_redescarga"] = [str(v) for v in data.get("selected", [])]
        job["state"]["duplicate_event"].set()
    return jsonify({"ok": True})

@bp.route("/api/reset", methods=["POST"])
def reset_job_route():
    data = request.get_json(silent=True) or request.form or {}
    periodo = data.get("periodo", "").strip()
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

    if periodo:
        periodo_dir = DOWNLOAD_DIR / periodo
        if periodo_dir.exists():
            for progreso_file in periodo_dir.glob("*/progreso.json"):
                try:
                    n_migradas = historial_db.migrar_progreso_json(progreso_file.parent)
                    progreso_file.unlink()
                    log(job, f"🗑️ Progreso eliminado: {progreso_file} ({n_migradas} factura(s) archivadas en el historial)")
                except Exception as e:
                    log(job, f"⚠️ Error al borrar {progreso_file}: {e}", "warn")
        else:
            log(job, f"⚠️ No existe la carpeta del período '{periodo}'.", "warn")
    else:
        log(job, "⚠️ No se especificó período, no se borró progreso.", "warn")

    if job:
        reset_state(job)
    return jsonify({"ok": True, "message": "Estado reiniciado y progreso eliminado."})

@bp.route("/api/status")
def get_status():
    empresa_id_raw = request.args.get("empresa_id", "")
    if empresa_id_raw:
        empresa_id = resolve_empresa_id(empresa_id_raw)
        job = get_job_or_none(empresa_id)
        if not job:
            return jsonify({"running": False, "finished": False, "error": None,
                             "stats": {"total": 0, "descargadas": 0, "errores": 0}, "logs": []})
        with job["lock"]:
            return jsonify({
                "running": job["state"]["running"],
                "finished": job["state"]["finished"],
                "error": job["state"]["error"],
                "stats": job["state"]["stats"],
                "logs": job["state"]["logs"][-200:],
                "duplicados_pendientes": job["state"].get("duplicados_pendientes"),
                "errores_excel_url": job["state"].get("errores_excel_url"),
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
        return jsonify({"running": len(activos) > 0, "finished": False, "error": None,
                         "stats": stats, "logs": [], "procesos": procesos})

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
    periodo = request.args.get("periodo", "")
    folder = DOWNLOAD_DIR / periodo if periodo else DOWNLOAD_DIR
    files = []
    if folder.exists():
        for f in sorted(folder.iterdir()):
            if f.is_file():
                files.append({"name": f.name, "size": f.stat().st_size, "path": str(f), "periodo": periodo})
    return jsonify({"files": files})

@bp.route("/api/files", methods=["DELETE"])
def delete_all_files():
    import shutil
    periodo = request.args.get("periodo", "")
    folder = DOWNLOAD_DIR / periodo if periodo else DOWNLOAD_DIR
    if not folder.exists():
        return jsonify({"ok": True, "message": "No hay archivos que eliminar"})
    try:
        eliminados = 0
        for item in list(folder.iterdir()):
            if item.is_file() and item.name != "progreso.json":
                item.unlink()
                eliminados += 1
            elif item.is_dir():
                # Dentro de subcarpetas: borrar archivos pero conservar progreso.json
                for sub in list(item.iterdir()):
                    if sub.is_file() and sub.name != "progreso.json":
                        sub.unlink()
                        eliminados += 1
                # Si la subcarpeta quedó vacía (o solo tiene progreso), dejarla
        log(None, f"🗑️ Soportes eliminados: {eliminados} archivo(s) en '{folder}' (progreso conservado)")
        return jsonify({"ok": True, "message": f"Se eliminaron {eliminados} soporte(s). El progreso se conservó.", "eliminados": eliminados})
    except Exception as e:
        log(None, f"⚠️ Error al eliminar soportes: {e}", "error")
        return jsonify({"ok": False, "error": str(e)}), 500

@bp.route("/downloads/<path:filename>")
def download_file(filename):
    return send_from_directory(DOWNLOAD_DIR, filename, as_attachment=True)

@bp.route("/api/periodos")
def get_periodos():
    periodos = []
    for d in DOWNLOAD_DIR.iterdir():
        if d.is_dir():
            count = len(list(d.glob("**/*.pdf")))
            periodos.append({"name": d.name, "count": count})
    return jsonify({"periodos": sorted(periodos, key=lambda x: x["name"], reverse=True)})

@bp.route("/api/upload", methods=["POST"])
def upload_facturas():
    empresa_id = resolve_empresa_id(request.form.get("usuario", ""))
    job = get_or_create_job(empresa_id)
    if 'file' not in request.files:
        return jsonify({"ok": False, "error": "No se envió ningún archivo"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"ok": False, "error": "Archivo vacío"}), 400

    try:
        filename = file.filename.lower()
        facturas = []
        if filename.endswith('.csv'):
            raw = file.read()
            # Detectar encoding: utf-8-sig cubre BOM, latin-1 cubre Windows
            for enc in ('utf-8-sig', 'utf-8', 'latin-1', 'cp1252'):
                try:
                    csv_text = raw.decode(enc)
                    break
                except (UnicodeDecodeError, LookupError):
                    continue
            else:
                csv_text = raw.decode('latin-1', errors='replace')
            reader = csv.DictReader(csv_text.splitlines())
            for row in reader:
                for col, val in row.items():
                    if 'factura' in col.lower():
                        facturas.append(val.strip())
                        break
        elif filename.endswith(('.xls', '.xlsx')):
            if not EXCEL_AVAILABLE:
                return jsonify({"ok": False, "error": "openpyxl no instalado"}), 500
            wb = openpyxl.load_workbook(BytesIO(file.read()), data_only=True)
            ws = wb.active
            col_idx = None
            for cell in ws[1]:
                if cell.value and 'factura' in str(cell.value).lower():
                    col_idx = cell.column
                    break
            if col_idx is None:
                return jsonify({"ok": False, "error": "No se encontró columna con 'factura'"}), 400
            for row in ws.iter_rows(min_row=2, values_only=True):
                val = row[col_idx-1]
                if val:
                    facturas.append(str(val).strip())
        else:
            return jsonify({"ok": False, "error": "Formato no soportado. Use CSV o Excel"}), 400

        facturas_limpias = [re.sub(r'\D', '', f) for f in facturas if re.sub(r'\D', '', f)]
        if not facturas_limpias:
            return jsonify({"ok": False, "error": "No se encontraron números de factura válidos"}), 400

        with job["lock"]:
            job["state"]["facturas_permitidas"] = facturas_limpias
        log(job, f"📄 Se cargaron {len(facturas_limpias)} facturas desde el archivo.")
        return jsonify({"ok": True, "count": len(facturas_limpias), "facturas": facturas_limpias[:10]})
    except Exception as e:
        return jsonify({"ok": False, "error": f"Error al procesar archivo: {str(e)}"}), 500

@bp.route("/api/progreso")
def get_progreso():
    periodo = request.args.get("periodo", "")
    ips = request.args.get("ips", "")
    if not periodo:
        return jsonify({"ok": False, "error": "Se requiere el parámetro 'periodo'"}), 400

    periodo_dir = DOWNLOAD_DIR / periodo
    if not periodo_dir.exists():
        return jsonify({"ok": True, "completadas": [], "mensaje": "No hay datos para este período"})

    if ips:
        ips_dir = periodo_dir / ips
        if not ips_dir.exists():
            return jsonify({"ok": False, "error": f"No existe la IPS '{ips}'"}), 404
    else:
        posibles = list(periodo_dir.iterdir())
        if not posibles:
            return jsonify({"ok": True, "completadas": [], "mensaje": "No hay subcarpetas de IPS"})
        ips_dir = None
        for d in posibles:
            if d.is_dir() and (d / "progreso.json").exists():
                ips_dir = d
                break
        if not ips_dir:
            ips_dir = posibles[0] if posibles[0].is_dir() else None
        if not ips_dir:
            return jsonify({"ok": True, "completadas": [], "mensaje": "No se encontró carpeta de IPS"})
        ips = ips_dir.name

    progreso_path = ips_dir / "progreso.json"
    if not progreso_path.exists():
        return jsonify({"ok": True, "completadas": [], "ips": ips, "mensaje": "Aún no hay facturas completadas"})

    try:
        with open(progreso_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        completadas = data.get("completadas", [])
        return jsonify({"ok": True, "completadas": completadas, "cantidad": len(completadas), "ips": ips, "actualizado": data.get("actualizado", "")})
    except Exception as e:
        return jsonify({"ok": False, "error": f"Error al leer progreso: {str(e)}"}), 500

@bp.route("/api/exportar_progreso")
def exportar_progreso_excel():
    periodo = request.args.get("periodo", "")
    if not periodo:
        return jsonify({"ok": False, "error": "Se requiere el parámetro 'periodo'"}), 400
    if not EXCEL_AVAILABLE:
        return jsonify({"ok": False, "error": "openpyxl no instalado"}), 500
    periodo_dir = DOWNLOAD_DIR / periodo
    if not periodo_dir.exists():
        return jsonify({"ok": False, "error": f"No existe la carpeta del período '{periodo}'"}), 404
    progreso_files = list(periodo_dir.glob("*/progreso.json"))
    if not progreso_files:
        return jsonify({"ok": False, "error": f"No se encontró progreso.json para el período '{periodo}'"}), 404
    progreso_path = progreso_files[0]
    ips_nombre = progreso_path.parent.name
    try:
        with open(progreso_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        completadas = data.get("completadas", [])
        actualizado = data.get("actualizado", "")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Facturas completadas"
        ws.append(["N° Factura", "Fecha de completado"])
        for factura in completadas:
            ws.append([factura, actualizado])
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 30
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"progreso_facturas_{periodo}_{ips_nombre}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify({"ok": False, "error": f"Error al generar Excel: {str(e)}"}), 500

if __name__ == "__main__":
    print("\n" + "=" * 55)
    print("  🏥 Activa IT — Descargador de Cartas Glosa")
    print("  🔷 Previsora SOAT (con nombres de IPS forzados desde el mapa)")
    print("=" * 55)
    print(f"  📂 Carpeta de descargas: {DOWNLOAD_DIR}")
    print(f"  🌐 Puerto: {port}")
    print("=" * 55 + "\n")
    app.run(host="0.0.0.0", port=port, debug=False)