"""
Panel de administrador -- capa de solo lectura + disparo sobre lo que
YA existe en cada bot (jobs, stop_job, progreso.json, ZIPs). No inventa
ningún mecanismo nuevo de descarga ni toca la lógica de los 4 bots.
"""
import os
import time
from pathlib import Path
from functools import wraps
from flask import Blueprint, render_template, request, jsonify, session, redirect

bp = Blueprint("admin", __name__)

ADMIN_USER = os.environ.get("ADMIN_USER", "Administrador")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "P@ssw0rd.2026")


def _ruta_bajo_base(ruta: str, bases) -> bool:
    """Comprueba si `ruta` está bajo alguna de las carpetas base (robusto a / vs \\)."""
    try:
        rp = Path(ruta).resolve()
    except Exception:
        return False
    for base in bases:
        if not base:
            continue
        try:
            bp = Path(base).resolve()
            rp.relative_to(bp)
            return True
        except (ValueError, OSError):
            # Fallback: comparación normalizada por string
            rs = str(rp).replace("\\", "/").lower()
            bs = str(bp).replace("\\", "/").lower()
            if rs == bs or rs.startswith(bs.rstrip("/") + "/"):
                return True
    return False



def _requiere_login(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("admin_ok"):
            return jsonify({"ok": False, "error": "No autenticado"}), 401
        return f(*args, **kwargs)
    return wrapper


def _bots_disponibles():
    """Import perezoso para evitar ciclos de import con app.py."""
    from bots import bolivar, previsora, mundial, estado_sura
    return {
        "bolivar": {"modulo": bolivar, "nombre": "Bolívar", "multi_empresa": False, "logo": "logo_bolivar.png"},
        "previsora": {"modulo": previsora, "nombre": "Previsora", "multi_empresa": False, "logo": "logo_previsora.png"},
        "mundial": {"modulo": mundial, "nombre": "Mundial", "multi_empresa": False, "logo": "logo_mundial.png"},
        "estado_sura": {"modulo": estado_sura, "nombre": "Estado/Sura", "multi_empresa": True, "logo": "logo_estado.png"},
    }


def _carpetas_escaneo(bot_id, info):
    """
    Las carpetas REALES y EXCLUSIVAS de cada bot para escanear
    progreso.json / *.zip. OJO: el DOWNLOAD_DIR de estado_sura es la
    carpeta 'downloads/' raíz (la comparte con los otros 3 bots, porque
    adentro tiene sus propias subcarpetas 'estado/' y 'sura/') -- si se
    escaneara ese DOWNLOAD_DIR completo, aparecería duplicado TODO lo de
    los otros bots también bajo "Estado/Sura". Por eso aquí se usan sus
    2 subcarpetas reales en vez del DOWNLOAD_DIR crudo.
    """
    modulo = info["modulo"]
    base = getattr(modulo, "DOWNLOAD_DIR", None)
    if not base:
        return []
    if info["multi_empresa"]:
        return [Path(base) / "estado", Path(base) / "sura"]
    return [Path(base)]


def _todas_las_carpetas_validas():
    """Lista plana de todas las carpetas exclusivas de los 4 bots, para validar rutas."""
    carpetas = []
    for bot_id, info in _bots_disponibles().items():
        carpetas.extend(_carpetas_escaneo(bot_id, info))
    return [str(c) for c in carpetas]


@bp.route("/admin")
def admin_home():
    if not session.get("admin_ok"):
        return render_template("admin_login.html")
    return render_template("admin_dashboard.html")


@bp.route("/admin/login", methods=["POST"])
def admin_login():
    data = request.get_json(silent=True) or request.form
    usuario = str(data.get("usuario", "")).strip()
    clave = str(data.get("password", "")).strip()
    if usuario == ADMIN_USER and clave == ADMIN_PASSWORD:
        session["admin_ok"] = True
        session.permanent = True
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Usuario o contraseña incorrectos"}), 401


@bp.route("/admin/logout", methods=["POST"])
def admin_logout():
    session.pop("admin_ok", None)
    return jsonify({"ok": True})


@bp.route("/admin/api/jobs")
@_requiere_login
def admin_jobs():
    """Todos los procesos activos, ahora mismo, en los 4 bots a la vez."""
    resultado = []
    ahora = time.time()
    for bot_id, info in _bots_disponibles().items():
        modulo = info["modulo"]
        jobs = getattr(modulo, "jobs", {})
        for clave, job in list(jobs.items()):
            with job["lock"]:
                estado = job["state"]
                if not estado.get("running"):
                    continue
                inicio = estado.get("_inicio_ts")
                minutos = round((ahora - inicio) / 60, 1) if inicio else None
                stats = estado.get("stats", {}) or {}
                total = stats.get("total") or 0
                descargadas = stats.get("descargadas") or 0
                porcentaje = round((descargadas / total) * 100, 1) if total else 0
                logo = info["logo"]
                if info["multi_empresa"]:
                    empresa_real = clave.split(":")[0]
                    logo = "logo_sura.png" if empresa_real == "sura" else "logo_estado.png"
                resultado.append({
                    "bot": bot_id,
                    "bot_nombre": info["nombre"],
                    "logo": logo,
                    "clave_job": clave,
                    "identidad": estado.get("identidad"),
                    "stats": stats,
                    "porcentaje": porcentaje,
                    "minutos_corriendo": minutos,
                    "atascado": bool(minutos and minutos > 20),
                })
    return jsonify({"ok": True, "jobs": resultado})


@bp.route("/admin/api/stop", methods=["POST"])
@_requiere_login
def admin_stop():
    data = request.get_json(silent=True) or {}
    bot_id, clave_job = data.get("bot"), data.get("clave_job")
    bots = _bots_disponibles()
    if bot_id not in bots:
        return jsonify({"ok": False, "error": "Bot no reconocido"}), 400
    modulo = bots[bot_id]["modulo"]
    job = getattr(modulo, "jobs", {}).get(clave_job)
    if not job:
        return jsonify({"ok": False, "error": "Ese proceso ya no existe"}), 404
    try:
        if bots[bot_id]["multi_empresa"]:
            empresa = clave_job.split(":")[0]
            modulo.stop_job(job, empresa)
        else:
            modulo.stop_job(job)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@bp.route("/admin/api/progresos")
@_requiere_login
def admin_progresos():
    """progreso.json en cada carpeta de descarga, con su antigüedad."""
    resultado = []
    ahora = time.time()
    for bot_id, info in _bots_disponibles().items():
        for base in _carpetas_escaneo(bot_id, info):
            if not base.exists():
                continue
            for p in base.rglob("progreso.json"):
                try:
                    edad_horas = round((ahora - p.stat().st_mtime) / 3600, 1)
                    resultado.append({
                        "bot": bot_id, "bot_nombre": info["nombre"], "logo": info["logo"],
                        "ruta": str(p), "carpeta": str(p.parent.relative_to(base)),
                        "edad_horas": edad_horas,
                    })
                except Exception:
                    continue
    resultado.sort(key=lambda r: -r["edad_horas"])
    return jsonify({"ok": True, "progresos": resultado})


@bp.route("/admin/api/borrar-progreso", methods=["POST"])
@_requiere_login
def admin_borrar_progreso():
    data = request.get_json(silent=True) or {}
    ruta = data.get("ruta", "")
    if not ruta or "progreso.json" not in ruta:
        return jsonify({"ok": False, "error": "Ruta inválida"}), 400
    p = Path(ruta)
    bases_validas = _todas_las_carpetas_validas()
    if not _ruta_bajo_base(ruta, bases_validas):
        return jsonify({"ok": False, "error": "Ruta fuera de las carpetas de descarga permitidas"}), 400
    try:
        if p.exists():
            p.unlink()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@bp.route("/admin/api/archivos")
@_requiere_login
def admin_archivos():
    """Todos los ZIP generados, en las 4 carpetas de descarga."""
    resultado = []
    for bot_id, info in _bots_disponibles().items():
        for base in _carpetas_escaneo(bot_id, info):
            if not base.exists():
                continue
            for p in base.rglob("*.zip"):
                try:
                    resultado.append({
                        "bot": bot_id, "bot_nombre": info["nombre"], "logo": info["logo"],
                        "ruta": str(p), "nombre": p.name,
                        "kb": round(p.stat().st_size / 1024, 1),
                        "modificado": p.stat().st_mtime,
                    })
                except Exception:
                    continue
    resultado.sort(key=lambda r: -r["modificado"])
    return jsonify({"ok": True, "archivos": resultado})


@bp.route("/admin/api/descargar")
@_requiere_login
def admin_descargar():
    from flask import send_file
    ruta = request.args.get("ruta", "")
    if not ruta or not ruta.endswith(".zip"):
        return jsonify({"ok": False, "error": "Ruta inválida"}), 400
    bases_validas = _todas_las_carpetas_validas()
    if not _ruta_bajo_base(ruta, bases_validas):
        return jsonify({"ok": False, "error": "Ruta fuera de las carpetas de descarga permitidas"}), 400
    p = Path(ruta)
    if not p.exists():
        return jsonify({"ok": False, "error": "El archivo ya no existe"}), 404
    return send_file(str(p), as_attachment=True, download_name=p.name)


@bp.route("/admin/api/borrar-archivo", methods=["POST"])
@_requiere_login
def admin_borrar_archivo():
    data = request.get_json(silent=True) or {}
    ruta = data.get("ruta", "")
    if not ruta or not ruta.endswith(".zip"):
        return jsonify({"ok": False, "error": "Ruta inválida"}), 400
    p = Path(ruta)
    bases_validas = _todas_las_carpetas_validas()
    if not _ruta_bajo_base(ruta, bases_validas):
        return jsonify({"ok": False, "error": "Ruta fuera de las carpetas de descarga permitidas"}), 400
    try:
        if p.exists():
            p.unlink()
            # También intentar borrar parciales asociados al mismo prefijo si aplica
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
