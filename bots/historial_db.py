"""Persistencia historica compatible con SQLite local y PostgreSQL/Neon."""
import json
import os
import sqlite3
import threading
from datetime import datetime, timezone
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
_USA_POSTGRES = bool(DATABASE_URL)
if _USA_POSTGRES:
    try:
        import psycopg2
        import psycopg2.extras
    except Exception as exc:
        raise RuntimeError("DATABASE_URL requiere psycopg2 disponible.") from exc
DB_PATH = Path(os.environ.get("HISTORIAL_DB_PATH", str(BASE_DIR / "downloads" / "historial.db")))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
_lock = threading.RLock()


def _id_type():
    return "SERIAL PRIMARY KEY" if _USA_POSTGRES else "INTEGER PRIMARY KEY AUTOINCREMENT"


def _marcador():
    return "%s" if _USA_POSTGRES else "?"


def _conectar():
    if _USA_POSTGRES:
        return psycopg2.connect(DATABASE_URL)
    conn = sqlite3.connect(str(DB_PATH), timeout=10, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _asegurar_columnas_descargas(cur):
    columnas = {
        "ips_id": "BIGINT", "ips_nit": "TEXT", "nombre_detectado": "TEXT",
        "metodo_identificacion": "TEXT", "primera_descarga": "TEXT",
        "ultima_descarga": "TEXT", "veces_procesada": "INTEGER NOT NULL DEFAULT 1",
        "ultima_ejecucion_id": "BIGINT", "ultimo_resultado": "TEXT",
    }
    if _USA_POSTGRES:
        for nombre, tipo in columnas.items():
            cur.execute(f"ALTER TABLE descargas ADD COLUMN IF NOT EXISTS {nombre} {tipo}")
        return
    existentes = {row[1] for row in cur.execute("PRAGMA table_info(descargas)").fetchall()}
    for nombre, tipo in columnas.items():
        if nombre not in existentes:
            cur.execute(f"ALTER TABLE descargas ADD COLUMN {nombre} {tipo}")


def inicializar():
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor()
            if _USA_POSTGRES:
                cur.execute("""CREATE TABLE IF NOT EXISTS descargas (
                    id SERIAL PRIMARY KEY, aseguradora TEXT NOT NULL, ips_nombre TEXT NOT NULL,
                    factura TEXT NOT NULL, siniestro TEXT, periodo TEXT, identidad TEXT,
                    fecha_descarga TEXT NOT NULL, fecha_migrado TEXT NOT NULL,
                    UNIQUE(aseguradora, ips_nombre, factura))""")
            else:
                cur.execute("""CREATE TABLE IF NOT EXISTS descargas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, aseguradora TEXT NOT NULL,
                    ips_nombre TEXT NOT NULL, factura TEXT NOT NULL, siniestro TEXT,
                    periodo TEXT, identidad TEXT, fecha_descarga TEXT NOT NULL,
                    fecha_migrado TEXT NOT NULL, UNIQUE(aseguradora, ips_nombre, factura))""")
            _asegurar_columnas_descargas(cur)
            cur.execute("CREATE INDEX IF NOT EXISTS idx_ips ON descargas(aseguradora, ips_nombre)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_factura ON descargas(aseguradora, ips_nombre, factura)")
            cur.execute(f"""CREATE TABLE IF NOT EXISTS ips (
                id {_id_type()}, nit TEXT NOT NULL UNIQUE, nombre_estandar TEXT NOT NULL,
                razon_social TEXT, activo INTEGER NOT NULL DEFAULT 1,
                fecha_creacion TEXT NOT NULL, fecha_actualizacion TEXT NOT NULL)""")
            cur.execute("""CREATE TABLE IF NOT EXISTS ips_alias (
                id %s, ips_id INTEGER NOT NULL, nombre_alias TEXT NOT NULL UNIQUE)""" %
                ("SERIAL PRIMARY KEY" if _USA_POSTGRES else "INTEGER PRIMARY KEY AUTOINCREMENT"))
            cur.execute(f"""CREATE TABLE IF NOT EXISTS ejecuciones (
                id {_id_type()}, persona TEXT NOT NULL, inicio TEXT NOT NULL, fin TEXT,
                estado TEXT NOT NULL, bot TEXT NOT NULL, aseguradora TEXT, ips_id INTEGER,
                ips_nit TEXT, ips_nombre_estandar TEXT, nombre_detectado TEXT,
                metodo_identificacion TEXT, periodo TEXT, ruta_destino TEXT, entorno TEXT,
                total_detectadas INTEGER DEFAULT 0, total_procesadas INTEGER DEFAULT 0,
                total_exitosas INTEGER DEFAULT 0, total_fallidas INTEGER DEFAULT 0,
                total_omitidas INTEGER DEFAULT 0, total_redescargadas INTEGER DEFAULT 0,
                decision_redescarga TEXT, facturas_previas TEXT, facturas_seleccionadas TEXT,
                facturas_descartadas TEXT)""")
            cur.execute(f"""CREATE TABLE IF NOT EXISTS ejecucion_facturas (
                id {_id_type()}, ejecucion_id INTEGER NOT NULL, factura TEXT NOT NULL,
                ips_id INTEGER, ips_nit TEXT, aseguradora TEXT, bot TEXT, periodo TEXT,
                estado TEXT NOT NULL, redescargada INTEGER NOT NULL DEFAULT 0,
                fecha_inicio TEXT, fecha_fin TEXT, archivo TEXT, error TEXT)""")
            cur.execute(f"""CREATE TABLE IF NOT EXISTS errores_ejecucion (
                id {_id_type()}, ejecucion_id INTEGER NOT NULL, factura TEXT, bot TEXT,
                aseguradora TEXT, ips_id INTEGER, ips_nit TEXT, tipo_error TEXT NOT NULL,
                mensaje TEXT NOT NULL, etapa TEXT, recuperable INTEGER, intento INTEGER,
                resultado_final TEXT, fecha TEXT NOT NULL)""")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_ejec_factura ON ejecucion_facturas(ips_nit, factura)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_ejecuciones_persona ON ejecuciones(persona, inicio)")
            conn.commit()
        finally:
            conn.close()


# Neon no se sustituye silenciosamente por SQLite cuando DATABASE_URL existe.
inicializar()


def _resolver_ips(nit=None, nombre=None):
    from .catalogo_ips import resolver
    return resolver(nit=nit, nombre_detectado=nombre)


def registrar_ips(ips):
    if not ips or not ips.get("nit"):
        return None
    ahora = datetime.now(timezone.utc).isoformat()
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor()
            if _USA_POSTGRES:
                cur.execute("""INSERT INTO ips (nit,nombre_estandar,fecha_creacion,fecha_actualizacion)
                    VALUES (%s,%s,%s,%s)
                    ON CONFLICT(nit) DO UPDATE SET nombre_estandar=EXCLUDED.nombre_estandar,
                    fecha_actualizacion=EXCLUDED.fecha_actualizacion RETURNING id""",
                    (ips["nit"], ips["nombre_estandar"], ahora, ahora))
                result = cur.fetchone()[0]
            else:
                cur.execute("INSERT OR IGNORE INTO ips (nit,nombre_estandar,fecha_creacion,fecha_actualizacion) VALUES (?,?,?,?)",
                            (ips["nit"], ips["nombre_estandar"], ahora, ahora))
                cur.execute("UPDATE ips SET nombre_estandar=?,fecha_actualizacion=? WHERE nit=?",
                            (ips["nombre_estandar"], ahora, ips["nit"]))
                cur.execute("SELECT id FROM ips WHERE nit=?", (ips["nit"],))
                result = cur.fetchone()[0]
            conn.commit()
            return result
        finally:
            conn.close()


def registrar_descargas(aseguradora, ips_nombre, periodo, identidad, items, ips=None):
    if not items:
        return 0
    ahora = datetime.now(timezone.utc).isoformat()
    ips = ips or _resolver_ips(nombre=ips_nombre)
    ips_id = registrar_ips(ips)
    filas = [(aseguradora, ips_nombre, str(item["factura"]), item.get("siniestro"), periodo,
              identidad, item.get("fecha_descarga", ahora), ahora, ips_id, ips.get("nit"),
              ips.get("nombre_detectado"), ips.get("metodo"), item.get("fecha_descarga", ahora),
              item.get("fecha_descarga", ahora)) for item in items]
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor()
            if _USA_POSTGRES:
                psycopg2.extras.execute_values(cur, """INSERT INTO descargas
                    (aseguradora,ips_nombre,factura,siniestro,periodo,identidad,fecha_descarga,fecha_migrado,
                     ips_id,ips_nit,nombre_detectado,metodo_identificacion,primera_descarga,ultima_descarga)
                    VALUES %s ON CONFLICT (aseguradora,ips_nombre,factura) DO UPDATE SET
                    siniestro=EXCLUDED.siniestro,periodo=EXCLUDED.periodo,identidad=EXCLUDED.identidad,
                    fecha_descarga=EXCLUDED.fecha_descarga,fecha_migrado=EXCLUDED.fecha_migrado,
                    ips_id=EXCLUDED.ips_id,ips_nit=EXCLUDED.ips_nit,nombre_detectado=EXCLUDED.nombre_detectado,
                    metodo_identificacion=EXCLUDED.metodo_identificacion,ultima_descarga=EXCLUDED.ultima_descarga""", filas)
            else:
                cur.executemany("""INSERT OR REPLACE INTO descargas
                    (aseguradora,ips_nombre,factura,siniestro,periodo,identidad,fecha_descarga,fecha_migrado,
                     ips_id,ips_nit,nombre_detectado,metodo_identificacion,primera_descarga,ultima_descarga)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", filas)
            conn.commit()
            return len(filas)
        finally:
            conn.close()


def buscar_ya_descargadas(aseguradora, ips_nombre, facturas):
    if not facturas:
        return {}
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador()
            placeholders = ",".join([m] * len(facturas))
            cur.execute(f"SELECT factura,fecha_descarga FROM descargas WHERE aseguradora={m} AND ips_nombre={m} AND factura IN ({placeholders})",
                        [aseguradora, ips_nombre] + [str(f) for f in facturas])
            return {str(a): b for a, b in cur.fetchall()}
        finally:
            conn.close()


def buscar_ya_descargadas_por_nit(aseguradora, ips_nit, facturas):
    if not ips_nit or not facturas:
        return {}
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador(); placeholders = ",".join([m] * len(facturas))
            cur.execute(f"SELECT factura,fecha_descarga FROM descargas WHERE aseguradora={m} AND ips_nit={m} AND factura IN ({placeholders})",
                        [aseguradora, ips_nit] + [str(f) for f in facturas])
            return {str(a): b for a, b in cur.fetchall()}
        finally:
            conn.close()


def registrar_facturas_ejecucion(ejecucion_id, bot, aseguradora, ips, periodo, items, redescargadas=None):
    if not ejecucion_id or not items:
        return 0
    ips_id = registrar_ips(ips) if ips else None
    ahora = datetime.now(timezone.utc).isoformat(); redescargadas = set(redescargadas or [])
    filas = [(ejecucion_id, str(item.get("factura")), ips_id, (ips or {}).get("nit"), aseguradora, bot, periodo,
              item.get("estado", "exitosa"), str(item.get("factura")) in redescargadas,
              item.get("fecha_inicio", ahora), item.get("fecha_fin", ahora), item.get("archivo"), item.get("error")) for item in items]
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor()
            if _USA_POSTGRES:
                psycopg2.extras.execute_values(cur, """INSERT INTO ejecucion_facturas
                    (ejecucion_id,factura,ips_id,ips_nit,aseguradora,bot,periodo,estado,redescargada,fecha_inicio,fecha_fin,archivo,error)
                    VALUES %s""", filas)
            else:
                cur.executemany("""INSERT INTO ejecucion_facturas
                    (ejecucion_id,factura,ips_id,ips_nit,aseguradora,bot,periodo,estado,redescargada,fecha_inicio,fecha_fin,archivo,error)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", filas)
            conn.commit(); return len(filas)
        finally:
            conn.close()


def registrar_error_ejecucion(ejecucion_id, factura, bot, aseguradora, ips, tipo_error, mensaje, etapa, recuperable, intento, resultado_final):
    if not ejecucion_id:
        return False
    values = (ejecucion_id, factura, bot, aseguradora, (ips or {}).get("nit"), tipo_error, mensaje, etapa,
              recuperable, intento, resultado_final, datetime.now(timezone.utc).isoformat())
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador()
            cur.execute("INSERT INTO errores_ejecucion (ejecucion_id,factura,bot,aseguradora,ips_nit,tipo_error,mensaje,etapa,recuperable,intento,resultado_final,fecha) VALUES (" + ",".join([m] * len(values)) + ")", values)
            conn.commit(); return True
        finally:
            conn.close()


def iniciar_ejecucion(persona, bot, aseguradora, ips, periodo, ruta_destino):
    ahora = datetime.now(timezone.utc).isoformat(); ips_id = registrar_ips(ips) if ips else None
    values = (persona, ahora, "iniciada", bot, aseguradora, ips_id, (ips or {}).get("nit"),
              (ips or {}).get("nombre_estandar"), (ips or {}).get("nombre_detectado"),
              (ips or {}).get("metodo"), periodo, ruta_destino, os.environ.get("APP_ENV", "Railway"))
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador()
            if _USA_POSTGRES:
                cur.execute("INSERT INTO ejecuciones (persona,inicio,estado,bot,aseguradora,ips_id,ips_nit,ips_nombre_estandar,nombre_detectado,metodo_identificacion,periodo,ruta_destino,entorno) VALUES (" + ",".join([m] * len(values)) + ") RETURNING id", values)
                result = cur.fetchone()[0]
            else:
                cur.execute("INSERT INTO ejecuciones (persona,inicio,estado,bot,aseguradora,ips_id,ips_nit,ips_nombre_estandar,nombre_detectado,metodo_identificacion,periodo,ruta_destino,entorno) VALUES (" + ",".join([m] * len(values)) + ")", values)
                result = cur.lastrowid
            if _USA_POSTGRES and result is None:
                cur.execute("SELECT currval(pg_get_serial_sequence('ejecuciones','id'))"); result = cur.fetchone()[0]
            conn.commit(); return result
        finally:
            conn.close()


def cerrar_ejecucion(ejecucion_id, estado, **totales):
    if not ejecucion_id:
        return
    permitidos = {"total_detectadas", "total_procesadas", "total_exitosas", "total_fallidas", "total_omitidas", "total_redescargadas", "decision_redescarga", "facturas_previas", "facturas_seleccionadas", "facturas_descartadas"}
    fields = {"fin": datetime.now(timezone.utc).isoformat(), "estado": estado}
    fields.update({k: v for k, v in totales.items() if k in permitidos})
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador()
            cur.execute("UPDATE ejecuciones SET " + ",".join(f"{k}={m}" for k in fields) + " WHERE id=" + m,
                        list(fields.values()) + [ejecucion_id])
            conn.commit()
        finally:
            conn.close()


def migrar_progreso_json(ips_dir):
    path = Path(ips_dir) / "progreso.json"
    if not path.exists():
        return 0
    try:
        data = json.loads(path.read_text(encoding="utf-8")); detalle = data.get("detalle") or {}; meta = data.get("meta") or {}
        items = [{"factura": factura, "siniestro": info.get("siniestro") or info.get("tipo"), "fecha_descarga": info.get("fecha_descarga")} for factura, info in detalle.items()]
        return registrar_descargas(meta.get("aseguradora") or "Desconocida", meta.get("ips_nombre") or "Desconocida", meta.get("periodo"), meta.get("identidad"), items)
    except Exception:
        return 0


def contar_historial(aseguradora=None, ips_nombre=None):
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador(); query = "SELECT COUNT(*) FROM descargas WHERE 1=1"; params = []
            if aseguradora: query += f" AND aseguradora={m}"; params.append(aseguradora)
            if ips_nombre: query += f" AND ips_nombre={m}"; params.append(ips_nombre)
            cur.execute(query, params); return cur.fetchone()[0]
        finally:
            conn.close()
