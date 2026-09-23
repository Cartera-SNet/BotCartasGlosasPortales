"""Persistencia historica compatible con SQLite local y PostgreSQL/Neon.

IMPORTANTE (Postgres/Neon): el esquema aquí es el NORMALIZADO -- ninguna
tabla guarda el nombre de la IPS como texto libre; todo apunta a
`ips.id` por `ips_id`. Esto evita nombres crudos con guion bajo y datos
duplicados/desincronizados. Para ver los datos legibles, se usa la vista
`descargas_detalladas` (o se hace JOIN a `ips` directamente).

SQLite (Local/.exe) se deja EXACTAMENTE como estaba -- fuera de alcance
por ahora, no se le aplica esta normalización.
"""
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


def _marcador():
    return "%s" if _USA_POSTGRES else "?"


def _conectar():
    if _USA_POSTGRES:
        return psycopg2.connect(DATABASE_URL)
    conn = sqlite3.connect(str(DB_PATH), timeout=10, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _asegurar_columnas_descargas_sqlite(cur):
    columnas = {
        "ips_id": "BIGINT", "ips_nit": "TEXT", "nombre_detectado": "TEXT",
        "metodo_identificacion": "TEXT", "primera_descarga": "TEXT",
        "ultima_descarga": "TEXT", "veces_procesada": "INTEGER NOT NULL DEFAULT 1",
        "ultima_ejecucion_id": "BIGINT", "ultimo_resultado": "TEXT", "sede": "TEXT",
    }
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
                cur.execute("""CREATE TABLE IF NOT EXISTS ips (
                    id SERIAL PRIMARY KEY, nit TEXT NOT NULL UNIQUE, nombre_estandar TEXT NOT NULL,
                    responsable TEXT, activo BOOLEAN NOT NULL DEFAULT true,
                    fecha_creacion TIMESTAMPTZ NOT NULL DEFAULT now(),
                    fecha_actualizacion TIMESTAMPTZ NOT NULL DEFAULT now())""")
                cur.execute("""CREATE TABLE IF NOT EXISTS ips_alias (
                    id SERIAL PRIMARY KEY, ips_id INTEGER NOT NULL REFERENCES ips(id),
                    nombre_alias TEXT NOT NULL, UNIQUE(ips_id, nombre_alias))""")
                cur.execute("""CREATE TABLE IF NOT EXISTS ejecuciones (
                    id BIGSERIAL PRIMARY KEY, persona TEXT NOT NULL, bot TEXT NOT NULL,
                    aseguradora TEXT NOT NULL, ips_id INTEGER REFERENCES ips(id), periodo TEXT,
                    ruta_destino TEXT, entorno TEXT, estado TEXT NOT NULL,
                    fecha_inicio TIMESTAMPTZ NOT NULL DEFAULT now(), fecha_fin TIMESTAMPTZ,
                    total_detectadas INTEGER NOT NULL DEFAULT 0, total_procesadas INTEGER NOT NULL DEFAULT 0,
                    total_exitosas INTEGER NOT NULL DEFAULT 0, total_fallidas INTEGER NOT NULL DEFAULT 0,
                    total_omitidas INTEGER NOT NULL DEFAULT 0, total_redescargadas INTEGER NOT NULL DEFAULT 0,
                    decision_redescarga TEXT, facturas_previas JSONB NOT NULL DEFAULT '[]',
                    facturas_seleccionadas JSONB NOT NULL DEFAULT '[]',
                    facturas_descartadas JSONB NOT NULL DEFAULT '[]')""")
                cur.execute("""CREATE TABLE IF NOT EXISTS descargas (
                    id BIGSERIAL PRIMARY KEY, ejecucion_id BIGINT REFERENCES ejecuciones(id),
                    aseguradora TEXT NOT NULL, ips_id INTEGER REFERENCES ips(id), factura TEXT NOT NULL,
                    siniestro TEXT, periodo TEXT, identidad TEXT, sede TEXT,
                    fecha_descarga TIMESTAMPTZ NOT NULL DEFAULT now(),
                    fecha_migrado TIMESTAMPTZ NOT NULL DEFAULT now(),
                    veces_procesada INTEGER NOT NULL DEFAULT 1,
                    UNIQUE(aseguradora, ips_id, factura))""")
                cur.execute("ALTER TABLE descargas ADD COLUMN IF NOT EXISTS sede TEXT")
                for col, expr in [
                    ("anio_descarga", "EXTRACT(YEAR FROM (fecha_descarga AT TIME ZONE 'UTC'))::integer"),
                    ("mes_descarga", "EXTRACT(MONTH FROM (fecha_descarga AT TIME ZONE 'UTC'))::integer"),
                    ("dia_descarga", "EXTRACT(DAY FROM (fecha_descarga AT TIME ZONE 'UTC'))::integer"),
                ]:
                    cur.execute(f"""DO $$ BEGIN
                        IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='descargas' AND column_name='{col}') THEN
                            ALTER TABLE descargas ADD COLUMN {col} integer GENERATED ALWAYS AS ({expr}) STORED;
                        END IF; END $$;""")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_descargas_anio_mes ON descargas(anio_descarga, mes_descarga)")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_descargas_ips ON descargas(ips_id)")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_descargas_factura ON descargas(aseguradora, factura)")
                cur.execute("""CREATE OR REPLACE VIEW descargas_detalladas AS
                    SELECT ROW_NUMBER() OVER (ORDER BY d.fecha_migrado, d.id) AS correlativo,
                        d.id, d.aseguradora, i.nombre_estandar AS ips_nombre, i.nit AS ips_nit,
                        i.responsable, d.factura, d.siniestro, d.periodo, d.identidad,
                        d.fecha_descarga, d.anio_descarga, d.mes_descarga, d.dia_descarga, d.sede
                    FROM descargas d LEFT JOIN ips i ON i.id = d.ips_id
                    ORDER BY d.fecha_migrado, d.id""")
                cur.execute("""CREATE TABLE IF NOT EXISTS ejecucion_facturas (
                    id BIGSERIAL PRIMARY KEY, ejecucion_id BIGINT NOT NULL REFERENCES ejecuciones(id),
                    factura TEXT NOT NULL, ips_id INTEGER REFERENCES ips(id), aseguradora TEXT NOT NULL,
                    bot TEXT NOT NULL, periodo TEXT, estado TEXT NOT NULL,
                    redescargada BOOLEAN NOT NULL DEFAULT false, fecha_inicio TIMESTAMPTZ, fecha_fin TIMESTAMPTZ,
                    intentos INTEGER NOT NULL DEFAULT 1, archivo TEXT, error TEXT,
                    UNIQUE(ejecucion_id, factura))""")
                cur.execute("""CREATE TABLE IF NOT EXISTS errores_ejecucion (
                    id BIGSERIAL PRIMARY KEY, ejecucion_id BIGINT REFERENCES ejecuciones(id),
                    ejecucion_factura_id BIGINT REFERENCES ejecucion_facturas(id), factura TEXT, bot TEXT,
                    aseguradora TEXT, ips_id INTEGER REFERENCES ips(id), tipo_error TEXT NOT NULL,
                    mensaje TEXT NOT NULL, etapa TEXT, recuperable BOOLEAN, intento INTEGER,
                    resultado_final TEXT, fecha TIMESTAMPTZ NOT NULL DEFAULT now())""")
                conn.commit()
            else:
                # SQLite (Local/.exe): esquema anterior, SIN cambios -- fuera
                # de alcance por ahora, se deja tal cual para no arriesgar nada.
                cur.execute("""CREATE TABLE IF NOT EXISTS descargas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, aseguradora TEXT NOT NULL,
                    ips_nombre TEXT NOT NULL, factura TEXT NOT NULL, siniestro TEXT,
                    periodo TEXT, identidad TEXT, fecha_descarga TEXT NOT NULL,
                    fecha_migrado TEXT NOT NULL, UNIQUE(aseguradora, ips_nombre, factura))""")
                _asegurar_columnas_descargas_sqlite(cur)
                cur.execute("CREATE INDEX IF NOT EXISTS idx_ips ON descargas(aseguradora, ips_nombre)")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_factura ON descargas(aseguradora, ips_nombre, factura)")
                cur.execute("""CREATE TABLE IF NOT EXISTS ips (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, nit TEXT NOT NULL UNIQUE, nombre_estandar TEXT NOT NULL,
                    responsable TEXT, razon_social TEXT, activo INTEGER NOT NULL DEFAULT 1,
                    fecha_creacion TEXT NOT NULL, fecha_actualizacion TEXT NOT NULL)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS ips_alias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, ips_id INTEGER NOT NULL, nombre_alias TEXT NOT NULL UNIQUE)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS ejecuciones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, persona TEXT NOT NULL, fecha_inicio TEXT NOT NULL, fecha_fin TEXT,
                    estado TEXT NOT NULL, bot TEXT NOT NULL, aseguradora TEXT, ips_id INTEGER,
                    ips_nit TEXT, ips_nombre_estandar TEXT, nombre_detectado TEXT,
                    metodo_identificacion TEXT, periodo TEXT, ruta_destino TEXT, entorno TEXT,
                    total_detectadas INTEGER DEFAULT 0, total_procesadas INTEGER DEFAULT 0,
                    total_exitosas INTEGER DEFAULT 0, total_fallidas INTEGER DEFAULT 0,
                    total_omitidas INTEGER DEFAULT 0, total_redescargadas INTEGER DEFAULT 0,
                    decision_redescarga TEXT, facturas_previas TEXT, facturas_seleccionadas TEXT,
                    facturas_descartadas TEXT)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS ejecucion_facturas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, ejecucion_id INTEGER NOT NULL, factura TEXT NOT NULL,
                    ips_id INTEGER, ips_nit TEXT, aseguradora TEXT, bot TEXT, periodo TEXT,
                    estado TEXT NOT NULL, redescargada INTEGER NOT NULL DEFAULT 0,
                    fecha_inicio TEXT, fecha_fin TEXT, archivo TEXT, error TEXT)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS errores_ejecucion (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, ejecucion_id INTEGER NOT NULL, factura TEXT, bot TEXT,
                    aseguradora TEXT, ips_id INTEGER, ips_nit TEXT, tipo_error TEXT NOT NULL,
                    mensaje TEXT NOT NULL, etapa TEXT, recuperable INTEGER, intento INTEGER,
                    resultado_final TEXT, fecha TEXT NOT NULL)""")
                cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS ejecucion_facturas_ejecucion_id_factura_key ON ejecucion_facturas(ejecucion_id, factura)")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_ejec_factura ON ejecucion_facturas(ips_nit, factura)")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_ejecuciones_persona ON ejecuciones(persona, fecha_inicio)")
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
    responsable = ips.get("responsable")
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor()
            if _USA_POSTGRES:
                cur.execute("""INSERT INTO ips (nit,nombre_estandar,responsable,fecha_creacion,fecha_actualizacion)
                    VALUES (%s,%s,%s,now(),now())
                    ON CONFLICT(nit) DO UPDATE SET nombre_estandar=EXCLUDED.nombre_estandar,
                    responsable=COALESCE(EXCLUDED.responsable, ips.responsable),
                    fecha_actualizacion=now() RETURNING id""",
                    (ips["nit"], ips["nombre_estandar"], responsable))
                result = cur.fetchone()[0]
            else:
                cur.execute("INSERT OR IGNORE INTO ips (nit,nombre_estandar,responsable,fecha_creacion,fecha_actualizacion) VALUES (?,?,?,?,?)",
                            (ips["nit"], ips["nombre_estandar"], responsable, ahora, ahora))
                cur.execute("UPDATE ips SET nombre_estandar=?,responsable=COALESCE(?,responsable),fecha_actualizacion=? WHERE nit=?",
                            (ips["nombre_estandar"], responsable, ahora, ips["nit"]))
                cur.execute("SELECT id FROM ips WHERE nit=?", (ips["nit"],))
                result = cur.fetchone()[0]
            conn.commit()
            return result
        finally:
            conn.close()


def registrar_descargas(aseguradora, ips_nombre, periodo, identidad, items, ips=None, sede=None):
    if not items:
        return 0
    ahora = datetime.now(timezone.utc).isoformat()
    ips = ips or _resolver_ips(nombre=ips_nombre)
    ips_id = registrar_ips(ips)
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor()
            if _USA_POSTGRES:
                filas = [(aseguradora, ips_id, str(item["factura"]), item.get("siniestro"), periodo,
                          identidad, sede, item.get("fecha_descarga", ahora)) for item in items]
                psycopg2.extras.execute_values(cur, """INSERT INTO descargas
                    (aseguradora,ips_id,factura,siniestro,periodo,identidad,sede,fecha_descarga)
                    VALUES %s ON CONFLICT (aseguradora,ips_id,factura) DO UPDATE SET
                    siniestro=EXCLUDED.siniestro,periodo=EXCLUDED.periodo,identidad=EXCLUDED.identidad,
                    sede=COALESCE(EXCLUDED.sede, descargas.sede),
                    fecha_descarga=EXCLUDED.fecha_descarga,fecha_migrado=now(),
                    veces_procesada=descargas.veces_procesada+1""", filas)
            else:
                filas = [(aseguradora, ips_nombre, str(item["factura"]), item.get("siniestro"), periodo,
                          identidad, item.get("fecha_descarga", ahora), ahora, ips_id, ips.get("nit"),
                          ips.get("nombre_detectado"), ips.get("metodo"), sede) for item in items]
                cur.executemany("""INSERT OR REPLACE INTO descargas
                    (aseguradora,ips_nombre,factura,siniestro,periodo,identidad,fecha_descarga,fecha_migrado,
                     ips_id,ips_nit,nombre_detectado,metodo_identificacion,sede)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", filas)
            conn.commit()
            return len(filas)
        finally:
            conn.close()


def _filtrar_por_mismo_mes(filas):
    """
    Nueva regla de negocio: una Carta Glosa SÍ se puede volver a descargar
    si la vez anterior fue en un mes/año distinto al actual -- solo se
    considera "ya descargada" (y por lo tanto dispara el modal de decisión)
    si la descarga previa fue en el MISMO año y mes de hoy.
    Si la fecha guardada no se puede interpretar, se es conservador y de
    todos modos se avisa (mejor preguntar de más que perder el aviso).
    """
    ahora = datetime.now(timezone.utc)
    resultado = {}
    for factura, fecha_str in filas:
        try:
            fecha = datetime.fromisoformat(str(fecha_str))
        except Exception:
            resultado[str(factura)] = fecha_str
            continue
        if fecha.year == ahora.year and fecha.month == ahora.month:
            resultado[str(factura)] = fecha_str
    return resultado


def buscar_ya_descargadas(aseguradora, ips_nombre, facturas):
    if not facturas:
        return {}
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador()
            placeholders = ",".join([m] * len(facturas))
            if _USA_POSTGRES:
                cur.execute(f"""SELECT d.factura, d.fecha_descarga FROM descargas d
                    JOIN ips i ON i.id = d.ips_id
                    WHERE d.aseguradora={m} AND i.nombre_estandar={m} AND d.factura IN ({placeholders})""",
                    [aseguradora, ips_nombre] + [str(f) for f in facturas])
            else:
                cur.execute(f"SELECT factura,fecha_descarga FROM descargas WHERE aseguradora={m} AND ips_nombre={m} AND factura IN ({placeholders})",
                            [aseguradora, ips_nombre] + [str(f) for f in facturas])
            return _filtrar_por_mismo_mes(cur.fetchall())
        finally:
            conn.close()


def buscar_ya_descargadas_por_nit(aseguradora, ips_nit, facturas):
    if not ips_nit or not facturas:
        return {}
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador(); placeholders = ",".join([m] * len(facturas))
            if _USA_POSTGRES:
                cur.execute(f"""SELECT d.factura, d.fecha_descarga FROM descargas d
                    JOIN ips i ON i.id = d.ips_id
                    WHERE d.aseguradora={m} AND i.nit={m} AND d.factura IN ({placeholders})""",
                    [aseguradora, ips_nit] + [str(f) for f in facturas])
            else:
                cur.execute(f"SELECT factura,fecha_descarga FROM descargas WHERE aseguradora={m} AND ips_nit={m} AND factura IN ({placeholders})",
                            [aseguradora, ips_nit] + [str(f) for f in facturas])
            return _filtrar_por_mismo_mes(cur.fetchall())
        finally:
            conn.close()


def registrar_facturas_ejecucion(ejecucion_id, bot, aseguradora, ips, periodo, items, redescargadas=None):
    if not ejecucion_id or not items:
        return 0
    ips_id = registrar_ips(ips) if ips else None
    ahora = datetime.now(timezone.utc).isoformat(); redescargadas = set(redescargadas or [])
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor()
            if _USA_POSTGRES:
                filas = [(ejecucion_id, str(item.get("factura")), ips_id, aseguradora, bot, periodo,
                          item.get("estado", "exitosa"), str(item.get("factura")) in redescargadas,
                          item.get("fecha_inicio", ahora), item.get("fecha_fin", ahora),
                          item.get("archivo"), item.get("error")) for item in items]
                psycopg2.extras.execute_values(cur, """INSERT INTO ejecucion_facturas
                    (ejecucion_id,factura,ips_id,aseguradora,bot,periodo,estado,redescargada,fecha_inicio,fecha_fin,archivo,error)
                    VALUES %s
                    ON CONFLICT (ejecucion_id, factura) DO UPDATE SET
                        ips_id=EXCLUDED.ips_id, aseguradora=EXCLUDED.aseguradora, bot=EXCLUDED.bot,
                        periodo=EXCLUDED.periodo, estado=EXCLUDED.estado, redescargada=EXCLUDED.redescargada,
                        fecha_inicio=EXCLUDED.fecha_inicio, fecha_fin=EXCLUDED.fecha_fin,
                        archivo=EXCLUDED.archivo, error=EXCLUDED.error""", filas)
            else:
                filas = [(ejecucion_id, str(item.get("factura")), ips_id, (ips or {}).get("nit"), aseguradora, bot, periodo,
                          item.get("estado", "exitosa"), str(item.get("factura")) in redescargadas,
                          item.get("fecha_inicio", ahora), item.get("fecha_fin", ahora), item.get("archivo"), item.get("error")) for item in items]
                cur.executemany("""INSERT OR IGNORE INTO ejecucion_facturas
                    (ejecucion_id,factura,ips_id,ips_nit,aseguradora,bot,periodo,estado,redescargada,fecha_inicio,fecha_fin,archivo,error)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", filas)
            conn.commit(); return len(filas)
        finally:
            conn.close()


def registrar_error_ejecucion(ejecucion_id, factura, bot, aseguradora, ips, tipo_error, mensaje, etapa, recuperable, intento, resultado_final):
    if not ejecucion_id:
        return False
    ips_id = registrar_ips(ips) if (ips and _USA_POSTGRES) else None
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador()
            if _USA_POSTGRES:
                values = (ejecucion_id, factura, bot, aseguradora, ips_id, tipo_error, mensaje, etapa,
                          recuperable, intento, resultado_final)
                cur.execute("INSERT INTO errores_ejecucion (ejecucion_id,factura,bot,aseguradora,ips_id,tipo_error,mensaje,etapa,recuperable,intento,resultado_final) VALUES (" + ",".join([m] * len(values)) + ")", values)
            else:
                values = (ejecucion_id, factura, bot, aseguradora, (ips or {}).get("nit"), tipo_error, mensaje, etapa,
                          recuperable, intento, resultado_final, datetime.now(timezone.utc).isoformat())
                cur.execute("INSERT INTO errores_ejecucion (ejecucion_id,factura,bot,aseguradora,ips_nit,tipo_error,mensaje,etapa,recuperable,intento,resultado_final,fecha) VALUES (" + ",".join([m] * len(values)) + ")", values)
            conn.commit(); return True
        finally:
            conn.close()


def iniciar_ejecucion(persona, bot, aseguradora, ips, periodo, ruta_destino):
    ips_id = registrar_ips(ips) if ips else None
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador()
            if _USA_POSTGRES:
                values = (persona, bot, aseguradora, ips_id, periodo, ruta_destino,
                          os.environ.get("APP_ENV", "Railway"), "iniciada")
                cur.execute("""INSERT INTO ejecuciones (persona,bot,aseguradora,ips_id,periodo,ruta_destino,entorno,estado)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""", values)
                result = cur.fetchone()[0]
            else:
                ahora = datetime.now(timezone.utc).isoformat()
                values = (persona, ahora, "iniciada", bot, aseguradora, ips_id, (ips or {}).get("nit"),
                          (ips or {}).get("nombre_estandar"), (ips or {}).get("nombre_detectado"),
                          (ips or {}).get("metodo"), periodo, ruta_destino, os.environ.get("APP_ENV", "Railway"))
                cur.execute("INSERT INTO ejecuciones (persona,fecha_inicio,estado,bot,aseguradora,ips_id,ips_nit,ips_nombre_estandar,nombre_detectado,metodo_identificacion,periodo,ruta_destino,entorno) VALUES (" + ",".join([m] * len(values)) + ")", values)
                result = cur.lastrowid
            conn.commit(); return result
        finally:
            conn.close()


def cerrar_ejecucion(ejecucion_id, estado, **totales):
    if not ejecucion_id:
        return
    permitidos = {"total_detectadas", "total_procesadas", "total_exitosas", "total_fallidas", "total_omitidas", "total_redescargadas", "decision_redescarga", "facturas_previas", "facturas_seleccionadas", "facturas_descartadas"}
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor(); m = _marcador()
            if _USA_POSTGRES:
                fields = {"fecha_fin": "now()", "estado": estado}
                sets = ["fecha_fin=now()", f"estado={m}"]
                params = [estado]
                for k, v in totales.items():
                    if k in permitidos:
                        if k in ("facturas_previas", "facturas_seleccionadas", "facturas_descartadas"):
                            sets.append(f"{k}={m}::jsonb")
                            params.append(json.dumps(v))
                        else:
                            sets.append(f"{k}={m}")
                            params.append(v)
                cur.execute(f"UPDATE ejecuciones SET {','.join(sets)} WHERE id={m}", params + [ejecucion_id])
            else:
                fields = {"fecha_fin": datetime.now(timezone.utc).isoformat(), "estado": estado}
                fields.update({k: v for k, v in totales.items() if k in permitidos})
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
            cur = conn.cursor(); m = _marcador()
            if _USA_POSTGRES:
                query = "SELECT COUNT(*) FROM descargas d LEFT JOIN ips i ON i.id = d.ips_id WHERE 1=1"; params = []
                if aseguradora: query += f" AND d.aseguradora={m}"; params.append(aseguradora)
                if ips_nombre: query += f" AND i.nombre_estandar={m}"; params.append(ips_nombre)
            else:
                query = "SELECT COUNT(*) FROM descargas WHERE 1=1"; params = []
                if aseguradora: query += f" AND aseguradora={m}"; params.append(aseguradora)
                if ips_nombre: query += f" AND ips_nombre={m}"; params.append(ips_nombre)
            cur.execute(query, params); return cur.fetchone()[0]
        finally:
            conn.close()
