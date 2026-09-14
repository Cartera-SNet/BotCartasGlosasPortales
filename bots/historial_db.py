"""
Historial permanente de facturas/consecutivos ya descargados, separado por
aseguradora e IPS. Vive aparte de progreso.json (que sigue sirviendo solo
para reanudar un lote a medio camino) — esta base de datos es el registro
de largo plazo que sobrevive aunque se borre o expire el progreso de una
corrida puntual.

DOS MODOS, la misma API hacia afuera (nada en los bots necesita cambiar):

- Por defecto: SQLite, un archivo dentro de downloads/ (sin dependencias
  nuevas, funciona igual en Railway, Local y el .exe).
- Si se configura la variable de entorno DATABASE_URL (el nombre estándar
  que usa Neon): Postgres real en la nube — el historial deja de depender
  por completo del disco de Railway, sobrevive a cualquier redeploy o
  reinicio sin necesitar el volumen persistente para esta parte.
"""
import os
import json
import sqlite3
import threading
from pathlib import Path
from datetime import datetime, timezone

BASE_DIR = Path(__file__).resolve().parent.parent

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
_USA_POSTGRES = bool(DATABASE_URL)

if _USA_POSTGRES:
    try:
        import psycopg2
        import psycopg2.extras
    except Exception as e:
        # Si psycopg2 no está disponible por cualquier razón, NUNCA debe
        # tumbar la aplicación entera al arrancar -- se cae de vuelta a
        # SQLite en silencio (con un aviso en consola) en vez de explotar.
        print(f"[historial_db] ⚠️ No se pudo cargar psycopg2 ({e}); usando SQLite en su lugar.")
        _USA_POSTGRES = False

# Solo aplica en modo SQLite -- vive DENTRO de downloads/, no en la raíz del
# proyecto, para quedar protegido igual que progreso.json si se monta el
# volumen persistente de Railway.
DB_PATH = Path(os.environ.get("HISTORIAL_DB_PATH", str(BASE_DIR / "downloads" / "historial.db")))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

_lock = threading.RLock()


def _conectar():
    if _USA_POSTGRES:
        return psycopg2.connect(DATABASE_URL)
    conn = sqlite3.connect(str(DB_PATH), timeout=10, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")  # evita bloqueos entre lecturas y escrituras simultáneas
    return conn


def _marcador():
    """Postgres usa %s, SQLite usa ? -- toda consulta usa esto en vez de escribir el símbolo a mano."""
    return "%s" if _USA_POSTGRES else "?"


def inicializar():
    with _lock:
        conn = _conectar()
        try:
            cur = conn.cursor()
            if _USA_POSTGRES:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS descargas (
                        id SERIAL PRIMARY KEY,
                        aseguradora TEXT NOT NULL,
                        ips_nombre TEXT NOT NULL,
                        factura TEXT NOT NULL,
                        siniestro TEXT,
                        periodo TEXT,
                        identidad TEXT,
                        fecha_descarga TEXT NOT NULL,
                        fecha_migrado TEXT NOT NULL,
                        UNIQUE(aseguradora, ips_nombre, factura)
                    )
                """)
            else:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS descargas (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        aseguradora TEXT NOT NULL,
                        ips_nombre TEXT NOT NULL,
                        factura TEXT NOT NULL,
                        siniestro TEXT,
                        periodo TEXT,
                        identidad TEXT,
                        fecha_descarga TEXT NOT NULL,
                        fecha_migrado TEXT NOT NULL,
                        UNIQUE(aseguradora, ips_nombre, factura)
                    )
                """)
            cur.execute("CREATE INDEX IF NOT EXISTS idx_ips ON descargas(aseguradora, ips_nombre)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_factura ON descargas(aseguradora, ips_nombre, factura)")
            conn.commit()
        finally:
            conn.close()


try:
    inicializar()
except Exception as e:
    if _USA_POSTGRES:
        print(f"[historial_db] ⚠️ No se pudo conectar a Postgres al arrancar ({e}); "
              f"usando SQLite en su lugar por esta sesión.")
        _USA_POSTGRES = False
        try:
            inicializar()
        except Exception as e2:
            print(f"[historial_db] ⚠️ Tampoco se pudo inicializar SQLite ({e2}); "
                  f"el historial quedará inactivo, pero el panel sigue funcionando.")
    else:
        print(f"[historial_db] ⚠️ No se pudo inicializar el historial ({e}); "
              f"quedará inactivo, pero el panel sigue funcionando.")


def registrar_descargas(aseguradora: str, ips_nombre: str, periodo: str, identidad: str, items: list):
    """
    items: lista de dicts, cada uno con al menos {"factura": ..., "fecha_descarga": ...}
           y opcionalmente {"siniestro": ...}. Nunca duplica ni truena si se
           migra dos veces el mismo progreso.json (upsert real en los dos motores).
    Devuelve cuántas filas se insertaron/actualizaron. Si la base de datos
    no responde (Neon caído, red, lo que sea), devuelve 0 en vez de tumbar
    el proceso que la llamó -- el historial es un beneficio adicional,
    nunca debe poder frenar una descarga real.
    """
    if not items:
        return 0
    ahora = datetime.now(timezone.utc).isoformat()
    try:
        with _lock:
            conn = _conectar()
            try:
                cur = conn.cursor()
                filas = [
                    (aseguradora, ips_nombre, str(it["factura"]), it.get("siniestro"), periodo,
                     identidad, it.get("fecha_descarga", ahora), ahora)
                    for it in items
                ]
                if _USA_POSTGRES:
                    psycopg2.extras.execute_values(cur, """
                        INSERT INTO descargas
                            (aseguradora, ips_nombre, factura, siniestro, periodo, identidad, fecha_descarga, fecha_migrado)
                        VALUES %s
                        ON CONFLICT (aseguradora, ips_nombre, factura) DO UPDATE SET
                            siniestro = EXCLUDED.siniestro,
                            periodo = EXCLUDED.periodo,
                            identidad = EXCLUDED.identidad,
                            fecha_descarga = EXCLUDED.fecha_descarga,
                            fecha_migrado = EXCLUDED.fecha_migrado
                    """, filas)
                else:
                    cur.executemany("""
                        INSERT OR REPLACE INTO descargas
                            (aseguradora, ips_nombre, factura, siniestro, periodo, identidad, fecha_descarga, fecha_migrado)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, filas)
                conn.commit()
                return len(filas)
            finally:
                conn.close()
    except Exception as e:
        print(f"[historial_db] ⚠️ No se pudo registrar en el historial ({e}); se continúa sin migrar esto.")
        return 0


def buscar_ya_descargadas(aseguradora: str, ips_nombre: str, facturas: list):
    """
    Devuelve un dict {factura: fecha_descarga} SOLO para las facturas de la
    lista que ya aparecen en el historial de esa aseguradora + IPS. Si la
    base de datos no responde, devuelve un dict vacío (equivalente a "no
    encontré nada repetido") en vez de romper el arranque del proceso.
    """
    if not facturas:
        return {}
    try:
        with _lock:
            conn = _conectar()
            try:
                cur = conn.cursor()
                m = _marcador()
                placeholders = ",".join([m] * len(facturas))
                cur.execute(f"""
                    SELECT factura, fecha_descarga FROM descargas
                    WHERE aseguradora = {m} AND ips_nombre = {m} AND factura IN ({placeholders})
                """, [aseguradora, ips_nombre] + [str(f) for f in facturas])
                filas = cur.fetchall()
                return {factura: fecha for factura, fecha in filas}
            finally:
                conn.close()
    except Exception as e:
        print(f"[historial_db] ⚠️ No se pudo consultar el historial ({e}); se continúa sin la advertencia de duplicados.")
        return {}


def migrar_progreso_json(ips_dir) -> int:
    """
    Lee un progreso.json (formato enriquecido: detalle + meta) y migra su
    contenido al historial permanente. Se llama SIEMPRE antes de borrar el
    archivo — nunca lanza una excepción; si algo sale mal simplemente no
    migra nada y devuelve 0, para no bloquear jamás el borrado real.
    """
    p = Path(ips_dir) / "progreso.json"
    if not p.exists():
        return 0
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return 0

    detalle = data.get("detalle") or {}
    if not detalle:
        return 0
    meta = data.get("meta") or {}

    items = [
        {
            "factura": factura,
            "siniestro": info.get("siniestro") or info.get("tipo"),
            "fecha_descarga": info.get("fecha_descarga"),
        }
        for factura, info in detalle.items()
    ]
    try:
        return registrar_descargas(
            aseguradora=meta.get("aseguradora") or "Desconocida",
            ips_nombre=meta.get("ips_nombre") or "Desconocida",
            periodo=meta.get("periodo"),
            identidad=meta.get("identidad"),
            items=items,
        )
    except Exception:
        return 0


def contar_historial(aseguradora: str = None, ips_nombre: str = None):
    """Utilidad de diagnóstico: cuántos registros hay, opcionalmente filtrado.
    Si la base de datos no responde, devuelve 0 en vez de romper."""
    try:
        with _lock:
            conn = _conectar()
            try:
                cur = conn.cursor()
                m = _marcador()
                query = "SELECT COUNT(*) FROM descargas WHERE 1=1"
                params = []
                if aseguradora:
                    query += f" AND aseguradora = {m}"
                    params.append(aseguradora)
                if ips_nombre:
                    query += f" AND ips_nombre = {m}"
                    params.append(ips_nombre)
                cur.execute(query, params)
                return cur.fetchone()[0]
            finally:
                conn.close()
    except Exception as e:
        print(f"[historial_db] ⚠️ No se pudo consultar el historial ({e}).")
        return 0
