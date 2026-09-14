"""
Registro de rutas de descarga PERSONALIZADAS que se hayan usado alguna vez,
para que el vigilante de progreso.json (24h) también las revise — no solo
las carpetas por defecto. Sin esto, un progreso.json guardado en una ruta
personalizada nunca expiraría solo.

Se guarda en downloads/rutas_personalizadas.json (dentro de la carpeta que
ya protege el volumen persistente de Railway, igual que historial.db).
"""
import json
import threading
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
ARCHIVO = BASE_DIR / "downloads" / "rutas_personalizadas.json"
ARCHIVO.parent.mkdir(parents=True, exist_ok=True)

_lock = threading.RLock()


def _cargar():
    if ARCHIVO.exists():
        try:
            return json.loads(ARCHIVO.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def _guardar(datos):
    try:
        ARCHIVO.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def registrar_ruta(aseguradora: str, ruta):
    """Se llama cada vez que un proceso arranca con una carpeta de descarga
    personalizada — nunca lanza, un fallo aquí no debe afectar el arranque."""
    if not ruta:
        return
    try:
        with _lock:
            datos = _cargar()
            entrada = {"aseguradora": aseguradora, "ruta": str(ruta)}
            if entrada not in datos:
                datos.append(entrada)
                _guardar(datos)
    except Exception:
        pass


def obtener_rutas():
    with _lock:
        return list(_cargar())


def quitar_ruta(aseguradora: str, ruta):
    try:
        with _lock:
            datos = _cargar()
            datos = [d for d in datos if not (d.get("aseguradora") == aseguradora and d.get("ruta") == str(ruta))]
            _guardar(datos)
    except Exception:
        pass
