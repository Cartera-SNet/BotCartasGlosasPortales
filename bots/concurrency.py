"""
Límite GLOBAL de automatizaciones simultáneas, contando entre las 5
aseguradoras juntas (no cada una por su lado). Existe para que, sin
importar cuántos procesos pesados de Playwright estén corriendo, siempre
queden hilos libres del servidor para lo básico: cargar el panel, subir
un archivo, consultar el progreso.

Se mantiene POR DEBAJO del número de hilos de gunicorn (ver Dockerfile) a
propósito — el margen entre los dos es lo que garantiza que el panel
nunca se quede sin hilos para responder, sin importar cuántas descargas
pesadas estén corriendo al mismo tiempo.
"""
import os
import threading

MAX_GLOBAL_JOBS = int(os.environ.get("MAX_GLOBAL_JOBS", "6"))

_lock = threading.RLock()
_contadores = {"estado": 0, "sura": 0, "bolivar": 0, "previsora": 0, "mundial": 0}


def puede_iniciar():
    with _lock:
        return sum(_contadores.values()) < MAX_GLOBAL_JOBS


def total_activos():
    with _lock:
        return sum(_contadores.values())


def registrar_inicio(bot):
    with _lock:
        _contadores[bot] = _contadores.get(bot, 0) + 1


def registrar_fin(bot):
    with _lock:
        if _contadores.get(bot, 0) > 0:
            _contadores[bot] -= 1
