"""
Vigilante de progreso.json abandonados. Cada hora (por defecto) revisa
todas las carpetas de descargas de las 5 aseguradoras (las de por defecto
Y cualquier ruta personalizada que se haya usado alguna vez); si un
progreso.json lleva más de 24h (por defecto) sin actividad Y esa
aseguradora no tiene NINGÚN proceso corriendo en este momento, lo migra al
historial permanente y lo borra en silencio — sin avisar a Telegram (eso
solo pasa cuando alguien lo borra a mano desde el panel).

El chequeo "¿hay algo corriendo?" es a propósito conservador: si CUALQUIER
proceso de esa aseguradora está activo, se salta el barrido completo de esa
aseguradora en esta pasada (mejor esperar una hora más que arriesgarse a
tocar algo que sigue en uso).
"""
import time
import threading
from pathlib import Path

from . import historial_db
from . import registro_rutas

HORAS_EXPIRACION_DEFECTO = 24
INTERVALO_SEGUNDOS_DEFECTO = 3600  # 1 hora


def _barrer_carpeta(nombre, root, esta_activa, ahora, limite_segundos, resultado):
    """Revisa una sola carpeta (por defecto o personalizada). Nunca lanza."""
    try:
        if esta_activa():
            return None  # hay algo corriendo en esta aseguradora, se salta por completo esta pasada

        root = Path(root)
        if not root.exists():
            return 0

        encontrados = 0
        for p in root.glob("**/progreso.json"):
            encontrados += 1
            try:
                edad_segundos = ahora - p.stat().st_mtime
                if edad_segundos < limite_segundos:
                    continue
                # Segundo chequeo, justo antes de tocar el archivo: si algo
                # arrancó justo a mitad de este barrido, se aborta aquí en
                # vez de confiar solo en el chequeo del inicio de la pasada.
                if esta_activa():
                    break
                n_migradas = historial_db.migrar_progreso_json(p.parent)
                p.unlink()
                encontrados -= 1
                resultado.append({
                    "aseguradora": nombre, "carpeta": str(p.parent),
                    "migradas": n_migradas, "edad_horas": round(edad_segundos / 3600, 1),
                })
            except Exception:
                continue
        return encontrados
    except Exception:
        return None


def ejecutar_barrido_una_vez(configs, horas_expiracion=HORAS_EXPIRACION_DEFECTO):
    """
    configs: lista de dicts {"nombre": str, "download_dir": Path, "esta_activa": callable() -> bool}
    Devuelve una lista de dicts con lo que se migró/borró en esta pasada,
    útil tanto para logging como para pruebas.
    Nunca lanza — cualquier error en una aseguradora no afecta a las demás.
    """
    resultado = []
    ahora = time.time()
    limite_segundos = horas_expiracion * 3600
    activa_por_nombre = {cfg.get("nombre", "?"): cfg["esta_activa"] for cfg in configs}

    for cfg in configs:
        _barrer_carpeta(cfg.get("nombre", "?"), cfg["download_dir"], cfg["esta_activa"],
                         ahora, limite_segundos, resultado)

    # Rutas personalizadas registradas alguna vez (fuera de las de por
    # defecto). Si ya no queda ningún progreso.json en una de ellas, se
    # quita del registro para no acumular entradas viejas para siempre.
    for entrada in registro_rutas.obtener_rutas():
        aseguradora = entrada.get("aseguradora", "?")
        ruta = entrada.get("ruta")
        esta_activa = activa_por_nombre.get(aseguradora)
        if not esta_activa or not ruta:
            continue
        restantes = _barrer_carpeta(aseguradora, ruta, esta_activa, ahora, limite_segundos, resultado)
        if restantes == 0:
            registro_rutas.quitar_ruta(aseguradora, ruta)

    return resultado


def iniciar_vigilante(configs, intervalo_segundos=INTERVALO_SEGUNDOS_DEFECTO,
                       horas_expiracion=HORAS_EXPIRACION_DEFECTO):
    """Arranca el hilo de fondo. No bloquea — se llama una vez al iniciar la app."""
    def _ciclo():
        while True:
            time.sleep(intervalo_segundos)
            try:
                ejecutar_barrido_una_vez(configs, horas_expiracion)
            except Exception:
                pass

    t = threading.Thread(target=_ciclo, daemon=True)
    t.start()
    return t

