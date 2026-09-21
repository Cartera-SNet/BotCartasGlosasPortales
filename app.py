"""
Activa IT - Descargador de Cartas Glosa (panel unificado)
============================================================================
Un solo programa, un solo puerto, una página de inicio con 5 bots:

  /estado      -> SIS / Seguros del Estado
  /sura        -> Suramericana
  /bolivar     -> Seguros Bolívar
  /previsora   -> Previsora SOAT
  /mundial     -> Seguros Mundial

Cada bot es un módulo (Flask Blueprint) independiente en bots/, con su
propio estado en memoria, su propia carpeta de descargas
(downloads/<bot>/...) y su propia automatización — viven en el mismo
proceso pero no comparten variables ni se pisan entre sí.

Estado/Sura comparten un mismo módulo (bots/estado_sura.py) porque su lógica
es idéntica (mismo tipo de portal); Bolívar, Previsora y Mundial son cada
uno su propia aplicación grande y madura, migrada tal cual a Blueprint —
se tocó lo mínimo posible (rutas, plantilla, estáticos, carpeta de
descargas) para no arriesgar romper algo que ya funciona en producción.
"""

import os
from flask import Flask, render_template

app = Flask(__name__)

# ---- Registrar cada bot como Blueprint ----
from bots import estado_sura, bolivar, previsora, mundial
from bots.estado_sura import bp as estado_sura_bp
from bots.bolivar import bp as bolivar_bp
from bots.previsora import bp as previsora_bp
from bots.mundial import bp as mundial_bp

app.register_blueprint(estado_sura_bp)
app.register_blueprint(bolivar_bp)
app.register_blueprint(previsora_bp)
app.register_blueprint(mundial_bp)

# ---- Vigilante de progreso.json abandonados (24h de gaveta) ----
# Se arranca aquí (a nivel de módulo, no dentro de __main__) para que
# funcione igual corriendo con "python app.py" o servido por gunicorn.
from bots import vigilante_progreso
vigilante_progreso.iniciar_vigilante([
    {"nombre": "SIS Estado", "download_dir": estado_sura.DOWNLOAD_DIR / "estado",
     "esta_activa": lambda: estado_sura.count_running_jobs("estado") > 0},
    {"nombre": "Suramericana", "download_dir": estado_sura.DOWNLOAD_DIR / "sura",
     "esta_activa": lambda: estado_sura.count_running_jobs("sura") > 0},
    {"nombre": "Bolívar", "download_dir": bolivar.DOWNLOAD_DIR,
     "esta_activa": lambda: bolivar.count_running_jobs() > 0},
    {"nombre": "Previsora", "download_dir": previsora.DOWNLOAD_DIR,
     "esta_activa": lambda: previsora.count_running_jobs() > 0},
    {"nombre": "Mundial", "download_dir": mundial.DOWNLOAD_DIR,
     "esta_activa": lambda: mundial.count_running_jobs() > 0},
])


# ---- Panel principal ----
def _rgba(hex_color, alpha):
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


BOTS = [
    {"id": "estado", "nombre": "SIS Estado", "empresa": "Seguros del Estado",
     "logo": "logo_estado.png", "color": "#a41e35", "url": "/estado",
     "status_url": "/api/estado/status"},
    {"id": "sura", "nombre": "Suramericana", "empresa": "Suramericana",
     "logo": "logo_sura.png", "color": "#0032a1", "url": "/sura",
     "status_url": "/api/sura/status"},
    {"id": "bolivar", "nombre": "Bolívar", "empresa": "Seguros Bolívar",
     "logo": "logo_bolivar.png", "color": "#f4b400", "texto_boton": "#1e293b", "url": "/bolivar",
     "status_url": "/bolivar/api/status"},
    {"id": "previsora", "nombre": "Previsora", "empresa": "Previsora SOAT",
     "logo": "logo_previsora.png", "color": "#08482e", "url": "/previsora",
     "status_url": "/previsora/api/status"},
    {"id": "mundial", "nombre": "Mundial", "empresa": "Seguros Mundial",
     "logo": "logo_mundial.png", "color": "#0088ff", "url": "/mundial",
     "status_url": "/mundial/api/status"},
]

for _b in BOTS:
    _b["sombra"] = _rgba(_b["color"], 0.45)
    _b["tinte"] = _rgba(_b["color"], 0.07)


@app.route("/")
@app.route("/landing")
@app.route("/landing.html")
def landing():
    return render_template("landing.html", bots=BOTS)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    print("\n" + "=" * 60)
    print("  🏥 Activa IT — Descargador de Cartas Glosa (panel unificado)")
    print("=" * 60)
    for b in BOTS:
        print(f"   http://localhost:{port}{b['url']}  -> {b['nombre']}")
    print("=" * 60 + "\n")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
