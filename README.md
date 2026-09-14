# Descargador de Cartas Glosa — Panel unificado

Un solo programa, un solo servicio de Railway, **5 aseguradoras**:

| Página | Aseguradora | Portal |
|---|---|---|
| `/estado` | SIS / Seguros del Estado | soatestado.sis.co |
| `/sura` | Suramericana | soatsura.sis.co |
| `/bolivar` | Seguros Bolívar | (portal Activa IT) |
| `/previsora` | Previsora SOAT | (portal Activa IT) |
| `/mundial` | Seguros Mundial | a2m-mundial.iqdigital.com.co |

La página principal (`/`) muestra las 5 como tarjetas para entrar a cada una.

## Arquitectura

Cada bot vive en su propio módulo dentro de `bots/` (`estado_sura.py`,
`bolivar.py`, `previsora.py`, `mundial.py`) como un **Flask Blueprint**
independiente, montado bajo su propia ruta. Estado y Sura comparten un mismo
módulo porque su lógica es idéntica (mismo tipo de portal); los otros tres
son aplicaciones grandes y maduras que ya existían por separado — se
tocó lo mínimo posible para no arriesgar romper algo que ya funciona:

- `@app.route` → `@bp.route` (con `url_prefix` por bot)
- Plantillas renombradas (`bolivar_index.html`, etc.) para no chocar entre sí
- Estáticos renombrados (`logo_bolivar.png`, `favicon_bolivar.ico`, etc.)
- Carpeta de descargas separada por bot (`downloads/<bot>/...`)
- Todas las llamadas del frontend (`fetch('/api/...')`) se reescribieron
  con el prefijo correspondiente (`/bolivar/api/...`, etc.)

Cada bot mantiene su propio estado en memoria (`jobs = {}` o `job_states`),
completamente aislado de los demás — se pueden correr varios a la vez sin
que se pisen.

## ⚠️ Cambios que hice y debes revisar

1. **Autenticación obligatoria por sesión**. El panel maneja documentos
   confidenciales de varias aseguradoras. Todas las páginas, APIs, acciones
   de los bots y descargas requieren una sesión iniciada en `/login`.
   Configura en Railway o en tu entorno local:
   - `PANEL_USER` = usuario autorizado
   - `PANEL_PASSWORD_HASH` = hash generado con Werkzeug
   - `SESSION_SECRET_KEY` = una clave aleatoria larga para firmar sesiones

   Para producción se recomienda crear una cuenta diferente para cada una de
   las 2 o 3 personas. Se pueden configurar varias cuentas con `PANEL_USERS`,
   usando un objeto JSON cuyos valores sean hashes:

   ```text
   PANEL_USERS={"Ana":"HASH_ANA","Carlos":"HASH_CARLOS","Laura":"HASH_LAURA"}
   ```

   `PANEL_USERS` reemplaza a `PANEL_USER` y `PANEL_PASSWORD_HASH` cuando se
   define. Cada persona podrá cerrar su propia sesión sin afectar las demás.

   Para generar el hash sin guardar la contraseña en el código:

   ```powershell
   python -c "from werkzeug.security import generate_password_hash; print(generate_password_hash('TU_CONTRASENA'))"
   ```

   La salida de ese comando debe configurarse como `PANEL_PASSWORD_HASH`.
   Sin estas tres variables, el acceso permanece bloqueado por seguridad.

   **Credenciales locales de prueba:**
   - Usuario: `Snet`
   - Contraseña: `1234`

   Estas credenciales están configuradas únicamente en `iniciar.bat` para
   pruebas locales. En producción se deben reemplazar por las variables
   seguras de Railway (`PANEL_USER`, `PANEL_PASSWORD_HASH` y
   `SESSION_SECRET_KEY`). No se debe subir una contraseña de producción al
   código ni al repositorio.

   El login permite como máximo **5 intentos fallidos consecutivos por
   origen**. Al superar el límite, el acceso queda bloqueado temporalmente
   durante 15 minutos. El bloqueo se restablece automáticamente al terminar
   ese tiempo o al reiniciar la aplicación local. En Railway, para forzar el
   restablecimiento se puede reiniciar el servicio; si el problema persiste,
   se deben revisar las variables de entorno y desplegar nuevamente.

   **Protección de enlaces directos y visibilidad compartida**. Los enlaces
   directos a cualquier aseguradora, por ejemplo `/estado`, `/bolivar` o
   `/previsora`, también están protegidos desde el backend. Sin sesión se
   redirigen obligatoriamente a `/login`; las APIs de los bots no entregan
   información ni permiten iniciar procesos sin una sesión válida.

   El estado de ejecución, las IPS procesadas y el progreso pertenecen al bot
   y se mantienen compartidos entre las sesiones autenticadas. Por eso
   cualquier usuario autorizado puede consultar el progreso de otro usuario,
   lo que permite coordinar qué IPS están ocupadas y evitar duplicar trabajo.

3. **`headless=True` + `--no-sandbox` en los 5 bots.** Bolívar, Previsora y
   Mundial ya traían `headless=True` (listos para Railway). Estado/Sura
   todavía tenía `headless=False` (pensado para tu PC) — lo cambié, porque
   un servidor sin pantalla no puede abrir un navegador visible. También
   agregué `--no-sandbox --disable-dev-shm-usage` a los que no lo tenían
   (Mundial ya lo traía) — es casi siempre necesario para que Chromium
   corra dentro de un contenedor Docker.

4. **Nota sobre los logos**: se revisaron los 5 archivos `logo_*.png` de
   `static/` y son distintos entre sí (Bolívar, Previsora y Mundial tienen
   cada uno su propio logo real, incluyendo el de Mundial con el texto
   "seguros mundial" visible) — no se encontró ningún logo genérico o
   placeholder compartido entre ellos. No se necesitó ningún cambio aquí.

5. **`sleepApplication: true`** en `railway.json` — se mantiene así a propósito
   (decisión ya tomada): ahorra costo dejando que Railway "duerma" la app
   sin tráfico, con el riesgo aceptado de que un proceso largo se corte si
   la duerme a mitad de camino. El volumen persistente (ver más abajo)
   ayuda a que ese corte no borre lo ya descargado.

6. **Versión de Playwright unificada** (`>=1.49.0`). Bolívar y Previsora
   traían fijado `playwright==1.44.0`; como todos los bots corren en el
   mismo entorno Python, solo puede haber una versión instalada. Usé la
   misma que ya está probada en Estado/Sura. No debería romper nada (las
   funciones que usan son estables entre esas versiones), pero conviene que
   pruebes Bolívar y Previsora con atención la primera vez.

## Cómo correrlo localmente

`iniciar.bat` — crea el entorno virtual, instala dependencias y Chromium,
y abre `http://localhost:8080`.

## Cómo desplegarlo en Railway

1. Sube esta carpeta a un repo de GitHub (o usa `railway up` desde la CLI).
2. En Railway, configura `PANEL_USER`, `PANEL_PASSWORD_HASH` y
   `SESSION_SECRET_KEY` (ver punto 1 arriba).
3. Si necesitas que los PDFs/ZIPs persistan entre reinicios, agrega un
   **Volume** de Railway montado en `/app/downloads` (si no, Railway borra
   el disco del contenedor en cada redeploy).
4. Railway detecta el `Dockerfile` automáticamente.

## Estructura

```
app.py                    <- panel principal + registro de los 5 blueprints
bots/
  estado_sura.py          <- SIS Estado + Suramericana
  bolivar.py              <- Seguros Bolívar
  previsora.py            <- Previsora SOAT
  mundial.py              <- Seguros Mundial
templates/
   login.html               <- inicio de sesión obligatorio
  landing.html            <- página de inicio con las 5 tarjetas
  estado_sura_index.html
  bolivar_index.html
  previsora_index.html
  mundial_index.html
static/
   favicon.ico              <- imagen de marca usada junto a Salud Net en el login
   logo_estado.png, logo_sura.png
  logo_bolivar.png, favicon_bolivar.ico
  logo_previsora.png, favicon_previsora.ico
  logo_mundial.png, favicon_mundial.ico
downloads/
  estado/  sura/  bolivar/  previsora/  mundial/
Dockerfile, railway.json, requirements.txt, .gitignore, iniciar.bat
```
