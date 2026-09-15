# Descargador de Cartas Glosa — Panel unificado

## Estado actual

Aplicacion Flask/Python desplegable como un unico servicio en Railway. El
servicio integra cinco entradas de automatizacion:

| Ruta | Aseguradora | Modulo |
|---|---|---|
| `/estado` | SIS / Seguros del Estado | `bots/estado_sura.py` |
| `/sura` | Suramericana | `bots/estado_sura.py` |
| `/bolivar` | Seguros Bolivar | `bots/bolivar.py` |
| `/previsora` | Previsora SOAT | `bots/previsora.py` |
| `/mundial` | Seguros Mundial | `bots/mundial.py` |

Estado y Sura comparten logica y proceso Python, pero tienen rutas, portales,
configuracion y estados de job separados. Cada bot usa Playwright y conserva
su progreso en disco.

## Arquitectura y estructura

- `app.py`: crea Flask, registra los blueprints e inicia el vigilante de
   `progreso.json`.
- `bots/estado_sura.py`: bots SIS Estado y Sura.
- `bots/bolivar.py`: automatizacion de Bolivar.
- `bots/previsora.py`: automatizacion de Previsora.
- `bots/mundial.py`: automatizacion de Mundial.
- `bots/catalogo_ips.py`: resolucion centralizada de IPS por NIT y aliases.
- `bots/historial_db.py`: SQLite local o PostgreSQL cuando existe
   `DATABASE_URL`.
- `bots/vigilante_progreso.py`: migra y elimina `progreso.json` abandonados
   despues de 24 horas; no elimina por si mismo PDFs ni ZIPs.
- `bots/registro_rutas.py`: registra rutas personalizadas de descarga.
- `bots/concurrency.py`: limita automatizaciones simultaneas.
- `templates/`: interfaces de los cuatro paneles y el landing.
- `static/audit_ui.js`: modal comun para decidir redescargas.
- `downloads/`: datos locales o persistidos mediante volumen de Railway.

## Funcionalidades verificadas

### Identidad de la persona

Los cuatro endpoints de inicio exigen el campo `identidad` antes de crear un
job, reservar concurrencia o iniciar Playwright. La interfaz muestra la
advertencia junto a `¿Quien eres?` y la oculta al seleccionar una persona.

La identidad se guarda como metadata de progreso y en el historial basico de
descargas. La auditoria completa por ejecucion sigue pendiente de integracion.

### Duplicados historicos

El sistema consulta facturas ya registradas antes de descargarlas. El usuario
puede elegir:

- no volver a descargar ninguna;
- volver a descargar todas;
- seleccionar facturas concretas.

Bolivar y Previsora tambien pueden pausar el job cuando descubren repetidos
despues de consultar el portal. La decision se envia mediante
`/api/duplicate-decision` dentro del blueprint correspondiente.

### Progreso y soportes

`progreso.json` es temporal y sirve para reanudar. El historial de descargas
se registra separadamente. El vigilante de 24 horas migra primero el progreso
y luego elimina solo `progreso.json`.

La politica actual de PDFs, ZIPs y soportes es manual o propia del flujo del
bot. No existe una tarea general que los elimine automaticamente por edad.
Mundial elimina archivos ZIP temporales despues de extraerlos y algunos bots
eliminan ZIPs parciales obsoletos al generar uno nuevo o el final.

### Suspension de Railway

`railway.json` mantiene `sleepApplication: true`. El landing no realiza
polling mientras la pestaña esta oculta. Los paneles internos hacen polling
durante una ejecucion activa para mostrar progreso.

## Base de datos PostgreSQL / Neon

El codigo soporta dos modos:

- `DATABASE_URL` configurada: PostgreSQL/Neon mediante `psycopg2`.
- sin `DATABASE_URL`: SQLite en `downloads/historial.db`.

La tabla historica original es `descargas`. El inicializador actual crea o
verifica estas tablas adicionales:

- `ips`: catalogo central con NIT unico y nombre estandar.
- `ips_alias`: nombres alternativos.
- `ejecuciones`: estructura preparada para auditar una corrida.
- `ejecucion_facturas`: estructura preparada para facturas por corrida y
   redescargas.
- `errores_ejecucion`: estructura preparada para errores por etapa e intento.

### Estado de Neon verificado por evidencia externa

El 2026-09-15 el usuario ejecuto la migracion en Neon y reporto que las
tablas fueron creadas. La captura compartida muestra `ips` con 21 registros y
las tablas nuevas. La conexion remota no puede consultarse desde este entorno,
por lo que los siguientes puntos quedan pendientes de verificacion directa:

- que todos los registros de `descargas` tengan `ips_id`;
- que no existan duplicados por `(aseguradora, ips_id, factura)`;
- que el indice unico `ux_descargas_aseguradora_ips_factura` tenga la
   definicion esperada;
- que la restriccion antigua basada en `ips_nombre` siga o haya sido retirada.

No eliminar la restriccion antigua hasta completar esa verificacion y probar
una ejecucion real.

### Estado de la auditoria

Los cuatro endpoints de inicio crean una fila en `ejecuciones` y los flujos
cierran esa fila con estado y contadores al finalizar. Las descargas exitosas
se escriben en `descargas` y en `ejecucion_facturas` cuando se guarda el
progreso. Los errores criticos emitidos por los jobs se escriben en
`errores_ejecucion`.

La auditoria por factura de errores de portal todavia es parcial: algunos
errores conservan detalle en memoria, Excel y logs, pero no todos incluyen aun
factura y etapa especifica en `errores_ejecucion`. Queda como mejora pendiente.

Cuando `DATABASE_URL` esta definida y PostgreSQL no conecta, la aplicacion ya
no cambia silenciosamente a SQLite: el arranque falla para evitar aparentar
que el historial quedo registrado en Neon cuando no fue asi.

Telegram no forma parte del proyecto actual ni de esta documentacion.

## Configuracion y despliegue

### Local

`iniciar.bat` prepara el entorno, instala dependencias y Chromium, y abre el
panel local en `http://localhost:8080`.

### Railway

- Builder: `Dockerfile`.
- Servidor: Gunicorn, un worker y 24 threads.
- Puerto: variable `PORT`, con valor predeterminado 8080.
- `sleepApplication: true`.
- Para conservar descargas se requiere un volumen montado en `/app/downloads`.
- `DATABASE_URL` debe apuntar a Neon si se desea historial PostgreSQL.

No se documentan valores de contrasenas, tokens ni secretos.

## Validaciones realizadas

- Compilacion de todos los modulos Python.
- Rechazo de peticiones sin identidad en los cuatro bots.
- Resolucion de IPS por NIT, alias y caso `IPS_NO_IDENTIFICADA`.
- Creacion local del esquema de tablas de historial.
- Verificacion de `sleepApplication: true`.
- Revision de rutas de eliminacion de progreso, PDF, ZIP y carpetas.
- Confirmacion de que no hay integracion activa de Telegram.

## Historial de cambios

### Estado anterior documentado — baseline

El README anterior describia la arquitectura de los cinco bots, el uso de
Docker/Playwright, el despliegue en Railway, el volumen persistente y la
version unificada de Playwright. Ese contenido se conserva conceptualmente en
las secciones actuales, corregido donde no coincide con los archivos
presentes.

### 2026-09-15 — Identidad obligatoria, redescargas y catalogo IPS

#### Codigo

- Se agrego `bots/catalogo_ips.py`.
- Se agregaron tablas de catalogo y auditoria al inicializador de
   `bots/historial_db.py`.
- Se agrego registro basico de descargas al guardar progreso.
- Se corrigio la lectura de solicitudes no JSON para evitar respuestas 415 en
   formularios normales.

#### Base de datos

- Se agregaron `ips`, `ips_alias`, `ejecuciones`, `ejecucion_facturas` y
   `errores_ejecucion` al esquema inicializable.
- Neon fue actualizado por el usuario con esas tablas y el catalogo inicial
   de 21 IPS, segun evidencia compartida.
- Se agregaron columnas de trazabilidad a `descargas` en Neon, segun el SQL
   ejecutado por el usuario.
- El indice unico por `(aseguradora, ips_id, factura)` fue creado o ya existia,
   segun el mensaje de PostgreSQL `already exists, skipping`.
- Pendiente de verificacion: migracion completa de todos los registros y
   retiro de la restriccion antigua por `ips_nombre`.

#### Frontend / UI

- Se mejoro el bloque `¿Quien eres?` en las cuatro plantillas.
- Se agrego advertencia contextual cuando falta la persona.
- Se creo `static/audit_ui.js` con el modal de facturas repetidas.
- Se sustituyeron confirmaciones simples por decisiones explicitas de
   redescarga.
- El landing deja de hacer polling cuando la pestaña esta oculta.

#### Backend

- Los cuatro endpoints de inicio rechazan identidad vacia.
- Las decisiones de duplicados se aplican antes de la descarga.
- Bolivar y Previsora exponen una pausa para duplicados detectados despues de
   consultar el portal.
- El registro historico basico se actualiza sin depender de la eliminacion de
   `progreso.json`.

### 2026-09-15 — Auditoria de persistencia y correcciones de contratos

#### Codigo

- `historial_db.py` fue reconstruido para que la migracion de `descargas` sea
   idempotente tambien en SQLite, agregando las columnas de IPS si la tabla ya
   existia.
- Se agregaron escrituras de ejecucion, facturas por ejecucion, errores
   criticos y cierre de ejecucion en los cuatro bots.
- La identidad persistida de una descarga incluye `ips_id`, `ips_nit`, nombre
   detectado y metodo de identificacion cuando el catalogo puede resolverla.
- Se corrigio el caso en que una segunda peticion confirmaba duplicados pero
   no aplicaba la decision de ninguna/seleccionadas.
- Se corrigio la obtencion del ID de ejecucion en PostgreSQL usando `RETURNING
   id`; SQLite conserva `lastrowid`.
- Se elimino el fallback silencioso de Neon a SQLite cuando `DATABASE_URL`
   esta configurada.
- El guardado de progreso ahora pasa la identidad IPS resuelta por NIT al
   historial `descargas`, evitando volver a resolverla solo por nombre cuando
   el bot ya conocia el NIT.
- Se corrigieron los endpoints de control para aceptar solicitudes sin JSON
   sin responder 415 innecesariamente.
- Estado/Sura y Previsora ahora hacen polling secuencial y dejan de sondear
   al terminar, igual que Bolivar y Mundial.

#### Verificacion

- Todos los modulos Python compilan.
- Se probó desde una SQLite vacia el ciclo de IPS, ejecucion, descarga,
   `ejecucion_facturas`, error y cierre.
- Se verifico que una `DATABASE_URL` invalida no se degrada silenciosamente.
- Se probaron los cuatro guardas de identidad y los endpoints de reset sin
   obtener 415.

#### Pendientes

- Probar la ruta PostgreSQL real con `psycopg2` instalado y la `DATABASE_URL`
   productiva de Neon.
- Registrar todos los errores de factura con etapa exacta, no solo errores
   criticos emitidos por el logger.
- Verificar directamente en Neon los indices, restricciones y migracion de
   filas antiguas.

### 2026-09-15 — Correccion de arranque en Railway por nombres de columnas

#### Problema encontrado

Railway no podia iniciar Gunicorn. Neon ya tenia `ejecuciones` creada con las
columnas `fecha_inicio` y `fecha_fin`, pero el codigo intentaba crear el indice
`idx_ejecuciones_persona` usando una columna inexistente llamada `inicio`.
PostgreSQL produjo `psycopg2.errors.UndefinedColumn` y el worker se cerraba al
importar `app.py`.

#### Correccion

- El esquema inicializable usa ahora `fecha_inicio` y `fecha_fin`.
- El indice usa `(persona, fecha_inicio)`.
- El INSERT de ejecuciones usa `fecha_inicio`.
- El cierre de ejecuciones actualiza `fecha_fin`.
- Se verifico compatibilidad local contra una tabla preexistente con los
   nombres de Neon.

#### Tablas actualmente creadas en Neon

Segun las capturas y SQL ejecutado durante esta configuracion, el esquema
contiene:

| Tabla | Funcion |
|---|---|
| `descargas` | Resumen historico compatible de facturas descargadas. |
| `ips` | Catalogo central de IPS con NIT unico y nombre estandar. |
| `ips_alias` | Alias de nombres de IPS. |
| `ejecuciones` | Una fila por corrida, con persona, bot, aseguradora, IPS, periodo, estado y contadores. Usa `fecha_inicio`/`fecha_fin`. |
| `ejecucion_facturas` | Facturas participantes en cada ejecucion y redescargas. |
| `errores_ejecucion` | Errores asociados a ejecuciones, factura, etapa e intento. |

La existencia remota de cada columna se basa en la evidencia compartida y en
el error de Railway. La comprobacion directa desde este entorno queda
`Pendiente de verificacion`.

### 2026-09-15 — Rediseño visual del selector de identidad

#### Frontend / UI

- Se reemplazo el bloque visual antiguo de `¿Quién eres?` en los cuatro
   paneles: Estado/Sura, Bolivar, Previsora y Mundial.
- Se agrego `static/identity.css` como hoja compartida para evitar que cada
   plantilla mantenga un diseño divergente.
- El nuevo componente incluye icono, texto de contexto, tarjetas de opcion,
   iconos por persona, marca visual de seleccion, foco accesible, campo para
   otra persona y estado de error destacado.
- La logica existente ahora cambia clases (`is-selected` y `has-error`) en
   lugar de imponer estilos inline que ocultaban el nuevo diseño.

#### Resultado

- La seleccion conserva su obligatoriedad y la advertencia sigue apareciendo
   junto al campo.
- El diseño deja de ser un par de botones planos dentro de un contenedor gris
   y pasa a ser un bloque visual consistente con el panel.
- Las cuatro plantillas cargan `/static/identity.css`.

#### Verificacion

- No quedan errores reportados en las cuatro plantillas ni en `identity.css`.
- Se verifico que el bloque antiguo de identidad fue reemplazado en las cuatro
   pantallas.

#### Configuracion / despliegue

- Se mantiene `sleepApplication: true`.
- Telegram queda fuera del alcance y no existe integracion activa.

#### Resultado y pendientes

- El sistema ya evita iniciar sin persona y evita descargar repetidos sin una
   decision explicita.
- La auditoria detallada por ejecucion, factura, intento y error aun no esta
   conectada completamente a los cuatro bots.

## Regla de mantenimiento de esta documentacion

Este README es la fuente oficial de documentacion e historial tecnico del
proyecto. Cada cambio futuro debe:

1. verificarse contra el estado real del repositorio;
2. implementarse y validarse;
3. revisar si afecta Neon, Railway, configuracion o dependencias;
4. agregar una entrada cronologica a este historial sin borrar entradas
    anteriores;
5. marcar como `Pendiente de verificacion` todo lo que no pueda comprobarse.

Ningun cambio se considera completo hasta actualizar este README.
