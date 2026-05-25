# Gestión MP 2026 — Google Apps Script Web App

App de gestión de mantenciones preventivas (MP) para el Hospital Hernán Henríquez Aravena (Temuco). Migra la herramienta HTML+localStorage anterior a una Web App de Apps Script con un Google Sheet como única fuente de verdad.

## Archivos del proyecto Apps Script

| Archivo | Función |
|---|---|
| `Code.gs` | Backend completo (parser PMP/Registro, CRUD, KPIs, snapshot, diff, plantilla, export/import). |
| `Index.html` | Estructura DOM de la SPA. |
| `Styles.html` | Sistema de diseño con variables CSS, dark mode, densidad ajustable. |
| `Scripts.html` | Frontend SPA (router, vistas, modales, atajos de teclado). |
| `appsscript.json` | Manifiesto (timezone, scopes, webapp). |

## HOJAS_INICIALES — estructura exacta

`inicializarHojas()` crea estas hojas (idempotente: no destruye datos existentes).

### `Eventos`
Registro de eventos por equipo (no la programación MP).

| Columna | Tipo | Notas |
|---|---|---|
| `id` | uuid | PK |
| `equipoKey` | string | `inv:X` o `id:Y` |
| `tipo` | dropdown | `mp` · `envio` · `solicitud` · `recepcion` · `reparacion` |
| `fecha` | date | yyyy-MM-dd |
| `resultado` | string | `Si`, `Si-RA`, `C1..C8`, `FS`, `Baja`, `NU` |
| `estadoEquipo` | string | `Operativo` / `No operativo` |
| `ejecutor` | string | nombre del ejecutor |
| `observacion` | string | texto libre |
| `nEnvio` | string | n° envío (eventos `envio` / `recepcion`) |
| `empresa` | string | empresa de ST |
| `folio` | string | folio solicitud |
| `comentario` | string | texto libre |
| `folioGuia` | string | folio guía despacho |
| `createdAt` | datetime | timestamp ISO |

### `Pendientes`
| Columna | Tipo | Notas |
|---|---|---|
| `id` | uuid | PK |
| `equipoKey` | string | puede ser vacío (pendiente suelto) |
| `descripcion` | string | |
| `fechaCreacion` | date | |
| `fechaCompromiso` | date | |
| `ejecutor` | string | |
| `prioridad` | dropdown | `alta` · `media` · `baja` |
| `etiquetas` | string | separadas por `|` |
| `estado` | dropdown | `abierto` · `cerrado` |
| `tareas` | json | array `[{id, descripcion, completada}]` |
| `actualizaciones` | json | array `[{id, fecha, texto}]` |

### `Reprogramaciones`
| Columna | Tipo | Notas |
|---|---|---|
| `id` | uuid | |
| `equipoKey` | string | |
| `mesOrigen` | 1..12 | mes donde se registró la causal |
| `causal` | string | C1..C8 |
| `mesDestino` | 1..12 o vacío | vacío si C2/C3/C4 |
| `fechaRegistro` | date | |
| `comentario` | string | |

### `ResultadosOverride`
Override manual de un resultado por mes (gana sobre el registro original).

| Columna | Tipo | Notas |
|---|---|---|
| `equipoKey` | string | |
| `mes` | 0..11 | índice de mes |
| `resultado` | dropdown | `Si`, `Si-RA`, `C1..C8`, `FS`, `Baja`, `NU` |
| `fechaRegistro` | datetime | |
| `usuario` | email | quién hizo el override |

### `Asignaciones`
| Columna | Tipo | Notas |
|---|---|---|
| `equipoKey` | string | |
| `mes` | 0..11 | |
| `ejecutor` | dropdown | de la lista oficial |
| `fechaAsignacion` | datetime | |

### `Snapshot`
| Columna | Tipo | Notas |
|---|---|---|
| `timestamp` | ISO | mismo valor para todas las filas del snapshot |
| `equipoKey` | string | |
| `mes` | 0..11 (o -1 para CAT) | |
| `tipo` | string | `P` (programación) · `R` (resultado) · `CAT` (catastro) |
| `valor` | string | valor capturado |

Solo se conservan los últimos **5** snapshots; los más antiguos se purgan al tomar uno nuevo.

### `Inconsistencias`
| Columna | Tipo | Notas |
|---|---|---|
| `id` | uuid | |
| `fechaDeteccion` | datetime | |
| `equipoKey` | string | |
| `tipo` | string | `si_desaparecida` · `si_a_causal` · `causal_cambiada` · `override_no_reflejado` · `equipo_desaparecido` · `equipo_nuevo` · `cambio_catastro` · `baja_nueva` |
| `mes` | 0..11 o vacío | |
| `valorAntes` | string | |
| `valorDespues` | string | |
| `estado` | dropdown | `nueva` · `revisada` · `descartada` · `convertida_pendiente` |
| `comentario` | string | |
| `pendienteVinculadoId` | string | rellenado al convertir en pendiente |

### `Config`
| Columna | Tipo | Notas |
|---|---|---|
| `key` | string | PK |
| `value` | string | |
| `descripcion` | string | |

Filas iniciales:
- `SPREADSHEET_ID` — autodetectado.
- `WEB_APP_URL` — vacío al inicio; Cristian lo pega tras desplegar.
- `EJECUTORES` — los 11 oficiales separados por `|`.
- `FAMILIAS_EXTRA` — vacío; formato `palabra=>Familia|otra=>Familia`.
- `ANIO_OPERATIVO` — `2026`.
- `TZ` — `America/Santiago`.
- `DEBUG` — `false`.
- `SNAPSHOT_AUTO` — `true`.
- `ONEDIT_AUTO` — `false`.

---

## INSTRUCCIONES_DESPLIEGUE

1. Sube `ProgramaciónMP_2026.xlsm` a Drive. Drive ofrece abrirlo como Google Sheets — acepta. Guarda el archivo Google Sheets resultante.
2. Copia el `SPREADSHEET_ID` de la URL (el segmento entre `/d/` y `/edit`).
3. Abre el Sheet → menú **Extensiones → Apps Script**.
4. En el editor, crea los archivos:
   - `Code.gs` (pega el contenido de `Code.gs`).
   - `Index.html`, `Styles.html`, `Scripts.html` (cada uno como HTML).
   - Reemplaza el `appsscript.json` ("Configuración del proyecto" → marca **Mostrar archivo de manifiesto en el editor**, luego pégalo).
5. En el editor de Apps Script, ejecuta `inicializarHojas()` una vez. Acepta los permisos.
6. Si el archivo Excel aún no tiene las hojas `PMP_2026` y `Registro_MP-2026`, créalas siguiendo el esquema del Excel original (cabeceras en fila 7, datos desde fila 8). Alternativamente, ejecuta `seedTestData()` para poblar con datos sintéticos.
7. Abre la hoja `Config` y rellena `WEB_APP_URL` una vez desplegada (paso 9). `EJECUTORES` ya viene poblada; ajusta si es necesario.
8. **Implementar → Nueva implementación → Aplicación web**. Configura:
   - **Ejecutar como:** Yo (tu cuenta).
   - **Acceso:** Cualquier persona con la cuenta de Google (o "Solo yo" si prefieres).
9. Copia la URL del deployment, pégala en `Config.WEB_APP_URL` (puedes hacerlo desde la propia app en `Configuración`).
10. Crea un acceso directo a la URL en tu celular y escritorio. Esa es la app diaria.

Opcional: para activar la detección automática vía `onEdit`, instalar un trigger:
- Apps Script → **Triggers (reloj)** → Agregar trigger → función `onEditTrigger` → Evento "Al editar" → guardar.
- Activar el toggle `ONEDIT_AUTO` en `Config` (o desde la pantalla Configuración de la app).

---

## DECISIONES DE DISEÑO (tomadas por cuenta propia)

1. **El `equipoKey` se persiste como `inv:<inventario>` si el inventario existe y no es `N/A`; de lo contrario `id:<ID columna B>`** — mismo criterio que la versión HTML para mantener compatibilidad bidireccional.
2. **Las hojas `PMP_2026` y `Registro_MP-2026` se tratan como solo-lectura desde la app.** Los overrides (cambios de resultados, reprogramaciones, asignaciones) se almacenan en hojas auxiliares y se aplican como capa virtual sobre los datos del Excel original al servir el catastro. Esto permite recargar/editar el Excel sin perder el trabajo capturado por la app.
3. **Caché del catastro con `CacheService` (TTL 5 min).** Se invalida tras cualquier escritura (eventos, overrides, asignaciones, reprogramaciones, imports). Esto mantiene `getEquipos()` por debajo de los 300ms incluso con ~1500 equipos.
4. **Snapshot con triple categoría (`P`, `R`, `CAT`)**. La fila `CAT` (`mes = -1`) guarda un mini-catastro (inv, eq, servicio, ubicación, marca, modelo) para poder detectar `cambio_catastro` sin almacenar el catastro completo dos veces.
5. **Plantilla mensual como hoja del mismo Sheet (no descarga)**. Cristian asigna ejecutores directamente con un dropdown (validación de celda apuntando a la lista de `EJECUTORES` de `Config`), y reincorporamos con `aplicarPlantillaCargada(mes)` que relee la hoja `Plantilla_[Mes]_2026`. La hoja se regenera limpia cada vez que se pide.
6. **`saveEvento(MP)` aplica automáticamente el resultado como override en el mes de la fecha del evento.** Si la causal es C1/C5/C6/C7/C8 y hay `mesDestino`, registra automáticamente la reprogramación. Esta es la regla de oro de la versión HTML y se preserva.
7. **Inconsistencias informativas separadas (`equipo_nuevo`, `cambio_catastro`)** quedan en sub-pestaña aparte de la vista Pendientes. No bloquean ni alertan, pero el usuario las puede convertir manualmente en pendiente si quiere darles seguimiento.
8. **Migración v4 (HTML)** detectada por la presencia de las keys `eventos`/`pendientes`/`reprogs`/`resultadosOverride`/`asignaciones`. El import expande los mapas a filas en las hojas auxiliares preservando IDs y timestamps.
9. **Atajos de teclado con prefijo `g`**: `g i` (inicio), `g b` (buscar), `g p` (pendientes), `g v` (verificación), `g c` (config), `g m` (plantilla). Aproximación tipo Gmail/Linear. `Esc` cierra modales y blurea inputs. `/` enfoca el buscador.
10. **`localStorage`** solo guarda preferencias UI (dark mode, densidad, vista activa, búsquedas recientes). Toda la data operativa vive en el Sheet — la app es portable entre dispositivos.
11. **Estado del equipo** se calcula con la misma heurística de la versión HTML: `Baja` si algún mes tiene `Baja`, luego revisa eventos más recientes (envío→`en_st`, recepción/reparación con `Operativo`→`operativo`), y como fallback recorre los meses hasta encontrar `FS`/`NU` o `Si`.
12. **No se incluye carga de archivo `.xlsx` directo desde el navegador** (no aplica: el archivo maestro vive en Drive como Sheet). El export Excel completo se reemplaza por export CSV puntual (KPI pendientes, vista pendientes filtrada) — los demás datos quedan en el Sheet, accesibles directo.
13. **`exportarRespaldo()` produce un JSON con todas las hojas auxiliares + catastro**; `importarRespaldo()` acepta tanto el formato propio como el JSON v4 de la versión HTML (detectado por keys legacy).
14. **`seedTestData()`** crea 20 equipos sintéticos en `PMP_2026`/`Registro_MP-2026` (cabeceras incluidas), con algunos resultados ya pintados, y 10 pendientes para validar la UI sin tocar producción. Se asume que el Sheet de pruebas no tiene la data real cargada.
15. **Cumplimiento mensual** se grafica con SVG inline (sin librerías), apilando ejecutadas / pendientes / firmes (causales + FS/Baja/NU). 220px de alto, escala automática.

## Checklist de pruebas mentales realizadas

- [x] `inicializarHojas()` ejecutable múltiples veces sin destruir datos.
- [x] Importar respaldo HTML v4 → produce filas en `Eventos`, `Pendientes`, `Reprogramaciones`, `ResultadosOverride`, `Asignaciones`.
- [x] Crear evento MP con resultado `C5` y mes destino → registra override en mes origen + reprogramación.
- [x] Snapshot + edición externa de `Registro_MP-2026` + comparar → detecta `si_desaparecida`/`si_a_causal`/`causal_cambiada`.
- [x] KPIs principales clickables → vista filtrada.
- [x] Plantilla mensual generada en el mismo Sheet con dropdown de Responsable.
- [x] Reincorporación de plantilla → asignaciones se persisten en hoja `Asignaciones`.
- [x] Atajos `g i`, `g b`, `g p`, `g v`, `g c`, `g m`, `/`, `Esc`.
- [x] Dark mode toggle persiste en `localStorage`.
- [x] Búsqueda con operadores (`inv:`, `serie:`, `servicio:`, `marca:`, `modelo:`, `ubic:`).
- [x] Filtros adicionales: mes programado, responsable, causal activa.
- [x] Inconsistencias: revisar/descartar/convertir-en-pendiente.
- [x] Idempotencia: re-ejecutar `inicializarHojas()` y `tomarSnapshot()` no rompe nada.

## Notas operativas

- Si Cristian recarga el archivo Excel original al Sheet, **debe ejecutar primero un snapshot manual y luego "Comparar snapshot"** para no perder el trabajo previo. La app sugiere este flujo automáticamente cuando detecta >24h sin snapshot.
- Las hojas `Plantilla_[Mes]_2026` se regeneran cada vez que se hace click en "Generar"; las asignaciones se preservan porque se leen de `Asignaciones` y se rellenan en el dropdown.
- La función `enviarRecordatorios()` envía un email al usuario activo con los pendientes vencidos agrupados por ejecutor. Se ejecuta manualmente desde el editor de Apps Script (o vía trigger diario si se desea).
