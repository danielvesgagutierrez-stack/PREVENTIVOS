# Preventivos UDES — prototipo funcional

App para el celular que reemplaza el "pintar de colores" en el Excel.

**Flujo nuevo:** subes el Excel → si tiene varias hojas, eliges la del mantenimiento → recorres las dependencias escaneando el código de barras → registras qué encontraste, horas y observaciones → al terminar generas un **Excel con el mismo formato del original**, listo para entregar.

| Estado | Color | Cuándo |
|---|---|---|
| `OK` | 🟢 verde | Hecho, sin novedades |
| `NOVEDAD` | 🟠 naranja | Pantalla partida, etc. (texto obligatorio y foto opcional) |
| `AUSENTE` | 🔴 rojo | No se encuentra en la dependencia |
| `ALTA` | 🔵 azul | Está en otra dependencia o es equipo nuevo |

## Reglas del informe final

- **Fecha ejecutada (G):** se llena con la fecha actual automáticamente.
- **Fecha seguimiento (K):** queda **en blanco** (la llena otra persona).
- **Tiempo ejecución (J):** se registra por equipo; la app muestra el total en el resumen.
- **Observaciones (L):** "sin novedades", la novedad escrita, "no se encuentra…" o "Reasignado desde…".
- **Cambio de dependencia:** si escaneas en un área distinta a la que aparece en la hoja, la app te avisa; al marcar en azul actualiza edificio/piso/oficina en el informe.
- Se conservan las columnas A-L del archivo original y se colorea el número de inventario según el estado.
- **Número de inventario en negro:** el fondo de la celda toma el color del estado, pero la fuente siempre se fuerza a negro (tanto al subir el cronograma como al guardar).
- **Equipos nuevos en su área:** al dar de alta o reasignar, el equipo se inserta justo después del último registro de su misma dependencia (por ejemplo junto a *Antiguo Masira*), no al final. El JS calcula el final real de cada área incluso cuando edificio/piso/oficina vienen en celdas combinadas verticalmente, y rellena todos los datos del alta (serial, tipo, fecha, horas, observaciones y color) en la fila nueva.
- **Escritura verificada:** el backend obtiene las combinaciones reales con `getMergedRanges()` y parte **solo la combinación completa** que toca la fila que se va a escribir (A, B, C, F, G y K son columnas combinadas en el cronograma). Después vuelve a unir el resto para conservar el formato, escribe serial, tipo, fecha, horas, observaciones y fondo, y comprueba que el serial **y el color** quedaron grabados. Si algo falla, lo registra y lo muestra en la app.
- **Diagnóstico visible:** en *Datos → Google Sheets* aparecen la **versión del script**, la hora del último envío, cuántas filas se escribieron y los errores reales de Apps Script. Si la versión dice «sin respuesta» o es anterior a `2026-11-20d`, el Web App sigue ejecutando código viejo: hay que crear una **nueva versión** en Implementar → Administrar.
- **Sin filas dobles:** si un sync anterior dejó una fila en blanco, el siguiente registro la reutiliza en vez de insertar otra. Tampoco se inserta nada al escanear/buscar un serial que ya está registrado.
- **Escanear no depende del "área actual":** al leer un código la app solo consulta si el serial está en el cronograma y muestra a qué área pertenece según la hoja. Ya no muestra el aviso de "márcalo en azul" por comparar con el área desde la que se entró. Si el serial no está, se ofrece darlo de alta eligiendo la dependencia donde se encontró.

## Estadísticas y pendientes reales

- Los seriales borrados, vacíos, "SS" o no numéricos **no cuentan** como equipo: se listan en **Alertas** y se excluyen de totales y porcentajes.
- El porcentaje solo avanza con lo que se grabó con **Guardar y seguir** (o con un cambio real traído del Sheet). Los colores que trae el Excel original son el **cronograma**, no un registro: una celda sin registro cuenta como **pendiente** aunque el Excel venga coloreado o tenga fecha/semestre rellenos.
- En el envío al Sheet, los pendientes viajan como `PENDIENTE` pero el script **no escribe nada en esas filas**: no les cambia fecha, horas, observaciones ni color. Solo se tocan las filas registradas.

## Estructura

| Archivo | Para qué |
|---|---|
| `index.html`, `styles.css`, `app.js` | La app (sin frameworks ni compilación) |
| `data/equipos.json` | Cronograma de ejemplo ya normalizado |
| `tools/normalize_excel.py` | Convierte un Excel nuevo en JSON desde la PC, con selección de hoja |
| `vendor/` | ZXing (barras/QR), SheetJS (XLSX), JsBarcode (etiquetas) y **ExcelJS** (informe final con estilos), locales para funcionar sin internet |
| `apps-script/Code.gs` | Backend opcional que mantiene una hoja `Informe` en Google Sheets |
| `sw.js`, `manifest.webmanifest`, `icons/` | Instalable como app y offline |
| `README.md` | Esta guía |

## 1) Probar en la PC

```powershell
cd preventivos-app
python -m http.server 8000
```

Abre <http://localhost:8000>. En `localhost` la cámara funciona. Para probar sin cámara usa el campo "Escribe el serial" (ej. `1020010500072`).

## 2) Usarlo en el celular

La cámara exige HTTPS. Publica la carpeta en **GitHub Pages / Netlify / Cloudflare Pages** (gratis) o expón el servidor con HTTPS (ngrok/localtunnel). Luego abre la URL en el celular → menú → **Agregar a pantalla de inicio**.

> No funciona abriendo el `index.html` con doble clic ni por HTTP a la IP de la PC.

## 3) Subir el cronograma

En Inicio o Datos: **Subir cronograma** → selecciona el `.xlsx`.

- La app detecta qué hojas tienen seriales. Si hay **una sola**, la usa directo. Si hay **varias**, muestra un listado y eliges la de tu mantenimiento (por ejemplo `2026-B`).
- El archivo original se guarda en el teléfono para que el informe final conserve su formato.
- Si el archivo no tiene seriales reconocibles, se muestran todas las hojas para que elijas.

Desde la PC también puedes generar el JSON:

```powershell
python tools/normalize_excel.py "CRONOGRAMA.xlsx" "data/equipos.json"
python tools/normalize_excel.py "CRONOGRAMA.xlsx" "data/equipos.json" --hoja=2026-B
```

Sin `--hoja`, si hay varias hojas con seriales, el script pregunta cuál usar.

## 4) Flujo diario

1. Elige la hoja del mantenimiento al subir el archivo.
2. Áreas → elige la dependencia (queda como área actual).
3. Escanear → la app vibra y abre el equipo con edificio, piso, oficina y tipo.
4. Toca el estado. Si es novedad, escribe qué pasó; si es cambio de dependencia, confirma dónde está; si no se encuentra, déjalo en rojo.
5. Ajusta el tiempo ejecutado (botones rápidos o texto libre) y toca **Guardar y seguir**.
6. Si el serial no existe → **Dar de alta** en azul.
7. Al terminar → **Generar informe final (.xlsx)** en Inicio o Datos.

Todo queda guardado en el teléfono; sincroniza cuando tengas internet si configuraste Google Sheets.

## 5) Google Sheets (opcional)

1. Abre **la hoja de Google Sheets donde quieres el informe** (puede ser la del cronograma) → **Extensiones → Apps Script**. Eso asocia el script a ese archivo. Si lo creas en <https://script.google.com> como proyecto independiente, `setup()` creará un spreadsheet nuevo.
2. Borra el `myFunction()` que aparece y pega todo `apps-script/Code.gs`. Guarda.
3. En el selector de funciones elige `setup()` y pulsa **Ejecutar**. Autoriza (Revisar permisos → Configuración avanzada → Ir al proyecto). `setup()` solo guarda el ID del Sheet: **no crea hojas auxiliares**.
4. **Implementar → Nueva implementación → Aplicación web**. Ejecutar como: **Yo**. Acceso: **Cualquier persona con el enlace**. Copia la URL `https://script.google.com/macros/s/.../exec`.
5. En la app (APK) ve a **Datos** → pega la URL en *URL del Web App*, confirma la clave `udes2026` (si la cambias, cámbiala también en `TOKEN` de `Code.gs`) y escribe tu nombre en *Técnico* → **Guardar datos**.

### Al subir el cronograma

Con la URL guardada, al elegir la hoja del Excel la app:

1. **Borra TODAS las hojas** que tenga actualmente el Google Sheet.
2. Crea **una sola hoja** con el nombre de la hoja elegida (por ejemplo `2026-B`).
3. Copia la hoja **tal cual**: valores, fechas, colores de fondo, contraste de fuente, celdas combinadas del encabezado, anchos de columna, altos de fila y fila congelada.
4. Aplica encima **únicamente** los equipos que registraste con *Guardar y seguir* (al subir un cronograma nuevo el progreso se reinicia).

Si no hay internet en ese momento, la carga queda pendiente y se hace automáticamente al volver la señal o al guardar la configuración.

> Las filas que no has registrado **no se tocan**: no se les escribe fecha, horas, observaciones ni color. Por eso el diseño deja de cambiar cada vez que entras o se corre el auto-sync. Solo se descombinan las filas que sí se registraron.

### Sincronización automática (bidireccional)

Con la URL guardada:

- **Al dar *Guardar y seguir*** (o dar de alta un equipo) se envía **solo ese registro** y se trae lo último del Sheet.
- **Al escanear o buscar un serial** se reenvían pendientes y se traen cambios.
- **Al abrir la APK**, **al volver la señal**, **al volver del background** y **cada 2 minutos**, la app envía la jornada completa y **trae los cambios** del Sheet.

El script busca el número de inventario en la pestaña creada y:

- pinta la celda del serial con el color del estado (verde/ámbar/rojo/azul);
- deja el **número en negro**;
- escribe fecha ejecutada, técnico, semestre, horas y observaciones;
- deja `FECHA SEGUIMIENTO` vacío y conserva `FECHA PROGRAMADA`.

Al traer cambios, gana el registro **más reciente** (se compara la fecha). La pestaña **Datos** muestra la última sincronización y un diagnóstico tipo *"Sheet leído: 231 con color · 231 con datos"*.

Si no hay internet, el registro queda en cola en el teléfono y se envía solo al volver la señal. La pestaña **Datos** muestra cuántos registros faltan por enviar.

> La app no lee el Sheet con CORS directo: usa JSONP (`<script>`) porque Apps Script no envía cabeceras CORS. Por eso el despliegue debe ser **"Cualquier persona con el enlace"** y el token debe coincidir. Las hojas viejas `Informe`, `Cambios`, `Resumen` y `_estados` ya no se usan: se eliminan al subir el cronograma.

## 6) Etiquetas

Datos → *Imprimir etiquetas* → filtra por edificio → Generar → imprime en papel adhesivo. Son Code128 con el serial exacto.

## 7) Hacer un APK para el celular

GitHub Actions **no tiene un botón mágico** para compilar APK: solo ejecuta flujos que tú defines en `.github/workflows/*.yml`. Ya dejé uno listo en este proyecto.

### Opción A: GitHub Actions (APK con Capacitor)

1. Sube a GitHub **TODOS los archivos de la app**, no solo el workflow: `index.html`, `app.js`, `styles.css`, `manifest.webmanifest`, `sw.js`, `version.js` y las carpetas `vendor/`, `data/`, `icons/`, además de `.github/workflows/build-apk.yml`. Si falta alguno, el workflow falla a propósito con un mensaje claro.
   - Si tu editor no muestra carpetas ocultas, crea el workflow desde la web: **Add file → Create new file** y en el nombre escribe literalmente `.github/workflows/build-apk.yml`.
   - Debe quedar en la **raíz del repositorio** y en la rama por defecto (`main`/`master`). Si subiste toda la app dentro de una carpeta `preventivos-app/`, el workflow también la detecta.
2. En GitHub entra a la pestaña **Actions**.
3. A la izquierda selecciona **Compilar APK** → botón **Run workflow** → **Run workflow**.
4. Espera 3-6 minutos. Al terminar, en la ejecución aparece **Artifacts → preventivos-apk**.
5. Descarga el `.zip`, extráelo y copia `app-debug.apk` al celular. Ábrelo e instala (hay que permitir "instalar aplicaciones desconocidas").

El APK trae la app completa dentro (no depende de internet), usa la cámara y guarda los datos en el teléfono. Si quieres actualizarlo, repite el paso 3 **después de haber subido los archivos modificados**.

### Cómo saber si el APK instalado es realmente el nuevo

Entra a **Datos → Diagnóstico**. Debe aparecer algo como:

```
App instalada: 2026-11-20-v26 · a1b2c3d · 2026-11-20T19:35Z
```

El texto después del `·` es el **commit de GitHub con el que se compiló**. Si dice `local` o un commit anterior, el celular todavía tiene la APK vieja o el navegador sirvió una copia en caché. En ese caso toca **↻ Forzar actualización (borrar caché y recargar)** y reinstala el APK.

> **Android conserva `localStorage` al actualizar una APK.** Por eso los cronogramas y registros de pruebas pueden seguir apareciendo aunque el código sea nuevo. En esta versión las claves de guardado cambiaron a `_v4`: la app nueva arranca con el cronograma que viene dentro del paquete y los datos antiguos se quedan guardados (no se borran) con los nombres `_v3`. Para volver a empezar en el celular, entra a **Datos → Olvidar cronograma** o reinstala borrando los datos de la app desde Ajustes de Android.

> Para subirlo a Google Play hace falta un APK/AAB **firmado** con tu propia llave y una cuenta de desarrollador; este flujo genera un `debug` listo para uso personal.

### Opción B: PWABuilder (sin tocar Actions)

1. Abre <https://www.pwabuilder.com/>, pega la URL HTTPS de tu PWA publicada y haz clic en **Start**.
2. Elige **Package for app stores → Android → Generate Android package**.
3. Descarga el APK y el archivo `assetlinks.json`.
4. Sube `assetlinks.json` a `https://TU-USUARIO.github.io/.well-known/assetlinks.json` (o a la raíz de tu dominio) para que se instale a pantalla completa sin barra del navegador. Con GitHub Pages en una subcarpeta (`/repo/`) ese archivo debe ir al repositorio `TU-USUARIO.github.io`.

## Límites conocidos

- Las fotos se guardan solo en el teléfono; no se suben a Sheets.
- El informe final se genera con ExcelJS a partir del archivo original; si el archivo tiene macros o elementos muy especiales, puede perderse algún detalle decorativo (los datos y colores de la hoja se conservan).
- El cronograma original tiene columnas combinadas (por ejemplo `FECHA EJECUTADA` abarca 20 filas). El informe separa esas celdas para poder escribir fecha, horas y observaciones **por equipo**. `FECHA PROGRAMADA` y `FECHA SEGUIMIENTO` se dejan como estaban (esta última, en blanco).
- La sincronización envía el estado completo de la jornada, no solo lo cambiado.
- Pensado para un técnico; para varios en paralelo hay que añadir login y bloqueo por jornada.
