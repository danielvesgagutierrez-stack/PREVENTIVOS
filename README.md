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

1. Crea un proyecto en <https://script.google.com> y pega `apps-script/Code.gs`.
2. Ejecuta `setup()` una vez y autoriza.
3. Implementar → Nueva implementación → **Aplicación web**, ejecutar como **Yo**, acceso **Cualquier persona con el enlace**.
4. Pega la URL en Datos, confirma la clave `udes2026` (cámbiala en ambos archivos si quieres) y escribe tu nombre en *Técnico*.
5. **Sincronizar ahora** crea/actualiza la hoja `Informe` con columnas A-L y colores, más `Cambios` y `Resumen`.

## 6) Etiquetas

Datos → *Imprimir etiquetas* → filtra por edificio → Generar → imprime en papel adhesivo. Son Code128 con el serial exacto.

## 7) Hacer un APK para el celular

GitHub Actions **no tiene un botón mágico** para compilar APK: solo ejecuta flujos que tú defines en `.github/workflows/*.yml`. Ya dejé uno listo en este proyecto.

### Opción A: GitHub Actions (APK con Capacitor)

1. Sube a GitHub **un solo archivo**: `.github/workflows/build-apk.yml`.
   - Si tu editor no muestra carpetas ocultas, crea el archivo desde la web: **Add file → Create new file** y en el nombre escribe literalmente `.github/workflows/build-apk.yml`.
   - Debe quedar en la **raíz del repositorio** y en la rama por defecto (`main`/`master`). Si subiste toda la app dentro de una carpeta `preventivos-app/`, el workflow también la detecta.
2. En GitHub entra a la pestaña **Actions**.
3. A la izquierda selecciona **Compilar APK** → botón **Run workflow** → **Run workflow**.
4. Espera 3-6 minutos. Al terminar, en la ejecución aparece **Artifacts → preventivos-apk**.
5. Descarga el `.zip`, extráelo y copia `app-debug.apk` al celular. Ábrelo e instala (hay que permitir "instalar aplicaciones desconocidas").

El APK trae la app completa dentro (no depende de internet), usa la cámara y guarda los datos en el teléfono. Si quieres actualizarlo, repite el paso 3 con los cambios ya subidos.

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
