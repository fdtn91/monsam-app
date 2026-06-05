# MONSAM App — Estado del Proyecto

## Stack
- **Electron** (frame:false)
- **better-sqlite3 v9.4.3** (compatible con Node 16 + electron-rebuild)
- **Three.js** v0.160 (CDN) para render 3D de STL
- **XLSX** para catálogo de STL (Excel)
- Node.js v16.13.2

## Arquitectura
- `main.js` — proceso principal, handlers IPC, SQLite
- `preload.js` — bridge contextBridge → renderer
- `index.html` — UI completa (HTML + CSS + JS vanilla, UN solo bloque script)
- `db.js` — schema SQLite compartido con e500-app
- `config.json` — configuración local

## Base de datos SQLite
- Ruta: `F:\DISEÑOS\Modelos 3D\control\datos.db` (COMPARTIDA con e500-app)
- Campo en config.json: `rutaDB`
- Tablas: `filamentos`, `inventario`, `clientes`, `costos`
- `filamentos` es propiedad de e500-app — monsam solo lee
- Polling de stock cada 10 segundos via `refresh-stock`

## Carpetas importantes
- `rutaAretes`: `F:\DISEÑOS\Modelos 3D\aretes`
- Fotos: `F:\DISEÑOS\Modelos 3D\aretes\fotos\{CODIGO}.jpg`
- STLs: `F:\DISEÑOS\Modelos 3D\aretes\{CARPETA}\{archivo}.stl`
- Excel catálogo: `F:\DISEÑOS\Modelos 3D\aretes\Catalogo_Aretes_3D.xlsx`

## Código de colores (SKU)
- Se genera desde la tabla `filamentos` de la DB compartida
- Regla: primeras 3 letras del color puro (ignorando tipo y marca)
- Palabras ignoradas: PLA, PETG, ABS, TPU, SUNLU, ESUN, BAMBU, MATTE, etc.
- Ejemplos: AMARILLO → `AMA1`, ROJO → `ROJ1`, segundo ROJO → `ROJ2`
- SKU completo: `MON-AR-CIR1-AMA1`

## Estado actual (Junio 2026)
- ✅ SQLite funcionando, DB compartida con e500-app
- ✅ Catálogo con toggle vista 3D / Foto
- ✅ Drag & drop de fotos en cada tarjeta del catálogo
- ✅ Modal viewer: foto ocupa todo el ancho, limpia imagen anterior al abrir
- ✅ Inventario con totales, filtros por modelo/color, autocomplete de modelo con preview
- ✅ SKU genera código corto (3 letras) sin marca ni tipo
- ✅ Stock polling cada 10 segundos desde DB compartida

## Bugs corregidos importantes

### Foto del viewer mostraba modelo anterior
El `openViewer` no limpiaba las `<img>` previas, solo los `<canvas>`.
Solución: `wrap.querySelectorAll('img').forEach(i => i.remove())` al abrir.

### Dos bloques script separados
El catálogo estaba en un `<script>` y el estado global en otro — scope separado.
Solución: fusionar en un solo `<script>`, variables del catálogo en el estado global.

### SKU usaba nombre completo del color
`get-colores` devolvía `nombre as codigo` — el código era "PLA AMARILLO MATTE SUNLU".
Solución: generar código corto de 3 letras ignorando tipo y marca.

## config.json (monsam)
```json
{
  "marca": "MONSAM",
  "prefijoSKU": "MON",
  "tipoAccesorio": "AR",
  "rutaAretes": "F:\\DISEÑOS\\Modelos 3D\\aretes",
  "rutaExcel": "F:\\DISEÑOS\\Modelos 3D\\aretes\\Catalogo_Aretes_3D.xlsx",
  "rutaDB": "F:\\DISEÑOS\\Modelos 3D\\control\\datos.db",
  "costoKwh": 2.8,
  "wattsPrinter": 350,
  "costoHerrajes": 3.5,
  "porcDesperdicio": 5,
  "costoEmpaque": 2,
  "moonrakerIp": "192.168.68.113:7125"
}
```
⚠️ Siempre guardar UTF-8 sin BOM en Windows.

## Ramas Git
- Rama local: `main` ✅
- Remote: `origin/main` ✅

## Pendiente para próxima sesión
- Verificar que drag & drop funciona en producción (usa `file.path` de Electron)
- Revisar funcionalidad de clientes con SQLite
- Probar flujo completo: costo → inventario → venta → descuento stock
