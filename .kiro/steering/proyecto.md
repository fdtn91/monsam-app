# MONSAM App — Estado del Proyecto

## Stack
- **Electron** (frame:false)
- **better-sqlite3 v9.4.3** (compatible con Node 16 + electron-rebuild)
- **Three.js** v0.160 (CDN) para render 3D de STL
- **XLSX** para catálogo de STL (Excel)
- Node.js v16.13.2

## Arquitectura
- `main.js` — proceso principal, handlers IPC, SQLite, servidor HTTP puerto 4000
- `preload.js` — bridge contextBridge → renderer
- `index.html` — UI completa (HTML + CSS + JS vanilla, UN solo bloque script)
- `db.js` — schema SQLite compartido con e500-app
- `config.json` — configuración local

## Base de datos SQLite
- Ruta: `F:\DISEÑOS\Modelos 3D\control\datos.db` (COMPARTIDA con e500-app)
- Campo en config.json: `rutaDB`
- Tablas: `filamentos`, `inventario`, `clientes`, `costos`, `pedidos`, `modelos_nuevos`, `ventas`
- `filamentos` es propiedad de e500-app — monsam solo lee
- Polling de stock cada 10 segundos via `refresh-stock`

## Carpetas importantes
- `rutaAretes`: `F:\DISEÑOS\Modelos 3D\aretes`
- Fotos: `F:\DISEÑOS\Modelos 3D\aretes\fotos\{CODIGO}.jpg`
- STLs: `F:\DISEÑOS\Modelos 3D\aretes\{CARPETA}\{archivo}.stl`
- Excel catálogo: `F:\DISEÑOS\Modelos 3D\aretes\Catalogo_Aretes_3D.xlsx`

## API REST (puerto 4000) — endpoints para app móvil
- GET  /api/ping
- GET  /api/catalogo
- GET  /api/foto/:codigo
- GET  /api/colores
- GET  /api/inventario
- GET  /api/precio-venta
- GET  /api/config  ← datos del ticket (marca, RFC, dirección, etc.)
- POST /api/pedidos
- GET  /api/pedidos
- PUT  /api/pedidos/:id/estado
- POST /api/ventas  ← registra venta y descuenta inventario

## Código de colores (SKU)
- Primeras 3 letras del color puro (ignorando tipo y marca)
- Palabras ignoradas: PLA, PETG, ABS, TPU, SUNLU, ESUN, BAMBU, MATTE, etc.
- Ejemplos: AMARILLO → `AMA1`, ROJO → `ROJ1`, segundo ROJO → `ROJ2`
- SKU completo: `MON-AR-CIR1-AMA1`

## Estados de Pedidos
- `pendiente` → `visto` → `en_revision` → `en_produccion` → `listo` → `entregado`
- Estados válidos en IPC y API REST

## Ticket de venta
- Campos en config.json: `ticketDireccion`, `ticketTelefono`, `ticketRazonSocial`, `ticketRFC`, `ticketLeyenda`
- IVA 16% solo si hay RFC configurado
- Descuento en % por ticket (modificable)
- Impresión via `window.open` + `window.print()`
- Endpoint `/api/config` expone datos del ticket al celular

## Estado actual (Junio 2026)
- ✅ SQLite funcionando, DB compartida con e500-app
- ✅ Catálogo con toggle vista 3D / Foto, drag & drop de fotos
- ✅ Inventario con totales, filtros, autocomplete con preview
- ✅ SKU genera código corto (3 letras) sin marca ni tipo
- ✅ Stock polling cada 10 segundos desde DB compartida
- ✅ Sección Pedidos con estados, revisar, agrupar por color/placa
- ✅ Agrupar pedidos divide en placas por pares totales (maxModelosPlaca)
- ✅ Botón "Nuevo pedido (PC)" — venta inmediata, estado Listo
- ✅ Sección Ventas con ticket, descuento %, IVA automático, impresión
- ✅ Ventas descuentan inventario al cerrar ticket
- ✅ Campos de ticket en Configuración (dirección, RFC, razón social, leyenda)
- ✅ Cargar pedido terminado en ticket de venta
- ✅ App móvil conectada via WiFi local

## Pendiente para próxima sesión
- Configurar `eas build:configure` en monsam-mobile para generar APK
- El usuario ya tiene cuenta en expo.dev
- Usar icono del repo de monsam-app para el APK
- Agregar `eas.json` con perfil `preview` para APK de prueba

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
  "moonrakerIp": "192.168.68.113:7125",
  "precioVentaPar": 0,
  "ticketDireccion": "",
  "ticketTelefono": "",
  "ticketRazonSocial": "",
  "ticketRFC": "",
  "ticketLeyenda": ""
}
```

## Ramas Git
- Rama local: `main` ✅
- Remote: `origin/main` ✅
