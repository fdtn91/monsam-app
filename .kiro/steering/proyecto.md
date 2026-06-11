# MONSAM App — Estado del Proyecto

## Stack
- **Electron** (frame:false)
- **better-sqlite3 v9.4.3**
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
- Colores degradados guardados como `hex1|hex2|hex3` en campo `color_hex`

## API REST (puerto 4000) — endpoints para app móvil
- GET  /api/ping
- GET  /api/catalogo
- GET  /api/foto/:codigo
- GET  /api/colores  ← devuelve hex con formato hex1|hex2|hex3 para degradados
- GET  /api/inventario
- GET  /api/precio-venta
- GET  /api/config
- POST /api/pedidos
- GET  /api/pedidos
- PUT  /api/pedidos/:id/estado
- POST /api/ventas

## Código de colores (SKU)
- `MON-AR-{MODELO}-{COLOR3LETRAS}` ejemplo: `MON-AR-CIR1-AMA1`

## Estados de Pedidos
- `pendiente` → `visto` → `en_revision` → `en_produccion` → `listo` → `entregado`
- Solo se puede avanzar de `pendiente` a `visto` como primer paso
- Cuando llega a `listo` aparece botón "📦 → Inventario" para agregar al inventario

## Funcionalidades recientes (Junio 2026)
- ✅ Botón "＋ SKU" en cada card del catálogo — agrega directo al inventario
- ✅ Inventario vista "Por modelo" agrupa por modelo y sub-agrupa por color
- ✅ Hover en pestaña Imprimir muestra preview de imagen del modelo
- ✅ Colores en grid de swatches (en lugar de lista)
- ✅ Soporte colores degradados bi/tricolor en modal de color
- ✅ Pedidos: botón "📦 → Inventario" cuando estado es "Listo"
- ✅ Bug fix: agregar SKU suma pares al existente en lugar de reemplazar
- ✅ Pedidos: solo puede avanzar de pendiente→visto como primer paso

## config.json
```json
{
  "marca": "MONSAM",
  "prefijoSKU": "MON",
  "tipoAccesorio": "AR",
  "rutaAretes": "F:\\DISEÑOS\\Modelos 3D\\aretes",
  "rutaExcel": "F:\\DISEÑOS\\Modelos 3D\\aretes\\Catalogo_Aretes_3D.xlsx",
  "rutaDB": "F:\\DISEÑOS\\Modelos 3D\\control\\datos.db",
  "moonrakerIp": "192.168.68.113:7125",
  "precioVentaPar": 0
}
```

## Ramas Git
- Repo: `fdtn91/monsam-app`
- Rama: `main`
