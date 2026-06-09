const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path    = require('path')
const fs      = require('fs')
const http    = require('http')
const { spawn } = require('child_process')
const XLSX    = require('xlsx')
const { getDB, migrateFromExcel } = require('./db')

// ╔══════════════════════════════════════════════════════════════════════════╗
//  CONFIG
// ╚══════════════════════════════════════════════════════════════════════════╝
const CONFIG_PATH = path.join(__dirname, 'config.json')

function loadConfig () {
  try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')) }
  catch { return {} }
}
function saveConfig (cfg) {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2), 'utf8')
}

ipcMain.handle('get-config', () => loadConfig())
ipcMain.handle('save-config', (_, cfg) => { saveConfig(cfg); return true })

// ╔══════════════════════════════════════════════════════════════════════════╗
//  DB — inicializar al arrancar
// ╚══════════════════════════════════════════════════════════════════════════╝
let db = null

function initDB () {
  const cfg    = loadConfig()
  const dbPath = cfg.rutaDB || path.join(__dirname, 'datos.db')
  db = getDB(dbPath)
  const migrated = migrateFromExcel(db, cfg)
  const total = Object.values(migrated).reduce((a, b) => a + b, 0)
  if (total > 0) console.log('Migración desde Excel completada:', migrated)
  // Migraciones incrementales — tablas que pueden no existir en DB antiguas
  db.exec(`
    CREATE TABLE IF NOT EXISTS modelos_nuevos (
      codigo    TEXT PRIMARY KEY,
      carpeta   TEXT DEFAULT '',
      archivo1  TEXT DEFAULT '',
      added_at  TEXT DEFAULT (datetime('now','localtime'))
    );
    CREATE TABLE IF NOT EXISTS pedidos (
      id           INTEGER PRIMARY KEY AUTOINCREMENT,
      fecha_pedido TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
      cliente      TEXT    DEFAULT '',
      notas        TEXT    DEFAULT '',
      items        TEXT    NOT NULL DEFAULT '[]',
      estado       TEXT    NOT NULL DEFAULT 'pendiente',
      synced_at    TEXT    DEFAULT NULL,
      created_at   TEXT    DEFAULT (datetime('now','localtime'))
    );
    CREATE INDEX IF NOT EXISTS idx_ped_estado ON pedidos (estado);
    CREATE INDEX IF NOT EXISTS idx_ped_fecha  ON pedidos (fecha_pedido);
    CREATE TABLE IF NOT EXISTS ventas (
      id         INTEGER PRIMARY KEY AUTOINCREMENT,
      fecha      TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
      cliente    TEXT    DEFAULT '',
      items      TEXT    NOT NULL DEFAULT '[]',
      total      REAL    DEFAULT 0,
      notas      TEXT    DEFAULT '',
      origen     TEXT    DEFAULT 'pc',
      created_at TEXT    DEFAULT (datetime('now','localtime'))
    );
    CREATE INDEX IF NOT EXISTS idx_ven_fecha ON ventas (fecha);
  `)
  return db
}

// ╔══════════════════════════════════════════════════════════════════════════╗
//  VENTANA
// ╚══════════════════════════════════════════════════════════════════════════╝
function createWindow () {
  const win = new BrowserWindow({
    width: 1200, height: 760,
    minWidth: 960, minHeight: 640,
    frame: false,
    show: false,
    backgroundColor: '#F0F2F5',
    icon: path.join(__dirname, 'icono.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  })
  win.loadFile('index.html')
  win.once('ready-to-show', () => win.show())
}

app.whenReady().then(() => {
  initDB()
  createWindow()
  startApiServer()
})
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })

ipcMain.on('win-minimize', e => BrowserWindow.fromWebContents(e.sender).minimize())
ipcMain.on('win-maximize', e => {
  const w = BrowserWindow.fromWebContents(e.sender)
  w.isMaximized() ? w.unmaximize() : w.maximize()
})
ipcMain.on('win-close', e => BrowserWindow.fromWebContents(e.sender).close())

// ╔══════════════════════════════════════════════════════════════════════════╗
//  API REST — servidor HTTP para la app móvil
//  Puerto: 4000 (configurable en config.json → apiPort)
//
//  GET  /api/ping                → { ok, version }
//  GET  /api/catalogo            → [ { carpeta, codigo, archivo1, tieneFoto } ]
//  GET  /api/foto/:codigo        → imagen binaria (jpg/png/webp)
//  GET  /api/colores             → [ { id, nombre, hex, stockGr } ]
//  POST /api/pedidos             → guarda pedido { cliente, notas, items, fecha_pedido }
//  GET  /api/pedidos             → lista pedidos
//  PUT  /api/pedidos/:id/estado  → { estado } actualiza estado
// ╚══════════════════════════════════════════════════════════════════════════╝
let apiServer = null

function startApiServer () {
  const cfg  = loadConfig()
  const port = cfg.apiPort || 4000

  apiServer = http.createServer((req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*')
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,OPTIONS')
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type')

    if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return }

    const url   = new URL(req.url, `http://localhost:${port}`)
    const route = url.pathname

    const json = (data, code = 200) => {
      res.writeHead(code, { 'Content-Type': 'application/json' })
      res.end(JSON.stringify(data))
    }
    const notFound = () => json({ error: 'Not found' }, 404)
    const bodyJSON = () => new Promise((resolve, reject) => {
      let raw = ''
      req.on('data', d => raw += d)
      req.on('end', () => {
        try { resolve(JSON.parse(raw || '{}')) }
        catch { reject(new Error('JSON inválido')) }
      })
    })

    try {
      // GET /api/ping
      if (req.method === 'GET' && route === '/api/ping') {
        return json({ ok: true, version: '1.0.0', app: 'MONSAN' })
      }

      // GET /api/catalogo
      if (req.method === 'GET' && route === '/api/catalogo') {
        const cfg2      = loadConfig()
        const rutaExcel = cfg2.rutaExcel || cfg2.rutaExcelMonsan
        if (!rutaExcel || !fs.existsSync(rutaExcel)) return json([])
        const wb   = XLSX.readFile(rutaExcel)
        const ws   = wb.Sheets[wb.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
        const rutaAretes = cfg2.rutaAretes || ''
        const items = []
        for (let i = 1; i < rows.length; i++) {
          const carpeta  = String(rows[i][0] || '').trim()
          const codigo   = String(rows[i][1] || '').trim()
          const archivo1 = String(rows[i][2] || '').trim()
          if (!codigo || codigo === 'TOTAL DE PARES') continue
          let tieneFoto = false
          if (rutaAretes) {
            const fotosDir = path.join(rutaAretes, 'fotos')
            for (const ext of ['.jpg', '.jpeg', '.png', '.webp']) {
              if (fs.existsSync(path.join(fotosDir, `${codigo}${ext}`))) { tieneFoto = true; break }
            }
          }
          let tieneStl = false
          if (rutaAretes && carpeta && archivo1) {
            tieneStl = fs.existsSync(path.join(rutaAretes, carpeta, archivo1))
          }
          items.push({ carpeta, codigo, archivo1, tieneFoto, tieneStl })
        }
        return json(items)
      }

      // GET /api/foto/:codigo
      if (req.method === 'GET' && route.startsWith('/api/foto/')) {
        const codigo     = decodeURIComponent(route.replace('/api/foto/', ''))
        const cfg2       = loadConfig()
        const rutaAretes = cfg2.rutaAretes || ''
        if (!rutaAretes) return json({ error: 'rutaAretes no configurada' }, 404)
        const fotosDir = path.join(rutaAretes, 'fotos')
        let fotoPath = null
        for (const ext of ['.jpg', '.jpeg', '.png', '.webp']) {
          const p = path.join(fotosDir, `${codigo}${ext}`)
          if (fs.existsSync(p)) { fotoPath = p; break }
        }
        if (!fotoPath) return notFound()
        const ext  = path.extname(fotoPath).toLowerCase()
        const mime = ext === '.png' ? 'image/png' : ext === '.webp' ? 'image/webp' : 'image/jpeg'
        res.writeHead(200, { 'Content-Type': mime })
        return fs.createReadStream(fotoPath).pipe(res)
      }

      // GET /api/colores
      if (req.method === 'GET' && route === '/api/colores') {
        const IGNORAR = /\b(PLA|PETG|ABS|TPU|ASA|NYLON|SILK|WOOD|METAL|SUNLU|ESUN|BAMBU|CREALITY|HATCHBOX|POLYMAKER|PRUSAMENT|BASICFIL|MEXICOMAKERS|MATTE|PLUS|PRO|MAX|LITE|BASIC)\b/gi
        const rows = db.prepare(`
          SELECT id, nombre, color_hex as hex, stock_gr as stockGr, notas
          FROM filamentos ORDER BY nombre
        `).all()
        return json(rows.map(r => ({
          ...r,
          nombre: r.nombre.replace(IGNORAR, '').replace(/\s+/g, ' ').trim()
        })))
      }

      // GET /api/inventario (para ventas desde móvil)
      if (req.method === 'GET' && route === '/api/inventario') {
        const rows = db.prepare('SELECT sku, modelo, color, pares, costo_produccion as costoProduccion FROM inventario WHERE pares > 0 ORDER BY sku').all()
        return json(rows)
      }

      // GET /api/precio-venta
      if (req.method === 'GET' && route === '/api/precio-venta') {
        const cfg2 = loadConfig()
        return json({ precio: cfg2.precioVentaPar || 0 })
      }

      // POST /api/ventas
      if (req.method === 'POST' && route === '/api/ventas') {
        bodyJSON().then(body => {
          const { cliente='', items=[], notas='' } = body
          if (!Array.isArray(items) || !items.length) return json({ ok:false, error:'items requerido' }, 400)
          const total  = items.reduce((s,i) => s + ((i.precio_par||0)*(i.pares||0)), 0)
          const result = db.prepare(`INSERT INTO ventas (fecha,cliente,items,total,notas,origen) VALUES (?,?,?,?,?,'mobile')`).run(new Date().toISOString(), cliente, JSON.stringify(items), total, notas)
          const descontar = db.transaction((items) => {
            for (const item of items) {
              if (!item.sku || !item.pares) continue
              db.prepare(`UPDATE inventario SET pares=MAX(0,pares-?), updated_at=datetime('now','localtime') WHERE sku=?`).run(item.pares, item.sku)
            }
          })
          descontar(items)
          BrowserWindow.getAllWindows().forEach(w => w.webContents.send('venta-nueva', { id: result.lastInsertRowid }))
          return json({ ok: true, id: result.lastInsertRowid, total })
        }).catch(e => json({ ok:false, error:e.message }, 400))
        return
      }

      // POST /api/pedidos
      if (req.method === 'POST' && route === '/api/pedidos') {
        bodyJSON().then(body => {
          const { cliente = '', notas = '', items = [], fecha_pedido } = body
          if (!Array.isArray(items) || items.length === 0)
            return json({ ok: false, error: 'items requerido' }, 400)
          const fecha  = fecha_pedido || new Date().toISOString()
          const result = db.prepare(`
            INSERT INTO pedidos (fecha_pedido, cliente, notas, items, estado, synced_at)
            VALUES (?, ?, ?, ?, 'pendiente', datetime('now','localtime'))
          `).run(fecha, cliente, notas, JSON.stringify(items))
          BrowserWindow.getAllWindows().forEach(w =>
            w.webContents.send('pedido-nuevo', { id: result.lastInsertRowid })
          )
          return json({ ok: true, id: result.lastInsertRowid })
        }).catch(e => json({ ok: false, error: e.message }, 400))
        return
      }

      // GET /api/pedidos
      if (req.method === 'GET' && route === '/api/pedidos') {
        const estado = url.searchParams.get('estado') || null
        const query  = estado
          ? db.prepare('SELECT * FROM pedidos WHERE estado=? ORDER BY fecha_pedido DESC')
          : db.prepare('SELECT * FROM pedidos ORDER BY fecha_pedido DESC')
        const rows   = estado ? query.all(estado) : query.all()
        return json(rows.map(r => ({ ...r, items: JSON.parse(r.items || '[]') })))
      }

      // PUT /api/pedidos/:id/estado
      if (req.method === 'PUT' && route.match(/^\/api\/pedidos\/\d+\/estado$/)) {
        const id = parseInt(route.split('/')[3])
        bodyJSON().then(body => {
          const { estado } = body
          const valid = ['pendiente','visto','en_revision','en_produccion','listo','entregado']
          if (!valid.includes(estado)) return json({ ok: false, error: 'estado inválido' }, 400)
          db.prepare('UPDATE pedidos SET estado=? WHERE id=?').run(estado, id)
          return json({ ok: true })
        }).catch(e => json({ ok: false, error: e.message }, 400))
        return
      }

      notFound()
    } catch (e) {
      console.error('[API]', e.message)
      json({ error: e.message }, 500)
    }
  })

  apiServer.listen(port, '0.0.0.0', () => {
    console.log(`[API] Servidor móvil escuchando en 0.0.0.0:${port}`)
  })
  apiServer.on('error', e => console.error('[API] Error:', e.message))
}

// ╔══════════════════════════════════════════════════════════════════════════╗
//  IPC — PEDIDOS
// ╚══════════════════════════════════════════════════════════════════════════╝
// ════════════════════════════════════════════════════════════
//  IPC — VENTAS
// ════════════════════════════════════════════════════════════
ipcMain.handle('get-ventas', (_, filtro) => {
  const { periodo } = filtro || {}
  const now = new Date()
  let query
  if (periodo === 'dia') {
    const hoy = now.toISOString().split('T')[0]
    query = db.prepare(`SELECT * FROM ventas WHERE date(fecha) = '${hoy}' ORDER BY fecha DESC`)
  } else if (periodo === 'semana') {
    query = db.prepare(`SELECT * FROM ventas WHERE fecha >= datetime('now', '-7 days') ORDER BY fecha DESC`)
  } else if (periodo === 'mes') {
    const mes = now.toISOString().slice(0, 7)
    query = db.prepare(`SELECT * FROM ventas WHERE strftime('%Y-%m', fecha) = '${mes}' ORDER BY fecha DESC`)
  } else if (periodo === 'anio') {
    const anio = now.getFullYear().toString()
    query = db.prepare(`SELECT * FROM ventas WHERE strftime('%Y', fecha) = '${anio}' ORDER BY fecha DESC`)
  } else {
    query = db.prepare(`SELECT * FROM ventas ORDER BY fecha DESC LIMIT 200`)
  }
  return query.all().map(r => ({ ...r, items: JSON.parse(r.items || '[]') }))
})

ipcMain.handle('save-venta', (_, venta) => {
  const items  = venta.items || []
  const total  = items.reduce((s, i) => s + ((i.precio_par || 0) * (i.pares || 0)), 0)
  const result = db.prepare(`
    INSERT INTO ventas (fecha, cliente, items, total, notas, origen)
    VALUES (?, ?, ?, ?, ?, ?)
  `).run(new Date().toISOString(), venta.cliente||'', JSON.stringify(items), total, venta.notas||'', venta.origen||'pc')
  // Descontar inventario por SKU
  const descontar = db.transaction((items) => {
    for (const item of items) {
      if (!item.sku || !item.pares) continue
      const inv = db.prepare('SELECT pares FROM inventario WHERE sku=?').get(item.sku)
      if (inv) {
        db.prepare(`UPDATE inventario SET pares=MAX(0,pares-?), updated_at=datetime('now','localtime') WHERE sku=?`)
          .run(item.pares, item.sku)
      }
    }
  })
  descontar(items)
  return { ok: true, id: result.lastInsertRowid, total }
})

ipcMain.handle('delete-venta', (_, id) => {
  const venta = db.prepare('SELECT items FROM ventas WHERE id=?').get(id)
  if (venta) {
    const items = JSON.parse(venta.items || '[]')
    const restaurar = db.transaction((items) => {
      for (const item of items) {
        if (!item.sku || !item.pares) continue
        db.prepare(`UPDATE inventario SET pares=pares+?, updated_at=datetime('now','localtime') WHERE sku=?`)
          .run(item.pares, item.sku)
      }
    })
    restaurar(items)
  }
  db.prepare('DELETE FROM ventas WHERE id=?').run(id)
  return { ok: true }
})

ipcMain.handle('get-precio-venta', () => {
  const cfg = loadConfig(); return cfg.precioVentaPar || 0
})
ipcMain.handle('save-precio-venta', (_, precio) => {
  const cfg = loadConfig(); cfg.precioVentaPar = parseFloat(precio) || 0; saveConfig(cfg); return true
})

// ════════════════════════════════════════════════════════════
ipcMain.handle('get-pedidos', (_, filtro) => {
  const estado = filtro?.estado || null
  const query  = estado
    ? db.prepare('SELECT * FROM pedidos WHERE estado=? ORDER BY fecha_pedido DESC')
    : db.prepare('SELECT * FROM pedidos ORDER BY fecha_pedido DESC')
  const rows = estado ? query.all(estado) : query.all()
  return rows.map(r => ({ ...r, items: JSON.parse(r.items || '[]') }))
})

ipcMain.handle('set-estado-pedido', (_, id, estado) => {
  const valid = ['pendiente','visto','en_revision','en_produccion','listo','entregado']
  if (!valid.includes(estado)) return { ok: false, error: 'estado inválido' }
  db.prepare('UPDATE pedidos SET estado=? WHERE id=?').run(estado, id)
  return { ok: true }
})

ipcMain.handle('delete-pedido', (_, id) => {
  db.prepare('DELETE FROM pedidos WHERE id=?').run(id)
  return { ok: true }
})

ipcMain.handle('get-api-port', () => {
  const cfg = loadConfig()
  return cfg.apiPort || 4000
})

ipcMain.handle('get-local-ip', () => {
  const { networkInterfaces } = require('os')
  const nets = networkInterfaces()
  for (const name of Object.keys(nets))
    for (const net of nets[name])
      if (net.family === 'IPv4' && !net.internal) return net.address
  return '127.0.0.1'
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  DIÁLOGOS
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('select-folder',    async (_, def) => {
  const r = await dialog.showOpenDialog({ properties: ['openDirectory'], defaultPath: def||'' })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('select-excel',     async (_, def) => {
  const r = await dialog.showOpenDialog({ filters:[{name:'Excel',extensions:['xlsx']}], defaultPath:def||'' })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('select-gcode-dir', async (_, def) => {
  const r = await dialog.showOpenDialog({ properties: ['openDirectory'], defaultPath: def||'' })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('select-gcode',     async (_, def) => {
  const r = await dialog.showOpenDialog({
    filters: [{name:'G-code', extensions:['gcode','g','gc']}], defaultPath: def||''
  })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('open-excel', async (_, p) => {
  if (p && fs.existsSync(p)) { await shell.openPath(p); return true }
  return false
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  CATÁLOGO — fotos y STL
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('select-foto', async (_, def) => {
  const r = await dialog.showOpenDialog({
    filters: [{name:'Imágenes', extensions:['jpg','jpeg','png','webp']}], defaultPath: def||''
  })
  return r.canceled ? null : r.filePaths[0]
})

ipcMain.handle('save-foto', async (_, rutaAretes, codigo, srcPath) => {
  try {
    const fotosDir = path.join(rutaAretes, 'fotos')
    if (!fs.existsSync(fotosDir)) fs.mkdirSync(fotosDir)
    const ext      = path.extname(srcPath)
    const destPath = path.join(fotosDir, `${codigo}${ext}`)
    fs.copyFileSync(srcPath, destPath)
    return destPath
  } catch { return null }
})

// Guardar preview STL generada desde canvas (dataUrl base64)
ipcMain.handle('save-preview-stl', async (_, rutaAretes, codigo, dataUrl) => {
  try {
    // Solo guardar si no existe foto real
    const fotosDir = path.join(rutaAretes, 'fotos')
    for (const ext of ['.jpg','.jpeg','.png','.webp']) {
      if (fs.existsSync(path.join(fotosDir, `${codigo}${ext}`))) return false // ya tiene foto
    }
    if (!fs.existsSync(fotosDir)) fs.mkdirSync(fotosDir)
    // dataUrl = "data:image/png;base64,..."
    const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '')
    const buffer = Buffer.from(base64, 'base64')
    const dest   = path.join(fotosDir, `${codigo}.png`)
    fs.writeFileSync(dest, buffer)
    return dest
  } catch { return null }
})

ipcMain.handle('get-foto', (_, rutaAretes, codigo) => {
  const fotosDir = path.join(rutaAretes, 'fotos')
  for (const ext of ['.jpg','.jpeg','.png','.webp']) {
    const p = path.join(fotosDir, `${codigo}${ext}`)
    if (fs.existsSync(p)) return p
  }
  return null
})

ipcMain.handle('get-stl-base64', (_, rutaAretes, carpeta, archivo) => {
  try {
    const stlPath = path.join(rutaAretes, carpeta, archivo)
    if (!fs.existsSync(stlPath)) return null
    return fs.readFileSync(stlPath).toString('base64')
  } catch { return null }
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  EXCEL helpers
// ╚══════════════════════════════════════════════════════════════════════════╝
function readWB (p)    { if (!p || !fs.existsSync(p)) return null; try { return XLSX.readFile(p) } catch { return null } }
function saveWB (wb,p) { try { XLSX.writeFile(wb,p); return true } catch { return false } }
function toRows (ws)   { return XLSX.utils.sheet_to_json(ws, { header:1, defval:'' }) }
function toSheet (r)   { return XLSX.utils.aoa_to_sheet(r) }

ipcMain.handle('get-catalogo-completo', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  const rows = toRows(wb.Sheets[wb.SheetNames[0]])
  return rows.slice(1)
    .filter(r => String(r[1]||'').trim() && String(r[1]).trim() !== 'TOTAL DE PARES')
    .map(r => ({ carpeta: String(r[0]||'').trim(), codigo: String(r[1]||'').trim(), archivo1: String(r[2]||'').trim() }))
})

ipcMain.handle('get-catalogo-codigos', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  const rows = toRows(wb.Sheets[wb.SheetNames[0]])
  const codigos = new Set()
  rows.slice(1).forEach(r => {
    const c = String(r[1]||'').trim()
    if (c && c !== 'TOTAL DE PARES') codigos.add(c)
  })
  return [...codigos].sort()
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  STL helpers
// ╚══════════════════════════════════════════════════════════════════════════╝
function getSTLStats (p) {
  try {
    const buf = fs.readFileSync(p)
    if (buf.length < 84) return null
    return { tris: buf.readUInt32LE(80), bytes: buf.length }
  } catch { return null }
}
function compareSTL (a, b, tol = 0.05) {
  if (!a || !b) return false
  if (a.bytes === b.bytes) return true
  const mx = Math.max(a.tris, b.tris)
  return mx > 0 && Math.abs(a.tris - b.tris) / mx <= tol
}

// ╔══════════════════════════════════════════════════════════════════════════╗
//  SINCRONIZAR STLs
// ╚══════════════════════════════════════════════════════════════════════════╝
function readExcelCodes (filePath) {
  const codes = new Set()
  const wb = readWB(filePath)
  if (!wb) return codes
  const rows = toRows(wb.Sheets[wb.SheetNames[0]])
  rows.slice(1).forEach(r => { const c = String(r[1]||'').trim(); if (c) codes.add(c) })
  return codes
}

function appendModelToExcel (filePath, carpeta, codigo, f1, f2) {
  try {
    const wb   = readWB(filePath)
    if (!wb) return false
    const ws   = wb.Sheets[wb.SheetNames[0]]
    const rows = toRows(ws)
    let insertAt = rows.length
    for (let i = 1; i < rows.length; i++) {
      const a = String(rows[i][0]||'').trim()
      if (a === 'TOTAL DE PARES') { insertAt = i; break }
      if (a === carpeta) insertAt = i + 1
    }
    rows.splice(insertAt, 0, [carpeta, codigo, f1, f2])
    wb.Sheets[wb.SheetNames[0]] = toSheet(rows)
    return saveWB(wb, filePath)
  } catch { return false }
}

ipcMain.handle('sync', async (event, { rutaBase, rutaExcel }) => {
  const send  = (type, data) => event.sender.send('sync-log', { type, data })
  const stats = { renombrados:0, agregados:0, advertencias:0, errores:0, _modelosAgregados:[] }
  if (!fs.existsSync(rutaBase)) { send('error','Carpeta de aretes no encontrada.'); return stats }
  send('info','Leyendo catálogo Excel...')
  const codigos = readExcelCodes(rutaExcel)
  send('ok', `Modelos en catálogo: ${codigos.size}`)
  send('divider','')
  const dirs = fs.readdirSync(rutaBase,{withFileTypes:true}).filter(d=>d.isDirectory()).map(d=>d.name)
  for (const carpeta of dirs) {
    const cp    = path.join(rutaBase, carpeta)
    const files = fs.readdirSync(cp).filter(f=>f.endsWith('.stl'))
    const o1    = files.filter(f=>f.startsWith('obj_1_'))
    const o3    = files.filter(f=>f.startsWith('obj_3_'))
    if (!o1.length && !o3.length) continue
    const nc   = carpeta.replace(/[^a-zA-Z0-9]/g,'')
    const pref = nc.substring(0, Math.min(3,nc.length)).toUpperCase()
    send('folder',`${carpeta}  →  ${pref}`)
    const ex  = files.filter(f=>new RegExp(`^${pref}\\d+_1\\.stl$`).test(f))
    let maxN  = 0
    for (const f of ex) { const m=f.match(new RegExp(`^${pref}(\\d+)_1\\.stl$`)); if (m&&+m[1]>maxN) maxN=+m[1] }
    let cnt = maxN + 1
    const processPair = (f1n, f2n, extra) => {
      const nd  = f1n.replace(/^obj_[13]_/,'')
      const f1p = path.join(cp, f1n)
      const f2p = path.join(cp, f2n)
      send('item', `${nd}${extra?' (par extra)':''}`)
      if (!fs.existsSync(f2p))       { send('warn',`Falta ${f2n}`);          stats.advertencias++; return }
      const s1=getSTLStats(f1p), s2=getSTLStats(f2p)
      if (!compareSTL(s1,s2))        { send('warn','Geometría diferente');    stats.advertencias++; return }
      const nn1=`${pref}${cnt}_1.stl`, nn2=`${pref}${cnt}_2.stl`, cod=`${pref}${cnt}`
      try {
        fs.renameSync(f1p, path.join(cp,nn1))
        fs.renameSync(f2p, path.join(cp,nn2))
        send('ok',`Renombrado: ${nn1} / ${nn2}`)
        stats.renombrados += 2
      } catch(e) { send('error',`Error al renombrar: ${e.message}`); stats.errores++; return }
      if (codigos.has(cod)) { send('muted','Ya existe en catálogo') }
      else if (fs.existsSync(rutaExcel)) {
        const ok = appendModelToExcel(rutaExcel, carpeta, cod, nn1, nn2)
        send(ok?'ok':'error', ok?'Agregado al catálogo':'Error al escribir Excel')
        if (ok) {
          codigos.add(cod)
          stats.agregados++
          stats._modelosAgregados.push({ codigo: cod, carpeta, archivo1: nn1 })
        } else stats.errores++
      }
      cnt++
    }
    for (const f of o1) processPair(f, `obj_2_${f.replace('obj_1_','')}`, false)
    for (const f of o3) processPair(f, `obj_4_${f.replace('obj_3_','')}`, true)
  }
  send('divider','')
  send('ok', stats.renombrados===0&&stats.advertencias===0&&stats.errores===0
    ? 'Todo al día — no había modelos nuevos.' : 'Sincronización completada.')
  stats.modelosAgregados = stats._modelosAgregados || []
  return stats
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  COMPARAR DB vs CARPETAS
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('compare-db-carpetas', (_, { rutaBase, rutaExcel }) => {
  const result = { soloCarpeta:[], soloDB:[], enAmbos:[], errores:[] }

  const catalogoSet = new Map()
  try {
    const wb = readWB(rutaExcel)
    if (wb) {
      const rows = toRows(wb.Sheets[wb.SheetNames[0]])
      rows.slice(1).forEach(r => {
        const codigo   = String(r[1]||'').trim()
        const carpeta  = String(r[0]||'').trim()
        const archivo1 = String(r[2]||'').trim()
        if (codigo && codigo !== 'TOTAL DE PARES')
          catalogoSet.set(codigo, { carpeta, archivo1 })
      })
    }
  } catch (e) { result.errores.push(`Error leyendo Excel: ${e.message}`) }

  const carpetaSet = new Map()
  try {
    if (fs.existsSync(rutaBase)) {
      const dirs = fs.readdirSync(rutaBase, { withFileTypes: true })
        .filter(d => d.isDirectory()).map(d => d.name)
      for (const carpeta of dirs) {
        const cp    = path.join(rutaBase, carpeta)
        const files = fs.readdirSync(cp).filter(f => f.endsWith('.stl'))
        const renombrados = files.filter(f => /^[A-Z]{2,5}\d+_1\.stl$/i.test(f))
        for (const f of renombrados) {
          const codigo = f.replace(/_1\.stl$/i, '')
          carpetaSet.set(codigo, { carpeta, archivo1: f, rutaCompleta: path.join(cp, f) })
        }
      }
    }
  } catch (e) { result.errores.push(`Error escaneando carpetas: ${e.message}`) }

  const todosCodigos = new Set([...catalogoSet.keys(), ...carpetaSet.keys()])
  for (const codigo of todosCodigos) {
    const enCatalogo = catalogoSet.has(codigo)
    const enCarpeta  = carpetaSet.has(codigo)
    if (enCatalogo && enCarpeta) {
      const info = carpetaSet.get(codigo)
      result.enAmbos.push({ codigo, carpeta: info.carpeta, archivo1: info.archivo1,
        archivoExiste: fs.existsSync(info.rutaCompleta) })
    } else if (enCarpeta && !enCatalogo) {
      const info = carpetaSet.get(codigo)
      result.soloCarpeta.push({ codigo, carpeta: info.carpeta, archivo1: info.archivo1 })
    } else if (enCatalogo && !enCarpeta) {
      const info    = catalogoSet.get(codigo)
      const rutaStl = path.join(rutaBase, info.carpeta, info.archivo1)
      result.soloDB.push({ codigo, carpeta: info.carpeta, archivo1: info.archivo1,
        archivoExiste: fs.existsSync(rutaStl) })
    }
  }
  return result
})

ipcMain.handle('agregar-a-catalogo', (_, { rutaExcel, modelos }) => {
  let agregados = 0
  for (const m of modelos) {
    const ok = appendModelToExcel(rutaExcel, m.carpeta, m.codigo, m.archivo1, m.archivo1.replace('_1.', '_2.'))
    if (ok) agregados++
  }
  return { ok: true, agregados }
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  ELIMINAR MODELO
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('eliminar-modelo', (_, { rutaAretes, rutaExcel, codigo }) => {
  const result = { ok:false, fotoEliminada:false, excelActualizado:false, dbActualizado:false, error:'' }
  try {
    // 1. Eliminar foto
    const fotosDir = path.join(rutaAretes, 'fotos')
    for (const ext of ['.jpg','.jpeg','.png','.webp']) {
      const fotoPath = path.join(fotosDir, `${codigo}${ext}`)
      if (fs.existsSync(fotoPath)) { fs.unlinkSync(fotoPath); result.fotoEliminada = true }
    }
    // 2. Eliminar del catálogo Excel
    const wb = readWB(rutaExcel)
    if (wb) {
      const rows = toRows(wb.Sheets[wb.SheetNames[0]])
        .filter((r, i) => i === 0 || String(r[1]||'').trim() !== codigo)
      wb.Sheets[wb.SheetNames[0]] = toSheet(rows)
      result.excelActualizado = saveWB(wb, rutaExcel)
    }
    // 3. Eliminar de modelos_nuevos
    db.prepare('DELETE FROM modelos_nuevos WHERE codigo=?').run(codigo)
    result.dbActualizado = true
    result.ok = true
  } catch (e) { result.error = e.message }
  return result
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  MODELOS NUEVOS
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('get-modelos-nuevos', () => {
  return db.prepare('SELECT codigo, carpeta, archivo1, added_at FROM modelos_nuevos ORDER BY added_at DESC').all()
})

ipcMain.handle('marcar-modelo-visto', (_, codigo) => {
  db.prepare('DELETE FROM modelos_nuevos WHERE codigo=?').run(codigo)
  return true
})

ipcMain.handle('marcar-todos-vistos', () => {
  db.prepare('DELETE FROM modelos_nuevos').run()
  return true
})

ipcMain.handle('registrar-modelos-nuevos', (_, modelos) => {
  const ins  = db.prepare('INSERT OR IGNORE INTO modelos_nuevos (codigo,carpeta,archivo1) VALUES (?,?,?)')
  const many = db.transaction((items) => { for (const m of items) ins.run(m.codigo, m.carpeta||'', m.archivo1||'') })
  many(modelos)
  return true
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  COLORES — lee de la tabla filamentos con código corto generado
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('get-colores', () => {
  const rows = db.prepare(`
    SELECT id, nombre, color_hex as hex, notas as descripcion,
           stock_gr as stockGr, costo_kg as costoPorKg
    FROM filamentos ORDER BY nombre
  `).all()

  const IGNORAR = /\b(PLA|PETG|ABS|TPU|ASA|NYLON|SILK|WOOD|METAL|SUNLU|ESUN|BAMBU|CREALITY|HATCHBOX|POLYMAKER|PRUSAMENT|BASICFIL|MEXICOMAKERS|MATTE|PLUS|PRO|MAX|LITE|BASIC)\b/gi
  const usedCodes = new Set()

  return rows.map(r => {
    const colorPuro = r.nombre.replace(IGNORAR, '').trim().split(/\s+/)
      .find(p => p.replace(/[^a-zA-Z]/g,'').length >= 2) || 'COL'
    const base   = colorPuro.replace(/[^a-zA-Z]/g,'').substring(0, 3).toUpperCase()
    let n = 1
    while (usedCodes.has(`${base}${n}`)) n++
    const codigo = `${base}${n}`
    usedCodes.add(codigo)
    return { ...r, codigo }
  })
})

ipcMain.handle('save-color', (_, __, color) => {
  const nombre = color.nombre || ''
  if (!nombre) return null
  let codigo = color.codigo
  if (!codigo) {
    const base     = nombre.replace(/[^a-zA-Z]/g,'').substring(0,3).toUpperCase()
    const existing = db.prepare("SELECT nombre FROM filamentos WHERE nombre LIKE ? || '%'").all(base)
    codigo = `${base}${existing.length + 1}`
  }
  db.prepare(`
    INSERT INTO filamentos (nombre, costo_kg, stock_gr, color_hex, notas)
    VALUES (?,?,?,?,?)
    ON CONFLICT(nombre) DO UPDATE SET
      costo_kg=excluded.costo_kg, stock_gr=excluded.stock_gr,
      color_hex=excluded.color_hex, notas=excluded.notas,
      updated_at=datetime('now','localtime')
  `).run(nombre, color.costoPorKg||0, color.stockGr||0, color.hex||'#888888', color.descripcion||'')
  return codigo
})

ipcMain.handle('delete-color', (_, __, nombre) => {
  db.prepare('DELETE FROM filamentos WHERE nombre=?').run(nombre)
  return true
})

ipcMain.handle('refresh-stock', () => {
  return db.prepare('SELECT nombre, stock_gr as stockGr, updated_at FROM filamentos ORDER BY nombre').all()
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  INVENTARIO
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('get-inventario', () => {
  return db.prepare(`
    SELECT id as _idx, sku, modelo, color, codigo_color as codigoColor,
           pares, costo_produccion as costoProduccion
    FROM inventario ORDER BY sku
  `).all()
})

ipcMain.handle('save-inventario', (_, __, item) => {
  db.prepare(`
    INSERT INTO inventario (sku, modelo, color, codigo_color, pares, costo_produccion)
    VALUES (?,?,?,?,?,?)
    ON CONFLICT(sku) DO UPDATE SET
      modelo=excluded.modelo, color=excluded.color,
      codigo_color=excluded.codigo_color, pares=excluded.pares,
      costo_produccion=excluded.costo_produccion,
      updated_at=datetime('now','localtime')
  `).run(item.sku, item.modelo||'', item.color||'', item.codigoColor||'',
         item.pares||0, item.costoProduccion||0)
  return true
})

ipcMain.handle('descontar-stock', (_, __, sku, qty) => {
  const row = db.prepare('SELECT pares FROM inventario WHERE sku=?').get(sku)
  if (!row) return { ok:false, msg:`SKU ${sku} no encontrado` }
  if (row.pares < qty) return { ok:false, msg:`Stock insuficiente (${row.pares} pares disponibles)` }
  db.prepare(`UPDATE inventario SET pares=pares-?, updated_at=datetime('now','localtime') WHERE sku=?`).run(qty, sku)
  return { ok:true, msg:`Stock actualizado: ${row.pares - qty} pares` }
})

ipcMain.handle('delete-inventario', (_, __, sku) => {
  db.prepare('DELETE FROM inventario WHERE sku=?').run(sku)
  return true
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  CLIENTES
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('get-clientes', () => {
  const rows   = db.prepare('SELECT * FROM clientes ORDER BY tipo, nombre').all()
  const result = { directo:[], distribuidor:[], mayoreo:[] }
  rows.forEach((r, i) => {
    const c = {
      _idx: i, nombre: r.nombre, contacto: r.contacto, direccion: r.direccion,
      rfc: r.rfc, notas: r.notas, fechaRegistro: r.fecha_registro,
      fechaUltimaCompra: r.fecha_ultima_compra, piezasUltimaCompra: r.piezas_ultima_compra,
      acumuladoHistorico: r.acumulado_historico,
      skus: r.skus ? String(r.skus).split(',').map(s=>s.trim()) : [], _id: r.id
    }
    if (result[r.tipo]) result[r.tipo].push(c)
  })
  return result
})

ipcMain.handle('save-cliente', (_, __, cliente) => {
  const skusStr = Array.isArray(cliente.skus) ? cliente.skus.join(', ') : (cliente.skus||'')
  if (cliente._id) {
    db.prepare(`
      UPDATE clientes SET tipo=?,nombre=?,contacto=?,direccion=?,rfc=?,notas=?,
        fecha_registro=?,fecha_ultima_compra=?,piezas_ultima_compra=?,acumulado_historico=?,skus=?
      WHERE id=?
    `).run(cliente.tipo, cliente.nombre, cliente.contacto||'', cliente.direccion||'',
           cliente.rfc||'', cliente.notas||'', cliente.fechaRegistro||'',
           cliente.fechaUltimaCompra||'', cliente.piezasUltimaCompra||0,
           cliente.acumuladoHistorico||0, skusStr, cliente._id)
  } else {
    db.prepare(`
      INSERT INTO clientes
        (tipo,nombre,contacto,direccion,rfc,notas,fecha_registro,fecha_ultima_compra,piezas_ultima_compra,acumulado_historico,skus)
      VALUES (?,?,?,?,?,?,?,?,?,?,?)
    `).run(cliente.tipo, cliente.nombre, cliente.contacto||'', cliente.direccion||'',
           cliente.rfc||'', cliente.notas||'', cliente.fechaRegistro||'',
           cliente.fechaUltimaCompra||'', cliente.piezasUltimaCompra||0,
           cliente.acumuladoHistorico||0, skusStr)
  }
  return true
})

ipcMain.handle('delete-cliente', (_, __, rowIndex, tipo) => {
  const row = db.prepare('SELECT id FROM clientes WHERE tipo=? ORDER BY nombre LIMIT 1 OFFSET ?').get(tipo, rowIndex)
  if (row) db.prepare('DELETE FROM clientes WHERE id=?').run(row.id)
  return true
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  G-CODE — parse y cálculo de costos
// ╚══════════════════════════════════════════════════════════════════════════╝
function parseGcodeFile (filePath) {
  try {
    const lines = fs.readFileSync(filePath, 'utf8').split('\n').slice(0,200)
    const find  = (patterns) => {
      for (const line of lines)
        for (const pat of patterns) { const m = line.match(pat); if (m) return m[1].trim() }
      return null
    }
    const pesoStr   = find([/;\s*filament\s+used\s*\[g\]\s*=\s*([\d.]+)/i, /;\s*total\s+filament\s+used\s*=\s*([\d.]+)\s*g/i])
    const tiempoStr = find([/;\s*estimated\s+printing\s+time\s*=\s*(.+)/i, /;\s*total\s+estimated\s+time\s*:\s*(.+)/i])
    let horas = 0
    if (tiempoStr) {
      const h = tiempoStr.match(/(\d+)h/), m = tiempoStr.match(/(\d+)m/), s = tiempoStr.match(/(\d+)s/)
      horas = (h?+h[1]:0) + (m?+m[1]/60:0) + (s?+s[1]/3600:0)
    }
    const filamento = find([/;\s*filament_type\s*=\s*(.+)/i, /;\s*filament\s+type\s*=\s*(.+)/i])
    return { pesoGr:pesoStr?parseFloat(pesoStr):null, horasImp:horas||null, filamento:filamento||null, tiempoRaw:tiempoStr||null }
  } catch (e) { return { error: e.message } }
}

ipcMain.handle('parse-gcode', (_, filePath) => parseGcodeFile(filePath))

ipcMain.handle('calc-costo', (_, p) => {
  const cf  = (p.pesoGr/1000)*p.costoPorKg
  const ce  = (p.horasImp*p.wattsPrinter/1000)*p.costoKwh
  const sub = cf + ce + p.costoHerrajes + p.costoEmpaque
  const des = sub*(p.porcDesperdicio/100)
  return {
    costoFilamento:    +cf.toFixed(2),
    costoElectricidad: +ce.toFixed(2),
    costoHerrajes:     +p.costoHerrajes.toFixed(2),
    costoEmpaque:      +p.costoEmpaque.toFixed(2),
    desperdicio:       +des.toFixed(2),
    total:             +(sub+des).toFixed(2)
  }
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  COSTOS
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('get-costos', () => {
  return db.prepare(`
    SELECT id as _idx, sku, modelo, color,
           peso_gr as pesoGr, horas_imp as horasImp,
           costo_filamento as costoFilamento, costo_elec as costoElec,
           costo_herrajes as costoHerrajes, costo_empaque as costoEmpaque,
           desperdicio, costo_total as costoTotal, fecha
    FROM costos ORDER BY rowid DESC
  `).all()
})

ipcMain.handle('save-costo', (_, __, item) => {
  db.prepare(`
    INSERT INTO costos
      (sku,modelo,color,peso_gr,horas_imp,costo_filamento,costo_elec,costo_herrajes,costo_empaque,desperdicio,costo_total,fecha)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
  `).run(item.sku, item.modelo, item.color, item.pesoGr, item.horasImp,
         item.costoFilamento, item.costoElec, item.costoHerrajes,
         item.costoEmpaque, item.desperdicio, item.costoTotal,
         new Date().toLocaleDateString('es-MX'))
  return true
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  MOONRAKER
// ╚══════════════════════════════════════════════════════════════════════════╝
function moonrakerRequest (ip, method, endpoint, body) {
  return new Promise((resolve, reject) => {
    const data    = body ? JSON.stringify(body) : null
    const options = {
      hostname: ip.split(':')[0],
      port:     parseInt(ip.split(':')[1]) || 7125,
      path:     endpoint, method,
      headers:  { 'Content-Type':'application/json', 'Content-Length': data?Buffer.byteLength(data):0 },
      timeout:  8000
    }
    const req = http.request(options, res => {
      let raw = ''
      res.on('data', d => raw += d)
      res.on('end', () => { try { resolve(JSON.parse(raw)) } catch { resolve({raw}) } })
    })
    req.on('error',   e => reject(e))
    req.on('timeout', () => { req.destroy(); reject(new Error('Timeout')) })
    if (data) req.write(data)
    req.end()
  })
}

ipcMain.handle('moonraker-status', async (_, ip) => {
  try {
    const r   = await moonrakerRequest(ip,'GET','/printer/objects/query?print_stats&heater_bed&extruder',null)
    const ps  = r?.result?.status?.print_stats||{}
    const bed = r?.result?.status?.heater_bed ||{}
    const ext = r?.result?.status?.extruder   ||{}
    return { ok:true, state:ps.state||'unknown', filename:ps.filename||'',
             progress:ps.print_duration||0, bedTemp:bed.temperature||0, extTemp:ext.temperature||0 }
  } catch(e) { return { ok:false, state:'offline', error:e.message } }
})

ipcMain.handle('moonraker-gcode', async (_, ip, script) => {
  try {
    const r = await moonrakerRequest(ip,'POST','/printer/gcode/script',{script})
    return { ok:true, result:r }
  } catch(e) { return { ok:false, error:e.message } }
})

ipcMain.handle('moonraker-print-status', async (_, ip) => {
  try {
    const r  = await moonrakerRequest(ip,'GET','/printer/objects/query?print_stats',null)
    const ps = r?.result?.status?.print_stats||{}
    return { ok:true, state:ps.state||'unknown', filename:ps.filename||'',
             duration:ps.print_duration||0, totalDur:ps.total_duration||0 }
  } catch(e) { return { ok:false, state:'offline', error:e.message } }
})

// ╔══════════════════════════════════════════════════════════════════════════╗
//  ORCASLICER
// ╚══════════════════════════════════════════════════════════════════════════╝
ipcMain.handle('open-orcaslicer', async (_, orcaPath, stlPaths) => {
  try {
    if (!fs.existsSync(orcaPath)) return { ok:false, error:`OrcaSlicer no encontrado en: ${orcaPath}` }
    const missing = stlPaths.filter(p => !fs.existsSync(p))
    if (missing.length) return { ok:false, error:`STLs no encontrados: ${missing.join(', ')}` }
    spawn(orcaPath, stlPaths, { detached:true, stdio:'ignore' }).unref()
    return { ok:true }
  } catch(e) { return { ok:false, error:e.message } }
})

ipcMain.handle('select-orca-exe', async (_, def) => {
  const r = await dialog.showOpenDialog({
    filters:[{name:'Ejecutable',extensions:['exe']}],
    defaultPath: def||'C:\\Program Files\\OrcaSlicer'
  })
  return r.canceled ? null : r.filePaths[0]
})
