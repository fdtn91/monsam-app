const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path    = require('path')
const fs      = require('fs')
const { spawn } = require('child_process')
const { getDB, migrateFromExcel } = require('./db')

// ════════════════════════════════════════════════════════════
//  CONFIG
// ════════════════════════════════════════════════════════════
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

// ════════════════════════════════════════════════════════════
//  DB — inicializar al arrancar
// ════════════════════════════════════════════════════════════
let db = null

function initDB () {
  const cfg    = loadConfig()
  const dbPath = cfg.rutaDB || path.join(__dirname, 'datos.db')
  db = getDB(dbPath)
  const migrated = migrateFromExcel(db, cfg)
  const total = Object.values(migrated).reduce((a, b) => a + b, 0)
  if (total > 0) console.log('Migración desde Excel completada:', migrated)
  return db
}

// ════════════════════════════════════════════════════════════
//  VENTANA
// ════════════════════════════════════════════════════════════
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
})
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })

ipcMain.on('win-minimize', e => BrowserWindow.fromWebContents(e.sender).minimize())
ipcMain.on('win-maximize', e => {
  const w = BrowserWindow.fromWebContents(e.sender)
  w.isMaximized() ? w.unmaximize() : w.maximize()
})
ipcMain.on('win-close', e => BrowserWindow.fromWebContents(e.sender).close())

// ════════════════════════════════════════════════════════════
//  DIÁLOGOS
// ════════════════════════════════════════════════════════════
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
ipcMain.handle('select-gcode', async (_, def) => {
  const r = await dialog.showOpenDialog({
    filters: [{name:'G-code', extensions:['gcode','g','gc']}], defaultPath: def||''
  })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('open-excel', async (_, p) => {
  if (p && fs.existsSync(p)) { await shell.openPath(p); return true }
  return false
})

// ════════════════════════════════════════════════════════════
//  CATÁLOGO — fotos y STL (sin cambios, siguen usando filesystem)
// ════════════════════════════════════════════════════════════
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

// ════════════════════════════════════════════════════════════
//  EXCEL helpers — solo para sync de catálogo STL
// ════════════════════════════════════════════════════════════
const XLSX = require('xlsx')
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

// ════════════════════════════════════════════════════════════
//  STL helpers
// ════════════════════════════════════════════════════════════
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

// ════════════════════════════════════════════════════════════
//  SINCRONIZAR STLs
// ════════════════════════════════════════════════════════════
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
  const stats = { renombrados:0, agregados:0, advertencias:0, errores:0 }
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
        if (ok) { codigos.add(cod); stats.agregados++ } else stats.errores++
      }
      cnt++
    }
    for (const f of o1) processPair(f, `obj_2_${f.replace('obj_1_','')}`, false)
    for (const f of o3) processPair(f, `obj_4_${f.replace('obj_3_','')}`, true)
  }
  send('divider','')
  send('ok', stats.renombrados===0&&stats.advertencias===0&&stats.errores===0
    ? 'Todo al día — no había modelos nuevos.' : 'Sincronización completada.')
  return stats
})

// ════════════════════════════════════════════════════════════
//  COLORES — ahora lee de la tabla filamentos (DB compartida)
//  monsam ve los filamentos de e500 como "colores"
// ════════════════════════════════════════════════════════════
ipcMain.handle('get-colores', () => {
  return db.prepare(`
    SELECT id, nombre, nombre as codigo, color_hex as hex, notas as descripcion,
           stock_gr as stockGr, costo_kg as costoPorKg
    FROM filamentos ORDER BY nombre
  `).all()
})

ipcMain.handle('save-color', (_, __, color) => {
  const nombre = color.nombre || ''
  if (!nombre) return null

  // Generar código si no viene
  let codigo = color.codigo
  if (!codigo) {
    const base = nombre.replace(/[^a-zA-Z]/g,'').substring(0,3).toUpperCase()
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

// ════════════════════════════════════════════════════════════
//  INVENTARIO
// ════════════════════════════════════════════════════════════
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
  db.prepare(`UPDATE inventario SET pares=pares-?, updated_at=datetime('now','localtime') WHERE sku=?`)
    .run(qty, sku)
  return { ok:true, msg:`Stock actualizado: ${row.pares - qty} pares` }
})

ipcMain.handle('delete-inventario', (_, __, sku) => {
  db.prepare('DELETE FROM inventario WHERE sku=?').run(sku)
  return true
})

// ════════════════════════════════════════════════════════════
//  CLIENTES
// ════════════════════════════════════════════════════════════
ipcMain.handle('get-clientes', () => {
  const rows = db.prepare('SELECT * FROM clientes ORDER BY tipo, nombre').all()
  const result = { directo:[], distribuidor:[], mayoreo:[] }
  rows.forEach((r, i) => {
    const c = {
      _idx:               i,
      nombre:             r.nombre,
      contacto:           r.contacto,
      direccion:          r.direccion,
      rfc:                r.rfc,
      notas:              r.notas,
      fechaRegistro:      r.fecha_registro,
      fechaUltimaCompra:  r.fecha_ultima_compra,
      piezasUltimaCompra: r.piezas_ultima_compra,
      acumuladoHistorico: r.acumulado_historico,
      skus:               r.skus ? String(r.skus).split(',').map(s=>s.trim()) : [],
      _id:                r.id
    }
    if (result[r.tipo]) result[r.tipo].push(c)
  })
  return result
})

ipcMain.handle('save-cliente', (_, __, cliente) => {
  const skusStr = Array.isArray(cliente.skus) ? cliente.skus.join(', ') : (cliente.skus||'')
  if (cliente._id) {
    db.prepare(`
      UPDATE clientes SET
        tipo=?, nombre=?, contacto=?, direccion=?, rfc=?, notas=?,
        fecha_registro=?, fecha_ultima_compra=?,
        piezas_ultima_compra=?, acumulado_historico=?, skus=?
      WHERE id=?
    `).run(cliente.tipo, cliente.nombre, cliente.contacto||'', cliente.direccion||'',
           cliente.rfc||'', cliente.notas||'', cliente.fechaRegistro||'',
           cliente.fechaUltimaCompra||'', cliente.piezasUltimaCompra||0,
           cliente.acumuladoHistorico||0, skusStr, cliente._id)
  } else {
    db.prepare(`
      INSERT INTO clientes
        (tipo, nombre, contacto, direccion, rfc, notas,
         fecha_registro, fecha_ultima_compra, piezas_ultima_compra, acumulado_historico, skus)
      VALUES (?,?,?,?,?,?,?,?,?,?,?)
    `).run(cliente.tipo, cliente.nombre, cliente.contacto||'', cliente.direccion||'',
           cliente.rfc||'', cliente.notas||'', cliente.fechaRegistro||'',
           cliente.fechaUltimaCompra||'', cliente.piezasUltimaCompra||0,
           cliente.acumuladoHistorico||0, skusStr)
  }
  return true
})

ipcMain.handle('delete-cliente', (_, __, rowIndex, tipo) => {
  const row = db.prepare('SELECT id FROM clientes WHERE tipo=? ORDER BY nombre LIMIT 1 OFFSET ?')
                .get(tipo, rowIndex)
  if (row) db.prepare('DELETE FROM clientes WHERE id=?').run(row.id)
  return true
})

// ════════════════════════════════════════════════════════════
//  G-CODE — parse y cálculo de costos
// ════════════════════════════════════════════════════════════
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
    return { pesoGr: pesoStr?parseFloat(pesoStr):null, horasImp:horas||null, filamento:filamento||null, tiempoRaw:tiempoStr||null }
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

// ════════════════════════════════════════════════════════════
//  COSTOS
// ════════════════════════════════════════════════════════════
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
      (sku, modelo, color, peso_gr, horas_imp,
       costo_filamento, costo_elec, costo_herrajes,
       costo_empaque, desperdicio, costo_total, fecha)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
  `).run(item.sku, item.modelo, item.color, item.pesoGr, item.horasImp,
         item.costoFilamento, item.costoElec, item.costoHerrajes,
         item.costoEmpaque, item.desperdicio, item.costoTotal,
         new Date().toLocaleDateString('es-MX'))
  return true
})

// ════════════════════════════════════════════════════════════
//  MOONRAKER
// ════════════════════════════════════════════════════════════
const http = require('http')

function moonrakerRequest (ip, method, endpoint, body) {
  return new Promise((resolve, reject) => {
    const data    = body ? JSON.stringify(body) : null
    const options = {
      hostname: ip.split(':')[0],
      port:     parseInt(ip.split(':')[1]) || 7125,
      path:     endpoint, method,
      headers: { 'Content-Type':'application/json', 'Content-Length': data?Buffer.byteLength(data):0 },
      timeout: 8000
    }
    const req = http.request(options, res => {
      let raw = ''
      res.on('data', d => raw += d)
      res.on('end', () => { try { resolve(JSON.parse(raw)) } catch { resolve({raw}) } })
    })
    req.on('error', e => reject(e))
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

// ════════════════════════════════════════════════════════════
//  ORCASLICER
// ════════════════════════════════════════════════════════════
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
