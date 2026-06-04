const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('path')
const fs   = require('fs')
const XLSX = require('xlsx')

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
//  VENTANA
// ════════════════════════════════════════════════════════════
function createWindow () {
  const win = new BrowserWindow({
    width: 1200, height: 760,
    minWidth: 960, minHeight: 640,
    frame: false,
    backgroundColor: '#080B10',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  })
  win.loadFile('index.html')
}

app.whenReady().then(createWindow)
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
ipcMain.handle('select-folder', async (_, def) => {
  const r = await dialog.showOpenDialog({ properties: ['openDirectory'], defaultPath: def || '' })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('select-excel', async (_, def) => {
  const r = await dialog.showOpenDialog({
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    defaultPath: def || ''
  })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('select-gcode-dir', async (_, def) => {
  const r = await dialog.showOpenDialog({ properties: ['openDirectory'], defaultPath: def || '' })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('select-gcode', async (_, def) => {
  const r = await dialog.showOpenDialog({
    filters: [{ name: 'G-code', extensions: ['gcode', 'g', 'gc'] }],
    defaultPath: def || ''
  })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('open-excel', async (_, p) => {
  if (fs.existsSync(p)) { await shell.openPath(p); return true }
  return false
})

// ════════════════════════════════════════════════════════════
//  CATÁLOGO — fotos y STL
// ════════════════════════════════════════════════════════════

// Seleccionar foto de imagen
ipcMain.handle('select-foto', async (_, def) => {
  const r = await dialog.showOpenDialog({
    filters: [{ name: 'Imágenes', extensions: ['jpg','jpeg','png','webp'] }],
    defaultPath: def || ''
  })
  return r.canceled ? null : r.filePaths[0]
})

// Guardar foto — copia el archivo a la carpeta fotos/ junto a los aretes
ipcMain.handle('save-foto', async (_, rutaAretes, codigo, srcPath) => {
  try {
    const fotosDir = path.join(rutaAretes, 'fotos')
    if (!fs.existsSync(fotosDir)) fs.mkdirSync(fotosDir)
    const ext    = path.extname(srcPath)
    const destPath = path.join(fotosDir, `${codigo}${ext}`)
    fs.copyFileSync(srcPath, destPath)
    return destPath
  } catch (e) { return null }
})

// Obtener ruta de foto si existe
ipcMain.handle('get-foto', (_, rutaAretes, codigo) => {
  const fotosDir = path.join(rutaAretes, 'fotos')
  for (const ext of ['.jpg','.jpeg','.png','.webp']) {
    const p = path.join(fotosDir, `${codigo}${ext}`)
    if (fs.existsSync(p)) return p
  }
  return null
})

// Leer archivo STL como base64 para Three.js
ipcMain.handle('get-stl-base64', (_, rutaAretes, carpeta, archivo) => {
  try {
    const stlPath = path.join(rutaAretes, carpeta, archivo)
    if (!fs.existsSync(stlPath)) return null
    const buf = fs.readFileSync(stlPath)
    return buf.toString('base64')
  } catch { return null }
})

// Leer catálogo completo con carpeta y archivo STL
ipcMain.handle('get-catalogo-completo', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  const ws   = wb.Sheets[wb.SheetNames[0]]
  const rows = toRows(ws)
  const result = []
  for (let i = 1; i < rows.length; i++) {
    const carpeta  = String(rows[i][0] || '').trim()
    const codigo   = String(rows[i][1] || '').trim()
    const archivo1 = String(rows[i][2] || '').trim()
    if (!codigo || codigo === 'TOTAL DE PARES') continue
    result.push({ carpeta, codigo, archivo1 })
  }
  return result
})

// ════════════════════════════════════════════════════════════
//  EXCEL — helpers generales
// ════════════════════════════════════════════════════════════
function readWB (filePath) {
  if (!fs.existsSync(filePath)) return null
  try { return XLSX.readFile(filePath) } catch { return null }
}
function saveWB (wb, filePath) {
  try { XLSX.writeFile(wb, filePath); return true } catch { return false }
}
function ensureSheet (wb, name, headers) {
  if (!wb.SheetNames.includes(name)) {
    const ws = XLSX.utils.aoa_to_sheet([headers])
    XLSX.utils.book_append_sheet(wb, ws, name)
  }
  return wb.Sheets[name]
}
function toRows (ws)   { return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) }
function toSheet (rows){ return XLSX.utils.aoa_to_sheet(rows) }

// ════════════════════════════════════════════════════════════
//  STL — helpers
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
//  SINCRONIZAR
// ════════════════════════════════════════════════════════════
const CAT_HEADERS = ['Carpeta', 'Codigo', 'Archivo_1', 'Archivo_2']

function readExcelCodes (filePath) {
  const codes = new Set()
  const wb = readWB(filePath)
  if (!wb) return codes
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = toRows(ws)
  for (let i = 1; i < rows.length; i++) {
    const c = String(rows[i][1] || '').trim()
    if (c) codes.add(c)
  }
  return codes
}

function appendModelToExcel (filePath, carpeta, codigo, f1, f2) {
  try {
    const wb = readWB(filePath)
    if (!wb) return false
    const ws   = wb.Sheets[wb.SheetNames[0]]
    const rows = toRows(ws)
    let insertAt = rows.length
    for (let i = 1; i < rows.length; i++) {
      const a = String(rows[i][0] || '').trim()
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
  const stats = { renombrados: 0, agregados: 0, advertencias: 0, errores: 0 }

  if (!fs.existsSync(rutaBase)) { send('error', 'Carpeta de aretes no encontrada.'); return stats }

  send('info', 'Leyendo catálogo Excel...')
  const codigos = readExcelCodes(rutaExcel)
  send('ok',   `Modelos en catálogo: ${codigos.size}`)
  send('divider', '')

  const dirs = fs.readdirSync(rutaBase, { withFileTypes: true })
    .filter(d => d.isDirectory()).map(d => d.name)

  for (const carpeta of dirs) {
    const cp    = path.join(rutaBase, carpeta)
    const files = fs.readdirSync(cp).filter(f => f.endsWith('.stl'))
    const o1    = files.filter(f => f.startsWith('obj_1_'))
    const o3    = files.filter(f => f.startsWith('obj_3_'))
    if (!o1.length && !o3.length) continue

    const nc   = carpeta.replace(/[^a-zA-Z0-9]/g, '')
    const pref = nc.substring(0, Math.min(3, nc.length)).toUpperCase()
    send('folder', `${carpeta}  →  ${pref}`)

    const ex   = files.filter(f => new RegExp(`^${pref}\\d+_1\\.stl$`).test(f))
    let maxN   = 0
    for (const f of ex) {
      const m = f.match(new RegExp(`^${pref}(\\d+)_1\\.stl$`))
      if (m && +m[1] > maxN) maxN = +m[1]
    }
    let cnt = maxN + 1

    const processPair = (f1n, f2n, extra) => {
      const nd  = f1n.replace(/^obj_[13]_/, '')
      const f1p = path.join(cp, f1n)
      const f2p = path.join(cp, f2n)
      send('item', `${nd}${extra ? ' (par extra)' : ''}`)

      if (!fs.existsSync(f2p))         { send('warn', `Falta ${f2n}`);              stats.advertencias++; return }
      const s1 = getSTLStats(f1p), s2 = getSTLStats(f2p)
      if (!compareSTL(s1, s2))         { send('warn', 'Geometría diferente');        stats.advertencias++; return }

      const nn1 = `${pref}${cnt}_1.stl`
      const nn2 = `${pref}${cnt}_2.stl`
      const cod = `${pref}${cnt}`

      try {
        fs.renameSync(f1p, path.join(cp, nn1))
        fs.renameSync(f2p, path.join(cp, nn2))
        send('ok', `Renombrado: ${nn1} / ${nn2}`)
        stats.renombrados += 2
      } catch (e) { send('error', `Error al renombrar: ${e.message}`); stats.errores++; return }

      if (codigos.has(cod)) {
        send('muted', 'Ya existe en catálogo')
      } else if (fs.existsSync(rutaExcel)) {
        const ok = appendModelToExcel(rutaExcel, carpeta, cod, nn1, nn2)
        send(ok ? 'ok' : 'error', ok ? 'Agregado al catálogo' : 'Error al escribir Excel')
        if (ok) { codigos.add(cod); stats.agregados++ } else stats.errores++
      }
      cnt++
    }

    for (const f of o1) processPair(f, `obj_2_${f.replace('obj_1_', '')}`, false)
    for (const f of o3) processPair(f, `obj_4_${f.replace('obj_3_', '')}`, true)
  }

  send('divider', '')
  send('ok', stats.renombrados === 0 && stats.advertencias === 0 && stats.errores === 0
    ? 'Todo al día — no había modelos nuevos.'
    : 'Sincronización completada.')
  return stats
})

// ════════════════════════════════════════════════════════════
//  COLORES
//  Hoja "Colores": Nombre | CodigoColor | Hex | Descripcion | StockGr | CostoPorKg
// ════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════
//  CATÁLOGO — leer códigos de modelos STL desde la hoja principal
// ════════════════════════════════════════════════════════════
ipcMain.handle('get-catalogo-codigos', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  const ws   = wb.Sheets[wb.SheetNames[0]]
  const rows = toRows(ws)
  const codigos = new Set()
  for (let i = 1; i < rows.length; i++) {
    const codigo = String(rows[i][1] || '').trim()
    if (codigo && codigo !== 'TOTAL DE PARES') codigos.add(codigo)
  }
  return [...codigos].sort()
})

const COL_HDR = ['Nombre', 'CodigoColor', 'Hex', 'Descripcion', 'StockGr', 'CostoPorKg']

ipcMain.handle('get-colores', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  ensureSheet(wb, 'Colores', COL_HDR)
  const rows = toRows(wb.Sheets['Colores'])
  return rows.slice(1).filter(r => r[0]).map(r => ({
    nombre:      r[0],
    codigo:      r[1],
    hex:         r[2],
    descripcion: r[3],
    stockGr:     r[4],
    costoPorKg:  r[5]
  }))
})

ipcMain.handle('save-color', (_, filePath, color) => {
  let wb = readWB(filePath)
  if (!wb) { wb = XLSX.utils.book_new() }
  ensureSheet(wb, 'Colores', COL_HDR)
  const ws   = wb.Sheets['Colores']
  const rows = toRows(ws)

  // Generar código automático usando el campo colorNombre si viene, si no el nombre completo
  let codigo = color.codigo
  if (!codigo) {
    const fuente = color.colorNombre || color.nombre
    const base = fuente.replace(/[^a-zA-Z]/g, '').substring(0, 3).toUpperCase()
    const existing = rows.slice(1).map(r => String(r[1] || ''))
    let n = 1
    while (existing.includes(`${base}${n}`)) n++
    codigo = `${base}${n}`
  }

  const row = [color.nombre, codigo, color.hex || '', color.descripcion || '',
               color.stockGr || 0, color.costoPorKg || 0]
  const idx = rows.findIndex((r, i) => i > 0 && r[0] === color.nombre)
  if (idx >= 0) rows[idx] = row
  else rows.push(row)

  wb.Sheets['Colores'] = toSheet(rows)
  const ok = saveWB(wb, filePath)
  return ok ? codigo : null   // devuelve el código generado
})

ipcMain.handle('delete-color', (_, filePath, nombre) => {
  const wb = readWB(filePath)
  if (!wb) return false
  ensureSheet(wb, 'Colores', COL_HDR)
  const rows = toRows(wb.Sheets['Colores']).filter((r, i) => i === 0 || r[0] !== nombre)
  wb.Sheets['Colores'] = toSheet(rows)
  return saveWB(wb, filePath)
})

// ════════════════════════════════════════════════════════════
//  INVENTARIO
//  Hoja "Inventario": SKU | Modelo | Color | CodigoColor | Pares | CostoProduccion
// ════════════════════════════════════════════════════════════
const INV_HDR = ['SKU', 'Modelo', 'Color', 'CodigoColor', 'Pares', 'CostoProduccion']

ipcMain.handle('get-inventario', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  ensureSheet(wb, 'Inventario', INV_HDR)
  const rows = toRows(wb.Sheets['Inventario'])
  return rows.slice(1).filter(r => r[0]).map(r => ({
    sku:             r[0],
    modelo:          r[1],
    color:           r[2],
    codigoColor:     r[3],
    pares:           Number(r[4]) || 0,
    costoProduccion: Number(r[5]) || 0
  }))
})

ipcMain.handle('save-inventario', (_, filePath, item) => {
  const wb = readWB(filePath)
  if (!wb) return false
  ensureSheet(wb, 'Inventario', INV_HDR)
  const ws   = wb.Sheets['Inventario']
  const rows = toRows(ws)
  const row  = [item.sku, item.modelo, item.color, item.codigoColor,
                item.pares, item.costoProduccion || 0]
  const idx  = rows.findIndex((r, i) => i > 0 && r[0] === item.sku)
  if (idx >= 0) rows[idx] = row
  else rows.push(row)
  wb.Sheets['Inventario'] = toSheet(rows)
  return saveWB(wb, filePath)
})

ipcMain.handle('descontar-stock', (_, filePath, sku, qty) => {
  const wb = readWB(filePath)
  if (!wb) return { ok: false, msg: 'Excel no encontrado' }
  ensureSheet(wb, 'Inventario', INV_HDR)
  const ws   = wb.Sheets['Inventario']
  const rows = toRows(ws)
  const idx  = rows.findIndex((r, i) => i > 0 && r[0] === sku)
  if (idx < 0) return { ok: false, msg: `SKU ${sku} no encontrado` }
  const actual = Number(rows[idx][4]) || 0
  if (actual < qty) return { ok: false, msg: `Stock insuficiente (${actual} pares disponibles)` }
  rows[idx][4] = actual - qty
  wb.Sheets['Inventario'] = toSheet(rows)
  const ok = saveWB(wb, filePath)
  return { ok, msg: ok ? `Stock actualizado: ${actual - qty} pares` : 'Error al guardar' }
})

ipcMain.handle('delete-inventario', (_, filePath, sku) => {
  const wb = readWB(filePath)
  if (!wb) return false
  ensureSheet(wb, 'Inventario', INV_HDR)
  const rows = toRows(wb.Sheets['Inventario']).filter((r, i) => i === 0 || r[0] !== sku)
  wb.Sheets['Inventario'] = toSheet(rows)
  return saveWB(wb, filePath)
})

// ════════════════════════════════════════════════════════════
//  CLIENTES
//  Hojas: Clientes_Directos | Distribuidores | Mayoreo
//  Cols: Nombre | Contacto | Direccion | RFC | Notas |
//        FechaRegistro | FechaUltimaCompra |
//        PiezasUltimaCompra | AcumuladoHistorico | SKUs
// ════════════════════════════════════════════════════════════
const CLI_HDR    = ['Nombre','Contacto','Direccion','RFC','Notas',
                    'FechaRegistro','FechaUltimaCompra',
                    'PiezasUltimaCompra','AcumuladoHistorico','SKUs']
const TIPO_SHEET = {
  directo:      'Clientes_Directos',
  distribuidor: 'Distribuidores',
  mayoreo:      'Mayoreo'
}

ipcMain.handle('get-clientes', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return { directo: [], distribuidor: [], mayoreo: [] }
  const result = {}
  for (const [tipo, sheet] of Object.entries(TIPO_SHEET)) {
    ensureSheet(wb, sheet, CLI_HDR)
    const rows = toRows(wb.Sheets[sheet])
    result[tipo] = rows.slice(1).filter(r => r[0]).map((r, i) => ({
      _idx:                i,
      nombre:              r[0],
      contacto:            r[1],
      direccion:           r[2],
      rfc:                 r[3],
      notas:               r[4],
      fechaRegistro:       r[5],
      fechaUltimaCompra:   r[6],
      piezasUltimaCompra:  Number(r[7]) || 0,
      acumuladoHistorico:  Number(r[8]) || 0,
      skus:                r[9] ? String(r[9]).split(',').map(s => s.trim()) : []
    }))
  }
  return result
})

ipcMain.handle('save-cliente', (_, filePath, cliente) => {
  let wb = readWB(filePath)
  if (!wb) { wb = XLSX.utils.book_new() }
  const sheet = TIPO_SHEET[cliente.tipo]
  if (!sheet) return false
  ensureSheet(wb, sheet, CLI_HDR)
  const ws   = wb.Sheets[sheet]
  const rows = toRows(ws)
  const skusStr = Array.isArray(cliente.skus) ? cliente.skus.join(', ') : (cliente.skus || '')
  const row  = [
    cliente.nombre, cliente.contacto || '', cliente.direccion || '',
    cliente.rfc || '', cliente.notas || '',
    cliente.fechaRegistro || '', cliente.fechaUltimaCompra || '',
    cliente.piezasUltimaCompra || 0, cliente.acumuladoHistorico || 0,
    skusStr
  ]
  const editIdx = cliente._editIndex !== undefined ? cliente._editIndex + 1 : -1
  if (editIdx > 0 && editIdx < rows.length) rows[editIdx] = row
  else rows.push(row)
  wb.Sheets[sheet] = toSheet(rows)
  return saveWB(wb, filePath)
})

ipcMain.handle('delete-cliente', (_, filePath, rowIndex, tipo) => {
  const wb = readWB(filePath)
  if (!wb) return false
  const sheet = TIPO_SHEET[tipo]
  if (!sheet) return false
  ensureSheet(wb, sheet, CLI_HDR)
  const rows = toRows(wb.Sheets[sheet])
  rows.splice(rowIndex + 1, 1)
  wb.Sheets[sheet] = toSheet(rows)
  return saveWB(wb, filePath)
})

// ════════════════════════════════════════════════════════════
//  G-CODE — parse y cálculo de costos
// ════════════════════════════════════════════════════════════
function parseGcodeFile (filePath) {
  try {
    const text = fs.readFileSync(filePath, 'utf8')
    const lines = text.split('\n').slice(0, 200) // solo encabezado

    const find = (patterns) => {
      for (const line of lines) {
        for (const pat of patterns) {
          const m = line.match(pat)
          if (m) return m[1].trim()
        }
      }
      return null
    }

    // Peso en gramos
    const pesoStr = find([
      /;\s*filament\s+used\s*\[g\]\s*=\s*([\d.]+)/i,
      /;\s*total\s+filament\s+used\s*=\s*([\d.]+)\s*g/i
    ])

    // Tiempo de impresión
    const tiempoStr = find([
      /;\s*estimated\s+printing\s+time\s*=\s*(.+)/i,
      /;\s*total\s+estimated\s+time\s*:\s*(.+)/i
    ])

    // Convertir tiempo a horas decimales
    let horas = 0
    if (tiempoStr) {
      const h = tiempoStr.match(/(\d+)h/)
      const m = tiempoStr.match(/(\d+)m/)
      const s = tiempoStr.match(/(\d+)s/)
      horas = (h ? +h[1] : 0) + (m ? +m[1] / 60 : 0) + (s ? +s[1] / 3600 : 0)
    }

    // Tipo de filamento
    const filamento = find([
      /;\s*filament_type\s*=\s*(.+)/i,
      /;\s*filament\s+type\s*=\s*(.+)/i
    ])

    return {
      pesoGr:    pesoStr ? parseFloat(pesoStr) : null,
      horasImp:  horas   || null,
      filamento: filamento || null,
      tiempoRaw: tiempoStr || null
    }
  } catch (e) {
    return { error: e.message }
  }
}

ipcMain.handle('parse-gcode', (_, filePath) => parseGcodeFile(filePath))

ipcMain.handle('calc-costo', (_, params) => {
  // params: { pesoGr, horasImp, costoPorKg, costoKwh, wattsPrinter,
  //           costoHerrajes, porcDesperdicio, costoEmpaque }
  const { pesoGr, horasImp, costoPorKg, costoKwh, wattsPrinter,
          costoHerrajes, porcDesperdicio, costoEmpaque } = params

  const costoFilamento   = (pesoGr / 1000) * costoPorKg
  const costoElectricidad = (horasImp * wattsPrinter / 1000) * costoKwh
  const subtotal1        = costoFilamento + costoElectricidad + costoHerrajes + costoEmpaque
  const desperdicio      = subtotal1 * (porcDesperdicio / 100)
  const total            = subtotal1 + desperdicio

  return {
    costoFilamento:    +costoFilamento.toFixed(2),
    costoElectricidad: +costoElectricidad.toFixed(2),
    costoHerrajes:     +costoHerrajes.toFixed(2),
    costoEmpaque:      +costoEmpaque.toFixed(2),
    desperdicio:       +desperdicio.toFixed(2),
    total:             +total.toFixed(2)
  }
})

// ════════════════════════════════════════════════════════════
//  COSTOS — historial en Excel
//  Hoja "Costos": SKU | Modelo | Color | PesoGr | HorasImp |
//                 CostoFilamento | CostoElec | CostoHerrajes |
//                 CostoEmpaque | Desperdicio | CostoTotal | Fecha
// ════════════════════════════════════════════════════════════
const COST_HDR = ['SKU','Modelo','Color','PesoGr','HorasImp',
                  'CostoFilamento','CostoElec','CostoHerrajes',
                  'CostoEmpaque','Desperdicio','CostoTotal','Fecha']

ipcMain.handle('get-costos', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  ensureSheet(wb, 'Costos', COST_HDR)
  const rows = toRows(wb.Sheets['Costos'])
  return rows.slice(1).filter(r => r[0]).map(r => ({
    sku:             r[0],  modelo:    r[1],  color:    r[2],
    pesoGr:          r[3],  horasImp:  r[4],
    costoFilamento:  r[5],  costoElec: r[6],  costoHerrajes: r[7],
    costoEmpaque:    r[8],  desperdicio: r[9],
    costoTotal:      r[10], fecha:     r[11]
  }))
})

ipcMain.handle('save-costo', (_, filePath, item) => {
  const wb = readWB(filePath)
  if (!wb) return false
  ensureSheet(wb, 'Costos', COST_HDR)
  const ws   = wb.Sheets['Costos']
  const rows = toRows(ws)
  rows.push([
    item.sku, item.modelo, item.color, item.pesoGr, item.horasImp,
    item.costoFilamento, item.costoElec, item.costoHerrajes,
    item.costoEmpaque, item.desperdicio, item.costoTotal,
    new Date().toLocaleDateString('es-MX')
  ])
  wb.Sheets['Costos'] = toSheet(rows)
  return saveWB(wb, filePath)
})


// ════════════════════════════════════════════════════════════
//  MOONRAKER — integración con impresora Klipper
// ════════════════════════════════════════════════════════════
const http = require('http')

function moonrakerRequest (ip, method, endpoint, body) {
  return new Promise((resolve, reject) => {
    const data    = body ? JSON.stringify(body) : null
    const options = {
      hostname: ip.split(':')[0],
      port:     parseInt(ip.split(':')[1]) || 7125,
      path:     endpoint,
      method,
      headers:  {
        'Content-Type':   'application/json',
        'Content-Length': data ? Buffer.byteLength(data) : 0
      },
      timeout: 8000
    }
    const req = http.request(options, res => {
      let raw = ''
      res.on('data', d => raw += d)
      res.on('end', () => {
        try { resolve(JSON.parse(raw)) }
        catch { resolve({ raw }) }
      })
    })
    req.on('error',   e => reject(e))
    req.on('timeout', () => { req.destroy(); reject(new Error('Timeout')) })
    if (data) req.write(data)
    req.end()
  })
}

// Estado de la impresora
ipcMain.handle('moonraker-status', async (_, ip) => {
  try {
    const r = await moonrakerRequest(ip, 'GET',
      '/printer/objects/query?print_stats&heater_bed&extruder', null)
    const ps = r?.result?.status?.print_stats || {}
    const bed = r?.result?.status?.heater_bed || {}
    const ext = r?.result?.status?.extruder   || {}
    return {
      ok:       true,
      state:    ps.state         || 'unknown',
      filename: ps.filename      || '',
      progress: ps.print_duration|| 0,
      bedTemp:  bed.temperature  || 0,
      extTemp:  ext.temperature  || 0
    }
  } catch (e) {
    return { ok: false, state: 'offline', error: e.message }
  }
})

// Enviar macro GCode (ZOffset etc.)
ipcMain.handle('moonraker-gcode', async (_, ip, script) => {
  try {
    const r = await moonrakerRequest(ip, 'POST',
      '/printer/gcode/script', { script })
    return { ok: true, result: r }
  } catch (e) {
    return { ok: false, error: e.message }
  }
})

// Monitorear progreso de impresión
ipcMain.handle('moonraker-print-status', async (_, ip) => {
  try {
    const r = await moonrakerRequest(ip, 'GET',
      '/printer/objects/query?print_stats', null)
    const ps = r?.result?.status?.print_stats || {}
    return {
      ok:       true,
      state:    ps.state         || 'unknown',
      filename: ps.filename      || '',
      duration: ps.print_duration|| 0,
      totalDur: ps.total_duration|| 0
    }
  } catch (e) {
    return { ok: false, state: 'offline', error: e.message }
  }
})

// ════════════════════════════════════════════════════════════
//  ORCASLICER — abrir con STLs seleccionados
// ════════════════════════════════════════════════════════════
const { spawn } = require('child_process')

ipcMain.handle('open-orcaslicer', async (_, orcaPath, stlPaths) => {
  try {
    if (!fs.existsSync(orcaPath)) {
      return { ok: false, error: `OrcaSlicer no encontrado en: ${orcaPath}` }
    }
    // Verificar que los STLs existen
    const missing = stlPaths.filter(p => !fs.existsSync(p))
    if (missing.length > 0) {
      return { ok: false, error: `STLs no encontrados: ${missing.join(', ')}` }
    }
    // Abrir OrcaSlicer con los STLs como argumentos
    spawn(orcaPath, stlPaths, {
      detached: true,
      stdio:    'ignore'
    }).unref()
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e.message }
  }
})

// Seleccionar ejecutable de OrcaSlicer
ipcMain.handle('select-orca-exe', async (_, def) => {
  const r = await dialog.showOpenDialog({
    filters:     [{ name: 'Ejecutable', extensions: ['exe'] }],
    defaultPath: def || 'C:\\Program Files\\OrcaSlicer'
  })
  return r.canceled ? null : r.filePaths[0]
})
