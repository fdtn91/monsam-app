const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('api', {

  // ── Ventana ──────────────────────────────────────────────────────────────
  minimize:       ()           => ipcRenderer.send('win-minimize'),
  maximize:       ()           => ipcRenderer.send('win-maximize'),
  close:          ()           => ipcRenderer.send('win-close'),

  // ── Diálogos de selección ────────────────────────────────────────────────
  selectFolder:   (def)        => ipcRenderer.invoke('select-folder',    def),
  selectExcel:    (def)        => ipcRenderer.invoke('select-excel',     def),
  selectGcodeDir: (def)        => ipcRenderer.invoke('select-gcode-dir', def),
  openExcel:      (p)          => ipcRenderer.invoke('open-excel',       p),

  // ── Catálogo / Fotos / STL ───────────────────────────────────────────────
  selectFoto:          (def)          => ipcRenderer.invoke('select-foto',           def),
  saveFoto:            (ra, cod, src) => ipcRenderer.invoke('save-foto',             ra, cod, src),
  getFoto:             (ra, cod)      => ipcRenderer.invoke('get-foto',              ra, cod),
  getStlBase64:        (ra, car, arc) => ipcRenderer.invoke('get-stl-base64',        ra, car, arc),
  getCatalogoCompleto: (p)            => ipcRenderer.invoke('get-catalogo-completo', p),

  // ── Configuración ────────────────────────────────────────────────────────
  getConfig:  ()    => ipcRenderer.invoke('get-config'),
  saveConfig: (cfg) => ipcRenderer.invoke('save-config', cfg),

  // ── Sincronizar STLs ─────────────────────────────────────────────────────
  sync:              (opts)  => ipcRenderer.invoke('sync',                opts),
  compareDbCarpetas: (opts)  => ipcRenderer.invoke('compare-db-carpetas', opts),
  agregarACatalogo:  (opts)  => ipcRenderer.invoke('agregar-a-catalogo',  opts),
  onSyncLog:  (cb)           => ipcRenderer.on('sync-log', (_, m) => cb(m)),
  offSyncLog: ()             => ipcRenderer.removeAllListeners('sync-log'),

  // ── Colores ──────────────────────────────────────────────────────────────
  getColores:         (p)   => ipcRenderer.invoke('get-colores',          p),
  saveColor:          (p,c) => ipcRenderer.invoke('save-color',           p, c),
  deleteColor:        (p,n) => ipcRenderer.invoke('delete-color',         p, n),
  getCatalogoCodigos: (p)   => ipcRenderer.invoke('get-catalogo-codigos', p),
  refreshStock:       ()    => ipcRenderer.invoke('refresh-stock'),

  // ── Inventario ───────────────────────────────────────────────────────────
  getInventario:   (p)           => ipcRenderer.invoke('get-inventario',    p),
  saveInventario:  (p, item)     => ipcRenderer.invoke('save-inventario',   p, item),
  descontarStock:  (p, sku, qty) => ipcRenderer.invoke('descontar-stock',   p, sku, qty),
  deleteInventario:(p, sku)      => ipcRenderer.invoke('delete-inventario', p, sku),

  // ── Clientes ─────────────────────────────────────────────────────────────
  getClientes:   (p)         => ipcRenderer.invoke('get-clientes',   p),
  saveCliente:   (p, c)      => ipcRenderer.invoke('save-cliente',   p, c),
  deleteCliente: (p, idx, t) => ipcRenderer.invoke('delete-cliente', p, idx, t),

  // ── Costos / G-code ──────────────────────────────────────────────────────
  parseGcode: (filePath) => ipcRenderer.invoke('parse-gcode',  filePath),
  selectGcode:(def)      => ipcRenderer.invoke('select-gcode', def),
  calcCosto:  (params)   => ipcRenderer.invoke('calc-costo',   params),
  saveCosto:  (p, item)  => ipcRenderer.invoke('save-costo',   p, item),
  getCostos:  (p)        => ipcRenderer.invoke('get-costos',   p),

  // ── Moonraker / Impresora ────────────────────────────────────────────────
  moonrakerStatus:      (ip)         => ipcRenderer.invoke('moonraker-status',       ip),
  moonrakerGcode:       (ip, script) => ipcRenderer.invoke('moonraker-gcode',        ip, script),
  moonrakerPrintStatus: (ip)         => ipcRenderer.invoke('moonraker-print-status', ip),

  // ── OrcaSlicer ───────────────────────────────────────────────────────────
  openOrcaSlicer: (orcaPath, stls) => ipcRenderer.invoke('open-orcaslicer', orcaPath, stls),
  selectOrcaExe:  (def)            => ipcRenderer.invoke('select-orca-exe', def),

  // ── Modelos nuevos ───────────────────────────────────────────────────────
  getModelosNuevos:      ()       => ipcRenderer.invoke('get-modelos-nuevos'),
  marcarModeloVisto:     (codigo) => ipcRenderer.invoke('marcar-modelo-visto',      codigo),
  marcarTodosVistos:     ()       => ipcRenderer.invoke('marcar-todos-vistos'),
  registrarModelosNuevos:(items)  => ipcRenderer.invoke('registrar-modelos-nuevos', items),
  eliminarModelo:        (opts)   => ipcRenderer.invoke('eliminar-modelo',          opts),

  // ── Pedidos (recibidos desde app móvil) ──────────────────────────────────
  getPedidos:      (filtro) => ipcRenderer.invoke('get-pedidos',       filtro),
  setEstadoPedido: (id, e)  => ipcRenderer.invoke('set-estado-pedido', id, e),
  deletePedido:    (id)     => ipcRenderer.invoke('delete-pedido',     id),
  onPedidoNuevo:   (cb)     => ipcRenderer.on('pedido-nuevo', (_, d) => cb(d)),
  offPedidoNuevo:  ()       => ipcRenderer.removeAllListeners('pedido-nuevo'),

  // ── Red local ────────────────────────────────────────────────────────────
  getLocalIP:  () => ipcRenderer.invoke('get-local-ip'),
  getApiPort:  () => ipcRenderer.invoke('get-api-port'),
})
