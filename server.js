const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const { nanoid } = require('nanoid');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Detectar directorio base
const baseDir = process.env.RENDER ? '/opt/render/project/src' : __dirname;
const publicDir = path.join(baseDir, 'public');
const dbPath = path.join(baseDir, 'reportes.db');

// Crear carpeta public si no existe
if (!fs.existsSync(publicDir)) {
  fs.mkdirSync(publicDir, { recursive: true });
}

// Base de datos SQLite
const db = new Database(dbPath);
db.exec(`
  CREATE TABLE IF NOT EXISTS reportes (
    id TEXT PRIMARY KEY,
    fecha TEXT,
    datos TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`);

// Limpiar reportes viejos (más de 7 días)
db.exec(`DELETE FROM reportes WHERE created_at < datetime('now', '-7 days')`);

app.use(express.static(publicDir));
app.use(express.json());

// HTML embebido para página principal
const indexHTML = `<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Procesador de Cobranza - Promocash</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        .gradient-bg { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%); }
        .card-glass { background: rgba(255,255,255,0.05); backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.1); }
        .drop-zone { border: 2px dashed #475569; transition: all 0.3s; }
        .drop-zone:hover, .drop-zone.dragover { border-color: #22d3ee; background: rgba(34,211,238,0.1); }
        .file-ok { border-color: #10b981 !important; background: rgba(16,185,129,0.1) !important; }
    </style>
</head>
<body class="gradient-bg min-h-screen text-white">
    <div class="container mx-auto px-4 py-8 max-w-2xl">
        <div class="text-center mb-8">
            <h1 class="text-3xl font-bold mb-2">📊 Procesador de Cobranza</h1>
            <p class="text-cyan-400 font-semibold">PROMOCASH</p>
        </div>
        <div class="card-glass rounded-2xl p-6">
            <h2 class="text-lg font-semibold mb-4">Sube los 3 archivos</h2>
            <form id="formUpload" enctype="multipart/form-data">
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
                    <div class="drop-zone rounded-xl p-4 text-center cursor-pointer" id="dropMora" onclick="document.getElementById('fileMora').click()">
                        <input type="file" name="mora" id="fileMora" accept=".xlsx,.xls,.csv" class="hidden" onchange="handleFile(this, 'Mora')" required>
                        <div class="text-3xl mb-1">📁</div>
                        <p class="font-medium text-amber-400 text-sm">MORA</p>
                        <p class="text-xs text-gray-400 mt-1" id="statusMora">Seleccionar</p>
                    </div>
                    <div class="drop-zone rounded-xl p-4 text-center cursor-pointer" id="dropVigente" onclick="document.getElementById('fileVigente').click()">
                        <input type="file" name="vigente" id="fileVigente" accept=".xlsx,.xls,.csv" class="hidden" onchange="handleFile(this, 'Vigente')" required>
                        <div class="text-3xl mb-1">📁</div>
                        <p class="font-medium text-violet-400 text-sm">VIGENTE</p>
                        <p class="text-xs text-gray-400 mt-1" id="statusVigente">Seleccionar</p>
                    </div>
                    <div class="drop-zone rounded-xl p-4 text-center cursor-pointer" id="dropCobranza" onclick="document.getElementById('fileCobranza').click()">
                        <input type="file" name="cobranza" id="fileCobranza" accept=".xlsx,.xls,.csv" class="hidden" onchange="handleFile(this, 'Cobranza')" required>
                        <div class="text-3xl mb-1">📁</div>
                        <p class="font-medium text-emerald-400 text-sm">COBRANZA</p>
                        <p class="text-xs text-gray-400 mt-1" id="statusCobranza">Seleccionar</p>
                    </div>
                </div>
                <button type="submit" id="btnProcesar" class="w-full py-3 bg-gradient-to-r from-cyan-600 to-blue-600 rounded-xl font-semibold disabled:opacity-50 hover:from-cyan-500 hover:to-blue-500 transition">
                    🚀 Procesar y Generar Link
                </button>
            </form>
            <div id="loading" class="hidden text-center py-4">
                <div class="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-cyan-400"></div>
                <p class="mt-2 text-gray-400">Procesando...</p>
            </div>
            <div id="resultado" class="hidden mt-6 p-4 bg-emerald-900/30 rounded-xl border border-emerald-500/30">
                <p class="text-emerald-400 font-medium mb-2">✅ ¡Listo!</p>
                <p class="text-sm text-gray-300 mb-3">Comparte este link:</p>
                <div class="flex gap-2">
                    <input type="text" id="linkGenerado" readonly class="flex-1 bg-slate-800 border border-slate-600 rounded-lg px-3 py-2 text-sm">
                    <button onclick="copiarLink()" class="px-4 py-2 bg-cyan-600 hover:bg-cyan-500 rounded-lg text-sm font-medium">📋</button>
                </div>
                <a id="linkIr" href="#" class="block mt-3 text-center py-2 bg-violet-600 hover:bg-violet-500 rounded-lg text-sm font-medium">👁️ Ver Reporte</a>
            </div>
        </div>
        <div class="text-center mt-6 text-gray-500 text-xs"><p>Los reportes se guardan por 7 días</p></div>
    </div>
    <script>
        ['Mora', 'Vigente', 'Cobranza'].forEach(tipo => {
            const drop = document.getElementById('drop' + tipo);
            drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('dragover'); });
            drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
            drop.addEventListener('drop', e => {
                e.preventDefault(); drop.classList.remove('dragover');
                const file = e.dataTransfer.files[0];
                if (file) { const input = document.getElementById('file' + tipo); const dt = new DataTransfer(); dt.items.add(file); input.files = dt.files; handleFile(input, tipo); }
            });
        });
        function handleFile(input, tipo) {
            if (input.files[0]) { document.getElementById('drop' + tipo).classList.add('file-ok'); document.getElementById('status' + tipo).textContent = '✓ ' + input.files[0].name.substring(0, 15); }
        }
        document.getElementById('formUpload').addEventListener('submit', async (e) => {
            e.preventDefault();
            const formData = new FormData(e.target);
            document.getElementById('btnProcesar').classList.add('hidden');
            document.getElementById('loading').classList.remove('hidden');
            document.getElementById('resultado').classList.add('hidden');
            try {
                const res = await fetch('/api/procesar', { method: 'POST', body: formData });
                const data = await res.json();
                if (data.success) {
                    const url = window.location.origin + data.url;
                    document.getElementById('linkGenerado').value = url;
                    document.getElementById('linkIr').href = data.url;
                    document.getElementById('resultado').classList.remove('hidden');
                } else { alert('Error: ' + data.error); }
            } catch (err) { alert('Error: ' + err.message); }
            document.getElementById('loading').classList.add('hidden');
            document.getElementById('btnProcesar').classList.remove('hidden');
        });
        function copiarLink() { navigator.clipboard.writeText(document.getElementById('linkGenerado').value); alert('✓ Link copiado!'); }
    </script>
</body>
</html>`;

// HTML para reporte
const reporteHTML = `<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte de Cobranza - Promocash</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        .gradient-bg { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%); }
        .card-glass { background: rgba(255,255,255,0.05); backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.1); }
    </style>
</head>
<body class="gradient-bg min-h-screen text-white">
    <div class="container mx-auto px-4 py-6 max-w-6xl">
        <div class="text-center mb-6">
            <h1 class="text-2xl font-bold mb-1">📊 Reporte de Cobranza</h1>
            <p class="text-cyan-400 font-semibold">PROMOCASH</p>
            <p class="text-gray-400 text-sm mt-1" id="fecha"></p>
        </div>
        <div id="loading" class="text-center py-20">
            <div class="inline-block animate-spin rounded-full h-10 w-10 border-b-2 border-cyan-400"></div>
            <p class="mt-3 text-gray-400">Cargando...</p>
        </div>
        <div id="contenido" class="hidden">
            <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
                <div class="bg-gradient-to-br from-emerald-600 to-emerald-800 rounded-xl p-4">
                    <p class="text-emerald-200 text-xs">Total Cobrado</p>
                    <p class="text-2xl font-bold" id="kpiMonto">$0</p>
                </div>
                <div class="bg-gradient-to-br from-blue-600 to-blue-800 rounded-xl p-4">
                    <p class="text-blue-200 text-xs">Cobertura</p>
                    <p class="text-2xl font-bold" id="kpiCobertura">0%</p>
                </div>
                <div class="bg-gradient-to-br from-amber-600 to-orange-700 rounded-xl p-4">
                    <p class="text-amber-200 text-xs">MORA</p>
                    <p class="text-2xl font-bold" id="kpiMora">0%</p>
                    <p class="text-amber-300 text-xs" id="kpiMoraDetalle"></p>
                </div>
                <div class="bg-gradient-to-br from-violet-600 to-purple-800 rounded-xl p-4">
                    <p class="text-violet-200 text-xs">VIGENTE</p>
                    <p class="text-2xl font-bold" id="kpiVigente">0%</p>
                    <p class="text-violet-300 text-xs" id="kpiVigenteDetalle"></p>
                </div>
            </div>
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
                <div class="card-glass rounded-xl p-4">
                    <h3 class="text-sm font-semibold mb-2">📍 Por Agencia</h3>
                    <div class="h-44"><canvas id="chartAgencias"></canvas></div>
                </div>
                <div class="card-glass rounded-xl p-4">
                    <h3 class="text-sm font-semibold mb-2">📈 Distribución</h3>
                    <div class="h-44"><canvas id="chartDona"></canvas></div>
                </div>
            </div>
            <div class="card-glass rounded-xl p-4 mb-6">
                <h3 class="text-sm font-semibold mb-2">🏢 Por Agencia</h3>
                <div class="overflow-x-auto">
                    <table class="w-full text-xs">
                        <thead><tr class="border-b border-gray-700">
                            <th class="text-left p-2">Agencia</th><th class="text-center p-2">Mora</th><th class="text-center p-2">Vigente</th>
                            <th class="text-center p-2">Total</th><th class="text-center p-2">Cobrados</th><th class="text-center p-2">%</th><th class="text-right p-2">Monto</th>
                        </tr></thead>
                        <tbody id="tablaAgencias"></tbody>
                    </table>
                </div>
            </div>
            <div class="card-glass rounded-xl p-4 mb-6">
                <h3 class="text-sm font-semibold mb-2">📥 Descargar</h3>
                <div class="grid grid-cols-3 gap-2">
                    <a id="dlMora" class="py-2 px-3 bg-amber-600 hover:bg-amber-500 rounded-lg text-xs font-medium text-center">⬇️ Mora</a>
                    <a id="dlVigente" class="py-2 px-3 bg-violet-600 hover:bg-violet-500 rounded-lg text-xs font-medium text-center">⬇️ Vigente</a>
                    <a id="dlReporte" class="py-2 px-3 bg-cyan-600 hover:bg-cyan-500 rounded-lg text-xs font-medium text-center">⬇️ Reporte</a>
                </div>
            </div>
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
                <div class="card-glass rounded-xl p-4">
                    <h3 class="text-sm font-semibold mb-2">⚠️ MORA con Pago (<span id="countMora">0</span>)</h3>
                    <div class="overflow-y-auto max-h-40 text-xs"><table class="w-full">
                        <thead class="sticky top-0 bg-slate-800"><tr class="border-b border-gray-700">
                            <th class="text-left p-1">Cliente</th><th class="text-left p-1">Agencia</th><th class="text-right p-1">Cobrado</th>
                        </tr></thead>
                        <tbody id="tablaMora"></tbody>
                    </table></div>
                </div>
                <div class="card-glass rounded-xl p-4">
                    <h3 class="text-sm font-semibold mb-2">✅ VIGENTE con Pago (<span id="countVigente">0</span>)</h3>
                    <div class="overflow-y-auto max-h-40 text-xs"><table class="w-full">
                        <thead class="sticky top-0 bg-slate-800"><tr class="border-b border-gray-700">
                            <th class="text-left p-1">Cliente</th><th class="text-left p-1">Agencia</th><th class="text-right p-1">Cobrado</th>
                        </tr></thead>
                        <tbody id="tablaVigente"></tbody>
                    </table></div>
                </div>
            </div>
            <a href="/" class="block mt-6 text-center py-3 bg-slate-700 hover:bg-slate-600 rounded-xl font-medium text-sm">➕ Crear nuevo reporte</a>
        </div>
        <div id="error" class="hidden text-center py-20">
            <p class="text-red-400 text-xl">❌ Reporte no encontrado</p>
            <p class="text-gray-400 mt-2">El link puede haber expirado</p>
            <a href="/" class="inline-block mt-4 px-6 py-2 bg-cyan-600 rounded-lg">Crear nuevo</a>
        </div>
    </div>
    <script>
        const id = window.location.pathname.split('/').pop();
        fetch('/api/reporte/' + id).then(r => r.json()).then(data => {
            if (data.error) throw new Error();
            document.getElementById('loading').classList.add('hidden');
            document.getElementById('contenido').classList.remove('hidden');
            const r = data.datos.resumen;
            document.getElementById('fecha').textContent = data.fecha;
            document.getElementById('kpiMonto').textContent = '$' + r.total.monto.toLocaleString();
            document.getElementById('kpiCobertura').textContent = r.total.porcentaje + '%';
            document.getElementById('kpiMora').textContent = r.mora.porcentaje + '%';
            document.getElementById('kpiMoraDetalle').textContent = r.mora.cobrados + '/' + r.mora.total + ' • $' + r.mora.monto.toLocaleString();
            document.getElementById('kpiVigente').textContent = r.vigente.porcentaje + '%';
            document.getElementById('kpiVigenteDetalle').textContent = r.vigente.cobrados + '/' + r.vigente.total + ' • $' + r.vigente.monto.toLocaleString();
            document.getElementById('dlMora').href = '/api/descargar/' + id + '/mora';
            document.getElementById('dlVigente').href = '/api/descargar/' + id + '/vigente';
            document.getElementById('dlReporte').href = '/api/descargar/' + id + '/reporte';
            const tabla = document.getElementById('tablaAgencias');
            let tM=0,tV=0,tC=0,tCob=0,tMonto=0;
            r.porAgencia.forEach(ag => {
                const pc = ag.porcentaje >= 50 ? 'text-emerald-400' : ag.porcentaje >= 20 ? 'text-yellow-400' : 'text-red-400';
                tabla.innerHTML += '<tr class="border-b border-gray-800"><td class="p-2">'+ag.agencia+'</td><td class="p-2 text-center text-amber-400">'+ag.mora+'</td><td class="p-2 text-center text-violet-400">'+ag.vigente+'</td><td class="p-2 text-center">'+ag.total+'</td><td class="p-2 text-center text-emerald-400">'+ag.cobrados+'</td><td class="p-2 text-center '+pc+'">'+ag.porcentaje+'%</td><td class="p-2 text-right text-emerald-400">$'+ag.monto.toLocaleString()+'</td></tr>';
                tM+=ag.mora;tV+=ag.vigente;tC+=ag.total;tCob+=ag.cobrados;tMonto+=ag.monto;
            });
            tabla.innerHTML += '<tr class="border-t-2 border-cyan-500 font-bold bg-slate-800/50"><td class="p-2">TOTAL</td><td class="p-2 text-center">'+tM+'</td><td class="p-2 text-center">'+tV+'</td><td class="p-2 text-center">'+tC+'</td><td class="p-2 text-center text-emerald-400">'+tCob+'</td><td class="p-2 text-center text-cyan-400">'+r.total.porcentaje+'%</td><td class="p-2 text-right text-emerald-400">$'+tMonto.toLocaleString()+'</td></tr>';
            document.getElementById('countMora').textContent = r.detalleMora.length;
            document.getElementById('countVigente').textContent = r.detalleVigente.length;
            document.getElementById('tablaMora').innerHTML = r.detalleMora.map(c => '<tr class="border-b border-gray-800"><td class="p-1">'+c.cliente+'</td><td class="p-1 text-amber-400">'+c.agencia+'</td><td class="p-1 text-right text-emerald-400">$'+c.cobranza.toLocaleString()+'</td></tr>').join('');
            document.getElementById('tablaVigente').innerHTML = r.detalleVigente.sort((a,b)=>b.cobranza-a.cobranza).map(c => '<tr class="border-b border-gray-800"><td class="p-1">'+c.cliente+'</td><td class="p-1 text-violet-400">'+c.agencia+'</td><td class="p-1 text-right text-emerald-400">$'+c.cobranza.toLocaleString()+'</td></tr>').join('');
            new Chart(document.getElementById('chartAgencias'),{type:'bar',data:{labels:r.porAgencia.map(a=>a.agencia),datasets:[{label:'Cobrados',data:r.porAgencia.map(a=>a.cobrados),backgroundColor:'#10b981',borderRadius:4},{label:'Sin Pago',data:r.porAgencia.map(a=>a.total-a.cobrados),backgroundColor:'#475569',borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,scales:{x:{stacked:true,grid:{display:false},ticks:{color:'#94a3b8',font:{size:9}}},y:{stacked:true,grid:{color:'#334155'},ticks:{color:'#94a3b8'}}},plugins:{legend:{display:false}}}});
            new Chart(document.getElementById('chartDona'),{type:'doughnut',data:{labels:['MORA sin pago','MORA cobrado','VIGENTE sin pago','VIGENTE cobrado'],datasets:[{data:[r.mora.total-r.mora.cobrados,r.mora.cobrados,r.vigente.total-r.vigente.cobrados,r.vigente.cobrados],backgroundColor:['#78350f','#f59e0b','#4c1d95','#8b5cf6'],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{color:'#94a3b8',font:{size:9},padding:6}}}}});
        }).catch(() => { document.getElementById('loading').classList.add('hidden'); document.getElementById('error').classList.remove('hidden'); });
    </script>
</body>
</html>`;

// Rutas
app.get('/', (req, res) => res.send(indexHTML));
app.get('/reporte/:id', (req, res) => res.send(reporteHTML));

// Procesar archivos
app.post('/api/procesar', upload.fields([
  { name: 'mora', maxCount: 1 },
  { name: 'vigente', maxCount: 1 },
  { name: 'cobranza', maxCount: 1 }
]), (req, res) => {
  try {
    const leerArchivo = (buffer) => {
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      return XLSX.utils.sheet_to_json(sheet);
    };

    const mora = leerArchivo(req.files.mora[0].buffer);
    const vigente = leerArchivo(req.files.vigente[0].buffer);
    const cobranza = leerArchivo(req.files.cobranza[0].buffer);

    const resultado = procesarCobranza(mora, vigente, cobranza);
    
    const id = nanoid(10);
    const fecha = new Date().toLocaleDateString('es-MX', { 
      weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' 
    });
    
    db.prepare('INSERT INTO reportes (id, fecha, datos) VALUES (?, ?, ?)').run(id, fecha, JSON.stringify(resultado));
    res.json({ success: true, id, url: '/reporte/' + id });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/api/reporte/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM reportes WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ error: 'No encontrado' });
  res.json({ fecha: row.fecha, datos: JSON.parse(row.datos) });
});

app.get('/api/descargar/:id/:tipo', (req, res) => {
  const row = db.prepare('SELECT * FROM reportes WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).send('No encontrado');
  
  const datos = JSON.parse(row.datos);
  let data, filename;
  
  if (req.params.tipo === 'mora') {
    data = datos.moraData; filename = 'mora_actualizado.xlsx';
  } else if (req.params.tipo === 'vigente') {
    data = datos.vigenteData; filename = 'vigente_actualizado.xlsx';
  } else {
    const wb = XLSX.utils.book_new();
    const resumenData = [
      ['REPORTE DE COBRANZA PROMOCASH'], ['Fecha:', row.fecha], [],
      ['RESUMEN GENERAL'], ['Cartera', 'Total', 'Cobrados', '%', 'Monto'],
      ['MORA', datos.resumen.mora.total, datos.resumen.mora.cobrados, datos.resumen.mora.porcentaje + '%', datos.resumen.mora.monto],
      ['VIGENTE', datos.resumen.vigente.total, datos.resumen.vigente.cobrados, datos.resumen.vigente.porcentaje + '%', datos.resumen.vigente.monto],
      ['TOTAL', datos.resumen.total.total, datos.resumen.total.cobrados, datos.resumen.total.porcentaje + '%', datos.resumen.total.monto],
      [], ['POR AGENCIA'], ['Agencia', 'Mora', 'Vigente', 'Total', 'Cobrados', '%', 'Monto']
    ];
    datos.resumen.porAgencia.forEach(ag => resumenData.push([ag.agencia, ag.mora, ag.vigente, ag.total, ag.cobrados, ag.porcentaje + '%', ag.monto]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumenData), 'Resumen');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(datos.moraData), 'Mora');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(datos.vigenteData), 'Vigente');
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', 'attachment; filename=reporte_cobranza.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return res.send(buffer);
  }
  
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Datos');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename=' + filename);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

function procesarCobranza(mora, vigente, cobranza) {
  const limpiarMonto = (v) => { if (!v) return 0; if (typeof v === 'number') return v; return parseFloat(String(v).replace(/[,$]/g, '')) || 0; };
  const normalizarNombre = (n) => (n || '').toString().toUpperCase().trim();
  const getAgencia = (r) => r.Agencia || r.AGENCIA || 'SIN AGENCIA';
  const getNombre = (r) => r.Cliente || r.CLIENTE || r.Nombre || r.NOMBRE || '';

  const pagos = {};
  cobranza.forEach(row => {
    const nombre = normalizarNombre(row.Nombre || row.NOMBRE || row.Cliente || row.CLIENTE);
    const monto = limpiarMonto(row.Cobrado || row.COBRADO || row.Monto || row.MONTO);
    if (nombre) { if (!pagos[nombre]) pagos[nombre] = 0; pagos[nombre] += monto; }
  });

  const moraData = mora.map(row => {
    const nombre = normalizarNombre(getNombre(row));
    return { ...row, Cobranza: pagos[nombre] || 0 };
  });

  const vigenteData = vigente.map(row => {
    const nombre = normalizarNombre(getNombre(row));
    const cob = pagos[nombre] || 0;
    return { ...row, Cobranza: cob, 'Cobranza Semanal': cob };
  });

  const agenciasSet = new Set();
  moraData.forEach(r => agenciasSet.add(getAgencia(r)));
  vigenteData.forEach(r => agenciasSet.add(getAgencia(r)));

  const porAgencia = [];
  agenciasSet.forEach(ag => {
    const moraAg = moraData.filter(r => getAgencia(r) === ag);
    const vigenteAg = vigenteData.filter(r => getAgencia(r) === ag);
    const total = moraAg.length + vigenteAg.length;
    const cobrados = moraAg.filter(r => r.Cobranza > 0).length + vigenteAg.filter(r => r.Cobranza > 0).length;
    const monto = moraAg.reduce((s, r) => s + r.Cobranza, 0) + vigenteAg.reduce((s, r) => s + r.Cobranza, 0);
    porAgencia.push({ agencia: ag, mora: moraAg.length, vigente: vigenteAg.length, total, cobrados, porcentaje: total > 0 ? parseFloat(((cobrados / total) * 100).toFixed(1)) : 0, monto });
  });
  porAgencia.sort((a, b) => b.monto - a.monto);

  const totalMora = moraData.length, cobradosMora = moraData.filter(r => r.Cobranza > 0).length, montoMora = moraData.reduce((s, r) => s + r.Cobranza, 0);
  const totalVigente = vigenteData.length, cobradosVigente = vigenteData.filter(r => r.Cobranza > 0).length, montoVigente = vigenteData.reduce((s, r) => s + r.Cobranza, 0);

  return {
    resumen: {
      mora: { total: totalMora, cobrados: cobradosMora, porcentaje: totalMora > 0 ? parseFloat(((cobradosMora / totalMora) * 100).toFixed(1)) : 0, monto: montoMora },
      vigente: { total: totalVigente, cobrados: cobradosVigente, porcentaje: totalVigente > 0 ? parseFloat(((cobradosVigente / totalVigente) * 100).toFixed(1)) : 0, monto: montoVigente },
      total: { total: totalMora + totalVigente, cobrados: cobradosMora + cobradosVigente, porcentaje: parseFloat((((cobradosMora + cobradosVigente) / (totalMora + totalVigente)) * 100).toFixed(1)), monto: montoMora + montoVigente },
      porAgencia,
      detalleMora: moraData.filter(r => r.Cobranza > 0).map(r => ({ cliente: getNombre(r), agencia: getAgencia(r), cobranza: r.Cobranza })),
      detalleVigente: vigenteData.filter(r => r.Cobranza > 0).map(r => ({ cliente: getNombre(r), agencia: getAgencia(r), cobranza: r.Cobranza }))
    },
    moraData,
    vigenteData
  };
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Servidor en puerto ' + PORT));
