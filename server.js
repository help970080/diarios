const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Almacenamiento JSON
const dataFile = path.join(__dirname, 'reportes.json');
let reportes = {};
try { if (fs.existsSync(dataFile)) reportes = JSON.parse(fs.readFileSync(dataFile, 'utf8')); } catch(e) {}
const guardar = () => { try { fs.writeFileSync(dataFile, JSON.stringify(reportes), 'utf8'); } catch(e) {} };
const limpiar = () => { const ahora = Date.now(); Object.keys(reportes).forEach(id => { if (ahora - reportes[id].timestamp > 604800000) delete reportes[id]; }); guardar(); };
limpiar(); setInterval(limpiar, 3600000);
const genId = () => Math.random().toString(36).substring(2, 12);

app.use(express.json());

// PAGINA PRINCIPAL
app.get('/', (req, res) => {
  res.send(`<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Cobranza - Promocash</title>
<script src="https://cdn.tailwindcss.com"></script>
<style>
.gradient-bg{background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%)}
.card-glass{background:rgba(255,255,255,0.05);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,0.1)}
.drop-zone{border:2px dashed #475569;transition:all 0.3s}
.drop-zone:hover,.drop-zone.dragover{border-color:#22d3ee;background:rgba(34,211,238,0.1)}
.file-ok{border-color:#10b981!important;background:rgba(16,185,129,0.1)!important}
</style>
</head>
<body class="gradient-bg min-h-screen text-white">
<div class="container mx-auto px-4 py-8 max-w-2xl">
<div class="text-center mb-8">
<h1 class="text-3xl font-bold mb-2">📊 Cobranza Diaria</h1>
<p class="text-cyan-400 font-semibold">PROMOCASH</p>
</div>
<div class="card-glass rounded-2xl p-6">
<h2 class="text-lg font-semibold mb-4">Sube los 3 archivos</h2>
<form id="f" enctype="multipart/form-data">
<div class="grid grid-cols-3 gap-3 mb-6">
<div class="drop-zone rounded-xl p-3 text-center cursor-pointer" id="d1" onclick="document.getElementById('f1').click()">
<input type="file" name="mora" id="f1" accept=".xlsx,.xls,.csv" class="hidden" onchange="ok(this,'d1','s1')" required>
<div class="text-2xl mb-1">📁</div>
<p class="font-medium text-amber-400 text-xs">MORA</p>
<p class="text-xs text-gray-400" id="s1">Subir</p>
</div>
<div class="drop-zone rounded-xl p-3 text-center cursor-pointer" id="d2" onclick="document.getElementById('f2').click()">
<input type="file" name="vigente" id="f2" accept=".xlsx,.xls,.csv" class="hidden" onchange="ok(this,'d2','s2')" required>
<div class="text-2xl mb-1">📁</div>
<p class="font-medium text-violet-400 text-xs">VIGENTE</p>
<p class="text-xs text-gray-400" id="s2">Subir</p>
</div>
<div class="drop-zone rounded-xl p-3 text-center cursor-pointer" id="d3" onclick="document.getElementById('f3').click()">
<input type="file" name="cobranza" id="f3" accept=".xlsx,.xls,.csv" class="hidden" onchange="ok(this,'d3','s3')" required>
<div class="text-2xl mb-1">📁</div>
<p class="font-medium text-emerald-400 text-xs">COBRANZA</p>
<p class="text-xs text-gray-400" id="s3">Subir</p>
</div>
</div>
<button type="submit" id="btn" class="w-full py-3 bg-gradient-to-r from-cyan-600 to-blue-600 rounded-xl font-semibold hover:from-cyan-500 hover:to-blue-500">🚀 Procesar</button>
</form>
<div id="load" class="hidden text-center py-4"><div class="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-cyan-400"></div></div>
<div id="res" class="hidden mt-6 p-4 bg-emerald-900/30 rounded-xl border border-emerald-500/30">
<p class="text-emerald-400 font-medium mb-2">✅ ¡Listo!</p>
<div class="flex gap-2 mb-3">
<input type="text" id="link" readonly class="flex-1 bg-slate-800 border border-slate-600 rounded-lg px-3 py-2 text-sm">
<button onclick="navigator.clipboard.writeText(document.getElementById('link').value);alert('✓ Copiado!')" class="px-3 py-2 bg-cyan-600 hover:bg-cyan-500 rounded-lg text-sm">📋</button>
</div>
<a id="ir" href="#" class="block text-center py-2 bg-violet-600 hover:bg-violet-500 rounded-lg text-sm font-medium">👁️ Ver Reporte</a>
</div>
</div>
<div class="text-center mt-6 text-gray-500 text-xs"><p>Reportes válidos por 7 días</p></div>
</div>
<script>
function ok(i,d,s){if(i.files[0]){document.getElementById(d).classList.add('file-ok');document.getElementById(s).textContent='✓'}}
document.getElementById('f').onsubmit=async e=>{
e.preventDefault();
document.getElementById('btn').classList.add('hidden');
document.getElementById('load').classList.remove('hidden');
document.getElementById('res').classList.add('hidden');
try{
const r=await fetch('/api/procesar',{method:'POST',body:new FormData(e.target)});
const d=await r.json();
if(d.success){document.getElementById('link').value=location.origin+d.url;document.getElementById('ir').href=d.url;document.getElementById('res').classList.remove('hidden')}
else alert('Error: '+d.error);
}catch(err){alert('Error: '+err.message)}
document.getElementById('load').classList.add('hidden');
document.getElementById('btn').classList.remove('hidden');
};
</script>
</body>
</html>`);
});

// PAGINA REPORTE
app.get('/reporte/:id', (req, res) => {
  res.send(`<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Reporte - Promocash</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>.gradient-bg{background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%)}.card-glass{background:rgba(255,255,255,0.05);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,0.1)}</style>
</head>
<body class="gradient-bg min-h-screen text-white">
<div class="container mx-auto px-4 py-6 max-w-6xl">
<div class="text-center mb-6">
<h1 class="text-2xl font-bold">📊 Reporte de Cobranza</h1>
<p class="text-cyan-400 font-semibold">PROMOCASH</p>
<p class="text-gray-400 text-sm" id="fecha"></p>
</div>
<div id="load" class="text-center py-20"><div class="inline-block animate-spin rounded-full h-10 w-10 border-b-2 border-cyan-400"></div></div>
<div id="cont" class="hidden">
<div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
<div class="bg-gradient-to-br from-emerald-600 to-emerald-800 rounded-xl p-4"><p class="text-emerald-200 text-xs">Total Cobrado</p><p class="text-2xl font-bold" id="k1">$0</p></div>
<div class="bg-gradient-to-br from-blue-600 to-blue-800 rounded-xl p-4"><p class="text-blue-200 text-xs">Cobertura</p><p class="text-2xl font-bold" id="k2">0%</p></div>
<div class="bg-gradient-to-br from-amber-600 to-orange-700 rounded-xl p-4"><p class="text-amber-200 text-xs">MORA</p><p class="text-2xl font-bold" id="k3">0%</p><p class="text-amber-300 text-xs" id="k3d"></p></div>
<div class="bg-gradient-to-br from-violet-600 to-purple-800 rounded-xl p-4"><p class="text-violet-200 text-xs">VIGENTE</p><p class="text-2xl font-bold" id="k4">0%</p><p class="text-violet-300 text-xs" id="k4d"></p></div>
</div>
<div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
<div class="card-glass rounded-xl p-4"><h3 class="text-sm font-semibold mb-2">📍 Por Agencia</h3><div class="h-44"><canvas id="c1"></canvas></div></div>
<div class="card-glass rounded-xl p-4"><h3 class="text-sm font-semibold mb-2">📈 Distribución</h3><div class="h-44"><canvas id="c2"></canvas></div></div>
</div>
<div class="card-glass rounded-xl p-4 mb-6">
<h3 class="text-sm font-semibold mb-2">🏢 Por Agencia</h3>
<div class="overflow-x-auto"><table class="w-full text-xs"><thead><tr class="border-b border-gray-700"><th class="text-left p-2">Agencia</th><th class="text-center p-2">Mora</th><th class="text-center p-2">Vigente</th><th class="text-center p-2">Total</th><th class="text-center p-2">Cobrados</th><th class="text-center p-2">%</th><th class="text-right p-2">Monto</th></tr></thead><tbody id="tabla"></tbody></table></div>
</div>
<div class="card-glass rounded-xl p-4 mb-6">
<h3 class="text-sm font-semibold mb-2">📥 Descargar</h3>
<div class="grid grid-cols-3 gap-2">
<a id="dl1" class="py-2 bg-amber-600 hover:bg-amber-500 rounded-lg text-xs font-medium text-center">⬇️ Mora</a>
<a id="dl2" class="py-2 bg-violet-600 hover:bg-violet-500 rounded-lg text-xs font-medium text-center">⬇️ Vigente</a>
<a id="dl3" class="py-2 bg-cyan-600 hover:bg-cyan-500 rounded-lg text-xs font-medium text-center">⬇️ Reporte</a>
</div>
</div>
<div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
<div class="card-glass rounded-xl p-4"><h3 class="text-sm font-semibold mb-2">⚠️ MORA con Pago (<span id="cm">0</span>)</h3><div class="overflow-y-auto max-h-40 text-xs"><table class="w-full"><tbody id="tm"></tbody></table></div></div>
<div class="card-glass rounded-xl p-4"><h3 class="text-sm font-semibold mb-2">✅ VIGENTE con Pago (<span id="cv">0</span>)</h3><div class="overflow-y-auto max-h-40 text-xs"><table class="w-full"><tbody id="tv"></tbody></table></div></div>
</div>
<a href="/" class="block mt-6 text-center py-3 bg-slate-700 hover:bg-slate-600 rounded-xl text-sm">➕ Nuevo reporte</a>
</div>
<div id="err" class="hidden text-center py-20"><p class="text-red-400 text-xl">❌ Reporte no encontrado</p><a href="/" class="inline-block mt-4 px-6 py-2 bg-cyan-600 rounded-lg">Crear nuevo</a></div>
</div>
<script>
const id=location.pathname.split('/').pop();
fetch('/api/reporte/'+id).then(r=>r.json()).then(d=>{
if(d.error)throw 0;
document.getElementById('load').classList.add('hidden');
document.getElementById('cont').classList.remove('hidden');
const r=d.datos.resumen;
document.getElementById('fecha').textContent=d.fecha;
document.getElementById('k1').textContent='$'+r.total.monto.toLocaleString();
document.getElementById('k2').textContent=r.total.porcentaje+'%';
document.getElementById('k3').textContent=r.mora.porcentaje+'%';
document.getElementById('k3d').textContent=r.mora.cobrados+'/'+r.mora.total+' $'+r.mora.monto.toLocaleString();
document.getElementById('k4').textContent=r.vigente.porcentaje+'%';
document.getElementById('k4d').textContent=r.vigente.cobrados+'/'+r.vigente.total+' $'+r.vigente.monto.toLocaleString();
document.getElementById('dl1').href='/api/descargar/'+id+'/mora';
document.getElementById('dl2').href='/api/descargar/'+id+'/vigente';
document.getElementById('dl3').href='/api/descargar/'+id+'/reporte';
let h='',tM=0,tV=0,tC=0,tCob=0,tMonto=0;
r.porAgencia.forEach(a=>{const pc=a.porcentaje>=50?'text-emerald-400':a.porcentaje>=20?'text-yellow-400':'text-red-400';h+='<tr class="border-b border-gray-800"><td class="p-2">'+a.agencia+'</td><td class="p-2 text-center text-amber-400">'+a.mora+'</td><td class="p-2 text-center text-violet-400">'+a.vigente+'</td><td class="p-2 text-center">'+a.total+'</td><td class="p-2 text-center text-emerald-400">'+a.cobrados+'</td><td class="p-2 text-center '+pc+'">'+a.porcentaje+'%</td><td class="p-2 text-right text-emerald-400">$'+a.monto.toLocaleString()+'</td></tr>';tM+=a.mora;tV+=a.vigente;tC+=a.total;tCob+=a.cobrados;tMonto+=a.monto;});
h+='<tr class="border-t-2 border-cyan-500 font-bold bg-slate-800/50"><td class="p-2">TOTAL</td><td class="p-2 text-center">'+tM+'</td><td class="p-2 text-center">'+tV+'</td><td class="p-2 text-center">'+tC+'</td><td class="p-2 text-center text-emerald-400">'+tCob+'</td><td class="p-2 text-center text-cyan-400">'+r.total.porcentaje+'%</td><td class="p-2 text-right text-emerald-400">$'+tMonto.toLocaleString()+'</td></tr>';
document.getElementById('tabla').innerHTML=h;
document.getElementById('cm').textContent=r.detalleMora.length;
document.getElementById('cv').textContent=r.detalleVigente.length;
document.getElementById('tm').innerHTML=r.detalleMora.map(c=>'<tr class="border-b border-gray-800"><td class="p-1">'+c.cliente+'</td><td class="p-1 text-amber-400">'+c.agencia+'</td><td class="p-1 text-right text-emerald-400">$'+c.cobranza.toLocaleString()+'</td></tr>').join('');
document.getElementById('tv').innerHTML=r.detalleVigente.sort((a,b)=>b.cobranza-a.cobranza).map(c=>'<tr class="border-b border-gray-800"><td class="p-1">'+c.cliente+'</td><td class="p-1 text-violet-400">'+c.agencia+'</td><td class="p-1 text-right text-emerald-400">$'+c.cobranza.toLocaleString()+'</td></tr>').join('');
new Chart(document.getElementById('c1'),{type:'bar',data:{labels:r.porAgencia.map(a=>a.agencia),datasets:[{data:r.porAgencia.map(a=>a.cobrados),backgroundColor:'#10b981',borderRadius:4},{data:r.porAgencia.map(a=>a.total-a.cobrados),backgroundColor:'#475569',borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,scales:{x:{stacked:true,grid:{display:false},ticks:{color:'#94a3b8',font:{size:9}}},y:{stacked:true,grid:{color:'#334155'},ticks:{color:'#94a3b8'}}},plugins:{legend:{display:false}}}});
new Chart(document.getElementById('c2'),{type:'doughnut',data:{labels:['MORA sin pago','MORA cobrado','VIGENTE sin pago','VIGENTE cobrado'],datasets:[{data:[r.mora.total-r.mora.cobrados,r.mora.cobrados,r.vigente.total-r.vigente.cobrados,r.vigente.cobrados],backgroundColor:['#78350f','#f59e0b','#4c1d95','#8b5cf6'],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{color:'#94a3b8',font:{size:9}}}}}});
}).catch(()=>{document.getElementById('load').classList.add('hidden');document.getElementById('err').classList.remove('hidden')});
</script>
</body>
</html>`);
});

// API PROCESAR
app.post('/api/procesar', upload.fields([
  { name: 'mora', maxCount: 1 },
  { name: 'vigente', maxCount: 1 },
  { name: 'cobranza', maxCount: 1 }
]), (req, res) => {
  try {
    const leer = b => XLSX.utils.sheet_to_json(XLSX.read(b, { type: 'buffer' }).Sheets[XLSX.read(b, { type: 'buffer' }).SheetNames[0]]);
    const mora = leer(req.files.mora[0].buffer);
    const vigente = leer(req.files.vigente[0].buffer);
    const cobranza = leer(req.files.cobranza[0].buffer);
    const resultado = procesar(mora, vigente, cobranza);
    const id = genId();
    const fecha = new Date().toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    reportes[id] = { fecha, datos: resultado, timestamp: Date.now() };
    guardar();
    res.json({ success: true, id, url: '/reporte/' + id });
  } catch (e) { console.error(e); res.status(500).json({ success: false, error: e.message }); }
});

// API REPORTE
app.get('/api/reporte/:id', (req, res) => {
  const r = reportes[req.params.id];
  if (!r) return res.status(404).json({ error: 'No encontrado' });
  res.json({ fecha: r.fecha, datos: r.datos });
});

// API DESCARGAR
app.get('/api/descargar/:id/:tipo', (req, res) => {
  const r = reportes[req.params.id];
  if (!r) return res.status(404).send('No encontrado');
  const d = r.datos;
  if (req.params.tipo === 'mora') {
    const ws = XLSX.utils.json_to_sheet(d.moraData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Mora');
    res.setHeader('Content-Disposition', 'attachment; filename=mora_actualizado.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return res.send(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  }
  if (req.params.tipo === 'vigente') {
    const ws = XLSX.utils.json_to_sheet(d.vigenteData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Vigente');
    res.setHeader('Content-Disposition', 'attachment; filename=vigente_actualizado.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return res.send(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  }
  const wb = XLSX.utils.book_new();
  const arr = [['REPORTE PROMOCASH'], ['Fecha:', r.fecha], [], ['RESUMEN'], ['', 'Total', 'Cobrados', '%', 'Monto'],
    ['MORA', d.resumen.mora.total, d.resumen.mora.cobrados, d.resumen.mora.porcentaje + '%', d.resumen.mora.monto],
    ['VIGENTE', d.resumen.vigente.total, d.resumen.vigente.cobrados, d.resumen.vigente.porcentaje + '%', d.resumen.vigente.monto],
    ['TOTAL', d.resumen.total.total, d.resumen.total.cobrados, d.resumen.total.porcentaje + '%', d.resumen.total.monto],
    [], ['POR AGENCIA'], ['Agencia', 'Mora', 'Vigente', 'Total', 'Cobrados', '%', 'Monto']];
  d.resumen.porAgencia.forEach(a => arr.push([a.agencia, a.mora, a.vigente, a.total, a.cobrados, a.porcentaje + '%', a.monto]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(arr), 'Resumen');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(d.moraData), 'Mora');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(d.vigenteData), 'Vigente');
  res.setHeader('Content-Disposition', 'attachment; filename=reporte_cobranza.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
});

// FUNCION PROCESAR
function procesar(mora, vigente, cobranza) {
  const limpiar = v => { if (!v) return 0; if (typeof v === 'number') return v; return parseFloat(String(v).replace(/[,$]/g, '')) || 0; };
  const norm = n => (n || '').toString().toUpperCase().trim();
  const getAg = r => r.Agencia || r.AGENCIA || 'SIN';
  const getNom = r => r.Cliente || r.CLIENTE || r.Nombre || r.NOMBRE || '';

  const pagos = {};
  cobranza.forEach(r => { const n = norm(r.Nombre || r.NOMBRE || r.Cliente || r.CLIENTE); const m = limpiar(r.Cobrado || r.COBRADO || r.Monto || r.MONTO); if (n) pagos[n] = (pagos[n] || 0) + m; });

  const moraData = mora.map(r => ({ ...r, Cobranza: pagos[norm(getNom(r))] || 0 }));
  const vigenteData = vigente.map(r => { const c = pagos[norm(getNom(r))] || 0; return { ...r, Cobranza: c, 'Cobranza Semanal': c }; });

  const ags = new Set();
  moraData.forEach(r => ags.add(getAg(r)));
  vigenteData.forEach(r => ags.add(getAg(r)));

  const porAgencia = [];
  ags.forEach(ag => {
    const mA = moraData.filter(r => getAg(r) === ag);
    const vA = vigenteData.filter(r => getAg(r) === ag);
    const t = mA.length + vA.length;
    const c = mA.filter(r => r.Cobranza > 0).length + vA.filter(r => r.Cobranza > 0).length;
    const m = mA.reduce((s, r) => s + r.Cobranza, 0) + vA.reduce((s, r) => s + r.Cobranza, 0);
    porAgencia.push({ agencia: ag, mora: mA.length, vigente: vA.length, total: t, cobrados: c, porcentaje: t > 0 ? +((c / t) * 100).toFixed(1) : 0, monto: m });
  });
  porAgencia.sort((a, b) => b.monto - a.monto);

  const tM = moraData.length, cM = moraData.filter(r => r.Cobranza > 0).length, mM = moraData.reduce((s, r) => s + r.Cobranza, 0);
  const tV = vigenteData.length, cV = vigenteData.filter(r => r.Cobranza > 0).length, mV = vigenteData.reduce((s, r) => s + r.Cobranza, 0);

  return {
    resumen: {
      mora: { total: tM, cobrados: cM, porcentaje: tM > 0 ? +((cM / tM) * 100).toFixed(1) : 0, monto: mM },
      vigente: { total: tV, cobrados: cV, porcentaje: tV > 0 ? +((cV / tV) * 100).toFixed(1) : 0, monto: mV },
      total: { total: tM + tV, cobrados: cM + cV, porcentaje: +(((cM + cV) / (tM + tV)) * 100).toFixed(1), monto: mM + mV },
      porAgencia,
      detalleMora: moraData.filter(r => r.Cobranza > 0).map(r => ({ cliente: getNom(r), agencia: getAg(r), cobranza: r.Cobranza })),
      detalleVigente: vigenteData.filter(r => r.Cobranza > 0).map(r => ({ cliente: getNom(r), agencia: getAg(r), cobranza: r.Cobranza }))
    },
    moraData,
    vigenteData
  };
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Servidor en puerto ' + PORT));
