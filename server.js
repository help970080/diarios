const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const { nanoid } = require('nanoid');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Base de datos SQLite
const db = new Database('./reportes.db');
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

app.use(express.static('public'));
app.use(express.json());

// Página principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

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

    // Procesar
    const resultado = procesarCobranza(mora, vigente, cobranza);
    
    // Guardar en DB
    const id = nanoid(10);
    const fecha = new Date().toLocaleDateString('es-MX', { 
      weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' 
    });
    
    db.prepare('INSERT INTO reportes (id, fecha, datos) VALUES (?, ?, ?)').run(
      id, fecha, JSON.stringify(resultado)
    );

    res.json({ success: true, id, url: `/reporte/${id}` });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// Ver reporte compartido
app.get('/reporte/:id', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'reporte.html'));
});

// API obtener reporte
app.get('/api/reporte/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM reportes WHERE id = ?').get(req.params.id);
  if (!row) {
    return res.status(404).json({ error: 'Reporte no encontrado' });
  }
  res.json({ fecha: row.fecha, datos: JSON.parse(row.datos) });
});

// Descargar Excel
app.get('/api/descargar/:id/:tipo', (req, res) => {
  const row = db.prepare('SELECT * FROM reportes WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).send('No encontrado');
  
  const datos = JSON.parse(row.datos);
  let data, filename;
  
  if (req.params.tipo === 'mora') {
    data = datos.moraData;
    filename = 'mora_actualizado.xlsx';
  } else if (req.params.tipo === 'vigente') {
    data = datos.vigenteData;
    filename = 'vigente_actualizado.xlsx';
  } else {
    // Reporte completo
    const wb = XLSX.utils.book_new();
    const resumenData = [
      ['REPORTE DE COBRANZA PROMOCASH'],
      ['Fecha:', row.fecha],
      [],
      ['RESUMEN GENERAL'],
      ['Cartera', 'Total', 'Cobrados', '%', 'Monto'],
      ['MORA', datos.resumen.mora.total, datos.resumen.mora.cobrados, datos.resumen.mora.porcentaje + '%', datos.resumen.mora.monto],
      ['VIGENTE', datos.resumen.vigente.total, datos.resumen.vigente.cobrados, datos.resumen.vigente.porcentaje + '%', datos.resumen.vigente.monto],
      ['TOTAL', datos.resumen.total.total, datos.resumen.total.cobrados, datos.resumen.total.porcentaje + '%', datos.resumen.total.monto],
      [],
      ['POR AGENCIA'],
      ['Agencia', 'Mora', 'Vigente', 'Total', 'Cobrados', '%', 'Monto']
    ];
    datos.resumen.porAgencia.forEach(ag => {
      resumenData.push([ag.agencia, ag.mora, ag.vigente, ag.total, ag.cobrados, ag.porcentaje + '%', ag.monto]);
    });
    
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
  
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

function procesarCobranza(mora, vigente, cobranza) {
  const limpiarMonto = (valor) => {
    if (!valor) return 0;
    if (typeof valor === 'number') return valor;
    return parseFloat(String(valor).replace(/[,$]/g, '')) || 0;
  };

  const normalizarNombre = (nombre) => (nombre || '').toString().toUpperCase().trim();
  const getAgencia = (row) => row.Agencia || row.AGENCIA || 'SIN AGENCIA';
  const getNombre = (row) => row.Cliente || row.CLIENTE || row.Nombre || row.NOMBRE || '';

  // Mapa de pagos
  const pagos = {};
  cobranza.forEach(row => {
    const nombre = normalizarNombre(row.Nombre || row.NOMBRE || row.Cliente || row.CLIENTE);
    const monto = limpiarMonto(row.Cobrado || row.COBRADO || row.Monto || row.MONTO);
    if (nombre) {
      if (!pagos[nombre]) pagos[nombre] = 0;
      pagos[nombre] += monto;
    }
  });

  // Procesar mora
  const moraData = mora.map(row => {
    const nombre = normalizarNombre(getNombre(row));
    return { ...row, Cobranza: pagos[nombre] || 0 };
  });

  // Procesar vigente
  const vigenteData = vigente.map(row => {
    const nombre = normalizarNombre(getNombre(row));
    const cobranza = pagos[nombre] || 0;
    return { ...row, Cobranza: cobranza, 'Cobranza Semanal': cobranza };
  });

  // Estadísticas por agencia
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
    porAgencia.push({
      agencia: ag,
      mora: moraAg.length,
      vigente: vigenteAg.length,
      total,
      cobrados,
      porcentaje: total > 0 ? parseFloat(((cobrados / total) * 100).toFixed(1)) : 0,
      monto
    });
  });
  porAgencia.sort((a, b) => b.monto - a.monto);

  // Totales
  const totalMora = moraData.length;
  const cobradosMora = moraData.filter(r => r.Cobranza > 0).length;
  const montoMora = moraData.reduce((s, r) => s + r.Cobranza, 0);

  const totalVigente = vigenteData.length;
  const cobradosVigente = vigenteData.filter(r => r.Cobranza > 0).length;
  const montoVigente = vigenteData.reduce((s, r) => s + r.Cobranza, 0);

  const resumen = {
    mora: {
      total: totalMora,
      cobrados: cobradosMora,
      porcentaje: totalMora > 0 ? parseFloat(((cobradosMora / totalMora) * 100).toFixed(1)) : 0,
      monto: montoMora
    },
    vigente: {
      total: totalVigente,
      cobrados: cobradosVigente,
      porcentaje: totalVigente > 0 ? parseFloat(((cobradosVigente / totalVigente) * 100).toFixed(1)) : 0,
      monto: montoVigente
    },
    total: {
      total: totalMora + totalVigente,
      cobrados: cobradosMora + cobradosVigente,
      porcentaje: parseFloat((((cobradosMora + cobradosVigente) / (totalMora + totalVigente)) * 100).toFixed(1)),
      monto: montoMora + montoVigente
    },
    porAgencia,
    detalleMora: moraData.filter(r => r.Cobranza > 0).map(r => ({ cliente: getNombre(r), agencia: getAgencia(r), cobranza: r.Cobranza })),
    detalleVigente: vigenteData.filter(r => r.Cobranza > 0).map(r => ({ cliente: getNombre(r), agencia: getAgencia(r), cobranza: r.Cobranza }))
  };

  return { resumen, moraData, vigenteData };
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));
