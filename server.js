import express from 'express';
import dotenv from 'dotenv';
import Anthropic from '@anthropic-ai/sdk';
import db from './db.js';
import path from 'path';
import { fileURLToPath } from 'url';
import multer from 'multer';
import XLSX from 'xlsx';

dotenv.config();

const app = express();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});
const upload = multer({ storage: multer.memoryStorage() });

// Ruta de prueba
app.get('/', (req, res) => {
  res.send('Servidor NOVAGRO Combustibles funcionando ✅');
});

// Agregar una carga/consumo de combustible (carga manual desde el formulario)
app.post('/api/combustible', async (req, res) => {
  const { fecha, cisterna, campo, vehiculo, litrosCargados, litrosConsumidos } = req.body;

  await db.read();
  db.data.combustible.push({
    id: Date.now(),
    fecha,
    cisterna,
    campo,
    vehiculo,
    litrosCargados,
    litrosConsumidos,
  });
  await db.write();

  res.json({ ok: true });
});

// Listar todas las cargas
app.get('/api/combustible', async (req, res) => {
  await db.read();
  res.json(db.data.combustible);
});

// Saldos de referencia por campo (el último saldo leído de la planilla, con su fecha)
app.get('/api/saldos-referencia', async (req, res) => {
  await db.read();
  res.json(db.data.saldosReferencia || {});
});

// Editar una carga existente
app.put('/api/combustible/:id', async (req, res) => {
  await db.read();
  const id = Number(req.params.id);
  const idx = db.data.combustible.findIndex(r => r.id === id);
  if (idx === -1) return res.status(404).json({ error: 'Registro no encontrado' });

  const { fecha, cisterna, campo, vehiculo, litrosCargados, litrosConsumidos, costo } = req.body;
  db.data.combustible[idx] = {
    ...db.data.combustible[idx],
    fecha, cisterna, campo, vehiculo,
    litrosCargados: Number(litrosCargados) || 0,
    litrosConsumidos: Number(litrosConsumidos) || 0,
    costo: costo !== undefined ? Number(costo) || 0 : db.data.combustible[idx].costo,
  };
  await db.write();
  res.json({ ok: true });
});

// Eliminar una carga
app.delete('/api/combustible/:id', async (req, res) => {
  await db.read();
  const id = Number(req.params.id);
  const antes = db.data.combustible.length;
  db.data.combustible = db.data.combustible.filter(r => r.id !== id);
  await db.write();
  res.json({ ok: true, eliminado: antes !== db.data.combustible.length });
});

// Importar el Excel que exporta Physis (rptMaquinarias) — cubre el período reciente
app.post('/api/importar-excel', upload.single('archivo'), async (req, res) => {
  try {
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const hoja = workbook.Sheets[workbook.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(hoja);

    // Nos quedamos solo con filas de combustible (GasOil o Nafta)
    const filasCombustible = filas.filter(f =>
      (typeof f.Insumo === 'string' && f.Insumo.trim().startsWith('GasOil')) ||
      (typeof f.Insumo === 'string' && f.Insumo.trim().toLowerCase().startsWith('nafta'))
    );

    await db.read();

    const clavesExistentes = new Set(
      db.data.combustible.map(r => `${r.fecha}|${r.vehiculo}|${r.litrosCargados}`)
    );

    let agregados = 0;
    for (const f of filasCombustible) {
      const fechaStr = f.Fecha instanceof Date
        ? f.Fecha.toISOString().split('T')[0]
        : String(f.Fecha);

      const clave = `${fechaStr}|${f.Maquinaria}|${f.Cantidad}`;
      if (clavesExistentes.has(clave)) continue;
      clavesExistentes.add(clave);

      db.data.combustible.push({
        id: Date.now() + agregados,
        fecha: fechaStr,
        cisterna: f.Labor || '',
        campo: f.descripcionDeposito ? f.descripcionDeposito.trim() : '',
        vehiculo: canonizarVehiculo(f.Maquinaria) || (f.Maquinaria || ''),
        litrosCargados: f.Cantidad || 0,
        litrosConsumidos: 0,
        costo: f.Costo || 0,
        combustible: f.Insumo ? f.Insumo.trim() : '',
      });
      agregados++;
    }

    await db.write();
    res.json({ ok: true, agregados, totalFilasExcel: filas.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'No se pudo procesar el archivo: ' + err.message });
  }
});

function limitarRango(hoja, maxFilas = 20000) {
  if (!hoja['!ref']) return hoja;
  const range = XLSX.utils.decode_range(hoja['!ref']);
  if (range.e.r > maxFilas) {
    range.e.r = maxFilas;
    hoja['!ref'] = XLSX.utils.encode_range(range);
  }
  return hoja;
}

// --- Utilidades de comparación difusa (para el importador histórico) ---

// Diccionario de alias: nombre oficial (como lo escribe Physis) -> variantes
// que aparecen escritas distinto en el histórico manual. Se arma a mano,
// a medida que se van encontrando coincidencias reales — es más confiable
// que intentar adivinar con un algoritmo genérico.
const ALIAS_VEHICULOS = {
  'TRACTOR MASSEY FERGUSON ZXC12': ['MF', 'MF CHICO'],
  'TRACTOR 4299 DFI54': ['MF 4299', 'MF-4299', 'MF4299', '4299', 'MF 299', 'MF299'],
  'TRACTOR 290 BPI59': ['290', 'M 290', 'MF 290', 'MF-290', 'MF290', 'MF 290 I', 'MF 290 II', 'MF 290-1', 'MF 290-2'],
  'TRACTOR 6711/2 ETL75': ['6711/2', '6711 / 2', 'M 6711-2', 'MF 6711-2', 'MF 6711 -2', 'MF 6711- 2'],
  'TRACTOR 265/2': ['MF 265', 'MF-265'],
  'TRACTOR 6713 EFT44': ['6713', 'M 6713', 'MF 6713', 'MF 6713-2'],
  'TRACTOR 6711/1 EFB75': ['6711/1', '6711 / 1', 'M 6711-1', 'MF 6711-1', 'MF 6711 -1', 'MF 6711- 1'],
  'TRACTOR 283 BPI71': ['MF 283', 'MF-283', 'MF283-2', 'MF 283-1', 'MF 283-2', 'MF 283-3', 'MF 283 (1)', 'MF 283 (2)', 'MF 283 (3)', 'MF 283 / 2', 'MF 283 /2', 'MF 283 RULLO', 'MF-283 1', 'MF-283 3', 'MF 2830'],
  'Retro CMF03': ['RETRO', 'MF RETRO', 'MF-RETRO', 'RETRO MF96', 'MF 96 RETRO'],

  // Rodados (camionetas y moto)
  'Toyota Prado LYT907': ['TOYOTA', 'PRADO', 'TOYOTA PRADO'],
  'Volkswagen Amarok AH781MH': ['AMK MH', 'AMAROK MH', 'AH781MH'],
  'Volkswagen Amarok LS': ['AMK LS', 'AMAROK LS'],
  'Volkswagen Amarok AH781MI': ['AMK MI', 'AMAROK MI', 'AH781MI'],
  'Toyota Hilux AE778WY': ['HILUX', 'TOYOTA 778', 'TOYOTA-WY', 'HILUX CLAUDIO'],
  'Nissan Frontier AF064AO': ['NISSAN', 'NISSAN 064', 'AF064'],
  'Toyota Hilux AI217HM': ['HILUX GONZALO', 'HILUX AI217', 'AI217HM'],
  'Moto Keller': ['MOTO', 'KELLER'],

  // Camiones
  'MB 1933 I GOS 164': ['MB 1933 I', 'MB 1933', 'GOS164', 'GOS 164'],
  'MB 2041 I AE623FH': ['MB 2041 I', 'MB 2041', 'AE623FH'],
  'JAULA AG341 VX': ['JAULA', 'AG341VX', 'AG341 VX'],
  'MB 2545 AH820MO': ['MB 2545', 'AH820MO'],
};

function canonizarVehiculo(nombre) {
  const n = (nombre || '').trim().toUpperCase();
  if (!n) return null;
  for (const [canonico, alias] of Object.entries(ALIAS_VEHICULOS)) {
    if (n === canonico.toUpperCase()) return canonico;
    if (alias.some(a => a.toUpperCase() === n)) return canonico;
  }
  return null;
}

// Se queda solo con dígitos y "/" — respaldo para vehículos que no están
// todavía en el diccionario de alias de arriba.
function codigoVehiculo(nombre) {
  return (nombre || '').toUpperCase().replace(/[^0-9/]/g, '');
}

function vehiculoCoincide(nombreHistorico, nombreExistente) {
  const c1 = canonizarVehiculo(nombreHistorico);
  const c2 = canonizarVehiculo(nombreExistente);
  if (c1 && c2) return c1 === c2;

  // Si alguno no está en el diccionario todavía, respaldo por número
  const h = codigoVehiculo(nombreHistorico);
  const e = codigoVehiculo(nombreExistente);
  if (!h || !e) return false;
  return e.includes(h) || h.includes(e);
}

// Tolera el típico error de tipeo de coma decimal (63 en vez de 0.63, o al revés)
function litrosCoinciden(a, b) {
  if (a == null || b == null) return false;
  const diffs = [Math.abs(a - b), Math.abs(a - b * 10), Math.abs(a * 10 - b)];
  return diffs.some(d => d < 1.5);
}

// Importar el histórico manual (Informe Cisternas y Campos).
// Compara cada fila contra lo ya cargado (idealmente ya importaste Physis antes):
// si encuentra una fila existente de la misma fecha, con vehículo y litros
// compatibles, la considera ya cubierta y la salta. Si no encuentra coincidencia,
// la agrega — sin importar si es una fecha "vieja" o "reciente".
app.post('/api/importar-historico', upload.single('archivo'), async (req, res) => {
  try {
    console.log('Iniciando importación histórica...');
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const hojasCisterna = ['PLANTA', 'PALMAR CHICO', 'LA LONJA', 'SELENE'];

    await db.read();

    // Índice por fecha de lo que ya está cargado, para comparar rápido
    const porFecha = {};
    for (const r of db.data.combustible) {
      if (!porFecha[r.fecha]) porFecha[r.fecha] = [];
      porFecha[r.fecha].push(r);
    }

    let agregados = 0;
    let procesadas = 0;
    let yaExistian = 0;

    function aFechaStr(valor) {
      if (valor instanceof Date) return valor.toISOString().split('T')[0];
      return String(valor || '').split('T')[0];
    }

    function yaExisteEnPhysis(fecha, vehiculo, litros, reloj) {
      const candidatas = porFecha[fecha] || [];
      return candidatas.some(r => {
        // Si ambos lados tienen número de reloj, es el criterio más confiable:
        // mismo vehículo + mismo reloj = la misma carga, sin dudas.
        if (reloj != null && r.reloj != null && vehiculoCoincide(vehiculo, r.vehiculo)) {
          return Number(r.reloj) === reloj;
        }
        // Si no hay reloj de un lado (por ejemplo viene de Physis, que no lo registra),
        // respaldo con la comparación anterior por vehículo + litros.
        return vehiculoCoincide(vehiculo, r.vehiculo) && litrosCoinciden(litros, Number(r.litrosCargados));
      });
    }

    function agregarRegistro(fecha, vehiculo, litros, tipo, campo, cisterna, costo, reloj) {
      if (yaExisteEnPhysis(fecha, vehiculo, litros, reloj)) { yaExistian++; return; }
      const nombreCanonico = canonizarVehiculo(vehiculo) || vehiculo; // usa el nombre oficial si está en el diccionario
      const nuevo = {
        id: Date.now() + agregados,
        fecha, cisterna, campo, vehiculo: nombreCanonico,
        litrosCargados: litros,
        litrosConsumidos: tipo === 'salida' ? litros : 0,
        costo: costo || 0,
        combustible: 'GasOil',
        reloj: reloj != null ? reloj : null,
      };
      db.data.combustible.push(nuevo);
      if (!porFecha[fecha]) porFecha[fecha] = [];
      porFecha[fecha].push(nuevo);
      agregados++;
    }

    for (const nombreHoja of hojasCisterna) {
      let hoja = workbook.Sheets[nombreHoja];
      if (!hoja) continue;
      hoja = limitarRango(hoja);
      const filas = XLSX.utils.sheet_to_json(hoja, { defval: null });
      console.log(`Hoja ${nombreHoja}: ${filas.length} filas`);
      procesadas += filas.length;

      for (const f of filas) {
        if (!f.FECHA) continue;
        const fechaStr = aFechaStr(f.FECHA);
        const reloj = (f.RELOJ !== undefined && f.RELOJ !== null && f.RELOJ !== '') ? Number(f.RELOJ) : null;

        // Algunas cargas no tienen VEHICULO (van directo a la cisterna, o quedó
        // sin anotar) — usamos OBSERVACIONES o un identificador genérico en vez
        // de perder la fila.
        let vehiculo = f.VEHICULO ? String(f.VEHICULO).trim() : null;
        if (!vehiculo && ((f.ENTRADA && Number(f.ENTRADA) > 0) || (f.SALIDA && Number(f.SALIDA) > 0))) {
          vehiculo = f.OBSERVACIONES ? String(f.OBSERVACIONES).trim() : `Sin identificar (${nombreHoja})`;
        }
        if (!vehiculo) continue;

        if (f.ENTRADA && Number(f.ENTRADA) > 0) {
          agregarRegistro(fechaStr, vehiculo, Number(f.ENTRADA), 'entrada', nombreHoja, 'Carga de combustible en Cisterna', Number(f.TOTAL), reloj);
        }
        if (f.SALIDA && Number(f.SALIDA) > 0) {
          agregarRegistro(fechaStr, vehiculo, Number(f.SALIDA), 'salida', nombreHoja, 'Consumo histórico', Number(f.TOTAL), reloj);
        }
      }

      // Guardamos el último saldo de esta hoja (el de la fecha más reciente
      // que tenga un valor en la columna SALDO) como punto de referencia.
      let mejorSaldo = null;
      for (const f of filas) {
        if (f.FECHA && f.SALDO != null && f.SALDO !== '') {
          const fechaFila = aFechaStr(f.FECHA);
          if (!mejorSaldo || fechaFila >= mejorSaldo.fecha) {
            mejorSaldo = { fecha: fechaFila, saldo: Number(f.SALDO) };
          }
        }
      }
      if (mejorSaldo) {
        const existente = db.data.saldosReferencia[nombreHoja];
        if (!existente || mejorSaldo.fecha >= existente.fecha) {
          db.data.saldosReferencia[nombreHoja] = mejorSaldo;
          console.log(`Saldo de referencia actualizado para ${nombreHoja}: ${mejorSaldo.saldo} L al ${mejorSaldo.fecha}`);
        }
      }
    }

    let hojaEstaciones = workbook.Sheets['ESTACIONES'];
    if (hojaEstaciones) {
      hojaEstaciones = limitarRango(hojaEstaciones);
      const filas = XLSX.utils.sheet_to_json(hojaEstaciones, { defval: null });
      console.log(`Hoja ESTACIONES: ${filas.length} filas`);
      procesadas += filas.length;

      for (const f of filas) {
        if (!f.FECHA || !f.VEHICULO) continue;
        const fechaStr = aFechaStr(f.FECHA);
        const vehiculo = String(f.VEHICULO).trim();

        let litros = 0;
        if (typeof f.COMBUSTIBLE === 'number') {
          litros = f.COMBUSTIBLE;
        } else if (typeof f.COMBUSTIBLE === 'string') {
          const match = f.COMBUSTIBLE.match(/[\d.,]+/);
          if (match) litros = parseFloat(match[0].replace(',', '.'));
        }
        agregarRegistro(fechaStr, vehiculo, litros, 'salida', 'ESTACION DE SERVICIO', 'Compra en Estación de Servicio', Number(f.TOTAL));
      }
    }

    console.log(`Importación terminada: ${agregados} registros nuevos de ${procesadas} filas procesadas (${yaExistian} ya estaban cubiertas por datos existentes).`);
    await db.write();
    res.json({ ok: true, agregados, filasProcesadas: procesadas, yaExistian });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'No se pudo procesar el archivo: ' + err.message });
  }
});

// Re-aplica el diccionario de alias a todo lo que ya está cargado — útil
// cada vez que se agrega una variante nueva a ALIAS_VEHICULOS, para no
// tener que borrar data.json y reimportar todo de nuevo.
app.post('/api/normalizar-vehiculos', async (req, res) => {
  await db.read();
  let cambios = 0;
  db.data.combustible.forEach(r => {
    const canonico = canonizarVehiculo(r.vehiculo);
    if (canonico && r.vehiculo !== canonico) {
      r.vehiculo = canonico;
      cambios++;
    }
  });
  await db.write();
  res.json({ ok: true, cambios });
});

app.post('/api/informe', async (req, res) => {
  await db.read();
  const datos = (req.body.datos && req.body.datos.length > 0) ? req.body.datos : db.data.combustible;

  if (datos.length === 0) {
    return res.json({ informe: 'No hay datos cargados todavía para generar un informe.' });
  }

  const mensaje = await anthropic.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 2000,
    messages: [{
      role: 'user',
      content: `Sos un asistente que arma informes de consumo de combustible para una empresa agropecuaria (NOVAGRO S.A., Cañada Rosquín).

Con estos datos en JSON (cada registro es una carga: "cisterna" es la labor/descripción, "campo" es la ubicación, "vehiculo" el vehículo o cisterna, "litrosCargados" los litros, "costo" el costo), escribí un informe en español, claro y prolijo, para un jefe que no es técnico, con esta estructura:

1. RESUMEN EJECUTIVO (2-3 líneas con lo más importante)
2. CONSUMO POR CAMPO (agrupá "DEPOSITO PRINCIPAL" como "PLANTA")
3. CONSUMO POR VEHÍCULO/MAQUINARIA (destacá los 3 que más consumieron)
4. ENTRADAS A CISTERNAS vs SALIDAS (las filas cuya "cisterna" dice "Carga de combustible en Cisterna" son entradas; el resto son salidas/consumo)
5. ALERTAS (algún vehículo con consumo fuera de lo normal comparado con el resto, o algún dato que llame la atención)

Datos:
${JSON.stringify(datos, null, 2)}`
    }]
  });

  const informe = mensaje.content[0].text;
  res.json({ informe });
});

app.listen(3000, () => {
  console.log('Servidor corriendo en http://localhost:3000');
});