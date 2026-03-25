const express = require('express');
const multer = require('multer');
const OpenAI = require('openai');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024, files: 100 } });

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

app.use(express.static(path.join(__dirname, 'public')));

const SYSTEM_PROMPT = `Eres un asistente especializado en extracción de datos de albaranes de agencias de viajes B2B. Tu salida será procesada automáticamente, por lo que debes seguir cada regla con precisión absoluta.

════════════════════════════════════════
FASE 1 — REGLAS DE EXTRACCIÓN
════════════════════════════════════════

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 1 — ESTRUCTURA DEL DOCUMENTO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Un albarán puede contener uno o varios BLOQUES de servicios. Cada bloque tiene:
  - Una cabecera con: [CODIGO_PROYECTO] DESCRIPCION_PROYECTO
  - Una lista de viajeros (SR./SRA. NOMBRE)
  - Una lista de servicios con sus precios

Los bloques están separados visualmente. Nunca mezcles viajeros ni servicios de bloques distintos.

Si el PDF contiene varios albaranes (varias páginas con distintos "Nº Albaran"), trátalo como documentos independientes, cada uno con su propio n_albaran y subtotal.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 2 — UNA FILA POR PERSONA Y SERVICIO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Cada fila de salida = exactamente 1 servicio para 1 persona.
Si hay 2 viajeros y 3 servicios → mínimo 6 filas para ese bloque (más tasas de emisión si las hay).

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 3 — ASIGNACIÓN DE PRECIOS (MUY IMPORTANTE)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Para determinar el importe_unitario de cada persona, analiza la columna TOTAL del albarán:

CASO A — Precio único compartido:
  Una sola línea de servicio. TOTAL = precio_unitario × nº_viajeros.
  → Divide el precio entre todos los viajeros del bloque.
  → Ejemplo: "TASAS DE EMISION  15,00  30,00" con 2 viajeros → cada uno: 15,00

CASO B — Precios individuales distintos:
  Varias líneas del mismo servicio con precios distintos. El TOTAL va acumulando (no es precio × viajeros).
  → NO dividas. Asigna cada línea al viajero en el ORDEN en que aparecen listados.
  → Ejemplo con TIMASHEV (1º) y HONCHAR (2º):
      Línea 1: "HOTEL EUROPA ARTEIXO  59,09  59,09"  → Timashev: 59,09
      Línea 2: "(sin texto)           56,36  112,73" → Honchar: 56,36
      El TOTAL acumulado (112,73 = 59,09 + 56,36) confirma precios distintos.

CASO C — Servicio exclusivo de un viajero:
  Cuando la descripción indica explícitamente a quién pertenece (ej: "SEGURO SR. GARCIA"), o cuando el precio solo cuadra asignándolo a una persona.
  → Genera solo 1 fila para esa persona, no la repliques al resto.

REGLA DE ORO: La suma de todos los importe_unitario SIEMPRE debe coincidir con el Subtotal del documento. Úsala como verificación final. Si no cuadra, revisa tu interpretación.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 4 — TASAS DE EMISIÓN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Siempre generan una fila propia por viajero. Nunca las agrupes con otros servicios.
Si hay una tasa de 15,00 con TOTAL 30,00 y hay 2 viajeros → 2 filas de 15,00 cada una.
descripcion_abreviada: OTROS
trayecto: null

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 5 — HOTEL Y NOCHES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Calcula siempre el número de noches a partir de las fechas (del 06 al 07 = 1 noche, del 11 al 14 = 3 noches).
Formato descripcion_servicio: "[NOMBRE HOTEL] // [TIPO HAB] // X NOCHE" o "// X NOCHES" (singular si es 1).
Si el precio es por noche y hay varias noches, el importe_unitario es el precio total de la estancia para esa persona (precio/noche × noches).
Fecha: la fecha de CHECK-IN (inicio de la estancia).
Trayecto: la CIUDAD donde está el hotel.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 6 — TRANSPORTE (AVIÓN, TREN, COCHE, TAXI)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Ida y vuelta en líneas separadas del albarán = 2 filas separadas en la salida, con sus fechas y trayectos propios.
Ida y vuelta en una sola línea (ej: "VALENCIA - PARIS - VALENCIA") = 1 fila, trayecto tal cual aparece.
Trayecto: siempre ORIGEN - DESTINO en mayúsculas.
Las tasas de aeropuerto incluidas en el billete (ej: "Tasas 9,38") forman parte del importe del billete si vienen en la misma línea; son TASAS DE EMISIÓN solo si aparecen en su propia línea con ese nombre.
Fecha: la fecha de salida/inicio del trayecto.
descripcion_abreviada: AVION / TREN / COCHE / TAXI según corresponda.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 7 — OTROS SERVICIOS (SEGUROS, TRANSFERS, ETC.)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Cualquier servicio que no sea transporte ni alojamiento → descripcion_abreviada: OTROS.
Ejemplos: seguro de viaje, parking, visado, servicio VIP, sala de reuniones, etc.
Si el servicio es para un subconjunto de los viajeros del bloque, aplica solo a esos viajeros (Regla 3, Caso C).
Trayecto: la ciudad relevante si aplica, o null si no hay.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 8 — NÚMERO Y DESCRIPCIÓN DE PROYECTO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
numero_proyecto: el código alfanumérico que precede a la descripción del proyecto en la cabecera del bloque (ej: "24VNR02", "INDITEX001"). Déjalo null si no hay código diferenciado.
descripcion_proyecto: el nombre del proyecto/evento/marca (ej: "VINCE FIXTURES 2024", "INDITEX FABRIC FAIR", "LOUIS VUITTON SAINT TROPEZ").
IMPORTANTE: Los códigos tipo "KS XXXXX" que aparecen al PIE del documento son referencias internas del agente de viajes, NO son número de proyecto. Ignóralos para este campo.
Si la cabecera del bloque solo tiene descripción sin código (ej: "VH - PARIS"), numero_proyecto = null y descripcion_proyecto = "VH - PARIS".

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 9 — SUBTOTAL vs TOTAL
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
El campo subtotal SIEMPRE es el "Subtotal" del documento (antes de IVA), nunca el "TOTAL" (que incluye IVA).
Este valor es el mismo en todas las filas del mismo albarán.
Si el documento no tiene línea "Subtotal" separada, usa el importe antes de impuestos.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 10 — LIMPIEZA DE NOMBRES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Elimina siempre: SR., SRA., D., DA., DON, DOÑA, MR., MRS., MS., DR., DRA.
Mantén el nombre completo (nombre + apellidos) tal como aparece, en mayúsculas si así está en el documento.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 11 — FECHAS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Formato siempre DD-MM-AAAA. Ejemplos: 06-03-2025, 11-03-2025.
Si el año no aparece explícitamente, dedúcelo del contexto del albarán (fecha del albarán u otros servicios).
Si no hay fecha posible para un servicio (ej: tasa de emisión sin fecha), usa la fecha del albarán.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 12 — CAMPOS VACÍOS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Usa null (no string vacío "") para campos sin valor: numero_proyecto, trayecto cuando no aplica.
Nunca inventes datos. Si algo no está en el documento, es null.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 13 — NÚMERO DE ALBARÁN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Extrae el número exactamente como aparece en el documento (ej: "A202502058", "ALB-2058", "2058").
No lo normalices ni lo modifiques.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REGLA 14 — SERVICIOS CON IMPORTE CERO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Si un servicio tiene importe 0,00 (ej: tasa de emisión bonificada, servicio gratuito incluido), inclúyelo igualmente como fila con importe_unitario: 0.

════════════════════════════════════════
VALIDACIÓN FINAL OBLIGATORIA
════════════════════════════════════════
Antes de generar la salida, haz este cálculo mentalmente:
  SUMA de todos los importe_unitario del albarán = Subtotal del documento?
Si NO coincide → revisa qué servicio has mal interpretado y corrígelo. No respondas hasta que cuadre.

════════════════════════════════════════
CAMPOS A EXTRAER (claves JSON exactas)
════════════════════════════════════════
- n_albaran           → número de albarán tal como aparece
- numero_proyecto     → código alfanumérico del proyecto (null si no hay)
- descripcion_proyecto → nombre/descripción del proyecto o evento
- nombre_quien_viaja  → nombre completo sin títulos
- fecha               → DD-MM-AAAA (fecha de inicio del servicio)
- trayecto            → ORIGEN - DESTINO para transporte; CIUDAD para hotel/otros; null para tasas
- descripcion_servicio → descripción completa y literal (con regla de noches aplicada si es hotel)
- descripcion_abreviada → exactamente uno de: TREN / HOTEL / AVION / COCHE / TAXI / OTROS
- importe_unitario    → número decimal, coste para esa persona
- subtotal            → número decimal, subtotal total del albarán (igual en todas las filas del mismo albarán)

════════════════════════════════════════
FASE 2 — FORMATO DE SALIDA
════════════════════════════════════════
Devuelve ÚNICAMENTE JSON válido. Sin texto antes ni después. Sin explicaciones. Sin markdown. Sin bloques de código.
Formato: una lista [] donde cada elemento {} es una fila con las claves exactas definidas arriba.`;

// Procesa un único PDF y devuelve sus filas
async function procesarPDF(file) {
  let uploadedFile;
  try {
    uploadedFile = await openai.files.create({
      file: new File([file.buffer], file.originalname, { type: 'application/pdf' }),
      purpose: 'assistants',
    });

    const response = await openai.responses.create({
      model: 'gpt-4o',
      instructions: SYSTEM_PROMPT,
      input: [
        {
          role: 'user',
          content: [
            { type: 'input_file', file_id: uploadedFile.id },
            { type: 'input_text', text: 'Extrae todos los datos de este albarán. Devuelve únicamente el JSON, sin texto adicional ni bloques de código.' },
          ],
        },
      ],
    });

    let content = response.output_text.trim();
    content = content.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();

    return JSON.parse(content);

  } finally {
    if (uploadedFile) {
      try { await openai.files.del(uploadedFile.id); } catch (_) {}
    }
  }
}

// Ejecuta tareas en paralelo con un límite de concurrencia
async function procesarEnParalelo(tareas, concurrencia) {
  const resultados = new Array(tareas.length);
  let indice = 0;

  async function worker() {
    while (indice < tareas.length) {
      const i = indice++;
      resultados[i] = await tareas[i]();
    }
  }

  const workers = Array.from({ length: concurrencia }, worker);
  await Promise.all(workers);
  return resultados;
}

// Endpoint principal con SSE para progreso en tiempo real
app.post('/api/procesar', upload.array('pdfs', 100), async (req, res) => {
  const files = req.files;
  if (!files || files.length === 0) {
    return res.status(400).json({ error: 'No se han enviado archivos.' });
  }

  // Cabeceras SSE
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  const send = (data) => res.write(`data: ${JSON.stringify(data)}\n\n`);

  const total = files.length;
  let completados = 0;
  const todasLasFilas = [];
  const errores = [];

  send({ tipo: 'inicio', total });

  const tareas = files.map(file => async () => {
    try {
      const filas = await procesarPDF(file);
      todasLasFilas.push(...filas);
      completados++;
      send({ tipo: 'progreso', completados, total, archivo: file.originalname, ok: true });
    } catch (err) {
      completados++;
      errores.push(file.originalname);
      send({ tipo: 'progreso', completados, total, archivo: file.originalname, ok: false, error: err.message });
    }
  });

  // 5 en paralelo: equilibrio entre velocidad y límites de la API
  await procesarEnParalelo(tareas, 5);

  if (todasLasFilas.length === 0) {
    send({ tipo: 'error', mensaje: 'No se pudo extraer datos de ningún albarán.' });
    return res.end();
  }

  // Generar Excel
  try {
    const cabeceras = ['Nº Albarán', 'Numero Proyecto', 'Descripcion Proyecto', 'Nombre quien viaja', 'Fecha', 'Trayecto', 'Descripción Servicio', 'Descripción Abreviada', 'Importe unitario', 'Subtotal'];
    const claves = ['n_albaran', 'numero_proyecto', 'descripcion_proyecto', 'nombre_quien_viaja', 'fecha', 'trayecto', 'descripcion_servicio', 'descripcion_abreviada', 'importe_unitario', 'subtotal'];

    const filas = todasLasFilas.map(row => claves.map(k => row[k] ?? ''));
    const ws = XLSX.utils.aoa_to_sheet([cabeceras, ...filas]);
    ws['!cols'] = [{ wch: 14 }, { wch: 16 }, { wch: 24 }, { wch: 28 }, { wch: 14 }, { wch: 22 }, { wch: 50 }, { wch: 20 }, { wch: 16 }, { wch: 14 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Albaranes');

    const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const excelBase64 = excelBuffer.toString('base64');

    send({
      tipo: 'listo',
      filas: todasLasFilas.length,
      errores: errores.length,
      archivosConError: errores,
      excel: excelBase64,
      fecha: new Date().toISOString().slice(0, 10),
    });
  } catch (err) {
    send({ tipo: 'error', mensaje: 'Error al generar el Excel: ' + err.message });
  }

  res.end();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor KS Albaranes corriendo en puerto ${PORT}`));
