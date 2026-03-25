const express = require('express');
const multer = require('multer');
const OpenAI = require('openai');
const path = require('path');
const pdfParse = require('pdf-parse');
const XLSX = require('xlsx');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

app.use(express.static(path.join(__dirname, 'public')));

const SYSTEM_PROMPT = `Eres un asistente especializado en extracción de datos de albaranes de agencias de viajes.

Tu tarea se divide en dos fases:

FASE 1 — EXTRACCIÓN (aplica estas reglas estrictamente):

Regla Maestra de Ámbito de Viajero: Si se listan varios viajeros bajo un mismo concepto o proyecto, todos los servicios detallados aplican a todos esos viajeros. El coste de cada servicio compartido se divide a partes iguales entre el número de viajeros, SALVO que se indique explícitamente a quién corresponde.

Una Fila por Persona y Servicio: Cada línea de salida = un servicio para una persona.

Tasas de Emisión: Siempre van en su propia línea separada, nunca agrupadas con otros conceptos.

Noches de Hotel: Calcula el número de noches (ej: del 11 al 14 = 3 noches) y usa el formato "(descripción) // X NOCHES".

Limpieza de Nombres: Extrae el nombre completo omitiendo títulos como SR., SRA., D., etc.

Validación: La suma de todos los "importe_unitario" debe coincidir EXACTAMENTE con el subtotal del documento. Si no cuadra, revisa y corrige.

CAMPOS A EXTRAER:
- n_albaran: número de albarán
- numero_proyecto: código del proyecto (solo si es numérico/alfanumérico; si solo hay descripción, dejar vacío)
- descripcion_proyecto: nombre de marca o descripción del proyecto
- nombre_quien_viaja: nombre del pasajero (sin títulos)
- fecha: fecha de inicio del servicio (formato DD-MM-AAAA)
- trayecto: ORIGEN - DESTINO para transporte, o CIUDAD para otros servicios
- descripcion_servicio: descripción completa y literal del servicio (aplicando regla de noches si aplica)
- descripcion_abreviada: categoría del servicio — solo una de: TREN, HOTEL, AVION, COCHE, TAXI, OTROS
- importe_unitario: coste del servicio para esa persona (ya dividido si era compartido)
- subtotal: subtotal total del albarán (igual para todas las filas del mismo documento)

FASE 2 — FORMATO DE SALIDA:
Devuelve ÚNICAMENTE un bloque JSON válido (sin texto adicional, sin explicaciones, sin markdown):
Una lista [] donde cada elemento {} representa una fila con las claves exactas indicadas arriba.`;

app.post('/api/procesar', upload.array('pdfs', 20), async (req, res) => {
  try {
    const files = req.files;
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No se han enviado archivos.' });
    }

    const todasLasFilas = [];

    for (const file of files) {
      let textoPDF;
      try {
        const parsed = await pdfParse(file.buffer);
        textoPDF = parsed.text;
      } catch (e) {
        throw new Error(`No se pudo leer el PDF "${file.originalname}". Asegúrate de que no está protegido con contraseña.`);
      }

      if (!textoPDF || textoPDF.trim().length < 20) {
        throw new Error(`El PDF "${file.originalname}" parece estar vacío o ser una imagen escaneada sin texto.`);
      }

      const response = await openai.chat.completions.create({
        model: 'gpt-4o',
        messages: [
          { role: 'system', content: SYSTEM_PROMPT },
          {
            role: 'user',
            content: `Extrae todos los datos de este albarán siguiendo las instrucciones del sistema. Devuelve únicamente el JSON.\n\nCONTENIDO DEL ALBARÁN:\n\n${textoPDF}`
          }
        ],
        max_tokens: 4096,
        temperature: 0
      });

      let content = response.choices[0].message.content.trim();
      content = content.replace(/^```json\s*/i, '').replace(/\s*```$/i, '').trim();

      let filas;
      try {
        filas = JSON.parse(content);
      } catch (e) {
        console.error('Error parseando JSON para archivo:', file.originalname, content);
        throw new Error(`No se pudo interpretar la respuesta del albarán "${file.originalname}". Intenta de nuevo.`);
      }

      todasLasFilas.push(...filas);
    }

    // Generar Excel en el servidor
    const cabeceras = ['Nº Albarán', 'Numero Proyecto', 'Descripcion Proyecto', 'Nombre quien viaja', 'Fecha', 'Trayecto', 'Descripción Servicio', 'Descripción Abreviada', 'Importe unitario', 'Subtotal'];
    const claves = ['n_albaran', 'numero_proyecto', 'descripcion_proyecto', 'nombre_quien_viaja', 'fecha', 'trayecto', 'descripcion_servicio', 'descripcion_abreviada', 'importe_unitario', 'subtotal'];

    const filas = todasLasFilas.map(row => claves.map(k => row[k] ?? ''));

    const ws = XLSX.utils.aoa_to_sheet([cabeceras, ...filas]);
    ws['!cols'] = [{ wch: 14 }, { wch: 16 }, { wch: 24 }, { wch: 28 }, { wch: 14 }, { wch: 22 }, { wch: 50 }, { wch: 20 }, { wch: 16 }, { wch: 14 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Albaranes');

    const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    const fecha = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="albaranes_${fecha}.xlsx"`);
    res.send(excelBuffer);

  } catch (err) {
    console.error('Error en /api/procesar:', err.message);
    res.status(500).json({ error: err.message || 'Error interno del servidor.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor KS Albaranes corriendo en puerto ${PORT}`));
