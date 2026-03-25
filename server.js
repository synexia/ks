const express = require('express');
const multer = require('multer');
const OpenAI = require('openai');
const path = require('path');
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

Tasas de Emisión: Siempre van en su propia línea separada por viajero, nunca agrupadas con otros conceptos.

Noches de Hotel: Calcula el número de noches (ej: del 11 al 14 = 3 noches) y usa el formato "(descripción) // X NOCHES".

Limpieza de Nombres: Extrae el nombre completo omitiendo títulos como SR., SRA., D., etc.

Validación: La suma de todos los "importe_unitario" debe coincidir EXACTAMENTE con el subtotal del documento. Si no cuadra, revisa y corrige antes de responder.

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
Devuelve ÚNICAMENTE JSON válido (sin texto adicional, sin explicaciones, sin markdown, sin bloques de código):
Una lista [] donde cada elemento {} representa una fila con las claves exactas indicadas arriba.`;

app.post('/api/procesar', upload.array('pdfs', 20), async (req, res) => {
  try {
    const files = req.files;
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No se han enviado archivos.' });
    }

    const todasLasFilas = [];

    for (const file of files) {
      // Subir el PDF a OpenAI Files API
      let uploadedFile;
      try {
        const blob = new Blob([file.buffer], { type: 'application/pdf' });
        const formData = new FormData();
        formData.append('file', blob, file.originalname);
        formData.append('purpose', 'assistants');

        uploadedFile = await openai.files.create({
          file: new File([file.buffer], file.originalname, { type: 'application/pdf' }),
          purpose: 'assistants',
        });
      } catch (e) {
        throw new Error(`No se pudo subir el PDF "${file.originalname}" a OpenAI: ${e.message}`);
      }

      let content;
      try {
        const response = await openai.responses.create({
          model: 'gpt-4o',
          instructions: SYSTEM_PROMPT,
          input: [
            {
              role: 'user',
              content: [
                {
                  type: 'input_file',
                  file_id: uploadedFile.id,
                },
                {
                  type: 'input_text',
                  text: 'Extrae todos los datos de este albarán siguiendo las instrucciones. Devuelve únicamente el JSON, sin texto adicional ni bloques de código.',
                },
              ],
            },
          ],
        });

        content = response.output_text.trim();
      } finally {
        // Eliminar el archivo de OpenAI tras procesarlo
        try { await openai.files.del(uploadedFile.id); } catch (_) {}
      }

      // Limpiar posibles bloques de código que el modelo añada
      content = content.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();

      let filas;
      try {
        filas = JSON.parse(content);
      } catch (e) {
        console.error('Error parseando JSON para:', file.originalname, '\nRespuesta:', content);
        throw new Error(`No se pudo interpretar la respuesta del albarán "${file.originalname}". Intenta de nuevo.`);
      }

      todasLasFilas.push(...filas);
    }

    // Generar Excel
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
