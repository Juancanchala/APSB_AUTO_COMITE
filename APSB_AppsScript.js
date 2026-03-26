// ============================================================
// APSB · SISTEMA DE SEGUIMIENTO DE CONTRATOS
// Apps Script — Google Sheets
// ============================================================
// INSTRUCCIONES DE CONFIGURACIÓN:
// 1. Abre tu Google Sheet
// 2. Extensiones → Apps Script
// 3. Pega este código completo
// 4. Reemplaza OPENAI_API_KEY con tu key real
// 5. Reemplaza SHEET_ID con el ID de tu Sheet (URL entre /d/ y /edit)
// 6. Guarda y ejecuta setupTriggers() UNA VEZ para activar los triggers
// 7. Publica como Web App: Implementar → Nueva implementación →
//    Tipo: App web → Ejecutar como: Yo → Acceso: Cualquier usuario
// ============================================================

// ══════════════════════════════════════════════
// CONFIGURACIÓN — EDITAR ESTOS VALORES
// ══════════════════════════════════════════════

var CONFIG = {
  OPENAI_API_KEY: 'TU_OPENAI_API_KEY_AQUI', // ← Reemplaza con tu key (no la subas a git)
  OPENAI_MODEL:   'gpt-4o-mini',
  SHEET_ID:       '1CUcvEz3uXK8BafnXAb1jpiXK1JO2WJ_otDjS9uJBpdI',    // ← ID DE TU SHEET
  HOJA_RAW:       'RAW',        // Nombre de la hoja donde pegas el Excel de Planner
  HOJA_PROCESADA: 'PROCESADA',  // Nombre de la hoja de salida estructurada
};

// ══════════════════════════════════════════════
// COLUMNAS HOJA PROCESADA
// ══════════════════════════════════════════════

var COLS_PROCESADA = [
  'semana',           // Fecha de la semana (ej: 2026-03-25)
  'fecha_exportacion',// Fecha de exportación del Excel Planner
  'id',               // ID único del proyecto (ej: manzanillo, um_sap)
  'nombre',           // Nombre del proyecto
  'seccion',          // regular | urgencia | estudios | transversal
  'corregimiento',
  'vereda',
  'contrato_obra',
  'contrato_interventoria',
  'avance',           // Número 0-100 o null
  'fecha_termino',
  'modificacion',
  'estado',           // ejecucion | ampliado | estructurando | en-revision | pendiente-cdp
  'objeto',
  'avances_semana',   // JSON array como string
  'pendientes',       // JSON array como string
  'alertas',          // JSON array como string
  'frentes',          // JSON array como string (solo urgencia)
  'compromisos',      // JSON array como string (solo transversal)
  'mod_items',        // JSON array como string
  'tipo_contrato',    // Solo estudios: Obra | Interventoría | Consultoría
  'estado_proceso',   // Solo estudios: estructurando | en-revision | pendiente-cdp
  'subtitulo',        // Solo transversal
  'entidades',        // JSON array como string (solo transversal)
  'raw_nombre',       // Nombre original de la tarea en Planner (para trazabilidad)
  'raw_deposito',     // Depósito original
  'raw_etiquetas',    // Etiquetas originales
];

// ══════════════════════════════════════════════
// PROMPT DEL AGENTE IA — ENTRENADO PARA APSB
// ══════════════════════════════════════════════

function buildPrompt(tareasJson, fechaExportacion) {
  return `Eres un asistente especializado en gestión de contratos de infraestructura de agua potable y saneamiento básico para la Alcaldía de Medellín, División Técnica APSB.

Tu tarea es procesar las tareas exportadas desde Microsoft Planner y convertirlas en objetos JSON estructurados para el sistema de seguimiento semanal.

FECHA DE EXPORTACIÓN: ${fechaExportacion}

REGLAS DE CLASIFICACIÓN POR DEPÓSITO:
- "En Ejecución de obra" + etiqueta "Urgencia Manifiesta" → seccion: "urgencia"
- "En Ejecución de obra" (sin Urgencia Manifiesta) → seccion: "regular"
- "Por contratar obra" o "Proyectos priorizados" → seccion: "estudios"
- "Para liquidar" o "Suspendidos" → OMITIR (no incluir en el resultado)
- Tareas con etiqueta "Automatización" → seccion: "regular" (proyectos en ejecución)

REGLAS DE EXTRACCIÓN:
1. El campo "Nombre de la tarea" usualmente empieza con el número de contrato (ej: 4600105619). Extráelo como contrato_obra.
2. El campo "Descripción" contiene el texto libre de avances semanales — es la fuente principal.
3. Las "Notas" en descripción con fechas (ej: "13/03/2026 Se realizó...") son avances de la semana actual.
4. Palabras clave para clasificar dentro de descripción:
   - "Pendiente", "falta", "requiere", "se debe", "está por" → pendientes[]
   - "Alerta", "riesgo", "sin avance", "suspendido", "vencido" → alertas[]
   - Frentes de obra: buscar nombres de sectores geográficos seguidos de actividades (Acueducto X, PTAR Y, Muro Z)
   - Compromisos: fechas futuras con verbo en futuro o "programado para"
5. Para el avance (%): buscar porcentajes explícitos en la descripción. Si no hay, usar null.
6. Para estado: 
   - "En curso" + depósito "En Ejecución" → "ejecucion" o "ampliado" (si menciona ampliación)
   - "No iniciado" + "Por contratar" → "estructurando" o "en-revision" o "pendiente-cdp"
7. El id debe ser en minúsculas, sin espacios, tipo slug: ej "manzanillo", "um_sap", "ep_llano"

MAPEO DE PROYECTOS CONOCIDOS (usa estos IDs exactos si reconoces el proyecto):
- 4600105619 Manzanillo → id: "manzanillo", corregimiento: "Altavista"
- 4600105522 Piedras Blancas → id: "piedrasblancas", corregimiento: "Santa Elena"  
- 4600106239 Pozos Sépticos → id: "pozosSepticos", corregimiento: "San Sebastián de Palmitas"
- 4600104896 SAP Urgencia → id: "um_sap", corregimiento: "San Antonio de Prado"
- 4600104897 Altavista Urgencia → id: "um_altavista", corregimiento: "Altavista"
- 4600104894 Carrotanque → id: "um_carrotanque", corregimiento: "San Antonio de Prado / Altavista"
- 4600104895 Interventoría UM → id: "um_interventoria", corregimiento: "San Antonio de Prado / Altavista"
- El Llano alcantarillado → id: "ep_llano", corregimiento: "Santa Elena"
- La Aldea obras complementarias → id: "ep_aldea", corregimiento: "San Sebastián de Palmitas"
- CTI sistema no convencional → id: "ep_cti", corregimiento: "San Sebastián de Palmitas"
- 4600105388 Automatización → id: "automatizacion_ptap", seccion: "regular"

FORMATO DE RESPUESTA:
Responde ÚNICAMENTE con un array JSON válido, sin texto adicional, sin bloques de código, sin explicaciones.
Cada elemento del array es un objeto con EXACTAMENTE estos campos:

{
  "id": string,
  "nombre": string,
  "seccion": "regular" | "urgencia" | "estudios" | "transversal",
  "corregimiento": string,
  "vereda": string,
  "contrato_obra": string,
  "contrato_interventoria": string | null,
  "avance": number | null,
  "fecha_termino": string,
  "modificacion": string,
  "estado": string,
  "objeto": string,
  "avances_semana": string[],
  "pendientes": string[],
  "alertas": string[],
  "frentes": [{"nombre": string, "actividades": string[]}],
  "compromisos": [{"fecha": string, "texto": string, "responsable": string}],
  "mod_items": [{"titulo": string, "valor": string, "detalle": string}],
  "tipo_contrato": string,
  "estado_proceso": string,
  "subtitulo": string,
  "entidades": string[],
  "raw_nombre": string,
  "raw_deposito": string,
  "raw_etiquetas": string
}

TAREAS A PROCESAR:
${tareasJson}`;
}

// ══════════════════════════════════════════════
// TRIGGER PRINCIPAL — onEdit
// ══════════════════════════════════════════════

function onEditTrigger(e) {
  var sheet = e.source.getActiveSheet();
  // Solo se activa cuando editas la hoja RAW
  if (sheet.getName() !== CONFIG.HOJA_RAW) return;
  // Solo si la edición fue en las primeras columnas (pegaste datos)
  if (e.range.getRow() > 1) {
    procesarRAW();
  }
}

// ══════════════════════════════════════════════
// CONFIGURAR TRIGGERS (ejecutar UNA VEZ)
// ══════════════════════════════════════════════

function setupTriggers() {
  // Eliminar triggers existentes para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  // Trigger onEdit para la hoja RAW
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log('✓ Triggers configurados correctamente');
}

// ══════════════════════════════════════════════
// LEER HOJA RAW Y PREPARAR DATOS
// ══════════════════════════════════════════════

function leerRAW() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var hoja  = ss.getSheetByName(CONFIG.HOJA_RAW);
  var datos = hoja.getDataRange().getValues();

  if (datos.length < 2) {
    Logger.log('Hoja RAW vacía o sin datos');
    return null;
  }

  var headers = datos[0];
  var tareas  = [];

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var tarea = {};
    headers.forEach(function(h, j) {
      tarea[h] = fila[j] || '';
    });
    // Solo incluir tareas con nombre
    if (tarea['Nombre de la tarea']) {
      tareas.push(tarea);
    }
  }

  // Obtener fecha de exportación del plan
  var hojaInfo = ss.getSheetByName('Nombre del plan');
  var fechaExportacion = '';
  if (hojaInfo) {
    var infoData = hojaInfo.getDataRange().getValues();
    infoData.forEach(function(row) {
      if (String(row[0]).indexOf('Fecha de exportación') !== -1) {
        fechaExportacion = String(row[1]);
      }
    });
  }

  return { tareas: tareas, fechaExportacion: fechaExportacion };
}

// ══════════════════════════════════════════════
// LLAMAR A OPENAI
// ══════════════════════════════════════════════

function llamarOpenAI(prompt) {
  var url     = 'https://api.openai.com/v1/chat/completions';
  var payload = {
    model: CONFIG.OPENAI_MODEL,
    messages: [
      {
        role: 'system',
        content: 'Eres un extractor de datos estructurados. Respondes ÚNICAMENTE con JSON válido, sin texto adicional ni bloques de código markdown.'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.1,
    max_tokens: 8000
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code     = response.getResponseCode();
  var body     = response.getContentText();

  if (code !== 200) {
    Logger.log('Error OpenAI: ' + code + ' — ' + body);
    throw new Error('OpenAI error ' + code);
  }

  var json    = JSON.parse(body);
  var content = json.choices[0].message.content.trim();

  // Limpiar por si acaso viene con bloques markdown
  content = content.replace(/^```json\s*/i, '').replace(/\s*```$/i, '').trim();

  return JSON.parse(content);
}

// ══════════════════════════════════════════════
// ESCRIBIR EN HOJA PROCESADA
// ══════════════════════════════════════════════

function escribirProcesada(proyectos, fechaExportacion) {
  var ss   = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var hoja = ss.getSheetByName(CONFIG.HOJA_PROCESADA);

  // Crear hoja si no existe
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA_PROCESADA);
  }

  // Escribir encabezados si la hoja está vacía
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(COLS_PROCESADA);
    // Estilo de encabezados
    var headerRange = hoja.getRange(1, 1, 1, COLS_PROCESADA.length);
    headerRange.setBackground('#231F20');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
  }

  // Fecha de semana actual (lunes de esta semana)
  var hoy   = new Date();
  var dia   = hoy.getDay();
  var lunes = new Date(hoy);
  lunes.setDate(hoy.getDate() - (dia === 0 ? 6 : dia - 1));
  var semana = Utilities.formatDate(lunes, 'America/Bogota', 'yyyy-MM-dd');

  // Agregar una fila por proyecto
  proyectos.forEach(function(p) {
    var fila = COLS_PROCESADA.map(function(col) {
      var val = p[col];
      if (col === 'semana')            return semana;
      if (col === 'fecha_exportacion') return fechaExportacion;
      // Arrays → JSON string
      if (Array.isArray(val)) return JSON.stringify(val);
      // Null → string vacío
      if (val === null || val === undefined) return '';
      return val;
    });
    hoja.appendRow(fila);
  });

  Logger.log('✓ Escritos ' + proyectos.length + ' proyectos en hoja PROCESADA para semana ' + semana);
}

// ══════════════════════════════════════════════
// PROCESO PRINCIPAL
// ══════════════════════════════════════════════

function procesarRAW() {
  try {
    Logger.log('Iniciando procesamiento RAW...');

    // 1. Leer datos de la hoja RAW
    var data = leerRAW();
    if (!data) return;

    Logger.log('Tareas encontradas: ' + data.tareas.length);

    // 2. Filtrar solo tareas relevantes (excluir Para liquidar y Suspendidos)
    var tareasRelevantes = data.tareas.filter(function(t) {
      var dep = String(t['Nombre del depósito'] || '');
      return dep !== 'Para liquidar' && dep !== 'Suspendidos';
    });

    Logger.log('Tareas relevantes (sin Para liquidar / Suspendidos): ' + tareasRelevantes.length);

    // 3. Simplificar para enviar a OpenAI (solo campos necesarios)
    var tareasSimplificadas = tareasRelevantes.map(function(t) {
      return {
        nombre:      t['Nombre de la tarea'],
        deposito:    t['Nombre del depósito'],
        progreso:    t['Progreso'],
        asignado_a:  t['Asignado a'],
        fecha_inicio:t['Fecha de inicio'],
        fecha_fin:   t['Fecha de vencimiento'],
        etiquetas:   t['Etiquetas'],
        descripcion: t['Descripción']
      };
    });

    // 4. Construir prompt y llamar a OpenAI
    var prompt    = buildPrompt(JSON.stringify(tareasSimplificadas, null, 2), data.fechaExportacion);
    Logger.log('Llamando a OpenAI...');
    var proyectos = llamarOpenAI(prompt);

    Logger.log('Proyectos procesados por IA: ' + proyectos.length);

    // 5. Escribir en hoja PROCESADA
    escribirProcesada(proyectos, data.fechaExportacion);

    Logger.log('✓ Proceso completado exitosamente');

  } catch(e) {
    Logger.log('ERROR en procesarRAW: ' + e.toString());
    // Notificar por email en caso de error
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: 'APSB Seguimiento — Error en procesamiento',
      body: 'Error al procesar la hoja RAW:\n\n' + e.toString()
    });
  }
}

// ══════════════════════════════════════════════
// WEB APP — GET (para el HTML)
// ══════════════════════════════════════════════

function doGet(e) {
  var params = e ? e.parameter : {};
  var mode   = params.mode || 'current';   // current | history | all
  var semana = params.semana || '';         // filtro por semana específica
  var seccion= params.seccion || '';        // filtro por sección

  try {
    var resultado = {};

    if (mode === 'current') {
      // Última semana disponible
      resultado = getUltimaSemana(seccion);
    } else if (mode === 'history') {
      // Todo el histórico (para el chat IA)
      resultado = getTodoHistorico(seccion);
    } else if (mode === 'semanas') {
      // Lista de semanas disponibles
      resultado = getListaSemanas();
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: resultado }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Obtener semana más reciente ──
function getUltimaSemana(seccionFiltro) {
  var ss   = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var hoja = ss.getSheetByName(CONFIG.HOJA_PROCESADA);
  if (!hoja || hoja.getLastRow() < 2) return buildDataVacia();

  var datos   = hoja.getDataRange().getValues();
  var headers = datos[0];

  // Encontrar la semana más reciente
  var colSemana = headers.indexOf('semana');
  var semanas   = [];
  for (var i = 1; i < datos.length; i++) {
    var s = datos[i][colSemana];
    if (s && semanas.indexOf(s) === -1) semanas.push(s);
  }
  semanas.sort().reverse();
  var ultimaSemana = semanas[0];

  // Filtrar filas de la última semana
  var filas = datos.filter(function(row, idx) {
    if (idx === 0) return false;
    if (row[colSemana] !== ultimaSemana) return false;
    if (seccionFiltro && row[headers.indexOf('seccion')] !== seccionFiltro) return false;
    return true;
  });

  return construirDATA(headers, filas);
}

// ── Obtener todo el histórico ──
function getTodoHistorico(seccionFiltro) {
  var ss   = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var hoja = ss.getSheetByName(CONFIG.HOJA_PROCESADA);
  if (!hoja || hoja.getLastRow() < 2) return [];

  var datos   = hoja.getDataRange().getValues();
  var headers = datos[0];

  var filas = datos.filter(function(row, idx) {
    if (idx === 0) return false;
    if (seccionFiltro && row[headers.indexOf('seccion')] !== seccionFiltro) return false;
    return true;
  });

  // Para histórico devolvemos array plano con semana incluida
  return filas.map(function(row) {
    var obj = {};
    headers.forEach(function(h, j) {
      var val = row[j];
      // Intentar parsear JSON strings
      if (typeof val === 'string' && (val.startsWith('[') || val.startsWith('{'))) {
        try { val = JSON.parse(val); } catch(e) {}
      }
      obj[h] = val;
    });
    return obj;
  });
}

// ── Lista de semanas disponibles ──
function getListaSemanas() {
  var ss   = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var hoja = ss.getSheetByName(CONFIG.HOJA_PROCESADA);
  if (!hoja || hoja.getLastRow() < 2) return [];

  var datos     = hoja.getDataRange().getValues();
  var headers   = datos[0];
  var colSemana = headers.indexOf('semana');
  var semanas   = [];

  for (var i = 1; i < datos.length; i++) {
    var s = datos[i][colSemana];
    if (s && semanas.indexOf(s) === -1) semanas.push(s);
  }
  return semanas.sort().reverse();
}

// ── Construir objeto DATA compatible con el HTML ──
function construirDATA(headers, filas) {
  var data = { regular: [], urgencia: [], estudios: [], transversal: [] };

  filas.forEach(function(row) {
    var obj = {};
    headers.forEach(function(h, j) {
      var val = row[j];
      if (typeof val === 'string' && (val.startsWith('[') || val.startsWith('{'))) {
        try { val = JSON.parse(val); } catch(e) {}
      }
      if (val === '') val = null;
      obj[h] = val;
    });

    var sec = obj.seccion;
    if (data[sec]) {
      data[sec].push(obj);
    }
  });

  return data;
}

// ── DATA vacía por defecto ──
function buildDataVacia() {
  return { regular: [], urgencia: [], estudios: [], transversal: [] };
}

// ══════════════════════════════════════════════
// UTILIDAD — Ejecutar manualmente para pruebas
// ══════════════════════════════════════════════

function testProcesamiento() {
  Logger.log('=== TEST DE PROCESAMIENTO ===');
  procesarRAW();
}

function testWebApp() {
  var resultado = doGet({ parameter: { mode: 'current' } });
  Logger.log(resultado.getContent());
}

function crearHojasIniciales() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  // Crear hoja RAW si no existe
  if (!ss.getSheetByName(CONFIG.HOJA_RAW)) {
    var raw = ss.insertSheet(CONFIG.HOJA_RAW);
    // Encabezados del Excel de Planner
    raw.appendRow([
      'Id. de tarea ', 'Nombre de la tarea', 'Nombre del depósito', 'Progreso',
      'Priority', 'Asignado a', 'Creado por', 'Fecha de creación', 'Fecha de inicio',
      'Fecha de vencimiento', 'Es periódica', 'Con retraso', 'Fecha de finalización',
      'Completado por', 'Elementos de la lista de comprobación completados',
      'Elementos de la lista de comprobación', 'Etiquetas', 'Descripción'
    ]);
    Logger.log('✓ Hoja RAW creada');
  }

  // Crear hoja PROCESADA si no existe
  if (!ss.getSheetByName(CONFIG.HOJA_PROCESADA)) {
    var proc = ss.insertSheet(CONFIG.HOJA_PROCESADA);
    proc.appendRow(COLS_PROCESADA);
    var h = proc.getRange(1, 1, 1, COLS_PROCESADA.length);
    h.setBackground('#231F20');
    h.setFontColor('#FFFFFF');
    h.setFontWeight('bold');
    Logger.log('✓ Hoja PROCESADA creada');
  }

  Logger.log('✓ Hojas iniciales listas');
}
