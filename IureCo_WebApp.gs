// ================================================================
// IURECO — Web App (Apps Script)
// Pegar en: Google Sheets → Extensiones → Apps Script
// Publicar como: Web App → Cualquier persona
// ================================================================

const SHEET_CONFIG = {
  SPREADSHEET_ID:  '',                        // Se toma del sheet activo
  FORM_RESPONSES:  'Respuestas de formulario 1',
  CONTROL_SHEET:   'Control',
  TICKET_PREFIX:   'IUR',
};

const REQUISITOS = [
  'Tres nombres tentativos de la sociedad',
  'Nombre tentativo empresa mercantil',
  'Copia DPI accionistas',
  'Profesión u oficio de cada accionista',
  'Copia DPI representante legal',
  'Distribución de acciones (%)',
  'Objeto de la sociedad',
  'Copia recibo agua o luz',
  'Nombre y NIT del Contador',
  'Régimen tributario',
  'Capital social a autorizar (Q)',
  'Capital suscrito y pagado (Q)',
  'Boleta de anticipo'
];

const PASOS_TRAMITE = [
  'Redacción de borrador de escritura',
  'Impresión y firma de escritura pública',
  'Preparación de documentación',
  'Subida a Registro Mercantil',
  'Seguimiento del proceso',
  'Descarga de patentes',
  'Entrega de documentación'
];

// ================================================================
// ROUTER PRINCIPAL
// ================================================================
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const action = params.action || (e.postData ? JSON.parse(e.postData.contents).action : '');
  const body   = e.postData ? JSON.parse(e.postData.contents) : params;

  let result;
  try {
    switch (action) {
      case 'getExpedientes':    result = getExpedientes();              break;
      case 'getExpediente':     result = getExpediente(body.ticket);    break;
      case 'updateEstatus':     result = updateEstatus(body);           break;
      case 'updatePaso':        result = updatePaso(body);              break;
      case 'updateNotas':       result = updateNotas(body);             break;
      case 'marcarFaltantes':   result = marcarFaltantes(body);         break;
      case 'generarFormCorr':   result = generarFormCorreccion(body);   break;
      case 'confirmarPago':     result = confirmarPago(body);           break;
      case 'agregarWhatsApp':   result = agregarWhatsApp(body);         break;
      default: result = { error: 'Acción no reconocida: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// FUNCIONES DE DATOS
// ================================================================

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function asegurarHojaControl() {
  const ss = getSSS();
  let ctrl = ss.getSheetByName(SHEET_CONFIG.CONTROL_SHEET);
  if (ctrl) return ctrl;

  ctrl = ss.insertSheet(SHEET_CONFIG.CONTROL_SHEET);
  const headers = [
    'Ticket', 'Cliente', 'WhatsApp', 'Fecha Ingreso', 'Estatus',
    'Paso Actual', 'Pago Anticipo', 'Pago Entrega', 'Notas',
    'Link Form Corrección'
  ].concat(REQUISITOS.map(r => 'FALTA: ' + r));

  ctrl.getRange(1, 1, 1, headers.length).setValues([headers]);
  ctrl.getRange(1, 1, 1, headers.length)
    .setBackground('#1A2744')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  ctrl.setFrozenRows(1);
  return ctrl;
}

function getSSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Obtener todos los expedientes para el panel
function getExpedientes() {
  const ss   = getSSS();
  const ctrl = ss.getSheetByName(SHEET_CONFIG.CONTROL_SHEET);
  if (!ctrl) return { expedientes: [] };

  const data    = ctrl.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  const expedientes = rows
    .filter(r => r[0] !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });

  return { expedientes };
}

// Obtener un expediente específico
function getExpediente(ticket) {
  const ss   = getSSS();
  const ctrl = ss.getSheetByName(SHEET_CONFIG.CONTROL_SHEET);
  const data = ctrl.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === ticket) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx]);
      return { expediente: obj, row: i + 1 };
    }
  }
  return { error: 'Ticket no encontrado' };
}

// Actualizar estatus
function updateEstatus(body) {
  const { ticket, estatus } = body;
  return updateField(ticket, 'Estatus', estatus);
}

// Actualizar paso actual del trámite
function updatePaso(body) {
  const { ticket, paso } = body;
  return updateField(ticket, 'Paso Actual', paso);
}

// Actualizar notas
function updateNotas(body) {
  const { ticket, notas } = body;
  return updateField(ticket, 'Notas', notas);
}

// Guardar número de WhatsApp del cliente
function agregarWhatsApp(body) {
  const { ticket, whatsapp } = body;
  return updateField(ticket, 'WhatsApp', whatsapp);
}

// Confirmar pago contra entrega
function confirmarPago(body) {
  const { ticket } = body;
  return updateField(ticket, 'Pago Entrega', new Date().toLocaleDateString('es-GT'));
}

// Marcar documentos faltantes
function marcarFaltantes(body) {
  const { ticket, faltantes } = body; // faltantes: array de nombres de requisitos
  const ss   = getSSS();
  const ctrl = ss.getSheetByName(SHEET_CONFIG.CONTROL_SHEET);
  const data = ctrl.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === ticket) {
      REQUISITOS.forEach(req => {
        const col = headers.indexOf('FALTA: ' + req);
        if (col >= 0) {
          ctrl.getRange(i + 1, col + 1).setValue(faltantes.includes(req));
        }
      });
      return { ok: true };
    }
  }
  return { error: 'Ticket no encontrado' };
}

// Generar formulario de corrección
function generarFormCorreccion(body) {
  const { ticket } = body;
  const ss   = getSSS();
  const ctrl = ss.getSheetByName(SHEET_CONFIG.CONTROL_SHEET);
  const data = ctrl.getDataRange().getValues();
  const headers = data[0];

  let filaCliente = null;
  let rowIndex    = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === ticket) {
      filaCliente = data[i];
      rowIndex    = i + 1;
      break;
    }
  }
  if (!filaCliente) return { error: 'Ticket no encontrado' };

  const faltantes = REQUISITOS.filter(req => {
    const col = headers.indexOf('FALTA: ' + req);
    return col >= 0 && filaCliente[col] === true;
  });

  if (faltantes.length === 0) return { error: 'No hay requisitos marcados como faltantes' };

  const clienteNombre = filaCliente[1];
  const form = FormApp.create(`Corrección — ${ticket}`);
  form.setTitle(`Corrección de documentos — ${ticket}`);
  form.setDescription(
    `Estimado/a ${clienteNombre},\n\n` +
    `Para continuar con su trámite necesitamos la siguiente documentación faltante.\n\n` +
    `Su número de expediente: ${ticket}`
  );
  form.setConfirmationMessage(`Gracias. Hemos recibido su corrección. Expediente: ${ticket}`);

  // Campo de ticket (referencia)
  form.addTextItem()
    .setTitle('Número de expediente')
    .setRequired(true);

  // Campos según tipo
  const TIPO = {
    'Copia DPI accionistas': 'archivo',
    'Copia DPI representante legal': 'archivo',
    'Copia recibo agua o luz': 'archivo',
    'Boleta de anticipo': 'archivo'
  };

  faltantes.forEach(req => {
    if (TIPO[req] === 'archivo') {
      form.addFileUploadItem().setTitle(req).setRequired(true);
    } else {
      form.addParagraphTextItem().setTitle(req).setRequired(true);
    }
  });

  const formUrl = form.getPublishedUrl();

  // Guardar link en Control
  const linkCol = headers.indexOf('Link Form Corrección');
  if (linkCol >= 0) ctrl.getRange(rowIndex, linkCol + 1).setValue(formUrl);
  updateField(ticket, 'Estatus', 'Corrección solicitada');

  return { ok: true, formUrl, faltantes, cliente: clienteNombre };
}

// ── Helper: actualizar un campo por nombre de columna ──────────
function updateField(ticket, campo, valor) {
  const ss   = getSSS();
  const ctrl = ss.getSheetByName(SHEET_CONFIG.CONTROL_SHEET);
  const data = ctrl.getDataRange().getValues();
  const headers = data[0];
  const col = headers.indexOf(campo);
  if (col < 0) return { error: 'Campo no encontrado: ' + campo };

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === ticket) {
      ctrl.getRange(i + 1, col + 1).setValue(valor);
      return { ok: true };
    }
  }
  return { error: 'Ticket no encontrado' };
}

// ================================================================
// TRIGGER: Al enviar el formulario
// ================================================================
function onFormSubmit(e) {
  const ss      = getSSS();
  const sheet   = ss.getSheetByName(SHEET_CONFIG.FORM_RESPONSES) || ss.getActiveSheet();
  const lastRow = sheet.getLastRow();

  const year     = new Date().getFullYear();
  const num      = String(lastRow - 1).padStart(3, '0');
  const ticketId = `${SHEET_CONFIG.TICKET_PREFIX}-${year}-${num}`;

  // Escribir ticket en la fila de respuestas
  const lastCol = sheet.getLastColumn();
  sheet.getRange(lastRow, lastCol + 1).setValue(ticketId).setBackground('#D9EAD3');

  // Asegurar hoja Control y agregar expediente
  const ctrl   = asegurarHojaControl();
  const values = e.values;
  const nombre = values[1] || 'Sin nombre';
  const fecha  = new Date().toLocaleDateString('es-GT');

  const fila = [ticketId, nombre, '', fecha, 'Pendiente revisión', '', 'No', 'No', '', ''];
  REQUISITOS.forEach(() => fila.push(false));

  ctrl.appendRow(fila);

  // Insertar checkboxes en columnas de requisitos
  const newRow  = ctrl.getLastRow();
  const colStart = 11; // columna donde inician los FALTA:
  ctrl.getRange(newRow, colStart, 1, REQUISITOS.length).insertCheckboxes();

  // Actualizar mensaje de confirmación del formulario con el ticket
  try {
    const form = FormApp.openByUrl(ss.getFormUrl());
    form.setConfirmationMessage(
      `✅ Documentación recibida.\n\nTu número de expediente es:\n${ticketId}\n\nAnótalo y envíalo al grupo de WhatsApp de IureCo para dar seguimiento.`
    );
  } catch(err) {
    Logger.log('No se pudo actualizar confirmación: ' + err);
  }
}

// ================================================================
// MENÚ
// ================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('IureCo')
    .addItem('Inicializar hoja Control', 'asegurarHojaControl')
    .addToUi();
}
