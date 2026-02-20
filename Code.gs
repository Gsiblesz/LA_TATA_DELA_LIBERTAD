const CONFIG = {
  spreadsheetId: '18WPHKhmnGtoNiHuALuK8486VuJeMq8LHF0tKZArq3hs',
  mainSheetName: 'DATA',
  catalogSheetName: 'PRODUCTOS',
  timeZone: Session.getScriptTimeZone() || 'America/Caracas',
  columns: {
    hora: 1,
    fecha: 2,
    familia: 3,
    codigo: 4,
    unidad: 5,
    producto: 6,
    sede: 7,
    cantidadSolicitada: 8,
    responsableSolicitud: 9,
    cantidadEntregada: 10,
    responsableEntrega: 11,
    merma: 12,
    mes: 13,
    timestamp: 14,
  },
};

function doGet(e) {
  const action = String(e?.parameter?.action || '').toLowerCase();
  try {
    if (action === 'getproducts') {
      const products = getProducts_();
      return buildResponse_(true, { products }, 'Catálogo sincronizado.');
    }

    if (!action || action === 'ping') {
      return buildResponse_(true, { ok: true }, 'Servicio disponible.');
    }

    return buildResponse_(false, null, 'Acción GET no soportada.');
  } catch (error) {
    return buildResponse_(false, null, error.message);
  }
}

function doPost(e) {
  try {
    const { action, payload } = parseBody_(e);
    const normalizedAction = String(action || '').toLowerCase();
    const result = withLock_(() => handleAction_(normalizedAction, payload || {}));
    return buildResponse_(true, result.data, result.message);
  } catch (error) {
    return buildResponse_(false, null, error.message);
  }
}

function handleAction_(action, payload) {
  switch (action) {
    case 'createsolicitud': {
      const data = createSolicitud_(payload);
      return { data, message: 'Solicitudes de Sedes registradas.' };
    }
    case 'recordentrega': {
      const data = recordEntrega_(payload);
      return { data, message: `Entregado a Sedes procesado: ${data.processed}` };
    }
    case 'recordmerma': {
      const data = recordMerma_(payload);
      return { data, message: `Producción registrada: ${data.processed}` };
    }
    default:
      throw new Error('Acción POST no soportada.');
  }
}

function createSolicitud_(payload) {
  validateRequired_(payload, ['hora', 'fecha', 'sede', 'responsable']);
  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) throw new Error('Debes enviar al menos un producto.');

  const registroAutomatico = new Date();

  const sanitizedItems = sanitizeSolicitudItems_(items);
  const sheet = getMainSheet_();
  if (isRecentSolicitudDuplicate_(sheet, payload, sanitizedItems)) {
    throw new Error('Esta solicitud ya fue registrada recientemente. Verifica antes de reenviar.');
  }
  const rows = sanitizedItems.map((item) => [
    payload.hora,
    payload.fecha,
    '',
    item.code,
    item.unit,
    item.description,
    payload.sede,
    item.quantity,
    payload.responsable,
    '',
    '',
    '',
    '',
    registroAutomatico,
  ]);

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { rowsInserted: rows.length };
}

function sanitizeSolicitudItems_(items) {
  return items.map((item, index) => {
    const code = String(item.code || '').trim();
    const description = String(item.description || '').trim();
    const unit = String(item.unit || '').trim();
    const quantity = Number(item.quantity);
    const label = code || description || `#${index + 1}`;

    if (!code) {
      throw new Error(`El producto ${label} necesita un código.`);
    }

    if (!description) {
      throw new Error(`El producto ${label} necesita una descripción.`);
    }

    if (!unit) {
      throw new Error(`La unidad del producto ${label} es obligatoria.`);
    }

    if (!Number.isFinite(quantity) || quantity <= 0) {
      throw new Error(`La cantidad solicitada de ${label} debe ser mayor a cero.`);
    }

    return { code, description, unit, quantity };
  });
}

function recordEntrega_(payload) {
  validateRequired_(payload, ['fecha', 'sede', 'responsableEntrega']);
  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) {
    throw new Error('Debes enviar al menos un producto.');
  }

  const sanitizedItems = sanitizeEntregaItems_(items);
  const sheet = getMainSheet_();
  const summary = { processed: 0, updated: 0, appended: 0 };
  const registroAutomatico = new Date();

  if (isRecentEntregaDuplicate_(sheet, payload, sanitizedItems)) {
    throw new Error('Esta entrega ya fue registrada recientemente. Verifica antes de reenviar.');
  }

  sanitizedItems.forEach((item) => {
    if (payload.sinSolicitud) {
      appendEntregaSinSolicitud_(sheet, payload, item, item.cantidadEntregada, registroAutomatico);
    } else {
      appendEntregaDirecta_(sheet, payload, item, item.cantidadEntregada, registroAutomatico);
    }

    summary.appended += 1;
    summary.processed += 1;
  });

  return summary;
}

function sanitizeEntregaItems_(items) {
  return items.map((item, index) => {
    const productCode = String(item.productCode || '').trim();
    const productName = String(item.productName || '').trim();
    const unit = String(item.unit || '').trim();
    const cantidadEntregada = Number(item.cantidadEntregada);
    const label = productCode || productName || `#${index + 1}`;

    if (!productCode) {
      throw new Error(`El producto ${label} necesita un código.`);
    }
    if (!productName) {
      throw new Error(`El producto ${label} necesita una descripción.`);
    }
    if (!unit) {
      throw new Error(`La unidad del producto ${label} es obligatoria.`);
    }
    if (!Number.isFinite(cantidadEntregada) || cantidadEntregada <= 0) {
      throw new Error(`La cantidad entregada debe ser mayor a cero (${label}).`);
    }

    return { productCode, productName, unit, cantidadEntregada };
  });
}

function appendEntregaDirecta_(sheet, payload, item, qty, registroAutomatico) {
  const row = [
    payload.hora || '',
    payload.fecha || '',
    '',
    item.productCode || '',
    item.unit || '',
    item.productName || '',
    payload.sede || '',
    '',
    '',
    qty,
    payload.responsableEntrega || '',
    '',
    '',
    registroAutomatico || new Date(),
  ];
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function appendEntregaSinSolicitud_(sheet, payload, item, qty, registroAutomatico) {
  const row = [
    payload.hora || '',
    payload.fecha || '',
    '',
    item.productCode || '',
    item.unit || '',
    item.productName || '',
    payload.sede || '',
    '',
    'SIN SOLICITUD',
    qty,
    payload.responsableEntrega || '',
    '',
    '',
    registroAutomatico || new Date(),
  ];
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function isRecentSolicitudDuplicate_(sheet, payload, items) {
  if (!sheet || !items.length) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || (lastRow - 1) < items.length) return false;

  const rows = getTailRows_(sheet, items.length);
  if (!rows.length) return false;

  const expectedDate = normalizeDate_(payload.fecha);
  const expectedHora = normalizeText_(payload.hora);
  const expectedSede = normalizeText_(payload.sede);
  const expectedResponsable = normalizeText_(payload.responsable);

  const headersMatch = rows.every((row) => {
    const rowDate = normalizeDate_(row[CONFIG.columns.fecha - 1]);
    const rowHora = normalizeText_(row[CONFIG.columns.hora - 1]);
    const rowSede = normalizeText_(row[CONFIG.columns.sede - 1]);
    const rowResponsable = normalizeText_(row[CONFIG.columns.responsableSolicitud - 1]);
    const qtySolicitada = Number(row[CONFIG.columns.cantidadSolicitada - 1]) || 0;
    const qtyEntregada = Number(row[CONFIG.columns.cantidadEntregada - 1]) || 0;
    return (
      rowDate === expectedDate &&
      rowHora === expectedHora &&
      rowSede === expectedSede &&
      rowResponsable === expectedResponsable &&
      qtySolicitada > 0 &&
      qtyEntregada <= 0
    );
  });

  if (!headersMatch) return false;

  const incomingSignature = items
    .map((item) => `${normalizeText_(item.code)}|${normalizeText_(item.unit)}|${Number(item.quantity)}`)
    .sort()
    .join('||');

  const existingSignature = rows
    .map((row) => {
      const code = row[CONFIG.columns.codigo - 1];
      const unit = row[CONFIG.columns.unidad - 1];
      const qty = Number(row[CONFIG.columns.cantidadSolicitada - 1]) || 0;
      return `${normalizeText_(code)}|${normalizeText_(unit)}|${qty}`;
    })
    .sort()
    .join('||');

  return incomingSignature === existingSignature;
}

function isRecentEntregaDuplicate_(sheet, payload, items) {
  if (!sheet || !items.length) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || (lastRow - 1) < items.length) return false;

  const rows = getTailRows_(sheet, items.length);
  if (!rows.length) return false;

  const expectedDate = normalizeDate_(payload.fecha);
  const expectedHora = normalizeText_(payload.hora);
  const expectedSede = normalizeText_(payload.sede);
  const expectedResponsable = normalizeText_(payload.responsableEntrega);
  const expectedSolicitudFlag = payload.sinSolicitud ? 'SIN SOLICITUD' : '';

  const headersMatch = rows.every((row) => {
    const rowDate = normalizeDate_(row[CONFIG.columns.fecha - 1]);
    const rowHora = normalizeText_(row[CONFIG.columns.hora - 1]);
    const rowSede = normalizeText_(row[CONFIG.columns.sede - 1]);
    const rowResponsable = normalizeText_(row[CONFIG.columns.responsableEntrega - 1]);
    const qtyEntregada = Number(row[CONFIG.columns.cantidadEntregada - 1]) || 0;
    const solicitudFlag = String(row[CONFIG.columns.responsableSolicitud - 1] || '').trim();
    return (
      rowDate === expectedDate &&
      rowHora === expectedHora &&
      rowSede === expectedSede &&
      rowResponsable === expectedResponsable &&
      qtyEntregada > 0 &&
      solicitudFlag === expectedSolicitudFlag
    );
  });

  if (!headersMatch) return false;

  const incomingSignature = items
    .map(
      (item) =>
        `${normalizeText_(item.productCode)}|${normalizeText_(item.unit)}|${Number(item.cantidadEntregada)}`
    )
    .sort()
    .join('||');

  const existingSignature = rows
    .map((row) => {
      const code = row[CONFIG.columns.codigo - 1];
      const unit = row[CONFIG.columns.unidad - 1];
      const qty = Number(row[CONFIG.columns.cantidadEntregada - 1]) || 0;
      return `${normalizeText_(code)}|${normalizeText_(unit)}|${qty}`;
    })
    .sort()
    .join('||');

  return incomingSignature === existingSignature;
}

function getTailRows_(sheet, count) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || count <= 0) return [];
  const totalDataRows = lastRow - 1;
  const rowsToRead = Math.min(count, totalDataRows);
  const startRow = lastRow - rowsToRead + 1;
  return sheet.getRange(startRow, 1, rowsToRead, 14).getValues();
}

function appendMermaSinSolicitud_(sheet, payload, item, qty) {
  const row = [
    payload.hora || '',
    payload.fecha || '',
    '',
    item.productCode || '',
    item.unit || '',
    item.productName || '',
    payload.sede || '',
    '',
    'SIN SOLICITUD',
    '',
    '',
    qty,
    '',
    new Date(),
  ];
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function recordMerma_(payload) {
  validateRequired_(payload, ['fecha', 'sede']);
  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) {
    throw new Error('Debes enviar al menos un producto.');
  }

  const sheet = getMainSheet_();
  const summary = { processed: 0, updated: 0, appended: 0 };
  items.forEach((item) => {
    const qty = Number(item.cantidadMerma) || 0;
    if (qty <= 0) {
      throw new Error(`La merma debe ser mayor a cero (${item.productCode || 'sin código'}).`);
    }

    appendMermaDirecta_(sheet, payload, item, qty);
    summary.appended += 1;
    summary.processed += 1;
  });

  return summary;
}

function appendMermaDirecta_(sheet, payload, item, qty) {
  const row = [
    payload.hora || '',
    payload.fecha || '',
    '',
    item.productCode || '',
    item.unit || '',
    item.productName || '',
    payload.sede || '',
    '',
    'SIN SOLICITUD',
    '',
    '',
    qty,
    '',
    new Date(),
  ];
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function getProducts_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.catalogSheetName);
  if (!sheet) throw new Error('No se encontró la pestaña PRODUCTOS.');
  const values = sheet.getDataRange().getValues();
  const [, ...rows] = values;
  return rows
    .filter((row) => row[0] && row[1])
    .map((row) => ({
      code: String(row[0]).trim(),
      description: String(row[1]).trim(),
      unit: String(row[2] || '').trim() || 'UND',
    }));
}

function parseBody_(e) {
  if (!e?.postData?.contents) throw new Error('Cuerpo vacío.');
  return JSON.parse(e.postData.contents);
}

function getSpreadsheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  if (!ss) throw new Error('No se pudo abrir el Spreadsheet.');
  return ss;
}

function getMainSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.mainSheetName);
  if (!sheet) throw new Error('No se encontró la pestaña principal.');
  return sheet;
}

function withLock_(callback) {
  const lock = getSafeLock_();
  lock.waitLock(20000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function getSafeLock_() {
  try {
    const docLock = LockService.getDocumentLock();
    if (docLock) {
      return docLock;
    }
  } catch (error) {
  }
  return LockService.getScriptLock();
}

function validateRequired_(payload, fields) {
  fields.forEach((field) => {
    const value = payload[field];
    const isString = typeof value === 'string';
    const normalized = isString ? value.trim() : value;
    if (isString) {
      payload[field] = normalized;
    }
    if (normalized === undefined || normalized === null) {
      throw new Error(`El campo ${field} es obligatorio.`);
    }
    if (typeof normalized === 'string' && normalized === '') {
      throw new Error(`El campo ${field} es obligatorio.`);
    }
  });
}

function normalizeDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, CONFIG.timeZone, 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (!text) return '';
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(text)) {
    const [day, month, year] = text.split('/');
    return `${year.padStart(4, '0')}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  }
  return text;
}

function normalizeText_(value) {
  return String(value || '').trim().toUpperCase();
}

function buildResponse_(success, data, message) {
  return ContentService.createTextOutput(
    JSON.stringify({ success, data, message })
  ).setMimeType(ContentService.MimeType.JSON);
}
