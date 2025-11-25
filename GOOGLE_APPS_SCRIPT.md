# Google Apps Script · Formularios LA TATA DE LA LIBERTAD

Sigue estos pasos para conectar los formularios con el Google Sheets mostrado en las capturas:

1. Abre el Google Sheets `LA TATA DE LA LIBERTAD` y navega a **Extensiones → Apps Script**.
2. Elimina cualquier código existente y pega el script que encontrarás más abajo.
3. Guarda el proyecto, asígnale un nombre (por ejemplo `formularios-latata`) y presiona **Deploy → New deployment**.
4. Selecciona **Web app**, elige *Anyone* o *Anyone with the link* para permitir que Vercel acceda, y copia la URL pública.
5. Sustituye la constante `window.APPS_SCRIPT_URL` en `index.html` con la URL copiada.

> Si actualizas el script en el futuro, recuerda crear un nuevo deployment o actualizar el existente para que mantenga la misma URL.

## Código del Apps Script

```javascript
const CONFIG = {
  spreadsheetId: '18WPHKhmnGtoNiHuALuK8486VuJeMq8LHF0tKZArq3hs',
  mainSheetName: 'LA TATA DE LA LIBERTAD',
  catalogSheetName: 'PRODUCTOS',
  timeZone: Session.getScriptTimeZone() || 'America/Caracas',
  columns: {
    hora: 1,
    fecha: 2,
    familia: 3,
    sede: 4,
    codigo: 5,
    producto: 6,
    unidad: 7,
    cantidadSolicitada: 8,
    responsableSolicitud: 9,
    cantidadEntregada: 10,
    responsableEntrega: 11,
    merma: 12,
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
      return { data, message: 'Solicitudes registradas.' };
    }
    case 'recordentrega': {
      const data = recordEntrega_(payload);
      return { data, message: `Registros procesados: ${data.processed}` };
    }
    case 'recordmerma': {
      const data = recordMerma_(payload);
      return { data, message: `Mermas registradas: ${data.processed}` };
    }
    default:
      throw new Error('Acción POST no soportada.');
  }
}

function createSolicitud_(payload) {
  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) throw new Error('Debes enviar al menos un producto.');

  const sheet = getMainSheet_();
  const rows = items.map((item) => [
    payload.hora || '',
    payload.fecha || '',
    '',
    payload.sede || '',
    item.code || '',
    item.description || '',
    item.unit || '',
    Number(item.quantity) || 0,
    payload.responsable || '',
    '',
    '',
    '',
  ]);

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { rowsInserted: rows.length };
}

function recordEntrega_(payload) {
  validateRequired_(payload, ['fecha', 'sede', 'responsableEntrega']);
  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) {
    throw new Error('Debes enviar al menos un producto.');
  }

  const sheet = getMainSheet_();
  const summary = { processed: 0, updated: 0, appended: 0 };

  items.forEach((item) => {
    const qty = Number(item.cantidadEntregada) || 0;
    if (qty <= 0) {
      throw new Error(`La cantidad entregada debe ser mayor a cero (${item.productCode || 'sin código'}).`);
    }

    if (payload.sinSolicitud) {
      appendEntregaSinSolicitud_(sheet, payload, item, qty);
      summary.appended += 1;
    } else {
      const rowIndex = findRowIndex_(sheet, payload.fecha, payload.sede, item.productCode);
      if (!rowIndex) {
        throw new Error(
          `No se encontró una solicitud con esa fecha, sede y producto (${item.productCode || 'sin código'}).`
        );
      }

      if (payload.hora) {
        sheet.getRange(rowIndex, CONFIG.columns.hora).setValue(payload.hora);
      }
      sheet.getRange(rowIndex, CONFIG.columns.cantidadEntregada).setValue(qty);
      sheet.getRange(rowIndex, CONFIG.columns.responsableEntrega).setValue(payload.responsableEntrega || '');
      summary.updated += 1;
    }

    summary.processed += 1;
  });

  return summary;
}

function appendEntregaSinSolicitud_(sheet, payload, item, qty) {
  const row = [
    payload.hora || '',
    payload.fecha || '',
    '',
    payload.sede || '',
    item.productCode || '',
    item.productName || '',
    item.unit || '',
    '',
    'SIN SOLICITUD',
    qty,
    payload.responsableEntrega || '',
    '',
  ];
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function recordMerma_(payload) {
  validateRequired_(payload, ['fecha']);
  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) {
    throw new Error('Debes enviar al menos un producto.');
  }

  const sheet = getMainSheet_();
  items.forEach((item) => {
    const qty = Number(item.cantidadMerma) || 0;
    if (qty <= 0) {
      throw new Error(`La merma debe ser mayor a cero (${item.productCode || 'sin código'}).`);
    }

    const rowIndex = findRowIndex_(sheet, payload.fecha, 'BC', item.productCode);
    if (!rowIndex) {
      throw new Error(
        `No se encontró una fila para esa fecha/producto en la sede BC (${item.productCode || 'sin código'}).`
      );
    }

    sheet.getRange(rowIndex, CONFIG.columns.merma).setValue(qty);
  });

  return { processed: items.length };
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

function findRowIndex_(sheet, dateValue, sede, productCode) {
  const values = sheet.getDataRange().getValues();
  const targetDate = normalizeDate_(dateValue);
  const targetSede = normalizeText_(sede);
  const targetCode = normalizeText_(productCode);

  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const rowDate = normalizeDate_(row[CONFIG.columns.fecha - 1]);
    const rowSede = normalizeText_(row[CONFIG.columns.sede - 1]);
    const rowCode = normalizeText_(row[CONFIG.columns.codigo - 1]);
    if (rowDate === targetDate && rowSede === targetSede && rowCode === targetCode) {
      return i + 1; // +1 porque getValues inicia en 0 y las filas en 1
    }
  }
  return null;
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
    // Sigue con el ScriptLock si el documento no existe (script standalone)
  }
  return LockService.getScriptLock();
}

function validateRequired_(payload, fields) {
  fields.forEach((field) => {
    if (!payload[field]) {
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
```

> El script utiliza `LockService` para evitar que dos usuarios escriban al mismo tiempo y se basa en la combinación `Fecha + Sede + Código` para ubicar las filas. Ajusta los nombres de las pestañas o columnas dentro del objeto `CONFIG` si tu hoja cambia en el futuro.
