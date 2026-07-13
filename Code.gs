/**
 * Informe de Predicación — Backend en Google Apps Script
 * Reemplaza el workflow de n8n (login / verificar / informe).
 *
 * CÓMO USAR:
 * 1. Abrí tu planilla de Google Sheets (la del Maestro y los Grupos).
 * 2. Extensiones → Apps Script. Pegá este código en Code.gs.
 * 3. Implementar → Nueva implementación → Aplicación web:
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 * 4. Copiá la URL (termina en /exec) y pegala en APPS_SCRIPT_URL del index.html.
 */

const SS_ID = '1Ujwram5ankf0iCtYT4w1o_5fnIWjdPWN__fYzEpqOSI';
const HOJA_MAESTRO = 'Maestro';
const TZ = 'America/Argentina/Buenos_Aires';

// ================== ROUTER ==================

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return responder({ error: 'Body inválido' });
  }

  let result;
  try {
    switch (body.accion) {
      case 'login':     result = login(body);           break;
      case 'verificar': result = verificar(body);       break;
      case 'informe':   result = guardarInforme(body);  break;
      default:          result = { error: 'Acción desconocida' };
    }
  } catch (err) {
    result = { error: String(err) };
  }

  return responder(result);
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================== LOGIN ==================
// Busca "Apellido Nombre" en la hoja Maestro (cualquier columna).
// El nombre de la columna donde aparece es el grupo.

function login(body) {
  const nombre = String(body.nombre || '').trim().toLowerCase();
  const apellido = String(body.apellido || '').trim().toLowerCase();
  const target = (apellido + ' ' + nombre).replace(/\s+/g, ' ').trim();

  const hoja = SpreadsheetApp.openById(SS_ID).getSheetByName(HOJA_MAESTRO);
  const valores = hoja.getDataRange().getValues();
  if (valores.length < 2) return { encontrado: false, grupo: null };

  const encabezados = valores[0];

  for (let fila = 1; fila < valores.length; fila++) {
    for (let col = 0; col < encabezados.length; col++) {
      const val = limpiar(valores[fila][col]);
      if (val && val === target) {
        return { encontrado: true, grupo: String(encabezados[col]) };
      }
    }
  }
  return { encontrado: false, grupo: null };
}

// Misma limpieza que el nodo Code de n8n:
// quita "(super)", "(aux)", puntos; normaliza espacios y mayúsculas.
function limpiar(v) {
  return String(v || '')
    .replace(/\(super\)/i, '')
    .replace(/\(aux\)/i, '')
    .replace(/\./g, '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

// ================== VERIFICAR ==================
// ¿Ya existe un informe con esta Clave en la hoja del grupo?

function verificar(body) {
  const clave = armarClave(body);
  const fila = buscarFilaPorClave(String(body.grupo || ''), clave);
  return { yaInformado: fila !== null };
}

function armarClave(body) {
  return [body.apellido, body.nombre, body.periodo]
    .map(function (v) { return String(v || '').toLowerCase().trim(); })
    .join('|');
}

/** Devuelve el número de fila (1-based) donde está la clave, o null. */
function buscarFilaPorClave(nombreGrupo, clave) {
  const hoja = SpreadsheetApp.openById(SS_ID).getSheetByName(nombreGrupo);
  if (!hoja) throw new Error('No existe la hoja "' + nombreGrupo + '"');

  const valores = hoja.getDataRange().getValues();
  if (valores.length < 2) return null;

  const colClave = valores[0].indexOf('Clave');
  if (colClave === -1) throw new Error('La hoja "' + nombreGrupo + '" no tiene columna "Clave"');

  for (let f = 1; f < valores.length; f++) {
    const val = String(valores[f][colClave] || '').toLowerCase().trim();
    if (val && val === clave) return f + 1; // fila real en la hoja
  }
  return null;
}

// ================== GUARDAR INFORME ==================
// Equivalente al "appendOrUpdate" de n8n: si la Clave existe actualiza
// la fila; si no, agrega una nueva. Mapea por NOMBRE de encabezado,
// así el orden de columnas de la hoja no importa.

function guardarInforme(body) {
  const grupo = String(body.grupo || '');
  const clave = armarClave(body);

  const datos = {
    'Fecha de registro':    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'),
    'Clave':                clave,
    'Nombre':               String(body.nombre || ''),
    'Apellido':             String(body.apellido || ''),
    'Periodo':              String(body.periodo || ''),
    'Tuvo actividad':       body.tuvoActividad ? 'Si' : 'No',
    'Tipo de actividad':    String(body.tipoActividad || '-'),
    'Horas':                String(body.horas || ''),
    'Cantidad de estudios': body.estudios || 0
  };

  const hoja = SpreadsheetApp.openById(SS_ID).getSheetByName(grupo);
  if (!hoja) throw new Error('No existe la hoja "' + grupo + '"');

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const filaNueva = encabezados.map(function (h) {
    return datos.hasOwnProperty(h) ? datos[h] : '';
  });

  // Bloqueo para evitar que dos envíos simultáneos dupliquen filas
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const filaExistente = buscarFilaPorClave(grupo, clave);
    if (filaExistente) {
      hoja.getRange(filaExistente, 1, 1, filaNueva.length).setValues([filaNueva]);
    } else {
      hoja.appendRow(filaNueva);
    }
  } finally {
    lock.releaseLock();
  }

  return { ok: true };
}
