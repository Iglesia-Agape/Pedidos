/**
 * Script de Ágape Coffee — Apps Script
 *
 * Tiene DOS funciones, distinguidas por el parámetro "accion" en la URL:
 *
 * 1) Sin parámetro (o accion=contador): el contador centralizado de número
 *    de orden, igual que antes. Esto es lo que usa la página de pedidos.
 *
 * 2) accion=resumen: lee el Sheet de respuestas del formulario y devuelve
 *    cuántos pedidos hay HOY y el total vendido HOY. Esto es lo que usa
 *    el panel de "pedidos de hoy" (panel.html), protegido con clave simple.
 *
 * IMPORTANTE: la función de resumen necesita saber el nombre exacto de la
 * pestaña de respuestas y de las columnas. Si cambias el nombre de la
 * pestaña o de alguna columna en tu Sheet, hay que actualizar las
 * constantes de abajo para que coincidan.
 */

// Nombre exacto de la pestaña de respuestas del formulario.
const NOMBRE_HOJA_RESPUESTAS = 'Respuestas de formulario 1';

// Nombres exactos de las columnas relevantes (deben coincidir con el
// encabezado real de la fila 1 de tu Sheet).
const COL_MARCA_TEMPORAL = 'Marca temporal';
const COL_TOTAL_A_PAGAR = 'Total a pagar';

function doGet(e) {
  const accion = (e && e.parameter && e.parameter.accion) || 'contador';

  if (accion === 'resumen') {
    return responderResumenDelDia();
  }

  return responderSiguienteOrden();
}

// ===================== CONTADOR DE ÓRDENES (ya existente) =====================

function responderSiguienteOrden() {
  const props = PropertiesService.getScriptProperties();
  const hoy = obtenerFechaHoy();
  const fechaGuardada = props.getProperty('fecha');
  let ultimo = parseInt(props.getProperty('ultimo') || '0', 10);

  if (fechaGuardada !== hoy) {
    ultimo = 0;
    props.setProperty('fecha', hoy);
  }

  ultimo += 1;
  props.setProperty('ultimo', String(ultimo));

  return jsonOutput({ orden: ultimo, fecha: hoy });
}

// ===================== RESUMEN DEL DÍA (nuevo) =====================

function responderResumenDelDia() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_RESPUESTAS);
    if (!hoja) {
      return jsonOutput({ error: 'No se encontró la hoja "' + NOMBRE_HOJA_RESPUESTAS + '".' });
    }

    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) {
      return jsonOutput({ pedidos: 0, total: 0, fecha: obtenerFechaHoy() });
    }

    const encabezados = datos[0];
    const idxFecha = encabezados.indexOf(COL_MARCA_TEMPORAL);
    const idxTotal = encabezados.indexOf(COL_TOTAL_A_PAGAR);

    if (idxFecha === -1 || idxTotal === -1) {
      return jsonOutput({
        error: 'No se encontraron las columnas esperadas. Revisa COL_MARCA_TEMPORAL y COL_TOTAL_A_PAGAR en el script.'
      });
    }

    const hoy = obtenerFechaHoy();
    let pedidosHoy = 0;
    let totalHoy = 0;

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const fechaFila = fila[idxFecha];
      if (!fechaFila) continue;

      const fechaFilaTexto = Utilities.formatDate(new Date(fechaFila), 'America/Tegucigalpa', 'yyyy-MM-dd');
      if (fechaFilaTexto !== hoy) continue;

      pedidosHoy += 1;

      // La columna "Total a pagar" se guarda como texto, ej: "L. 190.00".
      // Extraemos solo el número.
      const totalTexto = String(fila[idxTotal] || '');
      const numeroLimpio = totalTexto.replace(/[^0-9.]/g, '');
      const totalFila = parseFloat(numeroLimpio);
      if (!isNaN(totalFila)) totalHoy += totalFila;
    }

    return jsonOutput({
      pedidos: pedidosHoy,
      total: Math.round(totalHoy * 100) / 100,
      fecha: hoy
    });

  } catch (err) {
    return jsonOutput({ error: 'Error al leer el resumen: ' + err.message });
  }
}

// ===================== UTILIDADES =====================

function obtenerFechaHoy() {
  // Zona horaria de Honduras (Tegucigalpa), para que el "día" coincida
  // con el horario real del local y no con el de los servidores de Google.
  return Utilities.formatDate(new Date(), 'America/Tegucigalpa', 'yyyy-MM-dd');
}

function jsonOutput(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
