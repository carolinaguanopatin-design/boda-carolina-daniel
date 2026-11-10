/**
 * BODA CAROLINA & DANIEL — REGISTRO DE ASISTENCIA
 * 1) Crea una hoja de cálculo privada de Google Sheets.
 * 2) Extensiones > Apps Script: pega este archivo.
 * 3) Ejecuta configurarPanel() una sola vez y autoriza.
 * 4) Implementar > Nueva implementación > Aplicación web.
 *    Ejecutar como: Yo | Acceso: Cualquier persona (no requiere iniciar sesión).
 * 5) Comparte la URL /exec con quien configura la invitación.
 * IMPORTANTE: Este endpoint es público. No introduzcas información sensible.
 */
const HOJA_RESPUESTAS = 'Respuestas';
const HOJA_PANEL = 'Panel';
function configurarPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA_RESPUESTAS) || ss.insertSheet(HOJA_RESPUESTAS);
  if (sh.getLastRow() === 0) sh.appendRow(['Fecha y hora','Código de invitación','Invitados','¿Asistirán?','Pases reservados','Mensaje']);
  sh.setFrozenRows(1);
  sh.getRange('A1:F1').setFontWeight('bold').setBackground('#e9dfcf');
  sh.autoResizeColumns(1,6);
  let panel = ss.getSheetByName(HOJA_PANEL) || ss.insertSheet(HOJA_PANEL);
  panel.clear();
  panel.getRange('A1:B1').merge().setValue('CAROLINA & DANIEL · CONTROL DE ASISTENCIA').setFontSize(16).setFontWeight('bold');
  panel.getRange('A3').setValue('Invitaciones que respondieron');
  panel.getRange('B3').setFormula('=COUNTA(Respuestas!B2:B)');
  panel.getRange('A4').setValue('Pases confirmados');
  panel.getRange('B4').setFormula('=SUMIF(Respuestas!D2:D,"Sí",Respuestas!E2:E)');
  panel.getRange('A5').setValue('Pases que no asistirán');
  panel.getRange('B5').setFormula('=SUMIF(Respuestas!D2:D,"No",Respuestas!E2:E)');
  panel.getRange('A6').setValue('Respuestas: sí');
  panel.getRange('B6').setFormula('=COUNTIF(Respuestas!D2:D,"Sí")');
  panel.getRange('A7').setValue('Respuestas: no');
  panel.getRange('B7').setFormula('=COUNTIF(Respuestas!D2:D,"No")');
  panel.getRange('A9').setValue('Nota: si una invitación responde otra vez, conserva ambas filas; para un conteo definitivo revisa la última respuesta de cada código.');
  panel.setColumnWidth(1,330);
  panel.setColumnWidth(2,150);
  panel.getRange('A3:A7').setFontWeight('bold');
}
function doPost(e) {
  try {
    const p = e && e.parameter || {};
    const nombre = String(p.nombre || '').trim().slice(0,120);
    const codigo = String(p.codigo || '').trim().slice(0,35);
    const opcion = String(p.asistencia || '').trim();
    const pases = Number(p.pases || 1);
    const mensaje = String(p.mensaje || '').trim().slice(0,1000);
    if (!nombre || !['Sí','No'].includes(opcion) || !Number.isInteger(pases) || pases < 1 || pases > 15) {
      return respuesta('No pudimos registrar tu confirmación. Regresa a la invitación y completa los datos.', false);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(HOJA_RESPUESTAS);
    if (!hoja) { configurarPanel(); hoja = ss.getSheetByName(HOJA_RESPUESTAS); }
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { hoja.appendRow([new Date(),codigo,nombre,opcion,pases,mensaje]); }
    finally { lock.releaseLock(); }
    return respuesta(opcion === 'Sí' ? '¡Gracias! Tu asistencia quedó registrada.' : 'Gracias por avisarnos. Tu respuesta quedó registrada.', true);
  } catch (err) {
    console.error(err);
    return respuesta('Hubo un problema al guardar tu respuesta. Intenta nuevamente.', false);
  }
}
function respuesta(mensaje, exito) {
  return HtmlService.createHtmlOutput('<!doctype html><html lang="es"><meta name="viewport" content="width=device-width,initial-scale=1"><meta charset="utf-8"><title>Confirmación · Carolina y Daniel</title><body style="font-family:Georgia,serif;background:#f6f0e5;color:#283d32;text-align:center;padding:12vh 18px"><div style="max-width:500px;margin:auto"><div style="font-size:35px">Carolina & Daniel</div><p style="font-size:21px;line-height:1.7">'+mensaje+'</p><p>23 de enero de 2027</p></div></body></html>');
}
