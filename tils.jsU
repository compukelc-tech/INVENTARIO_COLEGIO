// =================================================================
// PROCESO 5 Y UTILIDADES
// =================================================================
function extraerFileIdGS(url) {
  if (!url) return null;
  if (url.indexOf('|') !== -1) return url.split('|')[1].trim();
  let m = url.match(/lh3\.googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/) || url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

function formatearFecha(valor) { 
  return valor instanceof Date ? Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy") : (valor || ""); 
}

function _ahora() { 
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"); 
}

function _getHojaAuditoria() {
  const ss = obtenerSpreadsheetActivo();
  let hoja = ss.getSheetByName(NOMBRE_AUDITORIA);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_AUDITORIA);
    hoja.appendRow(["fecha_hora","usuario","nombre_usuario","rol","accion","cod_bien","nom_bien","detalle","correo"]);
    hoja.getRange(1,1,1,9).setFontWeight("bold").setBackground("#2d3436").setFontColor("#ffffff");
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function _registrarAuditoria(sesion, accion, codBien, nomBien, detalle) {
  try {
    _getHojaAuditoria().appendRow([
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
      sesion.usuario||"", sesion.nombre||"", sesion.rol||"", accion||"", codBien||"", nomBien||"", detalle||"", sesion.correo||""
    ]);
  } catch(e) {}
}
