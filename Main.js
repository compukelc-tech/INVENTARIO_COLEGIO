// =================================================================
// RESOLVEDOR DINÁMICO DE BASE DE DATOS
// =================================================================
function obtenerSpreadsheetActivo() {
  try {
    const ssActivo = SpreadsheetApp.getActiveSpreadsheet();
    if (ssActivo && ssActivo.getId()) return ssActivo;
  } catch(e) {}

  try {
    const cache = CacheService.getScriptCache();
    let idAlmacenado = cache.get("ID_DATABASE_ACTIVA");
    if (idAlmacenado) return SpreadsheetApp.openById(idAlmacenado);

    const ssCentral = SpreadsheetApp.openByUrl(URL_CENTRAL_COMPUKELC);
    const sheetEmpresas = ssCentral.getSheetByName('Empresas');
    if (sheetEmpresas) {
      const data = sheetEmpresas.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString().toUpperCase() === ID_PROYECTO_COMPUKELC) {
          let idEncontrado = data[i][4] ? data[i][4].toString().trim() : ""; 
          if (idEncontrado) {
            cache.put("ID_DATABASE_ACTIVA", idEncontrado, 10800);
            return SpreadsheetApp.openById(idEncontrado);
          }
        }
      }
    }
  } catch (error) {}

  return SpreadsheetApp.openById("1cT177KBK-dNw1c4Q2tHB6jhRFM1Sjx3DgZ9aYLJA29g");
}

// =================================================================
// PROCESO 2: doGet — PUNTO DE ENTRADA CON CONVERSIÓN DE IMAGEN B64
// =================================================================
function doGet(e) {
  const estadoSistema = verificarEstadoServicio_();

  if (estadoSistema === 'Inactivo') {
    return HtmlService.createHtmlOutput(`<div style="text-align: center; padding: 50px;"><h2>Sistema Suspendido</h2><p>Acceso restringido administrativamente.</p><p>Proveedor: <b>compukelc</b></p></div>`).setTitle("Acceso Restringido - compukelc").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (estadoSistema === 'Mantenimiento') {
    return HtmlService.createHtmlOutput(`<div style="text-align: center; padding: 50px;"><h2>Mantenimiento Programado</h2><p>Actualización estructural en curso.</p></div>`).setTitle("Mantenimiento - compukelc").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (estadoSistema === 'No Encontrado') {
    return HtmlService.createHtmlOutput(`<div style="text-align: center; padding: 50px;"><h2>Error de Enlace Central</h2><p>Proyecto no registrado en matriz compukelc.</p></div>`).setTitle("Error de Registro").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  let script = '<script>window.codigoPrecargado = null; window.datosPrecargados = null;<\/script>';

  if (e && e.parameter && e.parameter.codigo) {
    const cod = String(e.parameter.codigo).trim();
    const datos = obtenerDatosPorCodigo(cod);
    script = `<script>window.codigoPrecargado = ${JSON.stringify(cod)}; window.datosPrecargados = ${JSON.stringify(datos)};<\/script>`;
  }

  const appConfig = getAppConfig();
  
  if (appConfig.url_logo && String(appConfig.url_logo).trim() !== "") {
    let idMatch = String(appConfig.url_logo).match(/[-\w]{25,}/);
    if (idMatch) {
      try {
        let archivoLogo = DriveApp.getFileById(idMatch[0]);
        let blob = archivoLogo.getBlob();
        appConfig.logoBase64 = "data:" + (blob.getContentType() || "image/png") + ";base64," + Utilities.base64Encode(blob.getBytes());
      } catch (e) {
        appConfig.logoBase64 = null;
      }
    }
  }

  const tituloPestaña = "Inventario · " + appConfig.nombre_institucion;
  const scriptConfig = `<script>window.appConfigDinamica = ${JSON.stringify(appConfig)};<\/script>`;

  const html = HtmlService.createHtmlOutputFromFile("Interfaz");
  return HtmlService.createHtmlOutput(
    html.getContent().replace(/<\/head>/i, script + scriptConfig + '</head>')
  )
  .setTitle(tituloPestaña)
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no");
}


