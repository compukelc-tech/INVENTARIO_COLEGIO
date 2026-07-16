// =================================================================
// PROCESO 0: CONEXIÓN CON EL PANEL CENTRAL compukelc
// =================================================================
const URL_CENTRAL_COMPUKELC = 'https://drive.google.com/drive/folders/1oeWmxSRd_6Cr4BgkhP50i7EeoIdWbLeu?usp=sharing';
const ID_PROYECTO_COMPUKELC = 'INV-COLEGIOS'; 

// =================================================================
// PROCESO 1: CONFIGURACIÓN GLOBAL ENLAZADA
// =================================================================
const NOMBRE_HOJA       = "Hoja 1";
const NOMBRE_USUARIOS  = "Hoja 2";
const NOMBRE_AUDITORIA = "Hoja 3";
const NOMBRE_CONFIG    = "Configuracion"; 
const CARPETA_ID       = "1Qdb_gE8fnT1fUnLGjd1rTGsc7VLW-AgS"; // ID extraído de la carpeta de evidencias
const LOGO_FALLBACK_ID = "1djaYd2BcOVi1NlsR6jvVSyUC-Cu89EaQ";

const COL_USUARIO     = 0; 
const COL_NOMBRE      = 1; 
const COL_DEPENDENCIA = 2; 
const COL_CARGO       = 3; 
const COL_ROL         = 4; 
const COL_CLAVE       = 5; 
const COL_ESTADO      = 6; 
const COL_PRIMER_ING  = 7; 
const COL_FECHA_REG   = 8; 
const COL_ULT_MOD     = 9; 
const COL_CORREO      = 10;

function verificarEstadoServicio_() {
  try {
    const ssCentral = SpreadsheetApp.openByUrl(URL_CENTRAL_COMPUKELC);
    const sheetEmpresas = ssCentral.getSheetByName('Empresas');
    if (!sheetEmpresas) return 'Activo'; 
    
    const data = sheetEmpresas.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toUpperCase() === ID_PROYECTO_COMPUKELC) {
        return data[i][3].toString().trim();
      }
    }
    return 'No Encontrado'; 
  } catch (error) {
    return 'Inactivo'; 
  }
}

function getAppConfig() {
  try {
    const ss = obtenerSpreadsheetActivo();
    let sheet = ss.getSheetByName("Configuracion") || ss.getSheetByName("Configuración");
    if (!sheet) return { nombre_institucion: "compukelc", url_logo: "https://drive.google.com/file/d/17dMeVDNgpStmVzC6PKYm07EYkzUtHJ24/view?usp=drive_link" + LOGO_FALLBACK_ID };
    
    const data = sheet.getDataRange().getValues();
    let config = {};
    
    data.forEach(row => {
      if (row[0] && typeof row[0] === 'string' && row[0].trim() !== "") {
        let clave = row[0].trim().toLowerCase();
        let valor = row[1] ? String(row[1]).trim() : "";
        config[clave] = valor;
      }
    });

    if (!config.nombre_institucion) config.nombre_institucion = "compukelc";
    if (!config.url_logo || config.url_logo.trim() === "") config.url_logo = "https://drive.google.com/file/d/17dMeVDNgpStmVzC6PKYm07EYkzUtHJ24/view?usp=drive_link" + LOGO_FALLBACK_ID;

    return config;
  } catch (error) {
    return { nombre_institucion: "compukelc", url_logo: "https://drive.google.com/file/d/17dMeVDNgpStmVzC6PKYm07EYkzUtHJ24/view?usp=drive_link" + LOGO_FALLBACK_ID };
  }
}
