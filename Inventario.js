// =================================================================
// PROCESO 3: INVENTARIO — OPERACIONES DE LECTURA Y FILTRADO
// =================================================================
function obtenerDatosPorCodigo(codigo) {
  const ss   = obtenerSpreadsheetActivo();
  const hoja = ss.getSheetByName(NOMBRE_HOJA);
  if (!hoja) return { error: "No se encontró la hoja: " + NOMBRE_HOJA };

  const datos      = hoja.getDataRange().getValues();
  const codBuscado = String(codigo).trim().toLowerCase();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] && String(datos[i][0]).trim().toLowerCase() === codBuscado) {
      const row = datos[i];
      return {
        cod_material:       row[0]  || "", nom_material:       row[1]  || "",
        cod_grupo_mat:      row[2]  || "", nom_grupo_mat:      row[3]  || "",
        cod_grupo_contable: row[4]  || "", nom_grupo_contable: row[5]  || "",
        cod_servicio:       row[6]  || "", nom_servicio:       row[7]  || "",
        cod_ubicacion:      row[8]  || "", nom_ubicacion:      row[9]  || "",
        nom_responsable:    row[10] || "", serial:             row[11] || "",
        tipo_adquisicion:   row[12] || "", nro_compra:         row[13] || "",
        fecha_compra:       formatearFecha(row[14]), fecha_entrada: formatearFecha(row[15]),
        estado_bien:        row[16] || "", modelo:             row[17] || "",
        marca:              row[18] || "", descripcion:        row[19] || "",
        fecha_baja:         formatearFecha(row[20]), motivo_baja:  row[21] || "",
        recurso:            row[22] || "", observacion:        row[23] || "",
        enlace_evidencia:   String(row[25] || "").trim()
      };
    }
  }
  return null;
}

function servirImagenBase64(enlaceEvidencia) {
  try {
    if (!enlaceEvidencia) return { ok: false, b64: "" };
    const fileId = extraerFileIdGS(enlaceEvidencia);
    if (!fileId) return { ok: false, b64: "" };
    const archivo = DriveApp.getFileById(fileId);
    const blob    = archivo.getBlob();
    return { ok: true, b64: "data:" + (blob.getContentType() || "image/jpeg") + ";base64," + Utilities.base64Encode(blob.getBytes()) };
  } catch(e) { return { ok: false, b64: "", error: e.message }; }
}

function obtenerOpcionesFiltros() {
  const ss = obtenerSpreadsheetActivo();
  const hoja = ss.getSheetByName(NOMBRE_HOJA);
  const datos = hoja.getDataRange().getValues();
  
  const sets = {
    materiales: new Set(),
    grupos: new Set(),
    ubicaciones: new Set(),
    responsables: new Set(),
    servicios: new Set()
  };
  
  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    if (!row[0]) continue; 
    
    if (row[1]) sets.materiales.add(String(row[1]).trim());
    if (row[5]) sets.grupos.add(String(row[5]).trim());
    if (row[7]) sets.servicios.add(String(row[7]).trim());
    if (row[9]) sets.ubicaciones.add(String(row[9]).trim());
    if (row[10]) sets.responsables.add(String(row[10]).trim());
  }
  
  return {
    materiales: Array.from(sets.materiales).sort(),
    grupos: Array.from(sets.grupos).sort(),
    servicios: Array.from(sets.servicios).sort(),
    ubicaciones: Array.from(sets.ubicaciones).sort(),
    responsables: Array.from(sets.responsables).sort()
  };
}

function filtrarInventarioAvanzado(criterios) {
  const hoja = obtenerSpreadsheetActivo().getSheetByName(NOMBRE_HOJA);
  const datos = hoja.getDataRange().getValues();
  const resultados = [];
  
  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    if (!row[0]) continue;
    
    let coincide = true;
    if (criterios.ubicacion && String(row[9]).trim() !== criterios.ubicacion) coincide = false;
    if (criterios.servicio && String(row[7]).trim() !== criterios.servicio) coincide = false;
    if (criterios.material && String(row[1]).trim() !== criterios.material) coincide = false;
    if (criterios.grupo && String(row[5]).trim() !== criterios.grupo) coincide = false;
    if (criterios.responsable && String(row[10]).trim() !== criterios.responsable) coincide = false;
    
    if (coincide) {
      resultados.push({
        cod_material:       row[0]  || "", nom_material:       row[1]  || "",
        cod_grupo_mat:      row[2]  || "", nom_grupo_mat:      row[3]  || "",
        cod_grupo_contable: row[4]  || "", nom_grupo_contable: row[5]  || "",
        cod_servicio:       row[6]  || "", nom_servicio:       row[7]  || "",
        cod_ubicacion:      row[8]  || "", nom_ubicacion:      row[9]  || "",
        nom_responsable:    row[10] || "", serial:             row[11] || "",
        tipo_adquisicion:   row[12] || "", nro_compra:         row[13] || "",
        fecha_compra:       formatearFecha(row[14]), fecha_entrada: formatearFecha(row[15]),
        estado_bien:        row[16] || "", modelo:             row[17] || "",
        marca:              row[18] || "", descripcion:        row[19] || "",
        fecha_baja:         formatearFecha(row[20]), motivo_baja:  row[21] || "",
        recurso:            row[22] || "", observacion:        row[23] || ""
      });
    }
  }
  return resultados;
}


