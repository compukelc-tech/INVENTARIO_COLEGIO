// =================================================================
// PROCESO 7: PANEL ADMINISTRATIVO PROTEGIDO
// =================================================================
function obtenerUsuariosPendientes() {
  try {
    const datos = _getHojaUsuarios().getDataRange().getValues();
    const lista = [];
    for (let i = 1; i < datos.length; i++) {
      const f = datos[i];
      if (!f || f.length === 0) continue;
      const usuario = f[COL_USUARIO] ? String(f[COL_USUARIO]).trim() : "";
      if (!usuario) continue;
      const estado = f[COL_ESTADO] ? String(f[COL_ESTADO]).trim().toLowerCase() : "";
      if (estado === "pendiente") {
        lista.push({
          usuario:     usuario,
          nombre:      f[COL_NOMBRE] ? String(f[COL_NOMBRE]) : "Sin Nombre",
          dependencia: f[COL_DEPENDENCIA] ? String(f[COL_DEPENDENCIA]) : "-",
          cargo:       f[COL_CARGO] ? String(f[COL_CARGO]) : "-",
          correo:      f[COL_CORREO] ? String(f[COL_CORREO]) : "Sin correo"
        });
      }
    }
    return { ok:true, lista:lista, total:lista.length };
  } catch(err) { return { ok:false, lista:[], total:0, mensaje:err.toString() }; }
}

function obtenerTodosUsuarios() {
  try {
    const datos = _getHojaUsuarios().getDataRange().getValues();
    const lista = [];
    for (let i = 1; i < datos.length; i++) {
      const f = datos[i];
      if (!f || f.length === 0) continue;
      const usuario = f[COL_USUARIO] ? String(f[COL_USUARIO]).trim() : "";
      if (!usuario) continue;
      const estado = f[COL_ESTADO] ? String(f[COL_ESTADO]).trim().toLowerCase() : "";
      if (estado === "pendiente" || estado === "eliminado") continue;
      lista.push({
        usuario:     usuario,
        nombre:      f[COL_NOMBRE] ? String(f[COL_NOMBRE]) : "Sin Nombre",
        rol:         f[COL_ROL] ? String(f[COL_ROL]) : "-",
        estado:      f[COL_ESTADO] ? String(f[COL_ESTADO]) : "-",
        correo:      f[COL_CORREO] ? String(f[COL_CORREO]) : "Sin correo"
      });
    }
    return { ok:true, lista:lista };
  } catch(err) { return { ok:false, lista:[], mensaje:err.toString() }; }
}

function aprobarUsuario(usuario, rol, claveTemporal, ejecutor) {
  try {
    const cuenta = _buscarUsuario(usuario);
    if (!cuenta) return { ok:false, mensaje:"Usuario no encontrado." };
    if (!claveTemporal || claveTemporal.trim() === "") return { ok:false, mensaje:"Falta contraseña temporal." };

    const rolNorm = String(rol).trim().toLowerCase();
    if (rolNorm.indexOf('super') !== -1) {
      const cuentaEjecutor = _buscarUsuario(ejecutor);
      if (!cuentaEjecutor || String(cuentaEjecutor.datos[COL_ROL]).toLowerCase().indexOf('super') === -1) {
        return { ok:false, mensaje:"Solo un Súper Administrador puede asignar cargos de nivel Súper." };
      }
    }

    const rolFinal = rolNorm.charAt(0).toUpperCase() + rolNorm.slice(1);
    const hoja = _getHojaUsuarios();
    hoja.getRange(cuenta.fila, COL_ROL+1).setValue(rolFinal);
    hoja.getRange(cuenta.fila, COL_CLAVE+1).setValue(claveTemporal.trim());
    hoja.getRange(cuenta.fila, COL_ESTADO+1).setValue("Aprobado");
    hoja.getRange(cuenta.fila, COL_PRIMER_ING+1).setValue("SI");

    _registrarAuditoria({ usuario:ejecutor, nombre:"Admin Panel", rol:"Admin", correo: "" }, "Aprobó usuario", "", "", "Usuario: "+usuario);
    return { ok:true, mensaje:"Usuario aprobado con rol "+rolFinal+"." };
  } catch(err) { return { ok:false, mensaje:err.message }; }
}

function rechazarUsuario(usuario) {
  try {
    const cuenta = _buscarUsuario(usuario);
    if (!cuenta) return { ok:false, mensaje:"Usuario no encontrado." };
    _getHojaUsuarios().getRange(cuenta.fila, COL_ESTADO+1).setValue("Eliminado");
    return { ok:true, mensaje:"Solicitud rechazada." };
  } catch(err) { return { ok:false, mensaje:err.message }; }
}

function cambiarEstadoUsuario(usuario, nuevoEstado, ejecutor) {
  try {
    const cuentaTarget = _buscarUsuario(usuario);
    if (!cuentaTarget) return { ok:false, mensaje:"Usuario no encontrado." };
    
    const rolTarget = String(cuentaTarget.datos[COL_ROL]).toLowerCase();
    if (rolTarget.indexOf('super') !== -1) {
      const cuentaEjecutor = _buscarUsuario(ejecutor);
      if (!cuentaEjecutor || String(cuentaEjecutor.datos[COL_ROL]).toLowerCase().indexOf('super') === -1) {
        return { ok:false, mensaje: "Acceso denegado. Las cuentas Súper Administradoras solo se gestionan por otro miembro Súper." };
      }
    }

    const estFinal = String(nuevoEstado).trim().charAt(0).toUpperCase() + String(nuevoEstado).trim().slice(1).toLowerCase();
    _getHojaUsuarios().getRange(cuentaTarget.fila, COL_ESTADO+1).setValue(estFinal);
    return { ok:true, mensaje:"Estado actualizado a "+estFinal+"." };
  } catch(err) { return { ok:false, mensaje:err.message }; }
}

function guardarEvidenciaFotografica(base64Data, mimeType, codigo, usuarioLogueado) {
  try {
    const ext = ({"image/jpeg":"jpg","image/png":"png","image/webp":"webp"})[mimeType] || "jpg";
    let b64 = base64Data.indexOf(",") !== -1 ? base64Data.split(",")[1] : base64Data;
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), mimeType, "EVIDENCIA_" + String(codigo).trim() + "." + ext);
    const carpeta = DriveApp.getFolderById(CARPETA_ID);
    const previos = carpeta.getFilesByName(blob.getName());
    while (previos.hasNext()) previos.next().setTrashed(true);
    
    const archivo = carpeta.createFile(blob);
    try { archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    const fileId = archivo.getId();
    const enlaceGuardado = "https://lh3.googleusercontent.com/d/" + fileId + "|" + fileId;

    const hojaInv = obtenerSpreadsheetActivo().getSheetByName(NOMBRE_HOJA);
    const datosInv = hojaInv.getDataRange().getValues();
    let fila = -1;
    for (let i = 1; i < datosInv.length; i++) {
      if (String(datosInv[i][0]).trim().toLowerCase() === String(codigo).trim().toLowerCase()) { fila = i + 1; break; }
    }
    if (fila !== -1) hojaInv.getRange(fila, 26).setValue(enlaceGuardado);

    return { ok:true, enlace:enlaceGuardado, fileId:fileId, mensaje:"Evidencia guardada." };
  } catch(err) { return { ok:false, mensaje:err.message }; }
}
