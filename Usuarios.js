// =================================================================
// PROCESO 4 Y 6: USUARIOS Y AUTENTICACIÓN
// =================================================================
function _getHojaUsuarios() {
  const ss   = obtenerSpreadsheetActivo();
  let  hoja = ss.getSheetByName(NOMBRE_USUARIOS);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_USUARIOS);
    hoja.appendRow(["usuario","nombre","dependencia","cargo","rol","contrasena","estado","primer_ingreso","fecha_registro","ultima_modificacion","correo"]);
    hoja.getRange(1,1,1,11).setFontWeight("bold").setBackground("#8b1a1a").setFontColor("#ffffff");
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function _buscarUsuario(usuario) {
  const datos = _getHojaUsuarios().getDataRange().getValues();
  const u = String(usuario).trim().toLowerCase();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][COL_USUARIO]).trim().toLowerCase() === u) return { fila: i + 1, datos: datos[i] };
  }
  return null;
}

function _buscarUsuarioPorCorreo(correo) {
  const datos = _getHojaUsuarios().getDataRange().getValues();
  const c = String(correo).trim().toLowerCase();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][COL_CORREO] && String(datos[i][COL_CORREO]).trim().toLowerCase() === c) return { fila: i + 1, datos: datos[i] };
  }
  return null;
}

function _insertarUsuario(fila) {
  const hoja = _getHojaUsuarios();
  const datos = hoja.getDataRange().getValues();
  let filaDestino = -1;
  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][COL_USUARIO] || String(datos[i][COL_USUARIO]).trim() === "") { filaDestino = i + 1; break; }
  }
  hoja.getRange(filaDestino === -1 ? datos.length + 1 : filaDestino, 1, 1, fila.length).setValues([fila]);
}

function _generarNombreUsuario(nombre) {
  if (!nombre) return "";
  const p = nombre.trim().toLowerCase().split(/\s+/);
  if (p.length === 0) return "";
  let ini = p.length === 2 ? p[1].charAt(0) : (p.length === 3 ? p[1].charAt(0)+p[2].charAt(0) : (p.length >= 4 ? p[p.length-2].charAt(0)+p[p.length-1].charAt(0) : ""));
  return (p[0] + ini).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
}

function registrarUsuario(info) {
  try {
    const nombreCompleto = String(info.nombre).trim(), correo = String(info.correo).trim().toLowerCase();
    if (!nombreCompleto || nombreCompleto.length < 4) return { ok:false, mensaje:"Nombre completo no válido." };
    if (!correo || correo.indexOf('@') === -1) return { ok:false, mensaje:"Correo electrónico no válido." };
    
    const existente = _buscarUsuarioPorCorreo(correo);
    if (existente) {
      const est = String(existente.datos[COL_ESTADO]).trim().toLowerCase();
      if (est === "pendiente" || est === "aprobado") return { ok:false, mensaje:"Este correo ya tiene cuenta o solicitud." };
    }
    
    const usuarioBase = _generarNombreUsuario(nombreCompleto);
    let usuarioFinal = usuarioBase, sufijo = 1;
    while (_buscarUsuario(usuarioFinal)) { usuarioFinal = usuarioBase + sufijo; sufijo++; }
    
    _insertarUsuario([usuarioFinal, nombreCompleto, String(info.dependencia).trim(), String(info.cargo).trim(), "Pendiente", "", "Pendiente", "SI", _ahora(), _ahora(), correo]);
    _registrarAuditoria({ usuario:usuarioFinal, nombre:nombreCompleto, rol:"Pendiente", correo: correo }, "Solicitó registro", "", "", "Dep: "+info.dependencia);
        
    return { ok:true, usuarioAsignado: usuarioFinal, mensaje: "Registro exitoso." };
  } catch(err) { return { ok:false, mensaje:err.message }; }
}

function iniciarSesion(usuario, clave) {
  try {
    const cuenta = _buscarUsuario(usuario);
    if (!cuenta) return { ok:false, mensaje:"Usuario no registrado." };
    const d = cuenta.datos, est = String(d[COL_ESTADO]).trim().toLowerCase(), rol = String(d[COL_ROL]).trim().toLowerCase();
    
    if (est !== "aprobado") return { ok:false, mensaje:"Cuenta " + est + "." };
    if (rol === "pendiente" || rol === "") return { ok:false, mensaje:"Rol no asignado." };
    
    const claveAlmacenada = String(d[COL_CLAVE]).trim();
    const hashIngresado = generarHashSHA256(clave);
    
    // Soporte retroactivo para el hash de seguridad implementado
    if (claveAlmacenada !== hashIngresado && claveAlmacenada !== String(clave).trim()) {
      return { ok:false, mensaje:"Contraseña incorrecta." };
    }
    
    let esPrimerIngreso = String(d[COL_PRIMER_ING]).trim().toUpperCase() === "SI";
    if (claveAlmacenada === String(clave).trim() && claveAlmacenada !== hashIngresado) {
       esPrimerIngreso = true; // Fuerza cambio para que la nueva quede encriptada
    }

    _registrarAuditoria({ usuario:d[COL_USUARIO], nombre:d[COL_NOMBRE], rol:d[COL_ROL], correo:d[COL_CORREO]?String(d[COL_CORREO]):"" }, "Inició sesión","","","");
    return { ok:true, usuario: d[COL_USUARIO], nombre: d[COL_NOMBRE], rol: rol, primerIngreso: esPrimerIngreso };
  } catch(err) { return { ok:false, mensaje:err.message }; }
}

function cambiarContrasena(usuario, actual, nueva) {
  try {
    const cuenta = _buscarUsuario(usuario);
    if (!cuenta) return { ok:false, mensaje:"Cuenta no encontrada." };
    
    const claveAlmacenada = String(cuenta.datos[COL_CLAVE]).trim();
    const hashActualIngresado = generarHashSHA256(actual);
    
    if (claveAlmacenada !== hashActualIngresado && claveAlmacenada !== String(actual).trim()) {
      return { ok:false, mensaje:"Contraseña actual incorrecta." };
    }
    if (!nueva || nueva.length < 6) return { ok:false, mensaje:"Mínimo 6 caracteres." };
    
    const hoja = _getHojaUsuarios();
    hoja.getRange(cuenta.fila, COL_CLAVE+1).setValue(generarHashSHA256(nueva));
    hoja.getRange(cuenta.fila, COL_PRIMER_ING+1).setValue("NO");
    hoja.getRange(cuenta.fila, COL_ULT_MOD+1).setValue(_ahora());
    
    _registrarAuditoria({ usuario:cuenta.datos[COL_USUARIO], nombre:cuenta.datos[COL_NOMBRE], rol:cuenta.datos[COL_ROL], correo:cuenta.datos[COL_CORREO]?String(cuenta.datos[COL_CORREO]):"" }, "Cambió contraseña","","","");
    return { ok:true, mensaje:"Contraseña actualizada." };
  } catch(err) { return { ok:false, mensaje:err.message }; }
}
