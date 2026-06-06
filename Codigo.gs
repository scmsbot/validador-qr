/*
================================================================================
  SISTEMA CONTROL DE ACCESO - CÓDIGO COMPLETO v8
  Correcciones: email con retorno de estado, errores no silenciosos,
  emailEnviado devuelto al frontend, cada correo en try-catch propio.
================================================================================
*/

var CONFIG = {
  EMAIL_ADMIN: 'scmsbot@gmail.com',
  EMPRESA: 'Hospital Felix Bulnes',
  HOJA_RESPUESTAS: 'Respuestas',
  HORAS_VALIDEZ: 24,

  EMAILS: {
    scms: ['M.poblete@scms.cl'],
    guardias: ['Marcelopoblete16@gmail.com'],
    inspeccion: ['Pobletechelo99@gmail.com']
  }
};

var COLS = {
  marcaTemporal: 1,
  emailFormulario: 2,
  horaIngreso: 3,
  horaSalida: 4,
  emailAutoriza: 5,
  nombreAutoriza: 6,
  nombreVisitante: 7,
  rutVisitante: 8,
  emailVisitante: 9,
  empresaProveedor: 10,
  patenteVisitante: 11,
  areaVisitar: 12,
  diaAutorizacion: 13,
  personaContacto: 14,
  observacionesAdicionales: 15,
  personaContacto2: 16,
  observaciones2: 17,
  codigoQR: 18,
  urlQR: 19,
  estado: 20,
  fechaQR: 21,
  fechaExpiracion: 22,
  fechaIngreso: 23,
  intentos: 24,
  observacionesIngreso: 25
};

// ================================================================================
// API PRINCIPAL - CON SOPORTE JSONP
// ================================================================================

function doGet(e) {
  var params = e.parameter;
  var callback = params.callback;
  var resultado = {};

  try {
    Logger.log('API llamada: action=' + params.action + ', qr=' + params.qr);

    if (params.action === 'validar') {
      resultado = validarQR(params.qr);
    } else if (params.action === 'confirmar') {
      resultado = confirmarIngreso(params.qr);
    } else if (params.action === 'ping') {
      resultado = { ok: true, mensaje: 'API funcionando' };
    } else {
      resultado = { ok: false, error: 'Accion no especificada' };
    }
  } catch (error) {
    Logger.log('ERROR API: ' + error.toString());
    resultado = { ok: false, error: error.toString() };
  }

  var jsonStr = JSON.stringify(resultado);
  Logger.log('Respuesta: ' + jsonStr);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + jsonStr + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================================
// VALIDAR QR
// ================================================================================

function validarQR(codigoQR) {
  if (!codigoQR || codigoQR.trim() === '') {
    return { ok: false, estado: 'NO_EXISTE', mensaje: 'Codigo vacio' };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_RESPUESTAS);
  if (!sheet) {
    return { ok: false, estado: 'ERROR', mensaje: 'Hoja no encontrada' };
  }

  var ultimaFila = sheet.getLastRow();
  if (ultimaFila < 2) {
    return { ok: false, estado: 'NO_EXISTE', mensaje: 'Sin registros' };
  }

  var codigos = sheet.getRange(2, COLS.codigoQR, ultimaFila - 1, 1).getValues();
  var fila = -1;
  for (var i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === codigoQR.trim()) {
      fila = i + 2;
      break;
    }
  }

  if (fila === -1) {
    return { ok: false, estado: 'NO_EXISTE', mensaje: 'Codigo no valido' };
  }

  var datos = sheet.getRange(fila, 1, 1, 25).getValues()[0];
  var estado = String(datos[COLS.estado - 1]).trim();
  var fechaExp = datos[COLS.fechaExpiracion - 1];
  var intentos = parseInt(datos[COLS.intentos - 1] || 0) + 1;

  sheet.getRange(fila, COLS.intentos).setValue(intentos);

  if (intentos > 5) {
    return { ok: false, estado: 'BLOQUEADO', mensaje: 'Demasiados intentos' };
  }

  if (estado === 'Ingresado') {
    return { ok: false, estado: 'YA_USADO', mensaje: 'Codigo ya utilizado' };
  }

  if (estado === 'Cancelado') {
    return { ok: false, estado: 'CANCELADO', mensaje: 'Acceso cancelado' };
  }

  if (fechaExp && fechaExp !== '') {
    var ahora = new Date();
    var expiracion = new Date(fechaExp);
    if (ahora > expiracion) {
      sheet.getRange(fila, COLS.estado).setValue('Expirado');
      return { ok: false, estado: 'EXPIRADO', mensaje: 'Codigo expirado' };
    }
  }

  return {
    ok: true,
    estado: 'VALIDO',
    mensaje: 'Acceso autorizado',
    datos: {
      visitante: datos[COLS.nombreVisitante - 1],
      autoriza: datos[COLS.nombreAutoriza - 1],
      empresa: datos[COLS.empresaProveedor - 1],
      patente: datos[COLS.patenteVisitante - 1],
      area: datos[COLS.areaVisitar - 1],
      dia: formatearFecha(datos[COLS.diaAutorizacion - 1], 'dd/MM/yyyy')
    }
  };
}

// ================================================================================
// CONFIRMAR INGRESO
// ================================================================================

function confirmarIngreso(codigoQR) {
  try {
    Logger.log('=== CONFIRMAR INGRESO ===');
    Logger.log('QR: ' + codigoQR);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_RESPUESTAS);
    if (!sheet) {
      return { ok: false, mensaje: 'Hoja no encontrada' };
    }

    var ultimaFila = sheet.getLastRow();
    if (ultimaFila < 2) {
      return { ok: false, mensaje: 'Sin registros en la hoja' };
    }

    // Bloquear ejecucion concurrente para evitar doble confirmacion
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (lockErr) {
      return { ok: false, mensaje: 'Sistema ocupado, intente nuevamente' };
    }

    try {
      var codigos = sheet.getRange(2, COLS.codigoQR, ultimaFila - 1, 1).getValues();
      var fila = -1;

      for (var i = 0; i < codigos.length; i++) {
        if (String(codigos[i][0]).trim() === codigoQR.trim()) {
          fila = i + 2;
          break;
        }
      }

      if (fila === -1) {
        return { ok: false, mensaje: 'Codigo no encontrado' };
      }

      var datos = sheet.getRange(fila, 1, 1, 25).getValues()[0];
      var estado = String(datos[COLS.estado - 1]).trim();

      // Rechazar estados terminales con mensaje especifico
      if (estado === 'Ingresado') {
        return { ok: false, mensaje: 'Este codigo ya fue utilizado para ingresar' };
      }
      if (estado === 'Cancelado') {
        return { ok: false, mensaje: 'Este acceso fue cancelado' };
      }
      if (estado === 'Expirado') {
        return { ok: false, mensaje: 'Este codigo ha expirado' };
      }
      // 'Pendiente' o vacio (trigger no ejecutado) → permitir ingreso

      // Verificar expiracion aunque validarQR no la haya marcado
      var fechaExp = datos[COLS.fechaExpiracion - 1];
      if (fechaExp && fechaExp !== '') {
        var ahora = new Date();
        if (ahora > new Date(fechaExp)) {
          try { sheet.getRange(fila, COLS.estado).setValue('Expirado'); } catch(e) {}
          return { ok: false, mensaje: 'Este codigo ha expirado' };
        }
      }

      var ahora = new Date();

      // Actualizar hoja en try/catch separado — si falla, igual se envian los correos
      try {
        sheet.getRange(fila, COLS.estado).setValue('Ingresado');
        sheet.getRange(fila, COLS.fechaIngreso).setValue(ahora);
        Logger.log('Estado actualizado a Ingresado en fila ' + fila);
      } catch (writeErr) {
        Logger.log('WARN: Error actualizando hoja (continuando con email): ' + writeErr.toString());
      }

      // Enviar correos y capturar resultado
      var resultadoEmail = enviarEmailConfirmacionIngreso(fila, codigoQR, ahora);
      Logger.log('Resultado email: ' + JSON.stringify(resultadoEmail));

      var visitante = String(datos[COLS.nombreVisitante - 1] || '');

      return {
        ok: true,
        mensaje: 'Ingreso registrado',
        visitante: visitante,
        emailEnviado: resultadoEmail.ok,
        emailError: resultadoEmail.ok ? null : resultadoEmail.error
      };

    } finally {
      lock.releaseLock();
    }

  } catch (error) {
    Logger.log('ERROR confirmarIngreso: ' + error.toString());
    Logger.log(error.stack);
    return { ok: false, mensaje: 'Error: ' + error.toString() };
  }
}

// ================================================================================
// TRIGGER - ENVÍO DE EMAILS Y QR AL ENVIAR FORMULARIO
// ================================================================================

function alEnviarFormulario(e) {
  try {
    Logger.log('=== FORMULARIO ENVIADO ===');

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_RESPUESTAS);
    if (!sheet) {
      Logger.log('ERROR: Hoja no encontrada');
      return;
    }

    // e.range.getRow() es el unico metodo confiable con envios simultaneos
    var fila = (e && e.range) ? e.range.getRow() : sheet.getLastRow();
    Logger.log('Fila: ' + fila);

    var horaIngreso       = sheet.getRange(fila, COLS.horaIngreso).getValue();
    var horaSalida        = sheet.getRange(fila, COLS.horaSalida).getValue();
    var nombreAutoriza    = sheet.getRange(fila, COLS.nombreAutoriza).getValue();
    var nombreVisitante   = sheet.getRange(fila, COLS.nombreVisitante).getValue();
    var rutVisitante      = sheet.getRange(fila, COLS.rutVisitante).getValue();
    var emailVisitante    = sheet.getRange(fila, COLS.emailVisitante).getValue();
    var empresaProveedor  = sheet.getRange(fila, COLS.empresaProveedor).getValue();
    var patenteVisitante  = sheet.getRange(fila, COLS.patenteVisitante).getValue();
    var areaVisitar       = sheet.getRange(fila, COLS.areaVisitar).getValue();
    var diaAutorizacion   = sheet.getRange(fila, COLS.diaAutorizacion).getValue();
    var personaContacto   = sheet.getRange(fila, COLS.personaContacto).getValue();
    var observaciones     = sheet.getRange(fila, COLS.observacionesAdicionales).getValue();

    var codigoQR  = generarCodigoQR();
    var urlQR     = 'https://scmsbot.github.io/validador-qr?qr=' + encodeURIComponent(codigoQR);

    var ahora      = new Date();
    var expiracion = new Date(ahora.getTime() + CONFIG.HORAS_VALIDEZ * 60 * 60 * 1000);

    sheet.getRange(fila, COLS.codigoQR).setValue(codigoQR);
    sheet.getRange(fila, COLS.urlQR).setValue(urlQR);
    sheet.getRange(fila, COLS.estado).setValue('Pendiente');
    sheet.getRange(fila, COLS.fechaQR).setValue(ahora);
    sheet.getRange(fila, COLS.fechaExpiracion).setValue(expiracion);
    sheet.getRange(fila, COLS.intentos).setValue(0);

    Logger.log('QR generado: ' + codigoQR);

    enviarEmailSCMS(codigoQR, nombreVisitante, nombreAutoriza, empresaProveedor, patenteVisitante, areaVisitar, observaciones, diaAutorizacion, horaIngreso, horaSalida, rutVisitante);
    enviarEmailGuardias(codigoQR, nombreVisitante, empresaProveedor, patenteVisitante, areaVisitar, diaAutorizacion, horaIngreso, horaSalida, observaciones);
    enviarEmailInspeccion(codigoQR, nombreVisitante, nombreAutoriza, empresaProveedor, patenteVisitante, areaVisitar, observaciones, diaAutorizacion, horaIngreso, horaSalida, rutVisitante);
    enviarEmailVisitante(codigoQR, nombreVisitante, emailVisitante, nombreAutoriza, diaAutorizacion, horaIngreso, horaSalida, observaciones);

    Logger.log('=== FORMULARIO COMPLETADO ===');

  } catch (error) {
    Logger.log('ERROR alEnviarFormulario: ' + error.toString());
  }
}

// ================================================================================
// GENERADORES / UTILIDADES
// ================================================================================

function generarCodigoQR() {
  var ahora     = new Date();
  var timestamp = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  var random    = Math.random().toString(16).substring(2, 10).toUpperCase();
  return 'FB-' + timestamp + '-' + random;
}

// Descarga la imagen QR como blob para adjuntarla inline en el correo.
// Devuelve el blob o null si falla.
function obtenerQRBlob(codigo) {
  try {
    var url = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&margin=10&data=' +
      encodeURIComponent('https://scmsbot.github.io/validador-qr?qr=' + codigo);
    return UrlFetchApp.fetch(url).getBlob().setName('qr.png');
  } catch (e) {
    Logger.log('Error generando QR blob: ' + e.toString());
    return null;
  }
}

function formatearHora(valor) {
  if (!valor) return 'N/A';
  var str   = String(valor);
  var match = str.match(/(\d{1,2}):(\d{2}):(\d{2})/);
  if (match) return match[1] + ':' + match[2] + ':' + match[3];
  return str;
}

function formatearFecha(fecha, formato) {
  if (!fecha || fecha === '') return '';
  if (typeof fecha === 'string') fecha = new Date(fecha);
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), formato || 'dd/MM/yyyy HH:mm');
}

// ================================================================================
// EMAILS - REGISTRO INICIAL (al enviar formulario)
// ================================================================================

function enviarEmailSCMS(codigo, visitante, autoriza, empresa, patente, area, obs, dia, horaIni, horaFin, rut) {
  try {
    var dest    = CONFIG.EMAILS.scms.join(',');
    var asunto  = 'Nuevo acceso: ' + visitante + ' - ' + empresa;
    var horario = formatearHora(horaIni) + ' a ' + formatearHora(horaFin);
    var estiloCode = 'font-family:monospace;font-size:14px;background:#fff;padding:8px;border-radius:4px;display:block;margin-top:8px;word-break:break-all';
    var codigoHtml = '<code style="' + estiloCode + '">' + codigo + '</code>';
    var qrBlob = obtenerQRBlob(codigo);
    var imgTag = qrBlob
      ? '<div style="text-align:center;margin:16px 0"><img src="cid:qrCode" alt="QR" style="width:200px;height:200px;border:2px solid #0078D4;border-radius:6px"></div>'
      : '';

    var html = _htmlBase('#0078D4', 'Nuevo Registro de Visitante', 'Hospital Felix Bulnes - Sistema SCMS',
      _tabla([
        ['Visitante', '<strong>' + visitante + '</strong>'],
        ['RUT', rut || 'N/A'],
        ['Autoriza', autoriza],
        ['Empresa', empresa],
        ['Patente', patente || 'N/A'],
        ['Area', area],
        ['Dia', formatearFecha(dia, 'dd/MM/yyyy')],
        ['Horario', horario]
      ]) +
      _bloque('#FFF3CD', '#FFC107', 'Observaciones', obs || 'Sin observaciones') +
      _bloque('#E1F5FE', '#0078D4', 'Codigo de acceso', imgTag + codigoHtml)
    );

    var opciones = { htmlBody: html, name: CONFIG.EMPRESA };
    if (qrBlob) opciones.inlineImages = { qrCode: qrBlob };
    GmailApp.sendEmail(dest, asunto, '', opciones);
    Logger.log('Email SCMS OK');
  } catch (err) {
    Logger.log('ERROR email SCMS: ' + err.toString());
  }
}

function enviarEmailGuardias(codigo, visitante, empresa, patente, area, dia, horaIni, horaFin, obs) {
  try {
    var dest    = CONFIG.EMAILS.guardias.join(',');
    var asunto  = 'Visitante esperado: ' + visitante;
    var horario = formatearHora(horaIni) + ' a ' + formatearHora(horaFin);
    var qrBlob  = obtenerQRBlob(codigo);
    var imgTag  = qrBlob
      ? '<div style="text-align:center;margin:16px 0"><img src="cid:qrCode" alt="QR" style="width:200px;height:200px;border:2px solid #107C10;border-radius:6px"></div>'
      : '';

    var html = _htmlBase('#107C10', 'Nuevo Visitante Esperado', 'Informacion de ingreso en recepcion',
      imgTag +
      _tabla([
        ['Visitante', '<strong>' + visitante + '</strong>'],
        ['Empresa',   empresa],
        ['Patente',   '<strong>' + (patente || 'Sin vehiculo') + '</strong>'],
        ['Destino',   area],
        ['Fecha',     formatearFecha(dia, 'dd/MM/yyyy')],
        ['Horario',   horario]
      ]) +
      _bloque('#FFF3CD', '#FFC107', 'Observaciones', obs || 'Sin observaciones')
    );

    var opciones = { htmlBody: html, name: CONFIG.EMPRESA };
    if (qrBlob) opciones.inlineImages = { qrCode: qrBlob };
    GmailApp.sendEmail(dest, asunto, '', opciones);
    Logger.log('Email Guardias OK');
  } catch (err) {
    Logger.log('ERROR email Guardias: ' + err.toString());
  }
}

function enviarEmailInspeccion(codigo, visitante, autoriza, empresa, patente, area, obs, dia, horaIni, horaFin, rut) {
  try {
    var dest    = CONFIG.EMAILS.inspeccion.join(',');
    var asunto  = 'Acceso registrado: ' + visitante + ' - ' + empresa;
    var horario = formatearHora(horaIni) + ' a ' + formatearHora(horaFin);
    var qrBlob  = obtenerQRBlob(codigo);
    var imgTag  = qrBlob
      ? '<div style="text-align:center;margin:16px 0"><img src="cid:qrCode" alt="QR" style="width:200px;height:200px;border:2px solid #8B4513;border-radius:6px"></div>'
      : '';

    var html = _htmlBase('#8B4513', 'Registro de Acceso', 'Hospital Felix Bulnes',
      imgTag +
      _tabla([
        ['Visitante', visitante],
        ['RUT',       rut || 'N/A'],
        ['Empresa',   empresa],
        ['Autoriza',  autoriza],
        ['Vehiculo',  patente || 'N/A'],
        ['Codigo QR', '<code>' + codigo + '</code>'],
        ['Area',      area],
        ['Horario',   horario]
      ])
    );

    var opciones = { htmlBody: html, name: CONFIG.EMPRESA };
    if (qrBlob) opciones.inlineImages = { qrCode: qrBlob };
    GmailApp.sendEmail(dest, asunto, '', opciones);
    Logger.log('Email Inspeccion OK');
  } catch (err) {
    Logger.log('ERROR email Inspeccion: ' + err.toString());
  }
}

function enviarEmailVisitante(codigo, visitante, email, autoriza, dia, horaIni, horaFin, obs) {
  if (!email || email.trim() === '') {
    Logger.log('Sin email visitante — omitiendo');
    return;
  }
  try {
    var asunto  = 'Tu acceso autorizado a ' + CONFIG.EMPRESA;
    var horario = formatearHora(horaIni) + ' a ' + formatearHora(horaFin);

    // Intentar obtener QR como imagen inline
    var qrBlob = obtenerQRBlob(codigo);
    var imgTag = qrBlob
      ? '<img src="cid:qrCode" alt="Codigo QR" style="width:280px;height:280px;border:3px solid #0078D4;border-radius:8px">'
      : '<p style="color:#666;font-size:13px">(Imagen QR no disponible)</p>';

    var estiloCode2 = 'font-family:monospace;font-size:15px;background:#f0f0f0;padding:12px;border-radius:4px;display:block;word-break:break-all;border:1px solid #ddd';
    var html = _htmlBase('#0078D4', 'Acceso Autorizado', CONFIG.EMPRESA,
      '<p style="font-size:16px">Hola <strong>' + visitante + '</strong></p>' +
      '<p>Tu acceso ha sido autorizado por <strong>' + autoriza + '</strong>. Presenta este codigo QR en recepcion para ingresar.</p>' +
      '<div style="text-align:center;margin:30px 0;background:#f9f9f9;padding:24px;border-radius:8px">' + imgTag + '</div>' +
      '<div style="text-align:center;margin:16px 0">' +
      '<p style="font-size:12px;color:#666;margin-bottom:8px">Codigo de respaldo (si la imagen no carga):</p>' +
      '<code style="' + estiloCode2 + '">' + codigo + '</code></div>' +
      _tabla([
        ['Fecha',          formatearFecha(dia, 'dd/MM/yyyy')],
        ['Horario',        horario],
        ['Autorizado por', autoriza]
      ]) +
      (obs ? _bloque('#FFF3CD', '#FFC107', 'Informacion adicional', obs) : '') +
      _bloque('#E8F5E9', '#107C10', 'Instrucciones',
        '1. Llega a recepcion con este correo abierto<br>' +
        '2. Muestra el codigo QR al guardia<br>' +
        '3. El sistema validara tu acceso automaticamente') +
      '<p style="color:#666;font-size:12px;margin-top:20px;text-align:center">Este codigo expira en ' + CONFIG.HORAS_VALIDEZ + ' horas.<br>' + CONFIG.EMPRESA + '</p>'
    );

    var opciones = { htmlBody: html, name: CONFIG.EMPRESA };
    if (qrBlob) opciones.inlineImages = { qrCode: qrBlob };

    GmailApp.sendEmail(email.trim(), asunto, '', opciones);
    Logger.log('Email Visitante OK' + (qrBlob ? ' (con QR inline)' : ' (sin QR)'));
  } catch (err) {
    Logger.log('ERROR email Visitante: ' + err.toString());
  }
}

// ================================================================================
// EMAIL: CONFIRMACIÓN DE INGRESO
// Retorna { ok: true } o { ok: false, error: '...' }
// ================================================================================

function enviarEmailConfirmacionIngreso(fila, codigo, fechaIngreso) {
  var errores = [];

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_RESPUESTAS);
    if (!sheet) {
      return { ok: false, error: 'Hoja no encontrada: ' + CONFIG.HOJA_RESPUESTAS };
    }
    var datos = sheet.getRange(fila, 1, 1, 25).getValues()[0];

    var visitante      = String(datos[COLS.nombreVisitante - 1]  || 'N/A').trim();
    var autoriza       = String(datos[COLS.nombreAutoriza - 1]   || 'N/A').trim();
    var empresa        = String(datos[COLS.empresaProveedor - 1] || 'N/A').trim();
    var area           = String(datos[COLS.areaVisitar - 1]      || 'N/A').trim();
    var emailVisitante = String(datos[COLS.emailVisitante - 1]   || '').trim();
    var horarioIngreso = formatearFecha(fechaIngreso, 'dd/MM/yyyy HH:mm:ss');

    Logger.log('Datos confirmacion - Visitante: ' + visitante + ', Empresa: ' + empresa);

    var tablaBase = _tabla([
      ['Visitante',         '<strong>' + visitante + '</strong>'],
      ['Empresa',           empresa],
      ['Area',              area],
      ['Autorizado por',    autoriza],
      ['Fecha/Hora Ingreso','<strong>' + horarioIngreso + '</strong>'],
      ['Codigo QR',         '<code>' + codigo + '</code>']
    ]);
    var bloqueInhabilitado = _bloque('#E8F5E9', '#107C10', 'Estado', 'El codigo ha sido INHABILITADO y no puede usarse nuevamente.');

    // EMAIL 1: SCMS
    try {
      var htmlSCMS = _htmlBase('#107C10', 'Ingreso Confirmado', 'Hospital Felix Bulnes - Sistema SCMS',
        '<p style="font-size:16px;color:#107C10"><strong>El visitante ha ingresado correctamente.</strong></p>' +
        tablaBase + bloqueInhabilitado
      );
      GmailApp.sendEmail(CONFIG.EMAILS.scms.join(','), 'CONFIRMADO - Ingreso de ' + visitante, '', { htmlBody: htmlSCMS, name: CONFIG.EMPRESA });
      Logger.log('Email SCMS confirmacion OK');
    } catch (err) {
      Logger.log('ERROR email SCMS confirmacion: ' + err.toString());
      errores.push('SCMS: ' + err.message);
    }

    // EMAIL 2: GUARDIAS
    try {
      var htmlGuardias = _htmlBase('#107C10', 'Visitante Ingresado', 'Confirmacion de acceso',
        '<p style="font-size:16px"><strong>' + visitante + '</strong> ha ingresado correctamente.</p>' +
        _tabla([
          ['Visitante',   '<strong>' + visitante + '</strong>'],
          ['Empresa',     empresa],
          ['Area',        area],
          ['Hora Ingreso','<strong>' + horarioIngreso + '</strong>']
        ]) +
        _bloque('#E8F5E9', '#107C10', 'Estado', 'El codigo QR ya no es valido.')
      );
      GmailApp.sendEmail(CONFIG.EMAILS.guardias.join(','), 'INGRESADO - ' + visitante, '', { htmlBody: htmlGuardias, name: CONFIG.EMPRESA });
      Logger.log('Email Guardias confirmacion OK');
    } catch (err) {
      Logger.log('ERROR email Guardias confirmacion: ' + err.toString());
      errores.push('Guardias: ' + err.message);
    }

    // EMAIL 3: INSPECCIÓN
    try {
      var htmlInspeccion = _htmlBase('#107C10', 'Ingreso Registrado', 'Hospital Felix Bulnes',
        tablaBase +
        _bloque('#E8F5E9', '#107C10', 'Estado', 'Codigo INHABILITADO - No puede reutilizarse.')
      );
      GmailApp.sendEmail(CONFIG.EMAILS.inspeccion.join(','), 'CONFIRMADO - Ingreso de ' + visitante + ' - ' + empresa, '', { htmlBody: htmlInspeccion, name: CONFIG.EMPRESA });
      Logger.log('Email Inspeccion confirmacion OK');
    } catch (err) {
      Logger.log('ERROR email Inspeccion confirmacion: ' + err.toString());
      errores.push('Inspeccion: ' + err.message);
    }

    // EMAIL 4: VISITANTE (solo si tiene email)
    if (emailVisitante !== '') {
      try {
        var htmlVisitante = _htmlBase('#107C10', 'Ingreso Confirmado', CONFIG.EMPRESA,
          '<p style="font-size:16px">Hola <strong>' + visitante + '</strong>,</p>' +
          '<p>Tu ingreso ha sido confirmado correctamente.</p>' +
          _tabla([
            ['Empresa',       empresa],
            ['Area',          area],
            ['Autorizado por',autoriza],
            ['Hora Ingreso',  '<strong>' + horarioIngreso + '</strong>']
          ]) +
          _bloque('#E8F5E9', '#107C10', 'Estado', 'Tu codigo QR ha sido inhabilitado tras confirmar tu ingreso.') +
          '<p style="color:#666;font-size:12px;margin-top:20px;text-align:center">' + CONFIG.EMPRESA + ' - Sistema de Control de Acceso</p>'
        );
        GmailApp.sendEmail(emailVisitante, 'Ingreso Confirmado - ' + CONFIG.EMPRESA, '', { htmlBody: htmlVisitante, name: CONFIG.EMPRESA });
        Logger.log('Email Visitante confirmacion OK');
      } catch (err) {
        Logger.log('ERROR email Visitante confirmacion: ' + err.toString());
        errores.push('Visitante: ' + err.message);
      }
    }

  } catch (err) {
    Logger.log('ERROR CRITICO en enviarEmailConfirmacionIngreso: ' + err.toString());
    return { ok: false, error: 'Error leyendo datos: ' + err.message };
  }

  if (errores.length === 0) {
    return { ok: true };
  }
  return { ok: false, error: errores.join(' | ') };
}

// ================================================================================
// HELPERS HTML (evitan repetición)
// ================================================================================

function _htmlBase(color, titulo, subtitulo, cuerpo) {
  return '<div style="font-family:Arial;max-width:600px;margin:0 auto">' +
    '<div style="background:' + color + ';color:white;padding:25px;text-align:center">' +
    '<h2 style="margin:0">' + titulo + '</h2>' +
    '<p style="margin:5px 0 0 0">' + subtitulo + '</p></div>' +
    '<div style="padding:25px">' + cuerpo + '</div></div>';
}

function _tabla(filas) {
  var html = '<table style="width:100%;border-collapse:collapse;margin:20px 0">';
  for (var i = 0; i < filas.length; i++) {
    var bg = (i % 2 === 0) ? '#f3f2f1' : '#ffffff';
    html += '<tr style="background:' + bg + '">' +
      '<td style="padding:12px;font-weight:bold;border:1px solid #ddd;width:40%">' + filas[i][0] + '</td>' +
      '<td style="padding:12px;border:1px solid #ddd">' + filas[i][1] + '</td></tr>';
  }
  return html + '</table>';
}

function _bloque(bgColor, borderColor, titulo, contenido) {
  return '<div style="background:' + bgColor + ';padding:15px;border-left:4px solid ' + borderColor + ';border-radius:4px;margin:20px 0">' +
    '<strong>' + titulo + ':</strong><br><span style="font-size:14px;color:#333">' + contenido + '</span></div>';
}

// ================================================================================
// INICIALIZACIÓN
// ================================================================================

function probarEmailConfirmacion() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_RESPUESTAS);
  var ultimaFila = sheet.getLastRow();
  if (ultimaFila < 2) {
    Logger.log('No hay filas con datos para probar');
    return;
  }
  Logger.log('Probando con fila: ' + ultimaFila);
  var resultado = enviarEmailConfirmacionIngreso(ultimaFila, 'TEST-CODIGO-PRUEBA', new Date());
  Logger.log('Resultado: ' + JSON.stringify(resultado));
}

function diagnosticar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = ss.getSheets();
  var nombres = hojas.map(function(h) { return '"' + h.getName() + '"'; });
  Logger.log('Hojas disponibles: ' + nombres.join(', '));
  Logger.log('Hoja configurada: "' + CONFIG.HOJA_RESPUESTAS + '"');
  var sheet = ss.getSheetByName(CONFIG.HOJA_RESPUESTAS);
  Logger.log('Sheet encontrada: ' + (sheet ? 'SI' : 'NO - nombre incorrecto'));
}

function crearTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'alEnviarFormulario') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('alEnviarFormulario')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  SpreadsheetApp.getUi().alert('Trigger configurado correctamente');
}
