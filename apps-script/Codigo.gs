/**
 * ============================================================================
 *  UNIDAD DE GESTIÓN MÉDICA · IMAGENOLOGÍA
 *  Nuevo Hospital Rosales · Red Nacional de Hospitales
 * ----------------------------------------------------------------------------
 *  Backend (Google Apps Script) del sistema de control de asistencia de
 *  pacientes a estudios de imagenología.
 *
 *  Vistas:
 *    /exec                 → Técnicos (control directo del paciente)
 *    /exec?vista=supervisor → Supervisores (radar en vivo, reportes, admin)
 *
 *  La hoja "Asistencia_Diaria" contiene la agenda del turno actual.
 *  La hoja "Historial" acumula todos los turnos cerrados (una fila por cita),
 *  lo que permite sacar reportes por rango de fechas.
 * ============================================================================
 */

var CONFIG = {
  APP_NOMBRE: 'Imagenología · UGM',
  HOJA_AGENDA: 'Asistencia_Diaria',
  HOJA_HISTORIAL: 'Historial',
  CARPETA_REPORTES: 'Reportes Imagenología UGM',
  ZONA_HORARIA: 'America/El_Salvador'
};

// Orden de columnas de la agenda. NO cambiar el orden sin migrar las hojas.
var ENCABEZADOS = [
  'ID', 'Expediente', 'Paciente', 'Ubicación', 'Estudio', 'Equipo', 'Modalidad',
  'Fecha cita', 'Hora cita', 'Médico solicitante', 'Estado', 'Motivo', 'Comentario',
  'Técnico', 'Registrado', 'Cargado'
];
var ENCABEZADOS_HISTORIAL = ENCABEZADOS.concat(['Turno cerrado']);

// Índices (base 0) de cada columna dentro de ENCABEZADOS.
var C = {
  ID: 0, EXPEDIENTE: 1, PACIENTE: 2, UBICACION: 3, ESTUDIO: 4, EQUIPO: 5, MODALIDAD: 6,
  FECHA: 7, HORA: 8, MEDICO: 9, ESTADO: 10, MOTIVO: 11, COMENTARIO: 12,
  TECNICO: 13, REGISTRADO: 14, CARGADO: 15, TURNO_CERRADO: 16
};

var ESTADOS = { PENDIENTE: 'Pendiente', ASISTIO: 'Asistió', AUSENTE: 'Ausente' };

// ─────────────────────────────────────────────────────────────────────────────
//  ENTRADA WEB
// ─────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  asegurarHojas_();
  var vista = (e && e.parameter && e.parameter.vista) || 'tecnico';
  var archivo = vista === 'supervisor' ? 'Supervisor' : 'Tecnico';
  var titulo = vista === 'supervisor' ? 'Radar Imagenología · UGM' : 'Control de Pacientes · Imagenología';
  return HtmlService.createTemplateFromFile(archivo)
    .evaluate()
    .setTitle(titulo)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://www.gstatic.com/images/icons/material/system/2x/local_hospital_black_48dp.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Inserta el contenido de otro archivo HTML del proyecto (estilos, scripts, logos). */
function include(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

/** Igual que include(), pero evalúa los <?!= ?> del archivo incluido. */
function incluirPlantilla(nombre) {
  return HtmlService.createTemplateFromFile(nombre).evaluate().getContent();
}

// ─────────────────────────────────────────────────────────────────────────────
//  HOJAS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Crea las hojas si no existen. Si "Asistencia_Diaria" tiene el formato viejo
 * (sin columna ID), se renombra a "Asistencia_Diaria_ANTIGUA_<fecha>" para no
 * perder nada y se crea una nueva con el formato actual.
 */
function asegurarHojas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var agenda = ss.getSheetByName(CONFIG.HOJA_AGENDA);
  if (agenda && agenda.getLastColumn() > 0 && agenda.getRange(1, 1).getValue() !== 'ID') {
    agenda.setName(CONFIG.HOJA_AGENDA + '_ANTIGUA_' + fechaTexto_(new Date(), 'yyyyMMdd_HHmm'));
    agenda = null;
  }
  if (!agenda) agenda = crearHoja_(ss, CONFIG.HOJA_AGENDA, ENCABEZADOS);
  if (!ss.getSheetByName(CONFIG.HOJA_HISTORIAL)) crearHoja_(ss, CONFIG.HOJA_HISTORIAL, ENCABEZADOS_HISTORIAL);
  return agenda;
}

function crearHoja_(ss, nombre, encabezados) {
  var hoja = ss.insertSheet(nombre);
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
    .setBackground('#10567F').setFontColor('#FFFFFF').setFontWeight('bold');
  hoja.setFrozenRows(1);
  // Expediente/fechas/horas como texto: evita que Sheets los convierta.
  hoja.getRange('B:B').setNumberFormat('@');
  hoja.getRange('H:I').setNumberFormat('@');
  return hoja;
}

function hojaAgenda_() { return asegurarHojas_(); }

function leerFilas_(hoja, columnas) {
  var ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  return hoja.getRange(2, 1, ultima - 1, columnas).getValues();
}

// ─────────────────────────────────────────────────────────────────────────────
//  CARGA DE AGENDA
//  El Excel se lee en el navegador (SheetJS) y llega aquí ya normalizado:
//  [{expediente, paciente, ubicacion, estudio, equipo, modalidad, fecha, hora, medico}]
// ─────────────────────────────────────────────────────────────────────────────

function registrarAgenda(filas) {
  if (!filas || !filas.length) return { exito: false, mensaje: 'El archivo no tiene pacientes.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var hoja = hojaAgenda_();
    var existentes = {};
    leerFilas_(hoja, ENCABEZADOS.length).forEach(function (f) { existentes[claveCita_(f[C.EXPEDIENTE], f[C.ESTUDIO], f[C.FECHA], f[C.HORA])] = true; });

    var ahora = fechaTexto_(new Date(), 'yyyy-MM-dd HH:mm:ss');
    var nuevas = [], duplicados = 0;
    filas.forEach(function (p) {
      var expediente = limpiar_(p.expediente);
      if (!expediente) return;
      var clave = claveCita_(expediente, p.estudio, p.fecha, p.hora);
      if (existentes[clave]) { duplicados++; return; }
      existentes[clave] = true;
      var fila = new Array(ENCABEZADOS.length).fill('');
      fila[C.ID] = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
      fila[C.EXPEDIENTE] = expediente;
      fila[C.PACIENTE] = limpiar_(p.paciente).toUpperCase();
      fila[C.UBICACION] = limpiar_(p.ubicacion);
      fila[C.ESTUDIO] = limpiar_(p.estudio);
      fila[C.EQUIPO] = limpiar_(p.equipo);
      fila[C.MODALIDAD] = limpiar_(p.modalidad) || detectarModalidad_(p.equipo + ' ' + p.estudio);
      fila[C.FECHA] = limpiar_(p.fecha);
      fila[C.HORA] = limpiar_(p.hora);
      fila[C.MEDICO] = limpiar_(p.medico);
      fila[C.ESTADO] = ESTADOS.PENDIENTE;
      fila[C.CARGADO] = ahora;
      nuevas.push(fila);
    });

    if (nuevas.length) hoja.getRange(hoja.getLastRow() + 1, 1, nuevas.length, ENCABEZADOS.length).setValues(nuevas);
    var msg = 'Se cargaron ' + nuevas.length + ' pacientes a la agenda.';
    if (duplicados) msg += ' Se omitieron ' + duplicados + ' citas que ya estaban cargadas.';
    return { exito: true, mensaje: msg, insertados: nuevas.length, duplicados: duplicados };
  } catch (err) {
    return { exito: false, mensaje: 'No se pudo cargar la agenda: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function claveCita_(expediente, estudio, fecha, hora) {
  return [expediente, estudio, fecha, hora].map(function (v) { return limpiar_(v).toLowerCase(); }).join('|');
}

/** Clasifica el estudio/equipo en una modalidad de imagen. */
function detectarModalidad_(texto) {
  var t = quitarAcentos_(String(texto || '')).toUpperCase();
  var reglas = [
    ['RM', /RESONANCIA|\bRMN?\b|\bMRI\b|\bIRM\b/],
    ['TAC', /TOMOGRAF|\bTAC\b|\bTC\b|\bCT\b|ANGIOTAC|UROTAC/],
    ['MAMOGRAFÍA', /MAMOGRAF|\bMAMO\b|TOMOSINTESIS/],
    ['DENSITOMETRÍA', /DENSITOMETR|\bDMO\b|\bDXA\b/],
    ['FLUOROSCOPÍA', /FLUORO|ESOFAGOGRAMA|SERIE GASTRO|TRANSITO INTESTINAL|COLON POR ENEMA|UROGRAF|CISTOGRAF|HISTEROSALPING|FISTULOGRAF|DEFECOGRAF/],
    ['INTERVENCIONISMO', /ANGIOGRAF|INTERVENC|EMBOLIZ|ARTERIOGRAF|FLEBOGRAF|BIOPSIA|DRENAJE/],
    ['ULTRASONIDO', /ULTRASON|ULTRASONOGRAF|\bUSG?\b|ECOGRAF|\bECO\b|DOPPLER|SONOGRAF/],
    ['RAYOS X', /RAYOS|\bRX\b|RADIOGRAF|\bRX\.|PLACA|PORTATIL/]
  ];
  for (var i = 0; i < reglas.length; i++) if (reglas[i][1].test(t)) return reglas[i][0];
  return 'OTROS';
}

// ─────────────────────────────────────────────────────────────────────────────
//  CONSULTAS
// ─────────────────────────────────────────────────────────────────────────────

/** Agenda completa del turno actual (todas las citas con su estado). */
function obtenerAgenda() {
  return leerFilas_(hojaAgenda_(), ENCABEZADOS.length)
    .filter(function (f) { return f[C.ID]; })
    .map(filaAObjeto_);
}

function filaAObjeto_(f) {
  return {
    id: String(f[C.ID]),
    expediente: String(f[C.EXPEDIENTE]),
    paciente: String(f[C.PACIENTE]),
    ubicacion: String(f[C.UBICACION]),
    estudio: String(f[C.ESTUDIO]),
    equipo: String(f[C.EQUIPO]),
    modalidad: String(f[C.MODALIDAD] || 'OTROS'),
    fecha: valorTexto_(f[C.FECHA], 'dd/MM/yyyy'),
    hora: valorTexto_(f[C.HORA], 'HH:mm'),
    medico: String(f[C.MEDICO]),
    estado: String(f[C.ESTADO] || ESTADOS.PENDIENTE),
    motivo: String(f[C.MOTIVO]),
    comentario: String(f[C.COMENTARIO]),
    tecnico: String(f[C.TECNICO]),
    registrado: valorTexto_(f[C.REGISTRADO], 'yyyy-MM-dd HH:mm:ss'),
    turno: f.length > C.TURNO_CERRADO ? valorTexto_(f[C.TURNO_CERRADO], 'yyyy-MM-dd HH:mm') : ''
  };
}

// ─────────────────────────────────────────────────────────────────────────────
//  REGISTRO DE ASISTENCIA
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Guarda el resultado de una cita. Si otra persona ya la registró se devuelve
 * {conflicto:true} y no se sobrescribe.
 */
function guardarAsistencia(id, estado, motivo, comentario, tecnico) {
  if ([ESTADOS.ASISTIO, ESTADOS.AUSENTE].indexOf(estado) === -1) return { exito: false, mensaje: 'Estado inválido.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var hoja = hojaAgenda_();
    var fila = buscarFila_(hoja, id);
    if (!fila) return { exito: false, mensaje: 'La cita ya no está en la agenda.' };
    var actual = hoja.getRange(fila, C.ESTADO + 1, 1, 4).getValues()[0];
    if (actual[0] && actual[0] !== ESTADOS.PENDIENTE && actual[3] && actual[3] !== tecnico) {
      return { exito: false, conflicto: true, mensaje: 'Ya fue registrada por ' + actual[3] + '.' };
    }
    var registrado = fechaTexto_(new Date(), 'yyyy-MM-dd HH:mm:ss');
    hoja.getRange(fila, C.ESTADO + 1, 1, 5).setValues([[
      estado,
      estado === ESTADOS.AUSENTE ? limpiar_(motivo) : '',
      limpiar_(comentario),
      limpiar_(tecnico),
      registrado
    ]]);
    return { exito: true, registrado: registrado };
  } finally {
    lock.releaseLock();
  }
}

/** Devuelve una cita a "Pendiente" (auditoría del supervisor o "deshacer" del técnico). */
function restaurarPaciente(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var hoja = hojaAgenda_();
    var fila = buscarFila_(hoja, id);
    if (!fila) return false;
    hoja.getRange(fila, C.ESTADO + 1, 1, 5).setValues([[ESTADOS.PENDIENTE, '', '', '', '']]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function buscarFila_(hoja, id) {
  var ultima = hoja.getLastRow();
  if (ultima < 2) return 0;
  var ids = hoja.getRange(2, 1, ultima - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

// ─────────────────────────────────────────────────────────────────────────────
//  CIERRE DE TURNO
// ─────────────────────────────────────────────────────────────────────────────

/** Mueve toda la agenda al "Historial" (con la fecha de cierre) y la limpia. */
function cerrarTurno() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = hojaAgenda_();
    var filas = leerFilas_(hoja, ENCABEZADOS.length).filter(function (f) { return f[C.ID]; });
    if (!filas.length) return { exito: false, mensaje: 'No hay pacientes en la agenda.' };
    var cierre = fechaTexto_(new Date(), 'yyyy-MM-dd HH:mm');
    var historial = ss.getSheetByName(CONFIG.HOJA_HISTORIAL);
    var datos = filas.map(function (f) { return f.concat([cierre]); });
    historial.getRange(historial.getLastRow() + 1, 1, datos.length, ENCABEZADOS_HISTORIAL.length).setValues(datos);
    hoja.getRange(2, 1, hoja.getLastRow() - 1, ENCABEZADOS.length).clearContent();
    var asistio = filas.filter(function (f) { return f[C.ESTADO] === ESTADOS.ASISTIO; }).length;
    var ausente = filas.filter(function (f) { return f[C.ESTADO] === ESTADOS.AUSENTE; }).length;
    return {
      exito: true,
      mensaje: 'Turno archivado (' + filas.length + ' citas: ' + asistio + ' asistieron, ' + ausente + ' no asistieron).'
    };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  REPORTES
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Datos para el reporte.
 * filtros.origen: 'actual' (turno en curso) | 'historial' (turnos cerrados)
 * filtros.desde / filtros.hasta: 'yyyy-MM-dd' (solo historial, por fecha de cierre)
 */
function obtenerDatosReporte(filtros) {
  filtros = filtros || {};
  if (filtros.origen !== 'historial') return obtenerAgenda();
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_HISTORIAL);
  var desde = filtros.desde || '0000-00-00', hasta = (filtros.hasta || '9999-99-99') + ' 99';
  return leerFilas_(hoja, ENCABEZADOS_HISTORIAL.length)
    .filter(function (f) { return f[C.ID]; })
    .map(filaAObjeto_)
    .filter(function (p) { return p.turno >= desde && p.turno <= hasta; });
}

/**
 * Crea un Excel con el reporte (Resumen, Asistieron, No asistieron, Pendientes)
 * en la carpeta de Drive "Reportes Imagenología UGM" y devuelve el enlace de
 * descarga .xlsx. `filas` llega ya filtrado desde el navegador.
 */
function generarReporteExcel(filas, info) {
  info = info || {};
  if (!filas || !filas.length) return { exito: false, mensaje: 'No hay datos para el reporte.' };

  var asistio = filas.filter(function (p) { return p.estado === ESTADOS.ASISTIO; });
  var ausente = filas.filter(function (p) { return p.estado === ESTADOS.AUSENTE; });
  var pendiente = filas.filter(function (p) { return p.estado !== ESTADOS.ASISTIO && p.estado !== ESTADOS.AUSENTE; });
  var procesados = asistio.length + ausente.length;

  var nombre = 'Reporte_Imagenologia_' + fechaTexto_(new Date(), 'yyyy-MM-dd_HHmm');
  var ss = SpreadsheetApp.create(nombre);
  var AZUL = '#10567F', CIAN = '#00B2E3';

  // ── Resumen ──
  var r = ss.getSheets()[0];
  r.setName('Resumen');
  r.getRange('A1:F1').merge().setValue('NUEVO HOSPITAL ROSALES · RED NACIONAL DE HOSPITALES')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(AZUL).setHorizontalAlignment('center').setFontSize(13);
  r.getRange('A2:F2').merge().setValue('Reporte de asistencia · Imagenología · ' + (info.titulo || 'Turno actual'))
    .setFontColor('#FFFFFF').setBackground(CIAN).setHorizontalAlignment('center').setFontWeight('bold');
  r.getRange('A3:F3').merge().setValue('Generado: ' + fechaTexto_(new Date(), 'dd/MM/yyyy HH:mm') + (info.generadoPor ? ' · por ' + info.generadoPor : ''))
    .setFontColor('#64748B').setHorizontalAlignment('center').setFontStyle('italic');

  var kpis = [
    ['Indicador', 'Valor'],
    ['Citas en el reporte', filas.length],
    ['Asistieron', asistio.length],
    ['No asistieron', ausente.length],
    ['Pendientes / sin registrar', pendiente.length],
    ['Tasa de asistencia (sobre registrados)', procesados ? asistio.length / procesados : 0],
    ['Tasa de inasistencia (sobre registrados)', procesados ? ausente.length / procesados : 0]
  ];
  r.getRange(5, 1, kpis.length, 2).setValues(kpis);
  r.getRange(10, 2, 2, 1).setNumberFormat('0.0%');
  estiloTabla_(r.getRange(5, 1, kpis.length, 2), AZUL);

  // Asistencia por modalidad
  var porModalidad = {};
  filas.forEach(function (p) {
    var m = p.modalidad || 'OTROS';
    porModalidad[m] = porModalidad[m] || [0, 0, 0];
    porModalidad[m][p.estado === ESTADOS.ASISTIO ? 0 : p.estado === ESTADOS.AUSENTE ? 1 : 2]++;
  });
  var tablaMod = [['Modalidad', 'Asistieron', 'No asistieron', 'Pendientes', 'Total']];
  Object.keys(porModalidad).sort().forEach(function (m) {
    var v = porModalidad[m];
    tablaMod.push([m, v[0], v[1], v[2], v[0] + v[1] + v[2]]);
  });
  var filaMod = 5 + kpis.length + 2;
  r.getRange(filaMod - 1, 1).setValue('ASISTENCIA POR MODALIDAD').setFontWeight('bold').setFontColor(AZUL);
  r.getRange(filaMod, 1, tablaMod.length, 5).setValues(tablaMod);
  estiloTabla_(r.getRange(filaMod, 1, tablaMod.length, 5), AZUL);

  // Motivos de inasistencia
  var motivos = {};
  ausente.forEach(function (p) { var m = p.motivo || 'Sin motivo'; motivos[m] = (motivos[m] || 0) + 1; });
  var tablaMot = [['Motivo de inasistencia', 'Pacientes', '% de ausentes']];
  Object.keys(motivos).sort(function (a, b) { return motivos[b] - motivos[a]; }).forEach(function (m) {
    tablaMot.push([m, motivos[m], motivos[m] / ausente.length]);
  });
  if (tablaMot.length === 1) tablaMot.push(['(sin inasistencias)', 0, 0]);
  var filaMot = filaMod + tablaMod.length + 3;
  r.getRange(filaMot - 1, 1).setValue('MOTIVOS DE INASISTENCIA').setFontWeight('bold').setFontColor('#BE123C');
  r.getRange(filaMot, 1, tablaMot.length, 3).setValues(tablaMot);
  r.getRange(filaMot + 1, 3, tablaMot.length - 1, 1).setNumberFormat('0.0%');
  estiloTabla_(r.getRange(filaMot, 1, tablaMot.length, 3), '#BE123C');

  r.setColumnWidth(1, 300);
  r.setColumnWidths(2, 5, 120);

  // Gráficos
  r.getRange('H5:I7').setValues([['Estado', 'Pacientes'], ['Asistieron', asistio.length], ['No asistieron', ausente.length]]).setFontColor('#FFFFFF');
  r.insertChart(r.newChart().setChartType(Charts.ChartType.PIE).addRange(r.getRange('H5:I7'))
    .setPosition(5, 7, 0, 0).setOption('title', 'Asistencia de pacientes').setOption('pieHole', 0.45)
    .setOption('colors', ['#10B981', '#F43F5E']).setOption('width', 420).setOption('height', 260).build());
  if (tablaMot.length > 1 && ausente.length) {
    r.insertChart(r.newChart().setChartType(Charts.ChartType.BAR).addRange(r.getRange(filaMot, 1, tablaMot.length, 2))
      .setPosition(19, 7, 0, 0).setOption('title', 'Motivos de inasistencia').setOption('legend', { position: 'none' })
      .setOption('colors', ['#F43F5E']).setOption('width', 420).setOption('height', 260).build());
  }

  // ── Detalle ──
  var columnas = ['Expediente', 'Paciente', 'Estudio', 'Modalidad', 'Equipo', 'Ubicación', 'Fecha cita', 'Hora cita', 'Técnico', 'Registrado'];
  var base = function (p) { return [p.expediente, p.paciente, p.estudio, p.modalidad, p.equipo, p.ubicacion, p.fecha, p.hora, p.tecnico, p.registrado]; };
  hojaDetalle_(ss, 'Asistieron', columnas, asistio.map(base), '#059669');
  hojaDetalle_(ss, 'No asistieron', columnas.concat(['Motivo', 'Comentario']),
    ausente.map(function (p) { return base(p).concat([p.motivo, p.comentario]); }), '#BE123C');
  hojaDetalle_(ss, 'Pendientes', columnas.slice(0, 8), pendiente.map(function (p) { return base(p).slice(0, 8); }), '#D97706');

  SpreadsheetApp.flush();
  moverACarpeta_(ss.getId());
  return {
    exito: true,
    url: 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
    urlHoja: ss.getUrl(),
    mensaje: 'Reporte generado: ' + nombre
  };
}

function hojaDetalle_(ss, nombre, columnas, datos, color) {
  var h = ss.insertSheet(nombre);
  h.getRange(1, 1, 1, columnas.length).setValues([columnas]);
  if (datos.length) h.getRange(2, 1, datos.length, columnas.length).setNumberFormat('@').setValues(datos);
  else h.getRange(2, 1).setValue('Sin registros').setFontStyle('italic').setFontColor('#94A3B8');
  estiloTabla_(h.getRange(1, 1, Math.max(datos.length, 1) + 1, columnas.length), color);
  h.setFrozenRows(1);
  h.autoResizeColumns(1, columnas.length);
}

function estiloTabla_(rango, color) {
  rango.setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  rango.offset(0, 0, 1).setBackground(color).setFontColor('#FFFFFF').setFontWeight('bold');
  if (rango.getNumRows() > 1) rango.offset(1, 0, rango.getNumRows() - 1).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
}

function moverACarpeta_(idArchivo) {
  try {
    var it = DriveApp.getFoldersByName(CONFIG.CARPETA_REPORTES);
    var carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.CARPETA_REPORTES);
    DriveApp.getFileById(idArchivo).moveTo(carpeta);
  } catch (err) {
    console.warn('No se pudo mover el reporte a la carpeta: ' + err);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  UTILIDADES
// ─────────────────────────────────────────────────────────────────────────────

function limpiar_(v) { return v === null || v === undefined ? '' : String(v).replace(/\s+/g, ' ').trim(); }

function quitarAcentos_(s) { return s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }

function fechaTexto_(fecha, formato) {
  return Utilities.formatDate(fecha, CONFIG.ZONA_HORARIA || Session.getScriptTimeZone(), formato);
}

function valorTexto_(v, formato) {
  if (v instanceof Date) return fechaTexto_(v, formato);
  return limpiar_(v);
}
