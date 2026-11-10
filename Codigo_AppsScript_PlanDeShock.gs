/**
 * RENTA MAQ · Base de datos de verificaciones – Plan de Shock (Teck Carmen de Andacollo)
 * Recibe los datos del formulario HTML (con o sin red, el formulario reenvía al volver la conexión)
 * y los guarda en la planilla "Base de Datos Verificaciones - Plan de Shock"
 * dentro de Mi unidad / Andacollo / Plan de Shock. Las fotos y PDF se guardan en la subcarpeta "Evidencias".
 */
const SHEET_ID  = '14wDcdO-0AC-W9aG6CIHjnkaY2wWVYzD2lVDiIpLbDOg';   // planilla Base de Datos
const FOLDER_ID = '1r4OrIFNLnlWp12oie-nrDCLEvgbj1HI9';              // carpeta Andacollo / Plan de Shock

const HOJAS = {
  'Registros':  ['ID Registro','Fecha','Quien inspecciona','Área / Lugar','Supervisor evaluado','SÍ','NO','N/A','% Cumplimiento','Observaciones generales','Creado en dispositivo','Recibido en Drive'],
  'Respuestas': ['ID Registro','Fecha','Quien inspecciona','Área / Lugar','Supervisor evaluado','Sección','N°','Pregunta','Respuesta','Comentario','Evidencias'],
  'Hallazgos':  ['ID Hallazgo','ID Registro','Fecha','Quien inspecciona','Área / Lugar','Supervisor evaluado','Pregunta','Comentario','Prioridad','Responsable','Plazo','Estado','Medida / evidencia de cierre','Fecha cierre','Evidencias hallazgo','Evidencias cierre','Última actualización'],
  'Evidencias': ['ID Evidencia','Nombre','Tipo','Enlace','Recibido en Drive']
};

/** Ejecutar UNA vez desde el editor (botón ▶ Ejecutar) para crear las hojas y autorizar el acceso. */
function setup() {
  const ss = hojas_();
  carpetaEvidencias_();
  ss.getSheetByName('Registros').getRange('I:I').setNumberFormat('0%');
  Logger.log('Listo: ' + ss.getUrl());
}

/** GET: prueba de conexión y lectura de hallazgos para todos los equipos (admite JSONP con ?callback=). */
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out = { ok: true, app: 'RentaMaq Plan de Shock' };
  try {
    if (p.accion === 'hallazgos') out = { ok: true, hallazgos: leerHallazgos_() };
  } catch (err) { out = { ok: false, error: String(err) }; }
  const s = JSON.stringify(out);
  if (p.callback && /^[A-Za-z_$][\w$]*$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + s + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

function fecha_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v || ''); }

function leerHallazgos_() {
  const sh = hojas_().getSheetByName('Hallazgos'), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 16).getValues().filter(function (r) { return r[0]; }).map(function (r) {
    return { id: String(r[0]), registro: String(r[1]), fecha: fecha_(r[2]), verificador: String(r[3]), area: String(r[4]), lider: String(r[5]),
      pregunta: String(r[6]), comentario: String(r[7]), prioridad: String(r[8]), responsable: String(r[9]), plazo: fecha_(r[10]),
      estado: String(r[11]), cierre: String(r[12]), fechaCierre: fecha_(r[13]), enlaces: String(r[14] || ''), enlacesCierre: String(r[15] || '') };
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const m = JSON.parse(e.postData.contents);
    const ss = hojas_();
    if (m.tipo === 'ping')      return json_({ ok: true });
    if (m.tipo === 'evidencia') return json_(guardarEvidencia_(ss, m));
    if (m.tipo === 'registro')  return json_(guardarRegistro_(ss, m.registro));
    if (m.tipo === 'hallazgo')  return json_(guardarHallazgo_(ss, m.hallazgo));
    return json_({ ok: false, error: 'Tipo desconocido' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- utilidades ---------------- */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function hojas_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Object.keys(HOJAS).forEach(function (n) {
    let sh = ss.getSheetByName(n);
    if (!sh) sh = ss.insertSheet(n);
    if (sh.getLastRow() === 0) {
      const h = HOJAS[n];
      sh.appendRow(h);
      sh.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#3F3F3F').setFontColor('#FFFFFF');
      sh.getRange(1, 1, 1, h.length).setBorder(null, null, true, null, null, null, '#E88305', SpreadsheetApp.BorderStyle.SOLID_THICK);
      sh.setFrozenRows(1);
    }
  });
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  return ss;
}

function carpetaEvidencias_() {
  const p = DriveApp.getFolderById(FOLDER_ID);
  const it = p.getFoldersByName('Evidencias');
  return it.hasNext() ? it.next() : p.createFolder('Evidencias');
}

function filaDe_(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const v = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === String(id)) return i + 2;
  return 0;
}

function mapaEvidencias_(ss) {
  const sh = ss.getSheetByName('Evidencias'), n = sh.getLastRow(), m = {};
  if (n > 1) sh.getRange(2, 1, n - 1, 4).getValues().forEach(function (r) { m[r[0]] = r[3]; });
  return m;
}
function enlaces_(m, lista) {
  return (lista || []).map(function (v) { return m[v.id] || ('(pendiente) ' + v.name); }).join('\n');
}

/* ---------------- guardado ---------------- */
function guardarEvidencia_(ss, m) {
  const sh = ss.getSheetByName('Evidencias');
  if (filaDe_(sh, m.id)) return { ok: true, duplicado: true };
  const blob = Utilities.newBlob(Utilities.base64Decode(m.data.split(',')[1]), m.type, m.id + '_' + m.name);
  const f = carpetaEvidencias_().createFile(blob);
  sh.appendRow([m.id, m.name, m.type, f.getUrl(), new Date()]);
  return { ok: true, url: f.getUrl() };
}

function guardarRegistro_(ss, r) {
  const reg = ss.getSheetByName('Registros');
  if (filaDe_(reg, r.id)) return { ok: true, duplicado: true };
  const d = r.datos, m = mapaEvidencias_(ss);
  const c = function (v) { return r.respuestas.filter(function (x) { return x.valor === v; }).length; };
  const si = c('SI'), no = c('NO'), na = c('NA');
  reg.appendRow([r.id, d.fecha, d.verificador, d.area, d.lider, si, no, na, (si + no) ? si / (si + no) : '', r.obs || '', r.creado, new Date()]);
  const et = { SI: 'SÍ', NO: 'NO', NA: 'N/A' };
  const filas = r.respuestas.map(function (x) {
    return [r.id, d.fecha, d.verificador, d.area, d.lider, x.seccion || '', x.n, x.pregunta, et[x.valor] || 'Sin respuesta', x.comentario || '', enlaces_(m, x.evid)];
  });
  const sh = ss.getSheetByName('Respuestas');
  if (filas.length) sh.getRange(sh.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
  return { ok: true };
}

function guardarHallazgo_(ss, h) {
  const sh = ss.getSheetByName('Hallazgos'), m = mapaEvidencias_(ss);
  const r = filaDe_(sh, h.id);
  // Si el hallazgo se edita desde otro equipo (sin las fotos locales), se conservan los enlaces ya guardados.
  const previo = r ? sh.getRange(r, 15, 1, 2).getValues()[0] : ['', ''];
  const ev = enlaces_(m, h.evid) || previo[0], evc = enlaces_(m, h.evidCierre) || previo[1];
  const fila = [h.id, h.registro, h.fecha, h.verificador, h.area, h.lider, h.pregunta, h.comentario, h.prioridad, h.responsable, h.plazo, h.estado, h.cierre, h.fechaCierre, ev, evc, new Date()];
  if (r) sh.getRange(r, 1, 1, fila.length).setValues([fila]); else sh.appendRow(fila);
  const cel = sh.getRange(r || sh.getLastRow(), 12);
  cel.setBackground(h.estado === 'CERRADO' ? '#E6F3E7' : '#FDF1E2').setFontColor(h.estado === 'CERRADO' ? '#2E7D32' : '#C76E00').setFontWeight('bold');
  return { ok: true };
}
