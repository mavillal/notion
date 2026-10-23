/**
 * Envío "Lo que no baja" · Google Apps Script ligado al Sheet del CRM.
 *
 * Instalación: Extensiones > Apps Script > pegar este archivo > guardar > recargar el Sheet.
 * Uso: menú "Lo que no baja" > Crear columnas > Enviar prueba > Enviar tanda.
 *
 * Reglas que el script hace cumplir:
 *  - Un correo por persona, personalizado, desde tu Gmail. Sin CC ni BCC.
 *  - Solo filas con Estado = "Por enviar", email válido y sin "No contactar".
 *  - Nunca reenvía: una vez "Mensaje enviado", la fila queda cerrada.
 *  - Sin recordatorios, sin cifras de fatales, sin preguntas.
 */

const CFG = {
  SHEET_NAME: 'Envio_LQNB',           // pestaña de envío; si no existe usa la activa
  SENDER_NAME: 'Mario Villalobos',
  REPLY_TO: '',                       // vacío = tu propio correo
  SUBJECT: 'Lo que no baja',
  URL: 'https://villalobos.ai/analisis/lo-que-no-baja/',
  CAMPAIGN: 'crm-chile',              // utm_campaign = CAMPAIGN-<tanda>
  MAX_PER_RUN: 40,                    // tope por ejecución (una tanda)
  PAUSE_MS: 4000,                     // pausa entre correos
  MAX_RUNTIME_MS: 5 * 60 * 1000,      // Apps Script corta a los 6 min
  COLS: {                             // nombres de columna (sin distinguir tildes ni mayúsculas)
    nombre: 'Nombre',
    email: 'Email',
    estado: 'Estado',
    tanda: 'Tanda',
    enviado: 'Enviado el',
    error: 'Error',
    noContactar: 'No contactar'
  },
  ST: { pending: 'Por enviar', sent: 'Mensaje enviado', error: 'Error' }
};

// ---------- Menú ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lo que no baja')
    .addItem('1. Crear columnas faltantes', 'setupColumns')
    .addItem('2. Enviar prueba a mi correo', 'sendTest')
    .addSeparator()
    .addItem('3. Enviar una tanda…', 'sendBatch')
    .addItem('Enviar solo filas seleccionadas', 'sendSelected')
    .addToUi();
}

// ---------- Acciones ----------

function setupColumns() {
  const sh = getSheet_();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(norm_);
  const missing = Object.keys(CFG.COLS)
    .map(k => CFG.COLS[k])
    .filter(name => headers.indexOf(norm_(name)) === -1);
  if (missing.length) {
    sh.getRange(1, lastCol + (sh.getLastColumn() ? 1 : 0), 1, missing.length).setValues([missing]);
  }
  SpreadsheetApp.getUi().alert(missing.length
    ? 'Columnas creadas: ' + missing.join(', ')
    : 'Todas las columnas ya existen.');
}

function sendTest() {
  const me = Session.getActiveUser().getEmail();
  if (!me) return SpreadsheetApp.getUi().alert('No pude leer tu correo.');
  deliver_(me, 'Nombre', 'prueba');
  SpreadsheetApp.getUi().alert('Prueba enviada a ' + me + '. Revísala en el celular y en el computador.');
}

function sendBatch() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Tanda', 'Número de tanda a enviar (columna "Tanda"). Déjalo vacío para tomar las primeras ' +
    CFG.MAX_PER_RUN + ' pendientes.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  run_({ tanda: r.getResponseText().trim() });
}

function sendSelected() {
  const sh = getSheet_();
  const sel = sh.getActiveRangeList();
  if (!sel) return SpreadsheetApp.getUi().alert('Selecciona filas primero.');
  const rows = {};
  sel.getRanges().forEach(rg => {
    for (let i = 0; i < rg.getNumRows(); i++) rows[rg.getRow() + i] = true;
  });
  run_({ rows: rows });
}

// ---------- Núcleo ----------

function run_(filter) {
  const ui = SpreadsheetApp.getUi();
  const sh = getSheet_();
  const col = columnMap_(sh);
  if (!col) return;

  const data = sh.getDataRange().getValues();
  const sentEmails = {};
  const queue = [];
  const seen = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;
    const email = String(row[col.email]).trim().toLowerCase();
    const estado = norm_(row[col.estado]);
    if (estado === norm_(CFG.ST.sent) && email) sentEmails[email] = true;

    if (estado !== norm_(CFG.ST.pending)) continue;
    if (filter.rows && !filter.rows[rowNum]) continue;
    if (filter.tanda && String(row[col.tanda]).trim() !== filter.tanda) continue;
    if (col.noContactar > -1 && String(row[col.noContactar]).trim() !== '') continue;
    if (!isEmail_(email)) { mark_(sh, col, rowNum, CFG.ST.error, 'Email inválido'); continue; }
    if (/^(info|contacto|contact|hello|helpdesk|help|soporte|ventas|admin|noreply|no-reply|reclutamiento|rrhh|prensa)[.\-_@]/.test(email)) {
      mark_(sh, col, rowNum, CFG.ST.error, 'Correo genérico'); continue;
    }
    if (seen[email]) continue;
    seen[email] = true;
    queue.push({ rowNum: rowNum, email: email, nombre: row[col.nombre], tanda: String(row[col.tanda]).trim() });
  }

  // Nunca repetir a quien ya recibió (mismo correo en otra fila)
  const todo = queue.filter(q => !sentEmails[q.email]).slice(0, CFG.MAX_PER_RUN);
  if (!todo.length) return ui.alert('No hay filas pendientes que cumplan el filtro.');

  const quota = MailApp.getRemainingDailyQuota();
  if (quota < todo.length) return ui.alert('Cuota diaria de Gmail insuficiente: quedan ' + quota + ', necesitas ' + todo.length + '.');

  const ok = ui.alert('Confirmar envío',
    'Se enviarán ' + todo.length + ' correos individuales desde tu Gmail.\nPrimero: ' + todo[0].email +
    '\nÚltimo: ' + todo[todo.length - 1].email + '\n\n¿Enviar?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return ui.alert('Hay otro envío en curso. Espera a que termine.');

  const t0 = Date.now();
  let sent = 0, failed = 0, stopped = false;
  try {
    for (const q of todo) {
      if (Date.now() - t0 > CFG.MAX_RUNTIME_MS) { stopped = true; break; }
      try {
        deliver_(q.email, q.nombre, q.tanda);
        mark_(sh, col, q.rowNum, CFG.ST.sent, '', new Date());
        sent++;
      } catch (e) {
        mark_(sh, col, q.rowNum, CFG.ST.error, String(e.message || e).slice(0, 200));
        failed++;
      }
      SpreadsheetApp.flush();
      Utilities.sleep(CFG.PAUSE_MS);
    }
  } finally {
    lock.releaseLock();
  }
  ui.alert('Listo.\nEnviados: ' + sent + '\nCon errores: ' + failed +
    (stopped ? '\nSe detuvo por tiempo; vuelve a ejecutar para continuar.' : ''));
}

function deliver_(email, nombre, tanda) {
  const url = CFG.URL + '?utm_source=email&utm_medium=crm&utm_campaign=' +
    encodeURIComponent(CFG.CAMPAIGN + '-' + (tanda || '0'));
  const first = firstName_(nombre);
  const opts = {
    name: CFG.SENDER_NAME,
    htmlBody: html_(first, url),
    noReply: false
  };
  if (CFG.REPLY_TO) opts.replyTo = CFG.REPLY_TO;
  const subject = tanda === 'prueba' ? '[PRUEBA] ' + CFG.SUBJECT : CFG.SUBJECT;
  GmailApp.sendEmail(email, subject, text_(first, url), opts);
}

// ---------- Contenido ----------

function text_(first, url) {
  return (first ? 'Hola ' + first + ',' : 'Hola,') + '\n\n' +
    'Hace tiempo que no hablamos y quise compartirte esta publicación, que podría ser de tu interés.\n\n' +
    'Si quieres contrastar más datos, los dejé en este repositorio:\n' + url + '\n\n' +
    'Un saludo,\n' + CFG.SENDER_NAME + '\n\n' +
    'Si prefieres no recibir más mensajes míos, respóndeme "no" y no volveré a escribirte.';
}

function html_(first, url) {
  const font = "'Helvetica Neue',Helvetica,Arial,sans-serif";
  const serif = "Georgia,'Times New Roman',serif";
  const hi = first ? 'Hola ' + esc_(first) + ',' : 'Hola,';
  return '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<meta name="color-scheme" content="light only"></head>' +
    '<body style="margin:0;padding:0;background:#FFFFFF;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#FFFFFF;">' +
    '<tr><td align="center" style="padding:56px 24px 64px;">' +
    '<table role="presentation" width="520" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:520px;">' +

    // Marca / título
    '<tr><td style="font-family:' + font + ';font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#8A94A6;padding-bottom:28px;">Análisis</td></tr>' +
    '<tr><td style="font-family:' + serif + ';font-size:34px;line-height:40px;font-weight:normal;color:#1A2A4A;padding-bottom:36px;">Lo que no baja</td></tr>' +
    '<tr><td style="border-top:1px solid #E6E9EF;font-size:0;line-height:0;height:1px;">&nbsp;</td></tr>' +

    // Cuerpo
    '<tr><td style="font-family:' + font + ';font-size:16px;line-height:26px;color:#1F2933;padding:36px 0 0;">' + hi + '</td></tr>' +
    '<tr><td style="font-family:' + font + ';font-size:16px;line-height:26px;color:#1F2933;padding:16px 0 0;">Hace tiempo que no hablamos y quise compartirte esta publicación, que podría ser de tu interés.</td></tr>' +
    '<tr><td style="font-family:' + font + ';font-size:16px;line-height:26px;color:#1F2933;padding:16px 0 0;">Si quieres contrastar más datos, los dejé en este repositorio.</td></tr>' +

    // Botón
    '<tr><td style="padding:36px 0 0;">' +
    '<a href="' + esc_(url) + '" style="display:inline-block;font-family:' + font + ';font-size:14px;line-height:14px;letter-spacing:.3px;color:#FFFFFF;background:#1A2A4A;text-decoration:none;padding:15px 28px;border-radius:2px;">Ver el análisis</a>' +
    '</td></tr>' +

    // Firma
    '<tr><td style="font-family:' + font + ';font-size:16px;line-height:26px;color:#1F2933;padding:44px 0 0;">Un saludo,<br><span style="color:#1A2A4A;">' + esc_(CFG.SENDER_NAME) + '</span></td></tr>' +

    // Pie
    '<tr><td style="font-family:' + font + ';font-size:12px;line-height:18px;color:#8A94A6;padding:56px 0 0;">Si prefieres no recibir más mensajes míos, respóndeme &laquo;no&raquo; y no volveré a escribirte.</td></tr>' +

    '</table></td></tr></table></body></html>';
}

// ---------- Utilidades ----------

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(CFG.SHEET_NAME) || ss.getActiveSheet();
}

function columnMap_(sh) {
  const headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(norm_);
  const col = {};
  const missing = [];
  Object.keys(CFG.COLS).forEach(k => {
    col[k] = headers.indexOf(norm_(CFG.COLS[k]));
    if (col[k] === -1 && k !== 'noContactar') missing.push(CFG.COLS[k]);
  });
  if (missing.length) {
    SpreadsheetApp.getUi().alert('Faltan columnas: ' + missing.join(', ') + '.\nUsa "Crear columnas faltantes" o ajusta CFG.COLS.');
    return null;
  }
  return col;
}

function mark_(sh, col, rowNum, estado, error, date) {
  sh.getRange(rowNum, col.estado + 1).setValue(estado);
  sh.getRange(rowNum, col.error + 1).setValue(error || '');
  if (date) sh.getRange(rowNum, col.enviado + 1).setValue(date);
}

function firstName_(v) {
  const t = String(v || '').trim().split(/\s+/)[0] || '';
  if (!t || t.toLowerCase() === 'nombre') return t ? 'Nombre' : '';
  return t.charAt(0).toUpperCase() + t.slice(1).toLowerCase();
}

function isEmail_(s) { return /^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]{2,}$/.test(s); }

function norm_(s) {
  return String(s == null ? '' : s).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
