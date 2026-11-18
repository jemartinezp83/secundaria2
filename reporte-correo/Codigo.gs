/**
 * Reporte de avance por correo — Apps Script para las apps Secundaria 2.° y 3.°
 * Versión 1.1 · 17 de noviembre de 2026 (crea solo el activador automático; revisión cada 15 min; función diagnostico)
 *
 * Qué hace:
 *  1) Recibe el avance que envían las apps (iPad, iPhone, Android o PC) y lo guarda en la hoja "Avance secundaria – registro".
 *  2) Cada hora revisa la configuración que el padre o tutor eligió en la app (frecuencia, días y hora) y, cuando toca,
 *     manda el correo con el resumen de cada alumno.
 * Instalación (una sola vez): ver el instructivo. En resumen: pegar este código, ejecutar "instalar", autorizar,
 * e Implementar > Nueva implementación > Aplicación web (Ejecutar como: yo · Acceso: cualquier persona) y copiar la URL /exec.
 */
const CONFIG = {
  TOKEN: 'avance-secundaria-2026',        // debe coincidir con el de la app
  DESTINATARIOS_EXTRA: '',                // correos adicionales separados por coma (por ejemplo, el de la mamá). El dueño de la cuenta siempre lo recibe.
  ZONA: 'America/Mexico_City',
  NOMBRE_HOJA: 'Avance secundaria – registro'
};
const DOW = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const PROPS = () => PropertiesService.getScriptProperties();

/* ---------- Instalación ---------- */
function instalar() {
  const ss = libro_();
  asegurarDisparador_(true);
  Logger.log('Listo. Hoja de registro: ' + ss.getUrl());
  Logger.log('Revisión automática activa (cada 15 minutos). Si aún no lo has hecho: Implementar > Nueva implementación > Aplicación web (Ejecutar como: yo; Quién tiene acceso: cualquier persona) y copia la URL que termina en /exec.');
}
/* Crea (o repara) el activador que revisa cada 15 minutos si toca enviar el correo. Devuelve true si quedó activo. */
function asegurarDisparador_(rehacer) {
  try {
    const ts = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'revisarEnvios');
    if (ts.length === 1 && !rehacer) return true;
    ts.forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('revisarEnvios').timeBased().everyMinutes(15).create();
    return true;
  } catch (e) { PROPS().setProperty('err_disparador', String(e).slice(0, 200)); return false; }
}
/* Ejecútala desde el editor para ver en el registro por qué no llega un correo. */
function diagnostico() {
  const t = ahora_(), ts = ScriptApp.getProjectTriggers().map(x => x.getHandlerFunction() + ' (' + x.getEventType() + ')');
  Logger.log('Hora de México: ' + t.txt + ' · día: ' + DOW[t.dow]);
  Logger.log('Activadores: ' + (ts.join(', ') || 'NINGUNO → ejecuta la función "instalar"'));
  Logger.log('Última revisión automática: ' + (PROPS().getProperty('ultima_revision') || 'nunca'));
  ['sec2', 'sec3'].forEach(app => {
    const c = JSON.parse(PROPS().getProperty('cfg_' + app) || 'null');
    if (!c) return Logger.log(app + ': sin configuración recibida de la app.');
    Logger.log(app + ': ' + (c.on ? 'ACTIVADO' : 'DESACTIVADO en la app') + ' · días: ' + c.days.map(d => DOW[d]).join(', ') + ' · hora: ' + c.hour + ':00 · avisar sin práctica: ' + (c.empty ? 'sí' : 'no') + ' · último correo programado: ' + (PROPS().getProperty('sent_' + app) || 'ninguno'));
  });
  Logger.log('Último error: ' + (PROPS().getProperty('err_envio') || PROPS().getProperty('err_disparador') || 'ninguno'));
}
function estado_(app) {
  const c = JSON.parse(PROPS().getProperty('cfg_' + app) || 'null') || {};
  let disp = false; try { disp = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'revisarEnvios'); } catch (e) { }
  return { disparador: disp, ultimaRevision: PROPS().getProperty('ultima_revision') || '', ultimoCorreo: PROPS().getProperty('sent_' + app) || '', activado: !!c.on, error: PROPS().getProperty('err_envio') || '' };
}
function libro_() {
  const id = PROPS().getProperty('ss');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { } }
  const ss = SpreadsheetApp.create(CONFIG.NOMBRE_HOJA);
  const d = ss.getSheets()[0]; d.setName('Dias'); d.appendRow(['clave', 'fecha', 'app', 'dispositivo', 'alumno', 'preguntas', 'aciertos', 'minutos', 'detalle', 'recibido']);
  const e = ss.insertSheet('Estado'); e.appendRow(['clave', 'app', 'appNombre', 'dispositivoId', 'dispositivo', 'alumno', 'estado', 'recibido']);
  const c = ss.insertSheet('Envios'); c.appendRow(['fecha y hora', 'app', 'tipo', 'destinatarios', 'alumnos', 'periodo']);
  [d, e, c].forEach(s => s.setFrozenRows(1));
  PROPS().setProperty('ss', ss.getId());
  return ss;
}

/* ---------- Recepción de datos ---------- */
function doGet() { return json_({ ok: true, servicio: 'Reporte de avance secundaria', hora: ahora_().txt }); }
function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'datos inválidos' }); }
  if (!data || data.token !== CONFIG.TOKEN) return json_({ ok: false, error: 'código de acceso incorrecto' });
  if (!/^sec\d$/.test(String(data.app || ''))) return json_({ ok: false, error: 'app desconocida' });
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    guardarConfig_(data);
    guardarDatos_(data);
    if (data.cfg && data.cfg.on) asegurarDisparador_(false);
    let msg = 'sincronizado';
    if (data.action === 'test') {
      const n = Number(PROPS().getProperty('test_' + hoy_()) || 0);
      if (n >= 10) return json_({ ok: false, error: 'límite de 10 correos de prueba por día' });
      PROPS().setProperty('test_' + hoy_(), String(n + 1));
      enviarReporte_(data.app, true); msg = 'correo de prueba enviado';
    }
    return json_({ ok: true, msg, estado: estado_(data.app) });
  } finally { lock.releaseLock(); }
}
function guardarConfig_(data) {
  const c = data.cfg || {}, k = 'cfg_' + data.app, prev = JSON.parse(PROPS().getProperty(k) || '{"at":-1}');
  if (Number(c.at || 0) >= Number(prev.at || 0)) {
    const days = (c.days || []).map(Number).filter(d => d >= 0 && d <= 6);
    PROPS().setProperty(k, JSON.stringify({ on: !!c.on, days, hour: Math.min(23, Math.max(0, Number(c.hour) || 21)), empty: !!c.empty, freq: String(c.freq || ''), at: Number(c.at || 0), appName: String(data.appName || data.app).slice(0, 40) }));
  }
}
function guardarDatos_(data) {
  if (!data.alumnos || !data.alumnos.length) return;
  const ss = libro_(), sd = ss.getSheetByName('Dias'), se = ss.getSheetByName('Estado'), now = new Date();
  const dev = data.device || {}, devId = String(dev.id || 'x').slice(0, 20), devName = String(dev.name || 'Equipo').slice(0, 30);
  const idx = sh => { const v = sh.getDataRange().getValues(), m = {}; for (let i = 1; i < v.length; i++) m[v[i][0]] = i + 1; return m; };
  const mD = idx(sd), mE = idx(se);
  data.alumnos.slice(0, 10).forEach(a => {
    const alumno = String(a.alumno || 'Alumno').slice(0, 40);
    const est = { goal: a.goal, streak: a.streak, totalQ: a.totalQ, grades: a.grades || [], prio: a.prio || [], version: data.version };
    const ke = [data.app, devId, alumno].join('|'), rowE = [ke, data.app, data.appName, devId, devName, alumno, JSON.stringify(est), now];
    if (mE[ke]) se.getRange(mE[ke], 1, 1, rowE.length).setValues([rowE]); else { se.appendRow(rowE); mE[ke] = se.getLastRow(); }
    Object.keys(a.days || {}).slice(0, 10).forEach(f => {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(f)) return;
      const d = a.days[f], kd = [f, data.app, devId, alumno].join('|');
      const row = [kd, "'" + f, data.app, devName, alumno, Number(d.q) || 0, Number(d.ok) || 0, Number(d.min) || 0, JSON.stringify(d).slice(0, 45000), now];
      if (mD[kd]) { const old = sd.getRange(mD[kd], 6, 1, 3).getValues()[0]; if (old[0] === row[5] && old[2] === row[7] && old[1] === row[6]) return; sd.getRange(mD[kd], 1, 1, row.length).setValues([row]); }
      else { sd.appendRow(row); mD[kd] = sd.getLastRow(); }
    });
  });
}

/* ---------- Envío programado (disparador cada hora) ---------- */
function revisarEnvios() {
  const t = ahora_();
  PROPS().setProperty('ultima_revision', t.txt);
  ['sec2', 'sec3'].forEach(app => {
    try {
      const c = JSON.parse(PROPS().getProperty('cfg_' + app) || 'null');
      if (!c || !c.on || !c.days.map(Number).includes(t.dow) || t.hour < Number(c.hour)) return;
      if (PROPS().getProperty('sent_' + app) === t.fecha) return;     // ya se envió hoy
      if (t.hour > Number(c.hour) + 3) return;                        // no enviar muy tarde si el activador se atrasó demasiado
      enviarReporte_(app, false);
      PROPS().setProperty('sent_' + app, t.fecha);
      PROPS().deleteProperty('err_envio');
    } catch (e) {
      PROPS().setProperty('err_envio', t.txt + ' · ' + String(e).slice(0, 180));
      try { libro_().getSheetByName('Envios').appendRow([new Date(), app, 'ERROR', '', 0, String(e).slice(0, 200)]); } catch (e2) { }
    }
  });
}
/* Periodo que cubre el correo: desde el día siguiente al envío programado anterior hasta hoy (máximo 7 días). */
function periodo_(c, t) {
  const fechas = [t.fecha];
  for (let i = 1; i < 7; i++) { const d = new Date(t.date.getTime() - i * 86400000); const dw = Number(Utilities.formatDate(d, CONFIG.ZONA, 'u')) % 7; if (c.days.includes(dw)) break; fechas.unshift(Utilities.formatDate(d, CONFIG.ZONA, 'yyyy-MM-dd')); }
  return fechas;
}
function enviarReporte_(app, prueba) {
  const c = JSON.parse(PROPS().getProperty('cfg_' + app) || 'null') || { days: [0, 1, 2, 3, 4, 5, 6], hour: 21, empty: true, appName: app };
  const t = ahora_(), fechas = periodo_(c, t), ss = libro_();
  const dias = ss.getSheetByName('Dias').getDataRange().getValues().slice(1).filter(r => r[2] === app);
  const estados = ss.getSheetByName('Estado').getDataRange().getValues().slice(1).filter(r => r[1] === app);
  const ultimo = PROPS().getProperty('last_' + app); const ultimoT = ultimo ? new Date(ultimo) : null;
  const alumnos = [...new Set(estados.map(r => r[5]))];
  if (!alumnos.length) { if (prueba) enviar_(app, c, `Prueba · ${c.appName}`, html_(c, t, fechas, [], prueba), 'prueba', 0, fechas); return; }
  const bloques = alumnos.map(al => {
    const filas = dias.filter(r => r[4] === al), est = estados.filter(r => r[5] === al).sort((a, b) => b[7] - a[7]);
    const fstr = v => v instanceof Date ? Utilities.formatDate(v, CONFIG.ZONA, 'yyyy-MM-dd') : String(v);
    const porFecha = f => combinar_(filas.filter(r => fstr(r[1]) === f));
    const periodo = fechas.map(f => [f, porFecha(f)]);
    const tardias = ultimoT ? [...new Set(filas.filter(r => !fechas.includes(fstr(r[1])) && r[9] > ultimoT).map(r => fstr(r[1])))].sort().map(f => [f, porFecha(f)]) : [];
    return { alumno: al, periodo, tardias, estado: JSON.parse(est[0][6]), equipos: est.map(r => [r[4], r[7]]) };
  });
  const practico = bloques.some(b => b.periodo.some(p => p[1].q > 0) || b.tardias.length);
  if (!prueba && !practico && !c.empty) return;
  const asunto = `${prueba ? 'Prueba · ' : ''}Avance ${c.appName} · ${bloques.map(b => b.alumno).join(', ')} · ${fechas.length > 1 ? fechaTxt_(fechas[0]) + ' al ' + fechaTxt_(fechas[fechas.length - 1]) : fechaTxt_(t.fecha)}`;
  enviar_(app, c, asunto, html_(c, t, fechas, bloques, prueba), prueba ? 'prueba' : 'programado', bloques.length, fechas);
  if (!prueba) PROPS().setProperty('last_' + app, new Date().toISOString());
}
function enviar_(app, c, asunto, html, tipo, n, fechas) {
  const yo = Session.getEffectiveUser().getEmail();
  const to = [yo].concat(String(CONFIG.DESTINATARIOS_EXTRA || '').split(',').map(s => s.trim()).filter(s => /@/.test(s))).join(',');
  MailApp.sendEmail({ to, subject: asunto, htmlBody: html, name: 'Avance Secundaria' });
  libro_().getSheetByName('Envios').appendRow([new Date(), app, tipo, to, n, fechas[0] + ' a ' + fechas[fechas.length - 1]]);
}
function combinar_(filas) {
  const r = { q: 0, ok: 0, min: 0, subj: {}, topics: {}, ev: [], sims: [], equipos: [] };
  filas.forEach(f => {
    const d = JSON.parse(f[8] || '{}'); r.q += d.q || 0; r.ok += d.ok || 0; r.min += d.min || 0; if (d.q || d.min) r.equipos.push(f[3]);
    Object.keys(d.subj || {}).forEach(s => { const x = r.subj[s] || (r.subj[s] = [0, 0]); x[0] += d.subj[s][0]; x[1] += d.subj[s][1]; });
    (d.topics || []).forEach(tp => { const x = r.topics[tp[0]] || (r.topics[tp[0]] = [tp[0], tp[1], 0, 0, tp[4]]); x[2] += tp[2]; x[3] += tp[3]; x[4] = Math.max(x[4], tp[4]); });
    r.ev = r.ev.concat(d.ev || []); r.sims = r.sims.concat(d.sims || []);
  });
  r.topics = Object.values(r.topics).sort((a, b) => b[2] - a[2]);
  return r;
}

/* ---------- Correo (HTML de tema claro, compatible con Gmail y Outlook) ---------- */
const COL = { ok: '#15803D', mid: '#B45309', bad: '#B91C1C', azul: '#1D4ED8', gris: '#475569', borde: '#E2E8F0', fondo: '#F8FAFC' };
const esc_ = s => String(s == null ? '' : s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
const pct_ = (a, b) => b ? Math.round(100 * a / b) + '%' : '—';
const colAcc_ = (a, b) => !b ? COL.gris : a / b >= 0.8 ? COL.ok : a / b >= 0.6 ? COL.mid : COL.bad;
const colG_ = g => g == null ? COL.gris : g >= 9 ? COL.ok : g >= 7 ? COL.mid : COL.bad;
function fechaTxt_(f) { const [y, m, d] = f.split('-').map(Number); const dt = new Date(y, m - 1, d, 12); return DOW[dt.getDay()] + ' ' + d + ' de ' + ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'][m - 1]; }
function kpi_(v, l, col) { return `<td style="padding:10px;border:1px solid ${COL.borde};background:#fff;text-align:center;width:25%"><div style="font-size:22px;font-weight:700;color:${col || '#0F172A'}">${v}</div><div style="font-size:12px;color:${COL.gris}">${l}</div></td>`; }
function html_(c, t, fechas, bloques, prueba) {
  const css = 'font-family:Arial,Helvetica,sans-serif;color:#0F172A;font-size:14px;line-height:1.45';
  let h = `<div style="${css};max-width:680px;margin:auto;background:${COL.fondo};padding:16px">`;
  h += `<h2 style="margin:0 0 4px">${esc_(c.appName)} · reporte de avance</h2><div style="color:${COL.gris};font-size:13px">${fechas.length > 1 ? 'Periodo: ' + fechaTxt_(fechas[0]) + ' al ' + fechaTxt_(fechas[fechas.length - 1]) : fechaTxt_(t.fecha)} · generado ${t.txt}</div>`;
  if (prueba) h += `<p style="background:#EFF6FF;border:1px solid #BFDBFE;padding:8px;border-radius:6px">Este es un <b>correo de prueba</b>. Los correos programados llegarán ${esc_(frecTxt_(c))}, a las ${c.hour}:00 h (o unos minutos después).</p>`;
  if (!bloques.length) h += `<p>Todavía no hay datos de ningún alumno. Activa el reporte en la sección de padres de la app en cada equipo donde estudie.</p>`;
  bloques.forEach(b => {
    const tot = b.periodo.reduce((a, [, d]) => ({ q: a.q + d.q, ok: a.ok + d.ok, min: a.min + d.min }), { q: 0, ok: 0, min: 0 });
    const diasCon = b.periodo.filter(([, d]) => d.q > 0).length, e = b.estado, meta = e.goal || 25;
    h += `<div style="background:#fff;border:1px solid ${COL.borde};border-radius:8px;padding:14px;margin-top:14px"><h3 style="margin:0 0 8px">👤 ${esc_(b.alumno)}</h3>`;
    if (!tot.q && !tot.min) h += `<p style="background:#FEF2F2;border:1px solid #FECACA;padding:8px;border-radius:6px;color:${COL.bad}"><b>${fechas.length > 1 ? 'No practicó en el periodo.' : 'Hoy no practicó.'}</b> Si sí estudió, es posible que el equipo no se haya conectado a internet; los datos llegarán en cuanto lo haga.</p>`;
    h += `<table style="border-collapse:collapse;width:100%;margin:6px 0"><tr>${kpi_(tot.min + ' min', 'tiempo de estudio')}${kpi_(tot.q, fechas.length > 1 ? 'preguntas' : `preguntas (meta ${meta})`, fechas.length === 1 ? (tot.q >= meta ? COL.ok : tot.q ? COL.mid : COL.bad) : null)}${kpi_(pct_(tot.ok, tot.q), 'aciertos', colAcc_(tot.ok, tot.q))}${kpi_(fechas.length > 1 ? diasCon + '/' + fechas.length : (e.streak || 0), fechas.length > 1 ? 'días con estudio' : 'días seguidos')}</tr></table>`;
    if (fechas.length > 1) h += `<table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:6px"><tr style="background:${COL.fondo}"><th align="left" style="padding:5px;border:1px solid ${COL.borde}">Día</th><th style="padding:5px;border:1px solid ${COL.borde}">Min</th><th style="padding:5px;border:1px solid ${COL.borde}">Preguntas</th><th style="padding:5px;border:1px solid ${COL.borde}">Aciertos</th></tr>${b.periodo.map(([f, d]) => `<tr><td style="padding:5px;border:1px solid ${COL.borde}">${fechaTxt_(f)}</td><td align="center" style="padding:5px;border:1px solid ${COL.borde}">${d.min}</td><td align="center" style="padding:5px;border:1px solid ${COL.borde}">${d.q}</td><td align="center" style="padding:5px;border:1px solid ${COL.borde};color:${colAcc_(d.ok, d.q)}">${pct_(d.ok, d.q)}</td></tr>`).join('')}</table>`;
    const tot2 = combinarPeriodo_(b.periodo);
    const subj = Object.keys(tot2.subj);
    if (subj.length) h += `<h4 style="margin:12px 0 4px">Materias practicadas</h4><table style="border-collapse:collapse;width:100%;font-size:13px">${subj.sort((x, y) => tot2.subj[y][0] - tot2.subj[x][0]).map(s => `<tr><td style="padding:5px;border:1px solid ${COL.borde}">${esc_(s)}</td><td align="center" style="padding:5px;border:1px solid ${COL.borde}">${tot2.subj[s][0]} preguntas</td><td align="center" style="padding:5px;border:1px solid ${COL.borde};color:${colAcc_(tot2.subj[s][1], tot2.subj[s][0])};font-weight:700">${pct_(tot2.subj[s][1], tot2.subj[s][0])}</td></tr>`).join('')}</table>`;
    if (tot2.topics.length) h += `<h4 style="margin:12px 0 4px">Temas trabajados</h4><ul style="margin:0;padding-left:18px">${tot2.topics.slice(0, 10).map(x => `<li>${esc_(x[0])} <span style="color:${COL.gris}">(${esc_(x[1])} · ${x[2]} preguntas · <b style="color:${colAcc_(x[3], x[2])}">${pct_(x[3], x[2])}</b> · nivel ${x[4]} de 5)</span></li>`).join('')}</ul>`;
    const vistos = {}, logros = tot2.ev.filter(x => (x[1] === 'up' || x[1] === 'mastered') && !vistos[x.join('|')] && (vistos[x.join('|')] = 1));
    if (logros.length) h += `<h4 style="margin:12px 0 4px">🏅 Logros</h4><ul style="margin:0;padding-left:18px">${logros.slice(0, 10).map(x => `<li>${x[1] === 'mastered' ? '<b>Dominó</b>' : 'Subió a nivel ' + x[2] + ' en'} ${esc_(x[0])}</li>`).join('')}</ul>`;
    if (tot2.sims.length) h += `<h4 style="margin:12px 0 4px">📝 Simulacros de examen</h4><ul style="margin:0;padding-left:18px">${tot2.sims.map(x => `<li>${esc_(x[0])}: <b style="color:${colG_(x[1])}">${Number(x[1]).toFixed(1)}</b> (${x[2]} de ${x[3]})</li>`).join('')}</ul>`;
    if (b.tardias.length) h += `<p style="font-size:13px;color:${COL.gris}">Actividad registrada después del correo anterior: ${b.tardias.map(([f, d]) => `${fechaTxt_(f)}: ${d.q} preguntas, ${d.min} min, ${pct_(d.ok, d.q)} de aciertos`).join(' · ')}.</p>`;
    const gr = (e.grades || []).filter(g => g[1] != null);
    if (gr.length) h += `<h4 style="margin:12px 0 4px">Calificación estimada por materia (meta: 9)</h4><table style="border-collapse:collapse;width:100%;font-size:13px"><tr>${gr.map(g => `<td style="padding:6px;border:1px solid ${COL.borde};text-align:center"><div style="font-size:12px;color:${COL.gris}">${esc_(g[0])}</div><div style="font-size:18px;font-weight:700;color:${colG_(g[1])}">${Number(g[1]).toFixed(1)}</div>${g[2] < 100 ? `<div style="font-size:11px;color:${COL.gris}">${g[2]}% evaluado</div>` : ''}</td>`).join('')}</tr></table>`;
    if ((e.prio || []).length) h += `<h4 style="margin:12px 0 4px">🎯 Temas para reforzar</h4><ul style="margin:0;padding-left:18px">${e.prio.slice(0, 6).map(x => `<li>${esc_(x)}</li>`).join('')}</ul>`;
    h += `<p style="font-size:12px;color:${COL.gris};margin-top:10px">Equipos: ${b.equipos.map(([n, d]) => `${esc_(n)} (última conexión ${Utilities.formatDate(new Date(d), CONFIG.ZONA, 'dd/MM HH:mm')})`).join(' · ')}</p></div>`;
  });
  h += `<p style="font-size:12px;color:${COL.gris};margin-top:14px">La calificación estimada combina el nivel alcanzado y los aciertos recientes de los temas vistos en clase, más el último simulacro. Para cambiar la frecuencia, la hora o desactivar este correo: app &gt; Padres &gt; Configuración.</p></div>`;
  return h;
}
function combinarPeriodo_(periodo) {
  const r = { subj: {}, topics: {}, ev: [], sims: [] };
  periodo.forEach(([, d]) => { Object.keys(d.subj).forEach(s => { const x = r.subj[s] || (r.subj[s] = [0, 0]); x[0] += d.subj[s][0]; x[1] += d.subj[s][1]; });
    d.topics.forEach(tp => { const x = r.topics[tp[0]] || (r.topics[tp[0]] = [tp[0], tp[1], 0, 0, tp[4]]); x[2] += tp[2]; x[3] += tp[3]; x[4] = Math.max(x[4], tp[4]); });
    r.ev = r.ev.concat(d.ev); r.sims = r.sims.concat(d.sims); });
  r.topics = Object.values(r.topics).sort((a, b) => b[2] - a[2]); return r;
}
function frecTxt_(c) { const d = c.days || []; if (d.length === 7) return 'todos los días'; if (d.length === 5 && [1, 2, 3, 4, 5].every(x => d.includes(x))) return 'de lunes a viernes'; if (d.length === 1 && c.freq === 'semanal') return 'cada ' + DOW[d[0]]; return 'los ' + d.map(x => DOW[x]).join(', '); }

/* ---------- Utilidades ---------- */
function ahora_() { const d = new Date(); return { date: d, fecha: Utilities.formatDate(d, CONFIG.ZONA, 'yyyy-MM-dd'), hour: Number(Utilities.formatDate(d, CONFIG.ZONA, 'H')), dow: Number(Utilities.formatDate(d, CONFIG.ZONA, 'u')) % 7, txt: Utilities.formatDate(d, CONFIG.ZONA, 'dd/MM/yyyy HH:mm') }; }
const hoy_ = () => ahora_().fecha;
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
/* Para probar desde el editor: envía ahora el reporte de la app de 3.° (o cambia a 'sec2'). */
function probarCorreo() { enviarReporte_('sec3', true); }
