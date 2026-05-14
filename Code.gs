/**
 * CONVASES Task Manager — Apps Script backend
 *
 * SETUP (ver README.md):
 *   1. Script Properties (Configuración del proyecto > Script Properties):
 *        - API_TOKEN      → token secreto compartido con el frontend
 *        - GEMINI_API_KEY → key de Google AI Studio (https://aistudio.google.com/apikey)
 *        - SHEET_ID       → 1LzjQQLFh5719qyEzpBYSYclcbuIiF3IE_MSR3XZpXxA
 *        - EMAIL_TO       → moises.vasquez999@gmail.com
 *   2. Correr una vez manualmente: setupAll() — crea hojas e instala triggers
 *   3. Deploy > Nueva implementación > Web app
 *        - Execute as: Me (moises.vasquez999@gmail.com)
 *        - Who has access: Anyone
 *      Copia la URL y pégala en index.html (constante SCRIPT_URL)
 */

// =====================================================
// CONFIG
// =====================================================
const SH = {
  TAREAS:       'Tareas',
  BITACORA:     'Bitacora',
  PROYECTOS:    'Config_Proyectos',
  HERRAMIENTAS: 'Config_Herramientas',
  BLOQUEADORES: 'Config_Bloqueadores'
};

const HEADERS = {
  TAREAS: ['id','createdAt','updatedAt','title','desc','project','priority','tool','status','blocker','startDate','dueDate','evidenceLink','evidencePending','completedAt','deleted','source'],
  BITACORA: ['id','taskId','taskTitle','project','date','startTime','endTime','minutes'],
  CONFIG: ['id','nombre','activo']
};

const DEFAULTS = {
  PROYECTOS:    ['Chalatenango','UC3','Las Vegas','Cochera','Granja / PEMA','Licitaciones','Interno CONVASES'],
  HERRAMIENTAS: ['AutoCAD / Civil 3D','Revit','Excel','Campo presencial','Gestión / Trámite','WhatsApp / Correo','IA','Múltiple'],
  BLOQUEADORES: ['Marcos (MOP)','Luis (Campo)','Francisco (Campo)','Alicia (Finanzas)','Herbert (Oficina técnica)','Pa','Subcontratista','Proveedor','Alcaldía / CNR']
};

const STATUSES = ['Backlog','Esperando','Gabinete','Campo','Completado'];
const PRIORITIES = ['Alta','Media','Baja'];
const TZ = 'America/El_Salvador';

// =====================================================
// ENTRY POINTS
// =====================================================
function doGet(e) {
  return json({ ok:true, ping:'CONVASES Task Manager API', time:new Date().toISOString() });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.token !== prop('API_TOKEN')) return json({ ok:false, error:'unauthorized' });
    const action = body.action;
    switch (action) {
      case 'init':         return json({ ok:true, data: actionInit() });
      case 'saveTask':     return json({ ok:true, data: actionSaveTask(body) });
      case 'deleteTask':   return json({ ok:true, data: actionDeleteTask(body) });
      case 'startWork':    return json({ ok:true, data: actionStartWork(body) });
      case 'stopWork':     return json({ ok:true, data: actionStopWork(body) });
      case 'listBitacora': return json({ ok:true, data: actionListBitacora(body) });
      case 'parseAI':      return json({ ok:true, data: actionParseAI(body) });
      case 'saveConfig':   return json({ ok:true, data: actionSaveConfig(body) });
      case 'deleteConfig': return json({ ok:true, data: actionDeleteConfig(body) });
      default: return json({ ok:false, error:'unknown_action:' + action });
    }
  } catch (err) {
    return json({ ok:false, error: String(err && err.message || err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// ACTIONS — TASKS
// =====================================================
function actionInit() {
  ensureSheets();
  return {
    tasks:    listTasks(),
    bitacora: listTodayBitacora(),
    active:   findActiveBitacora(),
    config: {
      proyectos:    listConfig(SH.PROYECTOS),
      herramientas: listConfig(SH.HERRAMIENTAS),
      bloqueadores: listConfig(SH.BLOQUEADORES)
    },
    statuses:   STATUSES,
    priorities: PRIORITIES
  };
}

function listTasks() {
  const sh = getSheet(SH.TAREAS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .map(r => rowToObject(r, headers))
    .filter(t => !asBool(t.deleted));
}

function actionSaveTask(body) {
  const t = body.task || {};
  const sh = getSheet(SH.TAREAS);
  const now = new Date().toISOString();
  const headers = HEADERS.TAREAS;
  let row = findRowById(sh, t.id);

  if (!row) {
    // CREATE
    const obj = {
      id:               t.id || uid(),
      createdAt:        now,
      updatedAt:        now,
      title:            t.title || '',
      desc:             t.desc || '',
      project:          t.project || '',
      priority:         PRIORITIES.indexOf(t.priority) >= 0 ? t.priority : 'Media',
      tool:             t.tool || '',
      status:           STATUSES.indexOf(t.status) >= 0 ? t.status : 'Backlog',
      blocker:          t.blocker || '',
      startDate:        t.startDate || '',
      dueDate:          t.dueDate || '',
      evidenceLink:     t.evidenceLink || '',
      evidencePending:  !!t.evidencePending,
      completedAt:      t.status === 'Completado' ? now : '',
      deleted:          false,
      source:           t.source || 'manual'
    };
    sh.appendRow(objectToRow(obj, headers));
    return obj;
  } else {
    // UPDATE
    const existing = rowToObject(sh.getRange(row.rowIndex, 1, 1, headers.length).getValues()[0], headers);
    const wasCompleted = existing.status === 'Completado';
    const nowCompleted = t.status === 'Completado';
    const merged = Object.assign({}, existing, {
      title:            t.title ?? existing.title,
      desc:             t.desc ?? existing.desc,
      project:          t.project ?? existing.project,
      priority:         t.priority ?? existing.priority,
      tool:             t.tool ?? existing.tool,
      status:           t.status ?? existing.status,
      blocker:          t.blocker ?? existing.blocker,
      startDate:        t.startDate ?? existing.startDate,
      dueDate:          t.dueDate ?? existing.dueDate,
      evidenceLink:     t.evidenceLink ?? existing.evidenceLink,
      evidencePending:  typeof t.evidencePending === 'boolean' ? t.evidencePending : asBool(existing.evidencePending),
      updatedAt:        now,
      completedAt:      (!wasCompleted && nowCompleted) ? now : (nowCompleted ? existing.completedAt : '')
    });
    sh.getRange(row.rowIndex, 1, 1, headers.length).setValues([objectToRow(merged, headers)]);
    return merged;
  }
}

function actionDeleteTask(body) {
  const sh = getSheet(SH.TAREAS);
  const row = findRowById(sh, body.id);
  if (!row) return { id: body.id, deleted: false };
  const headers = HEADERS.TAREAS;
  const obj = rowToObject(sh.getRange(row.rowIndex, 1, 1, headers.length).getValues()[0], headers);
  obj.deleted = true;
  obj.updatedAt = new Date().toISOString();
  sh.getRange(row.rowIndex, 1, 1, headers.length).setValues([objectToRow(obj, headers)]);
  return { id: body.id, deleted: true };
}

// =====================================================
// ACTIONS — BITACORA
// =====================================================
function actionStartWork(body) {
  const sh = getSheet(SH.BITACORA);
  const tsh = getSheet(SH.TAREAS);
  const trow = findRowById(tsh, body.taskId);
  if (!trow) throw new Error('task_not_found');
  const t = rowToObject(tsh.getRange(trow.rowIndex, 1, 1, HEADERS.TAREAS.length).getValues()[0], HEADERS.TAREAS);

  // Cerrar cualquier activo antes
  const active = findActiveBitacora();
  if (active) stopBitacora(active.id);

  const now = new Date();
  const entry = {
    id:        uid(),
    taskId:    body.taskId,
    taskTitle: t.title,
    project:   t.project,
    date:      Utilities.formatDate(now, TZ, 'yyyy-MM-dd'),
    startTime: now.toISOString(),
    endTime:   '',
    minutes:   ''
  };
  sh.appendRow(objectToRow(entry, HEADERS.BITACORA));
  return entry;
}

function actionStopWork(body) {
  return stopBitacora(body.bitacoraId);
}

function stopBitacora(bitacoraId) {
  const sh = getSheet(SH.BITACORA);
  const row = findRowById(sh, bitacoraId);
  if (!row) throw new Error('bitacora_not_found');
  const entry = rowToObject(sh.getRange(row.rowIndex, 1, 1, HEADERS.BITACORA.length).getValues()[0], HEADERS.BITACORA);
  if (entry.endTime) return entry;
  const now = new Date();
  const start = new Date(entry.startTime);
  entry.endTime = now.toISOString();
  entry.minutes = Math.max(1, Math.round((now - start) / 60000));
  sh.getRange(row.rowIndex, 1, 1, HEADERS.BITACORA.length).setValues([objectToRow(entry, HEADERS.BITACORA)]);
  return entry;
}

function actionListBitacora(body) {
  const date = body.date || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  return listBitacoraByDate(date);
}

function listBitacoraByDate(date) {
  const sh = getSheet(SH.BITACORA);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .map(r => rowToObject(r, headers))
    .filter(b => b.date === date);
}

function listTodayBitacora() {
  return listBitacoraByDate(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
}

function findActiveBitacora() {
  const sh = getSheet(SH.BITACORA);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  for (let i = data.length - 1; i >= 1; i--) {
    const b = rowToObject(data[i], headers);
    if (b.startTime && !b.endTime) return b;
  }
  return null;
}

// =====================================================
// ACTIONS — CONFIG
// =====================================================
function listConfig(sheetName) {
  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .map(r => rowToObject(r, headers))
    .filter(c => asBool(c.activo) !== false);
}

function actionSaveConfig(body) {
  const sheetName = configSheetName(body.type);
  const sh = getSheet(sheetName);
  const item = body.item || {};
  const headers = HEADERS.CONFIG;
  let row = item.id ? findRowById(sh, item.id) : null;
  if (!row) {
    const obj = { id: item.id || uid(), nombre: item.nombre || '', activo: item.activo !== false };
    sh.appendRow(objectToRow(obj, headers));
    return obj;
  } else {
    const existing = rowToObject(sh.getRange(row.rowIndex, 1, 1, headers.length).getValues()[0], headers);
    const merged = Object.assign({}, existing, {
      nombre: item.nombre ?? existing.nombre,
      activo: typeof item.activo === 'boolean' ? item.activo : asBool(existing.activo)
    });
    sh.getRange(row.rowIndex, 1, 1, headers.length).setValues([objectToRow(merged, headers)]);
    return merged;
  }
}

function actionDeleteConfig(body) {
  // Soft-delete: marca activo=false (nunca borra)
  const sheetName = configSheetName(body.type);
  const sh = getSheet(sheetName);
  const row = findRowById(sh, body.id);
  if (!row) return { id: body.id, deleted: false };
  const headers = HEADERS.CONFIG;
  const obj = rowToObject(sh.getRange(row.rowIndex, 1, 1, headers.length).getValues()[0], headers);
  obj.activo = false;
  sh.getRange(row.rowIndex, 1, 1, headers.length).setValues([objectToRow(obj, headers)]);
  return { id: body.id, deleted: true };
}

function configSheetName(type) {
  if (type === 'proyecto')    return SH.PROYECTOS;
  if (type === 'herramienta') return SH.HERRAMIENTAS;
  if (type === 'bloqueador')  return SH.BLOQUEADORES;
  throw new Error('config_type_inválido: ' + type);
}

// =====================================================
// AI — GEMINI
// =====================================================
function actionParseAI(body) {
  const text = (body.text || '').trim();
  if (!text) return { tasks: [] };
  const parsed = callGemini(text);
  // Guardar todas en Backlog directamente (Moi las edita después)
  const saved = parsed.map(p => actionSaveTask({
    task: {
      title:    p.title || 'Tarea sin título',
      desc:     p.desc || '',
      project:  p.project || '',
      priority: p.priority || 'Media',
      tool:     p.tool || '',
      status:   p.status || 'Backlog',
      blocker:  p.blocker || '',
      dueDate:  p.dueDate || '',
      source:   'ai'
    }
  }));
  return { tasks: saved };
}

function callGemini(text) {
  const apiKey = prop('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY no configurada en Script Properties');

  const proyectos    = listConfig(SH.PROYECTOS).map(p => p.nombre);
  const herramientas = listConfig(SH.HERRAMIENTAS).map(h => h.nombre);
  const bloqueadores = listConfig(SH.BLOQUEADORES).map(b => b.nombre);

  const systemPrompt = [
    'Sos asistente de captura para Moi, ingeniero civil de CONVASES S.A. de C.V., El Salvador.',
    'Maneja múltiples proyectos de construcción y obra civil simultáneamente.',
    'Tu trabajo: leer texto libre (puede ser una idea muy rápida o incompleta) y devolver un array JSON de tareas.',
    '',
    'PROYECTOS válidos: ' + proyectos.join(', '),
    'HERRAMIENTAS válidas: ' + herramientas.join(', '),
    'ESTADOS válidos: Backlog, Esperando, Gabinete, Campo',
    'PRIORIDADES: Alta, Media, Baja',
    'BLOQUEADORES posibles: ' + bloqueadores.join(', '),
    '',
    'REGLAS:',
    '- Si menciona "Pa" o "urgente" → priority Alta.',
    '- Si menciona esperar a alguien o "bloqueado por" → status Esperando + blocker apropiado.',
    '- Si menciona CAD/Revit/dibujo/Excel/estimación → status Gabinete + tool apropiado.',
    '- Si menciona ir/visitar/inspeccionar/MOP/alcaldía → status Campo.',
    '- Si no hay info clara → status Backlog (que él lo edite luego).',
    '- Si menciona ChatGPT/IA/Claude/análisis con IA → tool "IA".',
    '- desc: detalles adicionales útiles, NO repitas el title.',
    '- due: solo si menciona fecha explícita (formato YYYY-MM-DD).',
    '- Una sola tarea por idea separada. Si el texto tiene varias tareas separadas por coma/punto/"y", devolvé varias.',
    '- Title corto y accionable, en imperativo cuando posible.',
    '',
    'Devolvé SOLO JSON, sin texto adicional.'
  ].join('\n');

  const schema = {
    type: 'ARRAY',
    items: {
      type: 'OBJECT',
      properties: {
        title:    { type: 'STRING' },
        desc:     { type: 'STRING' },
        project:  { type: 'STRING' },
        priority: { type: 'STRING', enum: PRIORITIES },
        tool:     { type: 'STRING' },
        status:   { type: 'STRING', enum: ['Backlog','Esperando','Gabinete','Campo'] },
        blocker:  { type: 'STRING' },
        dueDate:  { type: 'STRING' }
      },
      required: ['title']
    }
  };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=' + apiKey;
  const payload = {
    contents: [{ role:'user', parts:[{ text: text }] }],
    systemInstruction: { parts: [{ text: systemPrompt }] },
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: schema,
      temperature: 0.2
    }
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const txt = res.getContentText();
  if (code !== 200) throw new Error('Gemini ' + code + ': ' + txt.slice(0, 300));
  const data = JSON.parse(txt);
  const raw = data.candidates?.[0]?.content?.parts?.[0]?.text || '[]';
  let arr;
  try { arr = JSON.parse(raw); } catch (e) { throw new Error('Gemini devolvió JSON inválido: ' + raw.slice(0,200)); }
  if (!Array.isArray(arr)) arr = [arr];
  return arr;
}

// =====================================================
// EMAILS — TRIGGERS
// =====================================================
function dailyMorningEmail() {
  const tasks = listTasks();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const inSevenDays = Utilities.formatDate(new Date(Date.now() + 7*86400000), TZ, 'yyyy-MM-dd');

  const overdue   = tasks.filter(t => t.status !== 'Completado' && t.dueDate && fmtDate(t.dueDate) < today);
  const dueToday  = tasks.filter(t => t.status !== 'Completado' && t.dueDate && fmtDate(t.dueDate) === today);
  const dueWeek   = tasks.filter(t => t.status !== 'Completado' && t.dueDate && fmtDate(t.dueDate) > today && fmtDate(t.dueDate) <= inSevenDays);
  const blocked   = tasks.filter(t => t.status === 'Esperando');
  const urgentND  = tasks.filter(t => t.status !== 'Completado' && t.priority === 'Alta' && !t.dueDate);

  const sections = [
    { title: '🔴 VENCIDAS',                items: overdue },
    { title: '🟡 PARA HOY',                items: dueToday },
    { title: '📅 ESTA SEMANA',             items: dueWeek },
    { title: '⏳ BLOQUEADAS',              items: blocked },
    { title: '⚡ URGENTES SIN FECHA',      items: urgentND }
  ];

  const totalCount = sections.reduce((n, s) => n + s.items.length, 0);
  if (totalCount === 0) {
    sendEmail('CONVASES · ' + today + ' · sin pendientes', '<p>No hay tareas que reportar esta mañana. Buen día.</p>');
    return;
  }

  const body = '<div style="font-family:Arial,sans-serif;max-width:640px">' +
    '<h2 style="color:#e8a020;margin:0 0 4px">CONVASES · resumen matutino</h2>' +
    '<p style="color:#666;margin:0 0 16px;font-size:12px">' + today + ' — ' + totalCount + ' tarea(s) requieren atención</p>' +
    sections.filter(s => s.items.length).map(s => sectionHTML(s.title, s.items)).join('') +
    '<p style="color:#999;font-size:11px;margin-top:24px">App: https://moisesvasqueza00834613.github.io/convases-tareas-/</p>' +
    '</div>';
  sendEmail('CONVASES · ' + today + ' · ' + totalCount + ' pendientes', body);
}

function dailyEveningEmail() {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const entries = listBitacoraByDate(today).filter(b => b.endTime); // solo cerrados

  // Agrupar por proyecto
  const byProject = {};
  let totalMin = 0;
  entries.forEach(e => {
    const p = e.project || '(sin proyecto)';
    if (!byProject[p]) byProject[p] = { entries: [], total: 0 };
    byProject[p].entries.push(e);
    byProject[p].total += Number(e.minutes) || 0;
    totalMin += Number(e.minutes) || 0;
  });

  const tasks = listTasks();
  const evidencePending = tasks.filter(t => t.status === 'Completado' && asBool(t.evidencePending));

  if (entries.length === 0 && evidencePending.length === 0) {
    sendEmail('CONVASES · ' + today + ' · sin bitácora', '<p>No registraste tiempo hoy. Hasta mañana.</p>');
    return;
  }

  let html = '<div style="font-family:Arial,sans-serif;max-width:640px">' +
    '<h2 style="color:#e8a020;margin:0 0 4px">CONVASES · cierre del día</h2>' +
    '<p style="color:#666;margin:0 0 16px;font-size:12px">' + today + ' — ' + fmtMinutes(totalMin) + ' trabajados</p>';

  if (entries.length) {
    html += '<h3 style="margin:16px 0 8px;font-size:14px">⏱ BITÁCORA POR PROYECTO</h3>';
    Object.keys(byProject).sort((a,b) => byProject[b].total - byProject[a].total).forEach(p => {
      const d = byProject[p];
      html += '<div style="margin:8px 0 12px;padding:10px 12px;background:#f7f7f5;border-left:3px solid #e8a020;border-radius:4px">' +
        '<div style="font-weight:700">' + escapeHtml(p) + ' <span style="color:#888;font-weight:400;font-size:12px">— ' + fmtMinutes(d.total) + '</span></div>' +
        '<ul style="margin:6px 0 0;padding-left:20px;color:#444;font-size:13px">' +
          d.entries.map(e => '<li>' + escapeHtml(e.taskTitle) + ' <span style="color:#888">(' + fmtMinutes(e.minutes) + ')</span></li>').join('') +
        '</ul></div>';
    });
  }

  if (evidencePending.length) {
    html += '<h3 style="margin:20px 0 8px;font-size:14px">📎 EVIDENCIA PENDIENTE DE SUBIR</h3>' +
      '<ul style="margin:0;padding-left:20px;color:#444;font-size:13px">' +
        evidencePending.map(t => '<li>' + escapeHtml(t.title) + (t.project ? ' <span style="color:#888">— ' + escapeHtml(t.project) + '</span>' : '') + '</li>').join('') +
      '</ul>';
  }

  html += '<p style="color:#999;font-size:11px;margin-top:24px">App: https://moisesvasqueza00834613.github.io/convases-tareas-/</p></div>';
  sendEmail('CONVASES · ' + today + ' · ' + fmtMinutes(totalMin) + ' trabajados', html);
}

function sectionHTML(title, items) {
  return '<h3 style="margin:16px 0 6px;font-size:14px">' + title + ' <span style="color:#888;font-weight:400">(' + items.length + ')</span></h3>' +
    '<ul style="margin:0;padding-left:20px;color:#444;font-size:13px">' +
      items.map(t => {
        const parts = [escapeHtml(t.title)];
        const meta = [];
        if (t.project) meta.push(escapeHtml(t.project));
        if (t.priority === 'Alta') meta.push('Alta');
        if (t.dueDate) meta.push(fmtDate(t.dueDate));
        if (t.blocker) meta.push('⏳ ' + escapeHtml(t.blocker));
        if (meta.length) parts.push('<span style="color:#888"> — ' + meta.join(' · ') + '</span>');
        return '<li>' + parts.join('') + '</li>';
      }).join('') +
    '</ul>';
}

function sendEmail(subject, htmlBody) {
  const to = prop('EMAIL_TO');
  if (!to) throw new Error('EMAIL_TO no configurada');
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: htmlBody });
}

// =====================================================
// SETUP — corre estos manualmente UNA SOLA VEZ
// =====================================================
function setupAll() {
  ensureSheets();
  installTriggers();
  Logger.log('Setup completo. Verificá Script Properties: API_TOKEN, GEMINI_API_KEY, SHEET_ID, EMAIL_TO');
}

function ensureSheets() {
  const ss = SpreadsheetApp.openById(prop('SHEET_ID'));
  ensureSheet(ss, SH.TAREAS, HEADERS.TAREAS);
  ensureSheet(ss, SH.BITACORA, HEADERS.BITACORA);
  seedIfEmpty(ss, SH.PROYECTOS,    HEADERS.CONFIG, DEFAULTS.PROYECTOS);
  seedIfEmpty(ss, SH.HERRAMIENTAS, HEADERS.CONFIG, DEFAULTS.HERRAMIENTAS);
  seedIfEmpty(ss, SH.BLOQUEADORES, HEADERS.CONFIG, DEFAULTS.BLOQUEADORES);
}

function seedIfEmpty(ss, name, headers, defaults) {
  ensureSheet(ss, name, headers);
  const sh = ss.getSheetByName(name);
  if (sh.getLastRow() <= 1) seedConfig(sh, defaults);
}

function ensureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  const created = !sh;
  if (!sh) sh = ss.insertSheet(name);
  // headers si la hoja está vacía
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a1e2b').setFontColor('#e8a020');
  }
  return created;
}

function seedConfig(sh, names) {
  const headers = HEADERS.CONFIG;
  const rows = names.map(n => objectToRow({ id: uid(), nombre: n, activo: true }, headers));
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyMorningEmail' || t.getHandlerFunction() === 'dailyEveningEmail') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('dailyMorningEmail').timeBased().atHour(6).nearMinute(30).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('dailyEveningEmail').timeBased().atHour(17).nearMinute(15).everyDays(1).inTimezone(TZ).create();
  Logger.log('Triggers instalados: 6:30 AM y 5:15 PM (' + TZ + ')');
}

// =====================================================
// HELPERS
// =====================================================
function getSheet(name) {
  const ss = SpreadsheetApp.openById(prop('SHEET_ID'));
  let sh = ss.getSheetByName(name);
  if (!sh) { ensureSheets(); sh = ss.getSheetByName(name); }
  return sh;
}

function findRowById(sh, id) {
  if (!id) return null;
  const data = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return { rowIndex: i + 2 };
  }
  return null;
}

function rowToObject(row, headers) {
  const o = {};
  headers.forEach((h, i) => o[h] = row[i]);
  // normalizar fechas que vienen como Date
  ['createdAt','updatedAt','completedAt','startTime','endTime'].forEach(k => {
    if (o[k] instanceof Date) o[k] = o[k].toISOString();
  });
  ['startDate','dueDate','date'].forEach(k => {
    if (o[k] instanceof Date) o[k] = Utilities.formatDate(o[k], TZ, 'yyyy-MM-dd');
  });
  return o;
}

function objectToRow(obj, headers) {
  return headers.map(h => {
    const v = obj[h];
    if (v === undefined || v === null) return '';
    if (typeof v === 'boolean') return v;
    return v;
  });
}

function asBool(v) {
  if (typeof v === 'boolean') return v;
  if (v === 'TRUE' || v === 'true') return true;
  if (v === 'FALSE' || v === 'false') return false;
  return !!v;
}

function uid() {
  return Utilities.getUuid().slice(0, 8) + Date.now().toString(36);
}

function prop(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}

function fmtMinutes(m) {
  m = Number(m) || 0;
  if (m < 60) return m + ' min';
  const h = Math.floor(m / 60);
  const r = m % 60;
  return r ? h + 'h ' + r + 'min' : h + 'h';
}

function escapeHtml(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
