/**
 * CONVASES Task Manager — Apps Script backend (v3)
 *
 * SETUP:
 *   1. Script Properties (Configuración del proyecto > Script Properties):
 *        - API_TOKEN          → token secreto compartido con el frontend
 *        - ANTHROPIC_API_KEY  → key de Anthropic (https://console.anthropic.com/settings/keys)
 *        - SHEET_ID           → 1LzjQQLFh5719qyEzpBYSYclcbuIiF3IE_MSR3XZpXxA
 *        - EMAIL_TO           → moises.vasquez999@gmail.com
 *   2. Correr una vez manualmente: setupAll() — crea/migra hojas e instala triggers
 *   3. Deploy > Nueva implementación > Web app
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Copia la URL y pégala en index.html (constante SCRIPT_URL)
 *
 * MIGRACIÓN desde v2:
 *   - Hoja Tareas: se agregan columnas cancelledAt + justificante después de completedAt
 *   - Hoja Bitacora: se agregan anio + mes + semana después de project
 *   - Se crea hoja Historial nueva
 *   - Estados legacy se normalizan: Backlog→Por iniciar, Gabinete→En progreso,
 *     Esperando→En espera, Campo→Por iniciar
 */

// =====================================================
// CONFIG
// =====================================================
const SH = {
  TAREAS:       'Tareas',
  BITACORA:     'Bitacora',
  HISTORIAL:    'Historial',
  PROYECTOS:    'Config_Proyectos',
  HERRAMIENTAS: 'Config_Herramientas',
  BLOQUEADORES: 'Config_Bloqueadores'
};

const HEADERS = {
  TAREAS:    ['id','createdAt','updatedAt','title','desc','project','priority','tool','status','blocker','startDate','dueDate','evidenceLink','evidencePending','completedAt','cancelledAt','justificante','deleted','source'],
  BITACORA:  ['id','taskId','taskTitle','project','anio','mes','semana','date','startTime','endTime','minutes'],
  HISTORIAL: ['anio','mes','nombreMes','semana','diaSemana','fecha','taskId','titulo','proyecto','prioridad','herramienta','estado','justificante','tiempoTotalMinutos','evidenceLink','evidencePending'],
  CONFIG:    ['id','nombre','activo']
};

const DEFAULTS = {
  PROYECTOS:    ['Chalatenango','UC3','Las Vegas','Cochera','Granja / PEMA','Licitaciones','Interno CONVASES'],
  HERRAMIENTAS: ['AutoCAD / Civil 3D','Revit','Excel','Campo presencial','Gestión / Trámite','WhatsApp / Correo','IA','Múltiple'],
  BLOQUEADORES: ['Marcos (MOP)','Luis (Campo)','Francisco (Campo)','Alicia (Finanzas)','Herbert (Oficina técnica)','Pa','Subcontratista','Proveedor','Alcaldía / CNR']
};

const STATUSES   = ['Por iniciar','En progreso','En espera','Pausado','Cancelado','Completado'];
const PRIORITIES = ['Alta','Media','Baja'];
const TZ         = 'America/El_Salvador';

const LEGACY_STATUS_MAP = {
  'Backlog':    'Por iniciar',
  'Gabinete':   'En progreso',
  'Esperando':  'En espera',
  'Campo':      'Por iniciar'
};

const MESES_ES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const DIAS_ES  = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];

// =====================================================
// ENTRY POINTS
// =====================================================
function doGet(e) {
  return json({ ok:true, ping:'CONVASES Task Manager API v3', time:new Date().toISOString() });
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
    tasks:      listTasks(),
    bitacora:   listTodayBitacora(),
    active:     findActiveBitacora(),
    taskTotals: computeTaskTotals(),
    config: {
      proyectos:    listConfig(SH.PROYECTOS),
      herramientas: listConfig(SH.HERRAMIENTAS),
      bloqueadores: listConfig(SH.BLOQUEADORES)
    },
    statuses:   STATUSES,
    priorities: PRIORITIES,
    serverTime: new Date().toISOString()
  };
}

function listTasks() {
  const sh = getSheet(SH.TAREAS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .map(r => rowToObject(r, headers))
    .filter(t => !asBool(t.deleted))
    .map(t => { t.status = normalizeStatus(t.status); return t; });
}

function actionSaveTask(body) {
  const t = body.task || {};
  const sh = getSheet(SH.TAREAS);
  const now = new Date().toISOString();
  const headers = HEADERS.TAREAS;
  let row = findRowById(sh, t.id);

  const incomingStatus = normalizeStatus(t.status);
  // Cancelado requiere justificante
  if (incomingStatus === 'Cancelado' && !String(t.justificante || '').trim()) {
    throw new Error('Para cancelar una tarea se requiere justificante');
  }

  if (!row) {
    const obj = {
      id:              t.id || uid(),
      createdAt:       now,
      updatedAt:       now,
      title:           t.title || '',
      desc:            t.desc || '',
      project:         t.project || '',
      priority:        PRIORITIES.indexOf(t.priority) >= 0 ? t.priority : 'Media',
      tool:            t.tool || '',
      status:          incomingStatus || 'Por iniciar',
      blocker:         t.blocker || '',
      startDate:       t.startDate || '',
      dueDate:         t.dueDate || '',
      evidenceLink:    t.evidenceLink || '',
      evidencePending: !!t.evidencePending,
      completedAt:     incomingStatus === 'Completado' ? now : '',
      cancelledAt:     incomingStatus === 'Cancelado'  ? now : '',
      justificante:    t.justificante || '',
      deleted:         false,
      source:          t.source || 'manual'
    };
    sh.appendRow(objectToRow(obj, headers));
    maybeAppendHistorial(obj, null);
    return obj;
  }

  const existing = rowToObject(sh.getRange(row.rowIndex, 1, 1, headers.length).getValues()[0], headers);
  existing.status = normalizeStatus(existing.status);
  const wasTerminal = existing.status === 'Completado' || existing.status === 'Cancelado';
  const nowTerminal = incomingStatus === 'Completado' || incomingStatus === 'Cancelado';
  const transitionedToTerminal = !wasTerminal && nowTerminal;

  const merged = Object.assign({}, existing, {
    title:           t.title ?? existing.title,
    desc:            t.desc ?? existing.desc,
    project:         t.project ?? existing.project,
    priority:        t.priority ?? existing.priority,
    tool:            t.tool ?? existing.tool,
    status:          incomingStatus,
    blocker:         t.blocker ?? existing.blocker,
    startDate:       t.startDate ?? existing.startDate,
    dueDate:         t.dueDate ?? existing.dueDate,
    evidenceLink:    t.evidenceLink ?? existing.evidenceLink,
    evidencePending: typeof t.evidencePending === 'boolean' ? t.evidencePending : asBool(existing.evidencePending),
    justificante:    t.justificante ?? existing.justificante,
    updatedAt:       now,
    completedAt:     incomingStatus === 'Completado'
                       ? (existing.completedAt || now)
                       : (incomingStatus === 'Cancelado' ? existing.completedAt : ''),
    cancelledAt:     incomingStatus === 'Cancelado'
                       ? (existing.cancelledAt || now)
                       : ''
  });
  sh.getRange(row.rowIndex, 1, 1, headers.length).setValues([objectToRow(merged, headers)]);
  if (transitionedToTerminal) maybeAppendHistorial(merged, existing);
  return merged;
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
// HISTORIAL
// =====================================================
function maybeAppendHistorial(task, previous) {
  const status = normalizeStatus(task.status);
  if (status !== 'Completado' && status !== 'Cancelado') return;
  const sh = getSheet(SH.HISTORIAL);
  const eventDate = status === 'Completado'
    ? new Date(task.completedAt || new Date())
    : new Date(task.cancelledAt || new Date());
  const totalMin = computeTotalMinutesForTask(task.id);
  const row = {
    anio:                 eventDate.getFullYear(),
    mes:                  eventDate.getMonth() + 1,
    nombreMes:            MESES_ES[eventDate.getMonth()],
    semana:               getIsoWeek(eventDate),
    diaSemana:            DIAS_ES[eventDate.getDay()],
    fecha:                Utilities.formatDate(eventDate, TZ, 'yyyy-MM-dd'),
    taskId:               task.id,
    titulo:               task.title || '',
    proyecto:             task.project || '',
    prioridad:            task.priority || '',
    herramienta:          task.tool || '',
    estado:               status,
    justificante:         task.justificante || '',
    tiempoTotalMinutos:   totalMin,
    evidenceLink:         task.evidenceLink || '',
    evidencePending:      !!task.evidencePending
  };
  sh.appendRow(objectToRow(row, HEADERS.HISTORIAL));
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
    anio:      now.getFullYear(),
    mes:       now.getMonth() + 1,
    semana:    getIsoWeek(now),
    date:      Utilities.formatDate(now, TZ, 'yyyy-MM-dd'),
    startTime: now.toISOString(),
    endTime:   '',
    minutes:   ''
  };
  sh.appendRow(objectToRow(entry, HEADERS.BITACORA));

  // Si la tarea está en "Por iniciar" o "Pausado", pasarla a "En progreso"
  const currentStatus = normalizeStatus(t.status);
  if (currentStatus === 'Por iniciar' || currentStatus === 'Pausado') {
    actionSaveTask({ task: { id: body.taskId, status: 'En progreso' } });
  }

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

function computeTaskTotals() {
  const sh = getSheet(SH.BITACORA);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return {};
  const headers = data[0];
  const totals = {};
  for (let i = 1; i < data.length; i++) {
    const b = rowToObject(data[i], headers);
    if (!b.taskId) continue;
    const m = Number(b.minutes) || 0;
    totals[b.taskId] = (totals[b.taskId] || 0) + m;
  }
  return totals;
}

function computeTotalMinutesForTask(taskId) {
  const sh = getSheet(SH.BITACORA);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return 0;
  const headers = data[0];
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    const b = rowToObject(data[i], headers);
    if (b.taskId === taskId) total += Number(b.minutes) || 0;
  }
  return total;
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
  }
  const existing = rowToObject(sh.getRange(row.rowIndex, 1, 1, headers.length).getValues()[0], headers);
  const merged = Object.assign({}, existing, {
    nombre: item.nombre ?? existing.nombre,
    activo: typeof item.activo === 'boolean' ? item.activo : asBool(existing.activo)
  });
  sh.getRange(row.rowIndex, 1, 1, headers.length).setValues([objectToRow(merged, headers)]);
  return merged;
}

function actionDeleteConfig(body) {
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
// AI — CLAUDE
// =====================================================
function actionParseAI(body) {
  const text = (body.text || '').trim();
  if (!text) return { tasks: [] };
  const parsed = callClaude(text);
  const saved = parsed.map(p => actionSaveTask({
    task: {
      title:    p.title || 'Tarea sin título',
      desc:     p.desc || '',
      project:  p.project || '',
      priority: p.priority || 'Media',
      tool:     p.tool || '',
      status:   normalizeStatus(p.status) || 'Por iniciar',
      blocker:  p.blocker || '',
      dueDate:  p.dueDate || '',
      source:   'ai'
    }
  }));
  return { tasks: saved };
}

function callClaude(text) {
  const apiKey = prop('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY no configurada en Script Properties');

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
    'ESTADOS válidos para captura inicial: Por iniciar, En progreso, En espera',
    'PRIORIDADES: Alta, Media, Baja',
    'BLOQUEADORES posibles: ' + bloqueadores.join(', '),
    '',
    'REGLAS:',
    '- Si menciona "Pa" o "urgente" → priority Alta.',
    '- Si menciona esperar a alguien o "bloqueado por" → status En espera + blocker apropiado.',
    '- Si menciona CAD/Revit/dibujo/Excel/estimación → status En progreso + tool apropiado.',
    '- Si menciona ir/visitar/inspeccionar/MOP/alcaldía → tool "Campo presencial" + status Por iniciar.',
    '- Si no hay info clara → status Por iniciar.',
    '- Si menciona ChatGPT/IA/Claude/análisis con IA → tool "IA".',
    '- desc: detalles adicionales útiles, NO repitas el title.',
    '- dueDate: solo si menciona fecha explícita (formato YYYY-MM-DD).',
    '- Una sola tarea por idea separada. Si el texto tiene varias tareas separadas por coma/punto/"y", devolvé varias.',
    '- Title corto y accionable, en imperativo cuando posible.',
    '',
    'Devolvé SOLO un array JSON con esta estructura, sin texto adicional, sin markdown, sin code fences:',
    '[{"title":"","desc":"","project":"","priority":"Alta|Media|Baja","tool":"","status":"Por iniciar|En progreso|En espera","blocker":"","dueDate":""}]'
  ].join('\n');

  const payload = {
    model: 'claude-haiku-4-5',
    max_tokens: 2048,
    system: systemPrompt,
    messages: [{ role: 'user', content: text }]
  };

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const txt = res.getContentText();
  if (code !== 200) throw new Error('Anthropic ' + code + ': ' + txt.slice(0, 300));
  const data = JSON.parse(txt);
  const raw = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('').trim();
  const cleaned = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
  let arr;
  try { arr = JSON.parse(cleaned); } catch (e) { throw new Error('Claude devolvió JSON inválido: ' + cleaned.slice(0,200)); }
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
  const isOpen = t => ['Por iniciar','En progreso','En espera','Pausado'].indexOf(t.status) >= 0;

  const overdue   = tasks.filter(t => isOpen(t) && t.dueDate && fmtDate(t.dueDate) < today);
  const dueToday  = tasks.filter(t => isOpen(t) && t.dueDate && fmtDate(t.dueDate) === today);
  const dueWeek   = tasks.filter(t => isOpen(t) && t.dueDate && fmtDate(t.dueDate) > today && fmtDate(t.dueDate) <= inSevenDays);
  const blocked   = tasks.filter(t => t.status === 'En espera');
  const paused    = tasks.filter(t => t.status === 'Pausado');
  const urgentND  = tasks.filter(t => isOpen(t) && t.priority === 'Alta' && !t.dueDate);

  const sections = [
    { title: '🔴 VENCIDAS',           items: overdue },
    { title: '🟡 PARA HOY',           items: dueToday },
    { title: '📅 ESTA SEMANA',        items: dueWeek },
    { title: '⏳ EN ESPERA',          items: blocked },
    { title: '⏸ PAUSADAS',            items: paused },
    { title: '⚡ URGENTES SIN FECHA', items: urgentND }
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
  const entries = listBitacoraByDate(today).filter(b => b.endTime);

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
  const workedTodayIds = new Set(entries.map(e => e.taskId));
  // evidencia pendiente: tareas con evidencePending=true trabajadas hoy O cerradas hoy
  const evidencePending = tasks.filter(t => {
    if (!asBool(t.evidencePending)) return false;
    const closedToday = (t.completedAt && String(t.completedAt).slice(0,10) === today) ||
                        (t.cancelledAt && String(t.cancelledAt).slice(0,10) === today);
    return workedTodayIds.has(t.id) || closedToday;
  });

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
        evidencePending.map(t => {
          const parts = [escapeHtml(t.title)];
          const meta = [];
          if (t.project)   meta.push(escapeHtml(t.project));
          if (t.status)    meta.push(escapeHtml(t.status));
          if (meta.length) parts.push(' <span style="color:#888">— ' + meta.join(' · ') + '</span>');
          if (t.evidenceLink) {
            if (/^https?:/i.test(t.evidenceLink)) {
              parts.push('<br><a href="' + escapeHtml(t.evidenceLink) + '" style="color:#3a7bd5;font-size:12px">🔗 ' + escapeHtml(t.evidenceLink.slice(0,80)) + '</a>');
            } else {
              parts.push('<br><code style="color:#888;font-size:11px;background:#eee;padding:1px 4px;border-radius:3px">' + escapeHtml(t.evidenceLink) + '</code>');
            }
          }
          return '<li style="margin-bottom:6px">' + parts.join('') + '</li>';
        }).join('') +
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
        if (t.project)            meta.push(escapeHtml(t.project));
        if (t.priority === 'Alta') meta.push('Alta');
        if (t.dueDate)            meta.push(fmtDate(t.dueDate));
        if (t.blocker)            meta.push('⏳ ' + escapeHtml(t.blocker));
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
// SETUP — correr manualmente
// =====================================================
function setupAll() {
  ensureSheets();
  migrateSheets();
  installTriggers();
  Logger.log('Setup v3 completo. Verificá Script Properties: API_TOKEN, ANTHROPIC_API_KEY, SHEET_ID, EMAIL_TO');
}

function ensureSheets() {
  const ss = SpreadsheetApp.openById(prop('SHEET_ID'));
  ensureSheet(ss, SH.TAREAS,    HEADERS.TAREAS);
  ensureSheet(ss, SH.BITACORA,  HEADERS.BITACORA);
  ensureSheet(ss, SH.HISTORIAL, HEADERS.HISTORIAL);
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

/**
 * Migra hojas v2 → v3 sin destruir datos.
 * Idempotente: se puede correr varias veces.
 */
function migrateSheets() {
  const ss = SpreadsheetApp.openById(prop('SHEET_ID'));

  // Tareas: insertar cancelledAt + justificante después de completedAt
  const tareas = ss.getSheetByName(SH.TAREAS);
  if (tareas && tareas.getLastColumn() > 0) {
    const lastCol = tareas.getLastColumn();
    const headers = tareas.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h));
    const hasCanc = headers.indexOf('cancelledAt') !== -1;
    const hasJust = headers.indexOf('justificante') !== -1;
    if (!hasCanc || !hasJust) {
      const completedIdx = headers.indexOf('completedAt');
      if (completedIdx !== -1) {
        const insertAfter = completedIdx + 1; // 1-based column
        tareas.insertColumnsAfter(insertAfter, 2);
        tareas.getRange(1, insertAfter + 1).setValue('cancelledAt').setFontWeight('bold').setBackground('#1a1e2b').setFontColor('#e8a020');
        tareas.getRange(1, insertAfter + 2).setValue('justificante').setFontWeight('bold').setBackground('#1a1e2b').setFontColor('#e8a020');
        Logger.log('Migración Tareas: agregadas cancelledAt + justificante');
      }
    }
  }

  // Bitacora: insertar anio + mes + semana después de project
  const bita = ss.getSheetByName(SH.BITACORA);
  if (bita && bita.getLastColumn() > 0) {
    const lastCol = bita.getLastColumn();
    const headers = bita.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h));
    if (headers.indexOf('anio') === -1) {
      const projIdx = headers.indexOf('project');
      if (projIdx !== -1) {
        const insertAfter = projIdx + 1; // 1-based
        bita.insertColumnsAfter(insertAfter, 3);
        bita.getRange(1, insertAfter + 1).setValue('anio').setFontWeight('bold').setBackground('#1a1e2b').setFontColor('#e8a020');
        bita.getRange(1, insertAfter + 2).setValue('mes').setFontWeight('bold').setBackground('#1a1e2b').setFontColor('#e8a020');
        bita.getRange(1, insertAfter + 3).setValue('semana').setFontWeight('bold').setBackground('#1a1e2b').setFontColor('#e8a020');
        // Backfill: leer date (que se desplazó 3 columnas) y calcular anio/mes/semana
        const lastRow = bita.getLastRow();
        if (lastRow > 1) {
          const dateColIdx = headers.indexOf('date') + 1 + 3;
          const dates = bita.getRange(2, dateColIdx, lastRow - 1, 1).getValues();
          const fills = dates.map(([d]) => {
            if (!d) return ['', '', ''];
            const dt = d instanceof Date ? d : new Date(d);
            if (isNaN(dt)) return ['', '', ''];
            return [dt.getFullYear(), dt.getMonth() + 1, getIsoWeek(dt)];
          });
          if (fills.length) bita.getRange(2, insertAfter + 1, fills.length, 3).setValues(fills);
        }
        Logger.log('Migración Bitacora: agregadas anio/mes/semana con backfill');
      }
    }
  }

  // Normalizar estados legacy en Tareas
  const tareasSh = ss.getSheetByName(SH.TAREAS);
  if (tareasSh && tareasSh.getLastRow() > 1) {
    const headers = tareasSh.getRange(1, 1, 1, tareasSh.getLastColumn()).getValues()[0].map(h => String(h));
    const statusIdx = headers.indexOf('status');
    if (statusIdx !== -1) {
      const range = tareasSh.getRange(2, statusIdx + 1, tareasSh.getLastRow() - 1, 1);
      const vals = range.getValues();
      let changed = false;
      for (let i = 0; i < vals.length; i++) {
        const s = String(vals[i][0] || '');
        if (LEGACY_STATUS_MAP[s]) { vals[i][0] = LEGACY_STATUS_MAP[s]; changed = true; }
      }
      if (changed) {
        range.setValues(vals);
        Logger.log('Migración Tareas: estados legacy normalizados');
      }
    }
  }
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
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return { rowIndex: i + 2 };
  }
  return null;
}

function rowToObject(row, headers) {
  const o = {};
  headers.forEach((h, i) => o[h] = row[i]);
  ['createdAt','updatedAt','completedAt','cancelledAt','startTime','endTime'].forEach(k => {
    if (o[k] instanceof Date) o[k] = o[k].toISOString();
  });
  ['startDate','dueDate','date','fecha'].forEach(k => {
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

function normalizeStatus(s) {
  if (!s) return 'Por iniciar';
  if (LEGACY_STATUS_MAP[s]) return LEGACY_STATUS_MAP[s];
  if (STATUSES.indexOf(s) >= 0) return s;
  return 'Por iniciar';
}

function getIsoWeek(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
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
