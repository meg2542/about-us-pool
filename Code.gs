/**
 * ABOUT US POOL QUEUE — V1
 * Separate system from POS.
 * Google Apps Script + Google Sheets + HTML Service
 *
 * First run:
 *   1) Run setupPoolQueue() once.
 *   2) Deploy as Web App: Execute as Me / Who has access: Anyone.
 */

const POOL_CONFIG = {
  APP_NAME: 'ABOUT US POOL QUEUE',
  TIMEZONE: 'Asia/Bangkok',
  DEFAULTS: {
    IS_OPEN: 'TRUE',
    MAX_QUEUE: '15',
    CALL_TIMEOUT_MIN: '5',
    RESET_HOUR: '18',
    AVG_GAME_MIN: '15'
  },
  QUEUE_HEADERS: [
    'ID', 'ShiftKey', 'QueueNo', 'Name', 'ClientToken', 'Status',
    'JoinedAt', 'CalledAt', 'AcknowledgedAt', 'StartedAt',
    'FinishedAt', 'CancelledAt', 'Source', 'Note'
  ],
  LOG_HEADERS: ['Timestamp', 'Action', 'QueueID', 'QueueNo', 'Name', 'Detail']
};

const ACTIVE_STATUSES = ['WAITING', 'CALLED', 'PLAYING'];
const BACKEND_VERSION = 'V1.7.0-PUSH';
const POOL_WEB_URL = 'https://meg2542.github.io/about-us-pool/';
const PUSH_EXTERNAL_PREFIX = 'poolq_';

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};

  // API mode used by the GitHub Pages frontend.
  // The browser stays on github.io; Apps Script is only the data backend.
  if (String(p.api || '') === '1') {
    return handleJsonpApi_(p);
  }

  // Direct Apps Script pages remain available as a fallback,
  // but the QR no longer opens these URLs.
  const page = p.page === 'admin' ? 'Admin' : 'Index';
  const title = page === 'Admin' ? 'Pool Queue Admin' : 'About Us Pool Queue';
  return HtmlService.createTemplateFromFile(page)
    .evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * JSONP API for the GitHub Pages frontend.
 * Apps Script ContentService officially supports JSONP for browser calls.
 */
function handleJsonpApi_(p) {
  const callback = sanitizeJsonpCallback_(p.callback);
  let response;

  try {
    const route = String(p.route || '');

    if (route === 'publicState') {
      response = {
        ok: true,
        data: getPublicState(String(p.clientToken || ''))
      };

    } else if (route === 'customerAction') {
      const action = String(p.action || '');
      const payload = {};
      if (action === 'JOIN') payload.name = String(p.name || '');

      response = {
        ok: true,
        data: customerAction(
          String(p.clientToken || ''),
          action,
          payload
        )
      };

    } else if (route === 'adminState') {
      response = {
        ok: true,
        data: adminGetState()
      };

    } else if (route === 'adminAction') {
      const action = String(p.action || '');
      const payload = {};

      if (action === 'SET_OPEN') {
        payload.isOpen = String(p.isOpen || '').toLowerCase() === 'true';
      }
      if (action === 'ADD_QUEUE') {
        payload.name = String(p.name || '');
      }
      if (['START', 'SKIP', 'CANCEL', 'MOVE_END'].includes(action)) {
        payload.id = String(p.id || '');
      }

      response = {
        ok: true,
        data: adminAction(action, payload)
      };

    } else if (route === 'debugState') {
      response = {
        ok: true,
        data: buildDebugState_()
      };

    } else {
      throw new Error('API route ไม่ถูกต้อง: ' + route);
    }

  } catch (err) {
    response = {
      ok: false,
      error: err && err.message ? err.message : 'เกิดข้อผิดพลาด'
    };
  }

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(response) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * Write API for GitHub Pages.
 * The frontend sends a normal form POST in no-cors mode, then reads the
 * result through the read-only JSONP writeResult route.
 */
function doPost(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const requestId = sanitizeRequestId_(p.requestId);
  let result;

  try {
    const route = String(p.route || '');

    if (route === 'customerAction') {
      const action = String(p.action || '');
      const payload = {};
      if (action === 'JOIN') payload.name = String(p.name || '');

      customerAction(
        String(p.clientToken || ''),
        action,
        payload
      );

      result = { ok: true };

    } else if (route === 'adminAction') {
      const action = String(p.action || '');
      const payload = {};

      if (action === 'SET_OPEN') {
        payload.isOpen = String(p.isOpen || '').toLowerCase() === 'true';
      }
      if (action === 'ADD_QUEUE') {
        payload.name = String(p.name || '');
      }
      if (['START', 'SKIP', 'CANCEL', 'MOVE_END'].includes(action)) {
        payload.id = String(p.id || '');
      }

      adminAction(action, payload);
      result = { ok: true };

    } else {
      throw new Error('Write route ไม่ถูกต้อง');
    }

  } catch (err) {
    result = {
      ok: false,
      error: err && err.message ? err.message : 'เกิดข้อผิดพลาด'
    };
  }

  CacheService.getScriptCache().put(
    'POOL_WRITE_' + requestId,
    JSON.stringify(result),
    60
  );

  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function sanitizeRequestId_(raw) {
  const id = String(raw || '');
  if (!/^[A-Za-z0-9_-]{8,100}$/.test(id)) {
    throw new Error('Invalid request id');
  }
  return id;
}

function sanitizeJsonpCallback_(raw) {
  const callback = String(raw || '');
  if (!/^[A-Za-z_$][0-9A-Za-z_$]{0,80}$/.test(callback)) {
    throw new Error('Invalid callback');
  }
  return callback;
}

/**
 * Run once from Apps Script editor.
 * Creates a dedicated spreadsheet automatically and stores its ID in Script Properties.
 */
function setupPoolQueue() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('POOL_SPREADSHEET_ID');
  let ss;

  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create('ABOUT US POOL QUEUE DATA');
    ssId = ss.getId();
    props.setProperty('POOL_SPREADSHEET_ID', ssId);
  }

  ensureSheets_(ss);

  // Clean up legacy PIN setting from older V1 builds.
  props.deleteProperty('ADMIN_PIN');

  const settings = getSettingsSheet_();
  const existing = readSettings_();
  Object.keys(POOL_CONFIG.DEFAULTS).forEach(function(key) {
    if (!(key in existing)) settings.appendRow([key, POOL_CONFIG.DEFAULTS[key]]);
  });

  formatSheets_(ss);

  const result = {
    ok: true,
    spreadsheetUrl: ss.getUrl(),
    message: 'Setup complete. Backend ' + BACKEND_VERSION
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}


// ---------- Public / Customer API ----------

function getPublicState(clientToken) {
  expireCalledIfNeededThrottled_();
  return buildPublicState_(String(clientToken || ''));
}

function customerAction(clientToken, action, payload) {
  const token = String(clientToken || '').trim();
  if (!token) throw new Error('ไม่พบรหัสเครื่อง กรุณารีเฟรชหน้าเว็บ');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const act = String(action || '').trim().toUpperCase();
    const data = payload || {};

    if (act === 'JOIN') {
      joinQueue_(token, data.name);
    } else if (act === 'CANCEL') {
      cancelMyQueue_(token);
    } else if (act === 'ACK') {
      acknowledgeMyQueue_(token);
    } else if (act === 'TEST_PUSH') {
      sendPushToClientToken_(token, 'ทดสอบแจ้งเตือน 🎱', 'ระบบแจ้งเตือน ABOUT US POOL พร้อมใช้งานแล้ว', null);
    } else {
      throw new Error('คำสั่งไม่ถูกต้อง: ' + act);
    }

    // Commit while the lock is still held so another request cannot read
    // half-finished queue data or allocate the same queue number.
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return buildPublicState_(token);
}

// ---------- Admin API ----------

function adminGetState() {
  expireCalledIfNeededThrottled_();
  return buildAdminState_();
}

function adminAction(action, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const act = String(action || '').trim().toUpperCase();
    const data = payload || {};

    if (act === 'SET_OPEN') {
      setSetting_('IS_OPEN', data.isOpen ? 'TRUE' : 'FALSE');
      log_('SET_OPEN', null, 'isOpen=' + !!data.isOpen);
    } else if (act === 'ADD_QUEUE') {
      adminAddQueue_(data.name);
    } else if (act === 'CALL_NEXT') {
      if (findCurrentPlaying_()) throw new Error('ยังมีคิวกำลังเล่นอยู่');
      if (findCalled_()) throw new Error('มีคิวที่กำลังถูกเรียกอยู่แล้ว');
      callNextWaiting_();
    } else if (act === 'FINISH_CALL_NEXT') {
      finishCurrentAndCallNext_();
    } else if (act === 'START') {
      startCalledQueue_(data.id);
    } else if (act === 'SKIP') {
      skipQueue_(data.id);
    } else if (act === 'CANCEL') {
      adminCancelQueue_(data.id);
    } else if (act === 'MOVE_END') {
      moveQueueToEnd_(data.id);
    } else {
      throw new Error('คำสั่ง Admin ไม่ถูกต้อง: ' + act);
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return buildAdminState_();
}

// ---------- Queue mutations ----------

function joinQueue_(token, rawName) {
  const settings = readSettings_();
  if (!asBool_(settings.IS_OPEN)) throw new Error('ขณะนี้ปิดรับคิวชั่วคราว');

  const rows = readRecentQueueRows_();
  const currentRows = filterRowsCurrentShift_(rows, settings);
  const active = currentRows.filter(r => ACTIVE_STATUSES.includes(r.status));

  const existing = active.find(r => r.clientToken === token);
  if (existing) return existing;

  const waitingCount = active.filter(r => r.status === 'WAITING' || r.status === 'CALLED').length;
  const maxQueue = asInt_(settings.MAX_QUEUE, 15);
  if (waitingCount >= maxQueue) throw new Error('คิวเต็มแล้ว กรุณาลองใหม่ภายหลัง');

  const name = sanitizeName_(rawName);
  if (!name) throw new Error('กรุณาใส่ชื่อเล่น');

  const shiftKey = getShiftKeyFromSettings_(settings);
  const queueNo = nextQueueNo_(shiftKey, currentRows);
  const q = appendQueue_(name, token, 'CUSTOMER', shiftKey, queueNo);
  log_('JOIN', q, 'customer');
  return q;
}

function adminAddQueue_(rawName) {
  const settings = readSettings_();
  const rows = readRecentQueueRows_();
  const currentRows = filterRowsCurrentShift_(rows, settings);
  const active = currentRows.filter(r => ACTIVE_STATUSES.includes(r.status));
  const waitingCount = active.filter(r => r.status === 'WAITING' || r.status === 'CALLED').length;
  const maxQueue = asInt_(settings.MAX_QUEUE, 15);
  if (waitingCount >= maxQueue) throw new Error('คิวเต็มตามจำนวนสูงสุดที่ตั้งไว้');

  const name = sanitizeName_(rawName);
  if (!name) throw new Error('กรุณาใส่ชื่อเล่น');

  const shiftKey = getShiftKeyFromSettings_(settings);
  const queueNo = nextQueueNo_(shiftKey, currentRows);
  const q = appendQueue_(name, '', 'ADMIN', shiftKey, queueNo);
  log_('JOIN', q, 'admin');
}

function appendQueue_(name, clientToken, source, shiftKey, queueNo) {
  const sheet = getQueueSheet_();
  const id = Utilities.getUuid();
  const now = new Date();
  const row = sheet.getLastRow() + 1;

  // Keep ShiftKey as text, but current-shift detection also uses JoinedAt,
  // so old rows remain readable even if Sheets converted ShiftKey to a Date.
  sheet.getRange(row, 2).setNumberFormat('@');
  sheet.getRange(row, 1, 1, POOL_CONFIG.QUEUE_HEADERS.length).setValues([[
    id, String(shiftKey), queueNo, name, clientToken, 'WAITING',
    now, '', '', '', '', '', source, ''
  ]]);

  return {
    id: id, shiftKey: String(shiftKey), queueNo: queueNo, name: name,
    clientToken: clientToken, status: 'WAITING', joinedAt: now,
    calledAt: null, acknowledgedAt: null, startedAt: null,
    finishedAt: null, cancelledAt: null,
    source: source, note: '', row: row
  };
}

function cancelMyQueue_(token) {
  const active = getActiveCurrentShift_();
  const q = active.find(r => r.clientToken === token);
  if (!q) throw new Error('ไม่พบคิวที่กำลังใช้งาน');

  if (q.status === 'PLAYING') throw new Error('คิวนี้กำลังเล่นอยู่ กรุณาแจ้งพนักงาน');
  const wasCalled = q.status === 'CALLED';

  updateQueueRow_(q.row, {
    Status: 'CANCELLED',
    CancelledAt: new Date(),
    Note: 'Cancelled by customer'
  });
  log_('CANCEL_CUSTOMER', q, '');

  if (wasCalled && !findCurrentPlaying_()) callNextWaiting_();
}

function acknowledgeMyQueue_(token) {
  const active = getActiveCurrentShift_();
  const q = active.find(r => r.clientToken === token);
  if (!q || q.status !== 'CALLED') throw new Error('คิวนี้ยังไม่ได้ถูกเรียก');

  if (!q.acknowledgedAt) {
    updateQueueRow_(q.row, { AcknowledgedAt: new Date() });
    log_('ACK', q, 'customer is coming');
  }
}

function finishCurrentAndCallNext_() {
  const current = findCurrentPlaying_();
  if (!current) {
    if (findCalled_()) throw new Error('ตอนนี้มีคิวที่กำลังถูกเรียกอยู่');
    return callNextWaiting_();
  }

  updateQueueRow_(current.row, {
    Status: 'DONE',
    FinishedAt: new Date()
  });
  log_('FINISH', current, '');

  return callNextWaiting_();
}

function callNextWaiting_() {
  const waiting = getActiveCurrentShift_()
    .filter(r => r.status === 'WAITING')
    .sort((a, b) => a.joinedAt.getTime() - b.joinedAt.getTime());

  if (!waiting.length) return null;

  const next = waiting[0];
  updateQueueRow_(next.row, {
    Status: 'CALLED',
    CalledAt: new Date(),
    AcknowledgedAt: ''
  });
  log_('CALL', next, '');

  // Push is best-effort only. Queue flow must keep working even if
  // OneSignal is not configured or its API is temporarily unavailable.
  if (next.clientToken) {
    const settings = readSettings_();
    const timeoutMin = asInt_(settings.CALL_TIMEOUT_MIN, 5);
    sendPushToClientToken_(
      next.clientToken,
      'ถึงคิว ' + next.queueNo + ' แล้ว 🎱',
      'กรุณามาที่โต๊ะพูลภายใน ' + timeoutMin + ' นาที',
      next
    );
  }

  return next;
}

function startCalledQueue_(id) {
  if (findCurrentPlaying_()) throw new Error('มีคิวกำลังเล่นอยู่แล้ว');

  const q = findActiveById_(id);
  if (!q || q.status !== 'CALLED') throw new Error('ไม่พบคิวที่กำลังถูกเรียก');

  updateQueueRow_(q.row, {
    Status: 'PLAYING',
    StartedAt: new Date()
  });
  log_('START', q, '');
}

function skipQueue_(id) {
  const q = findActiveById_(id);
  if (!q) throw new Error('ไม่พบคิวนี้');
  if (q.status === 'PLAYING') throw new Error('คิวกำลังเล่นอยู่ ใช้ปุ่มจบเกมแทน');

  const wasCalled = q.status === 'CALLED';
  updateQueueRow_(q.row, {
    Status: 'SKIPPED',
    CancelledAt: new Date(),
    Note: 'Skipped by admin'
  });
  log_('SKIP', q, '');

  if (wasCalled && !findCurrentPlaying_()) callNextWaiting_();
}

function adminCancelQueue_(id) {
  const q = findActiveById_(id);
  if (!q) throw new Error('ไม่พบคิวนี้');
  if (q.status === 'PLAYING') throw new Error('คิวกำลังเล่นอยู่ ใช้ปุ่มจบเกมแทน');

  const wasCalled = q.status === 'CALLED';
  updateQueueRow_(q.row, {
    Status: 'CANCELLED',
    CancelledAt: new Date(),
    Note: 'Cancelled by admin'
  });
  log_('CANCEL_ADMIN', q, '');

  if (wasCalled && !findCurrentPlaying_()) callNextWaiting_();
}

function moveQueueToEnd_(id) {
  const q = findActiveById_(id);
  if (!q || q.status !== 'WAITING') throw new Error('ย้ายได้เฉพาะคิวที่กำลังรอ');

  updateQueueRow_(q.row, {
    JoinedAt: new Date(),
    Note: 'Moved to end by admin'
  });
  log_('MOVE_END', q, '');
}

// ---------- Automatic timeout ----------

function expireCalledIfNeededThrottled_() {
  const cache = CacheService.getScriptCache();
  const key = 'POOL_EXPIRE_CHECK_V16';
  if (cache.get(key)) return;
  cache.put(key, '1', 4);
  expireCalledIfNeeded_();
}


function expireCalledIfNeeded_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1500)) return;

  try {
    const called = findCalled_();
    if (!called || !called.calledAt) return;

    const settings = readSettings_();
    const timeoutMin = asInt_(settings.CALL_TIMEOUT_MIN, 5);
    const elapsed = (Date.now() - called.calledAt.getTime()) / 60000;

    if (elapsed >= timeoutMin) {
      updateQueueRow_(called.row, {
        Status: 'SKIPPED',
        CancelledAt: new Date(),
        Note: 'Auto skipped after call timeout'
      });
      log_('AUTO_SKIP', called, timeoutMin + ' min');
      if (!findCurrentPlaying_()) callNextWaiting_();
    }
  } finally {
    lock.releaseLock();
  }
}


function getShiftKeyFromSettings_(settings) {
  const resetHour = asInt_(settings.RESET_HOUR, 18);
  const shifted = new Date(Date.now() - resetHour * 60 * 60 * 1000);
  return Utilities.formatDate(shifted, POOL_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function getShiftWindowFromSettings_(settings) {
  const resetHour = Math.max(0, Math.min(23, asInt_(settings.RESET_HOUR, 18)));
  const shiftKey = getShiftKeyFromSettings_(settings);
  const hourText = String(resetHour).padStart(2, '0');
  const start = Utilities.parseDate(
    shiftKey + ' ' + hourText + ':00:00',
    POOL_CONFIG.TIMEZONE,
    'yyyy-MM-dd HH:mm:ss'
  );
  return {
    shiftKey: shiftKey,
    start: start,
    end: new Date(start.getTime() + 24 * 60 * 60 * 1000)
  };
}

function filterRowsCurrentShift_(rows, settings) {
  const w = getShiftWindowFromSettings_(settings);
  return (rows || []).filter(function(r) {
    // JoinedAt is the authoritative source because Sheets may auto-format
    // ShiftKey. Fallback to ShiftKey only for legacy rows without JoinedAt.
    if (r.joinedAt instanceof Date && !isNaN(r.joinedAt.getTime())) {
      return r.joinedAt.getTime() >= w.start.getTime() &&
             r.joinedAt.getTime() < w.end.getTime();
    }
    return normalizeShiftKey_(r.shiftKey) === w.shiftKey;
  });
}

function getAverageGameMinutesFromRows_(rows, settings) {
  const fallback = asInt_(settings.AVG_GAME_MIN, 15);
  const done = rows
    .filter(r => r.status === 'DONE' && r.startedAt && r.finishedAt)
    .slice(-20);

  if (done.length < 3) return fallback;

  const mins = done
    .map(r => (r.finishedAt.getTime() - r.startedAt.getTime()) / 60000)
    .filter(n => n >= 3 && n <= 120);

  if (mins.length < 3) return fallback;
  return Math.max(5, Math.round(mins.reduce((a, b) => a + b, 0) / mins.length));
}

// ---------- State builders ----------

function buildPublicState_(token) {
  const settings = readSettings_();
  const rows = readRecentQueueRows_();
  const currentRows = filterRowsCurrentShift_(rows, settings);
  const active = currentRows.filter(r => ACTIVE_STATUSES.includes(r.status));

  const playing = active.find(r => r.status === 'PLAYING') || null;
  const called = active.find(r => r.status === 'CALLED') || null;
  const waiting = active
    .filter(r => r.status === 'WAITING')
    .sort((a, b) => a.joinedAt.getTime() - b.joinedAt.getTime());

  const mine = token ? (active.find(r => r.clientToken === token) || null) : null;
  const avgMin = getAverageGameMinutesFromRows_(rows, settings);
  const maxQueue = asInt_(settings.MAX_QUEUE, 15);
  const waitingCount = waiting.length + (called ? 1 : 0);

  let mineView = null;
  if (mine) {
    const ordered = [];
    if (called) ordered.push(called);
    waiting.forEach(q => ordered.push(q));

    let ahead = 0;
    if (mine.status === 'WAITING') {
      const idx = ordered.findIndex(q => q.id === mine.id);
      ahead = idx < 0 ? 0 : idx;
    }

    const gamesBefore = mine.status === 'WAITING'
      ? ahead + (playing ? 1 : 0)
      : 0;

    const estimate = gamesBefore > 0
      ? {
          min: Math.max(1, Math.round(gamesBefore * avgMin * 0.8)),
          max: Math.max(2, Math.round(gamesBefore * avgMin * 1.2))
        }
      : { min: 0, max: 5 };

    mineView = publicQueue_(mine);
    mineView.ahead = ahead;
    mineView.estimateMin = estimate.min;
    mineView.estimateMax = estimate.max;
    mineView.callRemainingSec = mine.status === 'CALLED'
      ? getCallRemainingSec_(mine, settings)
      : null;
  }

  return {
    backendVersion: BACKEND_VERSION,
    appName: POOL_CONFIG.APP_NAME,
    isOpen: asBool_(settings.IS_OPEN),
    maxQueue: maxQueue,
    waitingCount: waitingCount,
    canJoin: asBool_(settings.IS_OPEN) && waitingCount < maxQueue && !mine,
    currentPlaying: playing ? publicQueue_(playing) : null,
    called: called ? publicQueue_(called) : null,
    mine: mineView,
    avgGameMin: avgMin,
    pushConfigured: isPushConfigured_(),
    pushAppId: getPushAppId_(),
    serverTime: new Date().toISOString()
  };
}

function buildAdminState_() {
  const settings = readSettings_();
  const rows = readRecentQueueRows_();
  const shiftKey = getShiftKeyFromSettings_(settings);
  const currentRows = filterRowsCurrentShift_(rows, settings);
  const active = currentRows.filter(r => ACTIVE_STATUSES.includes(r.status));

  const playing = active.find(r => r.status === 'PLAYING') || null;
  const called = active.find(r => r.status === 'CALLED') || null;
  const waiting = active
    .filter(r => r.status === 'WAITING')
    .sort((a, b) => a.joinedAt.getTime() - b.joinedAt.getTime());

  return {
    backendVersion: BACKEND_VERSION,
    isOpen: asBool_(settings.IS_OPEN),
    maxQueue: asInt_(settings.MAX_QUEUE, 15),
    callTimeoutMin: asInt_(settings.CALL_TIMEOUT_MIN, 5),
    avgGameMin: getAverageGameMinutesFromRows_(rows, settings),
    shiftKey: shiftKey,
    playing: playing ? adminQueue_(playing, settings) : null,
    called: called ? adminQueue_(called, settings) : null,
    waiting: waiting.map(q => adminQueue_(q, settings)),
    pushConfigured: isPushConfigured_(),
    serverTime: new Date().toISOString()
  };
}

function publicQueue_(q) {
  return {
    id: q.id,
    queueNo: q.queueNo,
    name: q.name,
    status: q.status,
    joinedAt: toIso_(q.joinedAt),
    calledAt: toIso_(q.calledAt),
    acknowledgedAt: toIso_(q.acknowledgedAt),
    startedAt: toIso_(q.startedAt)
  };
}

function adminQueue_(q, settings) {
  const obj = publicQueue_(q);
  obj.source = q.source;
  obj.waitMin = q.joinedAt ? Math.max(0, Math.floor((Date.now() - q.joinedAt.getTime()) / 60000)) : 0;
  obj.playMin = q.startedAt ? Math.max(0, Math.floor((Date.now() - q.startedAt.getTime()) / 60000)) : 0;
  obj.callRemainingSec = q.status === 'CALLED' ? getCallRemainingSec_(q, settings) : null;
  return obj;
}

function getCallRemainingSec_(q, settings) {
  if (!q.calledAt) return null;
  const timeoutSec = asInt_(settings.CALL_TIMEOUT_MIN, 5) * 60;
  const elapsed = Math.floor((Date.now() - q.calledAt.getTime()) / 1000);
  return Math.max(0, timeoutSec - elapsed);
}

// ---------- Data helpers ----------

function getActiveCurrentShift_() {
  const settings = readSettings_();
  return filterRowsCurrentShift_(readRecentQueueRows_(), settings)
    .filter(r => ACTIVE_STATUSES.includes(r.status));
}

function findCurrentPlaying_() {
  return getActiveCurrentShift_().find(r => r.status === 'PLAYING') || null;
}

function findCalled_() {
  return getActiveCurrentShift_().find(r => r.status === 'CALLED') || null;
}

function findActiveById_(id) {
  const target = String(id || '');
  return getActiveCurrentShift_().find(r => r.id === target) || null;
}

function normalizeShiftKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, POOL_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }

  const text = String(value == null ? '' : value).trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return match[1] + '-' + match[2] + '-' + match[3];
  return text;
}

function readRecentQueueRows_() {
  const sheet = getQueueSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const width = POOL_CONFIG.QUEUE_HEADERS.length;
  const startRow = Math.max(2, lastRow - 299);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, width).getValues();

  return values.map(function(v, i) {
    return {
      row: startRow + i,
      id: String(v[0] || ''),
      shiftKey: normalizeShiftKey_(v[1]),
      queueNo: String(v[2] || ''),
      name: String(v[3] || ''),
      clientToken: String(v[4] || '').trim(),
      status: String(v[5] || '').trim().toUpperCase(),
      joinedAt: asDate_(v[6]),
      calledAt: asDate_(v[7]),
      acknowledgedAt: asDate_(v[8]),
      startedAt: asDate_(v[9]),
      finishedAt: asDate_(v[10]),
      cancelledAt: asDate_(v[11]),
      source: String(v[12] || ''),
      note: String(v[13] || '')
    };
  });
}

function updateQueueRow_(row, patch) {
  const sheet = getQueueSheet_();
  const col = {};
  POOL_CONFIG.QUEUE_HEADERS.forEach((h, i) => col[h] = i + 1);

  Object.keys(patch).forEach(function(key) {
    if (!col[key]) return;
    sheet.getRange(row, col[key]).setValue(patch[key]);
  });
}

function nextQueueNo_(shiftKey, currentRows) {
  const props = PropertiesService.getScriptProperties();
  const savedShift = props.getProperty('QUEUE_COUNTER_SHIFT');
  let counter = savedShift === shiftKey
    ? parseInt(props.getProperty('QUEUE_COUNTER_VALUE') || '0', 10)
    : 0;
  if (!Number.isFinite(counter)) counter = 0;

  let maxFromRows = 0;
  (currentRows || []).forEach(function(r) {
    const m = String(r.queueNo || '').match(/^P(\d+)$/i);
    if (m) maxFromRows = Math.max(maxFromRows, parseInt(m[1], 10) || 0);
  });

  const next = Math.max(counter, maxFromRows) + 1;
  props.setProperties({
    QUEUE_COUNTER_SHIFT: shiftKey,
    QUEUE_COUNTER_VALUE: String(next)
  });
  return 'P' + String(next).padStart(2, '0');
}

function getShiftKey_() {
  const settings = readSettings_();
  return getShiftKeyFromSettings_(settings);
}

function getAverageGameMinutes_() {
  const settings = readSettings_();
  return getAverageGameMinutesFromRows_(readRecentQueueRows_(), settings);
}

function buildDebugState_() {
  const settings = readSettings_();
  const rows = readRecentQueueRows_();
  const currentRows = filterRowsCurrentShift_(rows, settings);
  const active = currentRows.filter(r => ACTIVE_STATUSES.includes(r.status));
  const ss = getSpreadsheet_();
  const sheet = getQueueSheet_();

  return {
    backendVersion: BACKEND_VERSION,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    shiftKey: getShiftKeyFromSettings_(settings),
    queueLastRow: sheet.getLastRow(),
    recentRowCount: rows.length,
    currentShiftRowCount: currentRows.length,
    activeCount: active.length,
    lastRows: rows.slice(-5).map(function(r) {
      return {
        row: r.row,
        shiftKey: r.shiftKey,
        queueNo: r.queueNo,
        status: r.status,
        joinedAt: toIso_(r.joinedAt),
        source: r.source
      };
    }),
    serverTime: new Date().toISOString()
  };
}

// ---------- Settings / sheets ----------

function ensureSheets_(ss) {
  let q = ss.getSheetByName('Queue');
  if (!q) q = ss.insertSheet('Queue');
  if (q.getLastRow() === 0) q.appendRow(POOL_CONFIG.QUEUE_HEADERS);

  let log = ss.getSheetByName('Log');
  if (!log) log = ss.insertSheet('Log');
  if (log.getLastRow() === 0) log.appendRow(POOL_CONFIG.LOG_HEADERS);

  let settings = ss.getSheetByName('Settings');
  if (!settings) settings = ss.insertSheet('Settings');
  if (settings.getLastRow() === 0) settings.appendRow(['Key', 'Value']);

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 3) {
    try { ss.deleteSheet(defaultSheet); } catch (err) {}
  }
}

function formatSheets_(ss) {
  ['Queue', 'Log', 'Settings'].forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.setFrozenRows(1);
    if (sh.getLastColumn() > 0) {
      sh.getRange(1, 1, 1, sh.getLastColumn())
        .setFontWeight('bold')
        .setBackground('#111827')
        .setFontColor('#ffffff');
    }
  });

  const q = ss.getSheetByName('Queue');
  if (q) {
    q.setColumnWidth(3, 90);
    q.setColumnWidth(4, 140);
    q.setColumnWidth(6, 100);
  }
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('POOL_SPREADSHEET_ID');
  if (!id) throw new Error('ยังไม่ได้ Setup กรุณารัน setupPoolQueue() ก่อน');
  return SpreadsheetApp.openById(id);
}

function getQueueSheet_() {
  return getSpreadsheet_().getSheetByName('Queue');
}

function getSettingsSheet_() {
  return getSpreadsheet_().getSheetByName('Settings');
}

function getLogSheet_() {
  return getSpreadsheet_().getSheetByName('Log');
}

function readSettings_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'POOL_SETTINGS_V16';
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (err) {}
  }

  const sh = getSettingsSheet_();
  const last = sh.getLastRow();
  const obj = Object.assign({}, POOL_CONFIG.DEFAULTS);

  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function(r) {
      const key = String(r[0] || '').trim();
      if (key) obj[key] = String(r[1] == null ? '' : r[1]).trim();
    });
  }

  cache.put(cacheKey, JSON.stringify(obj), 30);
  return obj;
}

function setSetting_(key, value) {
  const sh = getSettingsSheet_();
  const last = sh.getLastRow();
  let found = false;

  if (last >= 2) {
    const values = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === key) {
        sh.getRange(i + 2, 2).setValue(value);
        found = true;
        break;
      }
    }
  }

  if (!found) sh.appendRow([key, value]);
  CacheService.getScriptCache().remove('POOL_SETTINGS_V16');
}



// ---------- Push notifications (OneSignal) ----------

function getPushAppId_() {
  return String(PropertiesService.getScriptProperties().getProperty('ONESIGNAL_APP_ID') || '').trim();
}

function getPushRestApiKey_() {
  return String(PropertiesService.getScriptProperties().getProperty('ONESIGNAL_REST_API_KEY') || '').trim();
}

function isPushConfigured_() {
  return !!(getPushAppId_() && getPushRestApiKey_());
}

function pushExternalId_(clientToken) {
  const token = String(clientToken || '').trim();
  return token ? PUSH_EXTERNAL_PREFIX + token : '';
}

function sendPushToClientToken_(clientToken, heading, content, queueObj) {
  const appId = getPushAppId_();
  const apiKey = getPushRestApiKey_();
  const externalId = pushExternalId_(clientToken);

  // Never let Push configuration break queue operations.
  if (!appId || !apiKey || !externalId) {
    return { ok: false, skipped: true, reason: 'PUSH_NOT_CONFIGURED' };
  }

  const payload = {
    app_id: appId,
    target_channel: 'push',
    include_aliases: {
      external_id: [externalId]
    },
    headings: {
      en: String(heading || 'ABOUT US POOL'),
      th: String(heading || 'ABOUT US POOL')
    },
    contents: {
      en: String(content || 'ถึงคิวของคุณแล้ว'),
      th: String(content || 'ถึงคิวของคุณแล้ว')
    },
    url: POOL_WEB_URL,
    data: {
      type: queueObj ? 'POOL_CALLED' : 'POOL_TEST',
      queueId: queueObj ? String(queueObj.id || '') : '',
      queueNo: queueObj ? String(queueObj.queueNo || '') : ''
    }
  };

  try {
    const res = UrlFetchApp.fetch('https://api.onesignal.com/notifications', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Key ' + apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const status = res.getResponseCode();
    const body = res.getContentText() || '';

    if (status >= 200 && status < 300) {
      if (queueObj) log_('PUSH_SENT', queueObj, 'status=' + status);
      return { ok: true, status: status, body: body };
    }

    console.warn('OneSignal push failed: HTTP ' + status + ' ' + body);
    if (queueObj) log_('PUSH_FAIL', queueObj, 'HTTP ' + status);
    return { ok: false, status: status, body: body };

  } catch (err) {
    console.warn('OneSignal push exception: ' + (err && err.message ? err.message : err));
    if (queueObj) log_('PUSH_FAIL', queueObj, err && err.message ? err.message : 'exception');
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

// ---------- Logging ----------

function log_(action, q, detail) {
  try {
    getLogSheet_().appendRow([
      new Date(),
      action || '',
      q ? q.id : '',
      q ? q.queueNo : '',
      q ? q.name : '',
      detail || ''
    ]);
  } catch (err) {
    console.warn('Log failed: ' + err.message);
  }
}

// ---------- Utilities ----------

function sanitizeName_(raw) {
  let s = String(raw || '').trim().replace(/[\r\n\t]+/g, ' ').replace(/\s{2,}/g, ' ');
  s = s.substring(0, 30);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function asBool_(v) {
  return String(v || '').toUpperCase() === 'TRUE';
}

function asInt_(v, fallback) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : fallback;
}

function asDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function toIso_(d) {
  return d instanceof Date && !isNaN(d.getTime()) ? d.toISOString() : null;
}
