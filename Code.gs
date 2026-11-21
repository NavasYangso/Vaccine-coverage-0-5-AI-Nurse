/**
 * ระบบติดตามความครอบคลุมวัคซีนเด็ก 0-5 ปี (Google Apps Script backend)
 * 1) สร้าง Google Sheet แล้วเปิด Extensions > Apps Script
 * 2) วางไฟล์นี้และรัน setupSheets() หนึ่งครั้ง
 * 3) Deploy > New deployment > Web app > Execute as: Me
 * 4) กำหนดผู้เข้าถึงตามนโยบายหน่วยงาน แล้วนำ URL ไปใส่ใน index.html
 */

const APP = {
  version: '1.0.0',
  timezone: 'Asia/Bangkok',
  sheets: {
    SETTINGS: ['KEY', 'VALUE', 'DESCRIPTION'],
    CHILDREN: ['CHILD_ID', 'CID', 'HN', 'PREFIX', 'FIRST_NAME', 'LAST_NAME', 'BIRTH_DATE', 'SEX', 'GUARDIAN_NAME', 'PHONE', 'LINE_ID', 'ADDRESS', 'VILLAGE_NO', 'VHV_ID', 'SERVICE_UNIT', 'RESIDENCE_STATUS', 'NOTE', 'CREATED_AT', 'UPDATED_AT'],
    VACCINE_SCHEDULE: ['SCHEDULE_ID', 'VACCINE_CODE', 'VACCINE_NAME', 'DOSE', 'AGE_DAYS', 'EARLY_DAYS', 'LATE_DAYS', 'ACTIVE', 'SORT_ORDER'],
    COVERAGE_TARGETS: ['AGE_GROUP', 'VILLAGE_NO', 'TARGET_COUNT', 'IS_ASSUMPTION', 'NOTE'],
    COVERAGE_DEFINITIONS: ['COVERAGE_ID', 'AGE_GROUP', 'COVERAGE_LABEL', 'SOURCE_SCHEDULE_IDS', 'MATCH_MODE', 'SORT_ORDER', 'ACTIVE'],
    VACCINE_RECORDS: ['RECORD_ID', 'CHILD_ID', 'SCHEDULE_ID', 'VACCINE_CODE', 'DOSE', 'DUE_DATE', 'APPOINTMENT_DATE', 'GIVEN_DATE', 'SERVICE_PLACE', 'LOT_NO', 'DATA_SOURCE', 'RECORDED_BY', 'NOTE', 'CREATED_AT', 'UPDATED_AT'],
    FOLLOW_UP: ['FOLLOWUP_ID', 'CHILD_ID', 'SCHEDULE_ID', 'FOLLOWUP_DATE', 'METHOD', 'RESULT', 'REASON', 'NEW_APPOINTMENT_DATE', 'NEXT_FOLLOWUP_DATE', 'ASSIGNED_TO', 'RECORDED_BY', 'NOTE', 'CREATED_AT'],
    USERS: ['USER_ID', 'EMAIL', 'DISPLAY_NAME', 'ROLE', 'VILLAGE_NO', 'ACTIVE', 'CREATED_AT', 'USERNAME', 'PASSWORD_HASH', 'PASSWORD_SALT', 'APPROVAL_STATUS', 'REQUESTED_AT', 'OTHER_ROLE'],
    IMPORT_LOG: ['IMPORT_ID', 'TIMESTAMP', 'USER_EMAIL', 'IMPORT_TYPE', 'TOTAL_ROWS', 'SUCCESS_ROWS', 'ERROR_ROWS', 'ERROR_DETAIL'],
    AUDIT_LOG: ['LOG_ID', 'TIMESTAMP', 'USER_EMAIL', 'ACTION', 'ENTITY', 'ENTITY_ID', 'DETAIL']
  }
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบวัคซีนเด็ก')
    .addItem('ติดตั้ง/ซ่อมแซมชีต', 'setupSheets')
    .addItem('สร้างนัดที่ยังขาด', 'generateMissingAppointments')
    .addToUi();
}

/** สร้างชีต คอลัมน์ และข้อมูลตั้งต้น โดยไม่ลบข้อมูลเดิม */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(APP.sheets).forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    ensureHeaders_(sheet, APP.sheets[name]);
    formatSheet_(sheet, name);
  });
  seedSettings_();
  seedVaccineSchedule_();
  seedCoverageTargets_();
  seedCoverageDefinitions_();
  seedAdmin_();
  Object.keys(APP.sheets).forEach(function (name) { formatSheet_(ss.getSheetByName(name), name); });
  SpreadsheetApp.flush();
  return { ok: true, message: 'ติดตั้งชีตเรียบร้อย', spreadsheetId: ss.getId() };
}

// alias เผื่อเรียกตามชื่อที่ต้องการ
function SetupSheet() { return setupSheets(); }
function Setup() { return setupSheets(); }

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'health';
    const result = route_(action, e.parameter || {});
    return jsonOutput_({ ok: true, data: result, version: APP.version });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message, stack: String(err.stack || '') });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const result = route_(body.action, body.payload || {});
    return jsonOutput_({ ok: true, data: result, version: APP.version });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message, stack: String(err.stack || '') });
  }
}

function route_(action, payload) {
  payload = payload || {};
  const routes = {
    health: function () { return { status: 'ready', now: new Date() }; },
    login: login_,
    register: register_,
    requestPasswordReset: requestPasswordReset_,
    resetPassword: resetPassword_,
    changePassword: changePassword_,
    saveUser: saveUser_,
    bootstrap: getBootstrapData_,
    saveChild: saveChild_,
    saveVaccination: saveVaccination_,
    saveFollowUp: saveFollowUp_,
    bulkImport: bulkImport_,
    generateAppointments: generateMissingAppointments
  };
  if (!routes[action]) throw new Error('ไม่พบ action: ' + action);
  if (['health', 'login', 'register', 'requestPasswordReset', 'resetPassword'].indexOf(action) < 0) {
    const session = requireAuth_(payload);
    payload._session = session;
    if (action === 'saveUser') requireRole_(session, ['ADMIN']);
    if (action === 'saveChild') requireRole_(session, ['ADMIN', 'PUBLIC_HEALTH', 'DATA_ENTRY', 'EDITOR', 'VHV']);
    if (['saveVaccination', 'generateAppointments'].indexOf(action) >= 0) requireRole_(session, ['ADMIN', 'PUBLIC_HEALTH', 'DATA_ENTRY', 'EDITOR']);
    if (['bulkImport', 'saveFollowUp'].indexOf(action) >= 0) requireRole_(session, ['ADMIN', 'PUBLIC_HEALTH']);
  }
  return routes[action](payload || {});
}

function register_(p) {
  assertRequired_(p, ['username', 'password', 'displayName', 'role']);
  const username = clean_(p.username).toLowerCase(), password = String(p.password || '');
  const role = normalizeRole_(p.role);
  if (role === 'PARENT') throw new Error('ยกเลิกการสมัครด้วยบทบาทบิดา มารดา หรือผู้ปกครองเด็กแล้ว');
  if (role === 'VHV' && !clean_(p.villageNo)) throw new Error('กรุณาระบุหมู่บ้านที่รับผิดชอบสำหรับ อสม.');
  if (!/^[a-z0-9._-]{3,40}$/.test(username)) throw new Error('ชื่อผู้ใช้ต้องเป็น a-z, 0-9, จุด, ขีดกลาง หรือขีดล่าง จำนวน 3-40 ตัว');
  if (password.length < 6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
  const users = rowsAsObjects_('USERS');
  if (users.some(function (u) { return clean_(u.USERNAME).toLowerCase() === username; })) throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้ว');
  const email = clean_(p.email);
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('รูปแบบอีเมลไม่ถูกต้อง');
  const salt = Utilities.getUuid().replace(/-/g, '');
  if (role === 'OTHER' && !clean_(p.otherRole)) throw new Error('กรุณาระบุบทบาทอื่นๆ');
  const row = {
    USER_ID: makeId_('US'), EMAIL: email, DISPLAY_NAME: clean_(p.displayName), ROLE: role,
    VILLAGE_NO: clean_(p.villageNo), ACTIVE: false, CREATED_AT: new Date(), USERNAME: username,
    PASSWORD_HASH: passwordHash_(password, salt), PASSWORD_SALT: salt,
    APPROVAL_STATUS: 'PENDING', REQUESTED_AT: new Date(), OTHER_ROLE: clean_(p.otherRole)
  };
  appendObject_('USERS', row);
  audit_('REGISTER_REQUEST', 'USERS', row.USER_ID, username);
  return { registered: true, message: 'ส่งคำขอสมัครสมาชิกแล้ว กรุณารอผู้ดูแลระบบอนุมัติ' };
}

function login_(p) {
  const username = clean_(p.username).toLowerCase(), password = String(p.password || '');
  if (!username || !password) throw new Error('กรุณาระบุชื่อผู้ใช้และรหัสผ่าน');
  const users = rowsAsObjects_('USERS');
  const user = users.find(function (u) { return clean_(u.USERNAME).toLowerCase() === username; });
  if (!user || passwordHash_(password, user.PASSWORD_SALT) !== String(user.PASSWORD_HASH)) {
    Utilities.sleep(350);
    throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  }
  if (!isTrue_(user.ACTIVE)) throw new Error('บัญชียังไม่ได้รับอนุมัติจากผู้ดูแลระบบ');
  const role = normalizeRole_(user.ROLE);
  if (role === 'PARENT') throw new Error('บทบาทบิดา มารดา หรือผู้ปกครองเด็กถูกยกเลิก กรุณาติดต่อผู้ดูแลระบบ');
  const token = Utilities.getUuid() + Utilities.getUuid();
  const session = { userId: user.USER_ID, username: user.USERNAME, displayName: user.DISPLAY_NAME, role: role, villageNo: clean_(user.VILLAGE_NO), otherRole: clean_(user.OTHER_ROLE) };
  CacheService.getScriptCache().put('session:' + token, JSON.stringify(session), 21600);
  audit_('LOGIN', 'USERS', user.USER_ID, user.USERNAME);
  return { token: token, user: session, expiresIn: 21600 };
}

function requestPasswordReset_(p) {
  const username = clean_(p.username).toLowerCase(), email = clean_(p.email).toLowerCase();
  if (!username || !email) throw new Error('กรุณาระบุชื่อผู้ใช้และอีเมล');
  const generic = { sent: true, message: 'หากข้อมูลถูกต้อง ระบบจะส่งรหัสยืนยันไปยังอีเมลที่ลงทะเบียน' };
  const user = rowsAsObjects_('USERS').find(function (u) {
    return clean_(u.USERNAME).toLowerCase() === username && clean_(u.EMAIL).toLowerCase() === email && isTrue_(u.ACTIVE);
  });
  if (!user) { Utilities.sleep(350); return generic; }
  const cache = CacheService.getScriptCache(), rateKey = 'reset-rate:' + username;
  if (cache.get(rateKey)) throw new Error('เพิ่งส่งรหัสไปแล้ว กรุณารอประมาณ 1 นาทีก่อนขอใหม่');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  const salt = Utilities.getUuid().replace(/-/g, '');
  cache.put('reset:' + username, JSON.stringify({ email: email, hash: passwordHash_(code, salt), salt: salt, attempts: 0 }), 600);
  cache.put(rateKey, '1', 60);
  MailApp.sendEmail({
    to: user.EMAIL,
    subject: 'รหัสยืนยันสำหรับตั้งรหัสผ่านใหม่',
    htmlBody: '<p>เรียน ' + escapeHtml_(user.DISPLAY_NAME || user.USERNAME) + '</p><p>รหัสยืนยันของคุณคือ</p><p style="font-size:28px;font-weight:bold;letter-spacing:5px">' + code + '</p><p>รหัสนี้มีอายุ 10 นาที หากไม่ได้เป็นผู้ขอ โปรดละเว้นอีเมลฉบับนี้</p>',
    name: 'ระบบติดตามความครอบคลุมวัคซีนเด็ก 0-5 ปี'
  });
  audit_('REQUEST_PASSWORD_RESET', 'USERS', user.USER_ID, username);
  return generic;
}

function resetPassword_(p) {
  const username = clean_(p.username).toLowerCase(), email = clean_(p.email).toLowerCase();
  const code = clean_(p.code), newPassword = String(p.newPassword || '');
  if (!username || !email || !code) throw new Error('กรุณากรอกข้อมูลและรหัสยืนยันให้ครบ');
  if (newPassword.length < 6) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร');
  const cache = CacheService.getScriptCache(), key = 'reset:' + username, raw = cache.get(key);
  if (!raw) throw new Error('รหัสยืนยันหมดอายุ กรุณาขอรหัสใหม่');
  const reset = JSON.parse(raw);
  if (reset.email !== email || passwordHash_(code, reset.salt) !== reset.hash) {
    reset.attempts = Number(reset.attempts || 0) + 1;
    if (reset.attempts >= 5) cache.remove(key); else cache.put(key, JSON.stringify(reset), 600);
    throw new Error('รหัสยืนยันไม่ถูกต้อง');
  }
  const found = rowsAsObjects_('USERS').find(function (u) { return clean_(u.USERNAME).toLowerCase() === username && clean_(u.EMAIL).toLowerCase() === email && isTrue_(u.ACTIVE); });
  if (!found) throw new Error('ไม่สามารถตั้งรหัสผ่านใหม่ได้');
  const salt = Utilities.getUuid().replace(/-/g, '');
  found.PASSWORD_SALT = salt;
  found.PASSWORD_HASH = passwordHash_(newPassword, salt);
  upsertObject_('USERS', 'USER_ID', found);
  cache.remove(key);
  audit_('RESET_PASSWORD', 'USERS', found.USER_ID, username);
  return { changed: true };
}

function requireAuth_(p) {
  const token = clean_(p && p.token);
  if (!token) throw new Error('UNAUTHORIZED: กรุณาเข้าสู่ระบบ');
  const cache = CacheService.getScriptCache(), key = 'session:' + token, raw = cache.get(key);
  if (!raw) throw new Error('UNAUTHORIZED: session หมดอายุ กรุณาเข้าสู่ระบบใหม่');
  cache.put(key, raw, 21600);
  return JSON.parse(raw);
}

function requireRole_(session, allowedRoles) {
  const role = String(session && session.role || '').toUpperCase();
  if (allowedRoles.indexOf(role) < 0) throw new Error('FORBIDDEN: บัญชีนี้ไม่มีสิทธิ์แก้ไขข้อมูล');
}

function normalizeRole_(role) {
  role = String(role || '').toUpperCase();
  if (role === 'EDITOR') return 'DATA_ENTRY'; // รองรับบัญชีรุ่นเดิม
  const allowed = ['PUBLIC_HEALTH', 'VHV', 'DATA_ENTRY', 'PARENT', 'ADMIN', 'OTHER'];
  if (allowed.indexOf(role) < 0) throw new Error('บทบาทผู้ใช้งานไม่ถูกต้อง');
  return role;
}

function assertVillageAccess_(session, childId) {
  const child = findRow_('CHILDREN', 'CHILD_ID', childId);
  if (!child || !clean_(session.villageNo) || clean_(child.object.VILLAGE_NO) !== clean_(session.villageNo)) {
    throw new Error('FORBIDDEN: ไม่สามารถเข้าถึงข้อมูลเด็กนอกพื้นที่รับผิดชอบ');
  }
}

function changePassword_(p) {
  const session = requireAuth_(p);
  const oldPassword = String(p.oldPassword || ''), newPassword = String(p.newPassword || '');
  if (newPassword.length < 6) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร');
  const found = findRow_('USERS', 'USER_ID', session.userId);
  if (!found) throw new Error('ไม่พบบัญชีผู้ใช้');
  const user = found.object;
  if (passwordHash_(oldPassword, user.PASSWORD_SALT) !== String(user.PASSWORD_HASH)) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  const salt = Utilities.getUuid().replace(/-/g, '');
  user.PASSWORD_SALT = salt;
  user.PASSWORD_HASH = passwordHash_(newPassword, salt);
  upsertObject_('USERS', 'USER_ID', user);
  audit_('CHANGE_PASSWORD', 'USERS', user.USER_ID, user.USERNAME);
  return { changed: true };
}

function getBootstrapData_(p) {
  assertInstalled_();
  const session = requireAuth_(p);
  const role = normalizeRole_(session.role), allChildren = rowsAsObjects_('CHILDREN');
  let children = allChildren;
  if (role === 'VHV') children = allChildren.filter(function (c) { return clean_(c.VILLAGE_NO) === clean_(session.villageNo); });
  if (['PARENT', 'OTHER'].indexOf(role) >= 0) children = [];
  const schedule = rowsAsObjects_('VACCINE_SCHEDULE').filter(function (r) { return isTrue_(r.ACTIVE); });
  const childIds = {};
  children.forEach(function (c) { childIds[String(c.CHILD_ID)] = true; });
  let records = rowsAsObjects_('VACCINE_RECORDS');
  let followups = rowsAsObjects_('FOLLOW_UP');
  if (role === 'VHV') {
    records = records.filter(function (r) { return childIds[String(r.CHILD_ID)]; });
    followups = followups.filter(function (r) { return childIds[String(r.CHILD_ID)]; });
  }
  if (['PARENT', 'OTHER'].indexOf(role) >= 0) { records = []; followups = []; }
  const coverageDefinitions = rowsAsObjects_('COVERAGE_DEFINITIONS').filter(function (r) { return isTrue_(r.ACTIVE); });
  let coverageTargets = rowsAsObjects_('COVERAGE_TARGETS');
  if (role === 'VHV') coverageTargets = coverageTargets.filter(function (t) { return clean_(t.VILLAGE_NO) === clean_(session.villageNo); });
  const canTrack = ['ADMIN', 'PUBLIC_HEALTH'].indexOf(role) >= 0;
  const response = {
    children: children,
    schedule: schedule,
    records: records,
    followups: followups,
    tracker: canTrack ? buildTracker_(children, schedule, records, followups) : [],
    coverage: buildCoverage_(role === 'VHV' ? children : allChildren, role === 'VHV' ? records : rowsAsObjects_('VACCINE_RECORDS'), coverageDefinitions, coverageTargets),
    coverageDefinitions: coverageDefinitions,
    coverageTargets: coverageTargets,
    settings: settingsObject_(),
    serverTime: formatDate_(new Date(), true), currentUser: session
  };
  if (String(session.role).toUpperCase() === 'ADMIN') {
    response.users = rowsAsObjects_('USERS').map(function (u) {
      return { USER_ID: u.USER_ID, USERNAME: u.USERNAME, EMAIL: u.EMAIL, DISPLAY_NAME: u.DISPLAY_NAME, ROLE: normalizeRole_(u.ROLE), OTHER_ROLE: u.OTHER_ROLE, VILLAGE_NO: u.VILLAGE_NO, ACTIVE: isTrue_(u.ACTIVE), APPROVAL_STATUS: u.APPROVAL_STATUS || (isTrue_(u.ACTIVE) ? 'APPROVED' : 'PENDING'), CREATED_AT: u.CREATED_AT, REQUESTED_AT: u.REQUESTED_AT };
    });
  }
  return response;
}

function saveUser_(p) {
  assertRequired_(p, ['USERNAME', 'DISPLAY_NAME', 'ROLE']);
  const username = clean_(p.USERNAME).toLowerCase(), role = String(p.ROLE).toUpperCase();
  if (!/^[a-z0-9._-]{3,40}$/.test(username)) throw new Error('ชื่อผู้ใช้ต้องเป็น a-z, 0-9, จุด, ขีดกลาง หรือขีดล่าง จำนวน 3-40 ตัว');
  normalizeRole_(role);
  if (role === 'PARENT') throw new Error('ไม่สามารถกำหนดบทบาทบิดา มารดา หรือผู้ปกครองเด็กได้');
  if (role === 'VHV' && !clean_(p.VILLAGE_NO)) throw new Error('กรุณาระบุหมู่บ้านที่รับผิดชอบสำหรับ อสม.');
  if (role === 'OTHER' && !clean_(p.OTHER_ROLE)) throw new Error('กรุณาระบุบทบาทอื่นๆ');
  const users = rowsAsObjects_('USERS');
  const duplicate = users.find(function (u) { return clean_(u.USERNAME).toLowerCase() === username && String(u.USER_ID) !== String(p.USER_ID || ''); });
  if (duplicate) throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้ว');
  const existing = p.USER_ID ? findRow_('USERS', 'USER_ID', p.USER_ID) : null;
  const old = existing ? existing.object : {};
  const password = String(p.PASSWORD || '');
  if (!existing && password.length < 6) throw new Error('บัญชีใหม่ต้องมีรหัสผ่านอย่างน้อย 6 ตัวอักษร');
  if (password && password.length < 6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
  let salt = old.PASSWORD_SALT || '', hash = old.PASSWORD_HASH || '';
  if (password) { salt = Utilities.getUuid().replace(/-/g, ''); hash = passwordHash_(password, salt); }
  const row = {
    USER_ID: old.USER_ID || makeId_('US'), EMAIL: clean_(p.EMAIL) || old.EMAIL || '', DISPLAY_NAME: clean_(p.DISPLAY_NAME), ROLE: role,
    VILLAGE_NO: clean_(p.VILLAGE_NO), ACTIVE: p.ACTIVE === false || String(p.ACTIVE).toLowerCase() === 'false' ? false : true,
    CREATED_AT: old.CREATED_AT || new Date(), USERNAME: username, PASSWORD_HASH: hash, PASSWORD_SALT: salt,
    APPROVAL_STATUS: p.ACTIVE === false || String(p.ACTIVE).toLowerCase() === 'false' ? 'PENDING' : 'APPROVED', REQUESTED_AT: old.REQUESTED_AT || new Date(), OTHER_ROLE: clean_(p.OTHER_ROLE)
  };
  upsertObject_('USERS', 'USER_ID', row);
  audit_('SAVE_USER', 'USERS', row.USER_ID, username + '/' + role);
  return { userId: row.USER_ID };
}

function saveChild_(p) {
  assertRequired_(p, ['FIRST_NAME', 'LAST_NAME', 'BIRTH_DATE']);
  const session = p._session || {}, role = normalizeRole_(session.role);
  const now = new Date();
  const id = p.CHILD_ID || makeId_('CH');
  const existing = findRow_('CHILDREN', 'CHILD_ID', id);
  if (role === 'VHV') {
    if (!clean_(session.villageNo)) throw new Error('บัญชี อสม. ยังไม่ได้ระบุหมู่บ้านที่รับผิดชอบ กรุณาติดต่อผู้ดูแลระบบ');
    if (existing) assertVillageAccess_(session, id);
    p.VILLAGE_NO = clean_(session.villageNo);
    if (!clean_(p.VHV_ID)) p.VHV_ID = clean_(session.displayName) || clean_(session.username);
  }
  const row = {
    CHILD_ID: id, CID: clean_(p.CID), HN: clean_(p.HN), PREFIX: clean_(p.PREFIX),
    FIRST_NAME: clean_(p.FIRST_NAME), LAST_NAME: clean_(p.LAST_NAME), BIRTH_DATE: parseDate_(p.BIRTH_DATE),
    SEX: clean_(p.SEX), GUARDIAN_NAME: clean_(p.GUARDIAN_NAME), PHONE: clean_(p.PHONE),
    LINE_ID: clean_(p.LINE_ID), ADDRESS: clean_(p.ADDRESS), VILLAGE_NO: clean_(p.VILLAGE_NO),
    VHV_ID: clean_(p.VHV_ID), SERVICE_UNIT: clean_(p.SERVICE_UNIT),
    RESIDENCE_STATUS: clean_(p.RESIDENCE_STATUS) || 'อยู่ในพื้นที่', NOTE: clean_(p.NOTE),
    CREATED_AT: existing ? existing.object.CREATED_AT : now, UPDATED_AT: now
  };
  upsertObject_('CHILDREN', 'CHILD_ID', row);
  createAppointmentsForChild_(row);
  audit_('SAVE_CHILD', 'CHILDREN', id, row.FIRST_NAME + ' ' + row.LAST_NAME);
  return { childId: id };
}

function saveVaccination_(p) {
  assertRequired_(p, ['CHILD_ID', 'SCHEDULE_ID', 'GIVEN_DATE']);
  const schedule = findRow_('VACCINE_SCHEDULE', 'SCHEDULE_ID', p.SCHEDULE_ID);
  if (!schedule) throw new Error('ไม่พบรายการวัคซีน');
  const existing = findAppointment_(p.CHILD_ID, p.SCHEDULE_ID);
  const now = new Date();
  const s = schedule.object;
  const row = {
    RECORD_ID: existing ? existing.object.RECORD_ID : makeId_('VR'),
    CHILD_ID: p.CHILD_ID, SCHEDULE_ID: p.SCHEDULE_ID, VACCINE_CODE: s.VACCINE_CODE, DOSE: s.DOSE,
    DUE_DATE: existing ? existing.object.DUE_DATE : '', APPOINTMENT_DATE: parseDateOrBlank_(p.APPOINTMENT_DATE) || (existing ? existing.object.APPOINTMENT_DATE : ''),
    GIVEN_DATE: parseDate_(p.GIVEN_DATE), SERVICE_PLACE: clean_(p.SERVICE_PLACE), LOT_NO: clean_(p.LOT_NO),
    DATA_SOURCE: clean_(p.DATA_SOURCE) || 'หน่วยบริการนี้', RECORDED_BY: currentUser_(), NOTE: clean_(p.NOTE),
    CREATED_AT: existing ? existing.object.CREATED_AT : now, UPDATED_AT: now
  };
  upsertObject_('VACCINE_RECORDS', 'RECORD_ID', row);
  audit_('SAVE_VACCINATION', 'VACCINE_RECORDS', row.RECORD_ID, p.CHILD_ID + '/' + p.SCHEDULE_ID);
  return { recordId: row.RECORD_ID };
}

function saveFollowUp_(p) {
  assertRequired_(p, ['CHILD_ID', 'FOLLOWUP_DATE', 'METHOD', 'RESULT']);
  const row = {
    FOLLOWUP_ID: makeId_('FU'), CHILD_ID: p.CHILD_ID, SCHEDULE_ID: clean_(p.SCHEDULE_ID),
    FOLLOWUP_DATE: parseDate_(p.FOLLOWUP_DATE), METHOD: clean_(p.METHOD), RESULT: clean_(p.RESULT),
    REASON: clean_(p.REASON), NEW_APPOINTMENT_DATE: parseDateOrBlank_(p.NEW_APPOINTMENT_DATE),
    NEXT_FOLLOWUP_DATE: parseDateOrBlank_(p.NEXT_FOLLOWUP_DATE), ASSIGNED_TO: clean_(p.ASSIGNED_TO),
    RECORDED_BY: currentUser_(), NOTE: clean_(p.NOTE), CREATED_AT: new Date()
  };
  appendObject_('FOLLOW_UP', row);
  if (row.NEW_APPOINTMENT_DATE && row.SCHEDULE_ID) updateAppointment_(row.CHILD_ID, row.SCHEDULE_ID, row.NEW_APPOINTMENT_DATE);
  audit_('SAVE_FOLLOWUP', 'FOLLOW_UP', row.FOLLOWUP_ID, row.RESULT);
  return { followupId: row.FOLLOWUP_ID };
}

/** นำเข้าข้อมูลแบบชุดจาก CSV/ข้อมูลที่คัดลอกจาก Excel */
function bulkImport_(p) {
  const type = clean_(p.importType);
  const rows = Array.isArray(p.rows) ? p.rows : [];
  if (['children', 'vaccinations', 'targets'].indexOf(type) < 0) throw new Error('ประเภทการนำเข้าไม่ถูกต้อง');
  if (!rows.length) throw new Error('ไม่พบข้อมูลสำหรับนำเข้า');
  if (rows.length > 1000) throw new Error('นำเข้าได้ครั้งละไม่เกิน 1,000 แถว');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const errors = [], successIds = [];
  try {
    rows.forEach(function (raw, index) {
      try {
        const row = normalizeImportRow_(raw);
        if (type === 'children') successIds.push(importChildRow_(row));
        if (type === 'vaccinations') successIds.push(importVaccinationRow_(row));
        if (type === 'targets') successIds.push(importTargetRow_(row));
      } catch (err) {
        errors.push({ row: index + 2, message: err.message });
      }
    });
  } finally {
    lock.releaseLock();
  }
  const result = { total: rows.length, success: successIds.length, errors: errors };
  appendObject_('IMPORT_LOG', {
    IMPORT_ID: makeId_('IMP'), TIMESTAMP: new Date(), USER_EMAIL: currentUser_(), IMPORT_TYPE: type,
    TOTAL_ROWS: rows.length, SUCCESS_ROWS: successIds.length, ERROR_ROWS: errors.length,
    ERROR_DETAIL: errors.slice(0, 50).map(function (x) { return 'แถว ' + x.row + ': ' + x.message; }).join(' | ')
  });
  audit_('BULK_IMPORT', 'IMPORT_LOG', type, JSON.stringify({ total: rows.length, success: successIds.length, errors: errors.length }));
  return result;
}

function importChildRow_(p) {
  assertRequired_(p, ['FIRST_NAME', 'LAST_NAME', 'BIRTH_DATE']);
  let id = clean_(p.CHILD_ID);
  if (!id && p.CID) { const foundCid = findRow_('CHILDREN', 'CID', clean_(p.CID)); if (foundCid) id = foundCid.object.CHILD_ID; }
  if (!id && p.HN) { const foundHn = findRow_('CHILDREN', 'HN', clean_(p.HN)); if (foundHn) id = foundHn.object.CHILD_ID; }
  p.CHILD_ID = id || makeId_('CH');
  saveChild_(p);
  return p.CHILD_ID;
}

function importVaccinationRow_(p) {
  assertRequired_(p, ['CHILD_ID', 'SCHEDULE_ID', 'GIVEN_DATE']);
  if (!findRow_('CHILDREN', 'CHILD_ID', p.CHILD_ID)) throw new Error('ไม่พบ CHILD_ID: ' + p.CHILD_ID);
  if (!findRow_('VACCINE_SCHEDULE', 'SCHEDULE_ID', p.SCHEDULE_ID)) throw new Error('ไม่พบ SCHEDULE_ID: ' + p.SCHEDULE_ID);
  return saveVaccination_(p).recordId;
}

function importTargetRow_(p) {
  assertRequired_(p, ['AGE_GROUP', 'VILLAGE_NO', 'TARGET_COUNT']);
  if (['1y','2y','3y','5y'].indexOf(String(p.AGE_GROUP)) < 0) throw new Error('AGE_GROUP ต้องเป็น 1y, 2y, 3y หรือ 5y');
  const count = Number(p.TARGET_COUNT);
  if (!isFinite(count) || count < 0 || Math.floor(count) !== count) throw new Error('TARGET_COUNT ต้องเป็นจำนวนเต็มตั้งแต่ 0 ขึ้นไป');
  const sh = getSheet_('COVERAGE_TARGETS'), values = sh.getDataRange().getValues(), h = values[0];
  const ai = h.indexOf('AGE_GROUP'), vi = h.indexOf('VILLAGE_NO');
  let targetRow = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][ai]) === String(p.AGE_GROUP) && String(values[i][vi]) === String(p.VILLAGE_NO)) { targetRow = i + 1; break; }
  }
  const obj = { AGE_GROUP: String(p.AGE_GROUP), VILLAGE_NO: String(p.VILLAGE_NO), TARGET_COUNT: count, IS_ASSUMPTION: true, NOTE: clean_(p.NOTE) || 'นำเข้าจากไฟล์โดยเจ้าหน้าที่' };
  const row = APP.sheets.COVERAGE_TARGETS.map(function (x) { return obj[x] === undefined ? '' : obj[x]; });
  if (targetRow) sh.getRange(targetRow, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  return obj.AGE_GROUP + '-' + obj.VILLAGE_NO;
}

function normalizeImportRow_(row) {
  const aliases = {
    'รหัสเด็ก':'CHILD_ID','เลขบัตรประชาชน':'CID','เลขบัตรประชาชน 13 หลัก':'CID','HN':'HN','คำนำหน้า':'PREFIX',
    'ชื่อ':'FIRST_NAME','นามสกุล':'LAST_NAME','วันเกิด':'BIRTH_DATE','เพศ':'SEX','ชื่อผู้ปกครอง':'GUARDIAN_NAME',
    'โทรศัพท์':'PHONE','LINE':'LINE_ID','ที่อยู่':'ADDRESS','หมู่':'VILLAGE_NO','หมู่ที่':'VILLAGE_NO','อสม.':'VHV_ID',
    'หน่วยบริการ':'SERVICE_UNIT','สถานะ':'RESIDENCE_STATUS','หมายเหตุ':'NOTE','รหัสนัดวัคซีน':'SCHEDULE_ID',
    'วันที่ได้รับวัคซีน':'GIVEN_DATE','สถานที่ฉีด':'SERVICE_PLACE','เลข Lot':'LOT_NO','แหล่งข้อมูล':'DATA_SOURCE',
    'กลุ่มอายุ':'AGE_GROUP','จำนวนกลุ่มเป้าหมาย':'TARGET_COUNT'
  };
  const out = {};
  Object.keys(row || {}).forEach(function (key) {
    const cleanKey = String(key).replace(/^\uFEFF/, '').trim();
    out[aliases[cleanKey] || cleanKey.toUpperCase()] = row[key];
  });
  return out;
}

function generateMissingAppointments() {
  assertInstalled_();
  const children = rowsAsObjects_('CHILDREN');
  let created = 0;
  children.forEach(function (child) { created += createAppointmentsForChild_(child); });
  return { created: created };
}

function createAppointmentsForChild_(child) {
  const schedule = rowsAsObjects_('VACCINE_SCHEDULE').filter(function (r) { return isTrue_(r.ACTIVE); });
  const records = rowsAsObjects_('VACCINE_RECORDS');
  const birth = parseDate_(child.BIRTH_DATE);
  let created = 0;
  schedule.forEach(function (s) {
    const exists = records.some(function (r) { return String(r.CHILD_ID) === String(child.CHILD_ID) && String(r.SCHEDULE_ID) === String(s.SCHEDULE_ID); });
    if (exists) return;
    const due = addDays_(birth, Number(s.AGE_DAYS || 0));
    appendObject_('VACCINE_RECORDS', {
      RECORD_ID: makeId_('VR'), CHILD_ID: child.CHILD_ID, SCHEDULE_ID: s.SCHEDULE_ID,
      VACCINE_CODE: s.VACCINE_CODE, DOSE: s.DOSE, DUE_DATE: due, APPOINTMENT_DATE: due,
      GIVEN_DATE: '', SERVICE_PLACE: '', LOT_NO: '', DATA_SOURCE: '', RECORDED_BY: '', NOTE: '',
      CREATED_AT: new Date(), UPDATED_AT: new Date()
    });
    created++;
  });
  return created;
}

function buildTracker_(children, schedule, records, followups) {
  const today = stripTime_(new Date());
  const scheduleMap = {};
  schedule.forEach(function (s) { scheduleMap[s.SCHEDULE_ID] = s; });
  const childMap = {};
  children.forEach(function (c) { childMap[c.CHILD_ID] = c; });
  return records.filter(function (r) { return !r.GIVEN_DATE && childMap[r.CHILD_ID]; }).map(function (r) {
    const c = childMap[r.CHILD_ID];
    const s = scheduleMap[r.SCHEDULE_ID] || {};
    const due = stripTime_(parseDate_(r.APPOINTMENT_DATE || r.DUE_DATE));
    const diff = Math.floor((today.getTime() - due.getTime()) / 86400000);
    let status = 'UPCOMING', color = 'blue', label = 'มีนัดล่วงหน้า';
    if (diff >= 90) { status = 'CHRONIC'; color = 'black'; label = 'ขาดนัดเรื้อรัง'; }
    else if (diff >= 15) { status = 'URGENT'; color = 'red'; label = 'ต้องเร่งติดตาม'; }
    else if (diff >= 1) { status = 'MISSED'; color = 'orange'; label = 'ขาดนัดระยะแรก'; }
    else if (diff >= -7) { status = 'DUE_SOON'; color = 'yellow'; label = diff === 0 ? 'นัดวันนี้' : 'ใกล้ถึงกำหนด'; }
    const latest = followups.filter(function (f) { return f.CHILD_ID === r.CHILD_ID && (!f.SCHEDULE_ID || f.SCHEDULE_ID === r.SCHEDULE_ID); })
      .sort(function (a, b) { return new Date(b.FOLLOWUP_DATE) - new Date(a.FOLLOWUP_DATE); })[0];
    return {
      recordId: r.RECORD_ID, childId: c.CHILD_ID, childName: [c.PREFIX, c.FIRST_NAME, c.LAST_NAME].filter(Boolean).join(' '),
      birthDate: formatDate_(c.BIRTH_DATE), guardian: c.GUARDIAN_NAME, phone: c.PHONE, village: c.VILLAGE_NO,
      vhv: c.VHV_ID, scheduleId: r.SCHEDULE_ID, vaccine: s.VACCINE_NAME || r.VACCINE_CODE, dose: r.DOSE,
      dueDate: formatDate_(r.DUE_DATE), appointmentDate: formatDate_(r.APPOINTMENT_DATE || r.DUE_DATE),
      overdueDays: Math.max(0, diff), status: status, color: color, label: label,
      latestFollowup: latest ? { date: formatDate_(latest.FOLLOWUP_DATE), result: latest.RESULT } : null
    };
  }).sort(function (a, b) { return new Date(a.appointmentDate) - new Date(b.appointmentDate); });
}

/** สรุปความครอบคลุมให้ตรงโครงสร้างไฟล์อ้างอิง: เป้าหมาย / ได้รับวัคซีน / ร้อยละ */
function buildCoverage_(children, records, definitions, targets) {
  const childMap = {};
  children.forEach(function (c) { childMap[String(c.CHILD_ID)] = c; });
  return definitions.sort(function (a, b) { return Number(a.SORT_ORDER) - Number(b.SORT_ORDER); }).map(function (d) {
    const sourceIds = String(d.SOURCE_SCHEDULE_IDS || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
    const targetByVillage = {};
    targets.filter(function (t) { return String(t.AGE_GROUP) === String(d.AGE_GROUP); }).forEach(function (t) {
      targetByVillage[String(t.VILLAGE_NO)] = Number(t.TARGET_COUNT || 0);
    });
    const receivedByChild = {};
    records.filter(function (r) { return r.GIVEN_DATE && childMap[String(r.CHILD_ID)]; }).forEach(function (r) {
      const id = String(r.CHILD_ID);
      if (!receivedByChild[id]) receivedByChild[id] = {};
      receivedByChild[id][String(r.SCHEDULE_ID)] = true;
    });
    const receivedByVillage = {};
    Object.keys(receivedByChild).forEach(function (childId) {
      const hits = sourceIds.map(function (sid) { return !!receivedByChild[childId][sid]; });
      const matched = String(d.MATCH_MODE).toUpperCase() === 'ALL' ? hits.length > 0 && hits.every(Boolean) : hits.some(Boolean);
      if (!matched) return;
      const village = String(childMap[childId].VILLAGE_NO || 'ไม่ระบุ');
      receivedByVillage[village] = (receivedByVillage[village] || 0) + 1;
    });
    const villages = Object.keys(targetByVillage).sort(function (a, b) { return Number(a) - Number(b); });
    const detail = villages.map(function (v) {
      const target = targetByVillage[v] || 0, received = receivedByVillage[v] || 0;
      return { village: v, target: target, received: received, percent: target ? received * 100 / target : null };
    });
    const target = detail.reduce(function (s, x) { return s + x.target; }, 0);
    const received = detail.reduce(function (s, x) { return s + x.received; }, 0);
    return {
      coverageId: d.COVERAGE_ID, ageGroup: d.AGE_GROUP, label: d.COVERAGE_LABEL,
      target: target, received: received, percent: target ? received * 100 / target : null,
      assumedTarget: true, detail: detail
    };
  });
}

function seedSettings_() {
  const defaults = [
    ['APP_NAME', 'ระบบติดตามความครอบคลุมวัคซีนเด็ก 0-5 ปี', 'ชื่อระบบ'],
    ['APPOINTMENT_DOC_URL', 'https://docs.google.com/document/d/1j_i8wqZnIxw6OzARumGkRthTVAKbeoBhcFBZn0GysxU/edit', 'แม่แบบใบนัด Google Docs'],
    ['SERVICE_UNIT', 'รพ.สต.ตัวอย่าง', 'ชื่อหน่วยบริการ'],
    ['WARN_DAYS', '7', 'แจ้งเตือนก่อนนัดกี่วัน'],
    ['URGENT_DAYS', '15', 'จำนวนวันขาดนัดระดับเร่งด่วน'],
    ['CHRONIC_DAYS', '90', 'จำนวนวันขาดนัดเรื้อรัง']
  ];
  const existing = rowsAsObjects_('SETTINGS').map(function (r) { return r.KEY; });
  defaults.forEach(function (r) { if (existing.indexOf(r[0]) < 0) getSheet_('SETTINGS').appendRow(r); });
  const appName = findRow_('SETTINGS', 'KEY', 'APP_NAME');
  if (appName && String(appName.object.VALUE || '').indexOf('ระบบติดตามวัคซีนเด็ก') === 0) {
    appName.object.VALUE = 'ระบบติดตามความครอบคลุมวัคซีนเด็ก 0-5 ปี';
    upsertObject_('SETTINGS', 'KEY', appName.object);
  }
}

function seedVaccineSchedule_() {
  // ตารางตั้งต้นสำหรับสาธิต ต้องให้ผู้รับผิดชอบ EPI ตรวจสอบกับแนวทางล่าสุดก่อนใช้งานจริง
  const data = [
    ['BCG-1','BCG','BCG',1,0,0,30,true,10], ['HB1','HB','ตับอักเสบบี',1,0,0,1,true,20],
    ['DTPHBHIB1','DTP-HB-Hib','คอตีบ-บาดทะยัก-ไอกรน-ตับอักเสบบี-ฮิบ',1,60,0,30,true,30],
    ['OPV1','OPV','โปลิโอชนิดรับประทาน',1,60,0,30,true,40], ['IPV0','IPV','โปลิโอชนิดฉีด 2 เดือน',1,60,0,30,true,45], ['ROTA1','Rota','โรต้า',1,60,0,14,true,50], ['PCV1','PCV','นิวโมคอคคัส 2 เดือน',1,60,0,30,true,55],
    ['DTPHBHIB2','DTP-HB-Hib','คอตีบ-บาดทะยัก-ไอกรน-ตับอักเสบบี-ฮิบ',2,120,0,30,true,60],
    ['OPV2','OPV','โปลิโอชนิดรับประทาน',2,120,0,30,true,70], ['IPV1','IPV','โปลิโอชนิดฉีด 4 เดือน',2,120,0,30,true,80], ['PCV2','PCV','นิวโมคอคคัส 4 เดือน',2,120,0,30,true,85],
    ['ROTA2','Rota','โรต้า',2,120,0,14,true,90], ['DTPHBHIB3','DTP-HB-Hib','คอตีบ-บาดทะยัก-ไอกรน-ตับอักเสบบี-ฮิบ',3,180,0,30,true,100],
    ['OPV3','OPV','โปลิโอชนิดรับประทาน',3,180,0,30,true,110], ['ROTA3','Rota','โรต้า',3,180,0,14,true,120], ['PCV3','PCV','นิวโมคอคคัส 6 เดือน (นอกแผน)',3,180,0,30,true,125],
    ['MMR1','MMR','หัด-คางทูม-หัดเยอรมัน',1,270,0,60,true,130], ['LAJE1','LAJE','ไข้สมองอักเสบเจอี',1,365,0,90,true,140], ['PCV4','PCV','นิวโมคอคคัส 1 ปี',4,365,0,90,true,145],
    ['DTP4','DTP','คอตีบ-บาดทะยัก-ไอกรน',4,548,0,60,true,150], ['OPV4','OPV','โปลิโอชนิดรับประทาน',4,548,0,60,true,160],
    ['MMR2','MMR','หัด-คางทูม-หัดเยอรมัน',2,548,0,90,true,170], ['LAJE2','LAJE','ไข้สมองอักเสบเจอี',2,913,0,90,true,180],
    ['DTP5','DTP','คอตีบ-บาดทะยัก-ไอกรน',5,1461,0,180,true,190], ['OPV5','OPV','โปลิโอชนิดรับประทาน',5,1461,0,180,true,200]
  ];
  const existing = rowsAsObjects_('VACCINE_SCHEDULE').map(function (r) { return String(r.SCHEDULE_ID); });
  data.forEach(function (row) { if (existing.indexOf(String(row[0])) < 0) getSheet_('VACCINE_SCHEDULE').appendRow(row); });
}

function seedCoverageTargets_() {
  if (getSheet_('COVERAGE_TARGETS').getLastRow() > 1) return;
  // จำนวนกลุ่มเป้าหมายสมมุติจากไฟล์ Vaccine coverage_completed.xlsx
  const groups = {
    '1y': [424,271,131,252,73,103,341,101],
    '2y': [554,368,186,316,87,122,385,136],
    '3y': [616,414,203,355,113,132,399,119],
    '5y': [655,538,210,359,175,138,442,148]
  };
  const rows = [];
  Object.keys(groups).forEach(function (group) {
    groups[group].forEach(function (count, i) { rows.push([group, String(i + 1), count, true, 'จำนวนสมมุติตามไฟล์อ้างอิง']); });
  });
  getSheet_('COVERAGE_TARGETS').getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedCoverageDefinitions_() {
  if (getSheet_('COVERAGE_DEFINITIONS').getLastRow() > 1) return;
  const rows = [
    ['BCG-BIRTH','1y','BCG แรกเกิด','BCG-1','ANY',10,true],
    ['HBV-BIRTH','1y','HBV แรกเกิด','HB1','ANY',20,true],
    ['DTP-2M','1y','DTP 2 เดือน','DTPHBHIB1','ANY',30,true],
    ['DTP-4M','1y','DTP 4 เดือน','DTPHBHIB2','ANY',40,true],
    ['DTP-6M','1y','DTP 6 เดือน','DTPHBHIB3','ANY',50,true],
    ['HBV-2M','1y','HBV 2 เดือน','DTPHBHIB1','ANY',60,true],
    ['HBV-4M','1y','HBV 4 เดือน','DTPHBHIB2','ANY',70,true],
    ['HBV-6M','1y','HBV 6 เดือน','DTPHBHIB3','ANY',80,true],
    ['HIB-2M','1y','Hib 2 เดือน','DTPHBHIB1','ANY',90,true],
    ['HIB-4M','1y','Hib 4 เดือน','DTPHBHIB2','ANY',100,true],
    ['HIB-6M','1y','Hib 6 เดือน','DTPHBHIB3','ANY',110,true],
    ['IPV-2M','1y','IPV 2 เดือน','IPV0','ANY',120,true],
    ['IPV-4M','1y','IPV 4 เดือน','IPV1','ANY',130,true],
    ['OPV-6M','1y','OPV 6 เดือน','OPV3','ANY',140,true],
    ['MMR-9M','1y','MMR 9 เดือน','MMR1','ANY',150,true],
    ['ROTA-2M','1y','Rota 2 เดือน','ROTA1','ANY',160,true],
    ['ROTA-COMPLETE','1y','Rota (ครบถ้วนตามเกณฑ์)','ROTA1,ROTA2,ROTA3','ALL',170,true],
    ['PCV-2M','1y','PCV 2 เดือน','PCV1','ANY',180,true],
    ['PCV-4M','1y','PCV 4 เดือน','PCV2','ANY',190,true],
    ['PCV-6M','1y','PCV 6 เดือน (นอกแผน)','PCV3','ANY',200,true],
    ['PCV-1Y','2y','PCV 1 ปี','PCV4','ANY',210,true],
    ['LAJE-1Y','2y','LAJE/JE 1 ปี','LAJE1','ANY',220,true],
    ['DTP-18M','2y','DTP 1 ปี 6 เดือน','DTP4','ANY',230,true],
    ['OPV-18M','2y','OPV 1 ปี 6 เดือน','OPV4','ANY',240,true],
    ['MMR1-CATCHUP','2y','MMR1 เก็บตก','MMR1-CATCHUP','ANY',250,true],
    ['MMR-18M','2y','MMR 1 ปี 6 เดือน','MMR2','ANY',260,true],
    ['LAJE-30M','3y','LAJE/JE 2 ปี 6 เดือน','LAJE2','ANY',270,true],
    ['DTP-4Y','5y','DTP 4 ปี','DTP5','ANY',280,true],
    ['OPV-4Y','5y','OPV 4 ปี','OPV5','ANY',290,true]
  ];
  getSheet_('COVERAGE_DEFINITIONS').getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedAdmin_() {
  const users = rowsAsObjects_('USERS');
  const existing = users.find(function (u) { return clean_(u.USERNAME).toLowerCase() === 'admin'; });
  if (existing && existing.PASSWORD_HASH && existing.PASSWORD_SALT) return;
  const previousAdmin = existing || users.find(function (u) { return String(u.ROLE).toUpperCase() === 'ADMIN'; });
  const salt = Utilities.getUuid().replace(/-/g, '');
  const row = {
    USER_ID: previousAdmin ? previousAdmin.USER_ID : makeId_('US'),
    EMAIL: previousAdmin ? previousAdmin.EMAIL : currentUser_(), DISPLAY_NAME: previousAdmin ? previousAdmin.DISPLAY_NAME : 'ผู้ดูแลระบบ',
    ROLE: 'ADMIN', VILLAGE_NO: previousAdmin ? previousAdmin.VILLAGE_NO : '', ACTIVE: true,
    CREATED_AT: previousAdmin ? previousAdmin.CREATED_AT : new Date(), USERNAME: 'admin',
    PASSWORD_HASH: passwordHash_('admin', salt), PASSWORD_SALT: salt, APPROVAL_STATUS: 'APPROVED', REQUESTED_AT: new Date(), OTHER_ROLE: ''
  };
  upsertObject_('USERS', 'USER_ID', row);
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (h, i) { if (!current[i]) current[i] = h; });
  sheet.getRange(1, 1, 1, headers.length).setValues([current]);
}

function formatSheet_(sheet, name) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, APP.sheets[name].length).setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold');
  sheet.autoResizeColumns(1, APP.sheets[name].length);
  if (sheet.getFilter()) sheet.getFilter().remove();
  if (sheet.getLastRow() > 1) sheet.getRange(1, 1, sheet.getLastRow(), APP.sheets[name].length).createFilter();
}

function rowsAsObjects_(name) {
  const sh = getSheet_(name), values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(function (r) { return r.some(function (v) { return v !== ''; }); }).map(function (row) {
    const obj = {}; headers.forEach(function (h, i) { obj[h] = serialize_(row[i]); }); return obj;
  });
}

function appendObject_(name, obj) {
  const headers = APP.sheets[name];
  getSheet_(name).appendRow(headers.map(function (h) { return obj[h] === undefined ? '' : obj[h]; }));
}

function upsertObject_(name, key, obj) {
  const found = findRow_(name, key, obj[key]), headers = APP.sheets[name];
  const values = headers.map(function (h) { return obj[h] === undefined ? '' : obj[h]; });
  if (found) getSheet_(name).getRange(found.row, 1, 1, headers.length).setValues([values]);
  else getSheet_(name).appendRow(values);
}

function findRow_(name, key, value) {
  const sh = getSheet_(name), values = sh.getDataRange().getValues(), headers = values[0] || [];
  const idx = headers.indexOf(key); if (idx < 0) return null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx]) === String(value)) {
      const obj = {}; headers.forEach(function (h, j) { obj[h] = values[i][j]; });
      return { row: i + 1, object: obj };
    }
  }
  return null;
}

function findAppointment_(childId, scheduleId) {
  const sh = getSheet_('VACCINE_RECORDS'), values = sh.getDataRange().getValues(), h = values[0];
  const ci = h.indexOf('CHILD_ID'), si = h.indexOf('SCHEDULE_ID');
  for (let i = 1; i < values.length; i++) if (String(values[i][ci]) === String(childId) && String(values[i][si]) === String(scheduleId)) {
    const obj = {}; h.forEach(function (x, j) { obj[x] = values[i][j]; }); return { row: i + 1, object: obj };
  }
  return null;
}

function updateAppointment_(childId, scheduleId, date) {
  const found = findAppointment_(childId, scheduleId); if (!found) return;
  const col = APP.sheets.VACCINE_RECORDS.indexOf('APPOINTMENT_DATE') + 1;
  getSheet_('VACCINE_RECORDS').getRange(found.row, col).setValue(date);
}

function settingsObject_() { const out = {}; rowsAsObjects_('SETTINGS').forEach(function (r) { out[r.KEY] = r.VALUE; }); return out; }
function getSheet_(name) { const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); if (!sh) throw new Error('ไม่พบชีต ' + name + ' กรุณารัน setupSheets()'); return sh; }
function assertInstalled_() { if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CHILDREN')) throw new Error('ยังไม่ได้ติดตั้งระบบ กรุณารัน setupSheets()'); }
function assertRequired_(obj, fields) { fields.forEach(function (f) { if (obj[f] === undefined || obj[f] === null || obj[f] === '') throw new Error('กรุณาระบุ ' + f); }); }
function makeId_(prefix) { return prefix + '-' + Utilities.getUuid().split('-')[0].toUpperCase(); }
function clean_(v) { return v === undefined || v === null ? '' : String(v).trim(); }
function isTrue_(v) { return v === true || String(v).toLowerCase() === 'true' || String(v) === '1'; }
function currentUser_() { return Session.getActiveUser().getEmail() || 'webapp-user'; }
function passwordHash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + String(password), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { const v = b < 0 ? b + 256 : b; return ('0' + v.toString(16)).slice(-2); }).join('');
}

function escapeHtml_(value) {
  return String(value || '').replace(/[&<>"']/g, function (c) {
    return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c];
  });
}
function parseDate_(v) { const d = v instanceof Date ? new Date(v) : new Date(String(v)); if (isNaN(d.getTime())) throw new Error('รูปแบบวันที่ไม่ถูกต้อง: ' + v); return stripTime_(d); }
function parseDateOrBlank_(v) { return v ? parseDate_(v) : ''; }
function stripTime_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function formatDate_(v, withTime) { if (!v) return ''; const d = v instanceof Date ? v : new Date(v); if (isNaN(d)) return ''; return Utilities.formatDate(d, APP.timezone, withTime ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd'); }
function serialize_(v) { return v instanceof Date ? formatDate_(v, true) : v; }
function jsonOutput_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function audit_(action, entity, id, detail) { appendObject_('AUDIT_LOG', { LOG_ID: makeId_('LOG'), TIMESTAMP: new Date(), USER_EMAIL: currentUser_(), ACTION: action, ENTITY: entity, ENTITY_ID: id, DETAIL: detail }); }

