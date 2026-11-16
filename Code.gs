/* ============================================================
   우리들의 교환일기 — 서버 (구글 시트 + Apps Script)
   - 이 스크립트는 구글 시트에 붙여서(확장 프로그램 → Apps Script) 씁니다.
   - 처음 한 번 setup() 을 실행하면 시트 탭과 교사 초대 코드가 만들어집니다.
   - 배포: 배포 → 새 배포 → 웹 앱 (실행: 나, 액세스: 모든 사용자)
   ============================================================ */

var TEACHER_INVITE = 'DIARY';  // 교사 초대 코드 (바꾸려면 따옴표 안을 고치고 저장 → 새 버전으로 배포)
var ACCOUNT_API = 'https://script.google.com/macros/s/AKfycbzZlmUD9M4rfo4Y4eEYgNjZH0N8SppUA1Otuo9yMEOqcLfklw8XbS5ZmXj1D3KZY4EG/exec'; // 생각 오락실 통합 계정
var WEEKLY_LIMIT = 2;        // 학생 한 명이 일주일에 쓸 수 있는 일기
var MAX_READERS = 2;         // 일기 한 편을 뽑을 수 있는 사람 수
var TICKET_DAYS = 7;         // 뽑기권 기한
var READ_DAYS = 7;           // 뽑은 일기를 볼 수 있는 기간
var DEFAULT_MIN = 150, DEFAULT_MAX = 800;
var STICKERS = ['공감해요', '재밌어요', '응원해요', '나도 그래'];
var DAY = 86400000;

var TABLES = {
  Teachers: ['nick', 'createdAt'],
  Rooms: ['code', 'teacher', 'name', 'minLen', 'maxLen', 'createdAt', 'active'],
  Students: ['room', 'code', 'no', 'pen', 'pinHash', 'fails', 'lockUntil', 'createdAt', 'active'],
  Diaries: ['id', 'room', 'student', 'pen', 'text', 'original', 'weather', 'len', 'flags', 'status', 'reason', 'edited', 'createdAt', 'decidedAt', 'drawCount'],
  Tickets: ['id', 'room', 'student', 'diaryId', 'createdAt', 'expiresAt', 'usedAt'],
  Draws: ['id', 'room', 'student', 'diaryId', 'createdAt', 'expiresAt', 'sticker']
};

/* ---------------- 처음 설정 ---------------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(TABLES[name]);
  });
  Logger.log('준비 완료! 교사 초대 코드: ' + TEACHER_INVITE);
}

/* ---------------- 입구 ---------------- */
function doPost(e) {
  var out;
  try {
    var d = JSON.parse(e.postData.contents || '{}');
    out = handle(d);
  } catch (err) {
    out = { ok: false, error: err && err.userMsg ? err.userMsg : '서버에서 문제가 생겼어요. 잠시 뒤 다시 해 주세요.', code: err && err.code, detail: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, name: '우리들의 교환일기' })).setMimeType(ContentService.MimeType.JSON);
}

function fail(msg, code) { var e = new Error(msg); e.userMsg = msg; e.code = code || 'ERR'; throw e; }

var ACTIONS = {
  // 교사
  tLogin: tLogin, tRegister: tRegister, tCreateRoom: tCreateRoom, tRoom: tRoom,
  tAddStudents: tAddStudents, tStudent: tStudent, tSettings: tSettings,
  tPending: tPending, tDecide: tDecide, tDone: tDone, tRevise: tRevise,
  // 학생
  sRoster: sRoster, sLogin: sLogin, sSetPin: sSetPin, sSetPen: sSetPen, sHome: sHome, sWrite: sWrite, sCheck: sCheck, sDraw: sDraw, sReact: sReact
};
var WRITES = { tRegister: 1, tCreateRoom: 1, tAddStudents: 1, tStudent: 1, tSettings: 1, tDecide: 1, tRevise: 1, sLogin: 1, sSetPin: 1, sSetPen: 1, sWrite: 1, sDraw: 1, sReact: 1 };

function handle(d) {
  var fn = ACTIONS[d.action];
  if (!fn) fail('알 수 없는 요청이에요.');
  if (!WRITES[d.action]) return fn(d);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) fail('사람이 많아요. 잠시 뒤 다시 눌러 주세요.', 'BUSY');
  try { return fn(d); } finally { lock.releaseLock(); }
}

/* ---------------- 시트 도우미 ---------------- */
var _cache = {};
function table(name) {
  if (_cache[name]) return _cache[name];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) fail('시트가 준비되지 않았어요. setup()을 먼저 실행해 주세요.');
  var vals = sh.getLastRow() > 0 ? sh.getDataRange().getValues() : [TABLES[name]];
  var head = vals[0].map(String);
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var o = { _r: i + 1 };
    head.forEach(function (h, j) { o[h] = vals[i][j]; });
    rows.push(o);
  }
  _cache[name] = { sh: sh, head: head, rows: rows };
  return _cache[name];
}
function insert(name, obj) {
  var t = table(name);
  var line = t.head.map(function (h) { return obj[h] === undefined ? '' : obj[h]; });
  t.sh.appendRow(line);
  obj._r = t.sh.getLastRow();
  t.rows.push(obj);
  return obj;
}
function update(name, obj, fields) {
  var t = table(name);
  Object.keys(fields).forEach(function (k) {
    var c = t.head.indexOf(k);
    if (c < 0) return;
    obj[k] = fields[k];
    t.sh.getRange(obj._r, c + 1, 1, 1).setValues([[fields[k]]]);
  });
}
function find(name, pred) { var r = table(name).rows; for (var i = 0; i < r.length; i++) if (pred(r[i])) return r[i]; return null; }
function filter(name, pred) { return table(name).rows.filter(pred); }

var CODE_CHARS = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
function randCode(n) { var s = ''; for (var i = 0; i < n; i++) s += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length)); return s; }
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
function now() { return Date.now(); }
function num(v) { return Number(v) || 0; }
function bool(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1; }

/* 한국 시간 기준 이번 주 월요일 0시 */
function weekStart(t) {
  var k = new Date(t + 9 * 3600000);
  var dow = (k.getUTCDay() + 6) % 7;
  var mid = Date.UTC(k.getUTCFullYear(), k.getUTCMonth(), k.getUTCDate()) - 9 * 3600000;
  return mid - dow * DAY;
}

/* ---------------- 교사 인증 ---------------- */
function checkAccount(nick, pin) {
  nick = String(nick || '').trim(); pin = String(pin || '').trim();
  if (!nick || !/^\d{4}$/.test(pin)) fail('생각 오락실에 로그인해 주세요.', 'AUTH');
  var cache = CacheService.getScriptCache();
  var key = 'acc:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, nick + '|' + pin));
  if (cache.get(key)) return nick;
  var res = UrlFetchApp.fetch(ACCOUNT_API, { method: 'post', contentType: 'text/plain;charset=utf-8', payload: JSON.stringify({ action: 'me', nick: nick, pin: pin }), muteHttpExceptions: true, followRedirects: true });
  var j = {};
  try { j = JSON.parse(res.getContentText()); } catch (e) { fail('생각 오락실 계정을 확인하지 못했어요. 잠시 뒤 다시 해 주세요.'); }
  if (!j.ok) fail('생각 오락실 계정 정보가 맞지 않아요. 다시 로그인해 주세요.', 'AUTH');
  cache.put(key, '1', 600);
  return nick;
}
function teacher(d) {
  var nick = checkAccount(d.nick, d.pin);
  if (!find('Teachers', function (r) { return String(r.nick) === nick; })) fail('교사 초대 코드를 먼저 넣어 주세요.', 'NOTTEACHER');
  return nick;
}
function myRoom(d, nick) {
  var code = String(d.room || '').toUpperCase();
  var room = find('Rooms', function (r) { return String(r.code) === code && bool(r.active); });
  if (!room || String(room.teacher) !== nick) fail('이 방을 찾을 수 없어요.', 'NOROOM');
  return room;
}
function roomInfo(r) {
  return { code: String(r.code), name: String(r.name), minLen: num(r.minLen) || DEFAULT_MIN, maxLen: num(r.maxLen) || DEFAULT_MAX };
}

function tLogin(d) {
  var nick = checkAccount(d.nick, d.pin);
  var isT = !!find('Teachers', function (r) { return String(r.nick) === nick; });
  var rooms = isT ? filter('Rooms', function (r) { return String(r.teacher) === nick && bool(r.active); }).map(roomInfo) : [];
  return { ok: true, teacher: isT, nick: nick, rooms: rooms };
}
function tRegister(d) {
  var nick = checkAccount(d.nick, d.pin);
  var inv = TEACHER_INVITE;
  if (!inv || String(d.invite || '').trim().toUpperCase() !== String(inv).toUpperCase()) fail('교사 초대 코드가 맞지 않아요.', 'INVITE');
  if (!find('Teachers', function (r) { return String(r.nick) === nick; })) insert('Teachers', { nick: nick, createdAt: now() });
  return tLogin(d);
}
function tCreateRoom(d) {
  var nick = teacher(d);
  var name = String(d.name || '').trim().slice(0, 20);
  if (!name) fail('반 이름을 써 주세요. 예) 6학년 2반');
  var code;
  do { code = randCode(5); } while (find('Rooms', function (r) { return String(r.code) === code; }));
  var count = Math.round(num(d.count));
  if (count < 1 || count > 40) fail('인원 수는 1~40명으로 넣어 주세요.');
  var r = insert('Rooms', { code: code, teacher: nick, name: name, minLen: DEFAULT_MIN, maxLen: DEFAULT_MAX, createdAt: now(), active: true });
  addNumbers(r, count);
  return { ok: true, room: roomInfo(r) };
}
function studentInfo(s) { return { code: String(s.code), no: num(s.no), pen: String(s.pen || ''), hasPin: !!s.pinHash, locked: num(s.lockUntil) > now() }; }
function activeStudents(code) {
  return filter('Students', function (s) { return String(s.room) === String(code) && bool(s.active); }).sort(function (a, b) { return num(a.no) - num(b.no); });
}
function tRoom(d) {
  var nick = teacher(d), room = myRoom(d, nick);
  var studs = activeStudents(room.code);
  var ws = weekStart(now());
  var mine = filter('Diaries', function (x) { return String(x.room) === String(room.code); });
  return {
    ok: true, room: roomInfo(room),
    students: studs.map(function (s) {
      var o = studentInfo(s);
      o.week = mine.filter(function (x) { return String(x.student) === o.code && num(x.createdAt) >= ws && String(x.status) !== 'rejected'; }).length;
      return o;
    }),
    pending: mine.filter(function (x) { return String(x.status) === 'pending'; }).length,
    approvedWeek: mine.filter(function (x) { return String(x.status) === 'approved' && num(x.decidedAt) >= ws; }).length
  };
}
function addNumbers(room, count) {
  var cur = activeStudents(room.code), maxNo = 0, used = {};
  cur.forEach(function (s) { maxNo = Math.max(maxNo, num(s.no)); });
  if (cur.length + count > 40) fail('한 방에는 40명까지 넣을 수 있어요.');
  filter('Students', function (s) { return String(s.room) === String(room.code); }).forEach(function (s) { used[String(s.code)] = 1; });
  for (var i = 1; i <= count; i++) {
    var c; do { c = randCode(6); } while (used[c]);
    used[c] = 1;
    insert('Students', { room: room.code, code: c, no: maxNo + i, pen: '', pinHash: '', fails: 0, lockUntil: '', createdAt: now(), active: true });
  }
}
function tAddStudents(d) {
  var nick = teacher(d), room = myRoom(d, nick);
  var count = Math.round(num(d.count));
  if (count < 1 || count > 40) fail('더할 인원 수를 1~40명으로 넣어 주세요.');
  addNumbers(room, count);
  return tRoom(d);
}
function tStudent(d) {
  var nick = teacher(d), room = myRoom(d, nick);
  var s = find('Students', function (x) { return String(x.room) === String(room.code) && String(x.code) === String(d.code) && bool(x.active); });
  if (!s) fail('학생을 찾을 수 없어요.');
  if (d.op === 'resetPin') update('Students', s, { pinHash: '', fails: 0, lockUntil: '' });
  else if (d.op === 'setPen') applyPen(room, s, d.pen);
  else if (d.op === 'resetPen') update('Students', s, { pen: '' });
  else if (d.op === 'remove') update('Students', s, { active: false });
  return tRoom(d);
}
function tSettings(d) {
  var nick = teacher(d), room = myRoom(d, nick);
  var mn = Math.round(num(d.minLen)), mx = Math.round(num(d.maxLen));
  if (mn < 30 || mn > 1000) fail('최소 글자 수는 30~1000 사이로 정해 주세요.');
  if (mx < mn + 50 || mx > 2000) fail('최대 글자 수는 최소보다 50자 이상 많고 2000자 이하로 정해 주세요.');
  var name = String(d.name || room.name).trim().slice(0, 20);
  update('Rooms', room, { minLen: mn, maxLen: mx, name: name });
  return { ok: true, room: roomInfo(room) };
}
function diaryForTeacher(x, studs) {
  var s = studs[String(x.student)];
  return {
    id: String(x.id), pen: String(x.pen), no: s ? num(s.no) : 0, text: String(x.text), original: String(x.original || ''),
    weather: String(x.weather), len: num(x.len), flags: parseJSON(x.flags, []), status: String(x.status), reason: String(x.reason || ''),
    edited: bool(x.edited), createdAt: num(x.createdAt), decidedAt: num(x.decidedAt), drawCount: num(x.drawCount),
    week: 0
  };
}
function studMap(room) { var m = {}; filter('Students', function (s) { return String(s.room) === String(room.code); }).forEach(function (s) { m[String(s.code)] = s; }); return m; }
function tPending(d) {
  var nick = teacher(d), room = myRoom(d, nick), studs = studMap(room);
  var list = filter('Diaries', function (x) { return String(x.room) === String(room.code) && String(x.status) === 'pending'; })
    .sort(function (a, b) { return num(a.createdAt) - num(b.createdAt); }).map(function (x) { return diaryForTeacher(x, studs); });
  return { ok: true, list: list };
}
function tDone(d) {
  var nick = teacher(d), room = myRoom(d, nick), studs = studMap(room);
  var since = now() - 14 * DAY;
  var list = filter('Diaries', function (x) { return String(x.room) === String(room.code) && String(x.status) !== 'pending' && num(x.decidedAt) >= since; })
    .sort(function (a, b) { return num(b.decidedAt) - num(a.decidedAt); }).slice(0, 80).map(function (x) {
      var o = diaryForTeacher(x, studs);
      o.stickers = stickerCount(String(x.id));
      return o;
    });
  return { ok: true, list: list };
}
function tDecide(d) {
  var nick = teacher(d), room = myRoom(d, nick);
  var x = find('Diaries', function (r) { return String(r.id) === String(d.id) && String(r.room) === String(room.code); });
  if (!x) fail('일기를 찾을 수 없어요.');
  if (String(x.status) !== 'pending') fail('이미 처리한 일기예요.', 'DONE');
  var t = now();
  if (d.decision === 'approve') {
    var fields = { status: 'approved', decidedAt: t, reason: '' };
    if (d.text != null && String(d.text) !== String(x.text)) {
      var nt = String(d.text).trim();
      if (!nt) fail('내용이 비어 있어요.');
      if (nt.length > 4000) fail('너무 길어요.');
      fields.original = x.text; fields.text = nt; fields.edited = true; fields.len = countLen(nt);
      var hard = hardBlocks(nt);
      if (hard.length) fail('고친 글에 아직 보낼 수 없는 말이 있어요: ' + hard.join(', '));
    }
    update('Diaries', x, fields);
    insert('Tickets', { id: uid(), room: room.code, student: x.student, diaryId: x.id, createdAt: t, expiresAt: t + TICKET_DAYS * DAY, usedAt: '' });
  } else if (d.decision === 'reject') {
    var reason = String(d.reason || '').trim().slice(0, 100) || '선생님이 다시 써 보자고 했어요.';
    update('Diaries', x, { status: 'rejected', decidedAt: t, reason: reason });
  } else fail('알 수 없는 처리예요.');
  return tPending(d);
}

/* 처리한 일기 다시 손보기: 고치기(승인된 일기) · 승인으로 바꾸기(반려된 일기) · 삭제 */
function cleanText(t) {
  var nt = String(t || '').replace(/\r\n/g, '\n').trim();
  if (!nt) fail('내용이 비어 있어요.');
  if (nt.length > 4000) fail('너무 길어요.');
  var hard = hardBlocks(nt);
  if (hard.length) fail('고친 글에 아직 보낼 수 없는 말이 있어요: ' + hard.join(', '));
  return nt;
}
function tRevise(d) {
  var nick = teacher(d), room = myRoom(d, nick), t = now();
  var x = find('Diaries', function (r) { return String(r.id) === String(d.id) && String(r.room) === String(room.code); });
  if (!x) fail('일기를 찾을 수 없어요.');
  var st = String(x.status);
  if (d.op === 'edit') {
    if (st !== 'approved') fail('승인된 일기만 고칠 수 있어요.');
    var nt = cleanText(d.text);
    if (nt !== String(x.text)) update('Diaries', x, { original: x.original || x.text, text: nt, edited: true, len: countLen(nt) });
  } else if (d.op === 'approve') {
    if (st !== 'rejected') fail('반려된 일기만 승인으로 바꿀 수 있어요.');
    var f = { status: 'approved', decidedAt: t, reason: '' };
    if (d.text != null && String(d.text).trim() !== String(x.text)) { var t2 = cleanText(d.text); f.original = x.original || x.text; f.text = t2; f.edited = true; f.len = countLen(t2); }
    update('Diaries', x, f);
    insert('Tickets', { id: uid(), room: room.code, student: x.student, diaryId: x.id, createdAt: t, expiresAt: t + TICKET_DAYS * DAY, usedAt: '' });
  } else if (d.op === 'delete') {
    if (st === 'deleted') fail('이미 삭제한 일기예요.');
    if (st === 'pending') fail('승인함에서 먼저 처리해 주세요.');
    update('Diaries', x, { status: 'deleted', reason: '선생님이 내렸어요.', decidedAt: t });
  } else fail('알 수 없는 처리예요.');
  return tDone(d);
}

/* ---------------- 글 확인(필터) ---------------- */
var BLOCK_WORDS = ['씨발', '시발', '씨바', '씨빨', 'ㅆㅂ', 'ㅅㅂ', '병신', 'ㅂㅅ', '븅신', '개새끼', '개새기', '개색기', '개색끼', '좆', '존나', '졸라', 'ㅈㄴ', '지랄', 'ㅈㄹ', '엠창', '느금', '니애미', '니미', '섹스', 'fuck', 'shit', 'bitch'];
var CARE_WORDS = ['죽고 싶', '죽고싶', '죽어버리', '자살', '자해', '사라지고 싶', '사라지고싶', '살기 싫', '살기싫', '태어나지 말', '맞았', '때렸', '때려', '괴롭', '왕따', '따돌림', '무서워', '혼자야', '아무도 없'];
var CHECK_WORDS = ['초등학교', '중학교', '학교 이름', '아파트', '동네', '사는 곳', '주소', '문구점', '편의점', '학원', '인스타', '카톡', '카카오', '오픈채팅', '페메', '페북', '디엠', '틱톡', '유튜브', '아이디', '번호', '전화', '연락', '만나자', '만날래', '실명', '내 이름', '미친', '새끼', '꺼져', '닥쳐', '바보', '멍청'];

function hardBlocks(text) {
  var found = [], low = text.toLowerCase(), compact = low.replace(/[\s\-.·_]/g, '');
  if (/01[016789][\s\-.]*\d{3,4}[\s\-.]*\d{4}/.test(text) || /\d{7,}/.test(compact)) found.push('전화번호');
  if (/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/.test(text)) found.push('이메일');
  if (/https?:\/\/|www\.|\.(com|net|kr|io|me|ly)\b/i.test(text)) found.push('인터넷 주소');
  if (/@\s*[A-Za-z0-9_.]{3,}/.test(text)) found.push('@아이디');
  BLOCK_WORDS.forEach(function (w) { if (compact.indexOf(w) >= 0) found.push('나쁜 말'); });
  return found.filter(function (v, i, a) { return a.indexOf(v) === i; });
}
function softFlags(text) {
  var flags = [], low = text.toLowerCase();
  function add(kind, word) { if (!flags.some(function (f) { return f.w === word; })) flags.push({ k: kind, w: word }); }
  CARE_WORDS.forEach(function (w) { if (low.indexOf(w) >= 0) add('care', w); });
  CHECK_WORDS.forEach(function (w) { if (low.indexOf(w) >= 0) add('check', w); });
  var m;
  var reAl = /[A-Za-z][A-Za-z0-9_.]{4,}/g;                  // 영어 아이디처럼 보이는 말
  while ((m = reAl.exec(text))) add('check', m[0]);
  var reSp = /(?:[A-Za-z0-9]\s){3,}[A-Za-z0-9]/g;             // 띄어 쓴 영어·숫자
  while ((m = reSp.exec(text))) add('check', m[0]);
  var reKn = /(?:공|영|일|이|삼|사|오|육|륙|칠|팔|구){7,}/g;    // 한글로 쓴 숫자
  while ((m = reKn.exec(text))) add('check', m[0]);
  var reNum = /\d{3,}/g;
  while ((m = reNum.exec(text))) add('check', m[0]);
  var reShop = /[가-힣]{1,8}(초등학교|초|문구|마트|학원|아파트)(?=[\s,.!?에의은는이가을를도]|$)/g;
  while ((m = reShop.exec(text))) add('check', m[0]);
  return flags;
}
function countLen(text) {
  var t = String(text);
  try { t = t.replace(/\p{Extended_Pictographic}/gu, ''); } catch (e) {}
  t = t.replace(/[ㄱ-ㅎㅏ-ㅣ]/g, '');            // ㅋㅋ ㅎㅎ ㅠㅠ 같은 낱자
  t = t.replace(/(.)\1{2,}/g, '$1');             // 같은 글자 반복은 한 글자로
  t = t.replace(/[\s~!.,?…·\-_=+*^]/g, '');     // 띄어쓰기와 꾸밈 기호
  return t.length;
}

/* ---------------- 학생 ---------------- */
function findRoom(d) {
  var rc = String(d.room || '').trim().toUpperCase();
  var cache = CacheService.getScriptCache(), fk = 'fail:' + rc;
  var fails = num(cache.get(fk));
  if (fails >= 30) fail('잘못 넣은 횟수가 많아요. 10분 뒤에 다시 해 주세요.', 'LOCK');
  var room = find('Rooms', function (r) { return String(r.code) === rc && bool(r.active); });
  if (!room) { cache.put(fk, String(fails + 1), 600); fail('방 코드를 다시 확인해 주세요.', 'NOROOM'); }
  return room;
}
function pinHash(room, s, pin) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'diary|' + room.code + '|' + s.code + '|' + pin));
}
function rosterStudent(room, no) {
  var s = find('Students', function (x) { return String(x.room) === String(room.code) && num(x.no) === num(no) && bool(x.active); });
  if (!s) fail('번호를 다시 골라 주세요.', 'NOSTUDENT');
  return s;
}
function student(d) {
  var room = findRoom(d), s = rosterStudent(room, d.no), t = now();
  if (!s.pinHash) fail('비밀번호를 먼저 정해 주세요.', 'NOPIN');
  if (num(s.lockUntil) > t) fail('비밀번호를 여러 번 틀려서 잠겼어요. ' + Math.ceil((num(s.lockUntil) - t) / 60000) + '분 뒤에 다시 해 주세요.', 'LOCKED');
  if (String(s.pinHash) !== pinHash(room, s, String(d.pin || ''))) {
    var f = num(s.fails) + 1;
    if (f >= 5) { update('Students', s, { fails: 0, lockUntil: t + 10 * 60000 }); fail('비밀번호를 5번 틀려서 10분 동안 잠겼어요. 잊어버렸다면 선생님께 말씀드려요.', 'LOCKED'); }
    update('Students', s, { fails: f });
    fail('비밀번호가 맞지 않아요. (' + f + '/5)', 'WRONGPIN');
  }
  if (num(s.fails)) update('Students', s, { fails: 0 });
  return { room: room, s: s };
}
function sRoster(d) {
  var room = findRoom(d);
  return { ok: true, roomName: String(room.name), list: activeStudents(room.code).map(function (s) { return { no: num(s.no), hasPin: !!s.pinHash }; }) };
}
function sSetPin(d) {
  var room = findRoom(d), s = rosterStudent(room, d.no);
  if (s.pinHash) fail('이미 비밀번호가 있어요. 잊어버렸다면 선생님께 말씀드려요.', 'HASPIN');
  var pin = String(d.pin || '');
  if (!/^\d{4}$/.test(pin)) fail('비밀번호는 숫자 4자리로 정해 주세요.');
  if (/^(\d)\1{3}$/.test(pin) || pin === '1234' || pin === '4321') fail('너무 쉬운 비밀번호예요. 다른 숫자로 정해 주세요.');
  update('Students', s, { pinHash: pinHash(room, s, pin), fails: 0, lockUntil: '' });
  return sLogin(d);
}
var PEN_RE = /^[가-힣A-Za-z0-9]{2,8}$/;
function sLogin(d) {
  var a = student(d);
  return { ok: true, pen: String(a.s.pen || ''), roomName: String(a.room.name), minLen: num(a.room.minLen) || DEFAULT_MIN, maxLen: num(a.room.maxLen) || DEFAULT_MAX };
}
function applyPen(room, st, pen) {
  pen = String(pen || '').trim();
  if (!PEN_RE.test(pen)) fail('필명은 한글·영어·숫자 2~8자로 써 주세요.');
  if (hardBlocks(pen).length) fail('필명으로 쓸 수 없는 말이 있어요.');
  if (find('Students', function (x) { return String(x.room) === String(room.code) && bool(x.active) && String(x.pen) === pen && x !== st; })) fail('우리 반에 같은 필명이 있어요. 다른 필명을 골라 주세요.');
  update('Students', st, { pen: pen });
  // 아직 선생님 확인 전인 일기는 새 필명으로 (이미 배달된 일기는 쓸 때의 필명 그대로)
  filter('Diaries', function (x) { return String(x.room) === String(room.code) && String(x.student) === String(st.code) && String(x.status) === 'pending'; })
    .forEach(function (x) { update('Diaries', x, { pen: pen }); });
  return pen;
}
function sSetPen(d) {
  var a = student(d);
  return { ok: true, pen: applyPen(a.room, a.s, d.pen) };
}
function stickerCount(diaryId) {
  var c = {};
  filter('Draws', function (r) { return String(r.diaryId) === diaryId && r.sticker; }).forEach(function (r) { c[r.sticker] = (c[r.sticker] || 0) + 1; });
  return c;
}
function liveTickets(a, t) {
  return filter('Tickets', function (k) { return String(k.room) === String(a.room.code) && String(k.student) === String(a.s.code) && !k.usedAt && num(k.expiresAt) > t; })
    .sort(function (x, y) { return num(x.expiresAt) - num(y.expiresAt); });
}
function poolSize(a) {
  var t = now(), mine = {};
  filter('Draws', function (r) { return String(r.room) === String(a.room.code) && String(r.student) === String(a.s.code); }).forEach(function (r) { mine[String(r.diaryId)] = 1; });
  return filter('Diaries', function (x) { return String(x.status) === 'approved' && String(x.room) !== String(a.room.code) && num(x.drawCount) < MAX_READERS && !mine[String(x.id)]; }).length;
}
function sHome(d) {
  var a = student(d), t = now(), ws = weekStart(t);
  var mine = filter('Diaries', function (x) { return String(x.room) === String(a.room.code) && String(x.student) === String(a.s.code); })
    .sort(function (x, y) { return num(y.createdAt) - num(x.createdAt); });
  var weekUsed = mine.filter(function (x) { return num(x.createdAt) >= ws && String(x.status) !== 'rejected'; }).length;
  var diaries = mine.slice(0, 12).map(function (x) {
    return { id: String(x.id), status: String(x.status), reason: String(x.reason || ''), createdAt: num(x.createdAt), text: String(x.text), original: String(x.original || ''), edited: bool(x.edited), weather: String(x.weather), readers: num(x.drawCount), stickers: stickerCount(String(x.id)) };
  });
  var received = filter('Draws', function (r) { return String(r.room) === String(a.room.code) && String(r.student) === String(a.s.code) && num(r.expiresAt) > t; })
    .sort(function (x, y) { return num(y.createdAt) - num(x.createdAt); })
    .map(function (r) {
      var x = find('Diaries', function (q) { return String(q.id) === String(r.diaryId); });
      if (!x || String(x.status) !== 'approved') return null;
      return { drawId: String(r.id), pen: String(x.pen), text: String(x.text), weather: String(x.weather), writtenAt: num(x.createdAt), drawnAt: num(r.createdAt), expiresAt: num(r.expiresAt), sticker: String(r.sticker || '') };
    }).filter(Boolean);
  return {
    ok: true, pen: String(a.s.pen || ''), roomName: String(a.room.name),
    minLen: num(a.room.minLen) || DEFAULT_MIN, maxLen: num(a.room.maxLen) || DEFAULT_MAX,
    weekUsed: weekUsed, weekLimit: WEEKLY_LIMIT,
    tickets: liveTickets(a, t).map(function (k) { return { id: String(k.id), expiresAt: num(k.expiresAt) }; }),
    pool: poolSize(a), diaries: diaries, received: received, stickers: STICKERS, now: t
  };
}
function sCheck(d) {
  var text = String(d.text || '');
  return { ok: true, len: countLen(text), blocks: hardBlocks(text), flags: softFlags(text).map(function (f) { return f.w; }) };
}
var WEATHERS = ['맑음', '흐림', '비', '눈'];
function sWrite(d) {
  var a = student(d);
  if (!a.s.pen) fail('필명을 먼저 정해 주세요.', 'NOPEN');
  var text = String(d.text || '').replace(/\r\n/g, '\n').trim();
  var mn = num(a.room.minLen) || DEFAULT_MIN, mx = num(a.room.maxLen) || DEFAULT_MAX;
  var len = countLen(text);
  if (len < mn) fail('아직 ' + (mn - len) + '자 더 써야 해요.', 'SHORT');
  if (len > mx) fail(mx + '자까지만 쓸 수 있어요.', 'LONG');
  if (text.length > 4000) fail('너무 길어요.', 'LONG');
  var hard = hardBlocks(text);
  if (hard.length) fail('보낼 수 없는 내용이 있어요: ' + hard.join(', '), 'BLOCK');
  var t = now(), ws = weekStart(t);
  var used = filter('Diaries', function (x) { return String(x.room) === String(a.room.code) && String(x.student) === String(a.s.code) && num(x.createdAt) >= ws && String(x.status) !== 'rejected'; }).length;
  if (used >= WEEKLY_LIMIT) fail('이번 주에는 ' + WEEKLY_LIMIT + '편을 다 썼어요. 다음 주 월요일에 또 써요!', 'LIMIT');
  var weather = WEATHERS.indexOf(d.weather) >= 0 ? d.weather : '';
  insert('Diaries', { id: uid(), room: a.room.code, student: a.s.code, pen: a.s.pen, text: text, original: '', weather: weather, len: len, flags: JSON.stringify(softFlags(text)), status: 'pending', reason: '', edited: false, createdAt: t, decidedAt: '', drawCount: 0 });
  return sHome(d);
}
function sDraw(d) {
  var a = student(d), t = now();
  var tickets = liveTickets(a, t);
  if (!tickets.length) fail('뽑기권이 없어요. 일기를 써서 선생님 확인을 받으면 생겨요.', 'NOTICKET');
  var mineDrawn = {}, recentAuthors = {};
  filter('Draws', function (r) { return String(r.room) === String(a.room.code) && String(r.student) === String(a.s.code); }).forEach(function (r) {
    mineDrawn[String(r.diaryId)] = 1;
    if (num(r.createdAt) > t - 7 * DAY) {
      var x = find('Diaries', function (q) { return String(q.id) === String(r.diaryId); });
      if (x) recentAuthors[String(x.room) + '/' + String(x.student)] = 1;
    }
  });
  var pool = filter('Diaries', function (x) {
    return String(x.status) === 'approved' && String(x.room) !== String(a.room.code) && num(x.drawCount) < MAX_READERS && !mineDrawn[String(x.id)];
  });
  if (!pool.length) fail('아직 도착한 쪽지가 없어요. 뽑기권은 그대로 남아 있으니 조금 뒤에 다시 와 주세요.', 'EMPTY');
  var fresh = pool.filter(function (x) { return !recentAuthors[String(x.room) + '/' + String(x.student)]; });
  if (fresh.length) pool = fresh;
  var least = Math.min.apply(null, pool.map(function (x) { return num(x.drawCount); }));
  pool = pool.filter(function (x) { return num(x.drawCount) === least; });
  var pick = pool[Math.floor(Math.random() * pool.length)];
  var tk = tickets[0];
  update('Tickets', tk, { usedAt: t });
  update('Diaries', pick, { drawCount: num(pick.drawCount) + 1 });
  var dr = insert('Draws', { id: uid(), room: a.room.code, student: a.s.code, diaryId: pick.id, createdAt: t, expiresAt: t + READ_DAYS * DAY, sticker: '' });
  return { ok: true, diary: { drawId: String(dr.id), pen: String(pick.pen), text: String(pick.text), weather: String(pick.weather), writtenAt: num(pick.createdAt), drawnAt: t, expiresAt: num(dr.expiresAt), sticker: '' } };
}
function sReact(d) {
  var a = student(d);
  if (STICKERS.indexOf(d.sticker) < 0 && d.sticker !== '') fail('알 수 없는 스티커예요.');
  var r = find('Draws', function (x) { return String(x.id) === String(d.drawId) && String(x.room) === String(a.room.code) && String(x.student) === String(a.s.code); });
  if (!r) fail('일기를 찾을 수 없어요.');
  if (num(r.expiresAt) <= now()) fail('볼 수 있는 기간이 지났어요.');
  update('Draws', r, { sticker: d.sticker });
  return { ok: true, sticker: d.sticker };
}

function parseJSON(s, dflt) { try { return s ? JSON.parse(s) : dflt; } catch (e) { return dflt; } }
