// ══════════════════════════════════════════
// 교복체크 — Apps Script 웹앱
// 설정 방법은 apps-script/README.md 참고
// ══════════════════════════════════════════
const SHEET_ID   = '1ik4BJZsIdVzRRCtwRLoBRtLZ6PEYmzNbvVpN_jgh3N8';
const TOKEN_DAYS = 30;   // 로그인 유지 기간
const MAX_FAILS  = 10;   // 10분 안에 이만큼 틀리면 잠금
const LOCK_SEC   = 600;

function doGet(){
  return json_({ status: 'error', msg: 'POST 요청만 허용됩니다' });
}

function doPost(e){
  let req;
  try{ req = JSON.parse(e.postData.contents); }
  catch(err){ return json_({ status: 'error', msg: '잘못된 요청' }); }
  try{
    if(req.action === 'login') return json_(login_(req.password));
    if(!verifyToken_(req.token)) return json_({ status: 'unauthorized' });
    if(req.action === 'read')    return json_({ status: 'ok', students: readStudents_(), logs: readLogs_() });
    if(req.action === 'saveLog') return json_({ status: 'ok', logs: saveLog_(req) });
    return json_({ status: 'error', msg: '알 수 없는 요청' });
  }catch(err){
    console.error(err);
    return json_({ status: 'error', msg: String(err && err.message || err) });
  }
}

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── 인증 ──
function login_(password){
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('login_fails') || 0);
  if(fails >= MAX_FAILS) return { status: 'locked' };

  const real = PropertiesService.getScriptProperties().getProperty('APP_PASSWORD');
  if(!real) return { status: 'error', msg: 'APP_PASSWORD가 설정되지 않았습니다' };

  if(typeof password !== 'string' || !safeEqual_(password, real)){
    cache.put('login_fails', String(fails + 1), LOCK_SEC);
    Utilities.sleep(1500); // 무작위 대입 지연
    return { status: 'denied' };
  }
  const exp = String(Date.now() + TOKEN_DAYS * 86400000);
  return { status: 'ok', token: exp + '.' + sign_(exp) };
}

// 토큰 = 만료시각.서명 — 비밀번호를 바꾸면 기존 토큰은 모두 무효
function verifyToken_(token){
  if(typeof token !== 'string') return false;
  const i = token.indexOf('.');
  if(i < 0) return false;
  const exp = token.slice(0, i), sig = token.slice(i + 1);
  if(!/^\d+$/.test(exp) || Number(exp) < Date.now()) return false;
  const s = sign_(exp);
  return !!s && safeEqual_(sig, s);
}

function sign_(text){
  const props = PropertiesService.getScriptProperties();
  const password = props.getProperty('APP_PASSWORD');
  if(!password) return '';
  let secret = props.getProperty('TOKEN_SECRET');
  if(!secret){
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('TOKEN_SECRET', secret);
  }
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text, secret + ':' + password));
}

function safeEqual_(a, b){
  a = String(a); b = String(b);
  let diff = a.length ^ b.length;
  for(let i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

// ── 시트 읽기 ──
function sheet_(name){
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if(!sh) throw new Error(`시트 "${name}" 탭이 없습니다`);
  return sh;
}

let tz_ = null;
function dateKey_(v){
  if(v instanceof Date){
    tz_ = tz_ || SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    return Utilities.formatDate(v, tz_, 'yyyy-MM-dd');
  }
  return String(v).trim();
}

// 학생 탭: A 학년 | B 반 | C 번호 | D 이름 | E 성별
function readStudents_(){
  return sheet_('학생').getDataRange().getValues()
    .map(r => ({ grade: parseInt(r[0]), cls: parseInt(r[1]), num: parseInt(r[2]) || 0, name: String(r[3]).trim(), gender: String(r[4]).trim() }))
    .filter(s => s.grade && s.cls && s.name); // 헤더·빈 행 제외
}

// 기록 탭: A 날짜 | B 학년 | C 반 | D 이름
function readLogs_(){
  return sheet_('기록').getDataRange().getValues()
    .map(r => ({ date: dateKey_(r[0]), grade: parseInt(r[1]), cls: parseInt(r[2]), name: String(r[3]).trim() }))
    .filter(x => /^\d{4}-\d{2}-\d{2}$/.test(x.date) && x.grade && x.cls && x.name);
}

// ── 기록 저장 ──
// 그 날짜 기록을 통째로 지우지 않고, 이 기기에서 추가·해제한 학생만 반영한다.
// (다른 기기에서 같은 날 저장한 기록이 덮어써지지 않도록)
function saveLog_(req){
  const date = String(req.date || '');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('날짜 형식 오류');
  const norm = list => (Array.isArray(list) ? list : [])
    .map(x => ({ grade: parseInt(x.grade), cls: parseInt(x.cls), name: String(x.name || '').trim() }))
    .filter(x => x.grade && x.cls && x.name);
  const key = x => `${x.grade}|${x.cls}|${x.name}`;
  const add = norm(req.add);
  const remove = new Set(norm(req.remove).map(key));

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    const sh = sheet_('기록');
    const values = sh.getDataRange().getValues();
    const existing = new Set();
    for(let i = values.length - 1; i >= 0; i--){
      const r = values[i];
      if(dateKey_(r[0]) !== date) continue;
      const k = key({ grade: parseInt(r[1]), cls: parseInt(r[2]), name: String(r[3]).trim() });
      if(remove.has(k)) sh.deleteRow(i + 1);
      else existing.add(k);
    }
    const rows = [];
    add.forEach(x => {
      const k = key(x);
      if(existing.has(k) || remove.has(k)) return;
      existing.add(k);
      rows.push([date, x.grade, x.cls, x.name]);
    });
    if(rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    SpreadsheetApp.flush();
  }finally{
    lock.releaseLock();
  }
  return readLogs_();
}
