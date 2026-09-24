/**
 * 태안 AI-DREAM Lab 오픈 랩 데이 투어 예약 (v1)
 * 구성: 구글 시트(컨테이너 바인딩) + Apps Script 웹앱 + 메일 알림
 * 관리자 처리: Reservations 시트의 상태 열을 승인 또는 반려로 바꾸면 안내 메일 자동 발송
 */

const TZ = 'Asia/Seoul';
const SHEET = { DAYS: 'OpenDays', RES: 'Reservations', CONFIG: 'Config', STATS: 'Stats' };

const DAY_HEADERS = ['회차ID', '날짜', '시작', '종료', '정원', '기관당 상한', '신청 마감일', '상태', '비고'];
const RES_HEADERS = ['예약번호', '신청 일시', '회차ID', '기관명', '기관 유형', '직위', '대표자 성명', '연락처',
  '이메일', '인원', '방문 목적', '공문 접수', '동의 일시', '상태', '처리 일시', '메일 발송 기록',
  '안내 메모(메일에 포함)', '내부 메모'];
const COL = { NO: 1, APPLIED: 2, SESSION: 3, ORG: 4, ORGTYPE: 5, TITLE: 6, NAME: 7, PHONE: 8, EMAIL: 9,
  COUNT: 10, PURPOSE: 11, DOC: 12, CONSENT: 13, STATUS: 14, DONE: 15, MAILED: 16, NOTE: 17, MEMO: 18 };

const DAY_STATUS = ['공개', '마감', '비공개'];
const RES_STATUS = ['신청', '승인', '반려', '취소'];
const ORG_TYPES = ['초등학교', '중학교', '고등학교', '교육청·교육지원청', '대학·연구기관', '기타 교육기관'];
const ACTIVE = ['신청', '승인'];

// 세션 기본값(제안치). 오픈 랩 데이 추가 메뉴에서 사용합니다.
const DEFAULT_SESSIONS = [{ start: '15:00', end: '15:40' }, { start: '15:50', end: '16:30' }];
const DEFAULT_CAPACITY = 15;
const DEFAULT_ORG_CAP = 5;
const DEADLINE_DAYS = 7;
// 하루 전체 신청 건수 상한(신청 일시 기준). Config 시트에 하루 신청 상한 항목을 추가하면 그 값을 씁니다.
const DAILY_SUBMIT_LIMIT = 40;

const CONFIG_DEFAULTS = [
  ['관리자 메일', ''],
  ['학교명', '태안초등학교'],
  ['장소', '태안초등학교 본관 3층 AI-DREAM Lab'],
  ['오시는 길 안내', '차량은 정문이 아닌 후문으로 들어와 교내 주차장에 주차하세요.'],
  ['촬영 안내', '공간 촬영은 가능하며, 학생이 포함된 촬영은 제한됩니다.'],
  ['공문 안내', '승인된 기관은 방문일 전까지 방문 협조 공문(수신: 태안초등학교)을 보내 주세요. 공문 접수 후 방문이 확정됩니다.'],
  ['보유 기간 안내', '방문일이 속한 학년도 종료 시까지 보유 후 파기'],
  ['파기 기준일', ''],
  ['예약 페이지 주소', '']
];

/* ───────── 공통 ───────── */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function isDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(s); }

function weekday_(dateStr) {
  const p = dateStr.split('-').map(Number);
  return ['일', '월', '화', '수', '목', '금', '토'][new Date(p[0], p[1] - 1, p[2]).getDay()];
}

function addDays_(dateStr, n) {
  const p = dateStr.split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2] + n);
  const mm = ('0' + (d.getMonth() + 1)).slice(-2);
  const dd = ('0' + d.getDate()).slice(-2);
  return d.getFullYear() + '-' + mm + '-' + dd;
}

/** 입력값 정리: 제어문자 제거, 길이 제한, 수식으로 해석될 수 있는 선행 문자 제거 */
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001F\u007F]/g, ' ').trim();
  s = s.replace(/^[=+\-@\s]+/, '');
  return s.length > max ? s.substring(0, max) : s;
}

function orgKey_(s) { return String(s).replace(/\s/g, '').toLowerCase(); }

/** 연락처: 숫자 11자리(0으로 시작)만 허용, 000-0000-0000 형태로 저장 */
function formatPhone_(raw) {
  const d = String(raw || '').replace(/\D/g, '');
  if (!/^0\d{10}$/.test(d)) return '';
  return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
}

function getConfig_() {
  const sh = ss_().getSheetByName(SHEET.CONFIG);
  const cfg = {};
  if (!sh || sh.getLastRow() < 2) return cfg;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(function (r) {
    if (r[0]) cfg[r[0].trim()] = r[1].trim();
  });
  return cfg;
}

function sendMail_(to, subject, body) {
  if (!to) return false;
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: '태안초등학교 AI-DREAM Lab' });
    return true;
  } catch (err) {
    console.error('메일 발송 실패: ' + to + ' / ' + err);
    return false;
  }
}

/* ───────── 데이터 읽기 ───────── */

function readSessions_() {
  const sh = ss_().getSheetByName(SHEET.DAYS);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, DAY_HEADERS.length).getDisplayValues();
  const list = [];
  rows.forEach(function (r) {
    const id = r[0].trim(), date = r[1].trim();
    if (!id || !isDate_(date)) return;
    const deadline = isDate_(r[6].trim()) ? r[6].trim() : date;
    list.push({
      id: id, date: date, weekday: weekday_(date), start: r[2].trim(), end: r[3].trim(),
      capacity: parseInt(r[4], 10) || 0, orgCap: parseInt(r[5], 10) || 0,
      deadline: deadline, status: r[7].trim()
    });
  });
  return list;
}

function readReservations_() {
  const sh = ss_().getSheetByName(SHEET.RES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, RES_HEADERS.length).getDisplayValues();
}

/** 회차별 점유 인원(신청 + 승인)과 기관별 인원 */
function usage_(rows, skipRowIndex) {
  const map = {};
  rows.forEach(function (r, i) {
    if (i === skipRowIndex) return;
    if (ACTIVE.indexOf(r[COL.STATUS - 1]) < 0) return;
    const sid = r[COL.SESSION - 1];
    const n = parseInt(r[COL.COUNT - 1], 10) || 0;
    if (!map[sid]) map[sid] = { total: 0, byOrg: {} };
    map[sid].total += n;
    const key = orgKey_(r[COL.ORG - 1]);
    map[sid].byOrg[key] = (map[sid].byOrg[key] || 0) + n;
  });
  return map;
}

function sessionLabel_(s) {
  return s.date + '(' + s.weekday + ') ' + s.start + '~' + s.end;
}

/* ───────── API (프런트엔드 분리형) ─────────
 * 예약 화면(index.html)은 별도 호스팅(GitHub Pages 등)에서 열리고,
 * 이 웹앱은 JSON만 주고받습니다. 브라우저 교차 출처 요청을 위해
 * POST 본문은 text/plain 으로 보내야 합니다(프런트엔드 api() 참고).
 */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'pageData') return json_(getPageData());
  return json_({ ok: false, message: '예약 페이지에서 접속하세요.', url: getConfig_()['예약 페이지 주소'] || '' });
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, message: '요청 형식이 잘못되었습니다.' });
  }
  const action = String(body.action || '');
  try {
    if (action === 'pageData') return json_(getPageData());
    if (action === 'submit') return json_(submitReservation(body.form || {}));
    if (action === 'lookup') return json_(guarded_(body.no, function () { return lookupReservation(body.no, body.phone); }));
    if (action === 'cancel') return json_(guarded_(body.no, function () { return cancelReservation(body.no, body.phone); }));
    if (action.indexOf('admin.') === 0) return json_(adminApi_(action.slice(6), body));
    return json_({ ok: false, message: '알 수 없는 요청입니다.' });
  } catch (err) {
    console.error('API 오류(' + action + '): ' + err);
    return json_({ ok: false, message: '처리 중 오류가 발생했습니다. 잠시 후 다시 시도하세요.' });
  }
}

/** 조회·취소 무차별 대입 완화: 같은 예약번호로 10분 내 실패 10회면 잠시 차단 */
function guarded_(no, fn) {
  const key = 'fail:' + String(no || '').trim().toUpperCase();
  const cache = CacheService.getScriptCache();
  const fails = parseInt(cache.get(key), 10) || 0;
  if (fails >= 10) return { ok: false, message: '시도 횟수를 초과했습니다. 10분 후 다시 시도하세요.' };
  const res = fn();
  if (!res.ok) cache.put(key, String(fails + 1), 600);
  return res;
}

/** 화면 초기 데이터: 공개 회차와 잔여 인원 */
function getPageData() {
  const today = today_();
  const use = usage_(readReservations_(), -1);
  const cfg = getConfig_();
  const sessions = readSessions_()
    .filter(function (s) { return s.status !== '비공개' && DAY_STATUS.indexOf(s.status) >= 0 && s.date >= today; })
    .map(function (s) {
      const taken = use[s.id] ? use[s.id].total : 0;
      const remaining = Math.max(0, s.capacity - taken);
      let closedReason = '';
      if (s.status === '마감') closedReason = '신청 마감';
      else if (s.deadline < today) closedReason = '신청 기간 종료';
      else if (remaining <= 0) closedReason = '정원 마감';
      return {
        id: s.id, date: s.date, weekday: s.weekday, start: s.start, end: s.end,
        capacity: s.capacity, orgCap: s.orgCap, deadline: s.deadline,
        remaining: remaining, closedReason: closedReason
      };
    })
    .sort(function (a, b) { return (a.date + a.start).localeCompare(b.date + b.start); });
  return {
    sessions: sessions,
    orgTypes: ORG_TYPES,
    place: cfg['장소'] || '',
    retention: cfg['보유 기간 안내'] || '방문일이 속한 학년도 종료 시까지 보유 후 파기'
  };
}

/** 예약 신청 */
function submitReservation(f) {
  f = f || {};
  if (f.website) return { ok: false, message: '신청을 처리할 수 없습니다.' };

  const org = clean_(f.org, 40), orgType = clean_(f.orgType, 20), title = clean_(f.title, 20);
  const name = clean_(f.name, 20), email = clean_(f.email, 80), purpose = clean_(f.purpose, 300);
  const phone = formatPhone_(f.phone);
  const count = parseInt(f.count, 10);
  const sessionId = clean_(f.sessionId, 30);

  if (!sessionId) return { ok: false, message: '방문할 회차를 선택하세요.' };
  if (!org || !title || !name) return { ok: false, message: '기관명, 직위, 대표자 성명을 입력하세요.' };
  if (ORG_TYPES.indexOf(orgType) < 0) return { ok: false, message: '기관 유형을 선택하세요.' };
  if (!phone) return { ok: false, message: '연락처는 010-0000-0000 형식(숫자 11자리)으로 입력하세요.' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return { ok: false, message: '이메일 주소를 확인하세요. @와 도메인이 있어야 합니다.' };
  if (!(count >= 1)) return { ok: false, message: '방문 인원을 입력하세요.' };
  if (!purpose) return { ok: false, message: '방문 목적을 입력하세요.' };
  if (f.consent !== true) return { ok: false, message: '개인정보 수집·이용에 동의해야 신청할 수 있습니다.' };

  const cfg = getConfig_();
  const dailyLimit = parseInt(cfg['하루 신청 상한'], 10) || DAILY_SUBMIT_LIMIT;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return { ok: false, message: '신청이 몰리고 있습니다. 잠시 후 다시 시도하세요.' };
  }

  let no, session;
  try {
    const today = today_();
    session = readSessions_().filter(function (s) { return s.id === sessionId; })[0];
    if (!session || session.status !== '공개' || session.date < today || session.deadline < today) {
      return { ok: false, message: '선택한 회차는 신청이 마감되었습니다. 다른 회차를 선택하세요.' };
    }
    const rows = readReservations_();
    const todayCount = rows.filter(function (r) { return String(r[COL.APPLIED - 1]).indexOf(today) === 0; }).length;
    if (todayCount >= dailyLimit) {
      return { ok: false, message: '오늘 접수 가능한 신청 건수를 넘었습니다. 내일 다시 신청하거나 학교로 문의하세요.' };
    }
    const phoneDigits = phone.replace(/\D/g, ''), emailKey = email.toLowerCase();
    const dup = rows.some(function (r) {
      if (r[COL.SESSION - 1] !== sessionId || ACTIVE.indexOf(r[COL.STATUS - 1]) < 0) return false;
      return r[COL.PHONE - 1].replace(/\D/g, '') === phoneDigits || r[COL.EMAIL - 1].toLowerCase() === emailKey;
    });
    if (dup) {
      return { ok: false, message: '이 회차에 같은 연락처 또는 이메일로 신청한 예약이 있습니다. 예약 조회에서 확인하세요.' };
    }
    const u = usage_(rows, -1)[sessionId] || { total: 0, byOrg: {} };
    const remaining = session.capacity - u.total;
    if (count > remaining) {
      return { ok: false, message: '잔여 인원은 ' + Math.max(0, remaining) + '명입니다. 인원을 조정하거나 다른 회차를 선택하세요.' };
    }
    const orgUsed = u.byOrg[orgKey_(org)] || 0;
    if (session.orgCap > 0 && orgUsed + count > session.orgCap) {
      return { ok: false, message: '한 기관은 회차당 ' + session.orgCap + '명까지 신청할 수 있습니다. (이미 신청한 인원 ' + orgUsed + '명)' };
    }

    no = newReservationNo_(rows);

    const stamp = now_();
    ss_().getSheetByName(SHEET.RES).appendRow([
      no, stamp, sessionId, org, orgType, title, name, phone, email, count, purpose,
      '', stamp, '신청', '', '', '', ''
    ]);
  } finally {
    lock.releaseLock();
  }

  const label = sessionLabel_(session);
  sendMail_(email, '[태안 AI-DREAM Lab] 투어 예약 신청이 접수되었습니다 (' + no + ')',
    [org + ' ' + name + ' 선생님께', '',
      '태안 AI-DREAM Lab 오픈 랩 데이 투어 예약 신청이 접수되었습니다.',
      '담당자 확인 후 승인 여부를 메일로 안내합니다.', '',
      '- 예약번호: ' + no, '- 일시: ' + label, '- 인원: ' + count + '명', '',
      '예약 조회와 취소에는 예약번호와 신청 시 입력한 연락처가 필요합니다.',
      cfg['예약 페이지 주소'] || ''].join('\n'));
  sendMail_(cfg['관리자 메일'], '[투어 예약 신청] ' + org + ' ' + count + '명 / ' + label,
    ['새 투어 예약 신청이 접수되었습니다.', '',
      '- 예약번호: ' + no, '- 일시: ' + label, '- 기관: ' + org + ' (' + orgType + ')',
      '- 대표자: ' + title + ' ' + name, '- 인원: ' + count + '명', '- 방문 목적: ' + purpose, '',
      'Reservations 시트의 상태 열을 승인 또는 반려로 변경하면 신청자에게 안내 메일이 발송됩니다.',
      ss_().getUrl()].join('\n'));

  return { ok: true, no: no, label: label, count: count };
}

/** 예약번호: 연도-순번 3자리 (예: 2026-001). 연도별로 1부터. 호출 측에서 스크립트 락을 잡고 있어야 함 */
function newReservationNo_(rows) {
  const year = Utilities.formatDate(new Date(), TZ, 'yyyy');
  let max = 0;
  rows.forEach(function (r) {
    const m = /^(\d{4})-(\d{3,})$/.exec(String(r[COL.NO - 1] || '').trim());
    if (m && m[1] === year) max = Math.max(max, parseInt(m[2], 10));
  });
  return year + '-' + String(max + 1).padStart(3, '0');
}

/** 예약번호 + 연락처(전체)로 행 찾기. 반환: { rowIndex(0부터), row } */
function findReservation_(no, phone) {
  no = String(no || '').trim();
  const digits = String(phone || '').replace(/\D/g, '');
  if (!/^\d{4}-\d{3,}$/.test(no) || digits.length !== 11) return null;
  const rows = readReservations_();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][COL.NO - 1] === no && rows[i][COL.PHONE - 1].replace(/\D/g, '') === digits) return { rowIndex: i, row: rows[i] };
  }
  return null;
}

function lookupReservation(no, phone) {
  const hit = findReservation_(no, phone);
  if (!hit) return { ok: false, message: '일치하는 예약이 없습니다. 예약번호와 연락처를 확인하세요.' };
  const r = hit.row;
  const s = readSessions_().filter(function (x) { return x.id === r[COL.SESSION - 1]; })[0];
  const status = r[COL.STATUS - 1];
  return {
    ok: true, no: r[COL.NO - 1], org: r[COL.ORG - 1], count: r[COL.COUNT - 1], status: status,
    label: s ? sessionLabel_(s) : r[COL.SESSION - 1],
    note: status === '반려' ? r[COL.NOTE - 1] : '',
    canCancel: ACTIVE.indexOf(status) >= 0 && (!s || s.date >= today_())
  };
}

function cancelReservation(no, phone) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return { ok: false, message: '잠시 후 다시 시도하세요.' };
  }
  let r, label;
  try {
    const hit = findReservation_(no, phone);
    if (!hit) return { ok: false, message: '일치하는 예약이 없습니다.' };
    r = hit.row;
    if (ACTIVE.indexOf(r[COL.STATUS - 1]) < 0) return { ok: false, message: '이미 ' + r[COL.STATUS - 1] + ' 처리된 예약입니다.' };
    const s = readSessions_().filter(function (x) { return x.id === r[COL.SESSION - 1]; })[0];
    if (s && s.date < today_()) return { ok: false, message: '방문일이 지난 예약은 취소할 수 없습니다.' };
    label = s ? sessionLabel_(s) : r[COL.SESSION - 1];
    const sh = ss_().getSheetByName(SHEET.RES);
    const rowNo = hit.rowIndex + 2;
    sh.getRange(rowNo, COL.STATUS).setValue('취소');
    sh.getRange(rowNo, COL.DONE).setValue(now_());
    sh.getRange(rowNo, COL.MAILED).setValue('본인 취소');
  } finally {
    lock.releaseLock();
  }
  const cfg = getConfig_();
  sendMail_(cfg['관리자 메일'], '[투어 예약 취소] ' + r[COL.ORG - 1] + ' ' + r[COL.COUNT - 1] + '명 / ' + label,
    '신청자가 예약을 취소했습니다.\n\n- 예약번호: ' + r[COL.NO - 1] + '\n- 일시: ' + label);
  sendMail_(r[COL.EMAIL - 1], '[태안 AI-DREAM Lab] 투어 예약이 취소되었습니다 (' + r[COL.NO - 1] + ')',
    r[COL.ORG - 1] + ' ' + r[COL.NAME - 1] + ' 선생님께\n\n아래 예약이 취소되었습니다.\n\n- 예약번호: ' +
    r[COL.NO - 1] + '\n- 일시: ' + label);
  return { ok: true, message: '예약이 취소되었습니다.' };
}

/* ───────── 관리자: 시트에서 상태 변경 시 메일 발송 ───────── */

/** 설치형 onEdit 트리거 핸들러 (초기 설정 메뉴에서 자동 등록) */
function handleEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET.RES) return;
  if (COL.STATUS < e.range.getColumn() || COL.STATUS > e.range.getLastColumn()) return;
  const first = Math.max(2, e.range.getRow());
  for (let rowNo = first; rowNo <= e.range.getLastRow(); rowNo++) processStatusRow_(sh, rowNo);
}

function processStatusRow_(sh, rowNo) {
  const r = sh.getRange(rowNo, 1, 1, RES_HEADERS.length).getDisplayValues()[0];
  const status = r[COL.STATUS - 1];
  if (status !== '승인' && status !== '반려') return;
  if (r[COL.MAILED - 1].indexOf(status + ' 안내') === 0) return; // 같은 상태로 이미 발송함
  if (!r[COL.EMAIL - 1]) return;

  const s = readSessions_().filter(function (x) { return x.id === r[COL.SESSION - 1]; })[0];
  const label = s ? sessionLabel_(s) : r[COL.SESSION - 1];
  const cfg = getConfig_();
  const head = r[COL.ORG - 1] + ' ' + r[COL.NAME - 1] + ' 선생님께';
  let subject, lines;

  if (status === '승인') {
    if (s) {
      const u = usage_(readReservations_(), -1)[s.id];
      if (u && u.total > s.capacity) {
        sh.getRange(rowNo, COL.MAILED).setValue('정원 초과(' + u.total + '/' + s.capacity + '), 발송 보류');
        return;
      }
    }
    subject = '[태안 AI-DREAM Lab] 투어 예약이 승인되었습니다 (' + r[COL.NO - 1] + ')';
    lines = [head, '', '태안 AI-DREAM Lab 오픈 랩 데이 투어 예약이 승인되었습니다.', '',
      '- 예약번호: ' + r[COL.NO - 1], '- 일시: ' + label, '- 인원: ' + r[COL.COUNT - 1] + '명',
      '- 장소: ' + (cfg['장소'] || ''), ''];
    if (cfg['오시는 길 안내']) lines.push(cfg['오시는 길 안내']);
    if (cfg['촬영 안내']) lines.push(cfg['촬영 안내']);
    if (cfg['공문 안내']) lines.push('', cfg['공문 안내']);
    if (r[COL.NOTE - 1]) lines.push('', r[COL.NOTE - 1]);
    lines.push('', '일정 확인과 취소는 예약 페이지의 예약 조회·취소에서 하실 수 있습니다.');
    if (cfg['예약 페이지 주소']) lines.push(cfg['예약 페이지 주소']);
  } else {
    subject = '[태안 AI-DREAM Lab] 투어 예약 결과 안내 (' + r[COL.NO - 1] + ')';
    lines = [head, '', '신청하신 투어 예약은 아래 사유로 승인되지 않았습니다.', '',
      '- 예약번호: ' + r[COL.NO - 1], '- 일시: ' + label,
      '- 사유: ' + (r[COL.NOTE - 1] || '학교 일정상 해당 회차 운영이 어렵습니다.'), '',
      '다른 회차는 예약 페이지에서 신청하실 수 있습니다.'];
    if (cfg['예약 페이지 주소']) lines.push(cfg['예약 페이지 주소']);
  }

  const sent = sendMail_(r[COL.EMAIL - 1], subject, lines.join('\n'));
  sh.getRange(rowNo, COL.DONE).setValue(now_());
  sh.getRange(rowNo, COL.MAILED).setValue(sent ? status + ' 안내 발송 ' + now_() : '메일 발송 실패');
}

/* ───────── 관리자 메뉴 ───────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('투어 예약 관리')
    .addItem('1. 초기 설정(최초 1회)', 'setup')
    .addItem('2. 오픈 랩 데이 추가', 'addOpenDay')
    .addItem('3. 관리자 계정 설정', 'setAdminAccount')
    .addSeparator()
    .addItem('통계 갱신', 'updateStats')
    .addItem('개인정보 파기(기준일 이전 방문분)', 'purgePersonalInfo')
    .addItem('개인정보 자동 파기 켜기', 'installPurgeTrigger')
    .addToUi();
}

function ensureSheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#E3F1EE');
    sh.setFrozenRows(1);
  }
  return sh;
}

function setup() {
  const days = ensureSheet_(SHEET.DAYS, DAY_HEADERS);
  days.getRange('A:D').setNumberFormat('@');
  days.getRange('G:G').setNumberFormat('@');
  days.getRange('H2:H500').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(DAY_STATUS, true).setAllowInvalid(false).build());

  const res = ensureSheet_(SHEET.RES, RES_HEADERS);
  res.getRange('A:I').setNumberFormat('@');
  res.getRange('K:R').setNumberFormat('@');
  res.getRange('N2:N2000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(RES_STATUS, true).setAllowInvalid(false).build());

  const cfg = ensureSheet_(SHEET.CONFIG, ['항목', '값']);
  cfg.getRange('A:B').setNumberFormat('@');
  if (cfg.getLastRow() < 2) {
    const rows = CONFIG_DEFAULTS.map(function (r) { return r.slice(); });
    rows[0][1] = Session.getEffectiveUser().getEmail();
    cfg.getRange(2, 1, rows.length, 2).setValues(rows);
    cfg.setColumnWidth(1, 160).setColumnWidth(2, 520);
  }

  ensureSheet_(SHEET.STATS, ['회차ID', '일시', '승인 기관 수', '승인 인원', '정원']);

  const has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'handleEdit'; });
  if (!has) ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss_()).onEdit().create();
  installPurgeTrigger_();

  SpreadsheetApp.getUi().alert('초기 설정을 마쳤습니다.\nConfig 시트의 값을 확인한 뒤 오픈 랩 데이를 추가하세요.');
}

/** 날짜를 입력받아 2개 세션 행을 비공개 상태로 추가 */
function addOpenDay() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('오픈 랩 데이 추가', '날짜를 입력하세요. (예: 2026-10-21)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const r = addOpenDayCore_(res.getResponseText().trim());
  ui.alert(r.message);
}

function addOpenDayCore_(date) {
  if (!isDate_(date)) return { ok: false, message: '날짜는 2026-10-21 형식으로 입력하세요.' };
  const sh = ss_().getSheetByName(SHEET.DAYS);
  const base = 'D' + date.replace(/-/g, '').slice(2);
  const exists = readSessions_().some(function (s) { return s.id.indexOf(base) === 0; });
  if (exists) return { ok: false, message: '이미 등록된 날짜입니다.' };
  DEFAULT_SESSIONS.forEach(function (t, i) {
    sh.appendRow([base + '-' + (i + 1), date, t.start, t.end, DEFAULT_CAPACITY, DEFAULT_ORG_CAP,
      addDays_(date, -DEADLINE_DAYS), '비공개', '']);
  });
  return { ok: true, message: date + ' 회차 2개를 비공개 상태로 추가했습니다. 시간과 정원을 확인한 뒤 상태를 공개로 바꾸세요.' };
}

/** 승인 건 기준 회차별 집계 (개인정보 미포함) */
function updateStats() {
  updateStatsCore_();
  SpreadsheetApp.getUi().alert('통계를 갱신했습니다.');
}

function updateStatsCore_() {
  const rows = readReservations_();
  const sessions = readSessions_();
  const out = sessions.map(function (s) {
    const orgs = {};
    let people = 0;
    rows.forEach(function (r) {
      if (r[COL.SESSION - 1] !== s.id || r[COL.STATUS - 1] !== '승인') return;
      orgs[orgKey_(r[COL.ORG - 1])] = true;
      people += parseInt(r[COL.COUNT - 1], 10) || 0;
    });
    return [s.id, sessionLabel_(s), Object.keys(orgs).length, people, s.capacity];
  });
  const sh = ss_().getSheetByName(SHEET.STATS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, 5).setValues(out);
  return out.length;
}

/** 파기 대상 열: 직위·성명·연락처·이메일·방문 목적·안내 메모·내부 메모 (기관명·기관 유형·인원은 통계용으로 유지) */
const PURGE_COLS = [COL.TITLE, COL.NAME, COL.PHONE, COL.EMAIL, COL.PURPOSE, COL.NOTE, COL.MEMO];

/** 방문일: OpenDays에서 찾고, 회차 행이 지워졌으면 회차ID(D261021-1)에서 복원 */
function visitDateOf_(sessionId, dateOf) {
  if (dateOf[sessionId]) return dateOf[sessionId];
  const m = /^D(\d{2})(\d{2})(\d{2})-/.exec(String(sessionId || ''));
  return m ? '20' + m[1] + '-' + m[2] + '-' + m[3] : '';
}

/** cutoff(yyyy-MM-dd) 이전 방문분의 개인정보를 지우고 건수를 반환. 호출 측에서 스크립트 락을 잡고 있어야 함 */
function purgeBefore_(cutoff) {
  const dateOf = {};
  readSessions_().forEach(function (s) { dateOf[s.id] = s.date; });
  const sh = ss_().getSheetByName(SHEET.RES);
  let n = 0;
  readReservations_().forEach(function (r, i) {
    const d = visitDateOf_(r[COL.SESSION - 1], dateOf);
    if (!d || d >= cutoff) return;
    const left = PURGE_COLS.some(function (c) { return r[c - 1] !== ''; });
    if (!left) return;
    PURGE_COLS.forEach(function (c) { sh.getRange(i + 2, c).clearContent(); });
    n++;
  });
  return n;
}

/** 메뉴: Config의 파기 기준일 이전 방문분을 수동 파기 */
function purgePersonalInfo() {
  const ui = SpreadsheetApp.getUi();
  const cutoff = getConfig_()['파기 기준일'] || '';
  if (!isDate_(cutoff)) { ui.alert('Config 시트의 파기 기준일을 2027-03-01 형식으로 입력하세요.'); return; }
  const ok = ui.alert('개인정보 파기', cutoff + ' 이전 방문분의 직위, 성명, 연락처, 이메일, 방문 목적, 메모를 삭제합니다. 되돌릴 수 없습니다.', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let n;
  try { n = purgeBefore_(cutoff); } finally { lock.releaseLock(); }
  ui.alert(n + '건의 개인정보를 파기했습니다.');
}

/** 자동 파기: 매일 새벽 실행. 지난 학년도(3월 1일 이전) 방문분을 파기하므로 사실상 3월 1일에 처리됨 */
function autoPurge() {
  const now = new Date();
  const y = parseInt(Utilities.formatDate(now, TZ, 'yyyy'), 10);
  const m = parseInt(Utilities.formatDate(now, TZ, 'M'), 10);
  const cutoff = (m >= 3 ? y : y - 1) + '-03-01';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  let n;
  try { n = purgeBefore_(cutoff); } finally { lock.releaseLock(); }
  if (n > 0) {
    sendMail_(getConfig_()['관리자 메일'], '[투어 예약] 개인정보 자동 파기 ' + n + '건',
      cutoff + ' 이전 방문분 ' + n + '건의 직위, 성명, 연락처, 이메일, 방문 목적, 메모를 삭제했습니다.\n' +
      '기관명, 기관 유형, 인원은 통계용으로 남겨 두었습니다.\n\n' +
      '이 시트의 사본이나 내려받은 파일, 메일함의 신청 알림 메일은 따로 삭제해야 합니다.\n' + ss_().getUrl());
  }
}

function installPurgeTrigger_() {
  const has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'autoPurge'; });
  if (!has) ScriptApp.newTrigger('autoPurge').timeBased().everyDays(1).atHour(3).inTimezone(TZ).create();
  return !has;
}

/** 메뉴: 자동 파기 트리거 등록 */
function installPurgeTrigger() {
  const added = installPurgeTrigger_();
  SpreadsheetApp.getUi().alert(added
    ? '자동 파기를 켰습니다. 매일 새벽 3시 무렵 점검하며, 3월 1일부터 지난 학년도 방문분을 파기합니다.'
    : '자동 파기가 이미 켜져 있습니다.');
}

/* ───────── 관리자 페이지 API (admin.html) ─────────
 * 인증: 시트 메뉴 3에서 설정한 아이디·비밀번호. 비밀번호는 솔트+SHA-256 해시로
 * 스크립트 속성에만 저장(시트에 남지 않음). 로그인 성공 시 6시간짜리 토큰 발급.
 * 실패 5회면 15분 잠금. 모든 admin.* 요청은 토큰이 있어야 처리.
 */
const ADMIN_TOKEN_TTL = 6 * 3600;
const ADMIN_MAX_FAIL = 5;
const ADMIN_LOCK_SEC = 900;

function hash_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}
function randomHex_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); }

/** 메뉴 3: 관리자 계정 설정 */
function setAdminAccount() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('관리자 계정 설정 (1/2)', '관리자 아이디 (영문·숫자 4~30자)', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const id = r1.getResponseText().trim();
  if (!/^[A-Za-z0-9_.-]{4,30}$/.test(id)) { ui.alert('아이디는 영문·숫자 4~30자로 입력하세요.'); return; }
  const r2 = ui.prompt('관리자 계정 설정 (2/2)', '비밀번호 (10자 이상, 영문과 숫자 포함)', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const pw = r2.getResponseText();
  if (pw.length < 10 || !/[A-Za-z]/.test(pw) || !/\d/.test(pw)) { ui.alert('비밀번호는 10자 이상, 영문과 숫자를 포함해야 합니다.'); return; }
  const salt = randomHex_().slice(0, 16);
  PropertiesService.getScriptProperties().setProperties({ ADMIN_ID: id, ADMIN_SALT: salt, ADMIN_HASH: hash_(salt + pw) });
  ui.alert('관리자 계정을 저장했습니다.\n비밀번호는 해시로만 보관되며 시트에는 남지 않습니다.\n기존 로그인 토큰은 그대로 유효합니다(최대 6시간).');
}

function adminLogin_(id, pw) {
  id = String(id || '').trim(); pw = String(pw || '');
  const cache = CacheService.getScriptCache();
  const failKey = 'alogin:' + id.toLowerCase();
  const fails = parseInt(cache.get(failKey), 10) || 0;
  if (fails >= ADMIN_MAX_FAIL) return { ok: false, message: '로그인 실패가 반복되어 15분간 잠겼습니다.' };
  const p = PropertiesService.getScriptProperties();
  const okId = p.getProperty('ADMIN_ID'), salt = p.getProperty('ADMIN_SALT'), h = p.getProperty('ADMIN_HASH');
  if (!okId || !salt || !h) return { ok: false, message: '관리자 계정이 아직 설정되지 않았습니다. 시트 메뉴 3에서 설정하세요.' };
  if (id !== okId || hash_(salt + pw) !== h) {
    cache.put(failKey, String(fails + 1), ADMIN_LOCK_SEC);
    return { ok: false, message: '아이디 또는 비밀번호가 올바르지 않습니다. (' + (fails + 1) + '/' + ADMIN_MAX_FAIL + ')' };
  }
  cache.remove(failKey);
  const token = randomHex_();
  cache.put('atok:' + token, id, ADMIN_TOKEN_TTL);
  return { ok: true, token: token, id: id };
}

function adminAuth_(token) {
  token = String(token || '');
  if (!/^[0-9a-f]{64}$/.test(token)) return false;
  const cache = CacheService.getScriptCache();
  const id = cache.get('atok:' + token);
  if (!id) return false;
  cache.put('atok:' + token, id, ADMIN_TOKEN_TTL); // 사용할 때마다 연장
  return true;
}

function adminApi_(action, body) {
  if (action === 'login') return adminLogin_(body.id, body.pw);
  if (!adminAuth_(body.token)) return { ok: false, auth: false, message: '로그인이 필요합니다.' };
  if (action === 'logout') { CacheService.getScriptCache().remove('atok:' + body.token); return { ok: true }; }
  if (action === 'overview') return adminOverview_();
  if (action === 'setStatus') return adminSetStatus_(body.no, body.status, body.note);
  if (action === 'setDoc') return adminSetDoc_(body.no, body.received === true);
  if (action === 'setDayStatus') return adminSetDayStatus_(body.id, body.status);
  if (action === 'addDay') { const r = addOpenDayCore_(String(body.date || '').trim()); return r; }
  if (action === 'updateStats') return { ok: true, message: updateStatsCore_() + '개 회차 통계를 갱신했습니다.' };
  return { ok: false, message: '알 수 없는 관리자 요청입니다.' };
}

/** 회차 요약 + 예약 전체(개인정보 포함, 관리자 전용) */
function adminOverview_() {
  const rows = readReservations_();
  const use = usage_(rows, -1);
  const sessions = readSessions_().map(function (s) {
    let pending = 0, approved = 0, orgs = {};
    rows.forEach(function (r) {
      if (r[COL.SESSION - 1] !== s.id) return;
      if (r[COL.STATUS - 1] === '신청') pending += parseInt(r[COL.COUNT - 1], 10) || 0;
      if (r[COL.STATUS - 1] === '승인') { approved += parseInt(r[COL.COUNT - 1], 10) || 0; orgs[orgKey_(r[COL.ORG - 1])] = true; }
    });
    const taken = use[s.id] ? use[s.id].total : 0;
    return {
      id: s.id, date: s.date, weekday: s.weekday, start: s.start, end: s.end, status: s.status,
      capacity: s.capacity, orgCap: s.orgCap, deadline: s.deadline,
      pending: pending, approved: approved, approvedOrgs: Object.keys(orgs).length,
      remaining: Math.max(0, s.capacity - taken)
    };
  }).sort(function (a, b) { return (b.date + b.start).localeCompare(a.date + a.start); });
  const list = rows.map(function (r, i) {
    return {
      row: i + 2, no: r[COL.NO - 1], applied: r[COL.APPLIED - 1], session: r[COL.SESSION - 1],
      org: r[COL.ORG - 1], orgType: r[COL.ORGTYPE - 1], title: r[COL.TITLE - 1], name: r[COL.NAME - 1],
      phone: r[COL.PHONE - 1], email: r[COL.EMAIL - 1], count: parseInt(r[COL.COUNT - 1], 10) || 0,
      purpose: r[COL.PURPOSE - 1], doc: r[COL.DOC - 1], status: r[COL.STATUS - 1], done: r[COL.DONE - 1],
      mailed: r[COL.MAILED - 1], note: r[COL.NOTE - 1], memo: r[COL.MEMO - 1]
    };
  }).filter(function (x) { return x.no; });
  return { ok: true, today: today_(), sessions: sessions, reservations: list, statuses: RES_STATUS, dayStatuses: DAY_STATUS };
}

function findRowByNo_(no) {
  no = String(no || '').trim().toUpperCase();
  const rows = readReservations_();
  for (let i = 0; i < rows.length; i++) if (rows[i][COL.NO - 1] === no) return { rowNo: i + 2, row: rows[i] };
  return null;
}

/** 승인·반려·취소 처리. 승인·반려는 processStatusRow_ 가 메일까지 처리, 취소는 여기서 발송 */
function adminSetStatus_(no, status, note) {
  status = String(status || '');
  if (['승인', '반려', '취소', '신청'].indexOf(status) < 0) return { ok: false, message: '잘못된 상태입니다.' };
  note = clean_(note, 300);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (err) { return { ok: false, message: '잠시 후 다시 시도하세요.' }; }
  let result;
  try {
    const hit = findRowByNo_(no);
    if (!hit) return { ok: false, message: '예약을 찾지 못했습니다.' };
    const sh = ss_().getSheetByName(SHEET.RES);
    const r = hit.row, rowNo = hit.rowNo;
    if (status === '반려' && !note && !r[COL.NOTE - 1]) return { ok: false, message: '반려 사유를 입력하세요.' };
    if (r[COL.STATUS - 1] === status && (status === '신청' || r[COL.MAILED - 1].indexOf(status + ' 안내') === 0)) {
      return { ok: false, message: '이미 ' + status + ' 상태입니다.' };
    }
    // 승인 또는 승인 대기로 바꾸기 전에 정원 확인: 이 예약을 뺀 점유 인원 + 이 예약 인원
    if (ACTIVE.indexOf(status) >= 0) {
      const s = readSessions_().filter(function (x) { return x.id === r[COL.SESSION - 1]; })[0];
      if (s) {
        const others = usage_(readReservations_(), rowNo - 2)[s.id];
        const taken = others ? others.total : 0;
        const count = parseInt(r[COL.COUNT - 1], 10) || 0;
        if (taken + count > s.capacity) {
          return { ok: false, message: '정원을 넘어 처리하지 않았습니다. (다른 예약 ' + taken + '명 + 이 예약 ' + count +
            '명 > 정원 ' + s.capacity + '명) 상태는 그대로입니다.' };
        }
      }
    }
    if (note) sh.getRange(rowNo, COL.NOTE).setValue(note);
    sh.getRange(rowNo, COL.STATUS).setValue(status);
    if (status === '신청') {
      sh.getRange(rowNo, COL.DONE).setValue('');
      sh.getRange(rowNo, COL.MAILED).setValue('');
      result = { ok: true, message: '승인 대기로 되돌렸습니다. 메일은 발송되지 않았습니다.' };
    } else if (status === '취소') {
      sh.getRange(rowNo, COL.DONE).setValue(now_());
      const s = readSessions_().filter(function (x) { return x.id === r[COL.SESSION - 1]; })[0];
      const label = s ? sessionLabel_(s) : r[COL.SESSION - 1];
      const sent = sendMail_(r[COL.EMAIL - 1], '[태안 AI-DREAM Lab] 투어 예약이 취소되었습니다 (' + r[COL.NO - 1] + ')',
        [r[COL.ORG - 1] + ' ' + r[COL.NAME - 1] + ' 선생님께', '', '아래 예약이 학교 사정으로 취소되었습니다.', '',
          '- 예약번호: ' + r[COL.NO - 1], '- 일시: ' + label, note ? '- 안내: ' + note : '', '',
          '문의: ' + (getConfig_()['학교명'] || '태안초등학교') + ' 정보과학부'].join('\n'));
      sh.getRange(rowNo, COL.MAILED).setValue(sent ? '취소 안내 발송 ' + now_() : '관리자 취소(메일 발송 실패)');
      result = { ok: true, message: sent ? '취소 처리하고 안내 메일을 보냈습니다.' : '취소 처리했지만 메일 발송에 실패했습니다.' };
    } else {
      processStatusRow_(sh, rowNo);
      const mailed = sh.getRange(rowNo, COL.MAILED).getDisplayValue();
      const held = mailed.indexOf('보류') >= 0 || mailed.indexOf('실패') >= 0;
      result = { ok: !held, message: held ? mailed : status + ' 처리하고 안내 메일을 보냈습니다.', mailed: mailed };
    }
  } finally {
    lock.releaseLock();
  }
  return result;
}

function adminSetDoc_(no, received) {
  const hit = findRowByNo_(no);
  if (!hit) return { ok: false, message: '예약을 찾지 못했습니다.' };
  ss_().getSheetByName(SHEET.RES).getRange(hit.rowNo, COL.DOC).setValue(received ? '접수 ' + today_() : '');
  return { ok: true };
}

function adminSetDayStatus_(id, status) {
  if (DAY_STATUS.indexOf(status) < 0) return { ok: false, message: '잘못된 상태입니다.' };
  const sh = ss_().getSheetByName(SHEET.DAYS);
  const rows = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0].trim() === id) { sh.getRange(i + 2, 8).setValue(status); return { ok: true }; }
  }
  return { ok: false, message: '회차를 찾지 못했습니다.' };
}
