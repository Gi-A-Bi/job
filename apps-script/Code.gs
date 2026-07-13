/**
 * 우리 반 직업 배정소 — 백엔드 (Google Apps Script)
 * ===================================================
 * 이 스크립트는 Google 스프레드시트에서
 *   [확장 프로그램] → [Apps Script]
 * 로 열어 만든 "바인딩 스크립트"입니다.
 * getSS_()가 자동으로 이 스프레드시트를 가리킵니다. (스프레드시트 ID 불필요)
 *
 * ── 회차(round) 모델 ──
 * 모든 신청/배정은 '회차' 단위로 묶입니다.
 *   상태(status): idle(열리지 않음) → open(신청받는중) → closed(마감/공개)
 *   - 교사가 '회차 열기'  → 현재 회차 +1, 상태 open  (학생 신청 가능)
 *   - 교사가 '자동 배정'  → 현재 회차 배정 계산
 *   - 교사가 '배정 마감'  → 상태 closed (학생이 결과 확인 가능)
 */

// ===== 시트 이름 =====
const SHEET_CLASSES      = 'Classes';
const SHEET_JOBS         = 'Jobs';
const SHEET_STUDENTS     = 'Students';
const SHEET_APPLICATIONS = 'Applications';
const SHEET_ASSIGNMENTS  = 'Assignments';
const SHEET_ROUNDS       = 'Rounds';

// ===== 시트 헤더 =====
const HEADERS = {
  'Classes':      ['classCode', 'className', 'passwordHash', 'currentRound', 'status', 'createdAt'],
  'Jobs':         ['classCode', 'jobId', 'jobName', 'icon', 'capacity', 'isPopular', 'sortOrder', 'description'],
  'Students':     ['classCode', 'studentId', 'number', 'name', 'createdAt', 'diligence'],
  'Applications': ['classCode', 'round', 'studentId', 'number', 'studentName', 'choice1', 'choice2', 'choice3', 'reason', 'submittedAt'],
  'Assignments':  ['classCode', 'round', 'studentId', 'number', 'studentName', 'jobId', 'jobName', 'choiceLevel', 'diligence', 'assignedAt'],
  'Rounds':       ['classCode', 'round', 'status', 'createdAt'],
};

// ===== 회차 상태 =====
const ST_OPEN = 'open';           // 신청 받는 중
const ST_PUBLISHED = 'published'; // 배정 공개됨

// ===== 보안 =====
const PW_SALT_PREFIX = 'classroom-jobs-2026:';

// ===== 배정 가중치 =====
const DILIGENCE_POINTS = { '잘함': 8, '보통': 4, '아쉬움': 0 };
const BONUS_NEW_POPULAR = 3;   // 인기직 한 번도 못 해본 학생
const BONUS_REASON      = 2;   // 지원 이유 5자 이상
const RANDOM_MAX        = 3;   // 추첨 0~3

// ===== 학급 코드 문자셋 (0/O, 1/I 등 헷갈리는 문자 제외) =====
const CODE_CHARS = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';


// ============================================================
// 엔트리 포인트
// ============================================================

function doGet() {
  ensureSheets_();
  const t = HtmlService.createTemplateFromFile('index');
  t.backendUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle('우리 반 직업 배정소')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let res;
  try {
    ensureSheets_();
    const body = JSON.parse(e.postData.contents);
    res = route_(body.action, body.payload || {});
  } catch (err) {
    res = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

// 데이터를 바꾸는 요청은 잠금(Lock) 안에서 실행 — 반 전체가 동시에 제출해도 안전
const MUTATING_ACTIONS = {
  createClass: 1, saveJobs: 1, saveStudents: 1, setStudentDiligence: 1, setDiligence: 1,
  openNextRound: 1, publishRound: 1, reopenRound: 1, submitApplication: 1,
  deleteApplication: 1, runAssignment: 1,
};

function route_(action, p) {
  if (MUTATING_ACTIONS[action]) {
    return withLock_(function () { return dispatch_(action, p); });
  }
  return dispatch_(action, p);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    throw new Error('지금 접속하는 친구가 많아요. 몇 초 뒤에 다시 시도해 주세요.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function dispatch_(action, p) {
  switch (action) {
    case 'createClass':         return createClass_(p);
    case 'teacherLogin':        return teacherLogin_(p);
    case 'getClassInfo':        return getClassInfo_(p);
    case 'listJobs':            return listJobs_(p);
    case 'saveJobs':            return saveJobs_(p);
    case 'listStudents':        return listStudents_(p);
    case 'saveStudents':        return saveStudents_(p);
    case 'setStudentDiligence': return setStudentDiligence_(p);
    case 'setDiligence':        return setStudentDiligence_(p);
    case 'openNextRound':       return openNextRound_(p);
    case 'publishRound':        return publishRound_(p);
    case 'reopenRound':         return reopenRound_(p);
    case 'submitApplication':   return submitApplication_(p);
    case 'getMyApplication':    return getMyApplication_(p);
    case 'getMyAppliedRounds':  return getMyAppliedRounds_(p);
    case 'listApplications':    return listApplications_(p);
    case 'deleteApplication':   return deleteApplication_(p);
    case 'runAssignment':       return runAssignment_(p);
    case 'listAssignments':     return listAssignments_(p);
    case 'getMyAssignment':     return getMyAssignment_(p);
    case 'getMyResults':        return getMyResults_(p);
    case 'getRounds':           return getRounds_(p);
    default: throw new Error('알 수 없는 요청입니다: ' + action);
  }
}


// ============================================================
// 학급 / 인증
// ============================================================

function createClass_(p) {
  const className = String(p.className || '').trim();
  const password  = String(p.password || '').trim();
  if (!className) throw new Error('학급 이름을 입력해 주세요.');
  if (password.length < 4) throw new Error('관리 비밀번호는 4자 이상으로 정해 주세요.');

  const classCode = makeUniqueClassCode_();
  getSheet_(SHEET_CLASSES).appendRow([
    classCode, className, hashPassword_(classCode, password), 1, ST_OPEN, new Date(),
  ]);
  setRoundStatus_(classCode, 1, ST_OPEN);   // 1회차 자동 열림
  return { ok: true, classCode: classCode, className: className, currentRound: 1, openRounds: [1] };
}

function teacherLogin_(p) {
  const cls = requireTeacher_(p);
  ensureRoundsBackfill_(cls);
  return {
    ok: true, className: cls.className, currentRound: cls.currentRound,
    rounds: roundsOf_(cls.classCode),
  };
}

function getRounds_(p) {
  const cls = requireTeacher_(p);
  ensureRoundsBackfill_(cls);
  return { ok: true, currentRound: cls.currentRound, rounds: roundsOf_(cls.classCode) };
}

function getClassInfo_(p) {
  const cls = getClassRow_(p.classCode);
  if (!cls) throw new Error('학급을 찾을 수 없어요. 코드를 다시 확인해 주세요.');
  ensureRoundsBackfill_(cls);
  return {
    ok: true, className: cls.className, currentRound: cls.currentRound,
    openRounds: openRoundsOf_(cls.classCode),
    publishedRounds: publishedRoundsOf_(cls.classCode),
  };
}

// 다음 회차 열기: 새 회차를 하나 열어 신청 받기 (여러 번 호출해 미리 여러 회차 개방 가능)
function openNextRound_(p) {
  const cls = requireTeacher_(p);
  ensureRoundsBackfill_(cls);
  const next = (cls.currentRound || 0) + 1;
  setClassFields_(cls.classCode, { currentRound: next });
  setRoundStatus_(cls.classCode, next, ST_OPEN);
  return { ok: true, currentRound: next, rounds: roundsOf_(cls.classCode) };
}

// 배정 마감(공개): 해당 회차 결과를 학생에게 공개
function publishRound_(p) {
  const cls = requireTeacher_(p);
  const round = Number(p.round) || cls.currentRound;
  if (!roundStatusOf_(cls.classCode, round)) throw new Error('열려 있지 않은 회차예요.');
  const has = readAll_(SHEET_ASSIGNMENTS).some(function (a) {
    return a.classCode === cls.classCode && Number(a.round) === round;
  });
  if (!has) throw new Error('먼저 자동 배정을 실행해 주세요.');
  setRoundStatus_(cls.classCode, round, ST_PUBLISHED);
  return { ok: true, round: round, rounds: roundsOf_(cls.classCode) };
}

// 공개 취소: 해당 회차를 다시 신청/배정 가능 상태로
function reopenRound_(p) {
  const cls = requireTeacher_(p);
  const round = Number(p.round) || cls.currentRound;
  setRoundStatus_(cls.classCode, round, ST_OPEN);
  return { ok: true, round: round, rounds: roundsOf_(cls.classCode) };
}


// ============================================================
// 직업
// ============================================================

function listJobs_(p) {
  const cls = getClassRow_(p.classCode);
  if (!cls) throw new Error('학급을 찾을 수 없어요.');
  return { ok: true, jobs: jobsOf_(p.classCode) };
}

function saveJobs_(p) {
  const cls = requireTeacher_(p);
  const incoming = Array.isArray(p.jobs) ? p.jobs : [];
  deleteRowsWhere_(SHEET_JOBS, function (r) { return r.classCode === cls.classCode; });
  const sheet = getSheet_(SHEET_JOBS);
  incoming.forEach(function (j, i) {
    const jobId = String(j.jobId || ('job_' + Utilities.getUuid().slice(0, 8)));
    sheet.appendRow([
      cls.classCode, jobId, String(j.jobName || '').trim(), String(j.icon || ''),
      Number(j.capacity) > 0 ? Number(j.capacity) : 1, truthy_(j.isPopular) ? 'Y' : '', i,
      String(j.description || '').trim(),
    ]);
  });
  return { ok: true, jobs: jobsOf_(cls.classCode) };
}


// ============================================================
// 학생
// ============================================================

function listStudents_(p) {
  const cls = requireTeacher_(p);
  return { ok: true, students: studentsOf_(cls.classCode) };
}

function saveStudents_(p) {
  const cls = requireTeacher_(p);
  const incoming = Array.isArray(p.students) ? p.students : [];

  const prevDil = {};
  readAll_(SHEET_STUDENTS).forEach(function (r) {
    if (r.classCode === cls.classCode) prevDil[r.studentId] = r.diligence || '';
  });

  deleteRowsWhere_(SHEET_STUDENTS, function (r) { return r.classCode === cls.classCode; });

  const sheet = getSheet_(SHEET_STUDENTS);
  incoming.forEach(function (s) {
    const number = String(s.number || '').trim();
    const name   = String(s.name || '').trim();
    if (!name) return;
    const studentId = 'stu_' + cls.classCode + '_' + number + '_' + name;
    const dil = s.diligence || prevDil[studentId] || '';
    sheet.appendRow([cls.classCode, studentId, number, name, new Date(), dil]);
  });
  return { ok: true, students: studentsOf_(cls.classCode) };
}

function setStudentDiligence_(p) {
  const cls = requireTeacher_(p);
  const items = Array.isArray(p.items) ? p.items : [];
  const map = {};
  items.forEach(function (it) { map[it.studentId] = it.diligence; });

  const sheet = getSheet_(SHEET_STUDENTS);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idxClass = header.indexOf('classCode');
  const idxStu   = header.indexOf('studentId');
  const idxDil   = header.indexOf('diligence');

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (row[idxClass] === cls.classCode && (row[idxStu] in map)) {
      sheet.getRange(r + 1, idxDil + 1).setValue(map[row[idxStu]]);
    }
  }
  return { ok: true, students: studentsOf_(cls.classCode) };
}


// ============================================================
// 지원 (회차 단위)
// ============================================================

function submitApplication_(p) {
  const cls = getClassRow_(p.classCode);
  if (!cls) throw new Error('학급을 찾을 수 없어요.');
  ensureRoundsBackfill_(cls);
  const round = Number(p.round) || cls.currentRound;
  if (roundStatusOf_(cls.classCode, round) !== ST_OPEN) {
    throw new Error('지금은 그 회차에 신청할 수 없어요. 열려 있는 회차를 골라 주세요.');
  }

  const number = String(p.number || '').trim();
  const name   = String(p.name || '').trim();
  if (!name) throw new Error('이름을 입력해 주세요.');

  const students = studentsOf_(cls.classCode);
  const me = students.find(function (s) {
    return sameNumber_(s.number, number) && s.name === name;
  });
  if (!me) throw new Error('명단에서 찾을 수 없어요. 번호와 이름을 다시 확인해 주세요.');

  // 이미 이 회차에 신청했으면 다시 신청 불가 (덮어쓰기 방지)
  // — 실수로 잘못 냈다면 선생님이 지원현황에서 삭제한 뒤 다시 신청하게 함
  var already = readAll_(SHEET_APPLICATIONS).some(function (rr) {
    return rr.classCode === cls.classCode && Number(rr.round) === round && rr.studentId === me.studentId;
  });
  if (already) {
    throw new Error('이미 이 회차에 신청했어요. 한 번 신청하면 다시 신청할 수 없어요. (수정이 필요하면 선생님께 말씀해 주세요.)');
  }

  getSheet_(SHEET_APPLICATIONS).appendRow([
    cls.classCode, round, me.studentId, number, name,
    String(p.choice1 || ''), String(p.choice2 || ''), String(p.choice3 || ''),
    String(p.reason || '').trim(), new Date(),
  ]);
  return { ok: true };
}

// 학생이 이미 신청한 회차 번호 목록 (신청 완료 표시용)
function getMyAppliedRounds_(p) {
  const cls = getClassRow_(p.classCode);
  if (!cls) throw new Error('학급을 찾을 수 없어요.');
  const number = String(p.number || '').trim();
  const name   = String(p.name || '').trim();
  const rounds = readAll_(SHEET_APPLICATIONS).filter(function (a) {
    return a.classCode === cls.classCode && validRound_(a.round) &&
           sameNumber_(a.number, number) && a.studentName === name;
  }).map(function (a) { return Number(a.round); });
  return { ok: true, rounds: rounds };
}

function getMyApplication_(p) {
  const cls = getClassRow_(p.classCode);
  if (!cls) throw new Error('학급을 찾을 수 없어요.');
  const number = String(p.number || '').trim();
  const name   = String(p.name || '').trim();
  const round  = Number(p.round) || cls.currentRound;
  const mine = readAll_(SHEET_APPLICATIONS).find(function (a) {
    return a.classCode === cls.classCode && Number(a.round) === round &&
           sameNumber_(a.number, number) && a.studentName === name;
  });
  return { ok: true, application: mine || null };
}

function listApplications_(p) {
  const cls = requireTeacher_(p);
  const round = Number(p.round) || cls.currentRound;
  const apps = readAll_(SHEET_APPLICATIONS).filter(function (a) {
    return a.classCode === cls.classCode && Number(a.round) === round;
  });
  return { ok: true, round: round, applications: apps };
}

function deleteApplication_(p) {
  const cls = requireTeacher_(p);
  const round = Number(p.round) || cls.currentRound;
  deleteRowsWhere_(SHEET_APPLICATIONS, function (r) {
    return r.classCode === cls.classCode && Number(r.round) === round && r.studentId === p.studentId;
  });
  return { ok: true };
}


// ============================================================
// 배정
// ============================================================

function runAssignment_(p) {
  const cls = requireTeacher_(p);
  const round = Number(p.round) || cls.currentRound;
  if (roundStatusOf_(cls.classCode, round) !== ST_OPEN) {
    throw new Error('지금은 배정을 실행할 수 없어요. (공개된 회차는 공개 취소 후 다시 시도)');
  }

  const jobs     = jobsOf_(cls.classCode);
  const students = studentsOf_(cls.classCode);
  const apps     = readAll_(SHEET_APPLICATIONS).filter(function (a) {
    return a.classCode === cls.classCode && Number(a.round) === round;
  });
  const history  = readAll_(SHEET_ASSIGNMENTS).filter(function (a) {
    return a.classCode === cls.classCode;
  });

  if (jobs.length === 0) throw new Error('직업을 먼저 등록해 주세요.');

  // 성실도 맵: 교사가 학생마다 설정한 값. 미설정은 '보통' 취급
  const diligenceMap = {};
  students.forEach(function (s) { diligenceMap[s.studentId] = s.diligence || ''; });

  // 인기직 미경험 판정용: 지금까지(이전 회차 포함) 인기직을 맡았던 학생
  const popularJobIds = {};
  jobs.forEach(function (j) { if (truthy_(j.isPopular)) popularJobIds[j.jobId] = true; });
  const heldPopular = {};
  history.forEach(function (a) { if (popularJobIds[a.jobId]) heldPopular[a.studentId] = true; });

  const jobById = {};
  const remaining = {};
  jobs.forEach(function (j) {
    jobById[j.jobId] = j;
    remaining[j.jobId] = Number(j.capacity) > 0 ? Number(j.capacity) : 1;
  });

  const appByStudent = {};
  apps.forEach(function (a) { appByStudent[a.studentId] = a; });

  const assignedJob = {};
  const result = [];

  function scoreOf(studentId, app) {
    var s = 0;
    var dp = DILIGENCE_POINTS[diligenceMap[studentId]];
    s += (dp === undefined ? 4 : dp);   // 미설정은 '보통'(4점)
    if (!heldPopular[studentId]) s += BONUS_NEW_POPULAR;
    if (app && String(app.reason || '').trim().length >= 5) s += BONUS_REASON;
    s += Math.random() * RANDOM_MAX;
    return s;
  }

  [['choice1', 1], ['choice2', 2], ['choice3', 3]].forEach(function (pair) {
    var field = pair[0], level = pair[1];
    var byJob = {};
    apps.forEach(function (app) {
      if (assignedJob[app.studentId]) return;
      var jobId = app[field];
      if (!jobId || !(jobId in remaining) || remaining[jobId] <= 0) return;
      (byJob[jobId] = byJob[jobId] || []).push(app);
    });
    Object.keys(byJob).forEach(function (jobId) {
      var applicants = byJob[jobId];
      var cap = remaining[jobId];
      var winners;
      if (applicants.length <= cap) {
        winners = applicants;
      } else {
        winners = applicants.map(function (app) {
          return { app: app, score: scoreOf(app.studentId, app) };
        }).sort(function (a, b) { return b.score - a.score; })
          .slice(0, cap).map(function (x) { return x.app; });
      }
      winners.forEach(function (app) {
        assignedJob[app.studentId] = jobId;
        remaining[jobId] -= 1;
        result.push(makeAssignRow_(cls.classCode, round, {
          studentId: app.studentId, number: app.number, studentName: app.studentName,
        }, jobId, jobById[jobId], level));
      });
    });
  });

  // 남은 학생 → 남은 직업에 무작위 배치
  var unassigned = students.filter(function (s) { return !assignedJob[s.studentId]; });
  shuffle_(unassigned);
  var openSlots = [];
  jobs.forEach(function (j) {
    for (var i = 0; i < remaining[j.jobId]; i++) openSlots.push(j);
  });
  shuffle_(openSlots);
  unassigned.forEach(function (s, idx) {
    var j = openSlots[idx];
    if (!j) return;
    assignedJob[s.studentId] = j.jobId;
    var app = appByStudent[s.studentId];
    result.push(makeAssignRow_(cls.classCode, round, {
      studentId: s.studentId, number: s.number,
      studentName: (app && app.studentName) || s.name,
    }, j.jobId, j, 0));
  });

  // 저장: 이번 회차 기존 배정 지우고 새로 씀
  deleteRowsWhere_(SHEET_ASSIGNMENTS, function (r) {
    return r.classCode === cls.classCode && Number(r.round) === round;
  });
  var sheet = getSheet_(SHEET_ASSIGNMENTS);
  result.forEach(function (r) {
    sheet.appendRow([
      r.classCode, r.round, r.studentId, r.number, r.studentName,
      r.jobId, r.jobName, r.choiceLevel, r.diligence, r.assignedAt,
    ]);
  });

  var leftover = unassigned.length - openSlots.length;
  return { ok: true, round: round, assignments: result, leftover: leftover > 0 ? leftover : 0 };
}

function makeAssignRow_(classCode, round, stu, jobId, job, level) {
  return {
    classCode: classCode, round: round,
    studentId: stu.studentId, number: stu.number, studentName: stu.studentName,
    jobId: jobId, jobName: job ? job.jobName : '', choiceLevel: level,
    diligence: '', assignedAt: new Date(),
  };
}

function listAssignments_(p) {
  const cls = requireTeacher_(p);
  const round = Number(p.round) || cls.currentRound;
  const rows = readAll_(SHEET_ASSIGNMENTS).filter(function (a) {
    return a.classCode === cls.classCode && Number(a.round) === round;
  });
  const st = roundStatusOf_(cls.classCode, round);
  return {
    ok: true, round: round, currentRound: cls.currentRound,
    roundStatus: st || ST_OPEN, published: st === ST_PUBLISHED,
    rounds: roundsOf_(cls.classCode), assignments: rows,
  };
}

// 학생: 마감(closed)된 경우에만 자기 직업을 볼 수 있음
function getMyAssignment_(p) {
  const cls = getClassRow_(p.classCode);
  if (!cls) throw new Error('학급을 찾을 수 없어요.');
  const round = Number(p.round) || cls.currentRound;
  if (roundStatusOf_(cls.classCode, round) !== ST_PUBLISHED) {
    return { ok: true, assignment: null, published: false, round: round };
  }
  const number = String(p.number || '').trim();
  const name   = String(p.name || '').trim();
  const mine = readAll_(SHEET_ASSIGNMENTS).find(function (a) {
    return a.classCode === cls.classCode && Number(a.round) === round &&
           sameNumber_(a.number, number) && a.studentName === name;
  });
  return { ok: true, assignment: mine || null, published: true, round: round };
}

// 학생: 공개된 회차들에서 내가 받은 직업 모두
function getMyResults_(p) {
  const cls = getClassRow_(p.classCode);
  if (!cls) throw new Error('학급을 찾을 수 없어요.');
  const number = String(p.number || '').trim();
  const name   = String(p.name || '').trim();
  const pubRounds = {};
  publishedRoundsOf_(cls.classCode).forEach(function (r) { pubRounds[r] = true; });
  const mine = readAll_(SHEET_ASSIGNMENTS).filter(function (a) {
    return a.classCode === cls.classCode && pubRounds[Number(a.round)] &&
           sameNumber_(a.number, number) && a.studentName === name;
  }).map(function (a) {
    return { round: Number(a.round), jobName: a.jobName, choiceLevel: a.choiceLevel };
  }).sort(function (x, y) { return y.round - x.round; });
  return { ok: true, results: mine, openRounds: openRoundsOf_(cls.classCode) };
}


// ============================================================
// 인증 / 클래스 헬퍼
// ============================================================

function requireTeacher_(p) {
  const cls = getClassRow_(p.classCode);
  if (!cls) throw new Error('학급을 찾을 수 없어요.');
  const given = hashPassword_(cls.classCode, String(p.password || '').trim());
  if (given !== cls.passwordHash) throw new Error('비밀번호가 올바르지 않아요.');
  return cls;
}

function hashPassword_(classCode, password) {
  const raw = PW_SALT_PREFIX + classCode + ':' + password;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function makeUniqueClassCode_() {
  for (var attempt = 0; attempt < 50; attempt++) {
    var code = '';
    for (var i = 0; i < 6; i++) code += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
    if (!getClassRow_(code)) return code;
  }
  throw new Error('학급 코드 생성에 실패했어요. 다시 시도해 주세요.');
}

// 클래스 행의 특정 칸들을 갱신
function setClassFields_(classCode, fields) {
  const sheet = getSheet_(SHEET_CLASSES);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idxCode = header.indexOf('classCode');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idxCode]).toUpperCase() === String(classCode).toUpperCase()) {
      Object.keys(fields).forEach(function (key) {
        var c = header.indexOf(key);
        if (c >= 0) sheet.getRange(r + 1, c + 1).setValue(fields[key]);
      });
      return;
    }
  }
}


// ============================================================
// 데이터 헬퍼
// ============================================================

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(HEADERS[name]);
  }
  return sheet;
}

function ensureSheets_() {
  var ss = getSS_();
  Object.keys(HEADERS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(HEADERS[name]);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS[name]);
    }
  });
}

/**
 * ★ 추천: 데이터 이전 (직업·학생·학급코드·주소 모두 유지) ★
 * 월 → 회차 모델로 바꾼 뒤 "한 번만" 실행하세요.
 * 사용법: Apps Script 편집기 상단 함수 목록에서 'migrateToRounds' 선택 → [실행].
 *   - 직업(Jobs) 시트: 그대로 보존 (손대지 않음)
 *   - 학생 명단·학급코드·비밀번호: 그대로 보존 (회차=1, 상태=신청받는중으로 시작)
 *   - 옛 신청/배정 기록(월 기반): 정리(삭제) — 새 회차부터 다시 쌓입니다.
 */
function migrateToRounds() {
  var ss = getSS_();

  // 학급/학생은 옛 데이터를 읽어 새 구조로 다시 기록 (코드·이름·비번 보존)
  var classes  = readAll_(SHEET_CLASSES);
  var students = readAll_(SHEET_STUDENTS);

  var cs = ss.getSheetByName(SHEET_CLASSES) || ss.insertSheet(SHEET_CLASSES);
  cs.clear();
  cs.getRange(1, 1, 1, HEADERS.Classes.length).setValues([HEADERS.Classes]);
  classes.forEach(function (c) {
    if (!c.classCode) return;
    cs.appendRow([c.classCode, c.className || '우리 반', c.passwordHash || '',
                  1, ST_OPEN, c.createdAt || new Date()]);
  });

  var st = ss.getSheetByName(SHEET_STUDENTS) || ss.insertSheet(SHEET_STUDENTS);
  st.clear();
  st.getRange(1, 1, 1, HEADERS.Students.length).setValues([HEADERS.Students]);
  students.forEach(function (s) {
    if (!s.name) return;
    st.appendRow([s.classCode, s.studentId, s.number, s.name,
                  s.createdAt || new Date(), s.diligence || '']);
  });

  // 신청/배정: 옛 기록은 비우고 새 머리글만 (Jobs는 건드리지 않음)
  [SHEET_APPLICATIONS, SHEET_ASSIGNMENTS].forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
  });

  // Rounds(회차 상태): 새로 만들고 각 학급의 1회차를 '열림'으로 시드
  var rd = ss.getSheetByName(SHEET_ROUNDS) || ss.insertSheet(SHEET_ROUNDS);
  rd.clear();
  rd.getRange(1, 1, 1, HEADERS.Rounds.length).setValues([HEADERS.Rounds]);
  classes.forEach(function (c) {
    if (!c.classCode) return;
    rd.appendRow([c.classCode, 1, ST_OPEN, new Date()]);
  });

  // Jobs 시트는 없을 때만 생성 (있으면 데이터 그대로)
  if (!ss.getSheetByName(SHEET_JOBS)) {
    ss.insertSheet(SHEET_JOBS).getRange(1, 1, 1, HEADERS.Jobs.length).setValues([HEADERS.Jobs]);
  }

  return '이전 완료 — 직업·학생 명단·학급코드·주소 그대로, 1회차가 열렸어요.';
}

/**
 * ☢ 전체 초기화 (모두 삭제) — 직업까지 전부 지웁니다. 완전히 새로 시작할 때만.
 * 직업을 살리려면 위의 migrateToRounds 를 쓰세요.
 */
function resetAllData() {
  var ss = getSS_();
  Object.keys(HEADERS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    else sheet.clear();
    sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
  });
  return '초기화 완료 — 같은 주소 그대로 새 회차 모델로 시작할 수 있어요.';
}

/**
 * 🧹 잘못된 회차 정리 — Rounds/Applications/Assignments 시트에서
 * 회차 값이 이상한(날짜·음수·소수 등) 행을 삭제합니다.
 * "-2209105671999회차" 같은 유령 회차가 보일 때 한 번 실행하세요.
 * 사용법: 편집기 상단 함수 목록에서 'cleanupBadRounds' 선택 → [실행].
 * 정상 회차(1, 2, 3 …)와 그 지원/배정 기록은 그대로 유지됩니다.
 */
function cleanupBadRounds() {
  var removed = { Rounds: 0, Applications: 0, Assignments: 0 };
  [
    [SHEET_ROUNDS, 'Rounds'],
    [SHEET_APPLICATIONS, 'Applications'],
    [SHEET_ASSIGNMENTS, 'Assignments'],
  ].forEach(function (pair) {
    deleteRowsWhere_(pair[0], function (r) {
      if (validRound_(r.round)) return false;
      removed[pair[1]]++;
      return true;
    });
  });
  return '정리 완료 — 잘못된 회차 행 삭제: 회차 ' + removed.Rounds +
    '개 · 지원 ' + removed.Applications + '개 · 배정 ' + removed.Assignments + '개';
}

function readAll_(name) {
  var sheet = getSheet_(name);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var header = data[0];
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var obj = {};
    for (var c = 0; c < header.length; c++) obj[header[c]] = data[r][c];
    out.push(obj);
  }
  return out;
}

function deleteRowsWhere_(name, predicate) {
  var sheet = getSheet_(name);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var header = data[0];
  for (var r = data.length - 1; r >= 1; r--) {
    var obj = {};
    for (var c = 0; c < header.length; c++) obj[header[c]] = data[r][c];
    if (predicate(obj)) sheet.deleteRow(r + 1);
  }
}

function getClassRow_(classCode) {
  if (!classCode) return null;
  var code = String(classCode).trim().toUpperCase();
  var rows = readAll_(SHEET_CLASSES);
  var row = rows.find(function (r) { return String(r.classCode).toUpperCase() === code; });
  if (!row) return null;
  row.currentRound = Number(row.currentRound) || 1;
  row.status = row.status || ST_OPEN;
  return row;
}

// ===== 회차별 상태 (Rounds 시트) =====
// 올바른 회차 값인지 검사 (1 이상의 정수). 시트에 날짜 등 이상한 값이 섞여
// -2209105671999 같은 유령 회차가 생기는 것을 막는다.
function validRound_(v) {
  var n = Number(v);
  return isFinite(n) && n >= 1 && n === Math.floor(n);
}

function roundsOf_(classCode) {
  return readAll_(SHEET_ROUNDS)
    .filter(function (r) { return r.classCode === classCode && validRound_(r.round); })
    .map(function (r) { return { round: Number(r.round), status: r.status || ST_OPEN }; })
    .sort(function (a, b) { return a.round - b.round; });
}
function roundStatusOf_(classCode, round) {
  var f = roundsOf_(classCode).find(function (r) { return r.round === Number(round); });
  return f ? f.status : null;   // null = 아직 열린 적 없는 회차
}
function openRoundsOf_(classCode) {
  return roundsOf_(classCode).filter(function (r) { return r.status === ST_OPEN; }).map(function (r) { return r.round; });
}
function publishedRoundsOf_(classCode) {
  return roundsOf_(classCode).filter(function (r) { return r.status === ST_PUBLISHED; }).map(function (r) { return r.round; });
}
function setRoundStatus_(classCode, round, status) {
  var sheet = getSheet_(SHEET_ROUNDS);
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var iC = header.indexOf('classCode'), iR = header.indexOf('round'), iS = header.indexOf('status');
  for (var r = 1; r < data.length; r++) {
    if (data[r][iC] === classCode && Number(data[r][iR]) === Number(round)) {
      sheet.getRange(r + 1, iS + 1).setValue(status);
      return;
    }
  }
  sheet.appendRow([classCode, Number(round), status, new Date()]);
}
// 옛 학급(회차 행이 없는 경우) 보정: currentRound를 열린 회차로 등록
function ensureRoundsBackfill_(cls) {
  if (!cls) return;
  if (roundsOf_(cls.classCode).length === 0 && cls.currentRound >= 1) {
    setRoundStatus_(cls.classCode, cls.currentRound, cls.status || ST_OPEN);
  }
}

function jobsOf_(classCode) {
  return readAll_(SHEET_JOBS)
    .filter(function (j) { return j.classCode === classCode; })
    .sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); })
    .map(function (j) {
      return {
        jobId: j.jobId, jobName: j.jobName, icon: j.icon,
        capacity: Number(j.capacity) > 0 ? Number(j.capacity) : 1,
        isPopular: truthy_(j.isPopular),
        description: j.description || '',
      };
    });
}

function studentsOf_(classCode) {
  return readAll_(SHEET_STUDENTS)
    .filter(function (s) { return s.classCode === classCode; })
    .map(function (s) {
      return { studentId: s.studentId, number: String(s.number), name: s.name, diligence: s.diligence || '' };
    })
    .sort(function (a, b) { return (Number(a.number) || 0) - (Number(b.number) || 0); });
}


// ============================================================
// 잡다한 유틸
// ============================================================

function truthy_(v) {
  return v === true || v === 'Y' || v === 'y' || v === 1 || v === '1' || v === 'TRUE' || v === 'true';
}

// 번호 비교: "1"과 "01", 숫자 1을 같은 번호로 취급 (시트에 숫자로 저장돼도 매칭)
function sameNumber_(a, b) {
  var x = String(a == null ? '' : a).trim();
  var y = String(b == null ? '' : b).trim();
  if (x === y) return true;
  if (x === '' || y === '') return false;
  var nx = Number(x), ny = Number(y);
  return !isNaN(nx) && !isNaN(ny) && nx === ny;
}

function shuffle_(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = arr[i]; arr[i] = arr[j]; arr[j] = t;
  }
  return arr;
}
