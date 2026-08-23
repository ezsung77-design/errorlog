/**
 * 아파트 펀치리스트 · Google Apps Script 백엔드
 * ------------------------------------------------------------------
 * 배포 방법
 *  1) 구글 스프레드시트 새로 만들기 → 확장 프로그램 → Apps Script
 *  2) 이 파일 내용을 Code.gs 에 붙여넣기
 *  3) (선택) PHOTO_FOLDER_ID 에 사진 저장용 드라이브 폴더 ID 입력
 *     비워두면 "펀치리스트_사진" 폴더를 내 드라이브 루트에 자동 생성
 *  4) 배포 → 새 배포 → 유형: 웹 앱
 *     - 실행 계정: 나
 *     - 액세스 권한: 링크가 있는 모든 사용자(사내 공유 시 '조직 내 사용자')
 *  5) 발급된 /exec URL 을 앱 설정(⚙) 화면에 입력
 *
 * 스프레드시트 시트 3개는 최초 호출 시 자동 생성됩니다.
 *  - PUNCH   : 펀치 원장(1행 = 1펀치)
 *  - MASTER  : 드롭다운 마스터(현장/동/공종/하자유형/업체)
 *  - LOG     : 상태 변경 이력
 */

/** 사진을 저장할 드라이브 폴더 ID. 비우면 자동 생성 */
var PHOTO_FOLDER_ID = '';

/** 시트 이름 */
var SH_PUNCH  = 'PUNCH';
var SH_MASTER = 'MASTER';
var SH_LOG    = 'LOG';

/** PUNCH 시트 컬럼 정의: [키, 헤더명] — 순서를 바꾸면 시트 열 순서도 바뀝니다 */
var COLS = [
  ['id',        'ID'],
  ['created_at','등록일시'],
  ['updated_at','수정일시'],
  ['site',      '현장'],
  ['building',  '동'],
  ['unit',      '호수'],
  ['room',      '실'],
  ['part',      '부위'],
  ['trade',     '공종'],
  ['defect',    '하자유형'],
  ['severity',  '중요도'],
  ['title',     '제목'],
  ['note',      '특기사항'],
  ['status',    '상태'],
  ['vendor',    '담당업체'],
  ['due',       '조치기한'],
  ['done_at',   '완료일시'],
  ['reporter',  '등록자'],
  ['photo_before','사진(조치전)'],
  ['photo_after', '사진(조치후)'],
  ['photo_meta',  '사진메타(JSON)'],
  ['deleted',   '삭제']
];

/** MASTER 시트 기본값 (열 = 항목 그룹) */
var MASTER_DEFAULT = {
  '현장'    : ['○○아파트 1단지'],
  '동'      : ['101동','102동','103동'],
  '실'      : ['거실','주방','현관','안방','침실1','침실2','드레스룸','욕실1','욕실2','발코니','다용도실','복도','계단실','공용부'],
  '부위'    : ['벽','천장','바닥','창호','문','걸레받이','몰딩','타일','위생기구','가구','조명','기타'],
  '공종'    : ['골조','조적','미장','방수','타일','도장','도배','바닥재','창호','유리','목공','가구','위생기구','기계설비','전기','소방','조경','청소','기타'],
  '하자유형': ['균열','누수','결로','오염','파손','스크래치','미시공','시공불량','마감불량','수직수평불량','작동불량','치수오차','이색','누락','기타'],
  '중요도'  : ['긴급','중요','보통'],
  '상태'    : ['접수','조치중','조치완료','확인완료','보류'],
  '담당업체': ['미지정']
};

/* ============================ 엔트리 포인트 ============================ */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'list';
  try {
    if (action === 'ping')    return json({ ok: true, version: 'punch-1.0.0' });
    if (action === 'masters') return json({ ok: true, masters: readMasters() });
    if (action === 'list') {
      var since = e.parameter.since || '';
      var site  = e.parameter.site  || '';
      return json({ ok: true, records: readRecords(since, site), masters: readMasters() });
    }
    return json({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); }
  catch (err) { return json({ ok: false, error: 'invalid JSON body' }); }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var action = body.action || 'upsert';
    if (action === 'photo')  return json(savePhoto(body));
    if (action === 'upsert') return json(upsertRecords(body.records || []));
    if (action === 'delete') return json(softDelete(body.ids || []));
    return json({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* ============================== 시트 준비 ============================== */

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function punchSheet() {
  var sh = ss().getSheetByName(SH_PUNCH);
  if (!sh) {
    sh = ss().insertSheet(SH_PUNCH);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS.map(function (c) { return c[1]; })]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setColumnWidth(colIndex('note'), 320);
    sh.setColumnWidth(colIndex('photo_meta'), 60);
  }
  return sh;
}

function masterSheet() {
  var sh = ss().getSheetByName(SH_MASTER);
  if (!sh) {
    sh = ss().insertSheet(SH_MASTER);
    var groups = Object.keys(MASTER_DEFAULT);
    var maxLen = 0;
    groups.forEach(function (g) { maxLen = Math.max(maxLen, MASTER_DEFAULT[g].length); });
    var rows = [groups];
    for (var i = 0; i < maxLen; i++) {
      rows.push(groups.map(function (g) { return MASTER_DEFAULT[g][i] || ''; }));
    }
    sh.getRange(1, 1, rows.length, groups.length).setValues(rows);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, groups.length).setFontWeight('bold').setBackground('#f1f3f4');
  }
  return sh;
}

function logSheet() {
  var sh = ss().getSheetByName(SH_LOG);
  if (!sh) {
    sh = ss().insertSheet(SH_LOG);
    sh.getRange(1, 1, 1, 5).setValues([['일시', '펀치ID', '변경필드', '변경값', '작업자']]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f1f3f4');
  }
  return sh;
}

function colIndex(key) {
  for (var i = 0; i < COLS.length; i++) if (COLS[i][0] === key) return i + 1;
  return -1;
}

/* ============================== 읽기 ============================== */

function readRecords(since, site) {
  var sh = punchSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, COLS.length).getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    var rec = {};
    for (var c = 0; c < COLS.length; c++) rec[COLS[c][0]] = cell(values[r][c]);
    if (!rec.id) continue;
    if (site && rec.site && rec.site !== site) continue;
    if (since && rec.updated_at && rec.updated_at <= since) continue;
    rec.deleted = rec.deleted === 'Y' || rec.deleted === true;
    rec.photo_before = splitList(rec.photo_before);
    rec.photo_after  = splitList(rec.photo_after);
    out.push(rec);
  }
  return out;
}

function readMasters() {
  var sh = masterSheet();
  var last = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (last < 1 || lastCol < 1) return {};
  var values = sh.getRange(1, 1, last, lastCol).getValues();
  var head = values[0];
  var out = {};
  for (var c = 0; c < head.length; c++) {
    var name = String(head[c] || '').trim();
    if (!name) continue;
    var list = [];
    for (var r = 1; r < values.length; r++) {
      var v = String(values[r][c] || '').trim();
      if (v) list.push(v);
    }
    out[name] = list;
  }
  return out;
}

function cell(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  return v === null || v === undefined ? '' : String(v);
}

function splitList(v) {
  if (!v) return [];
  return String(v).split(/[\n,]+/).map(function (s) { return s.trim(); }).filter(String);
}

/* ============================== 쓰기 ============================== */

/**
 * 레코드 upsert. id 기준으로 기존 행을 찾고, 없으면 추가.
 * updated_at 이 시트 값보다 최신인 경우에만 덮어씁니다(단순 LWW 병합).
 */
function upsertRecords(records) {
  if (!records.length) return { ok: true, saved: 0 };
  var sh = punchSheet();
  var last = sh.getLastRow();
  var idCol = colIndex('id'), upCol = colIndex('updated_at');
  var index = {};
  if (last >= 2) {
    var ids = sh.getRange(2, idCol, last - 1, 1).getValues();
    var ups = sh.getRange(2, upCol, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var id = String(ids[i][0] || '');
      if (id) index[id] = { row: i + 2, updated_at: cell(ups[i][0]) };
    }
  }

  var appends = [], saved = 0, logs = [];
  records.forEach(function (rec) {
    if (!rec || !rec.id) return;
    var row = toRow(rec);
    var hit = index[rec.id];
    if (hit) {
      if (rec.updated_at && hit.updated_at && rec.updated_at <= hit.updated_at) return; // 서버가 더 최신
      sh.getRange(hit.row, 1, 1, COLS.length).setValues([row]);
    } else {
      appends.push(row);
      index[rec.id] = { row: -1, updated_at: rec.updated_at || '' };
    }
    saved++;
    logs.push([new Date(), rec.id, '상태', rec.status || '', rec.reporter || '']);
  });

  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, COLS.length).setValues(appends);
  }
  if (logs.length) {
    logSheet().getRange(logSheet().getLastRow() + 1, 1, logs.length, 5).setValues(logs);
  }
  return { ok: true, saved: saved };
}

function toRow(rec) {
  return COLS.map(function (c) {
    var k = c[0], v = rec[k];
    if (k === 'deleted') return v ? 'Y' : '';
    if (k === 'photo_before' || k === 'photo_after') {
      return Array.isArray(v) ? v.join('\n') : (v || '');
    }
    if (k === 'photo_meta' && v && typeof v === 'object') return JSON.stringify(v);
    return v === null || v === undefined ? '' : String(v);
  });
}

function softDelete(ids) {
  var sh = punchSheet();
  var last = sh.getLastRow();
  if (last < 2 || !ids.length) return { ok: true, deleted: 0 };
  var idCol = colIndex('id'), delCol = colIndex('deleted'), upCol = colIndex('updated_at');
  var values = sh.getRange(2, idCol, last - 1, 1).getValues();
  var now = nowIso(), n = 0;
  for (var i = 0; i < values.length; i++) {
    if (ids.indexOf(String(values[i][0])) >= 0) {
      sh.getRange(i + 2, delCol).setValue('Y');
      sh.getRange(i + 2, upCol).setValue(now);
      n++;
    }
  }
  return { ok: true, deleted: n };
}

/* ============================== 사진 업로드 ============================== */

function photoFolder() {
  if (PHOTO_FOLDER_ID) return DriveApp.getFolderById(PHOTO_FOLDER_ID);
  var it = DriveApp.getFoldersByName('펀치리스트_사진');
  return it.hasNext() ? it.next() : DriveApp.createFolder('펀치리스트_사진');
}

function subFolder(parent, name) {
  if (!name) return parent;
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/**
 * body = { action:'photo', data:'<base64>', mime:'image/jpeg',
 *          name:'101-1502.jpg', site:'', building:'', kind:'before'|'after' }
 * 반환 { ok, fileId, url, thumb }
 */
function savePhoto(body) {
  if (!body.data) return { ok: false, error: 'no data' };
  var root = photoFolder();
  var folder = subFolder(subFolder(root, body.site || '기타현장'), body.building || '');
  var mime = body.mime || 'image/jpeg';
  var name = body.name || (Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.jpg');
  var blob = Utilities.newBlob(Utilities.base64Decode(body.data), mime, name);
  var file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // 조직 정책으로 외부 공유가 막힌 경우: 링크는 조직 내부에서만 열립니다.
  }
  var id = file.getId();
  return {
    ok: true,
    fileId: id,
    url: 'https://drive.google.com/file/d/' + id + '/view',
    thumb: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000'
  };
}

/* ============================== 유틸 ============================== */

function nowIso() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 스프레드시트 메뉴에서 수동으로 초기화할 때 사용 */
function setup() {
  punchSheet(); masterSheet(); logSheet();
  SpreadsheetApp.getUi().alert('시트 3개(PUNCH/MASTER/LOG)를 준비했습니다.');
}
