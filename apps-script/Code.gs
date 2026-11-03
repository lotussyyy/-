/**
 * 부엉이 사서의 마법 도서관 · 방명록 서버 (Google Apps Script)
 * -------------------------------------------------------------
 * 이 코드는 구글 시트에 붙어서 방명록 글을 저장하고 돌려줍니다.
 * README.md의 "방명록 연결하기" 순서대로 붙여 넣고 배포하세요.
 *
 * 시트 '방명록'의 열:
 *   A id | B 작성시각 | C 닉네임 | D 구분 | E 내용 | F 삭제키 | G 숨김
 *
 * 관리 방법
 *   - 글을 숨기려면 G열(숨김)에 체크(TRUE)하세요.
 *   - 글을 완전히 지우려면 그 행을 삭제하세요.
 *   - F열(삭제키)은 글쓴이 본인 확인용이니 고치지 마세요.
 */

var SHEET_NAME = '방명록';
var MAX_RETURN = 300;              // 화면에 보여 줄 최근 글 수
var MAX_POSTS_PER_MINUTE = 20;     // 1분에 받을 수 있는 최대 글 수 (도배 방지)
var ROLES = ['신입생', '재학생', '졸업생', '방문자'];

function doGet(e) {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return json_({ ok: true, entries: [] });
  var start = Math.max(2, last - MAX_RETURN + 1);
  var rows = sh.getRange(start, 1, last - start + 1, 7).getValues();
  var entries = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    if (!r[0]) continue;
    if (r[6] === true || String(r[6]).toUpperCase() === 'TRUE') continue; // 숨김
    entries.push({
      id: String(r[0]),
      ts: r[1] instanceof Date ? r[1].getTime() : new Date(r[1]).getTime(),
      nick: String(r[2]),
      role: String(r[3]),
      msg: String(r[4])
    });
  }
  return json_({ ok: true, entries: entries });
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (d.action === 'delete') return delete_(d);
  if (d.website) return json_({ ok: true, entry: { id: 'x', ts: Date.now(), nick: '', role: '', msg: '' }, key: '' }); // 자동 등록 봇 차단

  var nick = clean_(d.nick, 12) || '익명의 독자';
  var role = ROLES.indexOf(d.role) >= 0 ? d.role : '방문자';
  var msg = clean_(d.msg, 300);
  if (!msg) return json_({ ok: false, error: 'empty' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var cache = CacheService.getScriptCache();
    var count = Number(cache.get('rate') || 0);
    if (count >= MAX_POSTS_PER_MINUTE) return json_({ ok: false, error: 'busy' });
    cache.put('rate', String(count + 1), 60);

    var id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    var key = Utilities.getUuid();
    var now = new Date();
    sheet_().appendRow([id, now, safe_(nick), role, safe_(msg), key, false]);
    return json_({ ok: true, entry: { id: id, ts: now.getTime(), nick: nick, role: role, msg: msg }, key: key });
  } finally {
    lock.releaseLock();
  }
}

function delete_(d) {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2 || !d.id || !d.key) return json_({ ok: false, error: 'not_found' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var vals = sh.getRange(2, 1, last - 1, 6).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === String(d.id)) {
        if (String(vals[i][5]) !== String(d.key)) return json_({ ok: false, error: 'forbidden' });
        sh.deleteRow(i + 2);
        return json_({ ok: true });
      }
    }
    return json_({ ok: false, error: 'not_found' });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id', '작성시각', '닉네임', '구분', '내용', '삭제키', '숨김']);
    sh.setFrozenRows(1);
    sh.getRange('G2:G').insertCheckboxes();
    sh.setColumnWidth(5, 420);
  }
  return sh;
}

function clean_(s, n) {
  return String(s || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim().slice(0, n);
}

// '='·'+'·'-'·'@'로 시작하는 글이 시트에서 수식으로 실행되지 않도록 막습니다.
function safe_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 처음 한 번 실행해서 '방명록' 시트를 만들고 권한을 승인하는 용도입니다. */
function setup() {
  sheet_();
}
